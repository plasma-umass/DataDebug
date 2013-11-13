open System
open System.IO
open System.Collections.Generic
open Microsoft.Office.Interop.Excel

type App = Microsoft.Office.Interop.Excel.Application
type TextFieldParser = Microsoft.VisualBasic.FileIO.TextFieldParser
type FieldType = Microsoft.VisualBasic.FileIO.FieldType

// VisualBasic.NET has a handy-dandy CSV parser
let ParseCSV(path: string) =
    let parser = new TextFieldParser(path)
    parser.TextFieldType <- FieldType.Delimited
    parser.SetDelimiters(",")
    let mutable rows = []
    while not (parser.EndOfData) do
        rows <- parser.ReadFields() :: rows
    parser.Close()
    // convert to array
    let outarray = List.rev rows |> List.toArray
    // exclude the first element
    outarray.[1..outarray.Length-1]

// some constants
type Offsets =
    static member HITDEF_WIDTH = 7
    // MTurk prepends 27 fields to all HITs.
    static member OFFSET_WORKER_ID = 15
    // Then come the following:
    static member OFFSET_STATE_ID = 27
    // + (n * HITDEF_WIDTH) columns, where n = inputs_per_hit
    // the HITDEF_WIDTH columns describe the following:
    // path, workbook, worksheet, row, column, original string, image url
    static member OFFSET_PATH = 0
    static member OFFSET_WB = 1
    static member OFFSET_WS = 2
    static member OFFSET_ROW = 3
    static member OFFSET_COL = 4
    static member OFFSET_ORIG = 5
    static member OFFSET_URL = 6
    static member OFFSET_HITDEFS = 28
    // + n answers
    static member OFFSET_ANSWERS inputs_per_hit = 28 + (inputs_per_hit * Offsets.HITDEF_WIDTH)
    // last two columns are for "Approve" and "Reject"

// all processed data goes into one of these
type Data() =
    // address -> (original value, retyped value)
    let input_pairs = new Dictionary<AST.Address, string*string>()
    // worker_id -> address list
    let worker_ids = new Dictionary<string, AST.Address list>()
    // this data structure is for sanity-check purposes
    let addresses = new HashSet<AST.Address>()
    member self.LearnFromCSV(csvdata: string[][]) =
        // get width
        let inputs_per_hit = Data.CalculateNumInputs(csvdata)

        // each row
        Array.iteri (fun i row ->
            // create worker ID key in dictionary if it doesn't alreay exist
            let worker_id = csvdata.[i].[Offsets.OFFSET_WORKER_ID]
            if not (worker_ids.ContainsKey(worker_id)) then
                worker_ids.Add(worker_id, [])

            // each input
            Array.iter (fun j ->
                let path = csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_PATH]
                let workbook = csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_WB]
                let worksheet = csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_WS]
                let addr_r = Int32.Parse(csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_ROW])
                let addr_c = Int32.Parse(csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_COL])
                let image_url = csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_URL]
                let original_string = csvdata.[i].[Offsets.OFFSET_HITDEFS + j * Offsets.HITDEF_WIDTH + Offsets.OFFSET_ORIG]
                let retyped_string = csvdata.[i].[Offsets.OFFSET_ANSWERS(inputs_per_hit) + j]

                // get address object
                let addr = AST.Address.FromR1C1(addr_r, addr_c, worksheet, workbook, path)

                // add address to address hashset
                addresses.Add(addr) |> ignore

                // add address to worker dict
                worker_ids.[worker_id] <- addr :: worker_ids.[worker_id]

                // add input pairs
                input_pairs.Add(addr, (original_string, retyped_string))
            ) [|0..inputs_per_hit-1|]
        ) csvdata
    member self.OverallAccuracy =
        let numsame = Seq.fold (fun acc (pair: KeyValuePair<AST.Address,string*string>) ->
                        acc + if fst pair.Value = snd pair.Value then 1 else 0
                      ) 0 input_pairs
        System.Convert.ToDouble(numsame) / System.Convert.ToDouble(input_pairs.Count)
    member self.NumInputs = input_pairs.Count
    member self.NumWorkers = worker_ids.Count
    member self.MaxWorker = worker_ids.[self.MaxWorkerID].Length
    member self.MaxWorkerID : string =
        // get (worker_id, count)
        let w_counts = Seq.map (fun (pair: KeyValuePair<string, AST.Address list>) -> pair.Key, pair.Value.Length) worker_ids
        let max_w_pair = w_counts |> Seq.sortBy (fun (worker_id: string, count: int) -> -count) |> Seq.head
        fst max_w_pair
    member self.AssignmentsPerWorker worker_id = worker_ids.[worker_id].Length
    member self.WorkerAccuracy worker_id =
        let addrs = worker_ids.[worker_id]
        let numcorrect = (List.filter (fun addr -> fst input_pairs.[addr] = snd input_pairs.[addr]) addrs).Length
        System.Convert.ToDouble(numcorrect) / System.Convert.ToDouble(addrs.Length)
    member self.WorkerAssignments worker_id = worker_ids.[worker_id].Length
    // sorted by the most accurate worker
    member self.WorkerIDsSortedByAccuracy = 
        Seq.sortBy (fun (pair: KeyValuePair<string, AST.Address list>) -> -self.WorkerAccuracy pair.Key) worker_ids |>
        Seq.map (fun (pair: KeyValuePair<string, AST.Address list>) -> pair.Key)
    member self.StringPairs = Seq.map (fun (pair: KeyValuePair<AST.Address,string*string>) -> pair.Value) input_pairs
    static member private CalculateNumInputs(csvdata: string[][]) : int =
        let total_width = csvdata.[0].Length
        let payload_cols = total_width - Offsets.OFFSET_HITDEFS
        payload_cols / (Offsets.HITDEF_WIDTH + 1) // the +1 is to include the input's answer column

[<EntryPoint>]
let main argv = 
    if argv.Length < 1 then
        Console.WriteLine("Usage: SimulationApp.exe [MTurk input CSV 1] ... [MTurk input CSV N]") |> ignore
        1
    else
        // process args
        let files = argv
        let numfiles = argv.Length

        // parse the data
        let csvdatas = Array.map (fun f -> ParseCSV f) files

        // process data
        let data = Data()
        Array.iter (fun csvdata -> data.LearnFromCSV csvdata) csvdatas

        // get basic stats
        Console.WriteLine("{0:P} of inputs typed correctly.", data.OverallAccuracy) |> ignore
        Console.WriteLine("{0} workers participated.", data.NumWorkers) |> ignore
        Console.WriteLine("The fastest worker completed {0} data re-entries", data.MaxWorker) |> ignore
        Console.WriteLine("The fastest worker had an accuracy of {0:P}", data.WorkerAccuracy(data.MaxWorkerID)) |> ignore

        // pause
//        Console.WriteLine("Press any key to continue.") |> ignore
//        Console.ReadLine() |> ignore

        for worker_id in data.WorkerIDsSortedByAccuracy do
            Console.WriteLine("Worker {0} completed {1} assignments with an accuracy of {2:P}.", worker_id, data.WorkerAssignments(worker_id), data.WorkerAccuracy(worker_id)) |> ignore

        // train classifier
        let total_inputs = data.NumInputs
        let c = UserSimulation.Classification()
        Seq.iteri (fun i (orig,entered) ->
            Console.Write("\r{0:P} strings classified", System.Convert.ToDouble(i) / System.Convert.ToDouble(total_inputs))
            c.ProcessTypos(orig,entered) |> ignore
        ) data.StringPairs
        Console.Write("\n")

        0   // A-OK
