open System
open System.IO
open System.Collections.Generic
open Microsoft.Office.Interop.Excel

type App = Microsoft.Office.Interop.Excel.Application
type TextFieldParser = Microsoft.VisualBasic.FileIO.TextFieldParser
type FieldType = Microsoft.VisualBasic.FileIO.FieldType

let NUMTRIALS = 1
let NBOOTS = System.Convert.ToInt32(Math.E * 1000.0)
let SIGNIFICANCE = 0.95
let THRESHOLD = 0.05

// recursively get all spreadsheets that don't have the "bad" prefix in their filename
let EnumSpreadsheets(dir: string) : string[] =
    System.IO.Directory.GetFiles(dir, "*.xls", System.IO.SearchOption.AllDirectories)

[<EntryPoint>]
let main argv = 
    if argv.Length < 3 then
        Console.WriteLine("Usage: SimulationApp.exe [classification file] [spreadsheet dir] [output dir] [MTurk input CSV 1] ... [MTurk input CSV N]") |> ignore
        1
    else
        // process args
        let serfile = argv.[0]
        let xlsdir = argv.[1]
        let outdir = argv.[2]
        let files = argv.[3..]
        let numfiles = argv.Length

        // process data if we haven't already
        let c = if not (System.IO.File.Exists(serfile)) then
                    // parse the data
                    let csvdatas = Array.map (fun f -> MTurkParser.Data.ParseCSV f) files

                    let data = MTurkParser.Data()
                    Array.iter (fun csvdata -> data.LearnFromCSV csvdata) csvdatas

                    // get basic stats
                    Console.WriteLine("{0:P} of inputs typed correctly.", data.OverallAccuracy) |> ignore
                    Console.WriteLine("{0} workers participated.", data.NumWorkers) |> ignore
                    Console.WriteLine("The fastest worker completed {0} data re-entries", data.MaxWorker) |> ignore
                    Console.WriteLine("The fastest worker had an accuracy of {0:P}", data.WorkerAccuracy(data.MaxWorkerID)) |> ignore

//                    for worker_id in data.WorkerIDsSortedByAccuracy do
//                        Console.WriteLine("Worker {0} completed {1} assignments with an accuracy of {2:P}.", worker_id, data.WorkerAssignments(worker_id), data.WorkerAccuracy(worker_id)) |> ignore

                    // train classifier & save output as a side-effect
                    UserSimulation.Classification.Classify(data, serfile)
                else
                    Console.WriteLine("Reopening previously-saved classification database: {0}", serfile) |> ignore
                    UserSimulation.Classification.Deserialize(serfile)

        // start up a copy of Excel
        let a = DataDebugMethods.Utility.NewApplicationInstance() :?> Application

        // run user simulation experiment for every spreadsheet in xlsdir
        // NUMTRIALS times & save in outdir
//        let xlss = EnumSpreadsheets(xlsdir) |> Seq.toArray
//        let wbr = System.Text.RegularExpressions.Regex(@"\\\\(.+\\)*(.+\..+)$", System.Text.RegularExpressions.RegexOptions.Compiled)
//        let results = Array.map (fun xls ->
//                          Array.map (fun i ->
//                              let xlfile = wbr.Match(xls).Groups.[2].Value
//                              Console.WriteLine("Opening {0}", xlfile) |> ignore
//                              let usersim = new UserSimulation.Simulation()
//                              usersim.Run(NBOOTS, xls, SIGNIFICANCE, a, THRESHOLD, serfile, rng)
//                              usersim.Serialize(outdir + "\\" + xlfile + "_" + i.ToString() + ".bin")
//                              Console.WriteLine("Closing {0}", xlfile) |> ignore
//                              usersim
//                          ) [|0..NUMTRIALS-1|]
//                      ) xlss

        // print results
        Console.WriteLine("Done.") |> ignore

        0   // A-OK
