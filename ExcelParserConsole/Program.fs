open System
open ExcelParser

printfn "Excel Parser Console"
printfn "Type a formula expression or 'quit'."

let rec readAndProcess() =
    printf "excel: "
    match Console.ReadLine() with
    | "quit" -> ()
    | expr ->
        try
            printfn "Parsing..."
            ExcelParser.ConsoleTest expr
            
        with ex ->
            printfn "Unhandled Exception: %s" ex.Message

        readAndProcess()

readAndProcess()