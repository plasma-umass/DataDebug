open System
open DataAnalysis

printfn "creating database"
let db = DataAnalysis.MTurkData("data source=\"C:\Users\Dan Barowy\Desktop\TestDB.sql\"")
printfn "done"

match Console.ReadLine() with
    | "quit" -> ()
    | expr -> ()