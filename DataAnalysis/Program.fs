open System
open DataAnalysis
open System.IO

printfn "creating database"
let db = DataAnalysis.MTurkData("C:\Users\Dan Barowy\Desktop\TestDB.sql")
printfn "done"

match Console.ReadLine() with
    | "quit" -> ()
    | expr -> ()