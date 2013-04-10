open System
open DataAnalysis
open System.IO

printfn "creating database"
let db = DataAnalysis.MTurkData("C:\Users\dbarowy\Desktop\mtresults\ResultsDB.sql")
printfn "Importing C:\Users\dbarowy\Desktop\mtresults\Batch_1081584_batch_results.csv"
db.ImportCSV("C:\Users\dbarowy\Desktop\mtresults\Batch_1081584_batch_results.csv")
printfn "Importing C:\Users\dbarowy\Desktop\mtresults\Batch_1082662_batch_results.csv"
db.ImportCSV("C:\Users\dbarowy\Desktop\mtresults\Batch_1082662_batch_results.csv")
printfn "Importing C:\Users\dbarowy\Desktop\mtresults\Batch_1081143_batch_results.csv"
db.ImportCSV("C:\Users\dbarowy\Desktop\mtresults\Batch_1081143_batch_results.csv")
printfn "Importing C:\Users\dbarowy\Desktop\mtresults\Batch_1081138_batch_results.csv"
db.ImportCSV("C:\Users\dbarowy\Desktop\mtresults\Batch_1081138_batch_results.csv")
printfn "Importing C:\Users\dbarowy\Desktop\mtresults\Batch_1080232_batch_results.csv"
db.ImportCSV("C:\Users\dbarowy\Desktop\mtresults\Batch_1080232_batch_results.csv")
printfn "done"

match Console.ReadLine() with
    | "quit" -> ()
    | expr -> ()