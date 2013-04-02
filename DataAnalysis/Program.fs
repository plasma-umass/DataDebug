open System
open DataAnalysis
open System.IO

printfn "creating database"
let db = DataAnalysis.MTurkData("C:\Users\Dan Barowy\Desktop\TestDB.sql")
let csvfile = "C:\Users\Dan Barowy\Desktop\Batch_1082662_batch_results.csv"
let csv = File.ReadAllText(csvfile)
let rows = CSVParser.ParseCsv csv ","
let mutable count = 0
let fileid = db.AddFile("workbook.xls", csvfile)
for row in rows do
    if count <> 0 then
        let hitid = row.[0]
        let hittypeid = row.[1]
        let title = row.[2]
        let description = row.[3]
        let keywords = row.[4]
        let reward = Decimal.Parse(row.[5].Substring(1))    // drop leading '$'
        let creationtime = row.[6]
        let maxassignments = CSVParser.ZeroOrNum row.[7]
        let requesterannotation = row.[8]
        let assignmentdurationinseconds = CSVParser.ZeroOrNum row.[9]
        let autoapprovaldelayinseconds = CSVParser.ZeroOrNum row.[10]
        let expiration = row.[11]
        let numberofsimilarhits = CSVParser.ZeroOrNum row.[12]
        let lifetimeinseconds = CSVParser.ZeroOrNum row.[13]
        let assignmentid = row.[14]
        let workerid = row.[15]
        let assignmentstatus = row.[16]
        let accepttime = row.[17]
        let submittime = row.[18]
        let autoapprovaltime = row.[19]
        let approvaltime = row.[20]
        let rejectiontime = row.[21]
        let requesterfeedback = row.[22]
        let worktimeinseconds = CSVParser.ZeroOrNum row.[23]
        let lifetimeapprovalrate = row.[24]
        let last30daysapprovalrate = row.[25]
        let hitid = db.AddHIT(fileid,
                              hitid,
                              hittypeid,
                              title,
                              description,
                              keywords,
                              reward,
                              creationtime,
                              maxassignments,
                              requesterannotation,
                              assignmentdurationinseconds,
                              autoapprovaldelayinseconds,
                              expiration,
                              numberofsimilarhits,
                              lifetimeinseconds,
                              assignmentid,
                              workerid,
                              assignmentstatus,
                              accepttime,
                              submittime,
                              autoapprovaltime,
                              approvaltime,
                              rejectiontime,
                              requesterfeedback,
                              worktimeinseconds,
                              lifetimeapprovalrate,
                              last30daysapprovalrate)
        db.AddAnswerWithErrors(hitid, "blahblahblah", hitid, [DataAnalysis.ErrorType.ExtraDigit]) |> ignore
        db.AddAnswerWithErrors(hitid, "blahblahblah", hitid, [DataAnalysis.ErrorType.ExtraDigit]) |> ignore
    count <- count + 1

printfn "done"

match Console.ReadLine() with
    | "quit" -> ()
    | expr -> ()