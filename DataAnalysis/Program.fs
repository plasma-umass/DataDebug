open System
open DataAnalysis
open System.IO

printfn "creating database"
let db = DataAnalysis.MTurkData("C:\Users\Dan Barowy\Desktop\TestDB.sql")
db.AddFile("foo.xls", "blah.csv")
let hitid = db.AddHIT("hitid",
                      "hittypeid",
                      "title",
                      "description",
                      "keywords",
                      0.01m,
                      "Tue Mar 26 22:56:54 GMT 2013",
                      1,
                      "requesterannotation",
                      1,
                      1,
                      "Tue Mar 26 22:56:54 PDT 2013",
                      1,
                      1,
                      "assignmentstatus",
                      "Tue Mar 26 22:56:54 GMT 2013",
                      "Tue Mar 26 22:56:54 GMT 2013",
                      "Tue Mar 26 22:56:54 GMT 2013",
                      "Tue Mar 26 22:56:54 GMT 2013",
                      "Tue Mar 26 22:56:54 GMT 2013",
                      "requesterfeedback",
                      1,
                      "lifetimeapprovalrate",
                      "last30daysapprovalrate")
let answer1 = db.AddAnswerWithErrors(1, "blahblahblah", hitid, [DataAnalysis.ErrorType.ExtraDigit])
let answer2 = db.AddAnswerWithErrors(2, "foofoofoo", hitid, [DataAnalysis.ErrorType.MantissaError])

printfn "done"

match Console.ReadLine() with
    | "quit" -> ()
    | expr -> ()