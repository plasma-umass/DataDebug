module DataAnalysis
open System
open System.IO
open System.Data.SQLite
open System.Globalization
open DataDebugMethods

type ErrorType =
    | DigitTransposition = 0
    | ExtraDigit = 1
    | OmittedDigit = 2
    | SignError = 3
    | MantissaError = 4
    | Other = 10

type MTurkData(filename: string) =
    let mutable _conn: SQLiteConnection option = None
    let ToTimestamp(dt_utc: DateTime) : string =
        (dt_utc - (new DateTime(1970, 1, 1))).TotalSeconds.ToString()
    let FromMTurkTime(datestring: string) : DateTime =
        let utc = TimeZoneInfo.Utc
        try
            let pdtformatstring = "ddd MMM dd HH:mm:ss 'PDT' yyyy"
            let dt = DateTime.ParseExact(datestring, pdtformatstring, CultureInfo.InvariantCulture)
            let pdt = TimeZoneInfo.CreateCustomTimeZone("Pacific Daylight Time", new TimeSpan(-07, 00, 00), "(UTC-07:00) Pacific Daylight Time", "Pacific Daylight Time")
            TimeZoneInfo.ConvertTime(dt, pdt, utc)
        with
        | _ ->
            try
                let gmtformatstring = "ddd MMM dd HH:mm:ss 'GMT' yyyy"
                DateTime.ParseExact(datestring, gmtformatstring, CultureInfo.InvariantCulture)
            with
            | _ -> DateTime.Parse(datestring)
    let TurkTimeToTimestamp(datestring: string) =
        match datestring with
        | "" -> "null"
        | _ -> ToTimestamp(FromMTurkTime(datestring))
    let Connected() =
        match _conn with
        | Some(c) -> true
        | None -> false
    let CreateDatabase() =
        let conn = new SQLiteConnection("data source=\"" + filename + "\"")
        if (conn <> null) then
            _conn <- Some(conn)
        else
            failwith "Unable to connect to database."
        conn.Open()

        let cmd = new SQLiteCommand(conn)

        cmd.CommandText <- "CREATE TABLE files ( id INTEGER PRIMARY KEY AUTOINCREMENT," +
                                              " mturkfilename TEXT," +
                                              " benchmarkfilename TEXT" +
                                              ")" 
        cmd.ExecuteNonQuery() |> ignore

        cmd.CommandText <- "CREATE TABLE hits ( id INTEGER PRIMARY KEY AUTOINCREMENT," +
                                             " fileid INTEGER," +
                                             " hitid TEXT," +
                                             " hittypeid TEXT," +
                                             " title TEXT," +
                                             " description TEXT," +
                                             " keywords TEXT," +
                                             " reward NUMERIC," +
                                             " creationtime INTEGER," +
                                             " maxassignments NUMERIC," +
                                             " requesterannotation TEXT," +
                                             " assignmentdurationinseconds NUMERIC," +
                                             " autoapprovaldelayinseconds NUMERIC," +
                                             " expiration INTEGER," +
                                             " numberofsimilarhits NUMERIC," +
                                             " lifetimeinseconds NUMERIC," +
                                             " assignmentid TEXT," +
                                             " workerid TEXT," +
                                             " assignmentstatus TEXT," +
                                             " accepttime INTEGER," +
                                             " submittime INTEGER," +
                                             " autoapprovaltime INTEGER," +
                                             " approvaltime INTEGER," +
                                             " rejectiontime INTEGER," +
                                             " requesterfeedback TEXT," +
                                             " worktimeinseconds NUMERIC," +
                                             " lifetimeapprovalrate TEXT," +
                                             " last30daysapprovalrate TEXT," +
                                             " last7daysapprovalrate TEXT," +
                                             " FOREIGN KEY(fileid) REFERENCES files(id)" +
                                             ")"
        cmd.ExecuteNonQuery() |> ignore

        cmd.CommandText <- "CREATE TABLE answers ( id INTEGER PRIMARY KEY AUTOINCREMENT," +
                                                 " cell NUMERIC," +
                                                 " origdata TEXT," +
                                                 " userdata TEXT," +
                                                 " is_different BOOLEAN," +
                                                 " hitid INTEGER," +
                                                 " FOREIGN KEY(hitid) REFERENCES hits(id)" +
                                                 ")"
        cmd.ExecuteNonQuery() |> ignore

        cmd.CommandText <- "CREATE TABLE errortypes ( id INTEGER PRIMARY KEY," +
                                                    " name TEXT" +
                                                    ")"
        cmd.ExecuteNonQuery() |> ignore

        cmd.CommandText <- "CREATE TABLE answers_errors ( id INTEGER PRIMARY KEY AUTOINCREMENT," +
                                                        " answerid INTEGER," +
                                                        " errortypeid INTEGER," +
                                                        " FOREIGN KEY(answerid) REFERENCES answers(id)," +
                                                        " FOREIGN KEY(errortypeid) REFERENCES errortypes(id)" +
                                                        ")"
        cmd.ExecuteNonQuery() |> ignore

        // populate error table
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.DigitTransposition).ToString() + ",\"" +
                            ErrorType.DigitTransposition.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.ExtraDigit).ToString() + ",\"" +
                            ErrorType.ExtraDigit.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.OmittedDigit).ToString() + ",\"" +
                            ErrorType.OmittedDigit.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.SignError).ToString() + ",\"" +
                            ErrorType.SignError.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.MantissaError).ToString() + ",\"" +
                            ErrorType.MantissaError.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.Other).ToString() + ",\"" +
                            ErrorType.Other.ToString() + "\")"
        cmd.ExecuteNonQuery() |> ignore
    let OpenDatabase() =
        let conn = new SQLiteConnection("data source=\"" + filename + "\"")
        if (conn <> null) then
            printfn "Opening database."
            _conn <- Some(conn)
            conn.Open()
        else
            failwith "Unable to connect to database."
    let OpenOrCreate() =
        if System.IO.File.Exists(filename) then
            OpenDatabase()
        else
            CreateDatabase()
    do
        OpenOrCreate()
    member self.Command() : SQLiteCommand =
        match _conn with
        | Some(c) -> new SQLiteCommand(c)
        | None -> failwith "Unable to connect to database."
    member self.AddFile(mturkfilename: string, benchmarkfilename: string) : int =
        let cmd = self.Command()
        let querytxt = "INSERT INTO files (mturkfilename, benchmarkfilename) VALUES (\"" + mturkfilename + "\", \"" + benchmarkfilename + "\")"
        cmd.CommandText <- querytxt
        if cmd.ExecuteNonQuery() <> 1 then
            failwith ("INSERT failed: " + querytxt)
        cmd.CommandText <- "SELECT LAST_INSERT_ROWID();"
        let row_id = System.Convert.ToInt32(cmd.ExecuteScalar())
        cmd.CommandText <- "SELECT id FROM files WHERE ROWID = " + row_id.ToString() + ";"
        System.Convert.ToInt32(cmd.ExecuteScalar())
    member self.AddHIT(fileid: int,
                       hitid: string,
                       hittypeid: string,
                       title: string,
                       description: string,
                       keywords: string,
                       reward: decimal,
                       creationtime: string,
                       maxassignments: int,
                       requesterannotation: string,
                       assignmentdurationinseconds: int,
                       autoapprovaldelayinseconds: int,
                       expiration: string,
                       numberofsimilarhits: int,
                       lifetimeinseconds: int,
                       assignmentid: string,
                       workerid: string,
                       assignmentstatus: string,
                       accepttime: string,
                       submittime: string,
                       autoapprovaltime: string,
                       approvaltime: string,
                       rejectiontime: string,
                       requesterfeedback: string,
                       worktimeinseconds: int,
                       lifetimeapprovalrate: string,
                       last30daysapprovalrate: string,
                       last7daysapprovalrate: string
                       ) : int =
        let cmd = self.Command()
        let querystr = "INSERT INTO hits ( hitid, hittypeid, title, description," +
                                         " keywords, reward, creationtime, maxassignments," +
                                         " requesterannotation, assignmentdurationinseconds, " +
                                         " autoapprovaldelayinseconds, expiration, numberofsimilarhits," +
                                         " lifetimeinseconds, assignmentid, workerid," +
                                         " assignmentstatus, accepttime, submittime," +
                                         " autoapprovaltime, approvaltime, rejectiontime, requesterfeedback," +
                                         " worktimeinseconds, lifetimeapprovalrate, last30daysapprovalrate," +
                                         " last7daysapprovalrate,fileid )"
        let queryval1 = " VALUES (\"" + hitid + "\",\"" + hittypeid + "\",\"" + title + "\",\"" + description + "\",\""
                                    + keywords + "\"," + reward.ToString() + "," + TurkTimeToTimestamp(creationtime) + "," + maxassignments.ToString() + ",\""
                                    + requesterannotation + "\"," + assignmentdurationinseconds.ToString() + ","
        let queryval2 =               autoapprovaldelayinseconds.ToString() + "," + TurkTimeToTimestamp(expiration) + "," + numberofsimilarhits.ToString() + ","
                                    + lifetimeinseconds.ToString() + ",\"" + assignmentid + "\",\"" + workerid + "\",\""
                                    + assignmentstatus + "\"," + TurkTimeToTimestamp(accepttime) + "," + TurkTimeToTimestamp(submittime) + ","
                                    + TurkTimeToTimestamp(autoapprovaltime) + "," + TurkTimeToTimestamp(approvaltime) + "," + TurkTimeToTimestamp(rejectiontime) + ",\"" + requesterfeedback + "\","
                                    + worktimeinseconds.ToString() + ",\"" + lifetimeapprovalrate + "\",\"" + last30daysapprovalrate + "\",\"" + last7daysapprovalrate + "\","
                                    + fileid.ToString() + ")"
        let querytxt = querystr + queryval1 + queryval2
        cmd.CommandText <- querytxt
        if cmd.ExecuteNonQuery() <> 1 then
            failwith ("INSERT failed: " + querytxt)

        // return the id
        cmd.CommandText <- "SELECT LAST_INSERT_ROWID();"
        let row_id = System.Convert.ToInt32(cmd.ExecuteScalar())
        cmd.CommandText <- "SELECT id FROM hits WHERE ROWID = " + row_id.ToString() + ";"
        System.Convert.ToInt32(cmd.ExecuteScalar())
    member self.AddAnswerWithErrors(cell: int, origdata: string, userdata: string, is_different: bool, hitid: int, errors: ErrorType list) : int =
        if Connected() then
            let cmd = self.Command()
            let is_diff = if is_different then 1 else 0
            let querystr = "INSERT INTO answers (cell, origdata, userdata, is_different, hitid) VALUES (" + cell.ToString() + ",\"" + origdata + "\",\"" + userdata + "\"," + is_diff.ToString() + "," + hitid.ToString() + ")"
            cmd.CommandText <- querystr
            if cmd.ExecuteNonQuery() <> 1 then
                failwith ("INSERT failed: " + querystr)

            // return the id
            cmd.CommandText <- "SELECT LAST_INSERT_ROWID();"
            let row_id = System.Convert.ToInt32(cmd.ExecuteScalar())
            cmd.CommandText <- "SELECT id FROM answers WHERE ROWID = " + row_id.ToString() + ";"
            let answer_id = System.Convert.ToInt32(cmd.ExecuteScalar())

            // add all of the error classifications
            for error in errors do
                let querystr = "INSERT INTO answers_errors (answerid, errortypeid) VALUES (" + answer_id.ToString() + "," + (int error).ToString() + ")"
                cmd.CommandText <- querystr
                if cmd.ExecuteNonQuery() <> 1 then
                    failwith ("INSERT failed: " + querystr)

            // return answer id
            answer_id
        else
            failwith "Must be connected to a database."
    member self.ImportCSV(csvfile: string) =
        let csv = File.ReadAllText(csvfile)
        let rows = CSVParser.ParseCsv csv ","
        let wbname = rows.[1].[27]  // wbname is always here for DD experiments
        let mutable count = 0
        let fileid = self.AddFile(wbname, csvfile)
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
                let last7daysapprovalrate = row.[26]
                let hitid = self.AddHIT(fileid,
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
                                      last30daysapprovalrate,
                                      last7daysapprovalrate)
                for i in 0..9 do
                    let usertxt = row.[39 + i]
                    let origtxt = row.[29 + i]

                    // classify errors
                    let mutable errors = []
                    if ErrorClassifiers.TestSignError(usertxt, origtxt) then
                        errors <- ErrorType.SignError :: errors

                    // insert answer with errors into DB
                    self.AddAnswerWithErrors(hitid, origtxt, usertxt, not(usertxt.Equals(origtxt)), hitid, errors) |> ignore
            count <- count + 1