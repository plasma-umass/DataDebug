module DataAnalysis
open System
open System.Data.SQLite
open System.Globalization

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
        let formatstring = "ddd MMM dd HH:mm:ss 'PDT' yyyy"
        let dt = DateTime.ParseExact(datestring, formatstring, CultureInfo.InvariantCulture)
        let pdt = TimeZoneInfo.CreateCustomTimeZone("Pacific Daylight Time", new TimeSpan(-07, 00, 00), "(UTC-07:00) Pacific Daylight Time", "Pacific Daylight Time")
        let utc = TimeZoneInfo.Utc
        TimeZoneInfo.ConvertTime(dt, pdt, utc)
    let TurkTimeToTimestamp(datestring: string) =
        ToTimestamp(FromMTurkTime(datestring))
    let Connected =
        match _conn with
        | Some(c) -> true
        | None -> false
    override self.Finalize() =
        match _conn with
        | Some(c) -> c.Close()
        | None -> ()
    member self.CreateDatabase(dbfilename: string) =
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
                                             " FOREIGN KEY(fileid) REFERENCES files(id)" +
                                             ")"
        cmd.ExecuteNonQuery() |> ignore

        cmd.CommandText <- "CREATE TABLE answers ( id INTEGER PRIMARY KEY AUTOINCREMENT," +
                                                 " cell NUMERIC," +
                                                 " data TEXT," +
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
                            (int ErrorType.DigitTransposition).ToString() + "," +
                            ErrorType.DigitTransposition.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.ExtraDigit).ToString() + "," +
                            ErrorType.ExtraDigit.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.OmittedDigit).ToString() + "," +
                            ErrorType.OmittedDigit.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.SignError).ToString() + "," +
                            ErrorType.SignError.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.MantissaError).ToString() + "," +
                            ErrorType.MantissaError.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
        cmd.CommandText <- "INSERT INTO errortypes (id, name) VALUES (" +
                            (int ErrorType.Other).ToString() + "," +
                            ErrorType.Other.ToString() + ")"
        cmd.ExecuteNonQuery() |> ignore
    member self.OpenDatabase(dbfilename: string) =
        let conn = new SQLiteConnection("data source=\"" + filename + "\"")
        if (conn <> null) then
            _conn <- Some(conn)
        else
            failwith "Unable to connect to database."
    member self.Command =
        match _conn with
        | Some(c) -> new SQLiteCommand(c)
        | None -> failwith "Unable to connect to database."
    member self.AddFile(mturkfilename: string, benchmarkfilename: string) =
        if Connected then
            let cmd = self.Command
            let querytxt = "INSERT INTO files (mturkfilename, benchmarkfilename) VALUES (" + mturkfilename + ", " + benchmarkfilename + ")"
            cmd.CommandText <- querytxt
            if cmd.ExecuteNonQuery() <> 1 then
                failwith ("INSERT failed: " + querytxt)
        else
            failwith "Must be connected to a database."
    member self.AddHIT(hitid: string,
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
                       expiration: int,
                       numberofsimilarhits: int,
                       lifetimeinseconds: int,
                       assignmentstatus: string,
                       accepttime: string,
                       submittime: string,
                       autoapprovaltime: string,
                       approvaltime: string,
                       rejectiontime: string,
                       requesterfeedback: string,
                       worktimeinseconds: int,
                       lifetimeapprovalrate: string,
                       last30daysapprovalrate: string
                       ) : int =
        if Connected then
            let cmd = self.Command
            let querystr = "INSERT INTO files ( hitid, hittypeid, title, description," +
                                              " keywords, reward, creationtime, maxassignments," +
                                              " requesterannotation, assignmentdurationinseconds, " +
                                              " autoapprovaldelayinseconds, expiration, numberofsimilarhits," +
                                              " lifetimeinseconds, assignmentstatus, accepttime, submittime," +
                                              " autoapprovaltime, approvaltime, rejectiontime, requesterfeedback," +
                                              " worktimeinseconds, lifetimeapprovalrate, last30daysapprovalrate )"
            let queryval = " VALUES (" + hitid + "," + hittypeid + "," + title + "," + description + ","
                                       + keywords + "," + reward.ToString() + "," + TurkTimeToTimestamp(creationtime) + "," + maxassignments.ToString() + ","
                                       + requesterannotation + "," + assignmentdurationinseconds.ToString() + ","
                                       + autoapprovaldelayinseconds.ToString() + "," + expiration.ToString() + "," + numberofsimilarhits.ToString() + ","
                                       + lifetimeinseconds.ToString() + "," + assignmentstatus + "," + TurkTimeToTimestamp(accepttime) + "," + TurkTimeToTimestamp(submittime) + ","
                                       + TurkTimeToTimestamp(autoapprovaltime) + "," + TurkTimeToTimestamp(approvaltime) + "," + TurkTimeToTimestamp(rejectiontime) + "," + requesterfeedback + ","
                                       + worktimeinseconds.ToString() + "," + lifetimeapprovalrate + "," + last30daysapprovalrate
                                       + ")"
            cmd.CommandText <- querystr + queryval
            if cmd.ExecuteNonQuery() <> 1 then
                failwith ("INSERT failed: " + querystr + queryval)

            // return the id
            cmd.CommandText <- "SELECT LAST_INSERT_ROWID();"
            System.Convert.ToInt32(cmd.ExecuteScalar())
        else
            failwith "Must be connected to a database."
    member self.AddAnswerWithErrors(cell: int, data: string, hitid: int, errors: ErrorType list) : int =
        if Connected then
            let cmd = self.Command
            let querystr = "INSERT INTO answers (cell, data, hitid) VALUES (" + cell.ToString() + "," + data + "," + hitid.ToString() + ")"
            if cmd.ExecuteNonQuery() <> 1 then
                failwith ("INSERT failed: " + querystr)

            // return the id
            cmd.CommandText <- "SELECT LAST_INSERT_ROWID();"
            let answer_id = System.Convert.ToInt32(cmd.ExecuteScalar())

            // add all of the error classifications
            for error in errors do
                let querystr = "INSERT INTO answers_errors (answerid, errortypeid) VALUES (" + answer_id.ToString() + "," + error.ToString() + ")"
                cmd.CommandText <- querystr
                if cmd.ExecuteNonQuery() <> 1 then
                    failwith ("INSERT failed: " + querystr)

            // return answer id
            answer_id
        else
            failwith "Must be connected to a database."