module DataAnalysis
open System
open System.Data.SQLite

type MTurkData(filename: string) =
    let mutable _conn: SQLiteConnection = null
    member self.CreateDatabase(dbfilename: string) =
        _conn <- new SQLiteConnection("data source=\"" + filename + "\"")
        _conn.Open()

        let cmd = new SQLiteCommand(_conn)

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
    member self.OpenDatabase(dbfilename: string) =
        _conn <- new SQLiteConnection("data source=\"" + filename + "\"")
    member self.Connected =
        if _conn = null then
            false
        else
            true

    member self.AddFile(mturkfilename: string, benchmarkfilename: string) =
        if self.Connected then
            let cmd = new SQLiteCommand(_conn)
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
                       creationtime: DateTime,
                       maxassignments: int,
                       requesterannotation: string,
                       assignmentdurationinseconds: int,
                       autoapprovaldelayinseconds: int,
                       expiration: int,
                       numberofsimilarhits: int,
                       lifetimeinseconds: int,
                       assignmentstatus: string,
                       accepttime: DateTime,
                       submittime: DateTime,
                       autoapprovaltime: DateTime,
                       approvaltime: DateTime,
                       rejectiontime: DateTime,
                       requesterfeedback: string,
                       worktimeinseconds: int,
                       lifetimeapprovalrate: string,
                       last30daysapprovalrate: string
                       ) =
        if self.Connected then
            let cmd = new SQLiteCommand(_conn)
            let querystr = "INSERT INTO files ( hitid, hittypeid, title, description," +
                                              " keywords, reward, creationtime, maxassignments," +
                                              " requesterannotation, assignmentdurationinseconds, " +
                                              " autoapprovaldelayinseconds, expiration, numberofsimilarhits," +
                                              " lifetimeinseconds, assignmentstatus, accepttime, submittime," +
                                              " autoapprovaltime, approvaltime, rejectiontime, requesterfeedback," +
                                              " worktimeinseconds, lifetimeapprovalrate, last30daysapprovalrate )"
            let queryval = " VALUES (" + hitid + "," + hittypeid + "," + title + "," + description + ","
                                       + keywords + "," + reward.ToString() + "," + MTurkData.ToTimestamp(creationtime) + "," + maxassignments.ToString() + ","
                                       + requesterannotation + "," + assignmentdurationinseconds.ToString() + ","
                                       + autoapprovaldelayinseconds.ToString() + "," + expiration.ToString() + "," + numberofsimilarhits.ToString() + ","
                                       + lifetimeinseconds.ToString() + "," + assignmentstatus + "," + MTurkData.ToTimestamp(accepttime) + "," + MTurkData.ToTimestamp(submittime) + ","
                                       + MTurkData.ToTimestamp(autoapprovaltime) + "," + MTurkData.ToTimestamp(approvaltime) + "," + MTurkData.ToTimestamp(rejectiontime) + "," + requesterfeedback + ","
                                       + worktimeinseconds.ToString() + "," + lifetimeapprovalrate + "," + last30daysapprovalrate
                                       + ")"
            cmd.CommandText <- querystr + queryval
            if cmd.ExecuteNonQuery() <> 1 then
                failwith ("INSERT failed: " + querystr + queryval)
        else
            failwith "Must be connected to a database."
    static member ToTimestamp(dt: DateTime) : string =
        (dt.Ticks / Convert.ToInt64("10000000") - Convert.ToInt64("62136892800")).ToString()

//        SQLiteCommand cmd = new SQLiteCommand(conn);
//        cmd.CommandText = "select * from Customer";
//
//        SQLiteDataReader reader = cmd.ExecuteReader( );
//        while (reader.Read( ))
//        {
//                    // do something
//        }
//        reader.Close( );
//
//        cmd.CommandText = "delete from Customer where CustomerID = 33";
//        cmd.ExecuteScalar( );
//
//        conn.Close( );
//
//        cmd.CommandText <- "INSERT INTO test (foo) VALUES (1)"
//        cmd.ExecuteNonQuery() |> ignore
//        cmd.CommandText <- "INSERT INTO test (foo) VALUES (2)"
//        cmd.ExecuteNonQuery() |> ignore
//        cmd.CommandText <- "INSERT INTO test (foo) VALUES (3)"
//        cmd.ExecuteNonQuery() |> ignore
//        cmd.CommandText <- "INSERT INTO test (foo) VALUES (4)"
//        cmd.ExecuteNonQuery() |> ignore
//
//        cmd.CommandText <- "SELECT * FROM test"
//        let reader = cmd.ExecuteReader()
//        while reader.Read() do
//            data <- reader.["foo"] :: data
//        reader.Close()
//
//        conn.Close()
//    member self.GetData = String.Join("\n", data)