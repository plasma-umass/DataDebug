module DataAnalysis

open System.Data.SQLite

type MTurkData(filename: string) =
    do
        let conn = new SQLiteConnection(filename)
        conn.Open()

        let cmd = new SQLiteCommand(conn)
        cmd.CommandText <- "CREATE TABLE test (foo INTEGER)"
        
        cmd.ExecuteNonQuery() |> ignore

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