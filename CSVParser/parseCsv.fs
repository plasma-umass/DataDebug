module CSVParser

// taken from: https://gist.github.com/jb55/897600
open FParsec.Primitives
open FParsec.CharParsers

type CsvResult = { IsSuccess : bool; ErrorMsg : string; Result : string list list }

// Some simple helpers
let isWs s = Seq.map (isAnyOf "\t ") s
          |> Seq.reduce (&&)

let ws = spaces
let chr c = skipChar c
let st str = if isWs str then skipString str else skipString str .>> ws
let ch c = chr c .>> ws

let escapeChar = chr '\\' >>. anyOf "\"\\/bfnrt," 
                   |>> function
                     | 'b' -> '\b'
                     | 'f' -> '\u000C'
                     | 'n' -> '\n'
                     | 'r' -> '\r'
                     | 't' -> '\t'
                     | c   -> c


let nonQuotedCellChar delim = escapeChar <|> (noneOf (delim + "\r\n"))
let cellChar = escapeChar <|> (noneOf "\"")

// Trys to parse quoted cells, if that fails try to parse non-quoted cells
let cell delim = between (chr '\"') (chr '\"') (manyChars cellChar) <|> manyChars (nonQuotedCellChar delim)

// Cells are delimited by  a specified string
let row delim = sepBy (cell delim) (st delim)

// Rows are delimited by newlines
let csv delim = sepBy (row delim) newline .>> eof
let commaCsv = csv ","

let stripEmpty ls = List.filter (fun (row:'a list) -> row.Length <> 0) ls

let ParseCsv s delim: string[][]  =
  let res = run (csv delim) s in
    match res with
     | Success (rows, _, _) -> Array.ofList (List.map (fun row -> Array.ofList row) (stripEmpty rows))
     | Failure (s, _, _) -> failwith "Could not parse CSV."

// Misc utils
let detectDelimination s choices =
  let countChars ch = Seq.filter (fun c -> c = ch) s |> Seq.length
  let countChoices  = Seq.zip choices (Seq.map countChars choices)
   in Seq.sortBy (fun tup -> -(snd tup)) countChoices |> Seq.head |> fst

let DetectTabOrComma (s:string) = detectDelimination s "\t,"
let ParseTabOrComma s = ParseCsv s ((DetectTabOrComma s).ToString())
let ZeroOrNum s = match s with
                  | "" -> 0
                  | _ -> System.Convert.ToInt32(s)