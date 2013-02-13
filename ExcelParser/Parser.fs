module ExcelParser
    open FParsec
    open AST

    type UserState = unit
    type Parser<'t> = Parser<'t, UserState>

    let AddrRC: Parser<_> = (pstring "R" <|> pstring "C") >>. pint32
    let AddrR1C1: Parser<_> = pipe2 AddrRC AddrRC (fun r c -> Address(r,c))
    let MoreAddr: Parser<_> = (pstring ":" >>. AddrR1C1) |>> Some
    let NoMoreAddr: Parser<Address option> = pstring "" >>% None
    let RangeR1C1: Parser<_> = pipe2 AddrR1C1 (MoreAddr <|> NoMoreAddr) (fun r1 r2 -> Range(r1, r2))

    // monadic wrapper for success/failure
    let test p str =
        match run p str with
        | Success(result, _, _)   -> printfn "Success: %A" result
        | Failure(errorMsg, _, _) -> printfn "Failure: %s" errorMsg

    let GetRange str: Range option =
        match run RangeR1C1 str with
        | Success(range, _, _) -> Some(range)
        | Failure(errorMsg, _, _) -> None

    // helper function for mortals to comprehend
    let ConsoleTest(s: string) = test RangeR1C1 s