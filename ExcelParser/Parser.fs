module ExcelParser
    open FParsec
    open AST

    type UserState = unit
    type Parser<'t> = Parser<'t, UserState>

    // Addresses
    // We treat relative and absolute addresses the same-- they behave
    // exactly the same way unless you copy and paste them.
    let AddrR = pstring "R" >>. pint32
    let AddrC = pstring "C" >>. pint32
    let AddrR1C1 = pipe2 AddrR AddrC (fun r c -> Address(r,c,None,None))
    let AddrA = many1Satisfy isAsciiUpper
    let AddrAAbs = (pstring "$" <|> pstring "") >>. AddrA
    let Addr1 = pint32
    let Addr1Abs = (pstring "$" <|> pstring "") >>. Addr1
    let AddrA1 = pipe2 AddrAAbs Addr1Abs (fun col row -> Address(row,col,None,None))
    let AnyAddr = (attempt AddrR1C1) <|> AddrA1

    // Ranges
    let MoreAddrR1C1 = pstring ":" >>. AddrR1C1
    let RangeR1C1 = pipe2 AddrR1C1 MoreAddrR1C1 (fun r1 r2 -> Range(r1, r2))
    let MoreAddrA1 = pstring ":" >>. AddrA1
    let RangeA1 = pipe2 AddrA1 MoreAddrA1 (fun r1 r2 -> Range(r1, r2))
    let RangeAny = (attempt RangeR1C1) <|> RangeA1

    // Worksheet Names
    let WorksheetNameQuoted = between (pstring "'") (pstring "'") (many1Satisfy ((<>) '\''))
    let WorksheetNameUnquoted = many1Satisfy (fun c -> (isDigit c) || (isLetter c))
    let WorksheetName = WorksheetNameQuoted <|> WorksheetNameUnquoted

    // Workbook Names (this may be too restrictive)
    let WorkbookName = between (pstring "[") (pstring "]") (many1Satisfy (fun c -> c <> '[' && c <> ']'))
    let Workbook = (attempt WorkbookName) <|> (pstring "")

    // References
    // References consist of the following parts:
    //   An optional worksheet name prefix
    //   A single-cell ("Address") or multi-cell address ("Range")
    let RangeReferenceWorksheet = pipe2 (WorksheetName .>> pstring "!") RangeAny (fun wsname rng -> ReferenceRange(Some(wsname), rng) :> Reference)
    let RangeReferenceNoWorksheet = RangeAny |>> (fun rng -> ReferenceRange(None, rng) :> Reference)
    let RangeReference = (attempt RangeReferenceWorksheet) <|> RangeReferenceNoWorksheet

    let AddressReferenceWorksheet = pipe2 (WorksheetName .>> pstring "!") AnyAddr (fun wsname addr -> ReferenceAddress(Some(wsname), addr) :> Reference)
    let AddressReferenceNoWorksheet = AnyAddr |>> (fun addr -> ReferenceAddress(None, addr) :> Reference)
    let AddressReference = (attempt AddressReferenceWorksheet) <|> AddressReferenceNoWorksheet

    let NamedReferenceFirstChar = satisfy (fun c -> c = '_' || isLetter(c))
    let NamedReferenceLastChars = manySatisfy (fun c -> c = '_' || isLetter(c) || isDigit(c))
    let NamedReference = pipe2 NamedReferenceFirstChar NamedReferenceLastChars (fun c s -> ReferenceNamed(None, c.ToString() + s) :> Reference)

    let ReferenceKinds  = (attempt RangeReference) <|> (attempt AddressReference) <|> NamedReference
    let Reference = pipe2 Workbook ReferenceKinds (fun wbname ref -> ref.WorkbookName <- Some wbname; ref)

    // Functions
    let ArgumentList, ArgumentListImpl = createParserForwardedToRef()
    let FunctionName = many1Satisfy (fun c -> c <> '(' && c <> ')')
    let Function = pipe2 (FunctionName .>> pstring "(") (ArgumentList .>> pstring ")") (fun fname arglist -> ReferenceFunction(None, fname, arglist) :> Reference)
    let Argument = (attempt Function) <|> Reference
    do ArgumentListImpl := sepBy Reference (pstring ",")

    // Expressions
    let Expression = (attempt Function) <|> Reference

    // Formulas
    let Formula = pstring "=" >>. Expression .>> eof

    // monadic wrapper for success/failure
    let test p str =
        match run p str with
        | Success(result, _, _)   -> printfn "Success: %A" result
        | Failure(errorMsg, _, _) -> printfn "Failure: %s" errorMsg

    let GetAddress str ws: Address =
        match run (AddrR1C1 .>> eof) str with
        | Success(addr, _, _) -> addr
        | Failure(errorMsg, _, _) -> failwith errorMsg

    let GetRange str ws: Range option =
        match run (RangeR1C1 .>> eof) str with
        | Success(range, _, _) -> Some(range)
        | Failure(errorMsg, _, _) -> None

    let GetReference str ws: Reference option =
        match run (Reference .>> eof) str with
        | Success(reference, _, _) -> Some(reference)
        | Failure(errorMsg, _, _) -> None

    // helper function for mortals to comprehend; note that Formula looks for EOF
    let ConsoleTest(s: string) = test Formula s