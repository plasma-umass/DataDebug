module ExcelParser
    open FParsec
    open AST
    open Microsoft.Office.Interop.Excel

    type Workbook = Microsoft.Office.Interop.Excel.Workbook
    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type UserState = unit
    type Parser<'t> = Parser<'t, UserState>

    // Grammar forward references
    let ArgumentList, ArgumentListImpl = createParserForwardedToRef()
    let ExpressionSimple, ExpressionSimpleImpl = createParserForwardedToRef()
    let ExpressionDecl, ExpressionDeclImpl = createParserForwardedToRef()

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
    let Workbook = ((attempt WorkbookName) |>> Some) <|> ((pstring "") >>% None)

    // References
    // References consist of the following parts:
    //   An optional workbook name prefix
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

    let ConstantReference = pint32 |>> (fun r -> ReferenceConstant(None, r) :> Reference)

    let ReferenceKinds = (attempt RangeReference) <|> (attempt AddressReference) <|> (attempt ConstantReference) <|> NamedReference
    let Reference = pipe2 Workbook ReferenceKinds (fun wbname ref -> ref.WorkbookName <- wbname; ref)

    // Functions
    let FunctionName = many1Satisfy (fun c -> isLetter(c))
    let Function = pipe2 (FunctionName .>> pstring "(") (ArgumentList .>> pstring ")") (fun fname arglist -> ReferenceFunction(None, fname, arglist) :> Reference)
    let Argument = (attempt Function) <|> Reference
    do ArgumentListImpl := sepBy Reference (pstring ",")

    // Binary arithmetic operators
    let BinOpChar = satisfy (fun c -> c = '+' || c = '-' || c = '/' || c = '*')
    let BinOp = pipe2 BinOpChar ExpressionDecl (fun op rhs -> (op, rhs))

    // Unary operators
    let UnaryOpChar = satisfy (fun c -> c = '+' || c = '-')

    // Expressions
    let ParensExpr = (between (pstring "(") (pstring ")") ExpressionDecl) |>> ParensExpr
    let ExpressionAtom = ((attempt Function) <|> Reference) |>> ReferenceExpr
    do ExpressionSimpleImpl := ExpressionAtom <|> ParensExpr
    let UnaryOpExpr = pipe2 UnaryOpChar ExpressionDecl (fun op rhs -> UnaryOpExpr(op, rhs))
    let BinOpExpr = pipe2 ExpressionSimple BinOp (fun lhs (op, rhs) -> BinOpExpr(op, lhs, rhs))
    do ExpressionDeclImpl := (attempt UnaryOpExpr) <|> (attempt BinOpExpr) <|> (attempt ExpressionSimple)

    // Formulas
    let Formula = pstring "=" >>. ExpressionDecl .>> eof

    // Resolve all undefined references to the current worksheet and workbook
    let RefAddrResolve(ref: Reference)(wb: Workbook)(ws: Worksheet) = ref.Resolve wb ws
    let rec ExprAddrResolve(expr: Expression)(wb: Workbook)(ws: Worksheet) =
        match expr with
        | ReferenceExpr(r) ->
            RefAddrResolve r wb ws
        | BinOpExpr(op,e1,e2) ->
            ExprAddrResolve e1 wb ws
            ExprAddrResolve e2 wb ws
        | UnaryOpExpr(op, e) ->
            ExprAddrResolve e wb ws
        | ParensExpr(e) ->
            ExprAddrResolve e wb ws

    // monadic wrapper for success/failure
    let test p str =
        match run p str with
        | Success(result, _, _)   -> printfn "Success: %A" result
        | Failure(errorMsg, _, _) -> printfn "Failure: %s" errorMsg

    let GetAddress(str: string, wb: Workbook, ws: Worksheet): Address =
        match run (AddrR1C1 .>> eof) str with
        | Success(addr, _, _) ->
            addr.WorkbookName <- Some wb.Name
            addr.WorksheetName <- Some ws.Name
            addr
        | Failure(errorMsg, _, _) -> failwith errorMsg

    let GetRange str ws: AST.Range option =
        match run (RangeR1C1 .>> eof) str with
        | Success(range, _, _) -> Some(range)
        | Failure(errorMsg, _, _) -> None

    let GetReference str wb ws: Reference option =
        match run (Reference .>> eof) str with
        | Success(reference, _, _) ->
            RefAddrResolve reference wb ws
            Some(reference)
        | Failure(errorMsg, _, _) -> None

    let ParseFormula(str, wb, ws): Expression option =
        match run Formula str with
        | Success(formula, _, _) ->
            ExprAddrResolve formula wb ws
            Some(formula)
        | Failure(errorMsg, _, _) -> None

    // The parser REPL calls this; note that the
    // Formula parser looks for EOF
    let ConsoleTest(s: string) = test Formula s