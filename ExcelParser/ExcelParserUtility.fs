module ExcelParserUtility
    type Workbook = Microsoft.Office.Interop.Excel.Workbook
    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type XLRange = Microsoft.Office.Interop.Excel.Range

    let PuntedFunction(fnname: string) : bool =
        match fnname with
        | "INDEX" -> true
        | "HLOOKUP" -> true
        | "VLOOKUP" -> true
        | "LOOKUP" -> true
        | "OFFSET" -> true
        | _ -> false

    let rec GetRangeReferenceRanges(ref: AST.ReferenceRange) : AST.Range list = [ref.Range]

    and GetFunctionRanges(ref: AST.ReferenceFunction) : AST.Range list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetExprRanges(arg)) ref.ArgumentList |> List.concat
        
    and GetExprRanges(expr: AST.Expression) : AST.Range list =
        match expr with
        | AST.ReferenceExpr(r) -> GetRanges(r)
        | AST.BinOpExpr(op, e1, e2) -> GetExprRanges(e1) @ GetExprRanges(e2)
        | AST.UnaryOpExpr(op, e) -> GetExprRanges(e)
        | AST.ParensExpr(e) -> GetExprRanges(e)

    and GetRanges(ref: AST.Reference) : AST.Range list =
        match ref with
        | :? AST.ReferenceRange as r -> GetRangeReferenceRanges(r)
        | :? AST.ReferenceAddress -> []
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> GetFunctionRanges(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    let GetReferencesFromFormula(formula: string, wb: Workbook, ws: Worksheet) : seq<XLRange> =
        let wbpath = wb.FullName
        let app = wb.Application
        match ExcelParser.ParseFormula(formula, wbpath, wb, ws) with
        | Some(tree) ->
            let refs = GetExprRanges(tree)
            List.map (fun (r: AST.Range) ->
                        // temporarily bail if r refers to object in a different workbook
                        // TODO: we should open the other workbook and continue the analysis
                        let paths = Seq.filter (fun pth -> pth = wbpath) (r.GetPathNames())
                        if Seq.length(paths) = 0 then
                            None
                        else
                            Some(r.GetCOMObject(wb.Application))
                     ) refs |> List.choose id |> Seq.ofList
        | None -> [] |> Seq.ofList

    // single-cell variants:

    let rec GetSCExprRanges(expr: AST.Expression) : AST.Address list =
        match expr with
        | AST.ReferenceExpr(r) -> GetSCRanges(r)
        | AST.BinOpExpr(op, e1, e2) -> GetSCExprRanges(e1) @ GetSCExprRanges(e2)
        | AST.UnaryOpExpr(op, e) -> GetSCExprRanges(e)
        | AST.ParensExpr(e) -> GetSCExprRanges(e)

    and GetSCRanges(ref: AST.Reference) : AST.Address list =
        match ref with
        | :? AST.ReferenceRange -> []
        | :? AST.ReferenceAddress as r -> GetSCAddressReferenceRanges(r)
        | :? AST.ReferenceNamed -> []   // TODO: symbol table lookup
        | :? AST.ReferenceFunction as r -> GetSCFunctionRanges(r)
        | :? AST.ReferenceConstant -> []
        | :? AST.ReferenceString -> []
        | _ -> failwith "Unknown reference type."

    and GetSCAddressReferenceRanges(ref: AST.ReferenceAddress) : AST.Address list = [ref.Address]

    and GetSCFunctionRanges(ref: AST.ReferenceFunction) : AST.Address list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetSCExprRanges(arg)) ref.ArgumentList |> List.concat

    let rec GetFormulaNamesFromExpr(ast: AST.Expression): string list =
        match ast with
        | AST.ReferenceExpr(r) -> GetFormulaNamesFromReference(r)
        | AST.BinOpExpr(op, e1, e2) -> GetFormulaNamesFromExpr(e1) @ GetFormulaNamesFromExpr(e2)
        | AST.UnaryOpExpr(op, e) -> GetFormulaNamesFromExpr(e)
        | AST.ParensExpr(e) -> GetFormulaNamesFromExpr(e)

    and GetFormulaNamesFromReference(ref: AST.Reference): string list =
        match ref with
        | :? AST.ReferenceFunction as r -> [r.FunctionName]
        | _ -> []

    let GetSCFormulaNames(formula: string, path: string, ws: Worksheet, wb: Workbook) =
        let app = wb.Application
        match ExcelParser.ParseFormula(formula, path, wb, ws) with
        | Some(ast) -> GetFormulaNamesFromExpr(ast) |> Seq.ofList
        | None -> [] |> Seq.ofList

    let GetSingleCellReferencesFromFormula(formula: string, wb: Workbook, ws: Worksheet) : seq<AST.Address> =
        let app = wb.Application
        let path = wb.FullName
        match ExcelParser.ParseFormula(formula, path, wb, ws) with
        | Some(tree) -> GetSCExprRanges(tree) |> Seq.ofList
        | None -> [] |> Seq.ofList