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
        | _ -> false

    let rec GetRangeReferenceRanges(ref: AST.ReferenceRange) : AST.Range list = [ref.Range]

    and GetFunctionRanges(ref: AST.ReferenceFunction) : AST.Range list =
        if PuntedFunction(ref.FunctionName) then
            []
        else
            List.map (fun arg -> GetRanges(arg)) ref.ArgumentList |> List.concat
        
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
        | _ -> failwith "Unknown reference type."

    let GetReferencesFromFormula(formula: string, wb: Workbook, ws: Worksheet) : seq<XLRange> =
        let app = wb.Application
        match ExcelParser.ParseFormula(formula, wb, ws) with
        | Some(tree) ->
            let refs = GetExprRanges(tree)
            List.map (fun (r: AST.Range) -> r.GetCOMObject(wb.Application)) refs |> Seq.ofList
        | None -> [] |> Seq.ofList