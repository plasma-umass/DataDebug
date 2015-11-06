namespace ParcelCOMShim
    open System
    open Microsoft.Office.Interop.Excel
    open Parcel

    type XLRange = Microsoft.Office.Interop.Excel.Range

    type COMRef(unique_id: string,
                wb: Workbook,
                ws: Worksheet,
                range: XLRange,
                path: string,           // path excluding final separator and filename; option type because in-memory workbooks have no path
                workbook_name: string,
                worksheet_name: string,
                formula: string option,
                width: int,
                height: int) =
        let _wb = wb
        let _ws = ws
        let _r = range
        let _is_cell = width = 1 && height = 1
        let _interned_unique_id = String.Intern(unique_id)
        let _width = width
        let _height = height
        let _path = path
        let _workbook_name = workbook_name
        let _worksheet_name = worksheet_name
        let _formula = formula
        let mutable _do_not_perturb = match formula with | Some(f) -> true | None -> false

        member self.Width = _width
        member self.Height = _height
        member self.Workbook = _wb
        member self.Worksheet = _ws
        member self.Range = _r
        member self.IsFormula = match _formula with | Some(f) -> true | None -> false
        member self.Formula = match _formula with
            | Some(f) -> f
            | None -> failwith "Not a formula reference."
        member self.IsCell = _is_cell
        member self.UniqueID = _interned_unique_id
        member self.Path = _path
        member self.WorkbookName = _workbook_name
        member self.WorksheetName = _worksheet_name
        member self.DoNotPerturb
            with get() = _do_not_perturb
            and set(value) = _do_not_perturb <- value
        override self.GetHashCode() = _interned_unique_id.GetHashCode()

    module Address =
        let GetCOMObject(addr: AST.Address, app: Application) : XLRange =
            let wb: Workbook = app.Workbooks.Item(addr.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(addr.A1Worksheet()) :?> Worksheet
            let cell: XLRange = ws.Range(addr.A1Local())
            cell

        let AddressFromCOMObject(com: Microsoft.Office.Interop.Excel.Range, wb: Microsoft.Office.Interop.Excel.Workbook) : AST.Address =
            let wsname = com.Worksheet.Name
            let wbname = wb.Name
            let path = System.IO.Path.GetDirectoryName(wb.FullName)
            let addr = com.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, Type.Missing, Type.Missing)
            AST.Address.FromString(addr, wsname, wbname, path)

    module Range =
        let GetCOMObject(rng: AST.Range, app: Application) : XLRange =
            // tl and br must share workbook and worksheet (I think)
            let wb: Workbook = app.Workbooks.Item(rng.TopLeft.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(rng.TopLeft.A1Worksheet()) :?> Worksheet
            let range: XLRange = ws.Range(rng.TopLeft.A1Local(), rng.BottomRight.A1Local())
            range