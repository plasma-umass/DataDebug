module AST
    open System
    open System.Diagnostics
    open Microsoft.Office.Interop.Excel

    type Application = Microsoft.Office.Interop.Excel.Application
    type Workbook = Microsoft.Office.Interop.Excel.Workbook
    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type XLRange = Microsoft.Office.Interop.Excel.Range
    type XLRefStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle

    [<Serializable>]
    type Address() =
        let mutable R: int = 0
        let mutable C: int = 0
        let mutable _wsn = None
        let mutable _wbn = None
        let mutable _path = None
        static member FromR1C1(R: int, C: int, wsname: string, wbname: string, path: string) : Address =
            let addr = Address()
            addr.Row <- R
            addr.Col <- C
            addr.WorksheetName <- Some(wsname)
            addr.WorkbookName <- Some(wbname)
            addr.Path <- Some(path)
            addr
        static member NewFromR1C1(R: int, C: int, wsname: string option, wbname: string option, path: string option) : Address =
            let addr = Address()
            addr.Row <- R
            addr.Col <- C
            addr.WorksheetName <- wsname
            addr.WorkbookName <- wbname
            addr.Path <- path
            addr
        static member NewFromA1(row: int, col: string, wsname: string option, wbname: string option, path: string option) : Address =
            let addr = Address()
            addr.Row <- row
            addr.Col <- Address.CharColToInt(col)
            addr.WorksheetName <- wsname
            addr.WorkbookName <- wbname
            addr.Path <- path
            addr
        member self.A1Local() : string = Address.IntToColChars(self.X) + self.Y.ToString()
        member self.A1Path() : string =
            match _path with
            | Some(pth) -> pth
            | None -> failwith "Path string should never be unset."
        member self.A1Worksheet() : string =
            match _wsn with
            | Some(ws) -> ws
            | None -> failwith "Worksheet string should never be unset."
        member self.A1Workbook() : string =
            match _wbn with
            | Some(wb) -> wb
            | None -> failwith "Workbook string should never be unset."
        member self.A1FullyQualified() : string =
            "[" + self.A1Workbook() + "]" + self.A1Worksheet() + "!" + self.A1Local()
        member self.R1C1 =
            let wsstr = match _wsn with | Some(ws) -> ws + "!" | None -> ""
            let wbstr = match _wbn with | Some(wb) -> "[" + wb + "]" | None -> ""
            let pstr = match _path with | Some(pth) -> pth | None -> ""
            pstr + wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.X: int = C
        member self.Y: int = R
        member self.Row
            with get() = R
            and set(value) = R <- value
        member self.Col
            with get() = C
            and set(value) = C <- value
        member self.Path
            with get() = _path
            and set(value) = _path <- value
        member self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        member self.WorkbookName
            with get() = _wbn
            and set(value) = _wbn <- value
        member self.AddressAsInt32() =
            // convert to zero-based indices
            // the modulus catches overflow; collisions are OK because our equality
            // operator does an exact check
            // underflow should throw an exception
            let col_idx = (C - 1) % 65536       // allow 16 bits for columns
            let row_idx = (R - 1) % 65536       // allow 16 bits for rows
            Debug.Assert(col_idx >= 0 && row_idx >= 0)
            row_idx + (col_idx <<< 16)
        // Address is used as a Dictionary key, and reference equality
        // does not suffice, therefore GetHashCode and Equals are provided
        override self.GetHashCode() : int = self.AddressAsInt32()
        override self.Equals(obj: obj) : bool =
            let addr = obj :?> Address
            self.SameAs addr
        member self.SameAs(addr: Address) : bool =
            self.X = addr.X &&
            self.Y = addr.Y &&
            self.WorksheetName = addr.WorksheetName &&
            self.WorkbookName = addr.WorkbookName
        member self.InsideRange(rng: Range) : bool =
            not (self.X < rng.getXLeft() ||
                 self.Y < rng.getYTop() ||
                 self.X > rng.getXRight() ||
                 self.Y > rng.getYBottom())
        member self.InsideAddr(addr: Address) : bool =
            self.X = addr.X && self.Y = addr.Y
        member self.GetCOMObject(app: Application) : XLRange =
            let wb: Workbook = app.Workbooks.Item(self.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(self.A1Worksheet()) :?> Worksheet
            let cell: XLRange = ws.Range(self.A1Local())
            cell
        override self.ToString() =
            "(" + self.Y.ToString() + "," + self.X.ToString() + ")"
        static member CharColToInt(col: string) : int =
            let rec ccti(idx: int) : int =
                let ltr = (int col.[idx]) - 64
                let num = (int (Math.Pow(26.0, float (col.Length - idx - 1)))) * ltr
                if idx = 0 then
                    num
                else
                    num + ccti(idx - 1)
            ccti(col.Length - 1)
        static member AddressFromCOMObject(com: Microsoft.Office.Interop.Excel.Range, wb: Microsoft.Office.Interop.Excel.Workbook) : Address =
            let wsname = com.Worksheet.Name
            let wbname = wb.Name
            let path = wb.FullName
            let addr = com.get_Address(true, true, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, Type.Missing, Type.Missing)
            Address.FromString(addr, Some(wsname), Some(wbname), Some(path))
        static member FromString(addr: string, wsname: string option, wbname: string option, path: string option) : Address =
            let reg = System.Text.RegularExpressions.Regex("R(?<row>[0-9]+)C(?<column>[0-9]+)")
            let m = reg.Match(addr)
            let r = System.Convert.ToInt32(m.Groups.["row"].Value)
            let c = System.Convert.ToInt32(m.Groups.["column"].Value)
            Address.NewFromR1C1(r, c, wsname, wbname, path)
        static member IntToColChars(dividend: int) : string =
            let mutable quot = dividend / 26
            let rem = dividend % 26
            if rem = 0 then
                quot <- quot - 1
            let ltr = if rem = 0 then
                        'Z'
                      else
                        char (64 + rem)
            if quot = 0 then
                ltr.ToString()
            else
                Address.IntToColChars(quot) + ltr.ToString()

    and Range(topleft: Address, bottomright: Address) =
        let _tl = topleft
        let _br = bottomright
        override self.ToString() =
            let tlstr = topleft.ToString()
            let brstr = bottomright.ToString()
            tlstr + "," + brstr
        member self.getXLeft() : int = _tl.X
        member self.getXRight() : int = _br.X
        member self.getYTop() : int = _tl.Y
        member self.getYBottom() : int = _br.Y
        member self.InsideRange(rng: Range) : bool =
            not (self.getXLeft() < rng.getXLeft() ||
                 self.getYTop() < rng.getYTop() ||
                 self.getXRight() > rng.getXRight() ||
                 self.getYBottom() > rng.getYBottom())
        // Yup, weird case.  This is because we actually
        // distinguish between addresses and ranges, unlike Excel.
        member self.InsideAddr(addr: Address) : bool =
            not (self.getXLeft() < addr.X ||
                 self.getYTop() < addr.Y ||
                 self.getXRight() > addr.X ||
                 self.getYBottom() > addr.Y)
        member self.SetPathName(path: string option) : unit =
            _tl.Path <- path
            _br.Path <- path
        member self.SetWorksheetName(wsname: string option) : unit =
            _tl.WorksheetName <- wsname
            _br.WorksheetName <- wsname
        member self.SetWorkbookName(wbname: string option) : unit =
            _tl.WorkbookName <- wbname
            _br.WorkbookName <- wbname
        member self.GetWorkbookNames() : seq<string> =
            [_tl.WorkbookName; _br.WorkbookName] |> List.choose id |> List.toSeq
        member self.GetPathNames() : seq<string> =
            [_tl.Path; _br.Path] |> List.choose id |> List.toSeq
        member self.GetCOMObject(app: Application) : XLRange =
            // tl and br must share workbook and worksheet (I think)
            let wb: Workbook = app.Workbooks.Item(_tl.A1Workbook())
            let ws: Worksheet = wb.Worksheets.Item(_tl.A1Worksheet()) :?> Worksheet
            let range: XLRange = ws.Range(_tl.A1Local(), _br.A1Local())
            range

    type Reference(path: string option, wbname: string option, wsname: string option) =
        let mutable _path: string option = path
        let mutable _wbn: string option = wbname
        let mutable _wsn: string option = wsname
        abstract member InsideRef: Reference -> bool
        abstract member Path: string option with get, set
        abstract member Resolve: string -> Workbook -> Worksheet -> unit
        abstract member WorkbookName: string option with get, set
        abstract member WorksheetName: string option with get, set
        default self.Path
            with get() = _path
            and set(value) = _path <- value
        default self.WorkbookName
            with get() = _wbn
            and set(value) = _wbn <- value
        default self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        default self.InsideRef(ref: Reference) = false
        default self.Resolve(path: string)(wb: Workbook)(ws: Worksheet) : unit =
            // we assume that missing workbook and worksheet
            // names mean that the address is local to the current
            // workbook and worksheet
            _path <- match self.Path with
                     | Some(pth) -> Some pth
                     | None -> Some path
            _wbn <- match self.WorkbookName with
                    | Some(wbn) -> Some wbn
                    | None -> Some wb.Name
            _wsn <- match self.WorksheetName with
                    | Some(wsn) -> Some wsn
                    | None -> Some ws.Name

    and ReferenceRange(path: string option, wbname: string option, wsname: string option, rng: Range) =
        inherit Reference(path, wbname, wsname)
        do
            // also set the Workbook and Worksheet for the Range object itself
            rng.SetWorkbookName(wbname)
            rng.SetWorksheetName(wsname)
        override self.ToString() =
            let pth = match self.Path with
                      | Some(pth) -> pth
                      | None -> ""
            let wbn = match self.WorkbookName with
                      | Some(wbn) -> wbn
                      | None -> ""
            let wsn = match self.WorksheetName with
                      | Some(wsn) -> wsn
                      | None -> ""
            "ReferenceRange(" + pth + ",[" + wbn + "]," + wsn + "," + rng.ToString() + ")"
        override self.InsideRef(ref: Reference) : bool =
            match ref with
            | :? ReferenceAddress as ar -> rng.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> rng.InsideRange(rr.Range)
            | _ -> failwith "Unknown Reference subclass."
        member self.Range = rng
        override self.Resolve(path: string)(wb: Workbook)(ws: Worksheet) =
            // we assume that missing workbook and worksheet
            // names mean that the address is local to the current
            // workbook and worksheet
            self.Path <- match self.Path with
                         | Some(pth) ->
                            rng.SetPathName(Some pth)
                            Some pth
                         | None ->
                            rng.SetPathName(Some path)
                            Some path
            self.WorkbookName <- match self.WorkbookName with
                                 // If we know it, we also pass the wbname
                                 // down to ranges and addresses
                                 | Some(wbn) ->
                                      rng.SetWorkbookName(Some wbn)
                                      Some wbn
                                 | None ->
                                      rng.SetWorkbookName(Some wb.Name)
                                      Some wb.Name
            self.WorksheetName <- match self.WorksheetName with
                                  | Some(wsn) ->
                                      rng.SetWorksheetName(Some wsn)
                                      Some wsn
                                  | None ->
                                      rng.SetWorksheetName(Some ws.Name)
                                      Some ws.Name

    and ReferenceAddress(path: string option, wbname: string option, wsname: string option, addr: Address) =
        inherit Reference(path, wbname, wsname)
        do
            // also set the Workbook and Worksheet for the Address object itself
            addr.WorkbookName <- wbname
            addr.WorksheetName <- wsname
        override self.ToString() =
            let pth = match self.Path with
                      | Some(pth) -> pth
                      | None -> ""
            let wbn = match self.WorkbookName with
                      | Some(wbn) -> wbn
                      | None -> ""
            let wsn = match self.WorksheetName with
                      | Some(wsn) -> wsn
                      | None -> ""
            "ReferenceAddress(" + pth + ",[" + wbn + "]," + wsn + "," + addr.ToString() + ")"
        member self.Address = addr
        override self.InsideRef(ref: Reference) =
            match ref with
            | :? ReferenceAddress as ar -> addr.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> addr.InsideRange(rr.Range)
            | _ -> failwith "Invalid Reference subclass."
        override self.Resolve(path: string)(wb: Workbook)(ws: Worksheet) =
            // always resolve the workbook name when it is missing
            // but only resolve the worksheet name when the
            // workbook name is not set
            self.Path <- match self.Path with
                         | Some(pth) ->
                            addr.Path <- Some pth
                            Some pth
                         | None ->
                            addr.Path <- Some path
                            Some path
            self.WorkbookName <- match self.WorkbookName with
                                 // If we know it, we also pass the wbname
                                 // down to ranges and addresses
                                 | Some(wbn) ->
                                      addr.WorkbookName <- Some wbn
                                      Some wbn
                                 | None ->
                                      addr.WorkbookName <- Some wb.Name
                                      Some wb.Name
            self.WorksheetName <- match self.WorksheetName with
                                  | Some(wsn) ->
                                      addr.WorksheetName <- Some wsn
                                      Some wsn
                                  | None ->
                                      addr.WorksheetName <- Some ws.Name
                                      Some ws.Name

    and ReferenceFunction(wsname: string option, fnname: string, arglist: Expression list) =
        inherit Reference(None, None, wsname)
        member self.ArgumentList = arglist
        member self.FunctionName = fnname
        override self.ToString() =
            fnname + "(" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
        override self.Resolve(path: string)(wb: Workbook)(ws: Worksheet) =
            // pass wb and ws information down to arguments
            // wb and ws names do not matter for functions
            for expr in arglist do
                expr.Resolve path wb ws

    and ReferenceConstant(wsname: string option, value: int) =
        inherit Reference(None, None, wsname)
        override self.ToString() = "Constant(" + value.ToString() + ")"

    and ReferenceString(wsname: string option, value: string) =
        inherit Reference(None, None, wsname)
        override self.ToString() = "String(" + value + ")"

    and ReferenceNamed(wsname: string option, varname: string) =
        inherit Reference(None, None, wsname)
        override self.ToString() =
            match self.WorksheetName with
            | Some(wsn) -> "ReferenceName(" + wsn + ", " + varname + ")"
            | None -> "ReferenceName(None, " + varname + ")"

    and Expression =
    | ReferenceExpr of Reference
    | BinOpExpr of string * Expression * Expression
    | UnaryOpExpr of char * Expression
    | ParensExpr of Expression
        member self.Resolve(path: string)(wb: Workbook)(ws: Worksheet) =
            match self with
            | ReferenceExpr(r) -> r.Resolve path wb ws
            | BinOpExpr(op,e1,e2) ->
                e1.Resolve path wb ws
                e2.Resolve path wb ws
            | UnaryOpExpr(op, e) ->
                e.Resolve path wb ws
            | ParensExpr(e) ->
                e.Resolve path wb ws
        override self.ToString() =
            match self with
            | ReferenceExpr(r) -> "ReferenceExpr(" + r.ToString() + ")"
            | BinOpExpr(op,e1,e2) -> "BinOpExpr(" + op.ToString() + "," + e1.ToString() + "," + e2.ToString() + ")"
            | UnaryOpExpr(op, e) -> "UnaryOpExpr(" + op.ToString() + "," + e.ToString() + ")"
            | ParensExpr(e) -> "ParensExpr(" + e.ToString() + ")"