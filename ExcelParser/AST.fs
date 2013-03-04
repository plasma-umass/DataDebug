module AST
    open System
    open System.Diagnostics
    open Microsoft.Office.Interop.Excel

    type Application = Microsoft.Office.Interop.Excel.Application
    type Workbook = Microsoft.Office.Interop.Excel.Workbook
    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type XLRange = Microsoft.Office.Interop.Excel.Range
    type XLRefStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle

    type Address(R: int, C: int, wsname: string option, wbname: string option) =
        let mutable _wsn = wsname
        let mutable _wbn = wbname
        new(row: int, col: string, wsname: string option, wbname: string option) =
            Address(row, Address.CharColToInt(col), wsname, wbname)
        member self.A1(app: Application) : string =
            let wsstr = match _wsn with | Some(ws) -> ws + "!" | None -> ""
            let wbstr = match _wbn with | Some(wb) -> "[" + wb + "]" | None -> ""
            wbstr + wsstr + Address.IntToColChars(self.Y) + self.X.ToString()
        member self.R1C1 =
            let wsstr = match _wsn with | Some(ws) -> ws + "!" | None -> ""
            let wbstr = match _wbn with | Some(wb) -> "[" + wb + "]" | None -> ""
            wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.X: int = C
        member self.Y: int = R
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
            app.Range(self.A1(app), self.A1(app))
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
        static member IntToColChars(dividend: int) : string =
            let quot = dividend / 26
            let rem = dividend % 26
            let ltr = char (64 + rem)
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
        member self.SetWorksheetName(wsname: string option) : unit =
            _tl.WorksheetName <- wsname
            _br.WorksheetName <- wsname
        member self.SetWorkbookName(wbname: string option) : unit =
            _tl.WorkbookName <- wbname
            _br.WorkbookName <- wbname
        member self.GetCOMObject(app: Application) : XLRange =
            let tla1 = _tl.A1(app)
            let bra1 = _br.A1(app)
            app.Range(tla1, bra1)

    type Reference(wsname: string option) =
        let mutable _wbn: string option = None
        let mutable _wsn: string option = wsname
        abstract member InsideRef: Reference -> bool
        abstract member Resolve: Workbook -> Worksheet -> unit
        abstract member WorkbookName: string option with get, set
        abstract member WorksheetName: string option with get, set
        default self.WorkbookName
            with get() = _wbn
            and set(value) =
                _wbn <- value
        default self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        default self.InsideRef(ref: Reference) = false
        default self.Resolve(wb: Workbook)(ws: Worksheet) : unit =
            // always resolve the workbook name when it is missing
            // but only resolve the worksheet name when the
            // workbook name is not set
            _wbn <- match self.WorkbookName with
                    | Some(wbn) -> Some wbn
                    | None -> if wb.Name <> "" then
                                  Some wb.Name
                              else
                                  None
            _wsn <- match self.WorksheetName with
                    | Some(wsn) -> Some wsn
                    | None ->
                        match self.WorkbookName with
                        | Some(wbn) -> None
                        | None -> Some ws.Name

    and ReferenceRange(wsname: string option, rng: Range) =
        inherit Reference(wsname)
        do rng.SetWorksheetName(wsname)
        override self.ToString() =
            match self.WorksheetName with
            | Some(wsn) -> "ReferenceRange(" + wsn.ToString() + ", " + rng.ToString() + ")"
            | None -> "ReferenceRange(None, " + rng.ToString() + ")"
        override self.InsideRef(ref: Reference) : bool =
            match ref with
            | :? ReferenceAddress as ar -> rng.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> rng.InsideRange(rr.Range)
            | _ -> failwith "Unknown Reference subclass."
        member self.Range = rng
        override self.Resolve(wb: Workbook)(ws: Worksheet) =
            // always resolve the workbook name when it is missing
            // but only resolve the worksheet name when the
            // workbook name is not set
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
                                      match self.WorkbookName with
                                      | Some(wbn) -> None
                                      | None ->
                                          rng.SetWorksheetName(Some ws.Name)
                                          Some ws.Name

    and ReferenceAddress(wsname: string option, addr: Address) =
        inherit Reference(wsname)
        do addr.WorksheetName <- wsname
        override self.ToString() =
            match self.WorksheetName with
            | Some(wsn) -> "ReferenceAddress(" + wsn.ToString() + ", " + addr.ToString() + ")"
            | None -> "ReferenceAddress(None, " + addr.ToString() + ")"
        member self.Address = addr
        override self.InsideRef(ref: Reference) =
            match ref with
            | :? ReferenceAddress as ar -> addr.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> addr.InsideRange(rr.Range)
            | _ -> failwith "Invalid Reference subclass."
        override self.Resolve(wb: Workbook)(ws: Worksheet) =
            // always resolve the workbook name when it is missing
            // but only resolve the worksheet name when the
            // workbook name is not set
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
                                      match self.WorkbookName with
                                      | Some(wbn) -> None
                                      | None ->
                                          addr.WorksheetName <- Some ws.Name
                                          Some ws.Name

    and ReferenceFunction(wsname: string option, fnname: string, arglist: Reference list) =
        inherit Reference(wsname)
        member self.ArgumentList = arglist
        override self.ToString() =
            fnname + "(" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"
        override self.Resolve(wb: Workbook)(ws: Worksheet) =
            // pass wb and ws information down to arguments
            // wb and ws names do not matter for functions
            for arg in arglist do
                arg.Resolve wb ws

    and ReferenceConstant(wsname: string option, value: int) =
        inherit Reference(wsname)
        override self.ToString() = "Constant(" + value.ToString() + ")"

    and ReferenceNamed(wsname: string option, varname: string) =
        inherit Reference(wsname)
        override self.ToString() =
            match self.WorksheetName with
            | Some(wsn) -> "ReferenceName(" + wsn + ", " + varname + ")"
            | None -> "ReferenceName(None, " + varname + ")"

    and Expression =
    | ReferenceExpr of Reference
    | BinOpExpr of char * Expression * Expression
    | UnaryOpExpr of char * Expression
    | ParensExpr of Expression