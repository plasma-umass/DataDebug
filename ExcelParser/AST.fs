module AST
    open System
    open System.Diagnostics

    type Address(R: int, C: int, wsname: string option, wbname: string option) =
        let mutable _wsn = wsname
        let mutable _wbn = wbname
        new(row: int, col: string, wsname: string option, wbname: string option) =
            Address(row, Address.CharColToInt(col), wsname, wbname)
        member self.R1C1 =
            let wsstr = match wsname with | Some(ws) -> ws + "!" | None -> ""
            let wbstr = match wbname with | Some(wb) -> "[" + wb + "]" | None -> ""
            wbstr + wsstr + "R" + R.ToString() + "C" + C.ToString()
        member self.X = C
        member self.Y = R
        member self.WorksheetName
            with get() = _wsn
            and set(value) = _wsn <- value
        member self.WorkbookName
            with get() = _wbn
            and set(value) = _wbn <- value
        member self.XLAddress() : string = self.R1C1
        member self.AddressAsInt32() =
            // convert to zero-based indices
            // the modulus catches overflow; collisions are OK because our equality
            // operator does an exact check
            // underflow should throw an exception
            let col_idx = (C - 1) % 65536     // allow 16 bits for columns
            let row_idx = (R - 1) % 65536      // allow 16 bits for rows
            Debug.Assert(col_idx >= 0 && row_idx >= 0)
            row_idx + (col_idx <<< 16)
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

    type Reference(wsname: string option) =
        let mutable wbname = None
        abstract member InsideRef: Reference -> bool
        abstract member WorkbookName: string option with get, set
        abstract member WorksheetName: string option
        default self.WorkbookName
            with get() = wbname
            and set(value) = wbname <- value
        default self.WorksheetName = wsname
        default self.InsideRef(ref: Reference) = false

    and ReferenceRange(wsname: string option, rng: Range) =
        inherit Reference(wsname)
        do rng.SetWorksheetName(wsname)
        override self.ToString() =
            match wsname with
            | Some(wsn) -> "ReferenceRange(" + wsn.ToString() + ", " + rng.ToString() + ")"
            | None -> "ReferenceRange(None, " + rng.ToString() + ")"
        override self.InsideRef(ref: Reference) : bool =
            match ref with
            | :? ReferenceAddress as ar -> rng.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> rng.InsideRange(rr.Range)
            | _ -> failwith "Unknown Reference subclass."
        member self.Range = rng

    and ReferenceAddress(wsname: string option, addr: Address) =
        inherit Reference(wsname)
        do addr.WorksheetName <- wsname
        override self.ToString() =
            match wsname with
            | Some(wsn) -> "ReferenceAddress(" + wsn.ToString() + ", " + addr.ToString() + ")"
            | None -> "ReferenceAddress(None, " + addr.ToString() + ")"
        member self.Address = addr
        override self.InsideRef(ref: Reference) =
            match ref with
            | :? ReferenceAddress as ar -> addr.InsideAddr(ar.Address)
            | :? ReferenceRange as rr -> addr.InsideRange(rr.Range)
            | _ -> failwith "Invalid Reference subclass."

    and ReferenceFunction(wsname: string option, fnname: string, arglist: Reference list) =
        inherit Reference(wsname)
        override self.ToString() =
            fnname + "(" + String.Join(",", (List.map (fun arg -> arg.ToString()) arglist)) + ")"

    and ReferenceNamed(wsname: string option, varname: string) =
        inherit Reference(wsname)
        override self.ToString() =
            match wsname with
            | Some(wsn) -> "ReferenceName(" + wsn + ", " + varname + ")"
            | None -> "ReferenceName(None, " + varname + ")"