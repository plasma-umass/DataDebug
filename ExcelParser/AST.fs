module AST
    open System.Diagnostics
    open Microsoft.Office.Interop.Excel

    type Worksheet = Microsoft.Office.Interop.Excel.Worksheet
    type XLRange = Microsoft.Office.Interop.Excel.Range
    type XLRefStyle = Microsoft.Office.Interop.Excel.XlReferenceStyle

    type Address(R: int, C: int, ws: Worksheet) =
        override self.ToString() =
            "(" + R.ToString() + "," + C.ToString() + ")"
        member self.R1C1 = "R" + R.ToString() + "C" + C.ToString()
        member self.X = C
        member self.Y = R
        member self.worksheet_idx = ws.Index - 1
        member self.Cell() : XLRange = ws.Cells.Range(self.R1C1)
        member self.AddressAsInt32() =
            // convert to zero-based indices
            // the modulus catches overflow; collisions are OK because our equality
            // operator does an exact check
            // underflow should throw an exception
            let ws_idx = (ws.Index - 1) % 32 // allow 5 bits for worksheet index
            let col_idx = (C - 1) % 2048     // allow 11 bits for columns
            let row_idx = R - 1 % 65536      // allow 16 bits for rows
            Debug.Assert(ws_idx >= 0 && col_idx >= 0 && row_idx >= 0)
            row_idx + (col_idx <<< 11) + (ws_idx <<< 16)
        override self.GetHashCode() : int = self.AddressAsInt32()
        override self.Equals(obj: obj) : bool =
            let addr = obj :?> Address
            self.SameAs addr
        member self.SameAs(addr: Address) : bool =
            self.X = addr.X &&
            self.Y = addr.Y &&
            self.worksheet_idx = addr.worksheet_idx

    type Range(topleft: Address, bottomright: Address option) =
        let _tl = topleft
        // single-address range or multi-address range?
        let _br =
            match bottomright with
            | Some(br) -> br
            | None -> topleft    
        override self.ToString() =
            let tlstr = topleft.ToString()
            let brstr = match bottomright with
                        | Some(br) -> br.ToString()
                        | None -> "*"
            tlstr + "," + brstr
        member public self.getXLeft() : int = _tl.X
        member public self.getXRight() : int = _br.X
        member public self.getYTop() : int = _tl.Y
        member public self.getYBottom() : int = _br.Y