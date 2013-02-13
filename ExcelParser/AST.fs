module AST
    open System.Diagnostics

    type Address(R: int, C: int) =
        override self.ToString() =
            "(" + R.ToString() + "," + C.ToString() + ")"
        member self.X = R
        member self.Y = C
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