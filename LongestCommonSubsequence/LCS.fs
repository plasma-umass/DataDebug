module LongestCommonSubsequence
    open System
    open System.Diagnostics

    type Sign =
    | Plus = 0
    | Minus = 1
    | Empty = 2

    type Error = 
    // first int is index in orig string; second int is delta
    | TranspositionError of int*int
    // first int is index in orig string; char is the character we were supposed to type;
    // string is what was actually typed
    | TypoError of int*char*string
    // first sign is orig string; second is in retyped string
    | SignError of Sign*Sign

    // compute all of the LCS lengths
    let LCSLength(X: string, Y: string) : int[,] =
        let m = X.Length
        let n = Y.Length
        let C = Array2D.create (m + 1) (n + 1) 0
        for i = 1 to m do
            for j = 1 to n do
                if X.[i-1] = Y.[j-1] then
                    C.[i,j] <- C.[i-1,j-1] + 1
                else
                    C.[i,j] <- System.Math.Max(C.[i,j-1],C.[i-1,j])
        C

    // this function reads out all of the longest subsequences
    let rec backtrackAll(C: int[,], X: string, Y: string, i: int, j: int) : Set<string> =
        if i = 0 || j = 0 then
            set[""]
        else if X.[i-1] = Y.[j-1] then
            Set.map (fun (Z: string) -> Z + X.[i-1].ToString()) (backtrackAll(C, X, Y, i-1, j-1))
        else
            let mutable R = Set.empty
            if C.[i,j-1] >= C.[i-1,j] then
                R <- backtrackAll(C, X, Y, i, j-1)
            if C.[i-1,j] >= C.[i,j-1] then
                R <- Set.union R (backtrackAll(C, X, Y, i-1, j))
            R

    // like backtrack except that it returns a character pair sequence
    // instead of a string
    // for each character pair: (X pos, Y pos)
    let rec getCharPairs_single(C: int[,], X: string, Y: string, i: int, j: int) : (int*int) list =
        if i = 0 || j = 0 then
            []
        else if X.[i-1] = Y.[j-1] then
            // append instead of prepend so that alignments
            // are in ascending order
            // we adjust offsets because C is 1-based
            getCharPairs_single(C, X, Y, i-1, j-1) @ [(i-1,j-1)]
        else
            if C.[i, j-1] > C.[i-1, j] then
                getCharPairs_single(C, X, Y, i, j-1)
            else
                getCharPairs_single(C, X, Y, i-1, j)

    // like backtrackAll except that it returns a set of character pair
    // sequences instead of a set of strings
    // for each character pair: (X pos, Y pos)
    let rec getCharPairs(C: int[,], X: string, Y: string, i: int, j: int, sw: Stopwatch, timeout: TimeSpan) : Set<(int*int) list> =
        if sw.Elapsed > timeout then
            raise (TimeoutException())
        if i = 0 || j = 0 then
            set[[]]
        else if X.[i-1] = Y.[j-1] then
            let mutable ZS = Set.map (fun (Z: (int*int) list) -> Z @ [(i-1,j-1)] ) (getCharPairs(C, X, Y, i-1, j-1, sw, timeout))
            if (C.[i,j] = C.[i,j-1]) then 
                ZS <- Set.union ZS (getCharPairs(C, X, Y, i, j-1, sw, timeout))
            ZS
        else
            let mutable R = Set.empty
            if C.[i,j-1] >= C.[i-1,j] then
                R <- getCharPairs(C, X, Y, i, j-1, sw, timeout)
            if C.[i-1,j] >= C.[i,j-1] then
                R <- Set.union R (getCharPairs(C, X, Y, i-1, j, sw, timeout))
            R

    // compute the set of longest strings
    let LCS(X: string, Y: string) : Set<string> =
        let m = X.Length
        let n = Y.Length
        let C = LCSLength(X,Y)
        backtrackAll(C, X, Y, m, n)

    // compute the set of longest character pair sequences
    let LCS_Char(X: string, Y: string) : Set<(int*int) list> =
        let timeout = TimeSpan(0,0,1) // 1 second
        let m = X.Length
        let n = Y.Length
        let C = LCSLength(X,Y)
        let sw = new Stopwatch()
        sw.Start()
        try
            getCharPairs(C, X, Y, m, n, sw, timeout)
        with
        | :? TimeoutException -> set [getCharPairs_single(C, X, Y, m, n)]

    // "pull" each Y index to the left as far as it will go
    // this allows key-repeat typos to be counted correctly
    let LeftAlignedLCS(orig: string, entered: string) : (int*int) list =
        let charseqs = LCS_Char(orig, entered) |> Set.toArray
        // randomly choose one of the longest subsequences
        let rng = System.Random()
        let charseq = charseqs.[rng.Next(charseqs.Length - 1)]
        // new sequence
        let mutable newseq: (int*int) list = []
        // realign characters
        let mutable jstop = -1
        for (i,j) in charseq do
            // all of the candidates to the left of j, but only as far as jstop
            let candidate_indices = [jstop+1..j]
            let mutable keep_looking = true
            for k in candidate_indices do
                if keep_looking then
                    if entered.[k] = entered.[j] then
                        jstop <- k
                        keep_looking <- false
            newseq <- newseq @ [(i,jstop)]
        // return realigned char list
        newseq

    // find each missing character in original, by index into original string
    let GetMissingCharIndices(orig: string, align: (int*int) list) : int list =
        let all_indices = set[0..orig.Length - 1]
        let align_indices = List.map (fun (o,_) -> o) align |> Set.ofList
        Set.difference all_indices align_indices |> Set.toList

    // find each appended character in entered, by index into entered string
    let GetAddedCharIndices(entered: string, align: (int*int) list) : int list =
        let all_indices = set[0..entered.Length - 1]
        let align_indices = List.map (fun (_,e) -> e) align |> Set.ofList
        Set.difference all_indices align_indices |> Set.toList

    // returns:
    // corrected entered string
    // new list of alignments
    // new list of additions
    // new list of omissions
    // transposition deltas
    let FixTranspositions(al: (int*int) list, ad: int list, om: int list, orig: string, entered: string)
        : string * (int*int) list * int list * int list * int list =
        // remember that alignments are: (original position, entered position)
        let rec FT(al: (int*int) list, ad: int list, om: int list, entered: string, tdeltas: int list)
            : string * (int*int) list * int list * int list * int list =
            if ad.Length = 0 || om.Length = 0 then
                // we strip the (-1,-1) pseudo-alignment out of the list
                entered,al.Tail,ad,om,tdeltas
            else
                // get the location of the first omission
                let omloc = om.Head
                // get the character of the first omission
                let ochar = orig.[omloc]
                // find the intended location of the char in the entered string
                // this character needs to be moved to the right of the rightmost
                // alignment that occurs just before this character
                let rightmost = snd (List.rev (List.filter (fun (origpos,entpos) -> origpos < omloc) (List.sortBy (fun (o,e) -> o) al))).Head
                let insertpos = rightmost + 1
                // get additions to the left of insertpos
                let lhs = List.filter (fun i -> i < insertpos) ad
                // get additions to the right of insertpos
                let rhs = List.filter (fun i -> i >= insertpos) ad
                // get lhs character positions that match ochar
                let lhs_matches = List.rev (List.filter (fun i -> entered.[i] = ochar) lhs)
                // get rhs character positions that match ochar
                let rhs_matches = List.filter (fun i -> entered.[i] = ochar) rhs
                // choose the closest matching addition
                let is_lhs,a_idx = match lhs_matches,rhs_matches with
                                   | l::ls,r::rs -> if System.Math.Abs(omloc - r) <= System.Math.Abs(omloc - l) then false,Some(r) else true,Some(l)
                                   | [],r::rs -> false,Some(r)
                                   | l::ls,[] -> true,Some(l)
                                   | [],[] -> false,None
                match is_lhs,a_idx with
                | _,None ->
                    // if no characters match the current omitted character,
                    // discard the omission and move on
                    FT(al, ad, om.Tail, entered, tdeltas)
                | il,Some(wrong_pos) ->
                    // insert the character in the appropriate location,
                    // ensuring that the sting is lengthened if the location
                    // occurs after the end of the string
                    if insertpos > entered.Length then
                        failwith "insertpos > entered.Length"
                    let entered' = if insertpos = entered.Length then
                                       entered + ochar.ToString()
                                   else
                                       entered.Substring(0,insertpos) + ochar.ToString() + entered.Substring(insertpos)
                    // remove the character from the entered position
                    let entered'' = if wrong_pos < insertpos then
                                        entered'.Substring(0,wrong_pos) + entered'.Substring(wrong_pos + 1)
                                    else
                                        entered'.Substring(0,wrong_pos + 1) + entered'.Substring(wrong_pos + 2)

                    // adjust the omissions list
                    let omissions = om.Tail

                    // adjust the additions list
                    // remove the addition, and then, for all additions between the omission position
                    // and the insertion position, shift one to the right
                    let additions = if wrong_pos < insertpos then
                                        List.filter (fun a -> a <> wrong_pos) ad |> List.map (fun a -> if a >= wrong_pos && a < insertpos then a - 1 else a) 
                                    else
                                        List.filter (fun a -> a <> wrong_pos) ad |> List.map (fun a -> if a >= insertpos && a < wrong_pos then a + 1 else a) 

                    // correcting a transposition is guaranteed to produce an additional alignment, namely
                    let align = if wrong_pos < insertpos then (omloc, insertpos - 1) else (omloc, insertpos)

                    // adjust the alignments
                    let alignmentz = if wrong_pos < insertpos then
                                         align :: List.map (fun (o,e) -> if e >= wrong_pos && e < insertpos then (o,e-1) else (o,e)) al
                                     else
                                         align :: List.map (fun (o,e) -> if e >= insertpos && e < wrong_pos then (o,e+1) else (o,e)) al
                    // sort
                    let alignments = List.sortBy (fun (o,e) -> o) alignmentz

                    // to calculate the delta of this transposition, we want to calculate the number
                    // of alignments crossed by this character move
                    let delta: int = if il then
                                     // the character was accidentally entered to the LEFT of where it should have been
                                     -(List.filter (fun (o,e) -> o < omloc && e > wrong_pos) al).Length
                                     else
                                     // the character was accidentally entered to the RIGHT of where it should have been
                                     (List.filter (fun (o,e) -> o > omloc && e < wrong_pos) al).Length

                    // process the next transposition
                    FT(alignments, additions, omissions, entered'', delta::tdeltas)
        // add the pseudo-start char alignment to ensure that there is
        // always a rightmost alignment
        FT((-1,-1)::al, ad, om, entered, [])

    // rounds to the nearest positive number, including zero
    let rnd(z: int) = if z < 0 then 0 else z

    // return all typos
    // this method assumes that you have already removed all transpositions
    // alignments: (original position, entered position)
    let GetTypos(alignments: (int*int) list, orig: string, entered: string) : (char option*string) list =
        let rng = System.Random()
        let rec typoget(al: (int*int) list, typos: (char option*string) list) : (char option*string) list =
            match al with
            | a1::a2::als ->
                // get all of the characters of the entered string between snd a1 and snd a2-1 inclusive
                let extra_chars = entered.Substring((snd a1) + 1, (snd a2) - (snd a1 + 1))
                // get all of the missing characters between fst a1 and fst a2 - 1 inclusive
                let omitted_chars = orig.Substring((fst a1) + 1, (fst a2) - (fst a1 + 1))
                // create n random partitions of extra_chars, where n = omitted_chars.Length
                // the first partition is always (0,n >= 1)
                let parts = Seq.pairwise (List.sort (0 :: List.map (fun partition -> rng.Next(0,extra_chars.Length+1)) [1..omitted_chars.Length] @ [extra_chars.Length])) |> Seq.toList
                // the first partition is always conditioned on the character in
                // orig at position fst a1 which may or may not be the empty string
                let a1_typo = match (fst a1) with
                                | -1 -> None, extra_chars.Substring(0, snd parts.Head)
                                | _ -> Some(orig.[fst a1]), orig.[fst a1].ToString() + extra_chars.Substring(0, snd parts.Head)
                // prepend to typo list
                // divvy up the extra_chars among the omitted_chars
                let typos' = (List.mapi (fun idx (pstart,pend) ->
                                            let echars = extra_chars
                                            let condchar = Some(omitted_chars.[idx])
                                            let length = pend-pstart
                                            let typostr = echars.Substring(pstart,length)
                                            condchar,typostr
                                        ) parts.Tail
                             ) @ a1_typo :: typos
                // process remaining typos
                typoget(a2::als, typos')
            | a::[] ->
                // get all the remaining characters
                let extra_chars = entered.Substring(snd a + 1)
                let a_char = match (fst a) with
                             | -1 -> None
                             | _ -> Some(orig.[fst a])
                let typos' = match (snd a) with
                             | -1 -> (a_char, "") :: typos
                             | _ -> (a_char, entered.[snd a].ToString() + extra_chars) :: typos
                typoget([], typos')
            | [] -> List.rev typos
        // call recursive function, prepending a "start of string" alignment to the list
        typoget((-1,-1)::alignments, [])

    // this is for C# unit test use
    let LeftAlignedLCSList(orig: string, entered: string) : System.Collections.Generic.IEnumerable<(int*int)> =
        LeftAlignedLCS(orig, entered) |> List.toSeq

    let LCS_Hash(X: string, Y: string) : System.Collections.Generic.HashSet<string> =
        let hs = new System.Collections.Generic.HashSet<string>()
        for s in LCS(X,Y) do hs.Add(s) |> ignore
        hs

    let LCS_Hash_Char(X: string, Y: string) : System.Collections.Generic.HashSet<System.Collections.Generic.IEnumerable<int*int>> =
        let hs = new System.Collections.Generic.HashSet<System.Collections.Generic.IEnumerable<int*int>>()
        for ls in LCS_Char(X,Y) do hs.Add(ls |> List.toSeq) |> ignore
        hs

    let ToFSList(arr: 'a[]) : 'a list = List.ofArray arr