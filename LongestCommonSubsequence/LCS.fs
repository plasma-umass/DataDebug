module LongestCommonSubsequence
    type Sign =
    | Plus = 0
    | Minus = 1
    | Empty = 2

    type Error = 
    // first int is index in orig string; second int is delta
    | TranspositionError of int*int
    // first int is index in orig string; char is the character we we supposed to type;
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

    // like backtrackAll except that it returns a set of character pair
    // sequences instead of a set of strings
    // for each character pair: (X pos, Y pos)
    let rec getCharPairs(C: int[,], X: string, Y: string, i: int, j: int) : Set<(int*int) list> =
        if i = 0 || j = 0 then
            set[[]]
        else if X.[i-1] = Y.[j-1] then
            let mutable ZS = Set.map (fun (Z: (int*int) list) -> Z @ [(i-1,j-1)] ) (getCharPairs(C, X, Y, i-1, j-1))
            if (C.[i,j] = C.[i,j-1]) then 
                ZS <- Set.union ZS (getCharPairs(C, X, Y, i, j-1))
            ZS
        else
            let mutable R = Set.empty
            if C.[i,j-1] >= C.[i-1,j] then
                R <- getCharPairs(C, X, Y, i, j-1)
            if C.[i-1,j] >= C.[i,j-1] then
                R <- Set.union R (getCharPairs(C, X, Y, i-1, j))
            R

    // compute the set of longest strings
    let LCS(X: string, Y: string) : Set<string> =
        let m = X.Length
        let n = Y.Length
        let C = LCSLength(X,Y)
        backtrackAll(C, X, Y, m, n)

    // compute the set of longest character pair sequences
    let LCS_Char(X: string, Y: string) : Set<(int*int) list> =
        let m = X.Length
        let n = Y.Length
        let C = LCSLength(X,Y)
        getCharPairs(C, X, Y, m, n)

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
        let mutable jstop = 0
        for (i,j) in charseq do
            // all of the candidates to the left of j, but only as far as jstop
            let candidate_indices = List.rev [jstop..j]
            let mutable keep_looking = true
            for k in candidate_indices do
                if keep_looking then
                    if entered.[k] = entered.[j] then
                        jstop <- k
                    else
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

    // find all transpositions
    // returns (addition,omission) pairs
    // unusual structure to make call tail-recursive
    let rec GetTranspositions(additions: int list, omissions: int list, orig: string, entered: string, transpositions: (int*int) list) : (int*int) list =
        if additions.Length = 0 || omissions.Length = 0 then
            List.rev transpositions
        else
            // get the character of the first omission
            let ochar = orig.[omissions.Head]
            // get entered chars to the left of ochar
            let lhs = List.filter (fun i -> i <= omissions.Head) additions
            // get entered chars to the right of ochar
            let rhs = List.filter (fun i -> i > omissions.Head) additions
            // get lhs character positions that match ochar
            let lhs_matches = List.rev (List.filter (fun i -> entered.[i] = ochar) lhs)
            // get rhs character positions that match ochar
            let rhs_matches = List.filter (fun i -> entered.[i] = ochar) rhs
            // choose the closest match
            let is_lhs,a_idx = match lhs_matches,rhs_matches with
                             | l::ls,r::rs -> if System.Math.Abs(omissions.Head - r) <= System.Math.Abs(omissions.Head - l) then false,Some(r) else true,Some(l)
                             | [],r::rs -> false,Some(r)
                             | l::ls,[] -> true,Some(l)
                             | [],[] -> false,None
            match is_lhs,a_idx with
            | _,None ->
                // if no characters match the current omitted character,
                // discard the character and move on
                GetTranspositions(additions, omissions.Tail, orig, entered, transpositions)
            | il,Some(idx) ->
                let additions' = List.filter (fun a -> a <> idx) additions
                let omissions' = omissions.Tail
                GetTranspositions(additions', omissions', orig, entered, (omissions.Head,idx) :: transpositions)

    // rounds to the nearest positive number, including zero
    let rnd(z: int) = if z < 0 then 0 else z

    // return all typos
    // this method assumes that you have already removed all transpositions
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
                let parts = Seq.pairwise (List.sort (0 :: List.map (fun partition -> rng.Next(1,extra_chars.Length+1)) [1..omitted_chars.Length] @ [extra_chars.Length])) |> Seq.toList
                // the first partition is always conditioned on the character in
                // orig at position fst a1 which may or may not be the empty string
                let a1_typo = match (fst a1) with
                                | -1 -> None, extra_chars.Substring(0, snd parts.Head)
                                | _ -> Some(orig.[fst a1]), orig.[fst a1].ToString() + extra_chars.Substring(0, snd parts.Head)
                // prepend to typo list
                // pstart is inclusive
                // pend is exclusive
                let typos' = a1_typo :: (List.mapi (fun idx (pstart,pend) -> Some(orig.[idx]),entered.Substring(pstart,pend-pstart)) parts.Tail) @ typos
                // process remaining typos
                typoget(a2::als, typos')
            | a::[] ->
                // get all the remaining characters
                let extra_chars = entered.Substring(snd a + 1)
                let a_char = match (fst a) with
                             | -1 -> None
                             | _ -> Some(orig.[fst a])
                let typos' = (a_char,extra_chars) :: typos
                typoget([], typos')
            | [] -> typos
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