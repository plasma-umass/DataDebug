module LongestCommonSubsequence
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
            let sets = backtrackAll(C, X, Y, i-1, j-1)
            let ZS = Set.map (fun (Z: string) -> Z + X.[i-1].ToString()) sets
            ZS
        else
            let mutable R = Set.empty
            if C.[i,j-1] >= C.[i-1,j] then
                R <- backtrackAll(C, X, Y, i, j-1)
            if C.[i-1,j] >= C.[i,j-1] then
                R <- Set.union R (backtrackAll(C, X, Y, i-1, j))
            R

    // compute the set of longest strings
    let LCS(X: string, Y: string) : Set<string> =
        let m = X.Length
        let n = Y.Length
        let C = LCSLength(X,Y)
        backtrackAll(C, X, Y, m, n)

    let LCS_Hash(X: string, Y: string) : System.Collections.Generic.HashSet<string> =
        let hs = new System.Collections.Generic.HashSet<string>()
        for s in LCS(X,Y) do hs.Add(s) |> ignore
        hs
