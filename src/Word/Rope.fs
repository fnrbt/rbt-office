module RedBat.Word.Rope

open System

/// Maximum leaf size (1KB)
let [<Literal>] MAX_LEAF = 1024

/// Minimum leaf size for splitting (prevents excessive fragmentation)
let [<Literal>] MIN_LEAF = 511

/// Rope data structure - immutable balanced binary tree of text
type Rope =
    | Empty
    | Leaf of text: string
    | Node of left: Rope * right: Rope * info: NodeInfo

and NodeInfo = {
    Length: int       // Total character count
    Lines: int        // Number of newlines
    Height: int       // Tree height for balancing
}

/// Get length of rope
let rec length = function
    | Empty -> 0
    | Leaf s -> s.Length
    | Node (_, _, info) -> info.Length

/// Get line count (number of newlines)
let rec lines = function
    | Empty -> 0
    | Leaf s -> s |> Seq.filter ((=) '\n') |> Seq.length
    | Node (_, _, info) -> info.Lines

/// Get height for balancing
let height = function
    | Empty -> 0
    | Leaf _ -> 1
    | Node (_, _, info) -> info.Height

/// Create node info
let mkInfo left right = {
    Length = length left + length right
    Lines = lines left + lines right
    Height = 1 + max (height left) (height right)
}

/// Create balanced node
let mkNode left right =
    match left, right with
    | Empty, r -> r
    | l, Empty -> l
    | _ -> Node (left, right, mkInfo left right)

/// AVL tree balancing
let balance left right =
    let hl = height left
    let hr = height right

    if hl > hr + 2 then
        // Left-heavy
        match left with
        | Node (ll, lr, _) when height ll >= height lr ->
            // Left-left case: rotate right
            mkNode ll (mkNode lr right)
        | Node (ll, Node (lrl, lrr, _), _) ->
            // Left-right case: rotate left then right
            mkNode (mkNode ll lrl) (mkNode lrr right)
        | _ -> mkNode left right
    elif hr > hl + 2 then
        // Right-heavy
        match right with
        | Node (rl, rr, _) when height rr >= height rl ->
            // Right-right case: rotate left
            mkNode (mkNode left rl) rr
        | Node (Node (rll, rlr, _), rr, _) ->
            // Right-left case: rotate right then left
            mkNode (mkNode left rll) (mkNode rlr rr)
        | _ -> mkNode left right
    else
        mkNode left right

/// Create rope from string
let ofString (s: string) =
    if String.IsNullOrEmpty(s) then Empty
    elif s.Length <= MAX_LEAF then Leaf s
    else
        // Split into chunks
        let rec split acc start =
            if start >= s.Length then acc
            else
                let len = min MAX_LEAF (s.Length - start)
                let chunk = s.Substring(start, len)
                split (Leaf chunk :: acc) (start + len)

        // Build balanced tree from chunks
        let chunks = split [] 0 |> List.rev
        let rec buildTree = function
            | [] -> Empty
            | [x] -> x
            | xs ->
                let mid = xs.Length / 2
                let left = xs |> List.take mid |> buildTree
                let right = xs |> List.skip mid |> buildTree
                mkNode left right
        buildTree chunks

/// Convert rope to string
let rec toString = function
    | Empty -> ""
    | Leaf s -> s
    | Node (left, right, _) -> toString left + toString right

/// Insert text at position
let rec insert pos (text: string) rope =
    if String.IsNullOrEmpty(text) then rope
    else
        match rope with
        | Empty -> ofString text
        | Leaf s when s.Length + text.Length <= MAX_LEAF ->
            // Can fit in single leaf
            if pos <= 0 then Leaf (text + s)
            elif pos >= s.Length then Leaf (s + text)
            else Leaf (s.Substring(0, pos) + text + s.Substring(pos))
        | Leaf s ->
            // Need to split leaf
            let left = if pos > 0 then Leaf (s.Substring(0, pos)) else Empty
            let right = if pos < s.Length then Leaf (s.Substring(pos)) else Empty
            balance left (balance (ofString text) right)
        | Node (left, right, info) ->
            let leftLen = length left
            if pos <= leftLen then
                balance (insert pos text left) right
            else
                balance left (insert (pos - leftLen) text right)

/// Delete range [start, start + len)
let rec delete start len rope =
    if len <= 0 then rope
    else
        match rope with
        | Empty -> Empty
        | Leaf s ->
            if start <= 0 && len >= s.Length then Empty
            elif start <= 0 then Leaf (s.Substring(len))
            elif start >= s.Length then Leaf s
            elif start + len >= s.Length then Leaf (s.Substring(0, start))
            else Leaf (s.Substring(0, start) + s.Substring(start + len))
        | Node (left, right, info) ->
            let leftLen = length left
            let rightStart = max 0 (start - leftLen)
            let leftEnd = min leftLen (start + len)

            if start >= leftLen then
                // Delete entirely in right
                balance left (delete rightStart len right)
            elif start + len <= leftLen then
                // Delete entirely in left
                balance (delete start len left) right
            else
                // Delete spans both sides
                let leftDel = leftEnd - start
                let rightDel = len - leftDel
                balance (delete start leftDel left) (delete 0 rightDel right)

/// Get substring
let rec slice start len rope =
    if len <= 0 then ""
    else
        match rope with
        | Empty -> ""
        | Leaf s ->
            if start <= 0 && len >= s.Length then s
            elif start >= s.Length then ""
            else
                let actualStart = max 0 start
                let actualLen = min len (s.Length - actualStart)
                s.Substring(actualStart, actualLen)
        | Node (left, right, info) ->
            let leftLen = length left
            if start >= leftLen then
                slice (start - leftLen) len right
            elif start + len <= leftLen then
                slice start len left
            else
                let leftPart = slice start (leftLen - start) left
                let rightPart = slice 0 (len - leftLen + start) right
                leftPart + rightPart

/// Get character at position
let charAt pos rope =
    if pos < 0 || pos >= length rope then None
    else Some (slice pos 1 rope).[0]

/// Convert position to line/column
let positionToLineCol pos rope =
    let rec loop pos line col = function
        | Empty -> (line, col)
        | Leaf s ->
            let mutable line' = line
            let mutable col' = col
            for i = 0 to min pos s.Length - 1 do
                if s.[i] = '\n' then
                    line' <- line' + 1
                    col' <- 0
                else
                    col' <- col' + 1
            (line', col')
        | Node (left, right, _) ->
            let leftLen = length left
            if pos < leftLen then
                loop pos line col left
            else
                let (line', col') = loop leftLen line col left
                loop (pos - leftLen) line' col' right
    loop pos 0 0 rope

/// Convert line/column to position
let lineColToPosition line col rope =
    let rec loop targetLine targetCol currentLine currentCol currentPos rope =
        match rope with
        | Empty -> currentPos
        | Leaf s ->
            let mutable line' = currentLine
            let mutable col' = currentCol
            let mutable pos' = 0
            for i = 0 to s.Length - 1 do
                if line' = targetLine && col' = targetCol then
                    ()
                elif s.[i] = '\n' then
                    line' <- line' + 1
                    col' <- 0
                    pos' <- pos' + 1
                else
                    col' <- col' + 1
                    pos' <- pos' + 1
            currentPos + pos'
        | Node (left, right, _) ->
            let leftLines = lines left
            if targetLine < currentLine + leftLines then
                loop targetLine targetCol currentLine currentCol currentPos left
            else
                let leftLen = length left
                // Scan left to get final line/col
                let rec scanLeft l c rope =
                    match rope with
                    | Empty -> (l, c)
                    | Leaf s ->
                        let mutable line'' = l
                        let mutable col'' = c
                        for ch in s do
                            if ch = '\n' then
                                line'' <- line'' + 1
                                col'' <- 0
                            else
                                col'' <- col'' + 1
                        (line'', col'')
                    | Node (left, right, _) ->
                        let (line'', col'') = scanLeft l c left
                        scanLeft line'' col'' right
                let (line', col') = scanLeft currentLine currentCol left
                loop targetLine targetCol line' col' (currentPos + leftLen) right
    loop line col 0 0 0 rope
