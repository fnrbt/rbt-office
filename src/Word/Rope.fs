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
    TrailingColumn: int // Characters after the last newline in this subtree
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
    | Leaf s ->
        let mutable count = 0
        for i = 0 to s.Length - 1 do
            if s.[i] = '\n' then
                count <- count + 1
        count
    | Node (_, _, info) -> info.Lines

/// Get height for balancing
let height = function
    | Empty -> 0
    | Leaf _ -> 1
    | Node (_, _, info) -> info.Height

let private trailingColumnInString (s: string) =
    let mutable index = s.Length - 1
    let mutable found = false
    while not found && index >= 0 do
        if s.[index] = '\n' then
            found <- true
        else
            index <- index - 1
    if found then s.Length - index - 1 else s.Length

let private trailingColumn rope =
    let rec loop = function
        | Empty -> 0
        | Leaf s -> trailingColumnInString s
        | Node (left, right, _) ->
            if lines right > 0 then
                loop right
            else
                loop left + length right
    loop rope

let private trailingColumnFast = function
    | Empty -> 0
    | Leaf s -> trailingColumnInString s
    | Node (_, _, info) -> info.TrailingColumn

let private advanceLineCol line col rope =
    let newlineCount = lines rope
    if newlineCount = 0 then
        struct (line, col + length rope)
    else
        struct (line + newlineCount, trailingColumnFast rope)

/// Create node info
let mkInfo left right =
    let rightLines = lines right
    let trailing =
        if rightLines > 0 then
            trailingColumnFast right
        else
            trailingColumnFast left + length right
    {
        Length = length left + length right
        Lines = lines left + rightLines
        TrailingColumn = trailing
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
        let chunks = ResizeArray<Rope>()
        let mutable start = 0
        while start < s.Length do
            let len = min MAX_LEAF (s.Length - start)
            chunks.Add(Leaf (s.Substring(start, len)))
            start <- start + len

        let rec buildTree start count =
            if count <= 0 then Empty
            elif count = 1 then chunks.[start]
            else
                let mid = count / 2
                let left = buildTree start mid
                let right = buildTree (start + mid) (count - mid)
                mkNode left right
        buildTree 0 chunks.Count

/// Convert rope to string
let toString rope =
    match rope with
    | Empty -> ""
    | Leaf s -> s
    | Node _ ->
        let totalLength = length rope
        let sb = Text.StringBuilder(totalLength)
        let rec append = function
            | Empty -> ()
            | Leaf s -> sb.Append(s) |> ignore
            | Node (left, right, _) ->
                append left
                append right
        append rope
        sb.ToString()

let appendToStringBuilder (sb: Text.StringBuilder) rope =
    let rec append = function
        | Empty -> ()
        | Leaf s -> sb.Append(s) |> ignore
        | Node (left, right, _) ->
            append left
            append right
    append rope

let appendSliceToStringBuilder (sb: Text.StringBuilder) start len rope =
    if len > 0 then
        let rec appendSlice start len rope =
            if len > 0 then
                match rope with
                | Empty -> ()
                | Leaf s ->
                    if start <= 0 && len >= s.Length then
                        sb.Append(s) |> ignore
                    elif start < s.Length then
                        let actualStart = max 0 start
                        let actualLen = min len (s.Length - actualStart)
                        if actualLen > 0 then
                            sb.Append(s, actualStart, actualLen) |> ignore
                | Node (left, right, _) ->
                    let leftLen = length left
                    if start >= leftLen then
                        appendSlice (start - leftLen) len right
                    elif start + len <= leftLen then
                        appendSlice start len left
                    else
                        appendSlice start (leftLen - start) left
                        appendSlice 0 (len - leftLen + start) right
        appendSlice start len rope

let sliceEqualsString start len (text: string) rope =
    if isNull text || len <> text.Length || start < 0 || len < 0 then
        false
    elif len = 0 then
        true
    else
        let totalLength = length rope
        if start > totalLength - len then
            false
        else
            let mutable textIndex = 0
            let mutable equal = true
            let rec compareSlice start len rope =
                if equal && len > 0 then
                    match rope with
                    | Empty ->
                        equal <- false
                    | Leaf s ->
                        if start >= s.Length then
                            equal <- false
                        else
                            let actualStart = max 0 start
                            let actualLen = min len (s.Length - actualStart)
                            let mutable index = 0
                            while equal && index < actualLen do
                                if s.[actualStart + index] <> text.[textIndex + index] then
                                    equal <- false
                                index <- index + 1
                            textIndex <- textIndex + actualLen
                    | Node (left, right, _) ->
                        let leftLen = length left
                        if start >= leftLen then
                            compareSlice (start - leftLen) len right
                        elif start + len <= leftLen then
                            compareSlice start len left
                        else
                            compareSlice start (leftLen - start) left
                            compareSlice 0 (len - leftLen + start) right
            compareSlice start len rope
            equal && textIndex = text.Length

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
let slice start len rope =
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
        | Node _ ->
            let sb = Text.StringBuilder(min len (length rope))
            let rec appendSlice start len rope =
                if len > 0 then
                    match rope with
                    | Empty -> ()
                    | Leaf s ->
                        if start <= 0 && len >= s.Length then
                            sb.Append(s) |> ignore
                        elif start < s.Length then
                            let actualStart = max 0 start
                            let actualLen = min len (s.Length - actualStart)
                            if actualLen > 0 then
                                sb.Append(s, actualStart, actualLen) |> ignore
                    | Node (left, right, _) ->
                        let leftLen = length left
                        if start >= leftLen then
                            appendSlice (start - leftLen) len right
                        elif start + len <= leftLen then
                            appendSlice start len left
                        else
                            appendSlice start (leftLen - start) left
                            appendSlice 0 (len - leftLen + start) right
            appendSlice start len rope
            sb.ToString()

/// Get a rope containing the requested slice, preserving existing leaves that
/// are fully inside the slice.
let sliceRope start len rope =
    if len <= 0 then Empty
    else
        let total = length rope
        let boundedStart = max 0 (min start total)
        let boundedEnd = max boundedStart (min total (boundedStart + len))
        if boundedStart = 0 && boundedEnd = total then
            rope
        else
            let trimmedRight = delete boundedEnd (total - boundedEnd) rope
            delete 0 boundedStart trimmedRight

/// Get character at position
let charAt pos rope =
    if pos < 0 || pos >= length rope then None
    else
        let rec loop pos = function
            | Empty -> None
            | Leaf s -> Some s.[pos]
            | Node (left, right, _) ->
                let leftLen = length left
                if pos < leftLen then loop pos left
                else loop (pos - leftLen) right
        loop pos rope

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
                let struct (line', col') = advanceLineCol line col left
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
            let mutable found = ValueNone
            let mutable i = 0
            while ValueOption.isNone found && i < s.Length do
                if line' = targetLine && col' = targetCol then
                    found <- ValueSome pos'
                elif s.[i] = '\n' then
                    line' <- line' + 1
                    col' <- 0
                    pos' <- pos' + 1
                else
                    col' <- col' + 1
                    pos' <- pos' + 1
                i <- i + 1
            match found with
            | ValueSome pos -> currentPos + pos
            | ValueNone -> currentPos + pos'
        | Node (left, right, _) ->
            let struct (leftEndLine, leftEndCol) = advanceLineCol currentLine currentCol left
            if targetLine < leftEndLine || (targetLine = leftEndLine && targetCol <= leftEndCol) then
                loop targetLine targetCol currentLine currentCol currentPos left
            else
                let leftLen = length left
                loop targetLine targetCol leftEndLine leftEndCol (currentPos + leftLen) right
    loop line col 0 0 0 rope
