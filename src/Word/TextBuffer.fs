module RedBat.Word.TextBuffer

open RedBat.Word.Rope

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

/// Text style for a run
type TextStyle = {
    Style: string option
    Bold: bool
    Italic: bool
    Underline: bool
    /// w:u@val — single|double|thick|dotted|dash|dotDash|dotDotDash|wave|wavyDouble|… (None = single).
    UnderlineStyle: string option
    /// w:u@color — distinct underline colour (None = text colour).
    UnderlineColor: string option
    AllCaps: bool
    SmallCaps: bool
    Strike: bool
    DoubleStrike: bool          // w:dstrike
    Hidden: bool                // w:vanish — not rendered, takes no space
    Position: float             // w:position — baseline raise(+)/lower(-) in points
    EmphasisMark: string option // w:em@val — dot|comma|circle|underDot
    FontSize: float  // in points
    CharacterSpacing: float
    FontFamily: string option
    Color: string option
    Highlight: string option
    Hyperlink: string option
    CommentId: string option
    CommentText: string option
    CommentAuthor: string option
    CommentDate: string option
    NoteId: string option
    NoteType: string option
    NoteText: string option
    BookmarkStartId: string option
    BookmarkName: string option
    BookmarkEndId: string option
    FieldCode: string option
    FieldResult: string option
    RenderedPageBreakBefore: bool
    EquationXml: string option
    RevisionId: string option
    RevisionType: string option
    RevisionAuthor: string option
    RevisionDate: string option
    VertAlign: string option
}

/// Default text style
let defaultStyle = {
    Style = None
    Bold = false
    Italic = false
    Underline = false
    UnderlineStyle = None
    UnderlineColor = None
    AllCaps = false
    SmallCaps = false
    Strike = false
    DoubleStrike = false
    Hidden = false
    Position = 0.0
    EmphasisMark = None
    FontSize = 11.0  // 11pt default (like Word)
    CharacterSpacing = 0.0
    FontFamily = None
    Color = None
    Highlight = None
    Hyperlink = None
    CommentId = None
    CommentText = None
    CommentAuthor = None
    CommentDate = None
    NoteId = None
    NoteType = None
    NoteText = None
    BookmarkStartId = None
    BookmarkName = None
    BookmarkEndId = None
    FieldCode = None
    FieldResult = None
    RenderedPageBreakBefore = false
    EquationXml = None
    RevisionId = None
    RevisionType = None
    RevisionAuthor = None
    RevisionDate = None
    VertAlign = None
}

let characterSpacingSlots (text: string) =
    if System.String.IsNullOrEmpty text then 0.0
    else
        let mutable count = 0
        let mutable last = '\000'
        for i = 0 to text.Length - 1 do
            let ch = text.[i]
            if ch <> '\t' && ch <> '\r' && ch <> '\n' then
                count <- count + 1
                last <- ch
        if count = 0 then 0.0
        elif System.Char.IsWhiteSpace last then float count
        else float (count - 1)

let characterSpacingAdvance (text: string) (style: TextStyle) =
    style.CharacterSpacing * characterSpacingSlots text

/// Image data for inline images
type ImageRun = {
    Start: int           // Character offset (position in text)
    Src: string          // Data URL or image source
    Linked: bool
    Width: float option  // Width in pixels (None = intrinsic)
    Height: float option // Height in pixels (None = intrinsic)
    Alt: string option   // Alt text/description
    Name: string option
    Title: string option
    Floating: bool       // Imported floating/anchored image metadata
    Wrap: string option  // DOCX wrap mode when Floating is true
    PositionX: float option
    PositionY: float option
    PositionHRelativeFrom: string option
    PositionVRelativeFrom: string option
    BehindDoc: bool option
    AllowOverlap: bool option
    LayoutInCell: bool option
    DistanceTop: float option
    DistanceBottom: float option
    DistanceLeft: float option
    DistanceRight: float option
    RelativeHeight: string option
    Locked: bool option
    SimplePosition: bool option
    SimplePositionX: float option
    SimplePositionY: float option
    CropLeft: string option
    CropTop: string option
    CropRight: string option
    CropBottom: string option
}

/// Text run with consistent styling
type TextRun = {
    Start: int       // Character offset start
    Length: int      // Number of characters
    Style: TextStyle
}

/// Run can be either text or image
type Run =
    | Text of TextRun
    | Image of ImageRun

/// Text buffer with formatting information
type TextBuffer = {
    Rope: Rope
    Runs: Run list  // Non-overlapping runs covering the text (images use U+FFFC placeholder)
}

and FormatChange =
    | Bold
    | NoBold
    | Italic
    | NoItalic
    | Underline
    | NoUnderline
    | SetColor of string
    | ClearColor
    | SetFontSize of float

// Helper functions for Run type
module Run =
    let start = function
        | Text tr -> tr.Start
        | Image ir -> ir.Start

    let length = function
        | Text tr -> tr.Length
        | Image _ -> 1  // Images occupy 1 character (U+FFFC placeholder)

    let endPos run = start run + length run

    let shiftStart offset = function
        | Text tr -> Text { tr with Start = tr.Start + offset }
        | Image ir -> Image { ir with Start = ir.Start + offset }

type private RunArrayCacheEntry = {
    Items: Run array
    RunCount: int
    StringCharCount: int
}

type private RunArrayCache = {
    Entries: System.Collections.Generic.Dictionary<Run list, RunArrayCacheEntry>
    Order: System.Collections.Generic.Queue<Run list>
    mutable TotalRunCount: int
    mutable TotalStringCharCount: int
}

let mutable private runArrayCache: RunArrayCache option = None
let private runArrayCacheMaxEntries = 256
let private runArrayCacheMaxTotalRuns = 100000
let private runArrayCacheMaxTotalStringChars = 5000000

let private runListReferenceComparer =
    { new System.Collections.Generic.IEqualityComparer<Run list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }

let private runArrayCacheOptionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private runArrayCacheTextStyleStringCharCount (style: TextStyle) =
    let mutable total = 0
    let add (value: string option) =
        match value with
        | Some value -> total <- total + value.Length
        | None -> ()
    add style.Style
    add style.UnderlineStyle
    add style.UnderlineColor
    add style.EmphasisMark
    add style.FontFamily
    add style.Color
    add style.Highlight
    add style.Hyperlink
    add style.CommentId
    add style.CommentText
    add style.CommentAuthor
    add style.CommentDate
    add style.NoteId
    add style.NoteType
    add style.NoteText
    add style.BookmarkStartId
    add style.BookmarkName
    add style.BookmarkEndId
    add style.FieldCode
    add style.FieldResult
    add style.EquationXml
    add style.RevisionId
    add style.RevisionType
    add style.RevisionAuthor
    add style.RevisionDate
    add style.VertAlign
    total

let private runArrayCacheImageStringCharCount (image: ImageRun) =
    image.Src.Length
    + runArrayCacheOptionStringCharCount image.Alt
    + runArrayCacheOptionStringCharCount image.Name
    + runArrayCacheOptionStringCharCount image.Title
    + runArrayCacheOptionStringCharCount image.Wrap
    + runArrayCacheOptionStringCharCount image.PositionHRelativeFrom
    + runArrayCacheOptionStringCharCount image.PositionVRelativeFrom
    + runArrayCacheOptionStringCharCount image.RelativeHeight
    + runArrayCacheOptionStringCharCount image.CropLeft
    + runArrayCacheOptionStringCharCount image.CropTop
    + runArrayCacheOptionStringCharCount image.CropRight
    + runArrayCacheOptionStringCharCount image.CropBottom

let private runArrayCacheRunStringCharCount = function
    | Text textRun -> runArrayCacheTextStyleStringCharCount textRun.Style
    | Image image -> runArrayCacheImageStringCharCount image

let private runArrayCacheStringCharCount (runs: Run list) =
    let limit = runArrayCacheMaxTotalStringChars + 1
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = runs
    while total <= runArrayCacheMaxTotalStringChars
          && seen <= runArrayCacheMaxTotalRuns
          && not remaining.IsEmpty do
        total <- min limit (total + runArrayCacheRunStringCharCount remaining.Head)
        seen <- seen + 1
        remaining <- remaining.Tail
    min total limit

let private runArrayCacheStringCharCountFor runCount runs =
    if runCount > runArrayCacheMaxTotalRuns then
        runArrayCacheMaxTotalStringChars + 1
    else
        runArrayCacheStringCharCount runs

let private addRunArrayCacheEntry runs items (cache: RunArrayCache) =
    let entry = {
        Items = items
        RunCount = items.Length
        StringCharCount = runArrayCacheStringCharCountFor items.Length runs
    }
    let removeExisting () =
        let mutable existing = Unchecked.defaultof<RunArrayCacheEntry>
        if cache.Entries.TryGetValue(runs, &existing) then
            cache.TotalRunCount <- max 0 (cache.TotalRunCount - existing.RunCount)
            cache.TotalStringCharCount <- max 0 (cache.TotalStringCharCount - existing.StringCharCount)
            cache.Entries.Remove runs |> ignore
    if entry.RunCount > runArrayCacheMaxTotalRuns
       || entry.StringCharCount > runArrayCacheMaxTotalStringChars then
        removeExisting ()
    else
        if cache.Entries.ContainsKey runs then
            removeExisting ()
        else
            cache.Order.Enqueue runs
        cache.Entries.[runs] <- entry
        cache.TotalRunCount <- cache.TotalRunCount + entry.RunCount
        cache.TotalStringCharCount <- cache.TotalStringCharCount + entry.StringCharCount
        while (cache.Entries.Count > runArrayCacheMaxEntries
               || cache.TotalRunCount > runArrayCacheMaxTotalRuns
               || cache.TotalStringCharCount > runArrayCacheMaxTotalStringChars)
              && cache.Order.Count > 0 do
            let oldest = cache.Order.Dequeue()
            if not (obj.ReferenceEquals(oldest, runs)) then
                let mutable oldEntry = Unchecked.defaultof<RunArrayCacheEntry>
                if cache.Entries.TryGetValue(oldest, &oldEntry) then
                    cache.TotalRunCount <- max 0 (cache.TotalRunCount - oldEntry.RunCount)
                    cache.TotalStringCharCount <- max 0 (cache.TotalStringCharCount - oldEntry.StringCharCount)
                    cache.Entries.Remove oldest |> ignore
        if cache.Entries.Count > runArrayCacheMaxEntries
           || cache.TotalRunCount > runArrayCacheMaxTotalRuns
           || cache.TotalStringCharCount > runArrayCacheMaxTotalStringChars then
            let current = cache.Entries.[runs]
            cache.Entries.Clear()
            cache.Order.Clear()
            cache.Order.Enqueue runs
            cache.Entries.[runs] <- current
            cache.TotalRunCount <- current.RunCount
            cache.TotalStringCharCount <- current.StringCharCount

let private runArrayFor runs =
    match runArrayCache with
    | Some cached ->
        let mutable entry = Unchecked.defaultof<RunArrayCacheEntry>
        if cached.Entries.TryGetValue(runs, &entry) then
            entry.Items
        else
            let items = List.toArray runs
            addRunArrayCacheEntry runs items cached
            items
    | _ ->
        let items = List.toArray runs
        let entries = System.Collections.Generic.Dictionary<Run list, RunArrayCacheEntry>(runListReferenceComparer)
        let order = System.Collections.Generic.Queue<Run list>()
        let cache = { Entries = entries; Order = order; TotalRunCount = 0; TotalStringCharCount = 0 }
        addRunArrayCacheEntry runs items cache
        runArrayCache <- Some cache
        items

let private mergedRunList (runs: ResizeArray<Run>) =
    let mutable result = []
    for i = runs.Count - 1 downto 0 do
        match runs.[i], result with
        | Text current, Text next :: rest when current.Start + current.Length = next.Start && current.Style = next.Style ->
            result <- Text { current with Length = current.Length + next.Length } :: rest
        | run, _ ->
            result <- run :: result
    result

let private tryExtendLastTextRun bufferLen textLen (runs: Run list) =
    let values = ResizeArray<Run>()
    let mutable remaining = runs
    let mutable extendedLast = false
    let mutable canExtend = true
    while canExtend && not remaining.IsEmpty do
        match remaining with
        | [ Text tr ] when tr.Start + tr.Length = bufferLen ->
            values.Add(Text { tr with Length = tr.Length + textLen })
            extendedLast <- true
            remaining <- []
        | [ Text _ ] ->
            canExtend <- false
        | [ Image _ ] ->
            canExtend <- false
        | run :: tail ->
            values.Add run
            remaining <- tail
        | [] -> ()
    if extendedLast then Some (resizeArrayToList values) else None

/// Create empty buffer
let empty = {
    Rope = Empty
    Runs = []
}

/// Create from string
let ofString text =
    let len = String.length text
    {
        Rope = Rope.ofString text
        Runs = if len > 0 then [Text { Start = 0; Length = len; Style = defaultStyle }] else []
    }

/// Get length
let length buffer = Rope.length buffer.Rope

/// Get text content
let toString buffer = Rope.toString buffer.Rope

let appendToStringBuilder (sb: System.Text.StringBuilder) buffer =
    Rope.appendToStringBuilder sb buffer.Rope

let appendSliceToStringBuilder (sb: System.Text.StringBuilder) start len buffer =
    Rope.appendSliceToStringBuilder sb start len buffer.Rope

let sliceEqualsString start len (text: string) buffer =
    Rope.sliceEqualsString start len text buffer.Rope

/// Find non-overlapping literal matches without materializing the whole rope.
let iterMatchRanges (matchCase: bool) (query: string) buffer (emit: int -> int -> unit) =
    if System.String.IsNullOrEmpty query || query.Length > length buffer then
        ()
    else
        let normalize ch =
            if matchCase then ch else System.Char.ToUpperInvariant ch
        let firstQueryChar = normalize query.[0]
        let indexOfInString (text: string) start =
#if FABLE_COMPILER
            if start > text.Length - query.Length then
                -1
            elif matchCase then
                text.IndexOf(query, start)
            else
                let charEqualsIgnoreCase a b =
                    a = b || System.Char.ToUpperInvariant a = System.Char.ToUpperInvariant b
                let mutable found = -1
                let mutable index = start
                while found < 0 && index <= text.Length - query.Length do
                    let mutable offset = 0
                    let mutable matches = true
                    while matches && offset < query.Length do
                        if not (charEqualsIgnoreCase text.[index + offset] query.[offset]) then
                            matches <- false
                        offset <- offset + 1
                    if matches then
                        found <- index
                    index <- index + 1
                found
#else
            let comparison = if matchCase then System.StringComparison.Ordinal else System.StringComparison.OrdinalIgnoreCase
            text.IndexOf(query, start, comparison)
#endif

        let mutable nextAllowedStart = 0
        let emitIfAllowed start finish =
            if start >= nextAllowedStart then
                emit start finish
                nextAllowedStart <- finish

        let searchText absoluteStart (text: string) startLocal =
            let mutable index = indexOfInString text startLocal
            while index >= 0 do
                let start = absoluteStart + index
                emitIfAllowed start (start + query.Length)
                let nextLocal = max (index + 1) (nextAllowedStart - absoluteStart)
                index <- indexOfInString text nextLocal

        match buffer.Rope with
        | Leaf text ->
            searchText 0 text 0
        | _ ->
            let carryLimit = query.Length - 1
            let carryBuffer = if carryLimit > 0 then Array.zeroCreate<char> carryLimit else [||]
            let mutable carryLength = 0
            let mutable carryStart = 0
            let equalsQuery ch queryIndex =
                normalize ch = normalize query.[queryIndex]
            let updateCarry absoluteStart (text: string) =
                if carryLimit > 0 then
                    let takeFromText = min carryLimit text.Length
                    if takeFromText = carryLimit then
                        let sourceStart = text.Length - takeFromText
                        for i = 0 to takeFromText - 1 do
                            carryBuffer.[i] <- text.[sourceStart + i]
                        carryLength <- takeFromText
                        carryStart <- absoluteStart + sourceStart
                    else
                        let keepFromCarry = min carryLength (carryLimit - takeFromText)
                        let previousCarryLength = carryLength
                        let previousCarryStart = carryStart
                        if keepFromCarry > 0 then
                            let sourceStart = previousCarryLength - keepFromCarry
                            for i = 0 to keepFromCarry - 1 do
                                carryBuffer.[i] <- carryBuffer.[sourceStart + i]
                        for i = 0 to takeFromText - 1 do
                            carryBuffer.[keepFromCarry + i] <- text.[i]
                        carryLength <- keepFromCarry + takeFromText
                        carryStart <-
                            if keepFromCarry > 0 then previousCarryStart + previousCarryLength - keepFromCarry
                            else absoluteStart

            let mutable absolute = 0
            let rec walk rope =
                match rope with
                | Empty -> ()
                | Leaf text ->
                    let leafStart = absolute
                    if carryLimit > 0 && carryLength > 0 && text.Length > 0 then
                        let mutable startInCarry = 0
                        while startInCarry < carryLength do
                            let carryAvailable = carryLength - startInCarry
                            if carryAvailable < query.Length
                               && carryAvailable + text.Length >= query.Length
                               && normalize carryBuffer.[startInCarry] = firstQueryChar then
                                let mutable queryIndex = 0
                                let mutable matches = true
                                while matches && queryIndex < query.Length do
                                    let ch =
                                        if queryIndex < carryAvailable then
                                            carryBuffer.[startInCarry + queryIndex]
                                        else
                                            text.[queryIndex - carryAvailable]
                                    if not (equalsQuery ch queryIndex) then
                                        matches <- false
                                    queryIndex <- queryIndex + 1
                                if matches then
                                    let start = carryStart + startInCarry
                                    emitIfAllowed start (start + query.Length)
                            startInCarry <- startInCarry + 1

                    let startLocal = max 0 (nextAllowedStart - leafStart)
                    searchText leafStart text startLocal
                    updateCarry leafStart text
                    absolute <- absolute + text.Length
                | Node (left, right, _) ->
                    walk left
                    walk right
            walk buffer.Rope

/// True when the buffer contains no non-whitespace characters, without
/// materializing the rope as one string.
let isWhiteSpace buffer =
    let rec ropeIsWhiteSpace rope =
        match rope with
        | Empty -> true
        | Leaf text ->
            let mutable index = 0
            let mutable onlyWhitespace = true
            while onlyWhitespace && index < text.Length do
                if not (System.Char.IsWhiteSpace text.[index]) then
                    onlyWhitespace <- false
                index <- index + 1
            onlyWhitespace
        | Node (left, right, _) ->
            ropeIsWhiteSpace left && ropeIsWhiteSpace right
    ropeIsWhiteSpace buffer.Rope

/// True when the buffer is non-empty and contains only whitespace or NBSP,
/// without materializing the rope as one string.
let isNonbreakingBlank buffer =
    let isWhitespaceOrNbsp ch =
        System.Char.IsWhiteSpace ch || ch = '\u00A0'
    let rec ropeIsBlank rope =
        match rope with
        | Empty -> true
        | Leaf text ->
            let mutable index = 0
            let mutable onlyBlank = true
            while onlyBlank && index < text.Length do
                if not (isWhitespaceOrNbsp text.[index]) then
                    onlyBlank <- false
                index <- index + 1
            onlyBlank
        | Node (left, right, _) ->
            ropeIsBlank left && ropeIsBlank right
    length buffer > 0 && ropeIsBlank buffer.Rope

/// Insert text at position
let insert pos (text: string) buffer =
    if String.length text = 0 then buffer
    else
        let textLen = String.length text
        let bufferLen = length buffer

        let insertByScan () =
            let newRuns = ResizeArray<Run>()
            let mutable inserted = false
            for run in buffer.Runs do
                let runStart = Run.start run
                let runEnd = Run.endPos run
                if not inserted && pos >= runStart && pos <= runEnd then
                    // Insert within this run
                    match run with
                    | Image _ ->
                        // Can't split image - insert before it
                        newRuns.Add(Text { Start = pos; Length = textLen; Style = defaultStyle })
                        newRuns.Add(Run.shiftStart textLen run)
                    | Text tr ->
                        if pos = runStart then
                            // Insert at start of run
                            newRuns.Add(Text { Start = tr.Start; Length = textLen; Style = tr.Style })
                            newRuns.Add(Text { tr with Start = tr.Start + textLen })
                        elif pos = runEnd then
                            // Insert at end of run
                            newRuns.Add(run)
                            newRuns.Add(Text { Start = runEnd; Length = textLen; Style = tr.Style })
                        else
                            // Insert in middle - split into three runs
                            let beforeLen = pos - tr.Start
                            let afterLen = tr.Length - beforeLen
                            newRuns.Add(Text { tr with Length = beforeLen })
                            newRuns.Add(Text { Start = pos; Length = textLen; Style = tr.Style })
                            newRuns.Add(Text { Start = pos + textLen; Length = afterLen; Style = tr.Style })
                    inserted <- true
                elif inserted then
                    // Shift runs after insertion
                    newRuns.Add(Run.shiftStart textLen run)
                else
                    // Before insertion point
                    newRuns.Add(run)

            // If no runs or insertion at end, add new run
            if not inserted then
                newRuns.Add(Text { Start = pos; Length = textLen; Style = defaultStyle })

            { buffer with
                Rope = Rope.insert pos text buffer.Rope
                Runs = mergedRunList newRuns }

        if bufferLen = 0 && pos = 0 then
            ofString text
        elif pos = bufferLen then
            match buffer.Runs with
            | [ Text tr ] when tr.Start = 0 && tr.Length = bufferLen ->
                { buffer with
                    Rope = Rope.insert pos text buffer.Rope
                    Runs = [ Text { tr with Length = tr.Length + textLen } ] }
            | _ ->
                match tryExtendLastTextRun bufferLen textLen buffer.Runs with
                | Some runs ->
                    { buffer with
                        Rope = Rope.insert pos text buffer.Rope
                        Runs = runs }
                | None -> insertByScan ()
        else
            insertByScan ()

/// Delete range [start, start + len)
let delete start len buffer =
    if len <= 0 then buffer
    else
        let bufferLen = length buffer
        let deleteEnd = start + len
        let actualStart = max 0 (min bufferLen start)
        let actualEnd = max actualStart (min bufferLen deleteEnd)
        let actualLen = actualEnd - actualStart

        if actualLen <= 0 then
            buffer
        else
            let start = actualStart
            let len = actualLen
            let deleteEnd = actualEnd
            match buffer.Runs with
            | [ Text tr ] when tr.Start = 0 && tr.Length = bufferLen ->
                let remaining = bufferLen - actualLen
                { buffer with
                    Rope = Rope.delete actualStart actualLen buffer.Rope
                    Runs = if remaining > 0 then [ Text { tr with Length = remaining } ] else [] }
            | _ ->
                let newRuns = ResizeArray<Run>()

                for run in buffer.Runs do
                    let runStart = Run.start run
                    let runEnd = Run.endPos run
                    if runEnd <= start then
                        // Run is entirely before deletion
                        newRuns.Add(run)
                    elif runStart >= deleteEnd then
                        // Run is entirely after deletion - shift back
                        newRuns.Add(Run.shiftStart -len run)
                    elif runStart >= start && runEnd <= deleteEnd then
                        // Run is entirely within deletion - remove it
                        ()
                    elif runStart < start && runEnd > deleteEnd then
                        // Deletion is within run - shrink it
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr -> newRuns.Add(Text { tr with Length = tr.Length - len })
                    elif runStart < start then
                        // Run starts before, ends within deletion
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr ->
                            let newLen = start - tr.Start
                            newRuns.Add(Text { tr with Length = newLen })
                    else
                        // Run starts within deletion, ends after
                        match run with
                        | Image _ -> () // Image deleted
                        | Text tr ->
                            let overlap = deleteEnd - tr.Start
                            let newLen = tr.Length - overlap
                            newRuns.Add(Text { Start = start; Length = newLen; Style = tr.Style })

                { buffer with
                    Rope = Rope.delete start len buffer.Rope
                    Runs = mergedRunList newRuns }

/// Get substring
let slice start len buffer =
    Rope.slice start len buffer.Rope

/// Get a buffer slice while preserving text styles and image metadata.
let sliceBuffer start len buffer =
    if len <= 0 then empty
    else
        let bufferLen = length buffer
        let sliceStart = max 0 (min start bufferLen)
        let sliceEnd = max sliceStart (min bufferLen (sliceStart + len))
        if sliceStart = 0 && sliceEnd = bufferLen then
            buffer
        else
            let newRuns = ResizeArray<Run>()
            for run in buffer.Runs do
                let runStart = Run.start run
                let runEnd = Run.endPos run
                let overlapStart = max runStart sliceStart
                let overlapEnd = min runEnd sliceEnd
                if overlapStart < overlapEnd then
                    match run with
                    | Text tr ->
                        newRuns.Add(Text { tr with Start = overlapStart - sliceStart; Length = overlapEnd - overlapStart })
                    | Image ir ->
                        newRuns.Add(Image { ir with Start = ir.Start - sliceStart })
            { Rope = Rope.sliceRope sliceStart (sliceEnd - sliceStart) buffer.Rope
              Runs = mergedRunList newRuns }

/// Append two buffers without flattening either side to a string.
let append left right =
    let leftLen = length left
    if leftLen = 0 then right
    elif length right = 0 then left
    else
        let rightLen = length right
        match left.Runs, right.Runs with
        | [ Text leftRun ], [ Text rightRun ] when leftRun.Start = 0 && rightRun.Start = 0 && leftRun.Length = leftLen && rightRun.Length = rightLen ->
            let runs =
                if leftRun.Style = rightRun.Style then
                    [ Text { leftRun with Length = leftLen + rightLen } ]
                else
                    [ Text leftRun; Text { rightRun with Start = leftLen } ]
            { Rope = Rope.balance left.Rope right.Rope
              Runs = runs }
        | _ ->
            let runs = ResizeArray<Run>()
            for run in left.Runs do
                runs.Add run
            for run in right.Runs do
                runs.Add(Run.shiftStart leftLen run)
            { Rope = Rope.balance left.Rope right.Rope
              Runs = mergedRunList runs }

let trySliceNoCopyV start len buffer =
    match buffer.Rope with
    | Leaf text when start = 0 && len = text.Length -> ValueSome text
    | _ -> ValueNone

let trySliceNoCopy start len buffer =
    match trySliceNoCopyV start len buffer with
    | ValueSome text -> Some text
    | ValueNone -> None

let private styleAt pos buffer =
    if pos < 0 || pos >= length buffer then
        defaultStyle
    else
        match buffer.Runs with
        | [ Text tr ] when pos >= tr.Start && pos < tr.Start + tr.Length -> tr.Style
        | _ ->
            let runs = runArrayFor buffer.Runs
            let mutable lo = 0
            let mutable hi = runs.Length - 1
            let mutable found = defaultStyle
            let mutable searching = true
            while searching && lo <= hi do
                let mid = lo + ((hi - lo) / 2)
                let run = runs.[mid]
                let runStart = Run.start run
                if pos < runStart then
                    hi <- mid - 1
                else
                    let runEnd = runStart + Run.length run
                    if pos >= runEnd then
                        lo <- mid + 1
                    else
                        match run with
                        | Text tr -> found <- tr.Style
                        | Image _ -> ()
                        searching <- false
            found

/// Replace sorted, non-overlapping ranges in one pass using text already
/// materialized by the caller. Replacement text uses the same insertion-side
/// style as repeated delete+insert at each range start.
let replaceRangesWithText (originalText: string) (ranges: ResizeArray<struct (int * int)>) (replacement: string) buffer =
    if ranges.Count = 0 then
        buffer
    else
        let originalLength = originalText.Length
        let replacementLength = replacement.Length
        let mutable removedLength = 0
        let mutable sizingCursor = 0
        for i = 0 to ranges.Count - 1 do
            let struct (rawStart, rawFinish) = ranges.[i]
            let start = max sizingCursor (min originalLength rawStart)
            let finish = max start (min originalLength rawFinish)
            removedLength <- removedLength + (finish - start)
            sizingCursor <- finish
        let estimatedLength = int (min (int64 System.Int32.MaxValue) (max 0L (int64 originalLength - int64 removedLength + int64 ranges.Count * int64 replacementLength)))
        let text = System.Text.StringBuilder(estimatedLength)
        let appendTextSlice start finish =
            if finish > start then
#if FABLE_COMPILER
                text.Append(originalText.Substring(start, finish - start)) |> ignore
#else
                text.Append(originalText, start, finish - start) |> ignore
#endif
        match buffer.Runs with
        | [ Text tr ] when tr.Start = 0 && tr.Length = originalLength ->
            let mutable cursor = 0
            for i = 0 to ranges.Count - 1 do
                let struct (rawStart, rawFinish) = ranges.[i]
                let start = max cursor (min originalLength rawStart)
                let finish = max start (min originalLength rawFinish)
                appendTextSlice cursor start
                if replacementLength > 0 then
                    text.Append replacement |> ignore
                cursor <- finish
            appendTextSlice cursor originalLength
            { Rope = Rope.ofString (text.ToString())
              Runs = if estimatedLength > 0 then [ Text { tr with Length = estimatedLength } ] else [] }
        | _ ->
            let runs = buffer.Runs |> List.toArray
            let newRuns = ResizeArray<Run>(runs.Length + ranges.Count)
            let mutable copyRunIndex = 0
            let addCopiedRuns rangeStart rangeEnd outStart =
                if rangeEnd > rangeStart then
                    while copyRunIndex < runs.Length && Run.endPos runs.[copyRunIndex] <= rangeStart do
                        copyRunIndex <- copyRunIndex + 1
                    let mutable index = copyRunIndex
                    let mutable copying = true
                    while copying && index < runs.Length && Run.start runs.[index] < rangeEnd do
                        let run = runs.[index]
                        let runStart = Run.start run
                        let runEnd = Run.endPos run
                        let overlapStart = max runStart rangeStart
                        let overlapEnd = min runEnd rangeEnd
                        if overlapStart < overlapEnd then
                            let shiftedStart = outStart + (overlapStart - rangeStart)
                            match run with
                            | Text tr ->
                                newRuns.Add(Text { tr with Start = shiftedStart; Length = overlapEnd - overlapStart })
                            | Image ir ->
                                newRuns.Add(Image { ir with Start = shiftedStart })
                        if runEnd <= rangeEnd then
                            index <- index + 1
                        else
                            copying <- false
                    copyRunIndex <- index
            let mutable styleRunIndex = 0
            let styleAtAscending pos =
                while styleRunIndex < runs.Length && Run.endPos runs.[styleRunIndex] <= pos do
                    styleRunIndex <- styleRunIndex + 1
                if styleRunIndex < runs.Length && pos >= Run.start runs.[styleRunIndex] then
                    match runs.[styleRunIndex] with
                    | Text tr -> tr.Style
                    | Image _ -> defaultStyle
                else
                    defaultStyle
            let mutable cursor = 0
            let mutable outPos = 0
            for i = 0 to ranges.Count - 1 do
                let struct (rawStart, rawFinish) = ranges.[i]
                let start = max cursor (min originalLength rawStart)
                let finish = max start (min originalLength rawFinish)
                appendTextSlice cursor start
                addCopiedRuns cursor start outPos
                outPos <- outPos + (start - cursor)
                if replacementLength > 0 then
                    let style =
                        if start > 0 then styleAtAscending (start - 1)
                        elif start < originalLength then styleAtAscending start
                        else defaultStyle
                    text.Append replacement |> ignore
                    newRuns.Add(Text { Start = outPos; Length = replacementLength; Style = style })
                    outPos <- outPos + replacementLength
                cursor <- finish
            appendTextSlice cursor originalLength
            addCopiedRuns cursor originalLength outPos
            { Rope = Rope.ofString (text.ToString())
              Runs = mergedRunList newRuns }

/// Replace sorted, non-overlapping ranges in one pass. Replacement text uses
/// the same insertion-side style as repeated delete+insert at each range start.
let replaceRanges (ranges: ResizeArray<struct (int * int)>) (replacement: string) buffer =
    replaceRangesWithText (toString buffer) ranges replacement buffer

/// Object Replacement Character - used as placeholder for images
let objectReplacementChar = "\uFFFC"

/// Resize image at position
let resizeImage pos newWidth newHeight buffer =
    if pos < 0 || pos >= length buffer then
        buffer
    else
        let mutable found = false
        let mutable unchanged = false
        let mutable remaining = buffer.Runs
        while not found && not remaining.IsEmpty do
            match remaining with
            | Image imgRun :: tail when imgRun.Start = pos ->
                found <- true
                let widthUnchanged =
                    match imgRun.Width with
                    | Some width -> width = newWidth
                    | None -> false
                let heightUnchanged =
                    match imgRun.Height with
                    | Some height -> height = newHeight
                    | None -> false
                unchanged <- widthUnchanged && heightUnchanged
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()

        if not found || unchanged then
            buffer
        else
            let newRuns = ResizeArray<Run>()
            for run in buffer.Runs do
                match run with
                | Image imgRun when imgRun.Start = pos ->
                    newRuns.Add(Image { imgRun with Width = Some newWidth; Height = Some newHeight })
                | _ ->
                    newRuns.Add run
            { buffer with Runs = resizeArrayToList newRuns }

/// Insert image at position
let insertImageWithMetadata pos src linked width height alt name title floating wrap positionX positionY positionHRelativeFrom positionVRelativeFrom behindDoc allowOverlap layoutInCell distanceTop distanceBottom distanceLeft distanceRight relativeHeight locked simplePosition simplePositionX simplePositionY cropLeft cropTop cropRight cropBottom buffer =
    // Insert placeholder character in rope
    let newRope = Rope.insert pos objectReplacementChar buffer.Rope

    // Create image run
    let imageRun = Image {
        Start = pos
        Src = src
        Linked = linked
        Width = width
        Height = height
        Alt = alt
        Name = name
        Title = title
        Floating = floating
        Wrap = wrap
        PositionX = positionX
        PositionY = positionY
        PositionHRelativeFrom = positionHRelativeFrom
        PositionVRelativeFrom = positionVRelativeFrom
        BehindDoc = behindDoc
        AllowOverlap = allowOverlap
        LayoutInCell = layoutInCell
        DistanceTop = distanceTop
        DistanceBottom = distanceBottom
        DistanceLeft = distanceLeft
        DistanceRight = distanceRight
        RelativeHeight = relativeHeight
        Locked = locked
        SimplePosition = simplePosition
        SimplePositionX = simplePositionX
        SimplePositionY = simplePositionY
        CropLeft = cropLeft
        CropTop = cropTop
        CropRight = cropRight
        CropBottom = cropBottom
    }

    // Insert image run and shift subsequent runs
    let newRuns = ResizeArray<Run>()
    let mutable inserted = false
    for run in buffer.Runs do
        let runStart = Run.start run
        let runEnd = Run.endPos run
        if inserted then
            newRuns.Add(Run.shiftStart 1 run)
        elif pos < runStart then
            newRuns.Add imageRun
            newRuns.Add(Run.shiftStart 1 run)
            inserted <- true
        elif pos >= runStart && pos <= runEnd then
            match run with
            | Image _ ->
                if pos <= runStart then
                    newRuns.Add imageRun
                    newRuns.Add(Run.shiftStart 1 run)
                else
                    newRuns.Add run
                    newRuns.Add imageRun
                inserted <- true
            | Text tr ->
                let beforeLen = pos - tr.Start
                let afterLen = tr.Length - beforeLen
                if beforeLen > 0 then
                    newRuns.Add(Text { tr with Length = beforeLen })
                newRuns.Add imageRun
                if afterLen > 0 then
                    newRuns.Add(Text { Start = pos + 1; Length = afterLen; Style = tr.Style })
                inserted <- true
        else
            newRuns.Add run

    if not inserted then
        newRuns.Add imageRun

    { Rope = newRope; Runs = resizeArrayToList newRuns }

let insertImage pos src width height alt buffer =
    insertImageWithMetadata pos src false width height alt None None false None None None None None None None None None None None None None None None None None None None None None buffer

let tryMoveImageWithinBuffer offset insertOffset (img: ImageRun) buffer =
    let bufferAfterDeleteLen = max 0 (length buffer - 1)
    let insertOffset = max 0 (min bufferAfterDeleteLen insertOffset)
    if insertOffset = offset then
        Some buffer
    else
        let mutable canUseFastPath = true
        let newRuns = ResizeArray<Run>()
        let moved = Image { img with Start = insertOffset }
        let mutable inserted = false
        let mutable removed = false

        let addMovedIfBefore startAfterDelete =
            if canUseFastPath && not inserted && insertOffset <= startAfterDelete then
                newRuns.Add moved
                inserted <- true

        for run in buffer.Runs do
            match run with
            | Image image when image.Start = offset ->
                removed <- true
            | _ ->
                if canUseFastPath then
                    let start = Run.start run
                    let startAfterDelete = if start > offset then start - 1 else start
                    match run with
                    | Text tr ->
                        let endAfterDelete = startAfterDelete + tr.Length
                        if insertOffset > startAfterDelete && insertOffset < endAfterDelete then
                            canUseFastPath <- false
                    | Image _ -> ()
                    if canUseFastPath then
                        addMovedIfBefore startAfterDelete
                        let finalStart = if inserted && startAfterDelete >= insertOffset then startAfterDelete + 1 else startAfterDelete
                        newRuns.Add(Run.shiftStart (finalStart - start) run)

        if canUseFastPath && removed then
            if not inserted then
                newRuns.Add moved
            Some {
                Rope = buffer.Rope |> Rope.delete offset 1 |> Rope.insert insertOffset objectReplacementChar
                Runs = resizeArrayToList newRuns
            }
        else
            None

/// Apply style to range [start, start + len)
let applyStyle start len styleTransform buffer =
    if len <= 0 then buffer
    else
        let rangeEnd = start + len

        match buffer.Runs with
        | [ Text tr ] when tr.Start = 0 && tr.Length = length buffer ->
            let runStart = tr.Start
            let runEnd = tr.Start + tr.Length
            if runEnd <= start || runStart >= rangeEnd then
                buffer
            elif runStart >= start && runEnd <= rangeEnd then
                let transformedStyle = styleTransform tr.Style
                if transformedStyle = tr.Style then buffer
                else { buffer with Runs = [ Text { tr with Style = transformedStyle } ] }
            else
                let transformedStyle = styleTransform tr.Style
                if transformedStyle = tr.Style then
                    buffer
                else
                    let runs = ResizeArray<Run>(3)
                    let addText start length style =
                        if length > 0 then
                            runs.Add(Text { Start = start; Length = length; Style = style })
                    let overlapStart = max runStart start
                    let overlapEnd = min runEnd rangeEnd
                    addText runStart (overlapStart - runStart) tr.Style
                    addText overlapStart (overlapEnd - overlapStart) transformedStyle
                    addText overlapEnd (runEnd - overlapEnd) tr.Style
                    { buffer with Runs = mergedRunList runs }
        | _ ->
            let bufferLen = length buffer
            let fullRangeApply () =
                let newRuns = ResizeArray<Run>()
                let mutable changed = false
                let mutable hasPendingText = false
                let mutable pendingStart = 0
                let mutable pendingLength = 0
                let mutable pendingStyle = defaultStyle
                let flushPendingText () =
                    if hasPendingText then
                        newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                        hasPendingText <- false
                let addText start length style =
                    if length > 0 then
                        if hasPendingText && pendingStart + pendingLength = start && pendingStyle = style then
                            pendingLength <- pendingLength + length
                        else
                            flushPendingText ()
                            pendingStart <- start
                            pendingLength <- length
                            pendingStyle <- style
                            hasPendingText <- true
                for run in buffer.Runs do
                    match run with
                    | Text tr ->
                        let transformedStyle = styleTransform tr.Style
                        changed <- changed || transformedStyle <> tr.Style
                        addText tr.Start tr.Length transformedStyle
                    | Image _ ->
                        flushPendingText ()
                        newRuns.Add run
                if changed then
                    flushPendingText ()
                    { buffer with Runs = resizeArrayToList newRuns }
                else
                    buffer

            let partialRangeApply () =
                let mutable changed = false
                let mutable prefixCount = 0
                let mutable newRuns = Unchecked.defaultof<ResizeArray<Run>>
                let copyPrefix () =
                    let mutable copied = 0
                    let mutable prefix = buffer.Runs
                    while copied < prefixCount do
                        match prefix with
                        | run :: tail ->
                            newRuns.Add run
                            prefix <- tail
                            copied <- copied + 1
                        | [] -> copied <- prefixCount
                let ensureChanged () =
                    if not changed then
                        changed <- true
                        newRuns <- ResizeArray<Run>()
                        copyPrefix ()
                let addChangedText (tr: TextRun) runStart runEnd transformedStyle =
                    if runEnd <= start || runStart >= rangeEnd then
                        newRuns.Add(Text tr)
                    elif runStart >= start && runEnd <= rangeEnd then
                        newRuns.Add(Text { tr with Style = transformedStyle })
                    elif runStart < start && runEnd > rangeEnd then
                        let beforeLen = start - tr.Start
                        let rangeLen = len
                        let afterLen = runEnd - rangeEnd
                        newRuns.Add(Text { tr with Length = beforeLen })
                        newRuns.Add(Text { Start = start; Length = rangeLen; Style = transformedStyle })
                        newRuns.Add(Text { Start = rangeEnd; Length = afterLen; Style = tr.Style })
                    elif runStart < start then
                        let beforeLen = start - tr.Start
                        let overlapLen = runEnd - start
                        newRuns.Add(Text { tr with Length = beforeLen })
                        newRuns.Add(Text { Start = start; Length = overlapLen; Style = transformedStyle })
                    else
                        let overlapLen = rangeEnd - tr.Start
                        let afterLen = runEnd - rangeEnd
                        newRuns.Add(Text { Start = tr.Start; Length = overlapLen; Style = transformedStyle })
                        newRuns.Add(Text { Start = rangeEnd; Length = afterLen; Style = tr.Style })

                let mutable remaining = buffer.Runs
                while not remaining.IsEmpty do
                    match remaining with
                    | run :: _ when Run.start run >= rangeEnd ->
                        if changed then
                            let mutable rest = remaining
                            while not rest.IsEmpty do
                                match rest with
                                | tailRun :: tail ->
                                    newRuns.Add tailRun
                                    rest <- tail
                                | [] -> ()
                        remaining <- []
                    | Image _ as run :: tail ->
                        if changed then newRuns.Add run
                        else prefixCount <- prefixCount + 1
                        remaining <- tail
                    | Text tr as run :: tail ->
                        let runStart = tr.Start
                        let runEnd = tr.Start + tr.Length
                        if runEnd <= start || runStart >= rangeEnd then
                            if changed then newRuns.Add run
                            else prefixCount <- prefixCount + 1
                        else
                            let transformedStyle = styleTransform tr.Style
                            if transformedStyle = tr.Style then
                                if changed then newRuns.Add run
                                else prefixCount <- prefixCount + 1
                            else
                                ensureChanged ()
                                addChangedText tr runStart runEnd transformedStyle
                        remaining <- tail
                    | [] -> ()

                if changed then { buffer with Runs = mergedRunList newRuns }
                else buffer

            if start <= 0 && rangeEnd >= bufferLen then fullRangeApply ()
            else partialRangeApply ()

let private textRunsInRangeSatisfy start len predicate buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- predicate tr.Style
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveVertAlign start len align buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.VertAlign = align
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveColor start len color buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Color = color
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveHighlight start len color buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Highlight = color
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveHyperlink start len target buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Hyperlink = target
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveCharacterStyle start len styleId buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.Style = styleId
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveFontSize start len size buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    allMatch <- tr.Style.FontSize = size
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

let private textRunsInRangeHaveFontFamily start len family buffer =
    if len <= 0 then
        true
    else
        let rangeEnd = start + len
        let mutable allMatch = true
        let mutable remaining = buffer.Runs
        while allMatch && not remaining.IsEmpty do
            match remaining with
            | Text tr :: tail ->
                let runStart = tr.Start
                let runEnd = tr.Start + tr.Length
                if runEnd > start && runStart < rangeEnd then
                    match tr.Style.FontFamily with
                    | Some existingFamily -> allMatch <- existingFamily = family
                    | None -> allMatch <- false
                remaining <- tail
            | _ :: tail ->
                remaining <- tail
            | [] -> ()
        allMatch

/// Toggle bold for range
let toggleBold start len buffer =
    applyStyle start len (fun style -> { style with Bold = not style.Bold }) buffer

/// Toggle italic for range
let toggleItalic start len buffer =
    applyStyle start len (fun style -> { style with Italic = not style.Italic }) buffer

/// Toggle underline for range
let toggleUnderline start len buffer =
    applyStyle start len (fun style -> { style with Underline = not style.Underline }) buffer

/// Toggle all caps for range
let toggleAllCaps start len buffer =
    applyStyle start len (fun style -> { style with AllCaps = not style.AllCaps; SmallCaps = false }) buffer

/// Toggle small caps for range
let toggleSmallCaps start len buffer =
    applyStyle start len (fun style -> { style with SmallCaps = not style.SmallCaps; AllCaps = false }) buffer

/// Toggle strikethrough for range
let toggleStrike start len buffer =
    applyStyle start len (fun style -> { style with Strike = not style.Strike }) buffer

/// Set vertical alignment for range
let setVertAlign start len align buffer =
    if textRunsInRangeHaveVertAlign start len align buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with VertAlign = align }) buffer

/// Set text color for range
let setColor start len color buffer =
    if textRunsInRangeHaveColor start len color buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Color = color }) buffer

/// Set highlight color for range
let setHighlight start len color buffer =
    if textRunsInRangeHaveHighlight start len color buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Highlight = color }) buffer

/// Set hyperlink target for range
let setHyperlink start len target buffer =
    if textRunsInRangeHaveHyperlink start len target buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Hyperlink = target }) buffer

/// Set comment metadata for range
let setComment start len commentId commentText author date buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.CommentId = commentId
            && style.CommentText = commentText
            && style.CommentAuthor = author
            && style.CommentDate = date) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with CommentId = commentId; CommentText = commentText; CommentAuthor = author; CommentDate = date }) buffer

/// Set footnote/endnote metadata for range
let setNote start len noteId noteType noteText buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.NoteId = noteId
            && style.NoteType = noteType
            && style.NoteText = noteText) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with NoteId = noteId; NoteType = noteType; NoteText = noteText }) buffer

/// Set bookmark markers for range
let setBookmark start len bookmarkId bookmarkName buffer =
    applyStyle start len (fun style -> { style with BookmarkStartId = Some bookmarkId; BookmarkName = Some bookmarkName; BookmarkEndId = Some bookmarkId }) buffer

/// Set named character style for range
let setCharacterStyle start len styleId buffer =
    if textRunsInRangeHaveCharacterStyle start len styleId buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with Style = styleId }) buffer

/// Set revision metadata for range
let setRevision start len revisionType revisionId author date buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        style.RevisionType = Some revisionType
            && style.RevisionId = Some revisionId
            && style.RevisionAuthor = author
            && style.RevisionDate = date) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with RevisionType = Some revisionType; RevisionId = Some revisionId; RevisionAuthor = author; RevisionDate = date }) buffer

/// Clear revision metadata for range
let clearRevision start len buffer =
    if textRunsInRangeSatisfy start len (fun style ->
        Option.isNone style.RevisionType
            && Option.isNone style.RevisionId
            && Option.isNone style.RevisionAuthor
            && Option.isNone style.RevisionDate) buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None }) buffer

let private clearRevisionStyle style =
    { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None }

/// Resolve revision markup in one pass. Accepted insertions / rejected deletions
/// become normal text; accepted deletions / rejected insertions are removed.
let resolveRevisionsInRange accept selStart selEnd isCollapsed buffer =
    let intersects start endPos =
        if isCollapsed then selStart >= start && selStart <= endPos
        else start < selEnd && endPos > selStart
    let shouldDelete revisionType =
        match accept, revisionType with
        | true, "delete" -> true
        | false, "insert" -> true
        | _ -> false
    let mutable changed = false
    let mutable firstStart = None
    let mutable deletesText = false
    let mutable keepsContent = false
    let mutable remaining = buffer.Runs
    while not remaining.IsEmpty do
        match remaining with
        | Text tr :: tail ->
            match tr.Style.RevisionType with
            | Some _ when intersects tr.Start (tr.Start + tr.Length) ->
                if not changed then
                    changed <- true
                    firstStart <- Some tr.Start
                match tr.Style.RevisionType with
                | Some revisionType when shouldDelete revisionType -> deletesText <- true
                | _ -> keepsContent <- keepsContent || tr.Length > 0
                remaining <- tail
            | _ ->
                keepsContent <- keepsContent || tr.Length > 0
                remaining <- tail
        | Image _ :: tail ->
            keepsContent <- true
            remaining <- tail
        | [] -> ()
    if not changed then
        struct (buffer, None, false)
    elif deletesText && not keepsContent then
        struct (empty, firstStart, true)
    elif not deletesText then
        let clearedStyles = ResizeArray<TextStyle * TextStyle>()
        let clearCached style =
            let mutable found = ValueNone
            let mutable index = 0
            while ValueOption.isNone found && index < clearedStyles.Count do
                let original, cleared = clearedStyles.[index]
                if original = style then
                    found <- ValueSome cleared
                index <- index + 1
            match found with
            | ValueSome cleared -> cleared
            | ValueNone ->
                let cleared = clearRevisionStyle style
                clearedStyles.Add(style, cleared)
                cleared
        let newRuns = ResizeArray<Run>()
        let mutable hasPendingText = false
        let mutable pendingStart = 0
        let mutable pendingLength = 0
        let mutable pendingStyle = defaultStyle
        let sameStyle left right = obj.ReferenceEquals(left, right) || left = right
        let flushPendingText () =
            if hasPendingText then
                newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                hasPendingText <- false
        let addText start length style =
            if length > 0 then
                if hasPendingText && pendingStart + pendingLength = start && sameStyle pendingStyle style then
                    pendingLength <- pendingLength + length
                else
                    flushPendingText ()
                    pendingStart <- start
                    pendingLength <- length
                    pendingStyle <- style
                    hasPendingText <- true
        let addRun run =
            match run with
            | Text tr -> addText tr.Start tr.Length tr.Style
            | Image _ ->
                flushPendingText ()
                newRuns.Add run
        for run in buffer.Runs do
            match run with
            | Text tr ->
                match tr.Style.RevisionType with
                | Some _ when intersects tr.Start (tr.Start + tr.Length) ->
                    addText tr.Start tr.Length (clearCached tr.Style)
                | _ ->
                    addRun run
            | Image _ ->
                addRun run
        flushPendingText ()
        struct ({ buffer with Runs = resizeArrayToList newRuns }, firstStart, true)
    else
        let clearedStyles = ResizeArray<TextStyle * TextStyle>()
        let clearCached style =
            let mutable found = ValueNone
            let mutable index = 0
            while ValueOption.isNone found && index < clearedStyles.Count do
                let original, cleared = clearedStyles.[index]
                if original = style then
                    found <- ValueSome cleared
                index <- index + 1
            match found with
            | ValueSome cleared -> cleared
            | ValueNone ->
                let cleared = clearRevisionStyle style
                clearedStyles.Add(style, cleared)
                cleared
        let originalText = toString buffer
        let text = System.Text.StringBuilder(originalText.Length)
        let newRuns = ResizeArray<Run>()
        let mutable outPos = 0
        let mutable hasPendingText = false
        let mutable pendingStart = 0
        let mutable pendingLength = 0
        let mutable pendingStyle = defaultStyle
        let sameStyle left right = obj.ReferenceEquals(left, right) || left = right
        let flushPendingText () =
            if hasPendingText then
                newRuns.Add(Text { Start = pendingStart; Length = pendingLength; Style = pendingStyle })
                hasPendingText <- false
        let addBufferedText start length style =
            if length > 0 then
                if hasPendingText && pendingStart + pendingLength = start && sameStyle pendingStyle style then
                    pendingLength <- pendingLength + length
                else
                    flushPendingText ()
                    pendingStart <- start
                    pendingLength <- length
                    pendingStyle <- style
                    hasPendingText <- true
        let appendOriginal start len =
            if len > 0 then
#if FABLE_COMPILER
                text.Append(originalText.Substring(start, len)) |> ignore
#else
                text.Append(originalText, start, len) |> ignore
#endif
        let addTextRun (tr: TextRun) style =
            if tr.Length > 0 then
                appendOriginal tr.Start tr.Length
                addBufferedText outPos tr.Length style
                outPos <- outPos + tr.Length
        let addImageRun (ir: ImageRun) =
            flushPendingText ()
            text.Append(objectReplacementChar) |> ignore
            newRuns.Add(Image { ir with Start = outPos })
            outPos <- outPos + 1
        for run in buffer.Runs do
            match run with
            | Text tr ->
                match tr.Style.RevisionType with
                | Some revisionType when intersects tr.Start (tr.Start + tr.Length) ->
                    if firstStart.IsNone then
                        firstStart <- Some tr.Start
                    if not (shouldDelete revisionType) then
                        addTextRun tr (clearCached tr.Style)
                | _ ->
                    addTextRun tr tr.Style
            | Image ir ->
                addImageRun ir
        flushPendingText ()
        struct ({ Rope = Rope.ofString (text.ToString()); Runs = resizeArrayToList newRuns }, firstStart, true)

/// Set font size for range
let setFontSize start len size buffer =
    if textRunsInRangeHaveFontSize start len size buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with FontSize = size }) buffer

/// Set font family for range
let setFontFamily start len family buffer =
    if textRunsInRangeHaveFontFamily start len family buffer then
        buffer
    else
        applyStyle start len (fun style -> { style with FontFamily = Some family }) buffer

/// Clear direct character formatting for range while preserving annotation targets.
let clearFormatting start len buffer =
    applyStyle start len (fun style ->
        {
            defaultStyle with
                Hyperlink = style.Hyperlink
                CommentId = style.CommentId
                CommentText = style.CommentText
                CommentAuthor = style.CommentAuthor
                CommentDate = style.CommentDate
                NoteId = style.NoteId
                NoteType = style.NoteType
                NoteText = style.NoteText
                BookmarkStartId = style.BookmarkStartId
                BookmarkName = style.BookmarkName
                BookmarkEndId = style.BookmarkEndId
                FieldCode = style.FieldCode
                FieldResult = style.FieldResult
                RenderedPageBreakBefore = style.RenderedPageBreakBefore
                EquationXml = style.EquationXml
                RevisionId = style.RevisionId
                RevisionType = style.RevisionType
                RevisionAuthor = style.RevisionAuthor
                RevisionDate = style.RevisionDate
        }) buffer

/// Get style at position
let getStyleAt pos buffer =
    styleAt pos buffer

/// Convert position to line/column
let posToLineCol pos buffer =
    Rope.positionToLineCol pos buffer.Rope

/// Convert line/column to position
let lineColToPos line col buffer =
    Rope.lineColToPosition line col buffer.Rope
