module RedBat.Word.TextBuffer

open RedBat.Word.Rope

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
        let chars =
            text
            |> Seq.filter (fun ch -> ch <> '\t' && ch <> '\r' && ch <> '\n')
            |> Seq.toArray
        if chars.Length = 0 then 0.0
        elif System.Char.IsWhiteSpace chars.[chars.Length - 1] then float chars.Length
        else float (chars.Length - 1)

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

/// Insert text at position
let insert pos (text: string) buffer =
    if String.length text = 0 then buffer
    else
        let textLen = String.length text

        // Find the run containing the insertion point
        let mutable newRuns = []
        let mutable inserted = false
        for run in buffer.Runs do
            let runStart = Run.start run
            let runEnd = Run.endPos run
            if not inserted && pos >= runStart && pos <= runEnd then
                // Insert within this run
                match run with
                | Image _ ->
                    // Can't split image - insert before it
                    newRuns <- newRuns @ [Text { Start = pos; Length = textLen; Style = defaultStyle }]
                    newRuns <- newRuns @ [Run.shiftStart textLen run]
                | Text tr ->
                    if pos = runStart then
                        // Insert at start of run
                        newRuns <- newRuns @ [Text { Start = tr.Start; Length = textLen; Style = tr.Style }]
                        newRuns <- newRuns @ [Text { tr with Start = tr.Start + textLen }]
                    elif pos = runEnd then
                        // Insert at end of run
                        newRuns <- newRuns @ [run]
                        newRuns <- newRuns @ [Text { Start = runEnd; Length = textLen; Style = tr.Style }]
                    else
                        // Insert in middle - split into three runs
                        let beforeLen = pos - tr.Start
                        let afterLen = tr.Length - beforeLen
                        newRuns <- newRuns @ [Text { tr with Length = beforeLen }]
                        newRuns <- newRuns @ [Text { Start = pos; Length = textLen; Style = tr.Style }]
                        newRuns <- newRuns @ [Text { Start = pos + textLen; Length = afterLen; Style = tr.Style }]
                inserted <- true
            elif inserted then
                // Shift runs after insertion
                newRuns <- newRuns @ [Run.shiftStart textLen run]
            else
                // Before insertion point
                newRuns <- newRuns @ [run]

        // If no runs or insertion at end, add new run
        if not inserted then
            newRuns <- newRuns @ [Text { Start = pos; Length = textLen; Style = defaultStyle }]

        { buffer with
            Rope = Rope.insert pos text buffer.Rope
            Runs = newRuns }

/// Delete range [start, start + len)
let delete start len buffer =
    if len <= 0 then buffer
    else
        let deleteEnd = start + len
        let mutable newRuns = []

        for run in buffer.Runs do
            let runStart = Run.start run
            let runEnd = Run.endPos run
            if runEnd <= start then
                // Run is entirely before deletion
                newRuns <- newRuns @ [run]
            elif runStart >= deleteEnd then
                // Run is entirely after deletion - shift back
                newRuns <- newRuns @ [Run.shiftStart -len run]
            elif runStart >= start && runEnd <= deleteEnd then
                // Run is entirely within deletion - remove it
                ()
            elif runStart < start && runEnd > deleteEnd then
                // Deletion is within run - shrink it
                match run with
                | Image _ -> () // Image deleted
                | Text tr -> newRuns <- newRuns @ [Text { tr with Length = tr.Length - len }]
            elif runStart < start then
                // Run starts before, ends within deletion
                match run with
                | Image _ -> () // Image deleted
                | Text tr ->
                    let newLen = start - tr.Start
                    newRuns <- newRuns @ [Text { tr with Length = newLen }]
            else
                // Run starts within deletion, ends after
                match run with
                | Image _ -> () // Image deleted
                | Text tr ->
                    let overlap = deleteEnd - tr.Start
                    let newLen = tr.Length - overlap
                    newRuns <- newRuns @ [Text { Start = start; Length = newLen; Style = tr.Style }]

        { buffer with
            Rope = Rope.delete start len buffer.Rope
            Runs = newRuns }

/// Get substring
let slice start len buffer =
    Rope.slice start len buffer.Rope

/// Object Replacement Character - used as placeholder for images
let objectReplacementChar = "\uFFFC"

/// Resize image at position
let resizeImage pos newWidth newHeight buffer =
    let newRuns =
        buffer.Runs
        |> List.map (function
            | Image imgRun when imgRun.Start = pos ->
                Image { imgRun with Width = Some newWidth; Height = Some newHeight }
            | run -> run)

    { buffer with Runs = newRuns }

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
    let mutable newRuns = []
    let mutable inserted = false
    for run in buffer.Runs do
        let runStart = Run.start run
        if not inserted && pos <= runStart then
            newRuns <- newRuns @ [imageRun]
            newRuns <- newRuns @ [Run.shiftStart 1 run]
            inserted <- true
        elif inserted then
            newRuns <- newRuns @ [Run.shiftStart 1 run]
        else
            newRuns <- newRuns @ [run]

    if not inserted then
        newRuns <- newRuns @ [imageRun]

    { Rope = newRope; Runs = newRuns }

let insertImage pos src width height alt buffer =
    insertImageWithMetadata pos src false width height alt None None false None None None None None None None None None None None None None None None None None None None None None buffer

/// Apply style to range [start, start + len)
let applyStyle start len styleTransform buffer =
    if len <= 0 then buffer
    else
        let rangeEnd = start + len
        let mutable newRuns = []

        for run in buffer.Runs do
            let runStart = Run.start run
            let runEnd = Run.endPos run
            match run with
            | Image _ ->
                // Images don't have text styles - keep as is
                newRuns <- newRuns @ [run]
            | Text tr ->
                if runEnd <= start || runStart >= rangeEnd then
                    // Run is outside the range
                    newRuns <- newRuns @ [run]
                elif runStart >= start && runEnd <= rangeEnd then
                    // Run is entirely within range - apply style
                    newRuns <- newRuns @ [Text { tr with Style = styleTransform tr.Style }]
                elif runStart < start && runEnd > rangeEnd then
                    // Range is within run - split into three
                    let beforeLen = start - tr.Start
                    let rangeLen = len
                    let afterLen = runEnd - rangeEnd
                    newRuns <- newRuns @ [Text { tr with Length = beforeLen }]
                    newRuns <- newRuns @ [Text { Start = start; Length = rangeLen; Style = styleTransform tr.Style }]
                    newRuns <- newRuns @ [Text { Start = rangeEnd; Length = afterLen; Style = tr.Style }]
                elif runStart < start then
                    // Run overlaps start of range
                    let beforeLen = start - tr.Start
                    let overlapLen = runEnd - start
                    newRuns <- newRuns @ [Text { tr with Length = beforeLen }]
                    newRuns <- newRuns @ [Text { Start = start; Length = overlapLen; Style = styleTransform tr.Style }]
                else
                    // Run overlaps end of range
                    let overlapLen = rangeEnd - tr.Start
                    let afterLen = runEnd - rangeEnd
                    newRuns <- newRuns @ [Text { Start = tr.Start; Length = overlapLen; Style = styleTransform tr.Style }]
                    newRuns <- newRuns @ [Text { Start = rangeEnd; Length = afterLen; Style = tr.Style }]

        // Merge adjacent text runs with same style
        let rec mergeRuns acc = function
            | [] -> List.rev acc
            | [run] -> List.rev (run :: acc)
            | Text tr1 :: Text tr2 :: rest when tr1.Start + tr1.Length = tr2.Start && tr1.Style = tr2.Style ->
                // Merge text runs
                let merged = Text { tr1 with Length = tr1.Length + tr2.Length }
                mergeRuns acc (merged :: rest)
            | run1 :: run2 :: rest ->
                mergeRuns (run1 :: acc) (run2 :: rest)

        { buffer with Runs = mergeRuns [] newRuns }

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
    applyStyle start len (fun style -> { style with VertAlign = align }) buffer

/// Set text color for range
let setColor start len color buffer =
    applyStyle start len (fun style -> { style with Color = color }) buffer

/// Set highlight color for range
let setHighlight start len color buffer =
    applyStyle start len (fun style -> { style with Highlight = color }) buffer

/// Set hyperlink target for range
let setHyperlink start len target buffer =
    applyStyle start len (fun style -> { style with Hyperlink = target }) buffer

/// Set comment metadata for range
let setComment start len commentId commentText author date buffer =
    applyStyle start len (fun style -> { style with CommentId = commentId; CommentText = commentText; CommentAuthor = author; CommentDate = date }) buffer

/// Set footnote/endnote metadata for range
let setNote start len noteId noteType noteText buffer =
    applyStyle start len (fun style -> { style with NoteId = noteId; NoteType = noteType; NoteText = noteText }) buffer

/// Set bookmark markers for range
let setBookmark start len bookmarkId bookmarkName buffer =
    applyStyle start len (fun style -> { style with BookmarkStartId = Some bookmarkId; BookmarkName = Some bookmarkName; BookmarkEndId = Some bookmarkId }) buffer

/// Set named character style for range
let setCharacterStyle start len styleId buffer =
    applyStyle start len (fun style -> { style with Style = styleId }) buffer

/// Set revision metadata for range
let setRevision start len revisionType revisionId author date buffer =
    applyStyle start len (fun style -> { style with RevisionType = Some revisionType; RevisionId = Some revisionId; RevisionAuthor = author; RevisionDate = date }) buffer

/// Clear revision metadata for range
let clearRevision start len buffer =
    applyStyle start len (fun style -> { style with RevisionType = None; RevisionId = None; RevisionAuthor = None; RevisionDate = None }) buffer

/// Set font size for range
let setFontSize start len size buffer =
    applyStyle start len (fun style -> { style with FontSize = size }) buffer

/// Set font family for range
let setFontFamily start len family buffer =
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
    buffer.Runs
    |> List.tryFind (fun run -> pos >= Run.start run && pos < Run.endPos run)
    |> Option.bind (function
        | Text tr -> Some tr.Style
        | Image _ -> None)  // Images don't have text style
    |> Option.defaultValue defaultStyle

/// Convert position to line/column
let posToLineCol pos buffer =
    Rope.positionToLineCol pos buffer.Rope

/// Convert line/column to position
let lineColToPos line col buffer =
    Rope.lineColToPosition line col buffer.Rope
