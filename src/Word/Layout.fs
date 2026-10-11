module Rbt.Word.Layout

open System.Collections.Generic
open Rbt.Word.Document
open Rbt.Word.TextBuffer
open Rbt.Word.WordStyles

#if FABLE_COMPILER
open Fable.Core

[<Emit("try { if (new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1') console.log(new Date().toISOString() + ' word-layout:' + $0); } catch (_) {}")>]
let private trace (_label: string) : unit = jsNative

[<Emit("(function(){ try { return new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1'; } catch (_) { return false; } })()")>]
let private readTraceTimingEnabled () : bool = jsNative

// Read once at load. As an emitted expression it was re-evaluated (parsing
// the URL and reading localStorage) at every use, several times per paragraph
// on every layout pass.
let private traceTimingEnabled : bool = readTraceTimingEnabled ()

[<Emit("performance.now()")>]
let private nowMs () : float = jsNative
#else
let private trace (_label: string) : unit = ()
let private traceTimingEnabled = false

let private nowMs () =
    float (System.Diagnostics.Stopwatch.GetTimestamp()) * 1000.0 / float System.Diagnostics.Stopwatch.Frequency
#endif

let private timingStart () = if traceTimingEnabled then nowMs () else 0.0

let private traceSlow label start =
    if traceTimingEnabled then
        let elapsed = nowMs () - start
        if elapsed >= 25.0 then trace (sprintf "%s %.1fms" label elapsed)

let private traceTableLayout label =
    if traceTimingEnabled then trace ("table:" + label)

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private hasNonWhitespace (text: string) =
    let mutable index = 0
    let mutable found = false
    while not found && index < text.Length do
        if not (System.Char.IsWhiteSpace text.[index]) then
            found <- true
        index <- index + 1
    found

let private paragraphContentMetrics paragraphStyle metadata (buffer: TextBuffer.TextBuffer) =
    let mutable foundText = false
    let mutable contentFontSize = 0.0
    let mutable contentTextStyle = paragraphStyle.Text
    for run in buffer.Runs do
        match run with
        | Text textRun ->
            let displaySize = WordLayout.Text.displayFontSize textRun.Style
            if not foundText || displaySize > contentFontSize then
                foundText <- true
                contentFontSize <- displaySize
                contentTextStyle <- textRun.Style
        | Image _ -> ()
    if foundText then
        contentFontSize, contentTextStyle
    else
        metadata.ParagraphMarkFontSize |> Option.defaultValue paragraphStyle.Text.FontSize,
        paragraphStyle.Text

let private snippet (text: string) =
    let mutable start = 0
    let mutable finish = text.Length - 1
    while start <= finish && System.Char.IsWhiteSpace text.[start] do
        start <- start + 1
    while finish >= start && System.Char.IsWhiteSpace text.[finish] do
        finish <- finish - 1
    if start > finish then ""
    else
        let sb = System.Text.StringBuilder(min 48 (finish - start + 1))
        let mutable index = start
        while index <= finish && sb.Length < 48 do
            match text.[index] with
            | '\r' | '\n' | '\t' -> sb.Append(' ') |> ignore
            | ch -> sb.Append(ch) |> ignore
            index <- index + 1
        sb.ToString()

let private isAllWhitespace (text: string) =
    let mutable index = 0
    let mutable allWhitespace = true
    while allWhitespace && index < text.Length do
        if not (System.Char.IsWhiteSpace text.[index]) then
            allWhitespace <- false
        index <- index + 1
    allWhitespace

let private trimmedEqualsIgnoreCase (value: string) (expected: string) =
    if isNull value then false
    else
        let mutable start = 0
        let mutable finish = value.Length - 1
        while start <= finish && System.Char.IsWhiteSpace value.[start] do
            start <- start + 1
        while finish >= start && System.Char.IsWhiteSpace value.[finish] do
            finish <- finish - 1
        let length = finish - start + 1
        if length <> expected.Length then false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < length do
                let left = value.[start + index]
                let right = expected.[index]
                if left <> right && System.Char.ToUpperInvariant left <> System.Char.ToUpperInvariant right then
                    equal <- false
                index <- index + 1
            equal

let private equalsIgnoreCase (left: string) (right: string) =
    System.String.Equals(left, right, System.StringComparison.OrdinalIgnoreCase)

let private optionEqualsIgnoreCase expected value =
    match value with
    | Some actual -> equalsIgnoreCase actual expected
    | None -> false

let private hasDigitsFrom (start: int) (value: string) =
    let mutable index = start
    let mutable hasDigit = start < value.Length
    let mutable valid = hasDigit
    while valid && index < value.Length do
        if not (System.Char.IsDigit value.[index]) then
            valid <- false
        index <- index + 1
    hasDigit && valid

let private asciiSingleCharStrings =
    Array.init 256 (fun code -> string (char code))

let private sliceText (text: string) start length =
    if length = 1 then
        let ch = text.[start]
        let code = int ch
        if code >= 0 && code < asciiSingleCharStrings.Length then
            asciiSingleCharStrings.[code]
        else
            text.Substring(start, length)
    else
        text.Substring(start, length)

module Eq = Rbt.Word.Math

/// Segment can be text, image, or a laid-out math equation (OMML).
type Segment =
    | TextSegment of text: string * style: TextBuffer.TextStyle * x: float * width: float * startOffset: int
    | ImageSegment of src: string * x: float * yOffset: float * width: float * height: float * startOffset: int * floating: bool * wrap: string option * relativeHeight: string option * cropLeft: string option * cropTop: string option * cropRight: string option * cropBottom: string option
    /// A rendered equation: pre-laid-out primitives (relative to x and the line
    /// baseline), the run style (colour/font), box metrics, and the run's offset span.
    | MathSegment of prims: Eq.MathPrim list * style: TextBuffer.TextStyle * x: float * width: float * ascent: float * descent: float * startOffset: int * length: int

/// A picture segment's `relativeHeight` says where the picture stands among the
/// anchored drawings of its page (wp:anchor relativeHeight: the higher, the
/// further in front) - and, by this mark before the number, that it stands
/// behind the text (behindDoc). The painter reads both (see Paint.Drawing).
let behindTextMark = "behind:"

/// What a picture's segment carries as its relativeHeight (see behindTextMark).
let imageSegmentHeight (image: TextBuffer.ImageRun) : string option =
    if image.Floating && image.BehindDoc = Some true then Some (behindTextMark + (image.RelativeHeight |> Option.defaultValue "")) else image.RelativeHeight

/// Whether a segment's relativeHeight says "behind the text", and the height the file gave.
let splitSegmentHeight (value: string option) : bool * string option =
    match value with
    | Some text when text.StartsWith behindTextMark ->
        let rest = text.Substring behindTextMark.Length
        true, (if rest = "" then None else Some rest)
    | other -> false, other

/// Measured line for layout.
type MeasuredLine = {
    Segments: Segment list
    Y: float
    Width: float
    Height: float
    StartOffset: int
    EndOffset: int
}

let private lineStartsAfterWrapObstacle (line: MeasuredLine) =
    let mutable foundSegment = false
    let mutable minX = 0.0
    for segment in line.Segments do
        let x =
            match segment with
            | TextSegment (_, _, x, _, _)
            | ImageSegment (_, x, _, _, _, _, _, _, _, _, _, _, _)
            | MathSegment (_, _, x, _, _, _, _, _) -> x
        if not foundSegment || x < minX then
            foundSegment <- true
            minX <- x
    foundSegment && minX > 0.01

/// Layout result for a block.
type BlockLayout = {
    BlockId: BlockId
    Lines: MeasuredLine list
    Y: float
    Height: float
    AvailableWidth: float
    Metadata: BlockMetadata
    Prefix: string
    PrefixStyle: TextBuffer.TextStyle option
    PrefixWidth: float
    Indent: float
    // Vertical-alignment offset applied to this cell's lines (center/bottom
    // vAlign). Word's VBA position API reports the *unaligned* content origin,
    // so layout queries subtract this to match Microsoft Word.
    VAlignOffset: float
}

/// Table cell layout.
type CellLayout = {
    Paragraphs: BlockLayout list
    X: float
    Y: float
    Width: float
    Height: float
}

/// Table layout.
type TableLayout = {
    BlockId: BlockId
    Rows: CellLayout list list
    Y: float
    Height: float
}

type TableCellLayoutGroup = {
    VerticalAlignment: string option
    PaddingTop: float
    PaddingBottom: float
    ContentHeight: float
    /// Height the cell asks of its row(s): content, padding, borders, spacing.
    /// A vertically merged cell asks it of all the rows it spans together.
    RequiredHeight: float
    Layouts: BlockLayout list
}

type TableRowMeasurement = {
    Height: float
    CellGroups: TableCellLayoutGroup list
    Traces: string list
}

let private sortedTableRowBlockLines (measurement: TableRowMeasurement) =
    let lines = ResizeArray<BlockId * MeasuredLine>()
    for group in measurement.CellGroups do
        for layout in group.Layouts do
            for line in layout.Lines do
                lines.Add(layout.BlockId, line)
    let result = lines.ToArray()
    Array.sortInPlaceWith
        (fun (_, (firstLine: MeasuredLine)) (_, (secondLine: MeasuredLine)) ->
            compare
                (firstLine.Y, firstLine.StartOffset, firstLine.EndOffset)
                (secondLine.Y, secondLine.StartOffset, secondLine.EndOffset))
        result
    result

let private tableRowMeasurementLineSummary (measurement: TableRowMeasurement) =
    let blockIds = ResizeArray<BlockId>()
    let seenBlockIds = System.Collections.Generic.HashSet<BlockId>()
    let mutable lineCount = 0
    let mutable contentTopValue = System.Double.PositiveInfinity
    let mutable contentBottomValue = System.Double.NegativeInfinity
    for group in measurement.CellGroups do
        for layout in group.Layouts do
            for line in layout.Lines do
                lineCount <- lineCount + 1
                if line.Y < contentTopValue then contentTopValue <- line.Y
                let lineBottom = line.Y + line.Height
                if lineBottom > contentBottomValue then contentBottomValue <- lineBottom
                if seenBlockIds.Add layout.BlockId then
                    blockIds.Add layout.BlockId
    let contentTop = if lineCount = 0 then None else Some contentTopValue
    let contentBottom = if lineCount = 0 then None else Some contentBottomValue
    lineCount, contentTop, contentBottom, resizeArrayToList blockIds

type TableRowFragmentPlan = {
    Page: int
    RowPlanY: float
    RowContentY: float
    Y: float
    Height: float
    AvailableHeight: float
    RepeatedHeaderHeight: float
    ContentTop: float option
    ContentBottom: float option
    LineCount: int
    ConsumedLineStart: int
    ConsumedLineEnd: int
    BlockIds: BlockId list
    BreakReason: string
}

type TableLayoutPlanRow = {
    RowIndex: int
    Row: TableRow
    IsHeader: bool
    CanSplit: bool
    Y: float
    Height: float
    Measurement: TableRowMeasurement
    Fragments: TableRowFragmentPlan list
    /// The row starts a page the table goes on to, below the header rows
    /// repeated there: their height is left free above it.
    HeadersAbove: bool
}

type TableLayoutPlan = {
    Table: TableBlock
    TableX: float
    TableMaxWidth: float
    TableWidth: float
    OuterLeft: float
    OuterRight: float
    ColumnWidths: float list
    ResolvedRows: Lazy<WordLayout.Table.ResolvedRow array>
    HeaderRowCount: int
    Rows: TableLayoutPlanRow list
    NextY: float
}

type TableCellFragmentPlan = {
    RowIndex: int
    CellIndex: int
    StartColumn: int
    ColumnSpan: int
    RowSpan: int
    BorderTop: WordLayout.Table.PrintStroke option
    BorderRight: WordLayout.Table.PrintStroke option
    BorderBottom: WordLayout.Table.PrintStroke option
    BorderLeft: WordLayout.Table.PrintStroke option
    Page: int
    RowY: float
    RowHeight: float
    FragmentY: float
    FragmentHeight: float
    CellX: float
    CellWidth: float
    CellY: float
    CellHeight: float
    CellSpacing: float
    CellBackgroundColor: string option
    CellShading: Shading option
    PaddingTop: float
    PaddingBottom: float
    ContentY: float
    ContentHeight: float
    VisibleY: float option
    VisibleHeight: float option
    VisibleLineCount: int
    BlockIds: BlockId list
}

let private rowSpanHeightAt (rowHeights: float array) rowIndex rowSpan fallback =
    let span = max 1 rowSpan
    let last = min rowHeights.Length (rowIndex + span)
    let mutable total = 0.0
    let mutable index = max 0 rowIndex
    while index < last do
        total <- total + rowHeights.[index]
        index <- index + 1
    max total fallback

let private tablePlanRowHeights (plan: TableLayoutPlan) =
    let rowHeights = ResizeArray<float>()
    for row in plan.Rows do
        rowHeights.Add row.Height
    rowHeights.ToArray()

let private rowHeightPrefix (rowHeights: float array) =
    let prefix = Array.zeroCreate<float> (rowHeights.Length + 1)
    for index = 0 to rowHeights.Length - 1 do
        prefix.[index + 1] <- prefix.[index] + rowHeights.[index]
    prefix

let private rowSpanHeightFromPrefix (prefix: float array) rowIndex rowSpan fallback =
    if prefix.Length <= 1 then
        fallback
    else
        let start = max 0 (min rowIndex (prefix.Length - 1))
        let finish = min (prefix.Length - 1) (start + max 1 rowSpan)
        max (prefix.[finish] - prefix.[start]) fallback

let private resolvedCellRowSpanLookup (resolvedRows: WordLayout.Table.ResolvedRow array) =
    let lookup = Array.init resolvedRows.Length (fun _ -> [||])
    for rowIndex = 0 to resolvedRows.Length - 1 do
        let mutable maxSourceCellIndex = -1
        for cell in resolvedRows.[rowIndex].Cells do
            if cell.SourceCellIndex > maxSourceCellIndex then
                maxSourceCellIndex <- cell.SourceCellIndex
        if maxSourceCellIndex >= 0 then
            let rowLookup = Array.zeroCreate<int> (maxSourceCellIndex + 1)
            for cell in resolvedRows.[rowIndex].Cells do
                rowLookup.[cell.SourceCellIndex] <- cell.RowSpan
            lookup.[rowIndex] <- rowLookup
    lookup

let private resolvedCellRowSpanFromLookup (lookup: int array array) rowIndex sourceCellIndex =
    if rowIndex < 0 || rowIndex >= lookup.Length then
        1
    else
        let rowLookup = lookup.[rowIndex]
        if sourceCellIndex < 0 || sourceCellIndex >= rowLookup.Length then
            1
        else
            let rowSpan = rowLookup.[sourceCellIndex]
            if rowSpan > 0 then rowSpan else 1

let private tableCellVerticalAlignmentHeightFromLookup rowSpanLookup rowHeightPrefix rowIndex sourceCellIndex fallback =
    let rowSpan = resolvedCellRowSpanFromLookup rowSpanLookup rowIndex sourceCellIndex
    rowSpanHeightFromPrefix rowHeightPrefix rowIndex rowSpan fallback

/// The height a cell is painted with (fill, borders) in one fragment of its
/// row. A row that fits on its page has one fragment, and the cell is as
/// high as the rows it spans. A row split across pages has a fragment per
/// page: there the cell is as high as the fragment - the row's height runs
/// from its top on one page to its bottom on another and includes the page
/// break, so painting it would run the cell's sides through the bottom
/// margin, the gap between the pages and on below the row. The rows a cell
/// spans below a split row follow its last fragment.
let private tableCellPaintHeightInFragment (rowSpanLookup: int array array) (rowHeightPrefix: float array) rowIndex sourceCellIndex (fragmentCount: int) (fragmentIndex: int) (fragmentHeight: float) =
    if fragmentCount <= 1 then
        tableCellVerticalAlignmentHeightFromLookup rowSpanLookup rowHeightPrefix rowIndex sourceCellIndex fragmentHeight
    else
        let rowSpan = resolvedCellRowSpanFromLookup rowSpanLookup rowIndex sourceCellIndex
        if rowSpan > 1 && fragmentIndex = fragmentCount - 1 && rowHeightPrefix.Length > 1 then
            let start = max 0 (min (rowIndex + 1) (rowHeightPrefix.Length - 1))
            let finish = min (rowHeightPrefix.Length - 1) (rowIndex + rowSpan)
            fragmentHeight + max 0.0 (rowHeightPrefix.[finish] - rowHeightPrefix.[start])
        else fragmentHeight

let tableCellVerticalAlignmentHeightResolverForPlan (plan: TableLayoutPlan) =
    let rowHeights = tablePlanRowHeights plan
    let rowHeightPrefix = rowHeightPrefix rowHeights
    let rowSpanLookup = resolvedCellRowSpanLookup plan.ResolvedRows.Value
    fun rowIndex sourceCellIndex fallback ->
        tableCellVerticalAlignmentHeightFromLookup rowSpanLookup rowHeightPrefix rowIndex sourceCellIndex fallback

let tableCellVerticalAlignmentHeightForPlan (plan: TableLayoutPlan) rowIndex sourceCellIndex fallback =
    let resolveHeight = tableCellVerticalAlignmentHeightResolverForPlan plan
    resolveHeight rowIndex sourceCellIndex fallback

type LayoutSettings = {
    Width: float
    PageHeight: float
    FontSize: float
    LineHeight: float
    PaddingLeft: float
    PaddingTop: float
    MarginRight: float
    MarginBottom: float
    InTableCell: bool
    TableContinuationHeaderHeight: float
    // Long trailing-space tokens need backend-specific slack: PDF/native
    // shapers track Word more closely than browser canvas for some fonts.
    LongTrailingSpaceFitSafety: float
}

let pageHeight settings = settings.PageHeight
let pageGap = WordLayout.Page.gap

let private pageNumberForY (settings: LayoutSettings) y =
    int (floor (max 0.0 y / (pageHeight settings + pageGap))) + 1

let private minBodyTop settings pageTop = WordLayout.Page.minBodyTop settings.PaddingTop settings.FontSize pageTop

let private advanceToNextPage settings y =
    WordLayout.Page.advanceToNextPage (pageHeight settings) settings.PaddingTop y

let private pageFitTolerance = 0.5

let private applyRenderedPageBreakHint settings y (para: ParagraphBlock) =
    y

let private shiftLinesY delta lines =
    let result = ResizeArray<MeasuredLine>()
    for line in lines do
        result.Add { line with Y = line.Y + delta }
    resizeArrayToList result

let private moveLayoutToY newY (layout: BlockLayout) =
    let delta = newY - layout.Y
    { layout with
        Y = layout.Y + delta
        Lines = shiftLinesY delta layout.Lines }

let private paginateBlock settings y blockHeight =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    if y > contentTop && y + blockHeight > contentBottom + pageFitTolerance then advanceToNextPage settings y else y

let private availableContentHeight settings y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    max 0.0 (contentBottom - y)

let private pageIndexForY settings y =
    int (floor (max 0.0 y / (pageHeight settings + pageGap)))

let noteWrapChars (settings: LayoutSettings) =
    // ~6.3 px per char approximates Calibri 10pt so the reserve's line estimate
    // is not lower than the actual wrapped footnote height (which would let body
    // content overlap the notes).
    max 24 (int ((settings.Width - settings.PaddingLeft - settings.MarginRight) / 6.3))

let wrapNoteText settings (text: string) =
    let maxChars = noteWrapChars settings
    let lines = System.Collections.Generic.List<string>()
    let mutable current = System.Text.StringBuilder()
    let mutable currentLen = 0
    let mutable wordStart = -1
    let mutable sawWord = false

    let appendWord start length =
        for i = start to start + length - 1 do
            current.Append(text.[i]) |> ignore

    let flushWord wordStart wordLen =
        sawWord <- true
        if currentLen = 0 then
            appendWord wordStart wordLen
            currentLen <- wordLen
        elif currentLen + 1 + wordLen <= maxChars then
            current.Append(' ') |> ignore
            appendWord wordStart wordLen
            currentLen <- currentLen + 1 + wordLen
        else
            lines.Add(current.ToString())
            current <- System.Text.StringBuilder()
            appendWord wordStart wordLen
            currentLen <- wordLen

    for i = 0 to text.Length - 1 do
        match text.[i] with
        | ' ' | '\t' | '\r' | '\n' ->
            if wordStart >= 0 then
                flushWord wordStart (i - wordStart)
                wordStart <- -1
        | _ ->
            if wordStart < 0 then wordStart <- i
    if wordStart >= 0 then
        flushWord wordStart (text.Length - wordStart)

    if not sawWord then [text]
    else
        if currentLen > 0 then
            lines.Add(current.ToString())
        resizeArrayToList lines

let noteLineCount settings (text: string) =
    let maxChars = noteWrapChars settings
    let mutable lineCount = 0
    let mutable currentLen = 0
    let mutable wordStart = -1
    let flushWord wordLen =
        if currentLen = 0 then
            currentLen <- wordLen
        elif currentLen + 1 + wordLen <= maxChars then
            currentLen <- currentLen + 1 + wordLen
        else
            lineCount <- lineCount + 1
            currentLen <- wordLen
    for i = 0 to text.Length - 1 do
        match text.[i] with
        | ' ' | '\t' | '\r' | '\n' ->
            if wordStart >= 0 then
                flushWord (i - wordStart)
                wordStart <- -1
        | _ ->
            if wordStart < 0 then wordStart <- i
    if wordStart >= 0 then
        flushWord (text.Length - wordStart)
    if currentLen = 0 then 1 else lineCount + 1

let private lineHasNoteType noteTypeName (line: MeasuredLine) =
    let mutable found = false
    for segment in line.Segments do
        if not found then
            match segment with
            | TextSegment (_, style, _, _, _) ->
                match style.NoteId, style.NoteType, style.NoteText with
                | Some _, noteType, Some text when (noteType |> Option.defaultValue "footnote") = noteTypeName && text <> "" ->
                    found <- true
                | _ -> ()
            | ImageSegment _ | MathSegment _ -> ()
    found

let private layoutHasNoteType noteTypeName (layout: BlockLayout) =
    let mutable found = false
    for line in layout.Lines do
        if not found && lineHasNoteType noteTypeName line then
            found <- true
    found

let private notesInLine noteTypeName (line: MeasuredLine) =
    let mutable seen = Unchecked.defaultof<System.Collections.Generic.HashSet<string>>
    let mutable notes = Unchecked.defaultof<ResizeArray<string * string>>
    let mutable hasNotes = false
    for segment in line.Segments do
        match segment with
        | TextSegment (_, style, _, _, _) ->
            match style.NoteId, style.NoteType, style.NoteText with
            | Some id, noteType, Some text when (noteType |> Option.defaultValue "footnote") = noteTypeName && text <> "" ->
                if not hasNotes then
                    seen <- System.Collections.Generic.HashSet<string>()
                    notes <- ResizeArray<string * string>()
                    hasNotes <- true
                if seen.Add id then
                    notes.Add(id, text)
            | _ -> ()
        | ImageSegment _ | MathSegment _ -> ()
    if hasNotes then resizeArrayToList notes else []

let private footnotesInLine (line: MeasuredLine) =
    notesInLine "footnote" line

let private appendNewNotes (existing: (string * string) list) (additions: (string * string) list) =
    match additions with
    | [] -> existing
    | _ ->
        let seen = System.Collections.Generic.HashSet<string>()
        for id, _ in existing do
            seen.Add id |> ignore
        let merged = ResizeArray<string * string>(existing.Length + additions.Length)
        merged.AddRange existing
        let mutable changed = false
        for id, text in additions do
            if seen.Add id then
                merged.Add(id, text)
                changed <- true
        if changed then resizeArrayToList merged else existing

/// Endnotes (collected at the end of the document) found on a line, in order.
let private endnotesInLine (line: MeasuredLine) =
    notesInLine "endnote" line

/// BlockId prefix for the synthetic footnote-area paragraphs laid out at the
/// bottom of each page. Paint recognises these to render them and the separator.
let footnoteBlockIdPrefix = "__footnote__"
/// BlockId prefix for the synthetic endnote paragraphs laid out at the end of
/// the document (after all body content), with a separator above the first.
let endnoteBlockIdPrefix = "__endnote__"

/// Word numbers endnotes with lowercase roman numerals by default (footnotes use
/// arabic). Convert the (sequential) note id to roman for the endnote mark.
let private toLowerRoman (n: int) =
    let table = [ 1000, "m"; 900, "cm"; 500, "d"; 400, "cd"; 100, "c"; 90, "xc"; 50, "l"; 40, "xl"; 10, "x"; 9, "ix"; 5, "v"; 4, "iv"; 1, "i" ]
    let rec build rem acc =
        match table |> List.tryFind (fun (v, _) -> rem >= v) with
        | Some (v, s) -> build (rem - v) (acc + s)
        | None -> acc
    if n <= 0 then string n else build n ""

let private endnoteMark (id: string) =
    match System.Int32.TryParse id with
    | true, n -> toLowerRoman n
    | _ -> id
// Font sizes in the layout/buffer are CSS px (1pt = 96/72 px), as produced by
// the editor import (11pt body = 14.667px). Word's FootnoteText is 10pt.
let private footnoteFontSize = 10.0 * 96.0 / 72.0
// Footnotes inherit the document body font; Calibri is Word's default and the
// repro/most real docs use it. (A doc-default-font lookup would be more robust.)
let private footnoteFontFamily = "Calibri"
// Small gap left below the footnote block, above the bottom page margin.
let private footnoteBottomPad = 6.0
// Footnote single-line height (10pt body font) and the separator gap above the
// first note. The reserve must be >= the actual footnote-block height so body
// content does not overlap the notes (which are positioned by their real height).
let private footnoteLineHeight = 16.0
let private footnoteSeparatorGap = 16.0
// Gap below the last body line before the endnote separator + first endnote.
let private endnoteTopGap = 24.0

let private footnoteReserveHeight settings footnotes =
    match footnotes with
    | [] -> 0.0
    | notes ->
        let lineCount = notes |> List.sumBy (fun (id, text) -> noteLineCount settings (sprintf "%s. %s" id text))
        footnoteSeparatorGap + float lineCount * footnoteLineHeight + footnoteBottomPad

let private addLayoutFootnotes settings (footnotesByPage: Map<int, (string * string) list>) (layout: BlockLayout) =
    layout.Lines
    |> List.fold (fun acc line ->
        let pageIndex = pageIndexForY settings line.Y
        let lineFootnotes = footnotesInLine line
        if lineFootnotes.IsEmpty then
            acc
        else
            let existing = acc |> Map.tryFind pageIndex |> Option.defaultValue []
            let footnotes = appendNewNotes existing lineFootnotes
            if LanguagePrimitives.PhysicalEquality footnotes existing then acc else acc |> Map.add pageIndex footnotes) footnotesByPage

let private contentBottomWithFootnotes settings footnotesByPage y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let pageIndex = pageIndexForY settings y
    WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    - footnoteReserveHeight settings (footnotesByPage |> Map.tryFind pageIndex |> Option.defaultValue [])

let private paginateBlockWithFootnotes settings footnotesByPage y blockHeight =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
    let contentBottom = contentBottomWithFootnotes settings footnotesByPage y
    if y > contentTop && y + blockHeight > contentBottom + pageFitTolerance then advanceToNextPage settings y else y

let private lineVisibleFitHeight line =
    let mutable visibleHeight = 0.0
    let mutable hasVisibleSegment = false
    for segment in line.Segments do
        let candidate =
            match segment with
            | TextSegment (_, style, _, _, _) -> Some (WordLayout.Text.printFontSize style)
            | ImageSegment (_, _, yOffset, _, height, _, floating, _, _, _, _, _, _) ->
                if floating then None else Some (height + max 0.0 yOffset)
            | MathSegment (_, _, _, _, ascent, descent, _, _) -> Some (ascent + descent)
        match candidate with
        | Some height ->
            hasVisibleSegment <- true
            if height > visibleHeight then visibleHeight <- height
        | _ -> ()
    if hasVisibleSegment then visibleHeight else line.Height

let private tryLastLine lines =
    match lines with
    | [] -> None
    | first :: rest ->
        let mutable last = first
        for line in rest do
            last <- line
        Some last

let private paragraphVisibleFitHeight (paragraphStyle: ParagraphStyle) (layout: BlockLayout) =
    if layout.Metadata.ListStyle = Numbered then
        layout.Height
    else
        let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle layout.Metadata
        match layout.Lines with
        | [] -> max 0.0 (layout.Height - spacingAfter)
        // Word allows a single-line heading to sit at the page bottom even when
        // its after-spacing falls beyond the printable body.
        | [line] when paragraphStyle.OutlineLevel.IsSome || layout.Metadata.Style = Some "Heading1" ->
            max 0.0 ((line.Y - layout.Y) + lineVisibleFitHeight line)
        // Non-heading one-line paragraphs still need their full paragraph box,
        // including after-spacing, to fit on the page.
        | [_] -> layout.Height
        | _ ->
            match tryLastLine layout.Lines with
            | Some lastLine -> max 0.0 ((lastLine.Y - layout.Y) + lineVisibleFitHeight lastLine)
            | None -> max 0.0 (layout.Height - spacingAfter)

let private paginateTableRow settings (row: TableRow) y rowHeight =
    let canSplit = row.CanSplit |> Option.defaultValue true
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    let remaining = max 0.0 (contentBottom - y)
    if canSplit && y < contentBottom && rowHeight > remaining then y else paginateBlock settings y rowHeight

let private paginateTableRowWithFootnotes settings footnotesByPage (row: TableRow) y rowHeight =
    let canSplit = row.CanSplit |> Option.defaultValue true
    let contentBottom = contentBottomWithFootnotes settings footnotesByPage y
    let remaining = max 0.0 (contentBottom - y)
    if canSplit && y < contentBottom && rowHeight > remaining then y else paginateBlockWithFootnotes settings footnotesByPage y rowHeight

let private splitTableRowNeedsNextPage settings rowY contentBottom (measurement: TableRowMeasurement) =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) rowY
    let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
    if rowY <= contentTop + pageFitTolerance then
        false
    else
        let mutable hasLines = false
        let mutable anyLineFits = false
        for group in measurement.CellGroups do
            for layout in group.Layouts do
                for line in layout.Lines do
                    hasLines <- true
                    if rowY + line.Y + lineVisibleFitHeight line <= contentBottom + pageFitTolerance then
                        anyLineFits <- true
        hasLines && not anyLineFits

let private moveWidowParagraphToNextPage settings (lines: MeasuredLine list) =
    let shiftToNextPage (firstY: float) (lines: MeasuredLine list) : MeasuredLine list =
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) firstY
        let nextY = WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + settings.TableContinuationHeaderHeight
        let delta = nextY - firstY
        shiftLinesY delta lines

    match lines with
    | first :: second :: _ ->
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) first.Y
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
        let lineFits (line: MeasuredLine) = line.Y + lineVisibleFitHeight line <= contentBottom + pageFitTolerance
        if first.Y > contentTop && lineFits first && not (lineFits second) then
            shiftToNextPage first.Y lines
        else
            let mutable fitCount = 0
            let mutable stillFitting = true
            for line in lines do
                if stillFitting && lineFits line then
                    fitCount <- fitCount + 1
                else
                    stillFitting <- false
            if first.Y > contentTop && lines.Length <= 4 && fitCount = lines.Length - 1 then
                shiftToNextPage first.Y lines
            else
                lines
    | _ -> lines

/// Page height used when content is measured for its natural height rather
/// than placed on a page (table rows; see measureTableRow...RowLocal).
let private unpagedMeasurementPageHeight = 1000000.0

let private lineHasRenderedPageBreakBefore (line: MeasuredLine) =
    let rec loop segments =
        match segments with
        | [] -> false
        | TextSegment (_, style, _, _, _) :: _ when style.RenderedPageBreakBefore -> true
        | _ :: rest -> loop rest

    loop line.Segments

let private paginateLines settings widowControl (lines: MeasuredLine list) =
    let mutable delta = 0.0
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let lines = if widowControl then moveWidowParagraphToNextPage settings lines else lines
    let unpagedMeasurement = settings.PageHeight >= unpagedMeasurementPageHeight
    let pagedLines = ResizeArray<MeasuredLine>()
    for line in lines do
        let shiftedY = line.Y + delta
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) shiftedY
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        let pageIndex = pageIndexForY settings shiftedY
        let lineFootnotes = footnotesInLine line
        let pageExistingFootnotes = footnotesByPage |> Map.tryFind pageIndex |> Option.defaultValue []
        let pageFootnotes = appendNewNotes pageExistingFootnotes lineFootnotes
        let contentBottom =
            WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            - footnoteReserveHeight settings pageFootnotes
        // A lastRenderedPageBreak hint only means something on a real page.
        // Table rows are first measured on a 1,000,000px "page" to get their
        // natural height; honouring the hint there would move the line a
        // million pixels down and report that as the row height.
        let renderedBreakBefore = not unpagedMeasurement && lineHasRenderedPageBreakBefore line
        if shiftedY > contentTop && (renderedBreakBefore || shiftedY + lineVisibleFitHeight line > contentBottom + pageFitTolerance) then
            let nextY = WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + settings.TableContinuationHeaderHeight
            delta <- delta + (nextY - shiftedY)
            let nextPageIndex = pageIndexForY settings nextY
            let nextPageExistingFootnotes = footnotesByPage |> Map.tryFind nextPageIndex |> Option.defaultValue []
            let nextPageFootnotes = appendNewNotes nextPageExistingFootnotes lineFootnotes
            if not (LanguagePrimitives.PhysicalEquality nextPageFootnotes nextPageExistingFootnotes) then
                footnotesByPage <- footnotesByPage |> Map.add nextPageIndex nextPageFootnotes
            pagedLines.Add { line with Y = nextY }
        else
            if not (LanguagePrimitives.PhysicalEquality pageFootnotes pageExistingFootnotes) then
                footnotesByPage <- footnotesByPage |> Map.add pageIndex pageFootnotes
            pagedLines.Add { line with Y = shiftedY }
    resizeArrayToList pagedLines

let private linesSpanPages settings (lines: MeasuredLine list) =
    match lines with
    | [] -> false
    | first :: _ ->
        let firstPage = WordLayout.Page.pageTopForY (pageHeight settings) first.Y
        let mutable spansPages = false
        for line in lines do
            if not spansPages && WordLayout.Page.pageTopForY (pageHeight settings) line.Y <> firstPage then
                spansPages <- true
        spansPages

type TextMeasurer = string -> TextBuffer.TextStyle -> float
type TextRangeMeasurer = string -> TextBuffer.TextStyle -> int -> int -> float

let private createRangeMeasurerFromTextMeasurer (measureText: TextMeasurer) : TextRangeMeasurer =
    fun text style start length ->
        let startIndex = max 0 (min text.Length start)
        let endIndex = max startIndex (min text.Length (start + length))
        if endIndex = startIndex then 0.0
        else measureText (text.Substring(startIndex, endIndex - startIndex)) style

let rangeMeasurerFromTextMeasurer (measureText: TextMeasurer) : TextRangeMeasurer =
    createRangeMeasurerFromTextMeasurer measureText

let tableColumnWidths (maxWidth: float) (tableBlock: TableBlock) =
    WordLayout.Table.tableColumnWidths WordLayout.canvasScale (Some maxWidth) tableBlock

let private columnWidthPrefix (columnWidths: float list) =
    let widths = columnWidths |> List.toArray
    let prefix = Array.zeroCreate<float> (widths.Length + 1)
    for i = 0 to widths.Length - 1 do
        prefix.[i + 1] <- prefix.[i] + widths.[i]
    prefix

type internal TableColumnMetrics = {
    Prefix: float[]
    Count: int
    FallbackWidth: float
}

let internal tableColumnMetrics tableMaxWidth columnWidths =
    let prefix = columnWidthPrefix columnWidths
    let count = prefix.Length - 1
    {
        Prefix = prefix
        Count = count
        FallbackWidth = if count = 0 then tableMaxWidth else tableMaxWidth / float count
    }

let private startOffsetFromPrefix (prefix: float[]) columnIndex =
    if columnIndex <= 0 then 0.0
    else prefix.[min columnIndex (prefix.Length - 1)]

let private spannedWidthFromPrefix (prefix: float[]) start span fallback =
    let widthCount = prefix.Length - 1
    let mutable total = 0.0
    let mutable remaining = span
    let mutable index = start
    while remaining > 0 && index < 0 do
        total <- total + fallback
        index <- index + 1
        remaining <- remaining - 1
    if remaining > 0 && index < widthCount then
        let finish = min widthCount (index + remaining)
        total <- total + prefix.[finish] - prefix.[index]
        remaining <- remaining - (finish - index)
        index <- finish
    if remaining > 0 then
        total <- total + fallback * float remaining
    total

let private spacedStartOffsetFromPrefix prefix spacing columnIndex =
    startOffsetFromPrefix prefix columnIndex + spacing

let private spacedSpannedWidthFromPrefix prefix spacing start span fallback =
    max 0.0 (spannedWidthFromPrefix prefix start span fallback - spacing * 2.0)

// Word line boxes include font leading beyond the nominal font size. Auto
// w:line multiples (eg 259/276) still need this Calibri-leading factor.
let private wordLineBoxFactor = 1.224
let private wordSingleLineBoxFactor = 1.15
let private wordDefaultSpacingSingleLineBoxFactor = 1.316
let private wordTableFollowSingleLineBoxFactor = wordDefaultSpacingSingleLineBoxFactor
let private wordTocEntrySingleLineBoxFactor = 1.266
let private wordAutoSpacingPoints = 14.0
let private monospaceLineBoxFactor = 1.18
let private ubuntuLineBoxFactor = 1.148
let private tableCellDefaultLineBoxFactor = 1.336
// Matches Word's visual single-line pitch for explicit 11pt TableGrid cell paragraphs.
let private tableCellExplicitLineBoxFactor = 1.266
let private arialTableCellExplicitLineBoxFactor = 1.193

let private tableCellLineBoxFontSize fontSize =
    // All import paths now hand us fontSize in CSS-pixels at 96 DPI
    // (DocumentProcessors.fontSizeToPoints multiplies the pt value by 96/72).
    // The earlier 8-12pt special case was patching the native/PDF import
    // path's stale point values; it's no longer needed.
    fontSize

let private lineBoxFactorForTextStyle (style: TextBuffer.TextStyle) =
    match style.FontFamily with
    | Some family when trimmedEqualsIgnoreCase family "consolas"
                       || trimmedEqualsIgnoreCase family "courier new"
                       || trimmedEqualsIgnoreCase family "courier" -> monospaceLineBoxFactor
    | Some family when trimmedEqualsIgnoreCase family "ubuntu" -> ubuntuLineBoxFactor
    | _ -> wordLineBoxFactor

let private tableCellLineBoxFactorForTextStyle hasExplicitLineSpacing (style: TextBuffer.TextStyle) =
    if not hasExplicitLineSpacing then
        tableCellDefaultLineBoxFactor
    else
        match style.FontFamily with
        | Some family when trimmedEqualsIgnoreCase family "arial" -> arialTableCellExplicitLineBoxFactor
        | _ -> tableCellExplicitLineBoxFactor

let private usesDefaultParagraphSpacing (metadata: BlockMetadata) =
    metadata.SpacingBefore.IsNone && metadata.SpacingAfter.IsNone

let private isTocEntryStyle (metadata: BlockMetadata) =
    metadata.Style
    |> Option.exists (fun style ->
        style.StartsWith("TOC", System.StringComparison.OrdinalIgnoreCase)
        && hasDigitsFrom 3 style)

let private isHeadingStyle (metadata: BlockMetadata) =
    metadata.Style
    |> Option.exists (fun style ->
        style.StartsWith("Heading", System.StringComparison.OrdinalIgnoreCase)
        && hasDigitsFrom 7 style)

let private headingUsesDefaultLineBox (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
    let before = metadata.SpacingBefore |> Option.orElse paragraphStyle.SpacingBefore
    let after = metadata.SpacingAfter |> Option.orElse paragraphStyle.SpacingAfter
    let autoSpacing =
        before |> Option.exists (fun value -> abs (value - wordAutoSpacingPoints) < 0.001)
        && after |> Option.exists (fun value -> abs (value - wordAutoSpacingPoints) < 0.001)
    isHeadingStyle metadata
    && not autoSpacing

let private wordLineHeight textStyle fontSize lineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox =
    let factor =
        if abs (lineSpacing - 1.0) < 0.001 then
            if tocEntry then wordTocEntrySingleLineBoxFactor
            elif headingDefaultLineBox then wordDefaultSpacingSingleLineBoxFactor
            elif defaultParagraphSpacing then wordDefaultSpacingSingleLineBoxFactor
            else wordSingleLineBoxFactor
        else
            lineBoxFactorForTextStyle textStyle
    fontSize * lineSpacing * factor

let private lineHeightForSettings (settings: LayoutSettings) textStyle fontSize lineSpacing lineSpacingRule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox =
    // For "exact"/"atLeast" lineSpacing is the height in px (see
    // resolveLineSpacing) and the automatic box is single spacing.
    let absolute = lineSpacingRule = "exact" || lineSpacingRule = "atLeast"
    let multiple = if absolute then 1.0 else lineSpacing
    let automatic =
        if settings.InTableCell then
            let factor = tableCellLineBoxFactorForTextStyle hasExplicitLineSpacing textStyle
            tableCellLineBoxFontSize fontSize * multiple * factor
        else
            wordLineHeight textStyle fontSize multiple defaultParagraphSpacing tocEntry headingDefaultLineBox

    match lineSpacingRule with
    | "exact" -> max 1.0 lineSpacing
    | "atLeast" -> max automatic lineSpacing
    | _ -> automatic

/// Line spacing resolved to what layout needs: the multiple used for the
/// automatic line box, the rule, and (for "exact"/"atLeast") the line height
/// in layout pixels.
///
/// The importer delivers the value in two different units and the unit follows
/// where the value came from, so value and rule are resolved together:
///   * paragraph-level (BlockMetadata.LineSpacing): "auto" is a multiple of
///     single spacing (w:line / 240). For "exact"/"atLeast" it is ALSO a
///     multiple — the line height in px divided by the paragraph font size
///     (DocxReader.wordLineSpacingToMultiple), with that font size delivered
///     as ParagraphMarkFontSize;
///   * style-level (NamedStyle.LineSpacing): "auto" is the same multiple, but
///     "exact"/"atLeast" is the absolute height in px (twips / 15).
/// A paragraph that sets w:line without w:lineRule is "auto" even when its
/// style says "exact".
let private resolveLineSpacing (paragraphStyle: ParagraphStyle) (rawMetadata: BlockMetadata) contentFontSize =
    let isAbsoluteRule rule = rule = Some "exact" || rule = Some "atLeast"
    match rawMetadata.LineSpacing with
    | Some value ->
        if isAbsoluteRule rawMetadata.LineSpacingRule then
            let fontSize = rawMetadata.ParagraphMarkFontSize |> Option.defaultValue contentFontSize |> max 1.0
            struct (1.0, rawMetadata.LineSpacingRule.Value, value * fontSize)
        else
            struct (value, "auto", 0.0)
    | None ->
        match paragraphStyle.LineSpacing with
        | Some value ->
            let rule = rawMetadata.LineSpacingRule |> Option.orElse paragraphStyle.LineSpacingRule
            let styleValueIsPixels = isAbsoluteRule paragraphStyle.LineSpacingRule
            if isAbsoluteRule rule then
                struct (1.0, rule.Value, (if styleValueIsPixels then value else value * max 1.0 contentFontSize))
            elif styleValueIsPixels then
                struct (1.0, "auto", 0.0)
            else
                struct (value, "auto", 0.0)
        | None -> struct (1.0, "auto", 0.0)

let private resolvedLineHeightForSettings (settings: LayoutSettings) (paragraphStyle: ParagraphStyle) (rawMetadata: BlockMetadata) textStyle fontSize hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox =
    let struct (multiple, rule, pixels) = resolveLineSpacing paragraphStyle rawMetadata fontSize
    lineHeightForSettings settings textStyle fontSize (if rule = "auto" then multiple else pixels) rule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox

let private lineMaxTextFontSize fallback (line: MeasuredLine) =
    let mutable maxSize = fallback
    let mutable foundText = false
    for segment in line.Segments do
        match segment with
        | TextSegment (_, style, _, _, _) ->
            let fontSize = WordLayout.Text.displayFontSize style
            if not foundText || fontSize > maxSize then
                maxSize <- fontSize
                foundText <- true
        | ImageSegment _ | MathSegment _ -> ()
    maxSize

let private startsAtPageContentTop settings y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    abs (y - WordLayout.Page.contentTop settings.PaddingTop pageTop) < 0.5

let private suppressLayoutPageTopBeforeSpacing settings paragraphStyle (layout: BlockLayout) =
    let spacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle layout.Metadata
    match layout.Lines with
    | firstPaged :: _ when spacingBefore > 0.0 && startsAtPageContentTop settings layout.Y ->
        let firstPagedPageTop = WordLayout.Page.pageTopForY (pageHeight settings) firstPaged.Y
        let firstPagedContentTop = WordLayout.Page.contentTop settings.PaddingTop firstPagedPageTop
        let lineTopOffset = firstPaged.Y - firstPagedContentTop
        if abs (lineTopOffset - spacingBefore) > 0.5 then layout
        else
            let shift = min spacingBefore (max 0.0 lineTopOffset)
            let firstPagedPage = pageIndexForY settings firstPaged.Y
            { layout with
                Height = max 0.0 (layout.Height - shift)
                Lines =
                    layout.Lines
                    |> List.map (fun line ->
                        if pageIndexForY settings line.Y = firstPagedPage then { line with Y = line.Y - shift }
                        else line) }
    | _ -> layout

let private paragraphBorderAdvance (border: ParagraphBorder option) =
    match border with
    | Some value ->
        if optionEqualsIgnoreCase "none" value.Style || optionEqualsIgnoreCase "nil" value.Style then 0.0
        else (value.Width |> Option.defaultValue 0.5 |> max 0.25) + (value.Space |> Option.defaultValue 0.0)
    | None -> 0.0

let private paragraphBorderVerticalAdvance metadata =
    let visibleBefore, visibleAfter =
        match metadata.ParagraphBorders with
        | Some borders -> paragraphBorderAdvance borders.Top, paragraphBorderAdvance borders.Bottom
        | None -> 0.0, 0.0
    metadata.ParagraphBorderBefore |> Option.defaultValue visibleBefore,
    metadata.ParagraphBorderAfter |> Option.defaultValue visibleAfter

let private paragraphBorderRightAdvance metadata =
    match metadata.ParagraphBorders with
    | Some borders -> paragraphBorderAdvance borders.Right
    | None -> 0.0

let segmentHeight lineHeightFactor = function
    | TextSegment (_, style, _, _, _) -> WordLayout.Text.displayFontSize style * lineHeightFactor
    | ImageSegment (_, _, yOffset, _, height, _, floating, _, _, _, _, _, _) ->
        if floating then 0.0 else height + max 0.0 yOffset
    | MathSegment (_, _, _, _, ascent, descent, _, _) -> ascent + descent

let segmentStart = function
    | TextSegment (_, _, _, _, startOffset) -> startOffset
    | ImageSegment (_, _, _, _, _, startOffset, _, _, _, _, _, _, _) -> startOffset
    | MathSegment (_, _, _, _, _, _, startOffset, _) -> startOffset

let segmentEnd = function
    | TextSegment (text, _, _, _, startOffset) -> startOffset + text.Length
    | ImageSegment (_, _, _, _, _, startOffset, _, _, _, _, _, _, _) -> startOffset + 1
    | MathSegment (_, _, _, _, _, _, startOffset, length) -> startOffset + length

let segmentX = function
    | TextSegment (_, _, x, _, _) -> x
    | ImageSegment (_, x, _, _, _, _, _, _, _, _, _, _, _) -> x
    | MathSegment (_, _, x, _, _, _, _, _) -> x

let segmentWidth = function
    | TextSegment (_, _, _, width, _) -> width
    | ImageSegment (_, _, _, width, _, _, _, _, _, _, _, _, _) -> width
    | MathSegment (_, _, _, width, _, _, _, _) -> width

let alignmentOffset (layout: BlockLayout) (line: MeasuredLine) =
    match layout.Metadata.Alignment with
    // An RTL paragraph's default ("start") alignment is the right margin.
    | Left -> if layout.Metadata.Rtl then max 0.0 (layout.AvailableWidth - line.Width) else 0.0
    | Center -> max 0.0 ((layout.AvailableWidth - line.Width) / 2.0)
    | Right -> if layout.Metadata.Rtl then 0.0 else max 0.0 (layout.AvailableWidth - line.Width)
    // Justify stretches inter-word space so the visible content reaches the
    // right margin (baked into segment x by justifyLineToEdge), so no extra
    // whole-line shift is needed here — the line already starts at the margin.
    | Justify -> 0.0

/// Reposition a line's segments into right-to-left visual order (for w:bidi
/// paragraphs). The list stays in logical order (so offset→x mapping is intact);
/// each segment's x is reassigned so the first logical token is rightmost. The
/// browser/shaper already renders each token's glyphs in correct script order,
/// so reversing the token positions yields correct RTL reading order. Combined
/// with right alignment via alignmentOffset.
let private repositionRtl (line: MeasuredLine) : MeasuredLine =
    match line.Segments with
    | [] -> line
    | segs ->
        let mutable baseX = System.Double.PositiveInfinity
        let mutable totalW = 0.0
        for segment in segs do
            baseX <- min baseX (segmentX segment)
            totalW <- totalW + segmentWidth segment
        let mutable cursor = baseX + totalW
        let newSegs = ResizeArray<Segment>()
        for segment in segs do
            let w = segmentWidth segment
            cursor <- cursor - w
            newSegs.Add(
                match segment with
                | TextSegment (t, st, _, ww, off) -> TextSegment (t, st, cursor, ww, off)
                | ImageSegment (src, _, yOff, ww, h, off, fl, wr, rh, cl, ct, cr, cb) -> ImageSegment (src, cursor, yOff, ww, h, off, fl, wr, rh, cl, ct, cr, cb)
                | MathSegment (p, st, _, ww, asc, desc, off, len) -> MathSegment (p, st, cursor, ww, asc, desc, off, len))
        { line with Segments = resizeArrayToList newSegs }

/// Distribute the slack between a line's visible right edge and `rightEdge`
/// across its inter-word gaps, returning the line with stretched segment x.
/// `rightEdge` is in the same (post-indent) coordinate frame as the line's
/// segment x. A "gap" is a non-final segment whose token ends in a regular
/// space; tabs, nbsp, and images are not stretched. Caller excludes the
/// paragraph's last line (Word never justifies it).
let private justifyLineToEdge (rightEdge: float) (line: MeasuredLine) : MeasuredLine =
    let segs = line.Segments
    let endsInStretchSpace = function
        | TextSegment (t, _, _, _, _) when t.Length > 0 -> t.[t.Length - 1] = ' '
        | _ -> false
    // A tab advances to a fixed stop and an image is atomic — neither stretches,
    // and a line containing one is left ragged (Word stretches only the spaces,
    // which this simple model doesn't isolate around the fixed element).
    let mutable hasTabOrImage = false
    let mutable gapCount = 0
    let mutable baseX = System.Double.PositiveInfinity
    let rec scan previous remaining =
        match remaining with
        | [] ->
            previous |> Option.iter (fun segment ->
                baseX <- min baseX (segmentX segment))
        | segment :: rest ->
            baseX <- min baseX (segmentX segment)
            match previous with
            | Some prior when endsInStretchSpace prior -> gapCount <- gapCount + 1
            | _ -> ()
            match segment with
            | ImageSegment _ -> hasTabOrImage <- true
            | MathSegment _ -> hasTabOrImage <- true
            | TextSegment (t, _, _, _, _) when t.Contains("\t") -> hasTabOrImage <- true
            | TextSegment _ -> ()
            scan (Some segment) rest
    scan None segs
    let baseX = if System.Double.IsPositiveInfinity baseX then 0.0 else baseX
    let extra = rightEdge - (baseX + line.Width)
    if hasTabOrImage || gapCount = 0 || extra <= 0.5 then line
    else
        let perGap = extra / float gapCount
        let mutable shift = 0.0
        let newSegs = ResizeArray<Segment>()
        let rec build remaining =
            match remaining with
            | [] -> ()
            | segment :: rest ->
                newSegs.Add(
                    match segment with
                    | TextSegment (t, st, x, w, off) -> TextSegment (t, st, x + shift, w, off)
                    | ImageSegment (src, x, yOff, w, h, off, fl, wr, rh, cl, ct, cr, cb) ->
                        ImageSegment (src, x + shift, yOff, w, h, off, fl, wr, rh, cl, ct, cr, cb)
                    | MathSegment (p, st, x, w, asc, desc, off, len) -> MathSegment (p, st, x + shift, w, asc, desc, off, len))
                if (match rest with [] -> false | _ -> true) && endsInStretchSpace segment then
                    shift <- shift + perGap
                build rest
        build segs
        { line with Segments = resizeArrayToList newSegs }

let emptyLine y height : MeasuredLine = {
    Segments = []
    Y = y
    Width = 0.0
    Height = height
    StartOffset = 0
    EndOffset = 0
}

let private maxLineBottom defaultValue (layouts: BlockLayout list) =
    let mutable found = false
    let mutable bottom = 0.0
    for layout in layouts do
        for line in layout.Lines do
            let lineBottom = line.Y + line.Height
            if not found || lineBottom > bottom then
                found <- true
                bottom <- lineBottom
    if found then bottom else defaultValue

let private maxLineBottomAndLayoutCount defaultValue (layouts: BlockLayout list) =
    let mutable found = false
    let mutable bottom = 0.0
    let mutable layoutCount = 0
    for layout in layouts do
        layoutCount <- layoutCount + 1
        for line in layout.Lines do
            let lineBottom = line.Y + line.Height
            if not found || lineBottom > bottom then
                found <- true
                bottom <- lineBottom
    (if found then bottom else defaultValue), layoutCount

let private lineForOffsetArray (layout: BlockLayout) (lines: MeasuredLine array) (offset: int) =
    if lines.Length = 0 then
        emptyLine layout.Y layout.Height
    elif offset <= lines.[0].StartOffset then
        lines.[0]
    else
        let mutable low = 0
        let mutable high = lines.Length
        while low < high do
            let mid = low + (high - low) / 2
            if lines.[mid].EndOffset <= offset then
                low <- mid + 1
            else
                high <- mid
        if low >= lines.Length then
            lines.[lines.Length - 1]
        elif offset < lines.[low].StartOffset && low > 0 then
            lines.[low - 1]
        else
            lines.[low]

let private lineForOffsetList (layout: BlockLayout) (offset: int) : MeasuredLine =
    match layout.Lines with
    | [] -> emptyLine layout.Y layout.Height
    | first :: _ when offset <= first.StartOffset -> first
    | first :: rest ->
        let mutable candidate = first
        let mutable result = first
        let mutable found = offset < first.EndOffset
        let mutable remaining = rest
        let mutable keepGoing = not found
        while keepGoing do
            match remaining with
            | [] ->
                result <- candidate
                keepGoing <- false
            | line :: tail ->
                if offset < line.StartOffset then
                    result <- candidate
                    keepGoing <- false
                elif offset < line.EndOffset then
                    result <- line
                    keepGoing <- false
                else
                    candidate <- line
                    remaining <- tail
        if found then first else result

let lineForOffset (layout: BlockLayout) (offset: int) : MeasuredLine =
    lineForOffsetList layout offset

let maxOffsetInLayout (layout: BlockLayout) =
    match layout.Lines with
    | [] -> 0
    | first :: rest ->
        let mutable maxOffset = first.EndOffset
        for line in rest do
            if line.EndOffset > maxOffset then
                maxOffset <- line.EndOffset
        maxOffset

let private xForOffsetWithRangeUncached (measureRange: TextRangeMeasurer) (line: MeasuredLine) (offset: int) =
    let mutable x = 0.0
    let mutable found = false

    for segment in line.Segments do
        if not found then
            let segStart = segmentStart segment
            let segEnd = segmentEnd segment

            if offset <= segStart then
                x <- segmentX segment
                found <- true
            elif offset < segEnd then
                match segment with
                | TextSegment (text, style, segX, _, startOffset) ->
                    let charCount = max 0 (min text.Length (offset - startOffset))
                    x <- segX + measureRange (WordLayout.Text.displayText style text) (WordLayout.Text.displayStyle style) 0 charCount
                | ImageSegment (_, segX, _, width, _, startOffset, _, _, _, _, _, _, _) ->
                    x <- if offset <= startOffset then segX else segX + width
                | MathSegment (_, _, segX, width, _, _, startOffset, _) ->
                    x <- if offset <= startOffset then segX else segX + width
                found <- true
            else
                x <- segmentX segment + segmentWidth segment

    x

let xForOffsetWithRange (measureRange: TextRangeMeasurer) (line: MeasuredLine) (offset: int) =
    xForOffsetWithRangeUncached measureRange line offset

let xForOffset (measureText: TextMeasurer) (line: MeasuredLine) (offset: int) =
    xForOffsetWithRange (rangeMeasurerFromTextMeasurer measureText) line offset

let private layoutContainsY mouseY (layout: BlockLayout) =
    mouseY >= layout.Y && mouseY < layout.Y + layout.Height

let private lineContainsY (mouseY: float) (line: MeasuredLine) =
    mouseY >= line.Y && mouseY < line.Y + line.Height

let private tryFindLineForYList (lines: MeasuredLine list) (mouseY: float) =
    let mutable found = false
    let mutable result = Unchecked.defaultof<MeasuredLine>
    let mutable remaining = lines
    while not found && not remaining.IsEmpty do
        match remaining with
        | line :: tail ->
            if lineContainsY mouseY line then
                found <- true
                result <- line
            remaining <- tail
        | [] -> ()
    if found then ValueSome result else ValueNone

let private tryFindLineForYArray (lines: MeasuredLine array) (mouseY: float) =
    let mutable low = 0
    let mutable high = lines.Length
    while low < high do
        let mid = low + (high - low) / 2
        let line = lines.[mid]
        if line.Y + line.Height <= mouseY then
            low <- mid + 1
        else
            high <- mid

    let mutable index = low
    while index > 0 && lines.[index - 1].Y + lines.[index - 1].Height > mouseY do
        index <- index - 1

    let mutable found = false
    let mutable result = Unchecked.defaultof<MeasuredLine>
    while not found && index < lines.Length && lines.[index].Y <= mouseY do
        let line = lines.[index]
        if lineContainsY mouseY line then
            found <- true
            result <- line
        index <- index + 1
    if found then ValueSome result else ValueNone

let private tryFindLineForY (layout: BlockLayout) (mouseY: float) =
    tryFindLineForYList layout.Lines mouseY

let private hitTestLayoutWithRangeValue (measureRange: TextRangeMeasurer) (paddingLeft: float) (layout: BlockLayout) (mouseX: float) (mouseY: float) =
    let lineOpt = tryFindLineForY layout mouseY
    match lineOpt with
    | ValueNone -> ValueSome (struct (layout.BlockId, 0))
    | ValueSome line ->
        let relativeX = mouseX - paddingLeft - alignmentOffset layout line
        let mutable offset = line.EndOffset
        let mutable found = false

        if relativeX <= 0.0 then
            offset <- line.StartOffset
            found <- true

        for segment in line.Segments do
            if not found then
                match segment with
                | TextSegment (text, style, x, width, startOffset) ->
                    if relativeX <= x + width then
                        let displayText = WordLayout.Text.displayText style text
                        let displayStyle = WordLayout.Text.displayStyle style
                        let mutable low = 0
                        let mutable high = text.Length
                        while low < high do
                            let mid = (low + high) / 2
                            let textWidth = measureRange displayText displayStyle 0 mid
                            if x + textWidth < relativeX then
                                low <- mid + 1
                            else
                                high <- mid
                        // `low` is the first boundary at or right of the click,
                        // i.e. the far side of the clicked glyph. Put the caret
                        // on whichever side of that glyph is nearer.
                        if low > 0 then
                            let rightEdge =
                                if low >= text.Length then x + width
                                else x + measureRange displayText displayStyle 0 low
                            let leftEdge = x + measureRange displayText displayStyle 0 (low - 1)
                            if relativeX - leftEdge < rightEdge - relativeX then
                                low <- low - 1
                        offset <- startOffset + low
                        found <- true
                    else
                        offset <- startOffset + text.Length

                | ImageSegment (_, x, _, width, _, startOffset, _, _, _, _, _, _, _) ->
                    if relativeX <= x + width then
                        let midpoint = x + width / 2.0
                        offset <- if relativeX < midpoint then startOffset else startOffset + 1
                        found <- true
                    else
                        offset <- startOffset + 1

                | MathSegment (_, _, x, width, _, _, startOffset, length) ->
                    if relativeX <= x + width then
                        let midpoint = x + width / 2.0
                        offset <- if relativeX < midpoint then startOffset else startOffset + length
                        found <- true
                    else
                        offset <- startOffset + length

        ValueSome (struct (layout.BlockId, offset))

let private hitTestLayoutWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layout: BlockLayout) (mouseX: float) (mouseY: float) =
    match hitTestLayoutWithRangeValue measureRange paddingLeft layout mouseX mouseY with
    | ValueSome (struct (blockId, offset)) -> Some (blockId, offset)
    | ValueNone -> None

/// How far a point is, sideways, from where a layout's text stands (0 inside it).
let private horizontalDistanceToLayout (paddingLeft: float) (layout: BlockLayout) (mouseX: float) =
    let left = paddingLeft + layout.Indent
    let right = left + max 0.0 layout.AvailableWidth
    if mouseX < left then left - mouseX elif mouseX > right then mouseX - right else 0.0

/// The layout a point is in: of those at the point's height - the cells of a
/// table row, the columns of a section - the one the point is in sideways, or
/// nearest to. (It was the first at that height: a click in the second cell of
/// a row, or in the second column of a page, went to the first.)
let private layoutAtPoint (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) : BlockLayout voption =
    let mutable best = ValueNone
    let mutable bestDistance = System.Double.PositiveInfinity
    let mutable remaining = layouts
    while bestDistance > 0.0 && not remaining.IsEmpty do
        let layout = remaining.Head
        if layoutContainsY mouseY layout then
            let distance = horizontalDistanceToLayout paddingLeft layout mouseX
            if distance < bestDistance then
                best <- ValueSome layout
                bestDistance <- distance
        remaining <- remaining.Tail
    best

let hitTestLayoutsWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) =
    match layoutAtPoint paddingLeft layouts mouseX mouseY with
    | ValueNone -> None
    | ValueSome layout -> hitTestLayoutWithRange measureRange paddingLeft layout mouseX mouseY

let hitTestLayoutArrayWithRangeValue (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout array) (mouseX: float) (mouseY: float) =
    let mutable low = 0
    let mutable high = layouts.Length
    while low < high do
        let mid = (low + high) / 2
        let layout = layouts.[mid]
        if layout.Y + layout.Height <= mouseY then
            low <- mid + 1
        else
            high <- mid

    let mutable index = low
    while index > 0 && layouts.[index - 1].Y + layouts.[index - 1].Height > mouseY do
        index <- index - 1

    // (Of the layouts at that height - the cells of a row - the one the point is in sideways, or nearest to.)
    let mutable best = ValueNone
    let mutable bestDistance = System.Double.PositiveInfinity
    while bestDistance > 0.0 && index < layouts.Length && layouts.[index].Y <= mouseY do
        let layout = layouts.[index]
        if layoutContainsY mouseY layout then
            let distance = horizontalDistanceToLayout paddingLeft layout mouseX
            if distance < bestDistance then
                best <- ValueSome layout
                bestDistance <- distance
        index <- index + 1
    match best with
    | ValueSome layout -> hitTestLayoutWithRangeValue measureRange paddingLeft layout mouseX mouseY
    | ValueNone -> ValueNone

let hitTestLayoutArrayWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout array) (mouseX: float) (mouseY: float) =
    match hitTestLayoutArrayWithRangeValue measureRange paddingLeft layouts mouseX mouseY with
    | ValueSome (struct (blockId, offset)) -> Some (blockId, offset)
    | ValueNone -> None

let hitTestLayouts (measureText: TextMeasurer) (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) =
    hitTestLayoutsWithRange (rangeMeasurerFromTextMeasurer measureText) paddingLeft layouts mouseX mouseY

let hitTestForImageLayouts (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let mutable result = None
    for layout in layouts do
        if result.IsNone then
            for line in layout.Lines do
                if result.IsNone then
                    let lineAlignmentOffset = alignmentOffset layout line
                    let relativeX = mouseX - paddingLeft - lineAlignmentOffset
                    for segment in line.Segments do
                        if result.IsNone then
                            match segment with
                            | ImageSegment (_, x, yOffset, width, height, startOffset, _, _, _, _, _, _, _) ->
                                let imageY = line.Y + yOffset
                                if mouseY >= imageY && mouseY <= imageY + height && relativeX >= x && relativeX <= x + width then
                                    result <- Some (layout.BlockId, startOffset, width, height, paddingLeft + x + lineAlignmentOffset, imageY)
                            | TextSegment _ -> ()
                            | MathSegment _ -> ()
    result

let private hitTestForImageLayout (paddingLeft: float) (layout: BlockLayout) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let mutable result = None
    for line in layout.Lines do
        if result.IsNone then
            let lineAlignmentOffset = alignmentOffset layout line
            let relativeX = mouseX - paddingLeft - lineAlignmentOffset
            for segment in line.Segments do
                if result.IsNone then
                    match segment with
                    | ImageSegment (_, x, yOffset, width, height, startOffset, _, _, _, _, _, _, _) ->
                        let imageY = line.Y + yOffset
                        if mouseY >= imageY && mouseY <= imageY + height && relativeX >= x && relativeX <= x + width then
                            result <- Some (layout.BlockId, startOffset, width, height, paddingLeft + x + lineAlignmentOffset, imageY)
                    | TextSegment _ -> ()
                    | MathSegment _ -> ()
    result

let hitTestForImageLayoutArray (paddingLeft: float) (layouts: BlockLayout array) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let mutable low = 0
    let mutable high = layouts.Length
    while low < high do
        let mid = (low + high) / 2
        let layout = layouts.[mid]
        if layout.Y + layout.Height < mouseY then
            low <- mid + 1
        else
            high <- mid

    let mutable index = low
    while index > 0 && layouts.[index - 1].Y + layouts.[index - 1].Height >= mouseY do
        index <- index - 1

    let mutable result = None
    while result.IsNone && index < layouts.Length && layouts.[index].Y <= mouseY do
        result <- hitTestForImageLayout paddingLeft layouts.[index] mouseX mouseY
        index <- index + 1
    result

let imageBoxInLayout (paddingLeft: float) (layout: BlockLayout) (offset: int) : (float * float * float * float) option =
    let mutable result = None

    for line in layout.Lines do
        if result.IsNone then
            let lineAlignmentOffset = alignmentOffset layout line
            for segment in line.Segments do
                if result.IsNone then
                    match segment with
                    | ImageSegment (_, segX, yOffset, width, height, startOffset, _, _, _, _, _, _, _) when startOffset = offset ->
                        result <- Some (paddingLeft + segX + lineAlignmentOffset, line.Y + yOffset, width, height)
                    | _ -> ()

    result

let imageBoxForOffset (paddingLeft: float) (layouts: BlockLayout list) (blockId: BlockId) (offset: int) : (float * float * float * float) option =
    match layouts |> List.tryFind (fun l -> l.BlockId = blockId) with
    | None -> None
    | Some layout -> imageBoxInLayout paddingLeft layout offset

type FloatingObstacle = float * float * float * float
type PlacedFloatingBox = WordLayout.Floating.Box

let private isWhitespaceChar (ch: char) =
    ch = ' ' || ch = '\t' || ch = '\r' || ch = '\n'

let private shortTrailingSpaceFitSafety = 0.25
let private lineFitSafety = 0.25
let private tableCellLineFitSafety = 2.0
let private mixedMetricsLineFitSafety = 0.75

let private visibleTokenLength (text: string) =
    let mutable length = text.Length
    while length > 0 && isWhitespaceChar text.[length - 1] do
        length <- length - 1
    length

let private visibleTokenLengthRange (text: string) start length =
    let mutable visible = length
    while visible > 0 && isWhitespaceChar text.[start + visible - 1] do
        visible <- visible - 1
    visible

let private rangeContainsTab (text: string) start length =
    let stop = start + length
    let mutable index = start
    let mutable found = false
    while not found && index < stop do
        if text.[index] = '\t' then
            found <- true
        index <- index + 1
    found

let private trailingSpaceFitSafety settings visibleLength =
    if visibleLength <= 3 then shortTrailingSpaceFitSafety else settings.LongTrailingSpaceFitSafety

let private iterWordWrapTokenRanges (text: string) (f: int -> int -> unit) =
    let mutable index = 0
    while index < text.Length do
        let start = index
        if text.[index] = '\r' || text.[index] = '\n' then
            if text.[index] = '\r' && index + 1 < text.Length && text.[index + 1] = '\n' then
                index <- index + 2
            else
                index <- index + 1
            f start (index - start)
        elif text.[index] = '\t' then
            // Every tab is its own token: each one advances to the next tab
            // stop. A run of tabs (or spaces followed by a tab) in one token
            // would be given the width of a single tab.
            index <- index + 1
            f start 1
        elif isWhitespaceChar text.[index] then
            while index < text.Length && isWhitespaceChar text.[index] && text.[index] <> '\r' && text.[index] <> '\n' && text.[index] <> '\t' do
                index <- index + 1
            f start (index - start)
        else
            let mutable keepGoing = true
            while keepGoing && index < text.Length && not (isWhitespaceChar text.[index]) do
                let isHyphenBreak = text.[index] = '-' && index + 1 < text.Length && not (isWhitespaceChar text.[index + 1])
                index <- index + 1
                if isHyphenBreak then
                    keepGoing <- false
            // Absorb trailing whitespace into the same token, but exclude tabs:
            // tabs always advance to the next tab stop and need their own segment
            // so the paint backend doesn't stretch a "word + tab" segment's
            // glyphs (eg. TOC entry "1\t" being scaled by width/advance into a
            // wide blob).
            while index < text.Length && isWhitespaceChar text.[index] && text.[index] <> '\r' && text.[index] <> '\n' && text.[index] <> '\t' do
                index <- index + 1
            f start (index - start)

/// Layout buffer into lines with styled runs and images.
let layoutLinesStyledWithObstaclesWithRange (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) (contentY: float) (tabStops: TabStop list) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    let mutable lines = []
    let currentSegments = ResizeArray<Segment>()
    let mutable currentHasTextMetric = false
    let mutable currentFirstFontFamily = None
    let mutable currentFirstFontSize = 0.0
    let mutable currentFirstBold = false
    let mutable currentFirstItalic = false
    let mutable currentFirstVertAlign = None
    let mutable currentFirstSmallCaps = false
    let mutable currentHasMixedTextMetrics = false
    let mutable currentHasFlowContent = false
    let mutable currentX = 0.0
    let mutable y = minBodyTop settings 0.0
    let mutable flushedLineCount = 0
    let availableWidth = max 1.0 maxWidth
    let mutable currentLineEnd = availableWidth
    let mutable floatObstacles = initialObstacles
    let mutable placedBoxes = initialPlacedBoxes

    // A float is "already placed" when the anchor-extended second pass handed
    // us the box this very image run produced in the first pass. Identity is
    // the image run itself, not its Src: the same picture used twice is two
    // floats, and the second one must still become a wrap obstacle.
    let isPrePlacedImage (imgRun: TextBuffer.ImageRun) =
        let mutable found = false
        for box in initialPlacedBoxes do
            match box.Kind with
            | WordLayout.Floating.ImageBox image when obj.ReferenceEquals(image, imgRun) -> found <- true
            | _ -> ()
        found

    let noteTextMetric style =
        if not currentHasTextMetric then
            currentHasTextMetric <- true
            currentFirstFontFamily <- style.FontFamily
            currentFirstFontSize <- style.FontSize
            currentFirstBold <- style.Bold
            currentFirstItalic <- style.Italic
            currentFirstVertAlign <- style.VertAlign
            currentFirstSmallCaps <- style.SmallCaps
        elif not currentHasMixedTextMetrics
             && (currentFirstFontFamily <> style.FontFamily
                 || currentFirstFontSize <> style.FontSize
                 || currentFirstBold <> style.Bold
                 || currentFirstItalic <> style.Italic
                 || currentFirstVertAlign <> style.VertAlign
                 || currentFirstSmallCaps <> style.SmallCaps) then
            currentHasMixedTextMetrics <- true

    let wrapIntervalsForY lineY =
        // Return the list of free (start, end) intervals at this line y, ordered
        // left-to-right. Empty list means the entire line is blocked.
        let globalLineY = contentY + (lineY - settings.PaddingTop)
        let lineBottom = globalLineY + settings.LineHeight
        let active = ResizeArray<float * float>()
        let addActiveInterval startX endX =
            let mutable startX = startX
            let mutable endX = endX
            let mutable i = 0
            let mutable inserted = false
            while not inserted && i < active.Count do
                let activeStart, activeEnd = active.[i]
                if endX < activeStart then
                    active.Insert(i, (startX, endX))
                    inserted <- true
                elif startX > activeEnd then
                    i <- i + 1
                else
                    startX <- min startX activeStart
                    endX <- max endX activeEnd
                    active.RemoveAt i
            if not inserted then
                active.Add((startX, endX))
        for (x, top, width, height) in floatObstacles do
            if lineBottom > top && globalLineY < top + height then
                addActiveInterval x (x + width)
        let mutable gaps = []
        let mutable cursor = 0.0
        for (oxStart, oxEnd) in active do
            let gapEnd = min availableWidth oxStart
            if gapEnd > cursor then gaps <- (cursor, gapEnd) :: gaps
            cursor <- max cursor oxEnd
        if cursor < availableWidth then gaps <- (cursor, availableWidth) :: gaps
        // (No text in a band beside a floating object that is narrower than an inch - see
        // minimumWrapBand; a line with no band wide enough goes below the object.)
        if active.Count = 0 then List.rev gaps
        else
            let least = min WordLayout.Floating.minimumWrapBand availableWidth - 0.01
            gaps |> List.filter (fun (startX, endX) -> endX - startX >= least) |> List.rev

    let nextObstacleClearY lineY =
        let globalLineY = contentY + (lineY - settings.PaddingTop)
        let lineBottom = globalLineY + settings.LineHeight
        let mutable clearY = None
        for (_, top, _, height) in floatObstacles do
            let bottom = top + height
            if lineBottom > top && globalLineY < bottom then
                match clearY with
                | None -> clearY <- Some bottom
                | Some current when bottom < current -> clearY <- Some bottom
                | Some _ -> ()
        clearY

    // Intervals at the current y that we have not yet started filling. The
    // currentX/currentLineEnd pair always reflects the interval being filled.
    let mutable pendingIntervals: (float * float) list = []

    let private_setIntervalToFirst () =
        match pendingIntervals with
        | (s, e) :: rest ->
            currentX <- s
            currentLineEnd <- e
            pendingIntervals <- rest
        | [] ->
            currentX <- availableWidth
            currentLineEnd <- availableWidth

    let moveToWrappableLine () =
        if floatObstacles.IsEmpty then
            pendingIntervals <- []
            currentX <- 0.0
            currentLineEnd <- availableWidth
        else
            pendingIntervals <- wrapIntervalsForY y
            private_setIntervalToFirst ()
            while currentX >= currentLineEnd do
                match nextObstacleClearY y with
                | Some clearGlobalY ->
                    let clearY = settings.PaddingTop + (clearGlobalY - contentY)
                    if clearY > y + 0.01 then y <- clearY else y <- y + settings.LineHeight
                | None ->
                    y <- y + settings.LineHeight
                pendingIntervals <- wrapIntervalsForY y
                private_setIntervalToFirst ()

    let coalescedSegmentList (segments: ResizeArray<Segment>) =
        let output = ResizeArray<Segment>(segments.Count)
        let mutable pendingText: System.Text.StringBuilder = null
        let mutable pendingStyle = Unchecked.defaultof<TextBuffer.TextStyle>
        let mutable pendingX = 0.0
        let mutable pendingWidth = 0.0
        let mutable pendingStart = 0
        let flushPending () =
            if not (isNull pendingText) then
                output.Add(TextSegment (pendingText.ToString(), pendingStyle, pendingX, pendingWidth, pendingStart))
                pendingText <- null
                pendingWidth <- 0.0
                pendingStart <- 0
        for segment in segments do
            match segment with
            | TextSegment (text, style, x, width, startOffset)
                when not (rangeContainsTab text 0 text.Length) ->
                if not (isNull pendingText)
                   && pendingStyle = style
                   && pendingStart + pendingText.Length = startOffset
                   && abs ((pendingX + pendingWidth) - x) < 0.01 then
                    pendingText.Append(text) |> ignore
                    pendingWidth <- pendingWidth + width
                else
                    flushPending ()
                    pendingText <- System.Text.StringBuilder(text.Length)
                    pendingText.Append(text) |> ignore
                    pendingStyle <- style
                    pendingX <- x
                    pendingWidth <- width
                    pendingStart <- startOffset
            | _ ->
                flushPending ()
                output.Add segment
        flushPending ()
        resizeArrayToList output

    let flushLine () =
        if currentSegments.Count > 0 then
            let scale = settings.LineHeight / settings.FontSize
            let mutable hasInlineImage = false
            let mutable segmentMaxHeight = 0.0
            let mutable startOffset = 0
            let mutable endOffset = 0
            let mutable widthEnd = 0.0
            let mutable isFirstSegment = true
            for segment in currentSegments do
                if isFirstSegment then
                    startOffset <- segmentStart segment
                    isFirstSegment <- false
                endOffset <- segmentEnd segment
                match segment with
                | ImageSegment (_, _, _, _, _, _, floating, _, _, _, _, _, _) when not floating ->
                    hasInlineImage <- true
                | _ -> ()
                match segment with
                | ImageSegment (_, x, _, width, _, _, _, _, _, _, _, _, _) ->
                    widthEnd <- x + width
                | MathSegment (_, _, x, width, _, _, _, _) ->
                    widthEnd <- x + width
                | TextSegment (text, style, x, width, _) ->
                    let visibleLength = visibleTokenLength text
                    if visibleLength = text.Length then
                        widthEnd <- x + width
                    elif visibleLength > 0 then
                        let displayText = WordLayout.Text.displayText style text
                        let displayStyle = WordLayout.Text.displayStyle style
                        widthEnd <- x + measureRange displayText displayStyle 0 visibleLength
                let height = segmentHeight scale segment
                if height > segmentMaxHeight then
                    segmentMaxHeight <- height
            let inlineImageDescent = if hasInlineImage then settings.FontSize * 0.27 else 0.0
            let lineHeight = max settings.LineHeight (segmentMaxHeight + inlineImageDescent)
            let lineWidth = min currentX widthEnd
            let segments = coalescedSegmentList currentSegments
            lines <- {
                Segments = segments
                Y = y
                Width = lineWidth
                Height = lineHeight
                StartOffset = startOffset
                EndOffset = endOffset
            } :: lines
            currentSegments.Clear()
            currentHasTextMetric <- false
            currentFirstFontFamily <- None
            currentFirstFontSize <- 0.0
            currentFirstBold <- false
            currentFirstItalic <- false
            currentFirstVertAlign <- None
            currentFirstSmallCaps <- false
            currentHasMixedTextMetrics <- false
            currentHasFlowContent <- false
            flushedLineCount <- flushedLineCount + 1
            match pendingIntervals with
            | _ :: _ ->
                // Same y has another free strip (e.g. wrapSquare bothSides
                // splits a line into a left and right strip). Fill the next
                // strip before advancing y.
                private_setIntervalToFirst ()
            | [] ->
                y <- y + lineHeight
                moveToWrappableLine ()

    let runs = buffer.Runs |> List.toArray
    let fullText =
        if runs.Length > 1 then ValueSome (TextBuffer.toString buffer)
        else ValueNone
    let runText (textRun: TextBuffer.TextRun) =
        match fullText with
        | ValueSome text ->
            if textRun.Start = 0 && textRun.Length = text.Length then text
            else sliceText text textRun.Start textRun.Length
        | ValueNone ->
            TextBuffer.trySliceNoCopy textRun.Start textRun.Length buffer
            |> Option.defaultWith (fun () -> TextBuffer.slice textRun.Start textRun.Length buffer)

    let followingTextWidth runIndex currentTextStart (currentDisplayText: string) currentDisplayStyle =
        let mutable width = 0.0
        let addTextUntilBreak (displayText: string) displayStyle start =
            if start >= displayText.Length then false
            else
                let mutable stop = start
                while stop < displayText.Length && displayText.[stop] <> '\t' && displayText.[stop] <> '\r' && displayText.[stop] <> '\n' do
                    stop <- stop + 1
                let length = stop - start
                if length > 0 then
                    width <- width + measureRange displayText displayStyle start length
                stop < displayText.Length

        let mutable index = runIndex
        let mutable keepGoing = true
        while keepGoing && index < runs.Length do
            match runs.[index] with
            | TextBuffer.Text nextRun ->
                let rawText = runText nextRun
                let displayText =
                    if index = runIndex then currentDisplayText
                    else WordLayout.Text.displayText nextRun.Style rawText
                let displayStyle =
                    if index = runIndex then currentDisplayStyle
                    else WordLayout.Text.displayStyle nextRun.Style
                let start = if index = runIndex then currentTextStart else 0
                if addTextUntilBreak displayText displayStyle start then keepGoing <- false
            | TextBuffer.Image _ -> keepGoing <- false
            index <- index + 1
        width

    // Next non-bar tab stop strictly to the right of x.
    let nextTabStop x =
        let mutable best: TabStop option = None
        for stop in tabStops do
            if stop.Alignment <> "bar" && stop.Position > x + 0.01 then
                match best with
                | None -> best <- Some stop
                | Some current when stop.Position < current.Position -> best <- Some stop
                | Some _ -> ()
        best

    // Width of the immediate following text up to its decimal separator (for
    // decimal tabs) — that prefix's right edge lands on the tab stop.
    let followingWidthToDecimal currentTextStart (currentDisplayText: string) currentDisplayStyle =
        let mutable index = currentTextStart
        let mutable dotIndex = -1
        while index < currentDisplayText.Length
              && currentDisplayText.[index] <> '\t'
              && currentDisplayText.[index] <> '\r'
              && currentDisplayText.[index] <> '\n' do
            if dotIndex < 0 && currentDisplayText.[index] = '.' then
                dotIndex <- index
            index <- index + 1
        let upto = if dotIndex >= 0 then dotIndex - currentTextStart else index - currentTextStart
        if upto > 0 then measureRange currentDisplayText currentDisplayStyle currentTextStart upto else 0.0

    let tabWidth runIndex currentTextStart (currentDisplayText: string) currentDisplayStyle =
        let defaultTab = (floor (currentX / 48.0) + 1.0) * 48.0
        match nextTabStop currentX with
        | Some stop ->
            let target = min stop.Position currentLineEnd
            let following () = followingTextWidth runIndex currentTextStart currentDisplayText currentDisplayStyle
            match stop.Alignment with
            | "right" -> max 0.0 (target - currentX - following ())
            | "center" -> max 0.0 (target - currentX - following () / 2.0)
            | "decimal" -> max 0.0 (target - currentX - followingWidthToDecimal currentTextStart currentDisplayText currentDisplayStyle)
            | _ -> max 0.0 (target - currentX)   // left (default)
        | None ->
            max 12.0 (min defaultTab currentLineEnd - currentX)

    let addTextSegment text style x width startOffset =
        noteTextMetric style
        currentHasFlowContent <- true
        currentSegments.Add(TextSegment (text, style, x, width, startOffset))

    moveToWrappableLine ()

    for runIndex in 0 .. runs.Length - 1 do
        let run = runs.[runIndex]
        match run with
        | TextBuffer.Image imgRun ->
            let imgWidth, imgHeight = WordLayout.Image.canvasSize imgRun.Width imgRun.Height

            // (A picture in the line goes to the next line when it does not fit; an anchored one
            // takes no room in the line - its anchor broke the line where it stood.)
            if not imgRun.Floating && currentX + imgWidth > currentLineEnd && currentSegments.Count > 0 then
                flushLine ()

            let anchorPage : WordLayout.Floating.AnchorPage =
                { PageWidth = settings.Width; PageHeight = settings.PageHeight; MarginLeft = settings.PaddingLeft; MarginRight = settings.MarginRight
                  MarginTop = settings.PaddingTop; MarginBottom = settings.MarginBottom }
            let rawFloatingBox = WordLayout.Floating.imageCanvasBox anchorPage contentY currentX y settings.PaddingTop imgWidth imgHeight imgRun
            let floatingBox = if imgRun.AllowOverlap = Some false then WordLayout.Floating.avoidOverlap placedBoxes rawFloatingBox else rawFloatingBox
            let imageX = if imgRun.Floating then floatingBox.X else currentX
            let yOffset = if imgRun.Floating then floatingBox.Y - (contentY + (y - settings.PaddingTop)) else 0.0
            let segment = ImageSegment (TextBuffer.displayedImageSource imgRun, imageX, yOffset, imgWidth, imgHeight, imgRun.Start, imgRun.Floating, imgRun.Wrap, imageSegmentHeight imgRun, imgRun.CropLeft, imgRun.CropTop, imgRun.CropRight, imgRun.CropBottom)
            currentSegments.Add segment
            if not imgRun.Floating then
                currentHasFlowContent <- true
            // The same float may have been pre-placed by an anchor-extended pass
            // earlier in the document layout. Detect that by matching the image
            // source against the existing placed boxes; in that case the
            // obstacle is already in `floatObstacles` with an extended (upward)
            // top, so just emit the ImageSegment and keep wrapping as before.
            let alreadyPlaced = imgRun.Floating && isPrePlacedImage imgRun
            if imgRun.Floating && WordLayout.Floating.participatesInWrap floatingBox && not alreadyPlaced then
                let obstacleX, obstacleTop, obstacleWidth, obstacleHeight = WordLayout.Floating.obstacleForBox availableWidth floatingBox
                if WordLayout.Floating.wrapsTopAndBottom floatingBox.Wrap then
                    floatObstacles <- (0.0, obstacleTop, availableWidth, obstacleHeight) :: floatObstacles
                    moveToWrappableLine ()
                else
                    // (A picture with an outline to wrap along: the outline, slab by slab; else its box.)
                    match WordLayout.Floating.outlineObstacles availableWidth floatingBox with
                    | Some outline -> floatObstacles <- outline @ floatObstacles
                    | None -> floatObstacles <- (obstacleX, obstacleTop, obstacleWidth, obstacleHeight) :: floatObstacles
                    pendingIntervals <- wrapIntervalsForY y
                    private_setIntervalToFirst ()
            elif imgRun.Floating && alreadyPlaced then
                // pendingIntervals already reflects the pre-placed obstacle; no
                // change needed beyond emitting the segment.
                ()
            elif imgRun.Floating then
                // Floating but does not participate in wrap (behindDoc=true,
                // wrapNone). The picture is painted at its absolute anchor
                // position; the layout cursor does not reserve any space for
                // it on this line.
                ()
            else
                currentX <- currentX + imgWidth
            if imgRun.Floating && not alreadyPlaced then
                placedBoxes <- floatingBox :: placedBoxes

        | TextBuffer.Text textRun when textRun.Style.Hidden ->
            // w:vanish — hidden text is not laid out and takes no space.
            ()
        | TextBuffer.Text textRun when textRun.Style.EquationXml.IsSome ->
            // OMML equation: lay the math out into an atomic box (like an image).
            let basePx = WordLayout.Text.displayFontSize textRun.Style
            let measure : Eq.Measure = fun s px -> measureRange s { textRun.Style with FontSize = px; EquationXml = None } 0 s.Length
            let boxResult =
                try Ok (Eq.layoutEquation measure basePx textRun.Style.EquationXml.Value)
                with ex -> Error ex.Message
            let fallbackText (txt: string) =
                let dispStyle = WordLayout.Text.displayStyle textRun.Style
                let w = measureRange txt dispStyle 0 txt.Length
                if currentX + w > currentLineEnd && currentSegments.Count > 0 then flushLine ()
                noteTextMetric dispStyle
                currentSegments.Add(TextSegment (txt, dispStyle, currentX, w, textRun.Start))
                currentHasFlowContent <- true
                currentX <- currentX + w
            match boxResult with
            | Ok (Some box) when not box.Prims.IsEmpty ->
                if currentX + box.Width > currentLineEnd && currentSegments.Count > 0 then flushLine ()
                currentSegments.Add(MathSegment (box.Prims, textRun.Style, currentX, box.Width, box.Ascent, box.Descent, textRun.Start, textRun.Length))
                currentHasFlowContent <- true
                currentX <- currentX + box.Width
            | Ok _ | Error _ ->
                // Unparseable equation — fall back to the literal run text.
                fallbackText (WordLayout.Text.displayText textRun.Style (runText textRun))
        | TextBuffer.Text textRun ->
            if textRun.Style.RenderedPageBreakBefore && currentSegments.Count > 0 then
                flushLine ()

            let runText = runText textRun
            let displayRunText = WordLayout.Text.displayText textRun.Style runText
            let displayRunStyle = WordLayout.Text.displayStyle textRun.Style
            let mutable consumed = 0

            let fittingLength start length maxWidth =
                if length <= 1 then length
                else
                    let mutable low = 1
                    let mutable high = length
                    let mutable best = 1
                    while low <= high do
                        let mid = (low + high) / 2
                        let width = measureRange displayRunText displayRunStyle start mid
                        if width <= maxWidth then
                            best <- mid
                            low <- mid + 1
                        else
                            high <- mid - 1
                    best

            iterWordWrapTokenRanges runText (fun tokenStart tokenLength ->
                if tokenLength > 0 then
                    let isLineBreak =
                        match tokenLength with
                        | 1 -> runText.[tokenStart] = '\n' || runText.[tokenStart] = '\r'
                        | 2 -> runText.[tokenStart] = '\r' && runText.[tokenStart + 1] = '\n'
                        | _ -> false
                    if isLineBreak then
                        let hadSegments = currentSegments.Count > 0
                        flushLine ()
                        if not hadSegments then
                            // A break on an otherwise empty line still makes a
                            // line: emit it so the caret, selection and
                            // pagination have something to refer to.
                            let breakOffset = textRun.Start + consumed
                            lines <- {
                                Segments = []
                                Y = y
                                Width = 0.0
                                Height = settings.LineHeight
                                StartOffset = breakOffset
                                EndOffset = breakOffset
                            } :: lines
                            flushedLineCount <- flushedLineCount + 1
                            y <- y + settings.LineHeight
                            moveToWrappableLine ()
                        consumed <- consumed + tokenLength
                    else
                        let mutable remainingStart = 0
                        let mutable remainingLength = tokenLength
                        while remainingLength > 0 do
                            if currentSegments.Count = 0 && currentX >= currentLineEnd then
                                moveToWrappableLine ()

                            let tokenOffset = tokenStart + remainingStart
                            let hasTab = rangeContainsTab runText tokenOffset remainingLength
                            let width =
                                if hasTab then tabWidth runIndex (consumed + remainingLength) displayRunText displayRunStyle
                                else measureRange displayRunText displayRunStyle consumed remainingLength
                            let fitWidth =
                                if hasTab then width
                                else
                                    let visibleLength = visibleTokenLengthRange runText tokenOffset remainingLength
                                    if visibleLength > 0 && visibleLength < remainingLength then
                                        // Word fits the visible token; a trailing space may cross
                                        // the text boundary without forcing the word to the next line.
                                        let visibleWidth = measureRange displayRunText displayRunStyle consumed visibleLength
                                        if currentLineEnd - (currentX + visibleWidth) >= trailingSpaceFitSafety settings visibleLength then visibleWidth else width
                                    else width

                            let fullFitSafety =
                                if settings.InTableCell then
                                    if currentHasMixedTextMetrics then max tableCellLineFitSafety mixedMetricsLineFitSafety else tableCellLineFitSafety
                                elif currentHasMixedTextMetrics then mixedMetricsLineFitSafety
                                else lineFitSafety

                            let overflows = currentX + fitWidth > currentLineEnd - fullFitSafety
                            // For the strip-skip decision use the visible width (no trailing
                            // whitespace), so a short word like "Float" that fits the strip
                            // but whose trailing space pushes it over still gets placed.
                            let visibleOverflows =
                                let visibleLength = visibleTokenLengthRange runText tokenOffset remainingLength
                                if visibleLength > 0 && visibleLength < remainingLength && not hasTab then
                                    let vw = measureRange displayRunText displayRunStyle consumed visibleLength
                                    currentX + vw > currentLineEnd - fullFitSafety
                                else overflows
                            if overflows && currentHasFlowContent then
                                flushLine ()
                            elif visibleOverflows && not currentHasFlowContent && not pendingIntervals.IsEmpty && not hasTab then
                                // Narrow strip at this y can't hold the next whole word, but
                                // there's another strip at the same y — leave this one empty
                                // and try the next, matching Word's wrap-around-float behavior
                                // instead of force-breaking the word inside the narrow strip.
                                private_setIntervalToFirst ()
                            else
                                let segmentLength, segmentWidth =
                                    if overflows && not hasTab then
                                        let maxSegmentWidth = max 1.0 (currentLineEnd - currentX)
                                        let take = fittingLength consumed remainingLength maxSegmentWidth
                                        take, measureRange displayRunText displayRunStyle consumed take
                                    else
                                        remainingLength, width
                                let segmentText =
                                    if tokenOffset = 0 && segmentLength = runText.Length then
                                        runText
                                    else
                                        sliceText runText tokenOffset segmentLength
                                addTextSegment segmentText textRun.Style currentX segmentWidth (textRun.Start + consumed)
                                currentX <- currentX + segmentWidth
                                consumed <- consumed + segmentLength
                                remainingStart <- remainingStart + segmentLength
                                remainingLength <- remainingLength - segmentLength
                                if remainingLength > 0 then flushLine ())

    flushLine ()
    // Offset just after a line break that nothing visible follows (-1: none):
    // the last run that is laid out is text ending in a break.
    let trailingBreakOffset =
        let mutable index = runs.Length - 1
        let mutable result = -1
        let mutable searching = true
        while searching && index >= 0 do
            match runs.[index] with
            | TextBuffer.Text textRun when textRun.Style.Hidden || textRun.Length = 0 -> index <- index - 1
            | TextBuffer.Text textRun ->
                let last = Rope.charAtOrNull (textRun.Start + textRun.Length - 1) buffer.Rope
                if last = '\n' || last = '\r' then result <- textRun.Start + textRun.Length
                searching <- false
            | TextBuffer.Image _ -> searching <- false
        result
    if trailingBreakOffset >= 0 then
        // The paragraph ends in a line break: the break opens a last, empty
        // line, where the caret sits after the break and typing continues.
        lines <- {
            Segments = []
            Y = y
            Width = 0.0
            Height = settings.LineHeight
            StartOffset = trailingBreakOffset
            EndOffset = trailingBreakOffset
        } :: lines
    List.rev lines, floatObstacles, placedBoxes

let layoutLinesStyledWithObstacles (settings: LayoutSettings) (measureText: TextMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) (contentY: float) (tabStops: TabStop list) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    layoutLinesStyledWithObstaclesWithRange settings (rangeMeasurerFromTextMeasurer measureText) buffer maxWidth contentY tabStops initialObstacles initialPlacedBoxes

let layoutLinesStyled (settings: LayoutSettings) (measureText: TextMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) =
    let lines, _, _ = layoutLinesStyledWithObstacles settings measureText buffer maxWidth 0.0 [] [] []
    lines

let private trySingleLineStyledNoObstacles (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) (tabStops: TabStop list) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    if not tabStops.IsEmpty || not initialObstacles.IsEmpty || not initialPlacedBoxes.IsEmpty then
        None
    else
        match buffer.Runs with
        | [TextBuffer.Text textRun] when not textRun.Style.Hidden
                                      && textRun.Style.EquationXml.IsNone
                                      && not textRun.Style.RenderedPageBreakBefore ->
            let rawText =
                TextBuffer.trySliceNoCopy textRun.Start textRun.Length buffer
                |> Option.defaultWith (fun () -> TextBuffer.slice textRun.Start textRun.Length buffer)
            if rawText.Length = 0 || rangeContainsTab rawText 0 rawText.Length || rawText.Contains("\r") || rawText.Contains("\n") then
                None
            else
                let displayText = WordLayout.Text.displayText textRun.Style rawText
                let displayStyle = WordLayout.Text.displayStyle textRun.Style
                let width = measureRange displayText displayStyle 0 displayText.Length
                let availableWidth = max 1.0 maxWidth
                let fitSafety = if settings.InTableCell then tableCellLineFitSafety else lineFitSafety
                if width > availableWidth - fitSafety then
                    None
                else
                    let segment = TextSegment (rawText, textRun.Style, 0.0, width, textRun.Start)
                    let scale = settings.LineHeight / settings.FontSize
                    let lineHeight = max settings.LineHeight (segmentHeight scale segment)
                    let segments = [segment]
                    let lineWidth =
                        let visibleLength = visibleTokenLength rawText
                        if visibleLength = rawText.Length then width
                        elif visibleLength > 0 then measureRange displayText displayStyle 0 visibleLength
                        else 0.0
                    Some [{
                        Segments = segments
                        Y = minBodyTop settings 0.0
                        Width = lineWidth
                        Height = lineHeight
                        StartOffset = textRun.Start
                        EndOffset = textRun.Start + rawText.Length
                    }]
        | _ -> None

let private trySingleLineParagraphHeightNoObstacles (settings: LayoutSettings) (measureText: TextMeasurer) (maxWidth: float) (paragraphStyle: ParagraphStyle) (effectiveTextStyleForParagraph: ParagraphStyle -> TextBuffer.TextStyle -> TextBuffer.TextStyle) (para: ParagraphBlock) =
    let metadata = WordStyles.effectiveMetadata paragraphStyle para.Metadata
    if metadata.Rtl || not metadata.TabStops.IsEmpty then
        None
    else
        let prefixOpt, _ = WordLayout.Numbering.markerAndCounters metadata Map.empty
        if prefixOpt.IsSome then
            None
        else
            let indent = WordLayout.Paragraph.leftIndent metadata
            let rightIndent = WordLayout.Paragraph.rightIndent metadata
            let firstLineIndent = WordLayout.Paragraph.firstLineIndent metadata
            let spacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle metadata
            let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle metadata
            let hasExplicitLineSpacing =
                metadata.LineSpacing.IsSome
                || paragraphStyle.LineSpacing.IsSome
                || metadata.LineSpacingRule.IsSome
                || paragraphStyle.LineSpacingRule.IsSome
            let defaultParagraphSpacing = usesDefaultParagraphSpacing metadata
            let tocEntry = isTocEntryStyle metadata
            let headingDefaultLineBox = headingUsesDefaultLineBox paragraphStyle metadata
            let borderBefore, borderAfter = paragraphBorderVerticalAdvance metadata
            let borderRight = paragraphBorderRightAdvance metadata
            match para.Buffer.Runs with
            | [TextBuffer.Text textRun] when not textRun.Style.Hidden
                                          && textRun.Style.EquationXml.IsNone
                                          && not textRun.Style.RenderedPageBreakBefore ->
                let textStyle = effectiveTextStyleForParagraph paragraphStyle textRun.Style
                let rawText =
                    TextBuffer.trySliceNoCopy textRun.Start textRun.Length para.Buffer
                    |> Option.defaultWith (fun () -> TextBuffer.slice textRun.Start textRun.Length para.Buffer)
                if rawText.Length = 0 || rangeContainsTab rawText 0 rawText.Length || rawText.Contains("\r") || rawText.Contains("\n") then
                    None
                else
                    let availableWidth = maxWidth - indent - rightIndent - max 0.0 firstLineIndent - borderRight
                    let displayText = WordLayout.Text.displayText textStyle rawText
                    let displayStyle = WordLayout.Text.displayStyle textStyle
                    let width = measureText displayText displayStyle
                    let fitSafety = if settings.InTableCell then tableCellLineFitSafety else lineFitSafety
                    if width > max 1.0 availableWidth - fitSafety then
                        None
                    else
                        let contentFontSize = WordLayout.Text.displayFontSize textStyle
                        let contentTextStyle = textStyle
                        let lineHeight = resolvedLineHeightForSettings settings paragraphStyle para.Metadata contentTextStyle contentFontSize hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox
                        let lineScale = lineHeight / contentFontSize
                        let textHeight = contentFontSize * lineScale
                        Some (spacingBefore + borderBefore + max lineHeight textHeight + borderAfter + spacingAfter)
            | _ -> None

// ---------------------------------------------------------------------------
// Paragraph layout memo (the incremental layout).
//
// There is exactly one layout algorithm: the document pass below. What makes a
// keystroke cheap is that laying out one paragraph is a pure function of its
// inputs, so its results can be reused between passes:
//   * the broken lines depend only on the paragraph's text/formatting, the
//     wrap width and the line box (not on where the paragraph sits), and
//   * the finished BlockLayout additionally depends on y (pagination), the
//     list marker and the page geometry.
// A pass that is given a memo therefore returns precisely what a pass without
// one returns; only paragraphs whose inputs changed are recomputed. The memo
// belongs to one IncrementalLayoutState and is passed down the call chain as
// an argument: no layout state is shared between passes, threads or documents.
// Paragraphs that interact with floats (obstacles in or out) bypass the memo.
//
// The memo assumes the text measurer and the document's styles are the same
// between passes: layoutDocumentIncrementalWithRange drops it when the styles
// change, and hosts must clear the state when fonts/measurers change.
// ---------------------------------------------------------------------------

type ParagraphLineBreakKey = {
    LineFontSize: float
    LineHeight: float
    LineInTableCell: bool
    LinePaddingTop: float
    LineTrailingSpaceFitSafety: float
    LineAvailableWidth: float
    LineSingleLineFastPath: bool
}

type ParagraphPlacementKey = {
    PlacementSettings: LayoutSettings
    PlacementMaxWidth: float
    PlacementY: float
    PlacementXOffset: float
    PlacementPrefix: string
}

type ParagraphLayoutMemoEntry = {
    MemoBuffer: TextBuffer
    MemoRawMetadata: BlockMetadata
    MemoParagraphStyle: ParagraphStyle
    MemoMetadata: BlockMetadata
    MemoEffectiveBuffer: TextBuffer
    MemoHasFloatingImage: bool
    MemoContentFontSize: float
    MemoContentTextStyle: TextBuffer.TextStyle
    mutable MemoLines: (ParagraphLineBreakKey * MeasuredLine list) option
    mutable MemoPlacement: (ParagraphPlacementKey * BlockLayout) option
    /// The line breaks and the placement before the current ones: a pass can
    /// lay a paragraph out twice (tried at the foot of a page, then placed on
    /// the next), and with one slot each the two would evict each other at
    /// every pass.
    mutable MemoLinesOlder: (ParagraphLineBreakKey * MeasuredLine list) option
    mutable MemoPlacementOlder: (ParagraphPlacementKey * BlockLayout) option
    /// As MemoPlacement, for a paragraph laid out next to floating objects:
    /// the placement, the obstacles and placed boxes it was given, and the
    /// layout with the obstacles and boxes it handed on.
    mutable MemoObstaclePlacement: (ParagraphPlacementKey * FloatingObstacle list * PlacedFloatingBox list * BlockLayout * FloatingObstacle list * PlacedFloatingBox list) option
    mutable MemoGeneration: int
}

/// The state of a block-by-block pass (the column flow, the single-column
/// document pass) before a block: everything the blocks from there on are
/// laid out from, apart from the blocks themselves and the pass's inputs.
type PassCheckpoint = {
    CheckCol: int
    CheckPage: int
    CheckY: float
    CheckCounters: Map<string * int, int>
    CheckFootnotes: Map<int, (string * string) list>
    CheckPrevious: (ParagraphBlock * ParagraphStyle * BlockMetadata * float) option
    CheckPreviousLayout: BlockLayout option
    CheckObstacles: FloatingObstacle list
    CheckPlacedBoxes: PlacedFloatingBox list
}

/// What a block-by-block pass was given and what it made, kept for the next
/// pass over (nearly) the same blocks: the pass starts again at the first
/// block that differs, from the state it had there, and stops as soon as its
/// state before a block that did not change is what it was the last time -
/// from there on the layouts of the last pass stand, the same objects.
type PassRecord = {
    PassSettings: LayoutSettings
    PassColumns: SectionColumns option
    PassStartPage: int
    PassStartY: float
    PassPageTop: float
    PassNextPageTop: float
    PassNonUniform: bool
    PassBalanced: bool
    PassInitial: PassCheckpoint
    PassStyles: NamedStyle list
    PassPageSetup: PageSetup
    PassPackage: PackageData
    PassBlocks: Block array
    /// The state before each block, and after the last one (length + 1).
    PassCheckpoints: PassCheckpoint array
    /// Where each block's layouts start in PassResults (length + 1).
    PassStarts: int array
    PassResults: BlockLayout array
    /// Whether any of the results refers to an endnote.
    PassHasEndnotes: bool
    /// How many blocks at the start and (of the rest) at the end have the
    /// layouts - the same objects - of the pass this one started from.
    PassKeptPrefix: int
    PassKeptSuffix: int
}

/// The body layouts of the last pass block by block: the layouts of
/// SliceBlocks.[i] are SliceLayouts.[SliceStarts.[i] .. SliceStarts.[i + 1] - 1].
/// The layouts of blocks an edit did not touch or move are the same objects
/// from pass to pass, which is what the painter keeps its commands by.
type LayoutSlices = {
    SliceBlocks: Block array
    SliceStarts: int array
    SliceLayouts: BlockLayout array
    /// Footnote and endnote paragraphs, in the order of the layout list.
    SliceNotes: BlockLayout list
    /// A number of this pass, and of the pass it kept layouts from (0: none).
    /// The first SliceKeptPrefix and the last SliceKeptSuffix blocks have the
    /// layouts, the same objects, they had in that pass.
    SliceGeneration: int
    SliceBase: int
    SliceKeptPrefix: int
    SliceKeptSuffix: int
}

/// Numbers the passes that hand slices over (see LayoutSlices.SliceGeneration).
let mutable private sliceGenerations = 0
let private nextSliceGeneration () =
    sliceGenerations <- sliceGenerations + 1
    sliceGenerations

/// One page's footnote paragraphs and what they were laid out from.
type FootnoteAreaRecord = {
    AreaSettings: LayoutSettings
    AreaMaxWidth: float
    AreaContentBottom: float
    AreaNotes: (string * string) list
    AreaLayouts: BlockLayout list
}

/// The layouts of a section's pass (made with the section's first page at
/// y = 0) moved down to where that page stands in the document.
type PlacedSection = {
    PlacedSource: BlockLayout array
    PlacedOffset: float
    /// How far right of the document's left margin the section's own is.
    PlacedShift: float
    PlacedLayouts: BlockLayout array
}

type ParagraphLayoutMemo = {
    /// Block-by-block passes of the last layout, by the place they have in it.
    MemoPasses: System.Collections.Generic.Dictionary<string, PassRecord>
    MemoFootnoteAreas: System.Collections.Generic.Dictionary<string, FootnoteAreaRecord>
    /// Set by a pass that kept layouts of the last one (see the sweep in layoutDocumentPagesIncremental).
    mutable MemoKeptPass: bool
    /// Blocks the recorded passes of the layout in hand laid out (the others kept their layouts).
    mutable MemoBlocksLaidOut: int
    /// The block slices of the pass in hand, when it has them.
    mutable MemoSlices: LayoutSlices option
    MemoEntries: System.Collections.Generic.Dictionary<string, ParagraphLayoutMemoEntry>
    /// The rest of a paragraph that runs on in the next column or page, by
    /// paragraph id and the offset it starts at: the buffer it was cut from,
    /// where, and the paragraph made of the rest. Handing back the same
    /// paragraph lets its own memo entry stand from one pass to the next.
    MemoContinuations: System.Collections.Generic.Dictionary<string, TextBuffer * int * ParagraphBlock>
    /// The layouts of a section's pass where the section stands in the
    /// document (see sectionLayoutsAt), by the pass's place.
    MemoPlaced: System.Collections.Generic.Dictionary<string, PlacedSection>
    /// A page's footnote paragraphs moved sideways with their section (see
    /// layoutSections), by page: what they were moved from, how far, and the result.
    MemoPlacedNotes: System.Collections.Generic.Dictionary<int, BlockLayout list * float * BlockLayout list>
    mutable Generation: int
    mutable LineBreakCount: int
    mutable MemoStyles: NamedStyle list
    mutable MemoStyleContext: WordStyles.DocumentStyleContext option
}

let private createParagraphLayoutMemo () = {
    MemoPasses = System.Collections.Generic.Dictionary<string, PassRecord>()
    MemoFootnoteAreas = System.Collections.Generic.Dictionary<string, FootnoteAreaRecord>()
    MemoKeptPass = false
    MemoBlocksLaidOut = 0
    MemoSlices = None
    MemoEntries = System.Collections.Generic.Dictionary<string, ParagraphLayoutMemoEntry>()
    MemoContinuations = System.Collections.Generic.Dictionary<string, TextBuffer * int * ParagraphBlock>()
    MemoPlaced = System.Collections.Generic.Dictionary<string, PlacedSection>()
    MemoPlacedNotes = System.Collections.Generic.Dictionary<int, BlockLayout list * float * BlockLayout list>()
    Generation = 0
    LineBreakCount = 0
    MemoStyles = []
    MemoStyleContext = None
}

let private sameOrEqual (left: 'T) (right: 'T) =
    obj.ReferenceEquals(left, right) || left = right

/// Style context for a document, shared between passes of an incremental
/// layout while the style list is unchanged (it depends on nothing else).
let private styleContextForLayout (memo: ParagraphLayoutMemo option) (doc: Document) =
    match memo with
    | Some memo ->
        match memo.MemoStyleContext with
        | Some context when obj.ReferenceEquals(memo.MemoStyles, doc.Styles) -> context
        | _ ->
            let context = WordStyles.styleContext doc
            memo.MemoStyles <- doc.Styles
            memo.MemoStyleContext <- Some context
            context
    | None -> WordStyles.styleContext doc

let private bufferHasFloatingImage (buffer: TextBuffer) =
    let mutable found = false
    for run in buffer.Runs do
        match run with
        | Image image when image.Floating -> found <- true
        | _ -> ()
    found

let private paragraphMemoEntry (memo: ParagraphLayoutMemo) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) =
    let fresh () =
        let metadata = WordStyles.effectiveMetadata paragraphStyle para.Metadata
        let buffer = effectiveBufferForParagraph paragraphStyle para.Buffer
        let contentFontSize, contentTextStyle = paragraphContentMetrics paragraphStyle metadata buffer
        let entry = {
            MemoBuffer = para.Buffer
            MemoRawMetadata = para.Metadata
            MemoParagraphStyle = paragraphStyle
            MemoMetadata = metadata
            MemoEffectiveBuffer = buffer
            MemoHasFloatingImage = bufferHasFloatingImage para.Buffer
            MemoContentFontSize = contentFontSize
            MemoContentTextStyle = contentTextStyle
            MemoLines = None
            MemoPlacement = None
            MemoLinesOlder = None
            MemoPlacementOlder = None
            MemoObstaclePlacement = None
            MemoGeneration = memo.Generation
        }
        entry
    // A paragraph that breaks over a page is laid out as two paragraphs with
    // one id - the whole and the rest - at every pass; with one entry per id
    // each would put the other out. An id has a few slots: the one that holds
    // this paragraph is found by its buffer, a new entry takes a free slot or
    // one that this pass has not used.
    //
    // A slot nothing is in is taken before one whose paragraph the pass has
    // not used: that one is the paragraph as it was before the last change,
    // and undo brings it back - the very buffer, so its lines are found here
    // and not broken again. (Making a 600-page document bold and undoing it
    // broke, and measured, every line twice. What is kept for it is swept
    // with the entries of paragraphs that are gone, see
    // layoutDocumentPagesIncremental.)
    let slotKey (slot: int) = if slot = 0 then para.Id else para.Id + "\u0001" + string slot
    let slots = 3
    let mutable found = None
    let mutable empty = -1
    let mutable unused = -1
    let mutable slot = 0
    while found.IsNone && slot < slots do
        match memo.MemoEntries.TryGetValue (slotKey slot) with
        | true, entry ->
            if obj.ReferenceEquals(entry.MemoBuffer, para.Buffer)
               && sameOrEqual entry.MemoRawMetadata para.Metadata
               && sameOrEqual entry.MemoParagraphStyle paragraphStyle then
                found <- Some entry
            elif unused < 0 && entry.MemoGeneration <> memo.Generation then unused <- slot
        | false, _ -> if empty < 0 then empty <- slot
        slot <- slot + 1
    match found with
    | Some entry ->
        entry.MemoGeneration <- memo.Generation
        entry
    | None ->
        let entry = fresh ()
        memo.MemoEntries.[slotKey (if empty >= 0 then empty elif unused >= 0 then unused else slots - 1)] <- entry
        entry

let private layoutParagraphCore (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) (memoEntry: ParagraphLayoutMemoEntry option) =
    let mutable counters = numberedCounters
    let metadata =
        match memoEntry with
        | Some entry -> entry.MemoMetadata
        | None -> WordStyles.effectiveMetadata paragraphStyle para.Metadata
    let indent = WordLayout.Paragraph.leftIndent metadata
    let rightIndent = WordLayout.Paragraph.rightIndent metadata
    let firstLineIndent = WordLayout.Paragraph.firstLineIndent metadata
    let rawSpacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle metadata
    let spacingBefore = rawSpacingBefore
    let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle metadata
    let hasExplicitLineSpacing =
        metadata.LineSpacing.IsSome
        || paragraphStyle.LineSpacing.IsSome
        || metadata.LineSpacingRule.IsSome
        || paragraphStyle.LineSpacingRule.IsSome
    let defaultParagraphSpacing = usesDefaultParagraphSpacing metadata
    let tocEntry = isTocEntryStyle metadata
    let headingDefaultLineBox = headingUsesDefaultLineBox paragraphStyle metadata
    let borderBefore, borderAfter = paragraphBorderVerticalAdvance metadata
    let borderRight = paragraphBorderRightAdvance metadata
    let contentY = y + spacingBefore + borderBefore
    let buffer =
        match memoEntry with
        | Some entry -> entry.MemoEffectiveBuffer
        | None -> effectiveBufferForParagraph paragraphStyle para.Buffer

    let prefixOpt, updatedCounters = WordLayout.Numbering.markerAndCounters metadata counters
    counters <- updatedCounters
    let prefix = prefixOpt |> Option.defaultValue ""

    let prefixStyle =
        if prefix = "" then None
        else Some (WordLayout.Numbering.markerTextStyle metadata prefix paragraphStyle.Text)

    let prefixWidth =
        WordLayout.Numbering.markerWidth measureText prefix (prefixStyle |> Option.defaultValue paragraphStyle.Text)

    let isHangingList = prefix <> "" && firstLineIndent < 0.0
    let markerReserve = if isHangingList then 0.0 else prefixWidth
    let availableWidth = maxWidth - indent - rightIndent - markerReserve - max 0.0 firstLineIndent - borderRight
    let contentFontSize, contentTextStyle =
        match memoEntry with
        | Some entry -> entry.MemoContentFontSize, entry.MemoContentTextStyle
        | None -> paragraphContentMetrics paragraphStyle metadata buffer
    let paragraphLineHeight = resolvedLineHeightForSettings settings paragraphStyle para.Metadata contentTextStyle contentFontSize hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox
    let struct (lineSpacingMultiple, lineSpacingRule, _) = resolveLineSpacing paragraphStyle para.Metadata contentFontSize
    let lineSettings = { settings with FontSize = contentFontSize; LineHeight = paragraphLineHeight }
    let singleLineFastPath = prefix = "" && not metadata.Rtl
    let breakLines () =
        if singleLineFastPath then
            match trySingleLineStyledNoObstacles lineSettings measureRange buffer availableWidth metadata.TabStops initialObstacles initialPlacedBoxes with
            | Some lines -> lines, initialObstacles, initialPlacedBoxes
            | None -> layoutLinesStyledWithObstaclesWithRange lineSettings measureRange buffer availableWidth contentY metadata.TabStops initialObstacles initialPlacedBoxes
        else
            layoutLinesStyledWithObstaclesWithRange lineSettings measureRange buffer availableWidth contentY metadata.TabStops initialObstacles initialPlacedBoxes
    let lines, updatedObstacles, updatedPlacedBoxes =
        match memoEntry with
        | Some entry ->
            // Only reached without obstacles and without floating images, so
            // line breaking does not depend on contentY (see the memo notes).
            let key = {
                LineFontSize = lineSettings.FontSize
                LineHeight = lineSettings.LineHeight
                LineInTableCell = lineSettings.InTableCell
                LinePaddingTop = lineSettings.PaddingTop
                LineTrailingSpaceFitSafety = lineSettings.LongTrailingSpaceFitSafety
                LineAvailableWidth = availableWidth
                LineSingleLineFastPath = singleLineFastPath
            }
            match entry.MemoLines, entry.MemoLinesOlder with
            | Some (cachedKey, cachedLines), _ when cachedKey = key -> cachedLines, initialObstacles, initialPlacedBoxes
            | current, Some (olderKey, olderLines) when olderKey = key ->
                entry.MemoLinesOlder <- current
                entry.MemoLines <- Some (olderKey, olderLines)
                olderLines, initialObstacles, initialPlacedBoxes
            | current, _ ->
                let lines, obstacles, placedBoxes = breakLines ()
                entry.MemoLinesOlder <- current
                entry.MemoLines <- Some (key, lines)
                lines, obstacles, placedBoxes
        | None -> breakLines ()

    let mutable isFirstLine = true
    let mutable hasRenderedPageBreak = false
    let markerLineBoxFactor = prefixStyle |> Option.bind WordLayout.Numbering.markerLineBoxFactor
    let adjustedLines =
        let adjusted = ResizeArray<MeasuredLine>()
        for line in lines do
            let startsAfterWrapObstacle = lineStartsAfterWrapObstacle line
            let isFirst = isFirstLine
            isFirstLine <- false
            let lineFirstIndent =
                if isFirst then
                    if isHangingList || startsAfterWrapObstacle then 0.0 else firstLineIndent
                else 0.0
            let dx = xOffset + indent + markerReserve + lineFirstIndent
            let adjustedSegments =
                if abs dx < 0.001 then
                    for segment in line.Segments do
                        match segment with
                        | TextSegment (_, style, _, _, _) when style.RenderedPageBreakBefore ->
                            hasRenderedPageBreak <- true
                        | _ -> ()
                    line.Segments
                else
                    let adjustedSegments = ResizeArray<Segment>()
                    for segment in line.Segments do
                        match segment with
                        | TextSegment (text, style, x, width, startOffset) ->
                            if style.RenderedPageBreakBefore then
                                hasRenderedPageBreak <- true
                            adjustedSegments.Add(TextSegment (text, style, x + dx, width, startOffset))
                        | ImageSegment (src, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
                            adjustedSegments.Add(ImageSegment (src, x + dx, yOffset, width, height, startOffset, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom))
                        | MathSegment (prims, style, x, width, ascent, descent, startOffset, length) ->
                            adjustedSegments.Add(MathSegment (prims, style, x + dx, width, ascent, descent, startOffset, length))
                    resizeArrayToList adjustedSegments
            let lineHeight =
                match markerLineBoxFactor with
                | Some markerFactor when line.StartOffset = 0 ->
                    // An exact line height is not grown for the list marker.
                    if lineSpacingRule = "exact" then line.Height
                    else max line.Height (contentFontSize * lineSpacingMultiple * markerFactor)
                | _ -> line.Height
            adjusted.Add
                { line with
                    Y = contentY + (line.Y - settings.PaddingTop)
                    Segments = adjustedSegments
                    Height = lineHeight }
        resizeArrayToList adjusted

    // Justified text: stretch inter-word space on every line except the
    // paragraph's last. Skip entirely when floats are involved (updatedObstacles
    // non-empty ⟹ a float strip narrows some line and our flush-right target
    // wouldn't apply). The common right margin is the first line's reach, so
    // body lines must also stretch past their wrap width by the first-line
    // indent to end flush with it.
    let adjustedLines =
        if metadata.Alignment <> Justify || not (List.isEmpty updatedObstacles) then adjustedLines
        else
            let rightContentEdge = xOffset + indent + markerReserve + availableWidth + max 0.0 firstLineIndent
            let rec justifyNonLast acc = function
                | [] -> List.rev acc
                | [last] -> List.rev (last :: acc)
                | line :: rest -> justifyNonLast (justifyLineToEdge rightContentEdge line :: acc) rest
            justifyNonLast [] adjustedLines

    // RTL paragraphs: reposition each line's tokens right-to-left (guarded by
    // the w:bidi flag, so LTR layout is untouched).
    let adjustedLines =
        if metadata.Rtl then
            let repositioned = ResizeArray<MeasuredLine>()
            for line in adjustedLines do
                repositioned.Add(repositionRtl line)
            resizeArrayToList repositioned
        else adjustedLines

    let pagedLines =
        paginateLines settings ((widowControl paragraphStyle metadata) && not hasRenderedPageBreak) adjustedLines

    let blockHeight =
        if pagedLines.IsEmpty then paragraphLineHeight
        else
            match tryLastLine pagedLines with
            | Some lastLine -> (lastLine.Y - contentY) + lastLine.Height
            | None -> paragraphLineHeight

    ({
        BlockId = para.Id
        Lines = pagedLines
        Y = y
        Height = spacingBefore + borderBefore + blockHeight + borderAfter + spacingAfter
        AvailableWidth = availableWidth
        Metadata = metadata
        Prefix = prefix
        PrefixStyle = prefixStyle
        PrefixWidth = prefixWidth
        Indent = indent + xOffset
        VAlignOffset = 0.0
    }, counters, updatedObstacles, updatedPlacedBoxes)

/// Paragraph layout for a pass that may carry a memo. The memo is passed down
/// the call chain explicitly — there is no shared state, so concurrent
/// layouts (and layouts of different documents) cannot see each other's memo.
let private layoutParagraphWithMemo (memo: ParagraphLayoutMemo option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    let uncached () =
        layoutParagraphCore settings measureText measureRange maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters initialObstacles initialPlacedBoxes None
    match memo with
    | Some memo when not (initialObstacles.IsEmpty && initialPlacedBoxes.IsEmpty) ->
        // Next to floating objects: once one is on a page, every paragraph
        // after it is handed the obstacles. The same paragraph at the same
        // place with the same obstacles is laid out as it was.
        let entry = paragraphMemoEntry memo paragraphStyle effectiveBufferForParagraph para
        if entry.MemoHasFloatingImage then uncached ()
        else
            let prefixOpt, updatedCounters = WordLayout.Numbering.markerAndCounters entry.MemoMetadata numberedCounters
            let key = {
                PlacementSettings = settings
                PlacementMaxWidth = maxWidth
                PlacementY = y
                PlacementXOffset = xOffset
                PlacementPrefix = prefixOpt |> Option.defaultValue ""
            }
            match entry.MemoObstaclePlacement with
            | Some (cachedKey, obstaclesIn, boxesIn, layout, obstaclesOut, boxesOut)
                when cachedKey = key && sameOrEqual obstaclesIn initialObstacles && sameOrEqual boxesIn initialPlacedBoxes ->
                layout, updatedCounters, obstaclesOut, boxesOut
            | _ ->
                let layout, counters, obstacles, placedBoxes = uncached ()
                entry.MemoObstaclePlacement <- Some (key, initialObstacles, initialPlacedBoxes, layout, obstacles, placedBoxes)
                layout, counters, obstacles, placedBoxes
    | Some memo ->
        let entry = paragraphMemoEntry memo paragraphStyle effectiveBufferForParagraph para
        if entry.MemoHasFloatingImage then uncached ()
        else
            let prefixOpt, updatedCounters = WordLayout.Numbering.markerAndCounters entry.MemoMetadata numberedCounters
            let key = {
                PlacementSettings = settings
                PlacementMaxWidth = maxWidth
                PlacementY = y
                PlacementXOffset = xOffset
                PlacementPrefix = prefixOpt |> Option.defaultValue ""
            }
            match entry.MemoPlacement, entry.MemoPlacementOlder with
            | Some (cachedKey, layout), _ when cachedKey = key ->
                layout, updatedCounters, initialObstacles, initialPlacedBoxes
            | current, Some (olderKey, layout) when olderKey = key ->
                entry.MemoPlacementOlder <- current
                entry.MemoPlacement <- Some (olderKey, layout)
                layout, updatedCounters, initialObstacles, initialPlacedBoxes
            | current, _ ->
                let linesBefore = entry.MemoLines
                let linesOlderBefore = entry.MemoLinesOlder
                let layout, counters, obstacles, placedBoxes =
                    layoutParagraphCore settings measureText measureRange maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters initialObstacles initialPlacedBoxes (Some entry)
                // (Lines were broken when a new entry came in, not when the two slots changed places.)
                if not (obj.ReferenceEquals(linesBefore, entry.MemoLines)) && not (obj.ReferenceEquals(linesOlderBefore, entry.MemoLines)) then
                    memo.LineBreakCount <- memo.LineBreakCount + 1
                entry.MemoPlacementOlder <- current
                entry.MemoPlacement <- Some (key, layout)
                layout, counters, obstacles, placedBoxes
    | _ -> uncached ()

let layoutParagraphAtWithStyleAndObstaclesWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    layoutParagraphCore settings measureText measureRange maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters initialObstacles initialPlacedBoxes None

let layoutParagraphAtWithStyleAndObstacles (settings: LayoutSettings) (measureText: TextMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters initialObstacles initialPlacedBoxes

let layoutParagraphAtWithStyle (settings: LayoutSettings) (measureText: TextMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) =
    let layout, counters, _, _ = layoutParagraphAtWithStyleAndObstacles settings measureText maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters [] []
    layout, counters

let layoutParagraphAtWithStyleWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) =
    let layout, counters, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth y xOffset paragraphStyle effectiveBufferForParagraph para numberedCounters [] []
    layout, counters

let layoutParagraphAt (settings: LayoutSettings) (measureText: TextMeasurer) (maxWidth: float) (y: float) (xOffset: float) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) =
    layoutParagraphAtWithStyle settings measureText maxWidth y xOffset normal effectiveBuffer para numberedCounters

let layoutParagraphAtWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) =
    layoutParagraphAtWithStyleWithRange settings measureText measureRange maxWidth y xOffset normal effectiveBuffer para numberedCounters

let private defaultParagraphFirstLineOffsetInDocument memo doc =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInContext (styleContextForLayout memo doc) (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private defaultParagraphFirstLineOffsetInContext context =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInContext context (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private tableLeftOffsetForDocument memo doc (_tableBlock: TableBlock) =
    defaultParagraphFirstLineOffsetInDocument memo doc

let private tableLeftOffsetForContext context (_tableBlock: TableBlock) =
    defaultParagraphFirstLineOffsetInContext context

let private tableAlignmentOffset (alignment: string option) availableWidth tableWidth =
    if optionEqualsIgnoreCase "center" alignment then max 0.0 ((availableWidth - tableWidth) / 2.0)
    elif optionEqualsIgnoreCase "right" alignment then max 0.0 (availableWidth - tableWidth)
    else 0.0

let private paragraphIsVisuallyEmpty (para: ParagraphBlock) =
    TextBuffer.isWhiteSpace para.Buffer

/// A cell's content laid out in document order: paragraphs interleaved with
/// nested tables at the positions recorded by NestedTableAnchors.
type private CellContentItem =
    | CellParagraph of ParagraphBlock
    | CellNestedTable of TableBlock

let private maxNestedTableLayoutDepth = 512

let private orderedCellContent (cell: TableCell) : CellContentItem list =
    if cell.NestedTables.IsEmpty then
        cell.Content |> List.map CellParagraph
    else
        let content = cell.Content |> List.toArray
        let anchors = cell.NestedTableAnchors |> List.toArray
        // Pair each nested table with its anchor (paragraphs preceding it);
        // anchors beyond the available paragraph count fall at the end.
        let anchored =
            cell.NestedTables
            |> List.mapi (fun i table ->
                let anchor =
                    if i < anchors.Length then anchors.[i] else content.Length
                    |> max 0
                    |> min content.Length
                table, anchor)
            |> List.sortBy snd
            |> List.toArray
        let items = ResizeArray<CellContentItem>()
        let mutable tableIndex = 0
        let emitTablesAt paraIndex =
            while tableIndex < anchored.Length && snd anchored.[tableIndex] = paraIndex do
                items.Add(CellNestedTable (fst anchored.[tableIndex]))
                tableIndex <- tableIndex + 1
        for paraIndex in 0 .. content.Length do
            emitTablesAt paraIndex
            if paraIndex < content.Length then
                items.Add(CellParagraph content.[paraIndex])
        // Word requires a paragraph after a nested table; drop trailing empty
        // paragraph marks so a cell that ends with a table is not over-tall.
        let mutable last = items.Count - 1
        let mutable trimming = true
        while trimming && last >= 0 do
            match items.[last] with
            | CellParagraph para when paragraphIsVisuallyEmpty para -> last <- last - 1
            | _ -> trimming <- false
        let mutable result = []
        for i = last downto 0 do
            result <- items.[i] :: result
        result

// ---------------------------------------------------------------------------
// List numbering inside tables.
//
// Rows are measured one at a time and from several places (document layout,
// the table plan, painting), so the numbering state a row starts with cannot
// be threaded from caller to caller. It is derived from the table instead:
// counters start empty at the top of a top-level table and run through its
// cells in reading order (rows, then cells, then paragraphs and nested tables
// in content order). A list that continues from body text into a table (or
// from one table into the next) therefore restarts inside the table.
// ---------------------------------------------------------------------------

type private ListCounters = Map<string * int, int>

type private ListCounterWalk =
    | WalkRow of TableRow * int
    | WalkCell of TableCell * int
    | WalkParagraph of ParagraphBlock

/// Advance the counters through a row in reading order. Iterative (imported
/// tables can be nested hundreds deep) and allocation-free for the usual row
/// without nested tables.
let private advanceListCountersThroughRow (counters: ListCounters) (row: TableRow) : ListCounters =
    let mutable current = counters
    for cell in row.Cells do
        if cell.NestedTables.IsEmpty then
            for para in cell.Content do
                if para.Metadata.ListStyle = Numbered then
                    current <- snd (WordLayout.Numbering.markerAndCounters para.Metadata current)
        else
            let frames = System.Collections.Generic.Stack<ListCounterWalk>()
            frames.Push(WalkCell (cell, 0))
            while frames.Count > 0 do
                match frames.Pop() with
                | WalkParagraph para ->
                    current <- snd (WordLayout.Numbering.markerAndCounters para.Metadata current)
                | WalkRow (walkRow, depth) ->
                    for walkCell in List.rev walkRow.Cells do
                        frames.Push(WalkCell (walkCell, depth))
                | WalkCell (walkCell, depth) ->
                    if walkCell.NestedTables.IsEmpty then
                        for para in List.rev walkCell.Content do
                            if para.Metadata.ListStyle = Numbered then frames.Push(WalkParagraph para)
                    else
                        for item in List.rev (orderedCellContent walkCell) do
                            match item with
                            | CellParagraph para ->
                                if para.Metadata.ListStyle = Numbered then frames.Push(WalkParagraph para)
                            | CellNestedTable nestedTable ->
                                if depth < maxNestedTableLayoutDepth then
                                    for nestedRow in List.rev nestedTable.Rows do
                                        frames.Push(WalkRow (nestedRow, depth + 1))
    current

/// List counters each top-level row of one table starts with. Built once per
/// table by whoever measures its rows and handed to the row measurement; it is
/// a pure function of the table and is not kept anywhere.
type internal TableListNumbering = {
    /// None when the table contains no numbered paragraphs at all.
    BeforeRow: System.Collections.Generic.Dictionary<TableRow, ListCounters> option
}

/// True when list numbering has to be tracked through the table: it holds a
/// numbered paragraph, or a nested table (not looked into here).
let private tableMayHoldNumberedParagraphs (tableBlock: TableBlock) =
    let mutable hasNumbered = false
    for row in tableBlock.Rows do
        if not hasNumbered then
            for cell in row.Cells do
                if not hasNumbered then
                    if not cell.NestedTables.IsEmpty then hasNumbered <- true
                    else
                        for para in cell.Content do
                            if para.Metadata.ListStyle = Numbered then hasNumbered <- true
    hasNumbered

/// The list counters after the table, given those before it.
let private listCountersAfterTable (counters: ListCounters) (tableBlock: TableBlock) : ListCounters =
    if not (tableMayHoldNumberedParagraphs tableBlock) then counters
    else
        let mutable current = counters
        for tableRow in tableBlock.Rows do
            current <- advanceListCountersThroughRow current tableRow
        current

/// `initial` are the counters of the text before the table, so that a list
/// which starts in the body continues inside the table instead of restarting.
let internal tableListNumberingFrom (initial: ListCounters) (tableBlock: TableBlock) : TableListNumbering =
    if not (tableMayHoldNumberedParagraphs tableBlock) then { BeforeRow = None }
    else
        let beforeRow = System.Collections.Generic.Dictionary<TableRow, ListCounters>(HashIdentity.Reference)
        let mutable counters: ListCounters = initial
        for tableRow in tableBlock.Rows do
            if not (beforeRow.ContainsKey tableRow) then
                beforeRow.[tableRow] <- counters
            counters <- advanceListCountersThroughRow counters tableRow
        { BeforeRow = if counters.IsEmpty then None else Some beforeRow }

let internal tableListNumbering (tableBlock: TableBlock) : TableListNumbering =
    tableListNumberingFrom Map.empty tableBlock

/// The list counters in force where a top-level table of `doc` starts: those
/// of the numbered paragraphs and tables before it. For callers that measure a
/// table on its own (plan, paint); the document pass has them at hand.
let internal listCountersBeforeTable (doc: Document) (tableBlock: TableBlock) : ListCounters =
    if not (tableMayHoldNumberedParagraphs tableBlock) then Map.empty
    else
        let mutable counters: ListCounters = Map.empty
        let mutable remaining = doc.Blocks
        let mutable found = false
        while not found && not remaining.IsEmpty do
            match remaining with
            | Table candidate :: _ when obj.ReferenceEquals(candidate, tableBlock) || candidate.Id = tableBlock.Id ->
                found <- true
            | block :: tail ->
                match block with
                | Table earlier -> counters <- listCountersAfterTable counters earlier
                | Paragraph para
                | Heading (_, para) ->
                    if para.Metadata.ListStyle = Numbered then
                        counters <- snd (WordLayout.Numbering.markerAndCounters para.Metadata counters)
                | TextBox _ -> ()
                remaining <- tail
            | [] -> ()
        if found then counters else Map.empty

/// Counters a top-level row starts with; ValueNone when the table has no
/// numbered paragraphs (nothing to track).
let private listCountersBeforeRow (numbering: TableListNumbering) (row: TableRow) : ListCounters voption =
    match numbering.BeforeRow with
    | None -> ValueNone
    | Some beforeRow ->
        match beforeRow.TryGetValue row with
        | true, counters -> ValueSome counters
        | false, _ -> ValueSome Map.empty

/// The vertical merges of one table that cover more than one row. The content
/// of such a cell lives in its w:vMerge="restart" cell but belongs to all the
/// rows it spans: it must not inflate the restart row, and only what does not
/// fit in the spanned rows together makes a row taller (the last one, as in Word).
type internal TableVerticalMerges = {
    MergeRows: TableRow array
    /// Per row (by reference): for each cell, whether it is a restart cell
    /// spanning further rows. Rows without such a cell are absent.
    SpanningCells: System.Collections.Generic.Dictionary<TableRow, bool array>
    RowIndexes: System.Collections.Generic.Dictionary<TableRow, int>
    /// (restart row index, cell index, rows its content may flow over), > 1 only:
    /// the restart row and the empty continuation cells directly below it.
    Spans: struct (int * int * int) array
}

let internal tableVerticalMerges (rows: TableRow list) : TableVerticalMerges option =
    let mutable hasRestart = false
    for row in rows do
        if not hasRestart then
            for cell in row.Cells do
                if cell.VerticalMerge = Some RestartMerge then hasRestart <- true
    if not hasRestart then None
    else
        let rowArray = List.toArray rows
        // Start column -> cell index of the continuation cells of each row.
        let continuations =
            rowArray
            |> Array.map (fun row ->
                let mutable lookup : System.Collections.Generic.Dictionary<int, int> = null
                let mutable column = max 0 row.GridBefore
                let mutable cellIndex = 0
                for cell in row.Cells do
                    // A continuation cell that holds text of its own (content the
                    // importer spread over the merge, see distributeVerticalMergeContent)
                    // is laid out in its own row and ends the stretch the restart
                    // cell's content may flow over.
                    if cell.VerticalMerge = Some ContinueMerge
                       && cell.NestedTables.IsEmpty
                       && cell.Content |> List.forall (fun para -> TextBuffer.isWhiteSpace para.Buffer) then
                        if isNull lookup then lookup <- System.Collections.Generic.Dictionary<int, int>()
                        lookup.[column] <- cellIndex
                    column <- column + WordLayout.Table.cellSpan cell
                    cellIndex <- cellIndex + 1
                lookup)
        let spans = ResizeArray<struct (int * int * int)>()
        let spanningCells = System.Collections.Generic.Dictionary<TableRow, bool array>(HashIdentity.Reference)
        let rowIndexes = System.Collections.Generic.Dictionary<TableRow, int>(HashIdentity.Reference)
        for rowIndex = 0 to rowArray.Length - 1 do
            let row = rowArray.[rowIndex]
            if not (rowIndexes.ContainsKey row) then
                rowIndexes.[row] <- rowIndex
                let mutable column = max 0 row.GridBefore
                let mutable cellIndex = 0
                let mutable flags : bool array = null
                for cell in row.Cells do
                    if cell.VerticalMerge = Some RestartMerge then
                        let mutable span = 1
                        while rowIndex + span < rowArray.Length
                              && not (isNull continuations.[rowIndex + span])
                              && continuations.[rowIndex + span].ContainsKey column do
                            span <- span + 1
                        if span > 1 then
                            if isNull flags then flags <- Array.zeroCreate row.Cells.Length
                            flags.[cellIndex] <- true
                            spans.Add(struct (rowIndex, cellIndex, span))
                    column <- column + WordLayout.Table.cellSpan cell
                    cellIndex <- cellIndex + 1
                if not (isNull flags) then spanningCells.[row] <- flags
        if spans.Count = 0 then None
        else Some { MergeRows = rowArray; SpanningCells = spanningCells; RowIndexes = rowIndexes; Spans = spans.ToArray() }

/// The stretches of a page's text area that a table row (or anything from
/// `top` to `bottom`) covers, page by page: on the pages after the first one
/// below the table's repeated header rows (`continuationHeaderHeight`).
let internal pageStretches (settings: LayoutSettings) (continuationHeaderHeight: float) (top: float) (bottom: float) : (float * float) list =
    let stretches = ResizeArray<float * float>()
    let stride = pageHeight settings + pageGap
    let firstPageTop = WordLayout.Page.pageTopForY (pageHeight settings) top
    let mutable pageTop = firstPageTop
    let mutable guard = 0
    while pageTop < bottom && guard < 10000 do
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop + (if pageTop > firstPageTop then continuationHeaderHeight else 0.0)
        let stretchTop = max top contentTop
        let stretchBottom = min bottom (WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop)
        if stretchBottom > stretchTop + 0.01 then stretches.Add((stretchTop, stretchBottom))
        pageTop <- pageTop + stride
        guard <- guard + 1
    if stretches.Count = 0 then [ (top, bottom) ] else resizeArrayToList stretches

/// The layouts of a vertically merged cell laid over the stretches its rows
/// are on. `layouts` are the cell's paragraphs laid out in one piece from the
/// top of the first stretch (a row-local measurement moved there); a line
/// that does not fit in what is left of a stretch goes to the top of the
/// next one, as the rows it belongs to did, and the lines after it follow.
/// Returns the layouts and the bottom of the content (the last stretch is
/// not a limit: what is left over makes the last row taller).
let internal layoutsOverStretches (stretches: (float * float) array) (paddingTop: float) (paddingBottom: float) (layouts: BlockLayout list) : BlockLayout list * float =
    let mutable stretch = 0
    let mutable shift = 0.0
    let mutable contentBottom = if stretches.Length > 0 then fst stretches.[0] else 0.0
    let moved =
        layouts
        |> List.map (fun layout ->
            if layout.Lines.IsEmpty then
                let placed = { layout with Y = layout.Y + shift }
                contentBottom <- max contentBottom (placed.Y + placed.Height)
                placed
            else
                let mutable first = true
                let mutable layoutShift = shift
                let mutable lastBottom = layout.Y
                let lines =
                    layout.Lines
                    |> List.map (fun line ->
                        // (The last stretch takes whatever is left.)
                        while stretch < stretches.Length - 1
                              && line.Y + shift + line.Height > snd stretches.[stretch] - paddingBottom + pageFitTolerance do
                            stretch <- stretch + 1
                            shift <- fst stretches.[stretch] + paddingTop - line.Y
                        if first then
                            layoutShift <- shift
                            first <- false
                        lastBottom <- line.Y + shift + line.Height
                        if shift = 0.0 then line else { line with Y = line.Y + shift })
                let originalLastBottom = (List.last layout.Lines).Y + (List.last layout.Lines).Height
                let trailing = max 0.0 (layout.Y + layout.Height - originalLastBottom)
                let y = layout.Y + layoutShift
                contentBottom <- max contentBottom (lastBottom + trailing)
                { layout with Y = y; Height = lastBottom + trailing - y; Lines = lines })
    moved, contentBottom + paddingBottom

/// Extra height per row so that every merged cell fits in the rows it spans.
/// `measurements` are the rows measured with the merged cells left out of their
/// restart row's height. A deficit goes to the last row of the span; spans are
/// settled in the order they end, so nested spans see what inner ones added.
let private verticalMergeExtraHeights (merges: TableVerticalMerges) (measurements: TableRowMeasurement array) : float array =
    let extra = Array.zeroCreate<float> measurements.Length
    let groupArrays = Array.zeroCreate<TableCellLayoutGroup array> measurements.Length
    let ordered = merges.Spans |> Array.sortBy (fun struct (rowIndex, _, span) -> rowIndex + span)
    for struct (rowIndex, cellIndex, span) in ordered do
        if rowIndex < measurements.Length then
            if isNull groupArrays.[rowIndex] then
                groupArrays.[rowIndex] <- List.toArray measurements.[rowIndex].CellGroups
            let groups = groupArrays.[rowIndex]
            if cellIndex < groups.Length then
                let lastRow = min (measurements.Length - 1) (rowIndex + span - 1)
                let mutable available = 0.0
                for index = rowIndex to lastRow do
                    available <- available + measurements.[index].Height + extra.[index]
                let deficit = groups.[cellIndex].RequiredHeight - available
                if deficit > 0.0 then
                    extra.[lastRow] <- extra.[lastRow] + deficit
    extra

let private noSpanningCells : bool array = [||]

/// Rows of a table with vertical merges, measured once row-locally (merged
/// cells left out), plus the extra heights. Filled in by the first row
/// measurement that needs it: it takes the measurer, which the scope is built
/// without. It lives and dies with the scope of one pass over one table.
type internal TableMergeMeasurements = {
    Merges: TableVerticalMerges
    mutable RowLocal: TableRowMeasurement array
    mutable Extra: float array
}

/// What a row measurement needs besides geometry: the memo of the pass it
/// belongs to (if any), its table's list numbering and its vertical merges.
type internal TableMeasureScope = {
    ScopeMemo: ParagraphLayoutMemo option
    ScopeNumbering: TableListNumbering
    ScopeMerges: TableMergeMeasurements option
}

let internal tableMeasureScopeFrom (memo: ParagraphLayoutMemo option) (initialCounters: Map<string * int, int>) (tableBlock: TableBlock) : TableMeasureScope =
    { ScopeMemo = memo
      ScopeNumbering = tableListNumberingFrom initialCounters tableBlock
      ScopeMerges = tableVerticalMerges tableBlock.Rows |> Option.map (fun merges -> { Merges = merges; RowLocal = null; Extra = null }) }

let rec private measureTableRowWithStyleAtDepth nestedTableLayoutDepth (memo: ParagraphLayoutMemo option) (rowListCounters: ListCounters voption) (spanningCells: bool array) (mergeExtraHeight: float) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    let trackListCounters = rowListCounters.IsSome
    let mutable listCounters: ListCounters = match rowListCounters with ValueSome counters -> counters | ValueNone -> Map.empty
    let cellSettings = {
        settings with
            InTableCell = true
            MarginBottom = if row.CanSplit = Some false then -1000000.0 else settings.MarginBottom
            TableContinuationHeaderHeight = tableContinuationHeaderHeight
    }
    let numColumns = columnMetrics.Count
    let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
    let rowSpacingAdvance = WordLayout.Table.rowSpacingAdvance cellSpacing
    let outerStartInset = outerLeft + cellSpacing
    let mutable rowHeight = WordLayout.Table.rowMinHeight WordLayout.canvasScale settings.LineHeight row
    let mutable cellGroups = []
    let mutable traces = []

    let columnWidthPrefix = columnMetrics.Prefix
    let fallbackWidth = columnMetrics.FallbackWidth
    let mutable currentColumn = max 0 row.GridBefore
    let mutable cellIndex = 0
    for cell in row.Cells do
        let startColumn = currentColumn
        let span = WordLayout.Table.cellSpan cell
        let usesInheritedCellAfterSpacing =
            cell.Content
            |> List.exists (fun para -> para.Metadata.SpacingAfter.IsNone)
        let colWidth = spacedSpannedWidthFromPrefix columnWidthPrefix cellSpacing startColumn span fallbackWidth
        let cellPadding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row cell
        let textPadding =
            if cellSpacing > 0.0 then
                // Word still honors horizontal tcMar overrides in spaced cells, but vertical
                // text origin follows the row/table defaults.
                let rowPadding = WordLayout.Table.rowCellPaddingDefaults WordLayout.canvasScale tableBlock row
                { cellPadding with Top = rowPadding.Top; Bottom = rowPadding.Bottom }
            else
                cellPadding
        let verticalMarginExcess =
            if cellSpacing > 0.0 then WordLayout.Table.spacedVerticalMarginExcess textPadding cellPadding else 0.0
        let collapsedBorderExtent =
            if row.Height.IsNone && usesInheritedCellAfterSpacing then
                WordLayout.Table.collapsedHorizontalBorderExtent tableBlock row
            else
                0.0
        let verticalBorderExtent = max (WordLayout.Table.cellVerticalBorderExtent cell) collapsedBorderExtent
        let cellX = tableLeftOffset + spacedStartOffsetFromPrefix columnWidthPrefix cellSpacing startColumn + textPadding.Left - outerStartInset - WordLayout.Table.internalSpacingTextOffset cellSpacing startColumn
        let cellMaxWidth = WordLayout.Table.cellContentWidth colWidth textPadding
        let cellY = rowY + cellSpacing + WordLayout.Table.spacedCellTextTopOffset cellSpacing textPadding.Top
        let contentStartY = WordLayout.Table.cellContentStartY cellY textPadding
        let mutable paraY = contentStartY
        let cellLayouts = ResizeArray<BlockLayout>()

        let layoutNestedTable (nestedTable: TableBlock) nestedTableY =
            if nestedTableLayoutDepth >= maxNestedTableLayoutDepth then
                [], WordLayout.Table.emptyTableAdvance cellSettings.LineHeight
            else
                let nestedOuterLeft, nestedOuterRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale nestedTable
                let nestedTableMaxWidth = cellMaxWidth + nestedOuterLeft + nestedOuterRight
                let nestedColumnWidths = tableColumnWidths nestedTableMaxWidth nestedTable
                if List.isEmpty nestedColumnWidths then
                    [], WordLayout.Table.emptyTableAdvance cellSettings.LineHeight
                else
                    let nestedColumnMetrics = tableColumnMetrics nestedTableMaxWidth nestedColumnWidths
                    let nestedTableWidth = WordLayout.Table.tableWidthWithSpacing nestedColumnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale nestedTable)
                    let nestedHorizontalOffset = cellX + tableAlignmentOffset nestedTable.Alignment cellMaxWidth nestedTableWidth
                    let mutable nestedY = nestedTableY
                    let nestedLayouts = ResizeArray<BlockLayout>()
                    // Vertical merges of the nested table: measure its rows once
                    // with the merged cells left out to learn the extra heights.
                    let nestedMerges = tableVerticalMerges nestedTable.Rows
                    let nestedSpanning (nestedRow: TableRow) =
                        match nestedMerges with
                        | Some merges ->
                            match merges.SpanningCells.TryGetValue nestedRow with
                            | true, flags -> flags
                            | _ -> noSpanningCells
                        | None -> noSpanningCells
                    let nestedExtra =
                        match nestedMerges with
                        | Some merges ->
                            let nestedRowCounters = if trackListCounters then ValueSome listCounters else ValueNone
                            merges.MergeRows
                            |> Array.map (fun nestedRow ->
                                measureTableRowWithStyleAtDepth (nestedTableLayoutDepth + 1) memo nestedRowCounters (nestedSpanning nestedRow) 0.0 cellSettings measureText measureRange nestedTable nestedColumnMetrics nestedOuterLeft nestedHorizontalOffset nestedRow nestedTableY 0.0 paragraphStyleForParagraph effectiveBufferForParagraph)
                            |> verticalMergeExtraHeights merges
                        | None -> [||]
                    let mutable nestedRowIndex = 0

                    for nestedRow in nestedTable.Rows do
                        let nestedRowCounters = if trackListCounters then ValueSome listCounters else ValueNone
                        if trackListCounters then
                            listCounters <- advanceListCountersThroughRow listCounters nestedRow
                        let nestedRowExtra = if nestedRowIndex < nestedExtra.Length then nestedExtra.[nestedRowIndex] else 0.0
                        nestedRowIndex <- nestedRowIndex + 1
                        let measurement =
                            measureTableRowWithStyleAtDepth (nestedTableLayoutDepth + 1) memo nestedRowCounters (nestedSpanning nestedRow) nestedRowExtra cellSettings measureText measureRange nestedTable nestedColumnMetrics nestedOuterLeft nestedHorizontalOffset nestedRow nestedY 0.0 paragraphStyleForParagraph effectiveBufferForParagraph
                        for group in measurement.CellGroups do
                            let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment measurement.Height group.PaddingTop group.PaddingBottom group.ContentHeight
                            let adjustedLayouts =
                                group.Layouts
                                |> List.map (fun layout ->
                                    { layout with
                                        Y = layout.Y + offset
                                        VAlignOffset = layout.VAlignOffset + offset
                                        Lines = shiftLinesY offset layout.Lines })
                            nestedLayouts.AddRange adjustedLayouts
                        nestedY <- nestedY + measurement.Height

                    resizeArrayToList nestedLayouts, (nestedY - nestedTableY) + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale nestedTable

        let mutable pendingEmptyHeight = 0.0
        let layoutCellParagraph rawPara =
            let para = tableCellParagraphForLayout rawPara
            let paragraphStyle = paragraphStyleForParagraph para
            let layout, updatedListCounters, _, _ = layoutParagraphWithMemo memo cellSettings measureText measureRange cellMaxWidth paraY cellX paragraphStyle effectiveBufferForParagraph para listCounters [] []
            listCounters <- updatedListCounters
            let layout = if layout.Lines.IsEmpty then moveLayoutToY paraY layout else layout
            let layout =
                match layout.Lines with
                | firstLine :: _ when pendingEmptyHeight > 0.0 && pageNumberForY cellSettings layout.Y <> pageNumberForY cellSettings firstLine.Y ->
                    { layout with
                        Height = layout.Height + pendingEmptyHeight
                        Lines = shiftLinesY pendingEmptyHeight layout.Lines }
                | _ -> layout
            cellLayouts.Add layout
            pendingEmptyHeight <- if layout.Lines.IsEmpty then pendingEmptyHeight + layout.Height else 0.0
            paraY <- paraY + layout.Height

        if cell.NestedTables.IsEmpty then
            for rawPara in cell.Content do
                layoutCellParagraph rawPara
        else
            for item in orderedCellContent cell do
                match item with
                | CellNestedTable nestedTable ->
                    let nestedLayouts, nestedHeight = layoutNestedTable nestedTable paraY
                    cellLayouts.AddRange nestedLayouts
                    paraY <- paraY + nestedHeight
                    pendingEmptyHeight <- 0.0
                | CellParagraph rawPara ->
                    layoutCellParagraph rawPara

        let contentHeight = max 0.0 (paraY - contentStartY)
        let requiredHeight = contentHeight + textPadding.Top + textPadding.Bottom + verticalMarginExcess + verticalBorderExtent + rowSpacingAdvance
        // A merged cell's content belongs to all the rows it spans; what does
        // not fit in them arrives as mergeExtraHeight on the span's last row.
        let spansRows = cellIndex < spanningCells.Length && spanningCells.[cellIndex]
        if not spansRows then
            rowHeight <- max rowHeight requiredHeight
        let cellLayoutsList = resizeArrayToList cellLayouts
        cellGroups <- { VerticalAlignment = cell.VerticalAlignment; PaddingTop = textPadding.Top; PaddingBottom = textPadding.Bottom; ContentHeight = contentHeight; RequiredHeight = requiredHeight; Layouts = cellLayoutsList } :: cellGroups
        cellIndex <- cellIndex + 1
        let mutable totalLines = 0
        for layout in cellLayoutsList do
            totalLines <- totalLines + layout.Lines.Length
        if traceTimingEnabled then
            let cellSnippet =
                cell.Content
                |> List.map (fun para -> TextBuffer.toString para.Buffer)
                |> List.tryFind hasNonWhitespace
                |> Option.map snippet
                |> Option.defaultValue ""
            traces <- sprintf "c%d/%d w=%.2f cw=%.2f h=%.2f lines=%d text=%s" startColumn span colWidth cellMaxWidth contentHeight totalLines cellSnippet :: traces
        currentColumn <- currentColumn + span

    { Height = rowHeight + mergeExtraHeight; CellGroups = List.rev cellGroups; Traces = List.rev traces }

let internal tableMeasureScope (memo: ParagraphLayoutMemo option) (tableBlock: TableBlock) : TableMeasureScope =
    tableMeasureScopeFrom memo Map.empty tableBlock

/// The one way a top-level table row is measured: every caller (the table
/// plan, the block flow, and the two paint paths) goes through here with the
/// table's scope, so they agree on row heights, vertical merges included.
let internal measureTableRowWithStyleUsingMetrics (scope: TableMeasureScope) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    match scope.ScopeMerges with
    | None ->
        measureTableRowWithStyleAtDepth 0 scope.ScopeMemo (listCountersBeforeRow scope.ScopeNumbering row) noSpanningCells 0.0 settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph
    | Some state ->
        let merges = state.Merges
        let spanningFor (tableRow: TableRow) =
            match merges.SpanningCells.TryGetValue tableRow with
            | true, flags -> flags
            | _ -> noSpanningCells
        let rowLocalSettings = {
            settings with
                PageHeight = unpagedMeasurementPageHeight
                MarginBottom = 0.0
                TableContinuationHeaderHeight = 0.0
        }
        if isNull state.RowLocal then
            let rowLocal =
                merges.MergeRows
                |> Array.map (fun tableRow ->
                    measureTableRowWithStyleAtDepth 0 scope.ScopeMemo (listCountersBeforeRow scope.ScopeNumbering tableRow) (spanningFor tableRow) 0.0 rowLocalSettings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset tableRow 0.0 0.0 paragraphStyleForParagraph effectiveBufferForParagraph)
            state.Extra <- verticalMergeExtraHeights merges rowLocal
            state.RowLocal <- rowLocal
        match merges.RowIndexes.TryGetValue row with
        | true, rowIndex ->
            let extra = state.Extra.[rowIndex]
            let isRowLocal =
                rowY = 0.0 && tableContinuationHeaderHeight = 0.0
                && settings.PageHeight = unpagedMeasurementPageHeight && settings.MarginBottom = 0.0 && settings.TableContinuationHeaderHeight = 0.0
            if isRowLocal then
                // Exactly what the pre-pass measured; no need to lay the row out again.
                let measured = state.RowLocal.[rowIndex]
                if extra = 0.0 then measured else { measured with Height = measured.Height + extra }
            else
                measureTableRowWithStyleAtDepth 0 scope.ScopeMemo (listCountersBeforeRow scope.ScopeNumbering row) (spanningFor row) extra settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph
        | _ ->
            // Not a row of this table (a header row copy, say): measured on its own.
            measureTableRowWithStyleAtDepth 0 scope.ScopeMemo (listCountersBeforeRow scope.ScopeNumbering row) noSpanningCells 0.0 settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph

let measureTableRowWithStyle (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (tableMaxWidth: float) (columnWidths: float list) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowWithStyleUsingMetrics (tableMeasureScope None tableBlock) settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph

let internal tryMeasureTableRowHeightWithMetrics (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (row: TableRow) (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveTextStyleForParagraph: ParagraphStyle -> TextBuffer.TextStyle -> TextBuffer.TextStyle) =
    let cellSettings = {
        settings with
            InTableCell = true
            MarginBottom = if row.CanSplit = Some false then -1000000.0 else settings.MarginBottom
            TableContinuationHeaderHeight = 0.0
    }
    let numColumns = columnMetrics.Count
    // Rows taking part in a vertical merge depend on the other rows of the
    // span, which only the scoped measurement knows about.
    if numColumns = 0 || row.Cells |> List.exists (fun cell -> cell.VerticalMerge.IsSome) then
        None
    else
        let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
        let rowSpacingAdvance = WordLayout.Table.rowSpacingAdvance cellSpacing
        let mutable rowHeight = WordLayout.Table.rowMinHeight WordLayout.canvasScale settings.LineHeight row
        let mutable failed = false
        let columnWidthPrefix = columnMetrics.Prefix
        let fallbackWidth = columnMetrics.FallbackWidth
        let mutable currentColumn = max 0 row.GridBefore
        for cell in row.Cells do
            if not failed then
                let span = WordLayout.Table.cellSpan cell
                if not cell.NestedTables.IsEmpty then
                    failed <- true
                else
                    let usesInheritedCellAfterSpacing =
                        cell.Content
                        |> List.exists (fun para -> para.Metadata.SpacingAfter.IsNone)
                    let colWidth = spacedSpannedWidthFromPrefix columnWidthPrefix cellSpacing currentColumn span fallbackWidth
                    let cellPadding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row cell
                    let textPadding =
                        if cellSpacing > 0.0 then
                            let rowPadding = WordLayout.Table.rowCellPaddingDefaults WordLayout.canvasScale tableBlock row
                            { cellPadding with Top = rowPadding.Top; Bottom = rowPadding.Bottom }
                        else
                            cellPadding
                    let verticalMarginExcess =
                        if cellSpacing > 0.0 then WordLayout.Table.spacedVerticalMarginExcess textPadding cellPadding else 0.0
                    let collapsedBorderExtent =
                        if row.Height.IsNone && usesInheritedCellAfterSpacing then
                            WordLayout.Table.collapsedHorizontalBorderExtent tableBlock row
                        else
                            0.0
                    let verticalBorderExtent = max (WordLayout.Table.cellVerticalBorderExtent cell) collapsedBorderExtent
                    let cellMaxWidth = WordLayout.Table.cellContentWidth colWidth textPadding
                    let mutable contentHeight = 0.0
                    for rawPara in cell.Content do
                        if not failed then
                            let para = tableCellParagraphForLayout rawPara
                            let paragraphStyle = paragraphStyleForParagraph para
                            match trySingleLineParagraphHeightNoObstacles cellSettings measureText cellMaxWidth paragraphStyle effectiveTextStyleForParagraph para with
                            | Some height -> contentHeight <- contentHeight + height
                            | None -> failed <- true
                    if not failed then
                        rowHeight <- max rowHeight (contentHeight + textPadding.Top + textPadding.Bottom + verticalMarginExcess + verticalBorderExtent + rowSpacingAdvance)
                currentColumn <- currentColumn + span
        if failed then None else Some rowHeight

let tryMeasureTableRowHeightWithStyle (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (tableMaxWidth: float) (columnWidths: float list) (row: TableRow) (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveTextStyleForParagraph: ParagraphStyle -> TextBuffer.TextStyle -> TextBuffer.TextStyle) =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    tryMeasureTableRowHeightWithMetrics settings measureText tableBlock columnMetrics row paragraphStyleForParagraph effectiveTextStyleForParagraph

let internal measureTableRowForContextWithColumnMetrics (scope: TableMeasureScope) settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    measureTableRowWithStyleUsingMetrics scope settings measureText measureRange tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight (forTableCellParagraphInContext context) (effectiveBufferInContext context)

let measureTableRowForContext settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowForContextWithColumnMetrics (tableMeasureScope None tableBlock) settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

let measureTableRowForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    let context = WordStyles.styleContext doc
    measureTableRowForContext settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

let internal measureTableRowForContextRowLocalWithColumnMetrics (scope: TableMeasureScope) settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row =
    let rowLocalSettings = {
        settings with
            PageHeight = unpagedMeasurementPageHeight
            MarginBottom = 0.0
            TableContinuationHeaderHeight = 0.0
    }
    measureTableRowForContextWithColumnMetrics scope rowLocalSettings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row 0.0 0.0

let measureTableRowForContextRowLocal settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowForContextRowLocalWithColumnMetrics (tableMeasureScope None tableBlock) settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row

let measureTableRowForDocumentRowLocal settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row =
    let context = WordStyles.styleContext doc
    measureTableRowForContextRowLocal settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row

let moveTableRowMeasurementToY rowY (measurement: TableRowMeasurement) =
    { measurement with
        CellGroups =
            measurement.CellGroups
            |> List.map (fun group ->
                { group with
                    Layouts = group.Layouts |> List.map (fun layout -> moveLayoutToY (layout.Y + rowY) layout) }) }

let tableRowFragmentLayoutsWithHeight rowHeight (fragment: TableRowFragmentPlan) (measurement: TableRowMeasurement) =
    let fragmentBottom = fragment.Y + fragment.Height
    let contentOriginDelta = fragment.RowContentY - fragment.RowPlanY
    let fragmentBlockIds = System.Collections.Generic.HashSet<BlockId>(fragment.BlockIds)
    let result = ResizeArray<BlockLayout>()
    for group in measurement.CellGroups do
        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment rowHeight group.PaddingTop group.PaddingBottom group.ContentHeight
        let lineOffset = offset + contentOriginDelta
        for layout in group.Layouts do
            if fragmentBlockIds.Contains layout.BlockId then
                let visibleLines = ResizeArray<MeasuredLine>()
                let mutable y = System.Double.PositiveInfinity
                let mutable bottom = System.Double.NegativeInfinity
                for line in layout.Lines do
                    let shiftedY = line.Y + lineOffset
                    if shiftedY >= fragment.Y && shiftedY + line.Height <= fragmentBottom then
                        visibleLines.Add { line with Y = shiftedY }
                        if shiftedY < y then y <- shiftedY
                        let lineBottom = shiftedY + line.Height
                        if lineBottom > bottom then bottom <- lineBottom
                if visibleLines.Count > 0 then
                    result.Add { layout with Lines = resizeArrayToList visibleLines; Y = y; Height = bottom - y }
    resizeArrayToList result

let tableRowFragmentLayouts (fragment: TableRowFragmentPlan) (measurement: TableRowMeasurement) =
    tableRowFragmentLayoutsWithHeight measurement.Height fragment measurement

let tableCellFragmentPlans (plan: TableLayoutPlan) =
    let tableBlock = plan.Table
    let numColumns = WordLayout.Table.columnCount tableBlock
    let resolvedRows = plan.ResolvedRows.Value
    let rowHeights = tablePlanRowHeights plan
    let rowHeightPrefix = rowHeightPrefix rowHeights
    let rowSpanLookup = resolvedCellRowSpanLookup resolvedRows
    let result = ResizeArray<TableCellFragmentPlan>()
    for planRow in plan.Rows do
        let row = planRow.Row
        let positionedCells = WordLayout.Table.positionedCells row |> List.toArray
        let cellGroups = planRow.Measurement.CellGroups |> List.toArray
        let resolvedCells =
            resolvedRows
            |> Array.tryItem planRow.RowIndex
            |> Option.map (fun row -> row.Cells |> List.toArray)
            |> Option.defaultValue [||]
        let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
        let fragmentCount = planRow.Fragments.Length
        let mutable fragmentIndex = -1
        for fragment in planRow.Fragments do
            fragmentIndex <- fragmentIndex + 1
            let fragmentEnd = fragment.Y + fragment.Height
            let contentOriginDelta = fragment.RowContentY - fragment.RowPlanY
            let fragmentBlockIds = System.Collections.Generic.HashSet<BlockId>(fragment.BlockIds)
            for resolvedCell in resolvedCells do
                match positionedCells |> Array.tryItem resolvedCell.SourceCellIndex, cellGroups |> Array.tryItem resolvedCell.SourceCellIndex with
                | Some positioned, Some group ->
                    let fallbackWidth = plan.TableMaxWidth / float numColumns
                    let cellX =
                        plan.TableX
                        + WordLayout.Table.spacedStartOffset plan.ColumnWidths cellSpacing resolvedCell.StartColumn
                        - WordLayout.Table.spacedCellPaintXOffset cellSpacing resolvedCell.StartColumn
                    let cellWidth =
                        WordLayout.Table.spacedSpannedWidth plan.ColumnWidths cellSpacing resolvedCell.StartColumn resolvedCell.ColSpan fallbackWidth
                        - WordLayout.Table.spacedCellPaintWidthAdjustment cellSpacing numColumns resolvedCell.StartColumn resolvedCell.ColSpan
                        |> max 0.0
                    let cellPadding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row positioned.Cell
                    let textPadding =
                        if cellSpacing > 0.0 then
                            let rowPadding = WordLayout.Table.rowCellPaddingDefaults WordLayout.canvasScale tableBlock row
                            { cellPadding with Top = rowPadding.Top; Bottom = rowPadding.Bottom }
                        else
                            cellPadding
                    let paintTopAdjustment = WordLayout.Table.spacedCellPaintTopAdjustment cellSpacing textPadding cellPadding
                    let cellFragmentHeight = tableCellVerticalAlignmentHeightFromLookup rowSpanLookup rowHeightPrefix planRow.RowIndex resolvedCell.SourceCellIndex fragment.Height
                    let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment cellFragmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                    let visibleLines = ResizeArray<MeasuredLine>()
                    let blockIds = ResizeArray<BlockId>()
                    let seenBlockIds = System.Collections.Generic.HashSet<BlockId>()
                    let mutable visibleYValue = System.Double.PositiveInfinity
                    let mutable visibleEndValue = System.Double.NegativeInfinity
                    for layout in group.Layouts do
                        if fragmentBlockIds.Contains layout.BlockId then
                            if seenBlockIds.Add layout.BlockId then
                                blockIds.Add layout.BlockId
                            for line in layout.Lines do
                                let line = { line with Y = line.Y + offset + contentOriginDelta }
                                if line.Y >= fragment.Y && line.Y + line.Height <= fragmentEnd then
                                    visibleLines.Add line
                                    if line.Y < visibleYValue then visibleYValue <- line.Y
                                    let lineEnd = line.Y + line.Height
                                    if lineEnd > visibleEndValue then visibleEndValue <- lineEnd
                    let visibleY = if visibleLines.Count = 0 then None else Some visibleYValue
                    let visibleEnd = if visibleLines.Count = 0 then None else Some visibleEndValue
                    let visibleHeight =
                        match visibleY, visibleEnd with
                        | Some startY, Some endY -> Some (endY - startY)
                        | _ -> None
                    result.Add({
                        RowIndex = planRow.RowIndex
                        CellIndex = resolvedCell.SourceCellIndex
                        StartColumn = resolvedCell.StartColumn
                        ColumnSpan = resolvedCell.ColSpan
                        RowSpan = resolvedCell.RowSpan
                        BorderTop = resolvedCell.Borders.Top
                        BorderRight = resolvedCell.Borders.Right
                        BorderBottom = resolvedCell.Borders.Bottom
                        BorderLeft = resolvedCell.Borders.Left
                        Page = fragment.Page
                        RowY = planRow.Y
                        RowHeight = planRow.Height
                        FragmentY = fragment.Y
                        FragmentHeight = fragment.Height
                        CellX = cellX
                        CellWidth = cellWidth
                        CellY = fragment.Y + WordLayout.Table.spacedCellPaintTopOffset cellSpacing - paintTopAdjustment
                        CellHeight = WordLayout.Table.spacedCellPaintHeight (tableCellPaintHeightInFragment rowSpanLookup rowHeightPrefix planRow.RowIndex resolvedCell.SourceCellIndex fragmentCount fragmentIndex fragment.Height) cellSpacing + paintTopAdjustment
                        CellSpacing = cellSpacing
                        CellBackgroundColor = positioned.Cell.BackgroundColor
                        CellShading = positioned.Cell.Shading
                        PaddingTop = group.PaddingTop
                        PaddingBottom = group.PaddingBottom
                        ContentY = fragment.RowContentY + cellSpacing + group.PaddingTop + offset
                        ContentHeight = group.ContentHeight
                        VisibleY = visibleY
                        VisibleHeight = visibleHeight
                        VisibleLineCount = visibleLines.Count
                        BlockIds = resizeArrayToList blockIds
                    })
                | _ -> ()
    resizeArrayToList result

let tableCellPaintFragmentPlans (plan: TableLayoutPlan) =
    let tableBlock = plan.Table
    let numColumns = WordLayout.Table.columnCount tableBlock
    let resolvedRows = plan.ResolvedRows.Value
    let rowHeights = tablePlanRowHeights plan
    let rowHeightPrefix = rowHeightPrefix rowHeights
    let rowSpanLookup = resolvedCellRowSpanLookup resolvedRows
    let result = ResizeArray<TableCellFragmentPlan>()
    let columnWidthPrefix = columnWidthPrefix plan.ColumnWidths
    let fallbackWidth = if numColumns = 0 then plan.TableMaxWidth else plan.TableMaxWidth / float numColumns
    for planRow in plan.Rows do
        let row = planRow.Row
        let sourceCells = row.Cells |> List.toArray
        let cellGroups = planRow.Measurement.CellGroups |> List.toArray
        let resolvedCells =
            resolvedRows
            |> Array.tryItem planRow.RowIndex
            |> Option.map (fun row -> row.Cells |> List.toArray)
            |> Option.defaultValue [||]
        let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
        let fragmentCount = planRow.Fragments.Length
        let mutable fragmentIndex = -1
        for fragment in planRow.Fragments do
            fragmentIndex <- fragmentIndex + 1
            for resolvedCell in resolvedCells do
                match sourceCells |> Array.tryItem resolvedCell.SourceCellIndex, cellGroups |> Array.tryItem resolvedCell.SourceCellIndex with
                | Some cell, Some group ->
                    let cellX =
                        plan.TableX
                        + spacedStartOffsetFromPrefix columnWidthPrefix cellSpacing resolvedCell.StartColumn
                        - WordLayout.Table.spacedCellPaintXOffset cellSpacing resolvedCell.StartColumn
                    let cellWidth =
                        spacedSpannedWidthFromPrefix columnWidthPrefix cellSpacing resolvedCell.StartColumn resolvedCell.ColSpan fallbackWidth
                        - WordLayout.Table.spacedCellPaintWidthAdjustment cellSpacing numColumns resolvedCell.StartColumn resolvedCell.ColSpan
                        |> max 0.0
                    let cellPadding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row cell
                    let textPadding =
                        if cellSpacing > 0.0 then
                            let rowPadding = WordLayout.Table.rowCellPaddingDefaults WordLayout.canvasScale tableBlock row
                            { cellPadding with Top = rowPadding.Top; Bottom = rowPadding.Bottom }
                        else
                            cellPadding
                    let paintTopAdjustment = WordLayout.Table.spacedCellPaintTopAdjustment cellSpacing textPadding cellPadding
                    let cellFragmentHeight = tableCellVerticalAlignmentHeightFromLookup rowSpanLookup rowHeightPrefix planRow.RowIndex resolvedCell.SourceCellIndex fragment.Height
                    let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment cellFragmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                    result.Add({
                        RowIndex = planRow.RowIndex
                        CellIndex = resolvedCell.SourceCellIndex
                        StartColumn = resolvedCell.StartColumn
                        ColumnSpan = resolvedCell.ColSpan
                        RowSpan = resolvedCell.RowSpan
                        BorderTop = resolvedCell.Borders.Top
                        BorderRight = resolvedCell.Borders.Right
                        BorderBottom = resolvedCell.Borders.Bottom
                        BorderLeft = resolvedCell.Borders.Left
                        Page = fragment.Page
                        RowY = planRow.Y
                        RowHeight = planRow.Height
                        FragmentY = fragment.Y
                        FragmentHeight = fragment.Height
                        CellX = cellX
                        CellWidth = cellWidth
                        CellY = fragment.Y + WordLayout.Table.spacedCellPaintTopOffset cellSpacing - paintTopAdjustment
                        CellHeight = WordLayout.Table.spacedCellPaintHeight (tableCellPaintHeightInFragment rowSpanLookup rowHeightPrefix planRow.RowIndex resolvedCell.SourceCellIndex fragmentCount fragmentIndex fragment.Height) cellSpacing + paintTopAdjustment
                        CellSpacing = cellSpacing
                        CellBackgroundColor = cell.BackgroundColor
                        CellShading = cell.Shading
                        PaddingTop = group.PaddingTop
                        PaddingBottom = group.PaddingBottom
                        ContentY = fragment.RowContentY + cellSpacing + group.PaddingTop + offset
                        ContentHeight = group.ContentHeight
                        VisibleY = None
                        VisibleHeight = None
                        VisibleLineCount = 0
                        BlockIds = []
                    })
                | _ -> ()
    resizeArrayToList result

let private borderGridStroke priority (stroke: WordLayout.Table.PrintStroke option) =
    stroke
    |> Option.map (fun stroke -> ({
        Color = stroke.Color
        Width = stroke.Width
        Style = stroke.Style
        Priority = priority
    }: RedBat.BorderGrid.Stroke))

let tableCellFragmentBorderCells (cellFragments: TableCellFragmentPlan list) =
    let result = ResizeArray<RedBat.BorderGrid.Cell>(cellFragments.Length)
    for fragment in cellFragments do
        result.Add({
            Row = fragment.RowIndex
            Column = fragment.StartColumn
            RowSpan = fragment.RowSpan
            ColumnSpan = fragment.ColumnSpan
            X = fragment.CellX
            Y = fragment.CellY
            Width = fragment.CellWidth
            Height = fragment.CellHeight
            Edges = {
                Top = borderGridStroke 1 fragment.BorderTop
                Right = borderGridStroke 1 fragment.BorderRight
                Bottom = borderGridStroke 1 fragment.BorderBottom
                Left = borderGridStroke 1 fragment.BorderLeft
            }
        }: RedBat.BorderGrid.Cell)
    resizeArrayToList result

/// Vertically merged cells of a table whose rows are being placed (by the
/// layout pass and by the table's plan alike). The content of such a cell is
/// laid out with its first row, in one piece; where the rows it spans come
/// to be on more than one page - a page break between them, or a row that is
/// split - it is laid over the stretches those rows are on, once the last of
/// them has its place.
type internal MergedCellSplits = {
    SplitSpans: struct (int * int * int) array
    SplitRowStretches: ResizeArray<(float * float) list>
    /// (first row, cell) -> the cell laid out in one piece from its row's top.
    SplitWholes: System.Collections.Generic.Dictionary<struct (int * int), TableCellLayoutGroup>
}

let internal mergedCellSplitsFor (scope: TableMeasureScope) : MergedCellSplits option =
    match scope.ScopeMerges with
    | Some merges when merges.Merges.Spans.Length > 0 ->
        // (A merged cell that holds a table stays in one piece: the table's rows and borders
        // are placed with the cell's first row.)
        let spans =
            merges.Merges.Spans
            |> Array.filter (fun struct (rowIndex, cellIndex, _) ->
                match List.tryItem cellIndex merges.Merges.MergeRows.[rowIndex].Cells with
                | Some cell -> cell.NestedTables.IsEmpty
                | None -> false)
        Some { SplitSpans = spans; SplitRowStretches = ResizeArray<(float * float) list>()
               SplitWholes = System.Collections.Generic.Dictionary<struct (int * int), TableCellLayoutGroup>() }
    | _ -> None

/// Row `rowIndex` has its place (`rowY`, `rowHeight`; rows are told in
/// order): the merged cells that end with it and are split, as (first row,
/// cell, the cell's layouts where they belong now), and what the row has to
/// grow by for the content the last stretch does not hold.
let internal mergedCellsSettledAt (splits: MergedCellSplits) (settings: LayoutSettings) (continuationHeaderHeight: float) (rowIndex: int) (rowY: float) (rowHeight: float) (rowLocal: TableRowMeasurement) : (int * int * BlockLayout list) list * float =
    splits.SplitRowStretches.Add (pageStretches settings continuationHeaderHeight rowY (rowY + rowHeight))
    let mutable settled = []
    let mutable growth = 0.0
    for struct (firstRow, cellIndex, span) in splits.SplitSpans do
        if firstRow = rowIndex then
            match List.tryItem cellIndex (moveTableRowMeasurementToY rowY rowLocal).CellGroups with
            | Some whole -> splits.SplitWholes.[struct (firstRow, cellIndex)] <- whole
            | None -> ()
        if firstRow + span - 1 = rowIndex && firstRow < rowIndex then
            match splits.SplitWholes.TryGetValue(struct (firstRow, cellIndex)) with
            | true, whole ->
                // The stretches of the rows, those that touch as one.
                let stretches = ResizeArray<float * float>()
                for spanned in firstRow .. rowIndex do
                    for (top, bottom) in splits.SplitRowStretches.[spanned] do
                        if stretches.Count > 0 && abs (snd stretches.[stretches.Count - 1] - top) < 0.5 then
                            stretches.[stretches.Count - 1] <- (fst stretches.[stretches.Count - 1], bottom)
                        else stretches.Add((top, bottom))
                if stretches.Count > 1 then
                    let moved, contentBottom = layoutsOverStretches (stretches.ToArray()) whole.PaddingTop whole.PaddingBottom whole.Layouts
                    settled <- (firstRow, cellIndex, moved) :: settled
                    growth <- max growth (contentBottom - snd stretches.[stretches.Count - 1])
            | _ -> ()
    settled, (if growth > 0.5 then growth else 0.0)

/// "Keep with next" on a table row (on the first paragraph of its first
/// cell): the row is not the last of its page when the row after it starts
/// the next one - it goes there with it, and so do the rows before it that
/// are kept with it. Returns where the row goes: `pagedRowY`, or the top of
/// the next page.
let internal keepRowWithNextY (settings: LayoutSettings) (headerHeight: float) (contentBottom: float) (rows: TableRow array) (rowKeepsWithNext: TableRow -> bool) (measureRowLocal: TableRow -> TableRowMeasurement) (rowIndex: int) (rowY: float) (rowHeight: float) (pagedRowY: float) : float =
    if pagedRowY = rowY && rowIndex + 1 < rows.Length
       && contentBottom - (rowY + rowHeight) < 400.0
       && rowKeepsWithNext rows.[rowIndex] then
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) rowY
        let pageContentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        if rowY <= pageContentTop + headerHeight + pageFitTolerance then pagedRowY
        else
            let mutable total = rowHeight
            let mutable next = rowIndex + 1
            let mutable more = true
            while more && next < rows.Length do
                let nextRow = rows.[next]
                let nextMeasurement = measureRowLocal nextRow
                if rowKeepsWithNext nextRow && next + 1 < rows.Length then
                    total <- total + nextMeasurement.Height
                    next <- next + 1
                else
                    // (Of a row that may break, its first lines are enough.)
                    let needed =
                        if nextRow.CanSplit = Some false then nextMeasurement.Height
                        else
                            let mutable firstLines = 0.0
                            for group in nextMeasurement.CellGroups do
                                match group.Layouts |> List.tryPick (fun layout -> List.tryHead layout.Lines) with
                                | Some line -> firstLines <- max firstLines (line.Y + line.Height)
                                | None -> ()
                            if firstLines <= 0.0 then nextMeasurement.Height else min nextMeasurement.Height firstLines
                    total <- total + needed
                    more <- false
            let pageRoom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop - pageContentTop - headerHeight
            if rowY + total > contentBottom + pageFitTolerance && total <= pageRoom then advanceToNextPage settings rowY
            else pagedRowY
    else pagedRowY

let internal tableRowKeepsWithNext (styleContext: DocumentStyleContext) (row: TableRow) =
    match row.Cells with
    | cell :: _ ->
        match cell.Content with
        | para :: _ -> WordStyles.keepNext (forBlockInContext styleContext (Paragraph para)) para.Metadata
        | [] -> false
    | [] -> false

let planTableRowFragments (settings: LayoutSettings) firstTablePage repeatedHeaderHeight rowY rowHeight (measurement: TableRowMeasurement) =
    let rowBottom = rowY + rowHeight
    let mutable pageTop = WordLayout.Page.pageTopForY (pageHeight settings) rowY
    let pageContentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
    let pageContentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
    if rowY >= pageContentTop && rowBottom <= pageContentBottom then
        let lineCount, contentTop, contentBottom, blockIds = tableRowMeasurementLineSummary measurement
        let page = pageNumberForY settings rowY
        [{
            Page = page
            RowPlanY = rowY
            RowContentY = rowY
            Y = rowY
            Height = rowHeight
            AvailableHeight = pageContentBottom - rowY
            RepeatedHeaderHeight = if page > firstTablePage then repeatedHeaderHeight else 0.0
            ContentTop = contentTop
            ContentBottom = contentBottom
            LineCount = lineCount
            ConsumedLineStart = 0
            ConsumedLineEnd = lineCount
            BlockIds = blockIds
            BreakReason = "row-span"
        }]
    else
        let lines = sortedTableRowBlockLines measurement
        let fragments = ResizeArray<TableRowFragmentPlan>()
        let mutable consumedLineStart = 0
        while pageTop < rowBottom do
            let fragmentTop = max rowY (WordLayout.Page.contentTop settings.PaddingTop pageTop)
            let fragmentBottom = min rowBottom (WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop)
            if fragmentBottom > fragmentTop then
                let blockIds = ResizeArray<BlockId>()
                let seenBlockIds = System.Collections.Generic.HashSet<BlockId>()
                let mutable lineCount = 0
                let mutable contentTopValue = System.Double.PositiveInfinity
                let mutable contentBottomValue = System.Double.NegativeInfinity
                for blockId, line in lines do
                    if line.Y < fragmentBottom && line.Y + line.Height > fragmentTop then
                        lineCount <- lineCount + 1
                        if line.Y < contentTopValue then contentTopValue <- line.Y
                        let lineBottom = line.Y + line.Height
                        if lineBottom > contentBottomValue then contentBottomValue <- lineBottom
                        if seenBlockIds.Add blockId then
                            blockIds.Add blockId
                let consumedLineEnd = consumedLineStart + lineCount
                let contentTop = if lineCount = 0 then None else Some contentTopValue
                let contentBottom = if lineCount = 0 then None else Some contentBottomValue
                let page = pageNumberForY settings fragmentTop
                let repeatedHeaderHeight = if page > firstTablePage then repeatedHeaderHeight else 0.0
                fragments.Add {
                    Page = page
                    RowPlanY = rowY
                    RowContentY = rowY
                    Y = fragmentTop
                    Height = fragmentBottom - fragmentTop
                    AvailableHeight = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop - fragmentTop
                    RepeatedHeaderHeight = repeatedHeaderHeight
                    ContentTop = contentTop
                    ContentBottom = contentBottom
                    LineCount = lineCount
                    ConsumedLineStart = consumedLineStart
                    ConsumedLineEnd = consumedLineEnd
                    BlockIds = resizeArrayToList blockIds
                    BreakReason = "row-span"
                }
                consumedLineStart <- consumedLineEnd
            pageTop <- pageTop + pageHeight settings + pageGap
        resizeArrayToList fragments

let private contentBottomForY settings y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop

let private nextPageContentY settings y repeatedHeaderHeight =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + repeatedHeaderHeight

let private fragmentPlanForRowTransaction settings firstTablePage headerHeight rowY rowHeight (row: TableRow) (rowLocalMeasurement: TableRowMeasurement) =
    let rowLines = sortedTableRowBlockLines rowLocalMeasurement
    let fragments = ResizeArray<TableRowFragmentPlan>()
    let mutable currentY = rowY
    let mutable consumedLineStart = 0

    let repeatedHeaderHeightAt y =
        if not row.IsHeader && pageNumberForY settings y > firstTablePage then headerHeight else 0.0

    let addFragment fragmentY fragmentHeight repeatedHeaderHeight breakReason =
        let fragmentBottom = fragmentY + fragmentHeight
        let mutable consumedLineEnd = consumedLineStart
        while consumedLineEnd < rowLines.Length && currentY + (snd rowLines.[consumedLineEnd]).Y >= fragmentY && currentY + (snd rowLines.[consumedLineEnd]).Y + (snd rowLines.[consumedLineEnd]).Height <= fragmentBottom do
            consumedLineEnd <- consumedLineEnd + 1
        let visibleLineCount = consumedLineEnd - consumedLineStart
        let blockIds = ResizeArray<BlockId>()
        let seenBlockIds = System.Collections.Generic.HashSet<BlockId>()
        let mutable contentTopValue = System.Double.PositiveInfinity
        let mutable contentBottomValue = System.Double.NegativeInfinity
        let mutable lineIndex = consumedLineStart
        while lineIndex < consumedLineEnd do
            let blockId, line = rowLines.[lineIndex]
            let lineTop = currentY + line.Y
            let lineBottom = lineTop + line.Height
            if lineTop < contentTopValue then contentTopValue <- lineTop
            if lineBottom > contentBottomValue then contentBottomValue <- lineBottom
            if seenBlockIds.Add blockId then
                blockIds.Add blockId
            lineIndex <- lineIndex + 1
        let contentTop = if visibleLineCount = 0 then None else Some contentTopValue
        let contentBottom = if visibleLineCount = 0 then None else Some contentBottomValue
        fragments.Add {
            Page = pageNumberForY settings fragmentY
            RowPlanY = rowY
            RowContentY = currentY
            Y = fragmentY
            Height = fragmentHeight
            AvailableHeight = contentBottomForY settings fragmentY - fragmentY
            RepeatedHeaderHeight = repeatedHeaderHeight
            ContentTop = contentTop
            ContentBottom = contentBottom
            LineCount = visibleLineCount
            ConsumedLineStart = consumedLineStart
            ConsumedLineEnd = consumedLineEnd
            BlockIds = resizeArrayToList blockIds
            BreakReason = breakReason
        }
        consumedLineStart <- consumedLineEnd
        visibleLineCount

    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) currentY
    let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
    let contentBottom = contentBottomForY settings currentY
    let remainingHeight = max 0.0 (contentBottom - currentY)
    let repeatedHeaderHeight = repeatedHeaderHeightAt currentY
    if rowLines.Length > 0 && currentY > contentTop && remainingHeight < rowHeight then
        let visible = addFragment currentY remainingHeight repeatedHeaderHeight "no-full-line-fits"
        if visible = 0 then
            currentY <- nextPageContentY settings currentY (repeatedHeaderHeightAt (currentY + pageHeight settings + pageGap))

    let repeatedHeaderHeight = repeatedHeaderHeightAt currentY
    let reason = if consumedLineStart > 0 then "fits-after-advance" else if rowLines.Length = 0 then "empty-fragment" else "fits"
    addFragment currentY rowHeight repeatedHeaderHeight reason |> ignore
    currentY, resizeArrayToList fragments

// ---- table plans kept between passes -----------------------------------------
// The plan of a table - its rows measured, placed and split over pages - is
// what both the layout pass and the paint pass work from, and each built it
// again for every table of the document at every edit: two thirds of the cost
// of a keystroke in a long document with tables. A plan depends on the table
// (the same object while it is not edited), where it starts, the settings,
// the styles, the list counters before it, and the text measurers. The last
// are functions and cannot be compared, so plans are kept per *scope*: a
// renderer opens one for its layout state (`withTablePlanScope`), and gets a
// new one when its measurers change. Outside a scope nothing is kept.

type private KnownTablePlan = {
    Scope: obj
    Y: float
    Settings: LayoutSettings
    Styles: NamedStyle list
    Counters: Map<string * int, int> voption
    MaxWidth: float
    ColumnWidths: float list
    OuterLeft: float
    HorizontalOffset: float
    Plan: TableLayoutPlan
}

#if FABLE_COMPILER
let mutable private tablePlanScopeValue : obj = null
let private getTablePlanScope () : obj = tablePlanScopeValue
let private setTablePlanScope (scope: obj) = tablePlanScopeValue <- scope
#else
// (Per thread: native hosts lay documents out on several threads.)
type private TablePlanScopeHolder() =
    [<System.ThreadStatic; DefaultValue>]
    static val mutable private scope : obj
    static member Scope
        with get () : obj = TablePlanScopeHolder.scope
        and set (value: obj) = TablePlanScopeHolder.scope <- value
let private getTablePlanScope () : obj = TablePlanScopeHolder.Scope
let private setTablePlanScope (scope: obj) = TablePlanScopeHolder.Scope <- scope
#endif

/// Run `run` with table plans kept for `scope` (an object that stands for one
/// set of text measurers): a table that is the same object at the same place
/// as in the last pass of this scope is not measured again.
let withTablePlanScope (scope: obj) (run: unit -> 'T) : 'T =
    let previous = getTablePlanScope ()
    setTablePlanScope scope
    try run ()
    finally setTablePlanScope previous

/// The scope table plans are kept for at the moment (null: none).
let currentTablePlanScope () : obj = getTablePlanScope ()

#if FABLE_COMPILER
[<Fable.Core.Emit("new WeakMap()")>]
let private newTablePlanMap () : obj = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.get($1)")>]
let private tablePlanMapGet (map: obj) (key: TableBlock) : KnownTablePlan = Fable.Core.Util.jsNative
[<Fable.Core.Emit("$0.set($1, $2)")>]
let private tablePlanMapSet (map: obj) (key: TableBlock) (value: KnownTablePlan) : unit = Fable.Core.Util.jsNative
let private knownTablePlans = newTablePlanMap ()
let private tryKnownTablePlan (table: TableBlock) : KnownTablePlan option =
    let entry = tablePlanMapGet knownTablePlans table
    if isNull (box entry) then None else Some entry
let private rememberTablePlan (table: TableBlock) (entry: KnownTablePlan) = tablePlanMapSet knownTablePlans table entry
#else
let private knownTablePlans = System.Runtime.CompilerServices.ConditionalWeakTable<TableBlock, KnownTablePlan>()
let private tryKnownTablePlan (table: TableBlock) : KnownTablePlan option =
    match knownTablePlans.TryGetValue table with
    | true, entry -> Some entry
    | false, _ -> None
let private rememberTablePlan (table: TableBlock) (entry: KnownTablePlan) =
    lock knownTablePlans (fun () ->
        knownTablePlans.Remove table |> ignore
        knownTablePlans.Add(table, entry))
#endif

let rec private buildTableLayoutPlanWithMemo (memo: ParagraphLayoutMemo option) (initialCounters: Map<string * int, int> voption) settings measureText measureRange (doc: Document) (tableBlock: TableBlock) (tableMaxWidth: float) (columnWidths: float list) (outerLeft: float) (tableHorizontalOffset: float) (tableY: float) : TableLayoutPlan =
    let scope = getTablePlanScope ()
    let known =
        if isNull scope then None
        else
            match tryKnownTablePlan tableBlock with
            | Some entry when obj.ReferenceEquals(entry.Scope, scope)
                              && entry.Y = tableY
                              && obj.ReferenceEquals(entry.Styles, doc.Styles)
                              && entry.MaxWidth = tableMaxWidth
                              && entry.OuterLeft = outerLeft
                              && entry.HorizontalOffset = tableHorizontalOffset
                              && (obj.ReferenceEquals(entry.Settings, settings) || entry.Settings = settings)
                              && (obj.ReferenceEquals(entry.ColumnWidths, columnWidths) || entry.ColumnWidths = columnWidths)
                              // (A pass that does not state the counters - the paint pass - takes
                              // the plan of the layout pass that came before it.)
                              && (match initialCounters, entry.Counters with
                                  | ValueNone, _ -> true
                                  | ValueSome wanted, ValueSome had -> obj.ReferenceEquals(wanted, had) || wanted = had
                                  | ValueSome _, ValueNone -> false) ->
                Some entry.Plan
            | _ -> None
    match known with
    | Some plan -> plan
    | None ->
        let plan = buildTableLayoutPlanUncached memo initialCounters settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY
        if not (isNull scope) then
            rememberTablePlan tableBlock
                { Scope = scope; Y = tableY; Settings = settings; Styles = doc.Styles; Counters = initialCounters; MaxWidth = tableMaxWidth
                  ColumnWidths = columnWidths; OuterLeft = outerLeft; HorizontalOffset = tableHorizontalOffset; Plan = plan }
        plan

and private buildTableLayoutPlanUncached (memo: ParagraphLayoutMemo option) (initialCounters: Map<string * int, int> voption) settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY =
    let styleContext = styleContextForLayout memo doc
    // The list counters at the table: handed in by the document pass, else
    // worked out from the document, so the plan numbers cells as the pass does.
    let initialCounters =
        match initialCounters with
        | ValueSome counters -> counters
        | ValueNone -> listCountersBeforeTable doc tableBlock
    let measureScope = tableMeasureScopeFrom memo initialCounters tableBlock
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    let outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock |> snd
    let tableX = settings.PaddingLeft + tableHorizontalOffset - outerLeft
    let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
    let resolvedRows =
        lazy (
            WordLayout.Table.resolveTable WordLayout.canvasScale (Some tableMaxWidth) tableBlock
            |> fun resolved -> resolved.BodyRows |> List.toArray)
    let headerRows = WordLayout.Table.headerRows tableBlock
    let measureRow row rowY tableContinuationHeaderHeight =
        measureTableRowForContextWithColumnMetrics measureScope settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight
    let measureRowLocal row =
        measureTableRowForContextRowLocalWithColumnMetrics measureScope settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row
    let headerMeasurementsBuffer = ResizeArray()
    for row in headerRows do
        headerMeasurementsBuffer.Add(measureRowLocal row)
    let headerMeasurementsLocal = headerMeasurementsBuffer.ToArray()
    let headerRowCount = headerMeasurementsLocal.Length
    let headerHeight =
        let mutable totalHeight = 0.0
        for header in headerMeasurementsLocal do
            totalHeight <- totalHeight + header.Height
        totalHeight
    let mutable rowY = tableY
    let firstTablePage = pageNumberForY settings tableY
    let planRowArray = List.toArray tableBlock.Rows
    let mergedSplits = mergedCellSplitsFor measureScope
    let rows = ResizeArray<TableLayoutPlanRow>()
    let mutable rowIndex = 0
    for row in tableBlock.Rows do
        let continuationHeaderHeight = if row.IsHeader then 0.0 else headerHeight
        let rowLocalMeasurement = measureRowLocal row
        let mutable plannedMeasurement = moveTableRowMeasurementToY rowY rowLocalMeasurement
        let mutable plannedY = rowY
        let mutable rowHeight = rowLocalMeasurement.Height
        let pagedRowY = paginateTableRow settings row rowY rowHeight
        let contentBottom = contentBottomForY settings rowY
        let pagedRowY = keepRowWithNextY settings headerHeight contentBottom planRowArray (tableRowKeepsWithNext styleContext) measureRowLocal rowIndex rowY rowHeight pagedRowY
        let pagedRowY =
            if pagedRowY = rowY
               && (row.CanSplit |> Option.defaultValue true)
               && rowY + rowHeight > contentBottom
               && splitTableRowNeedsNextPage settings rowY contentBottom rowLocalMeasurement then
                advanceToNextPage settings rowY
            else
                pagedRowY
        let rowContentY =
            if pagedRowY <> rowY && not row.IsHeader && not headerRows.IsEmpty then
                pagedRowY + headerHeight
            else
                pagedRowY
        let rowFitsAtCurrentY = rowContentY = rowY && rowY + rowHeight <= contentBottom
        let rowFitsAfterMove = rowContentY <> rowY
        if rowFitsAtCurrentY || rowFitsAfterMove then
            plannedY <- rowContentY
            plannedMeasurement <- moveTableRowMeasurementToY rowContentY rowLocalMeasurement
        else
            plannedMeasurement <- measureRow row rowY continuationHeaderHeight
            rowHeight <- plannedMeasurement.Height
        let repeatedHeaderHeight = if row.IsHeader then 0.0 else headerHeight
        // (A merged cell that is split with its rows can make the last of them taller, as in the layout pass.)
        match mergedSplits with
        | Some splits ->
            let _, growth = mergedCellsSettledAt splits settings continuationHeaderHeight rowIndex plannedY rowHeight rowLocalMeasurement
            rowHeight <- rowHeight + growth
        | None -> ()
        let fragments = planTableRowFragments settings (pageNumberForY settings plannedY) repeatedHeaderHeight plannedY rowHeight plannedMeasurement
        rows.Add {
            RowIndex = rowIndex
            Row = row
            IsHeader = row.IsHeader
            CanSplit = row.CanSplit |> Option.defaultValue true
            Y = plannedY
            Height = rowHeight
            Measurement = plannedMeasurement
            Fragments = fragments
            HeadersAbove = pagedRowY <> rowY && not row.IsHeader && not headerRows.IsEmpty && plannedY = pagedRowY + headerHeight
        }
        rowY <- plannedY + rowHeight
        rowIndex <- rowIndex + 1
    {
        Table = tableBlock
        TableX = tableX
        TableMaxWidth = tableMaxWidth
        TableWidth = tableWidth
        OuterLeft = outerLeft
        OuterRight = outerRight
        ColumnWidths = columnWidths
        ResolvedRows = resolvedRows
        HeaderRowCount = headerRowCount
        Rows = resizeArrayToList rows
        NextY = rowY + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale tableBlock
    }

let buildTableLayoutPlanForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY =
    buildTableLayoutPlanWithMemo None ValueNone settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY

let private buildTableLayoutPlanAtWithMemo (memo: ParagraphLayoutMemo option) (initialCounters: Map<string * int, int> voption) settings measureText measureRange doc tableBlock tableY =
    let tableLeftOffset = tableLeftOffsetForDocument memo doc tableBlock
    let maxWidth = max 0.0 (settings.Width - settings.PaddingLeft - settings.MarginRight - tableLeftOffset)
    let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
    let tableMaxWidth = maxWidth + outerLeft + outerRight
    let columnWidths = tableColumnWidths tableMaxWidth tableBlock
    let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
    let tableHorizontalOffset = tableLeftOffset + tableAlignmentOffset tableBlock.Alignment maxWidth tableWidth
    buildTableLayoutPlanWithMemo memo initialCounters settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY

let buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY =
    buildTableLayoutPlanAtWithMemo None ValueNone settings measureText measureRange doc tableBlock tableY

/// Layout document blocks. Table cell paragraphs are included so shared hit/cursor logic works inside tables.
let private footnoteParagraphMetadataForLayout = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }

/// Lay out one page's footnotes in the reserved area at the bottom of the
/// page. The notes are laid out at a temporary origin to get their real height
/// (10pt body font, superscript mark number), then placed so the block ends
/// just above the bottom margin, the way Word stacks footnotes.
let private footnoteLayoutsForPage (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) maxWidth (pageIndex: int) pageContentBottom (notes: (string * string) list) : BlockLayout list =
    // An unpaged page: the notes are positioned explicitly below.
    let noteSettings = { settings with PageHeight = unpagedMeasurementPageHeight }
    let mutable tempY = 0.0
    let laidOut =
        notes |> List.mapi (fun noteIndex (id, text) ->
            let noteString = sprintf "%s %s" id (text.Trim())
            let buffer =
                TextBuffer.ofString noteString
                |> TextBuffer.setFontFamily 0 noteString.Length footnoteFontFamily
                |> TextBuffer.setFontSize 0 noteString.Length footnoteFontSize
                |> TextBuffer.setVertAlign 0 id.Length (Some "superscript")
            let para = { Id = sprintf "%s%d_%d" footnoteBlockIdPrefix pageIndex noteIndex; Buffer = buffer; Metadata = footnoteParagraphMetadataForLayout }
            let layout, _ = layoutParagraphAtWithRange noteSettings measureText measureRange maxWidth tempY 0.0 para Map.empty
            tempY <- tempY + layout.Height
            layout)
    let blockTop = pageContentBottom - footnoteBottomPad - tempY
    laidOut |> List.map (fun layout -> moveLayoutToY (blockTop + layout.Y) layout)

/// Endnote paragraphs for every endnote referenced in `bodyLayouts`, in document
/// order, stacked below `bodyBottom` (Word puts endnotes at the end of the
/// document); Paint draws the separator above the first. `pageAt y` gives the
/// content top and bottom of the page holding `y` and the content top of the
/// page after it, so the notes continue onto further pages.
let private endnoteLayoutsAfterBody (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (bodyLayouts: BlockLayout seq) (bodyBottom: float) (pageAt: float -> struct (float * float * float)) : BlockLayout list =
    let ordered = ResizeArray<BlockLayout>(bodyLayouts)
    ordered.Sort(System.Comparison<BlockLayout>(fun left right -> compare left.Y right.Y))
    let endnotes = ResizeArray<string * string>()
    let seenEndnotes = System.Collections.Generic.HashSet<string>()
    for layout in ordered do
        for line in layout.Lines do
            for id, text in endnotesInLine line do
                if seenEndnotes.Add id then
                    endnotes.Add(id, text)
    let result = ResizeArray<BlockLayout>()
    // An unpaged page: each note is measured for its natural height and placed explicitly.
    let noteSettings = { settings with PageHeight = unpagedMeasurementPageHeight }
    let mutable noteY = bodyBottom + endnoteTopGap
    for noteIndex = 0 to endnotes.Count - 1 do
        let id, text = endnotes.[noteIndex]
        let mark = endnoteMark id
        let noteString = sprintf "%s %s" mark (text.Trim())
        let buffer =
            TextBuffer.ofString noteString
            |> TextBuffer.setFontFamily 0 noteString.Length footnoteFontFamily
            |> TextBuffer.setFontSize 0 noteString.Length footnoteFontSize
            |> TextBuffer.setVertAlign 0 mark.Length (Some "superscript")
        let para = { Id = sprintf "%s%d" endnoteBlockIdPrefix noteIndex; Buffer = buffer; Metadata = footnoteParagraphMetadataForLayout }
        let probe, _ = layoutParagraphAtWithRange noteSettings measureText measureRange maxWidth 0.0 0.0 para Map.empty
        let height = probe.Height
        let struct (contentTop, contentBottom, nextContentTop) = pageAt noteY
        // A note that does not fit on the current page starts the next one.
        let startY = if noteY > contentTop && noteY + height > contentBottom then nextContentTop else noteY
        result.Add(moveLayoutToY startY probe)
        noteY <- startY + height
    resizeArrayToList result

/// Vertical position of a paragraph that follows another one directly, given
/// where it would start (`y`, right after the previous paragraph and its
/// spacing after): adjoining spacing collapses to the larger of the two,
/// contextual spacing drops it between paragraphs of one style, and a border
/// shared by both is not repeated.
let private collapseParagraphSpacing (y: float) (previousPara: ParagraphBlock) (previousStyle: ParagraphStyle) (previousMetadata: BlockMetadata) (previousSpacingAfter: float) (para: ParagraphBlock) (paragraphStyle: ParagraphStyle) (effectiveMetadata: BlockMetadata) =
    let contextualSpacingEnabled (style: ParagraphStyle) (metadata: BlockMetadata) =
        metadata.ContextualSpacing |> Option.orElse style.ContextualSpacing |> Option.defaultValue false
    let spacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle effectiveMetadata
    let stylesMatch = para.Metadata.Style.IsSome && previousPara.Metadata.Style = para.Metadata.Style
    let effectivePreviousSpacingAfter =
        if stylesMatch && contextualSpacingEnabled previousStyle previousMetadata then 0.0 else previousSpacingAfter
    let effectiveSpacingBefore =
        if stylesMatch && contextualSpacingEnabled paragraphStyle effectiveMetadata then 0.0 else spacingBefore
    let collapsed = (y - previousSpacingAfter) + max effectivePreviousSpacingAfter effectiveSpacingBefore - spacingBefore
    let sameBorders =
        previousMetadata.ParagraphBorders.IsSome
        && previousMetadata.ParagraphBorders = effectiveMetadata.ParagraphBorders
        && previousMetadata.ParagraphBorderBefore = effectiveMetadata.ParagraphBorderBefore
        && previousMetadata.ParagraphBorderAfter = effectiveMetadata.ParagraphBorderAfter
    if stylesMatch && sameBorders then
        let _, previousBorderAfter = paragraphBorderVerticalAdvance previousMetadata
        let currentBorderBefore, _ = paragraphBorderVerticalAdvance effectiveMetadata
        collapsed - previousBorderAfter - currentBorderBefore
    else collapsed

/// Whether two states of a block-by-block pass lead to the same layout of
/// the same blocks. (Of the paragraph above, the spacing rules look at its
/// style name, its resolved style and metadata and its spacing after.)
let private samePassState (left: PassCheckpoint) (right: PassCheckpoint) =
    left.CheckCol = right.CheckCol
    && left.CheckPage = right.CheckPage
    && left.CheckY = right.CheckY
    && sameOrEqual left.CheckCounters right.CheckCounters
    && sameOrEqual left.CheckFootnotes right.CheckFootnotes
    && (match left.CheckPrevious, right.CheckPrevious with
        | None, None -> true
        | Some (leftPara, leftStyle, leftMetadata, leftAfter), Some (rightPara, rightStyle, rightMetadata, rightAfter) ->
            leftAfter = rightAfter
            && (obj.ReferenceEquals(leftPara, rightPara) || leftPara.Metadata.Style = rightPara.Metadata.Style)
            && sameOrEqual leftStyle rightStyle
            && sameOrEqual leftMetadata rightMetadata
        | _ -> false)
    && (match left.CheckPreviousLayout, right.CheckPreviousLayout with
        | None, None -> true
        | Some leftLayout, Some rightLayout -> obj.ReferenceEquals(leftLayout, rightLayout)
        | _ -> false)
    && sameOrEqual left.CheckObstacles right.CheckObstacles
    && sameOrEqual left.CheckPlacedBoxes right.CheckPlacedBoxes

/// How many blocks two block arrays have in common at the start and (of the
/// rest) at the end: the same objects in the same order.
// ---- text boxes and shapes ---------------------------------------------------
// A shape's text is laid out in the shape: at its width less its insets
// (wps:bodyPr lIns / rIns), from its top inset down, or - by its vertical
// anchor - in the middle or at the bottom of its height. A shape that is as
// high as its text (a:spAutoFit) has no height of its own.

/// The block id of any block.
let private blockIdOfAny (block: Block) : BlockId =
    match block with
    | Paragraph para
    | Heading (_, para) -> para.Id
    | Table table -> table.Id
    | TextBox textBox -> textBox.Id

let private textBoxSettings (settings: LayoutSettings) : LayoutSettings =
    // (A shape's text is not broken over pages: no page to end on.)
    { settings with MarginBottom = 0.0; InTableCell = false }

/// The paragraphs of a shape's text from `top` down, their left edge at
/// `left` (down and across the canvas), each `width` wide at most.
let private textBoxParagraphLayouts (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (styleContext: DocumentStyleContext) (textBox: TextBoxBlock) (left: float) (top: float) (width: float) : BlockLayout list =
    let boxSettings = textBoxSettings settings
    let result = ResizeArray<BlockLayout>()
    let mutable y = top
    let mutable counters : Map<string * int, int> = Map.empty
    for para in textBox.Content do
        let layout, next =
            layoutParagraphAtWithStyleWithRange boxSettings measureText measureRange width y (left - settings.PaddingLeft) (forBlockInContext styleContext (Paragraph para)) (effectiveBufferInContext styleContext) para counters
        counters <- next
        result.Add layout
        y <- layout.Y + layout.Height
    List.ofSeq result

/// The width a shape's text is broken at.
let textBoxTextWidth (textBox: TextBoxBlock) : float =
    let left, _, right, _ = WordLayout.Floating.textBoxInsets textBox
    if textBox.NoTextWrap || textBox.FitText then 100000.0 else max 1.0 (textBox.Width - left - right)

/// The height of a shape's text.
let textBoxContentHeight (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (styleContext: DocumentStyleContext) (textBox: TextBoxBlock) : float =
    match textBoxParagraphLayouts settings measureText measureRange styleContext textBox settings.PaddingLeft 0.0 (textBoxTextWidth textBox) with
    | [] -> 0.0
    | layouts -> (let last = List.last layouts in last.Y + last.Height)

/// The height of a shape: its own, or that of its text and insets when it follows its text.
let textBoxFrameHeight (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (styleContext: DocumentStyleContext) (textBox: TextBoxBlock) : float =
    if textBox.AutoFit && not textBox.FitText && not textBox.Content.IsEmpty then
        let _, top, _, bottom = WordLayout.Floating.textBoxInsets textBox
        max 1.0 (textBoxContentHeight settings measureText measureRange styleContext textBox + top + bottom)
    else textBox.Height

/// The layouts of a shape's text, the shape's box at `x`, `y` (across and
/// down the canvas) and `height` high: where the painter draws the text and
/// where a click in the shape finds it.
let textBoxContentLayouts (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (styleContext: DocumentStyleContext) (textBox: TextBoxBlock) (x: float) (y: float) (height: float) : BlockLayout list =
    let left, top, _, bottom = WordLayout.Floating.textBoxInsets textBox
    let width = textBoxTextWidth textBox
    let layouts = textBoxParagraphLayouts settings measureText measureRange styleContext textBox (x + left) (y + top) width
    let contentHeight =
        match layouts with
        | [] -> 0.0
        | _ -> (let last = List.last layouts in last.Y + last.Height - (y + top))
    let room = height - top - bottom
    let shift =
        match textBox.VerticalAnchor with
        | Some "ctr" when contentHeight < room -> (room - contentHeight) / 2.0
        | Some "b" when contentHeight < room -> room - contentHeight
        | _ -> 0.0
    if shift = 0.0 then layouts
    else textBoxParagraphLayouts settings measureText measureRange styleContext textBox (x + left) (y + top + shift) width

/// The block a shape that text goes around is anchored in.
let private wrappingAnchorOf (block: Block) : BlockId option =
    match block with
    | TextBox textBox when textBox.AnchorBlockId.IsSome && not (textBox.BehindDoc = Some true) ->
        let wrap = WordLayout.Floating.textBoxWrap textBox
        if WordLayout.Floating.wrapsText wrap || WordLayout.Floating.wrapsTopAndBottom wrap then textBox.AnchorBlockId else None
    | _ -> None

let commonBlockEnds (before: Block array) (after: Block array) =
    let limit = min before.Length after.Length
    let mutable prefix = 0
    while prefix < limit && obj.ReferenceEquals(before.[prefix], after.[prefix]) do
        prefix <- prefix + 1
    let mutable suffix = 0
    while suffix < limit - prefix && obj.ReferenceEquals(before.[before.Length - 1 - suffix], after.[after.Length - 1 - suffix]) do
        suffix <- suffix + 1
    struct (prefix, suffix)

/// The record of the last pass in this place, when it was given what this
/// pass is given (the blocks apart).
let private tryPassRecord (memo: ParagraphLayoutMemo option) (slot: string) (settings: LayoutSettings) (columns: SectionColumns option) startPage (startY: float) (pageTop: float) (nextPageTop: float) nonUniform balanced (initial: PassCheckpoint) (doc: Document) =
    match memo with
    | Some memo when slot <> "" ->
        match memo.MemoPasses.TryGetValue slot with
        | true, record
            when record.PassStartPage = startPage
                 && record.PassStartY = startY
                 && record.PassPageTop = pageTop
                 && record.PassNextPageTop = nextPageTop
                 && record.PassNonUniform = nonUniform
                 && record.PassBalanced = balanced
                 && obj.ReferenceEquals(record.PassStyles, doc.Styles)
                 && obj.ReferenceEquals(record.PassPackage, doc.Package)
                 && sameOrEqual record.PassSettings settings
                 && sameOrEqual record.PassColumns columns
                 && sameOrEqual record.PassPageSetup doc.PageSetup
                 && samePassState record.PassInitial initial -> Some record
        | _ -> None
    | _ -> None

let private resultsHaveEndnotes (results: ResizeArray<BlockLayout>) first last =
    let mutable found = false
    let mutable index = first
    while not found && index < last do
        if layoutHasNoteType "endnote" results.[index] then found <- true
        index <- index + 1
    found

/// `footnoteLayoutsForPage`, kept from pass to pass: a page whose notes, place
/// and geometry are what they were has the paragraphs it had - the same objects.
let private footnoteLayoutsForPageKept (slot: string) (memo: ParagraphLayoutMemo option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) maxWidth (pageIndex: int) pageContentBottom (notes: (string * string) list) : BlockLayout list =
    let key = slot + ":" + string pageIndex
    match memo with
    | None -> footnoteLayoutsForPage settings measureText measureRange maxWidth pageIndex pageContentBottom notes
    | Some memo ->
        match memo.MemoFootnoteAreas.TryGetValue key with
        | true, area when area.AreaMaxWidth = maxWidth && area.AreaContentBottom = pageContentBottom
                          && sameOrEqual area.AreaNotes notes && sameOrEqual area.AreaSettings settings -> area.AreaLayouts
        | _ ->
            let layouts = footnoteLayoutsForPage settings measureText measureRange maxWidth pageIndex pageContentBottom notes
            memo.MemoFootnoteAreas.[key] <- { AreaSettings = settings; AreaMaxWidth = maxWidth; AreaContentBottom = pageContentBottom; AreaNotes = notes; AreaLayouts = layouts }
            layouts

let private layoutDocumentPassKept
        (slot: string)
        (memo: ParagraphLayoutMemo option)
        (settings: LayoutSettings)
        (measureText: TextMeasurer)
        (measureRange: TextRangeMeasurer)
        (doc: Document)
        (initialCounters: Map<string * int, int>)
        (initialObstacles: FloatingObstacle list)
        (initialPlacedBoxes: PlacedFloatingBox list) =
    // Multi-column sections are flowed by columnFlowLayout, which drives this
    // function per column at the column width; here we always use the full width.
    let styleContext = styleContextForLayout memo doc
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let mutable y = settings.PaddingTop
    let layouts = ResizeArray<BlockLayout>()
    let mutable numberedCounters = initialCounters
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let mutable floatingObstacles: FloatingObstacle list = initialObstacles
    let mutable placedFloatingBoxes: PlacedFloatingBox list = initialPlacedBoxes
    let mutable previousParagraphContext: (ParagraphBlock * ParagraphStyle * BlockMetadata * float) option = None
    let mutable previousParagraphLayout: BlockLayout option = None
    let mutable hasEndnoteRefs = false
    let blocks = doc.Blocks |> List.toArray
    let currentState () = {
        CheckCol = 0; CheckPage = 0; CheckY = y; CheckCounters = numberedCounters; CheckFootnotes = footnotesByPage
        CheckPrevious = previousParagraphContext; CheckPreviousLayout = previousParagraphLayout
        CheckObstacles = floatingObstacles; CheckPlacedBoxes = placedFloatingBoxes
    }
    let restoreState (state: PassCheckpoint) =
        y <- state.CheckY
        numberedCounters <- state.CheckCounters
        footnotesByPage <- state.CheckFootnotes
        previousParagraphContext <- state.CheckPrevious
        previousParagraphLayout <- state.CheckPreviousLayout
        floatingObstacles <- state.CheckObstacles
        placedFloatingBoxes <- state.CheckPlacedBoxes
    let initialState = currentState ()
    let lastPass = tryPassRecord memo slot settings None 0 settings.PaddingTop 0.0 0.0 false false initialState doc

    let addBodyLayout layout =
        layouts.Add layout
        if not hasEndnoteRefs && layoutHasNoteType "endnote" layout then
            hasEndnoteRefs <- true

    let nextParagraphBlock index =
        if index + 1 >= blocks.Length then None
        else
            match blocks.[index + 1] with
            | Paragraph para
            | Heading (_, para) -> Some (blocks.[index + 1], para)
            | _ -> None

    let paragraphBlockAt index =
        if index >= blocks.Length then None
        else
            match blocks.[index] with
            | Paragraph para
            | Heading (_, para) -> Some (blocks.[index], para)
            | _ -> None

    let paragraphTextIsEmpty (para: ParagraphBlock) =
        TextBuffer.isWhiteSpace para.Buffer

    let paragraphTextIsNonbreakingBlank (para: ParagraphBlock) =
        TextBuffer.isNonbreakingBlank para.Buffer

    let previousBlockIsTextTable index =
        index > 0
        && match blocks.[index - 1] with
           | Table table ->
               let mutable hasText = false
               for row in table.Rows do
                   if not hasText then
                       for cell in row.Cells do
                           if not hasText then
                               for para in cell.Content do
                                   if not hasText && not (paragraphTextIsEmpty para) then
                                       hasText <- true
               hasText
           | _ -> false

    let previousBlockIsTable index =
        index > 0
        && match blocks.[index - 1] with
           | Table _ -> true
           | _ -> false

    let nextBlockIsEmptyParagraph index =
        match paragraphBlockAt (index + 1) with
        | Some (_, nextPara) -> paragraphTextIsEmpty nextPara
        | None -> false

    let nextNonEmptyParagraphBlock index =
        let mutable probeIndex = index + 1
        let mutable result = None
        while result.IsNone && probeIndex < blocks.Length do
            match paragraphBlockAt probeIndex with
            | Some (nextBlock, nextPara) ->
                if paragraphTextIsEmpty nextPara then
                    probeIndex <- probeIndex + 1
                else
                    result <- Some (nextBlock, nextPara)
            | None ->
                probeIndex <- probeIndex + 1
        result

    let shouldMoveKeepGroup currentY groupHeight =
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) currentY
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
        let pageContentHeight = contentBottom - contentTop
        currentY > contentTop && groupHeight <= pageContentHeight && currentY + groupHeight > contentBottom

    let shouldMoveNearFooter currentY currentHeight lineCount spacingAfter nextLineCount =
        false

    let shouldProbeNearFooterNextLineCount currentY currentHeight lineCount spacingAfter =
        false

    let sameParagraphStyle (previous: ParagraphBlock) (current: ParagraphBlock) =
        current.Metadata.Style.IsSome && previous.Metadata.Style = current.Metadata.Style

    let contextualSpacingEnabled (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
        metadata.ContextualSpacing |> Option.orElse paragraphStyle.ContextualSpacing |> Option.defaultValue false

    let hasExplicitLineSpacing (paragraphStyle: ParagraphStyle) (metadata: BlockMetadata) =
        metadata.LineSpacing.IsSome
        || paragraphStyle.LineSpacing.IsSome
        || metadata.LineSpacingRule.IsSome
        || paragraphStyle.LineSpacingRule.IsSome

    let sameParagraphBorders (previous: BlockMetadata) (current: BlockMetadata) =
        previous.ParagraphBorders.IsSome
        && previous.ParagraphBorders = current.ParagraphBorders
        && previous.ParagraphBorderBefore = current.ParagraphBorderBefore
        && previous.ParagraphBorderAfter = current.ParagraphBorderAfter

    let resetPreviousParagraph () =
        previousParagraphContext <- None
        previousParagraphLayout <- None

    let tableFollowDefaultLineHeightAdjustment () =
        match previousParagraphContext, previousParagraphLayout with
        | Some (_, paragraphStyle, metadata, _), Some layout when not (hasExplicitLineSpacing paragraphStyle metadata) && not (usesDefaultParagraphSpacing metadata) ->
            let mutable adjustment = 0.0
            for line in layout.Lines do
                let fontSize = lineMaxTextFontSize paragraphStyle.Text.FontSize line
                adjustment <- adjustment + max 0.0 (fontSize * (wordTableFollowSingleLineBoxFactor - wordSingleLineBoxFactor))
            adjustment
        | _ -> 0.0

    // The shapes that text goes around, by the block they are anchored in. A
    // shape is placed when the block it is anchored in is laid out - from
    // where that block starts - and the lines of that block and of what
    // follows go around it, as around an anchored picture. (A shape was an
    // obstacle where it stood in the list of blocks - after all of them, for a
    // document that was read - and always on the first page.)
    let anchoredBoxes =
        let mutable found : Dictionary<BlockId, ResizeArray<TextBoxBlock>> = null
        for block in blocks do
            match block, wrappingAnchorOf block with
            | TextBox textBox, Some anchorId ->
                if isNull found then found <- Dictionary<BlockId, ResizeArray<TextBoxBlock>>()
                match found.TryGetValue anchorId with
                | true, boxes -> boxes.Add textBox
                | false, _ ->
                    let boxes = ResizeArray<TextBoxBlock>()
                    boxes.Add textBox
                    found.[anchorId] <- boxes
            | _ -> ()
        found
    let withAnchoredBoxes (anchorId: BlockId) (atY: float) (obstacles: FloatingObstacle list) (placed: PlacedFloatingBox list) : struct (FloatingObstacle list * PlacedFloatingBox list) =
        if isNull anchoredBoxes then struct (obstacles, placed)
        else
            match anchoredBoxes.TryGetValue anchorId with
            | true, boxes ->
                let page : WordLayout.Floating.AnchorPage =
                    { PageWidth = settings.Width; PageHeight = settings.PageHeight; MarginLeft = settings.PaddingLeft; MarginRight = settings.MarginRight
                      MarginTop = settings.PaddingTop; MarginBottom = settings.MarginBottom }
                let pageTop = if settings.PageHeight < 50000.0 then WordLayout.Page.pageTopForY settings.PageHeight atY else 0.0
                let mutable obstacles = obstacles
                let mutable placed = placed
                for textBox in boxes do
                    let height = textBoxFrameHeight settings measureText measureRange styleContext textBox
                    let box = WordLayout.Floating.anchoredTextBoxBox page pageTop atY height textBox
                    let obstacleX, obstacleTop, obstacleWidth, obstacleHeight = WordLayout.Floating.obstacleForBox maxWidth box
                    obstacles <-
                        (if WordLayout.Floating.wrapsTopAndBottom box.Wrap then (0.0, obstacleTop, maxWidth, obstacleHeight)
                         else (obstacleX, obstacleTop, obstacleWidth, obstacleHeight)) :: obstacles
                    placed <- box :: placed
                struct (obstacles, placed)
            | false, _ -> struct (obstacles, placed)

    let addTextBoxObstacle (textBox: TextBoxBlock) =
        let rawFloatingBox = WordLayout.Floating.textBoxCanvasBox settings.PaddingLeft settings.PaddingTop textBox
        let floatingBox = WordLayout.Floating.avoidOverlap placedFloatingBoxes rawFloatingBox
        let obstacleX, obstacleTop, obstacleWidth, obstacleHeight = WordLayout.Floating.relativeObstacleForBox settings.PaddingLeft maxWidth floatingBox
        floatingObstacles <- (obstacleX, obstacleTop, obstacleWidth, obstacleHeight) :: floatingObstacles
        placedFloatingBoxes <- floatingBox :: placedFloatingBoxes

    let layoutTableAt (tableBlock: TableBlock) tableY =
        let tableStart = timingStart ()
        let numColumns = WordLayout.Table.columnCount tableBlock
        let result =
            if numColumns > 0 then
                let tableLeftOffset = tableLeftOffsetForDocument memo doc tableBlock
                let tableAvailableWidth = max 0.0 (maxWidth - tableLeftOffset)
                let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
                let tableMaxWidth = tableAvailableWidth + outerLeft + outerRight
                let columnWidths = tableColumnWidths tableMaxWidth tableBlock
                let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
                let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
                let tableHorizontalOffset = tableLeftOffset + tableAlignmentOffset tableBlock.Alignment tableAvailableWidth tableWidth
                let mutable rowY = tableY
                let headerRows = WordLayout.Table.headerRows tableBlock
                let tableRowCountForTrace =
                    if traceTimingEnabled then Some (List.length tableBlock.Rows) else None
                let traceRows =
                    match tableRowCountForTrace with
                    | Some count -> count >= 10
                    | None -> false
                // The body's list counters at this table: a list that starts in
                // the body text continues in the cells, and on after the table.
                let countersBeforeTable = numberedCounters
                numberedCounters <- listCountersAfterTable countersBeforeTable tableBlock
                let tablePlan = lazy (buildTableLayoutPlanAtWithMemo memo (ValueSome countersBeforeTable) settings measureText measureRange doc tableBlock tableY)
                let tracePlanRows =
                    if traceRows then
                        tablePlan.Value.Rows |> List.map (fun row -> row.RowIndex, row) |> Map.ofList |> Some
                    else
                        None

                let hasRowSpanningVerticalAlignment =
                    let mutable found = false
                    for row in tableBlock.Rows do
                        if not found then
                            for cell in row.Cells do
                                if not found then
                                    match cell.VerticalMerge, cell.VerticalAlignment with
                                    | Some RestartMerge, Some alignment when alignment = "center" || alignment = "bottom" -> found <- true
                                    | _ -> ()
                    found
                let verticalAlignmentHeightResolver = lazy (tableCellVerticalAlignmentHeightResolverForPlan tablePlan.Value)

                let verticalAlignmentHeightForCell rowIndex sourceCellIndex rowHeight =
                    if hasRowSpanningVerticalAlignment then
                        verticalAlignmentHeightResolver.Value rowIndex sourceCellIndex rowHeight
                    else
                        rowHeight

                if traceRows then
                    let rowCount = tableRowCountForTrace |> Option.defaultValue 0
                    traceTableLayout (sprintf "start rows=%d y=%.2f page=%d widths=%s" rowCount tableY (pageNumberForY settings tableY) (columnWidths |> List.map (sprintf "%.2f") |> String.concat ","))

                let measureScope = tableMeasureScopeFrom memo countersBeforeTable tableBlock
                let measureRow (row: TableRow) rowY tableContinuationHeaderHeight =
                    measureTableRowForContextWithColumnMetrics measureScope settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

                let measureRowLocal (row: TableRow) =
                    measureTableRowForContextRowLocalWithColumnMetrics measureScope settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row

                let headerMeasurementsBuffer = ResizeArray()
                for row in headerRows do
                    headerMeasurementsBuffer.Add(measureRowLocal row)
                let headerMeasurementsLocal = headerMeasurementsBuffer.ToArray()
                let headerHeight =
                    let mutable totalHeight = 0.0
                    for header in headerMeasurementsLocal do
                        totalHeight <- totalHeight + header.Height
                    totalHeight

                let addAdjustedLayouts offset queryVAlignOffset (groupLayouts: BlockLayout list) =
                    for layout in groupLayouts do
                        if not layout.Lines.IsEmpty then
                            let adjustedLayout =
                                { layout with
                                    Y = layout.Y + offset
                                    VAlignOffset = layout.VAlignOffset + queryVAlignOffset
                                    Lines = shiftLinesY offset layout.Lines }
                            addBodyLayout adjustedLayout
                            footnotesByPage <- addLayoutFootnotes settings footnotesByPage adjustedLayout
                        else
                            // An empty paragraph of a cell has no line, but it is
                            // where the caret is in a cell nothing was typed into
                            // (every cell of a table that was just inserted): its
                            // place is part of the layout like any paragraph's.
                            addBodyLayout { layout with Y = layout.Y + offset; VAlignOffset = layout.VAlignOffset + queryVAlignOffset }

                let addHeaderLayoutsAt headerY =
                    let mutable headerRowY = headerY
                    let mutable headerRowIndex = 0
                    for headerRow in headerRows do
                        let headerMeasurement = measureRow headerRow headerRowY 0.0
                        let mutable groupIndex = 0
                        for group in headerMeasurement.CellGroups do
                            let alignmentHeight = verticalAlignmentHeightForCell headerRowIndex groupIndex headerMeasurement.Height
                            let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment alignmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                            let queryVAlignOffset = if alignmentHeight > headerMeasurement.Height + 0.5 then 0.0 else offset
                            addAdjustedLayouts offset queryVAlignOffset group.Layouts
                            groupIndex <- groupIndex + 1
                        headerRowY <- headerRowY + headerMeasurement.Height
                        headerRowIndex <- headerRowIndex + 1

                let tableRowArray = List.toArray tableBlock.Rows
                let rowKeepsWithNext = tableRowKeepsWithNext styleContext
                let mergedSplits = mergedCellSplitsFor measureScope
                // (first row, cell) -> where the merged cell's layouts are in `layouts`, and how many.
                let mergedLayoutsAt = System.Collections.Generic.Dictionary<struct (int * int), struct (int * int)>()

                let mutable rowIndex = 0
                for row in tableBlock.Rows do
                    let continuationHeaderHeight = if row.IsHeader then 0.0 else headerHeight
                    let rowLocalMeasurement = measureRowLocal row
                    let mutable measuredRow = rowLocalMeasurement
                    let mutable rowHeight = measuredRow.Height
                    let mutable cellLayoutGroups = moveTableRowMeasurementToY rowY rowLocalMeasurement |> fun measurement -> measurement.CellGroups

                    let pagedRowY = paginateTableRow settings row rowY rowHeight
                    let pagedRowY = paginateTableRowWithFootnotes settings footnotesByPage row pagedRowY rowHeight
                    let contentBottom = contentBottomWithFootnotes settings footnotesByPage rowY
                    let pagedRowY = keepRowWithNextY settings headerHeight contentBottom tableRowArray rowKeepsWithNext measureRowLocal rowIndex rowY rowHeight pagedRowY
                    let pagedRowY =
                        if pagedRowY = rowY
                           && (row.CanSplit |> Option.defaultValue true)
                           && rowY + rowHeight > contentBottom
                           && splitTableRowNeedsNextPage settings rowY contentBottom rowLocalMeasurement then
                            advanceToNextPage settings rowY
                        else
                            pagedRowY
                    let rowContentY =
                        if pagedRowY <> rowY && not row.IsHeader && not headerRows.IsEmpty then
                            pagedRowY + headerHeight
                        else
                            pagedRowY
                    let rowFitsAtCurrentY = rowContentY = rowY && rowY + rowHeight <= contentBottom
                    let rowFitsAfterMove = rowContentY <> rowY
                    if rowFitsAtCurrentY || rowFitsAfterMove then
                        if rowFitsAfterMove && not row.IsHeader && not headerRows.IsEmpty then
                            addHeaderLayoutsAt pagedRowY
                        cellLayoutGroups <- (moveTableRowMeasurementToY rowContentY rowLocalMeasurement).CellGroups
                        rowY <- rowContentY
                    else
                        measuredRow <- measureRow row rowY continuationHeaderHeight
                        rowHeight <- measuredRow.Height
                        cellLayoutGroups <- measuredRow.CellGroups
                        if not row.IsHeader && not headerRows.IsEmpty then
                            let fragments = planTableRowFragments settings (pageNumberForY settings rowY) continuationHeaderHeight rowY rowHeight measuredRow
                            for fragment in fragments do
                                if fragment.RepeatedHeaderHeight > 0.0 then
                                    addHeaderLayoutsAt fragment.Y

                    if traceRows then
                        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) rowY
                        let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
                        let rowSnippet =
                            row.Cells
                            |> List.collect (fun cell -> cell.Content)
                            |> List.map (fun para -> TextBuffer.toString para.Buffer)
                            |> List.tryFind hasNonWhitespace
                            |> Option.map snippet
                            |> Option.defaultValue ""
                        traceTableLayout (sprintf "row=%d y=%.2f page=%d h=%.2f bottom=%.2f contentBottom=%.2f split=%b header=%b text=%s cells=[%s]" rowIndex rowY (pageNumberForY settings rowY) rowHeight (rowY + rowHeight) contentBottom (row.CanSplit |> Option.defaultValue true) row.IsHeader rowSnippet (String.concat " | " measuredRow.Traces))
                        match tracePlanRows |> Option.bind (Map.tryFind rowIndex) with
                        | Some planRow ->
                            let fragmentPages = planRow.Fragments |> List.map (fun fragment -> string fragment.Page) |> List.distinct |> String.concat ","
                            let fragmentLineCounts = planRow.Fragments |> List.map (fun fragment -> string fragment.LineCount) |> String.concat ","
                            let fragmentReasons = planRow.Fragments |> List.map _.BreakReason |> String.concat ","
                            traceTableLayout (sprintf "planCompare row=%d actualY=%.2f actualH=%.2f planY=%.2f planH=%.2f dy=%.2f dh=%.2f fragments=%d pages=%s lines=%s reasons=%s" rowIndex rowY rowHeight planRow.Y planRow.Height (rowY - planRow.Y) (rowHeight - planRow.Height) planRow.Fragments.Length fragmentPages fragmentLineCounts fragmentReasons)
                        | None -> ()

                    for groupIndex, group in cellLayoutGroups |> List.indexed do
                        let alignmentHeight = verticalAlignmentHeightForCell rowIndex groupIndex rowHeight
                        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment alignmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                        // Word's VBA position API ignores single-row cell vAlign (reports the
                        // unaligned origin) but reflects vMerge row-span centering, so only the
                        // single-row offset is subtracted back out in layout queries.
                        let queryVAlignOffset = if alignmentHeight > rowHeight + 0.5 then 0.0 else offset
                        let layoutsBefore = layouts.Count
                        addAdjustedLayouts offset queryVAlignOffset group.Layouts
                        if mergedSplits.IsSome then
                            mergedLayoutsAt.[struct (rowIndex, groupIndex)] <- struct (layoutsBefore, layouts.Count - layoutsBefore)
                    match mergedSplits with
                    | Some splits ->
                        let settled, growth = mergedCellsSettledAt splits settings continuationHeaderHeight rowIndex rowY rowHeight rowLocalMeasurement
                        for (firstRow, cellIndex, moved) in settled do
                            match mergedLayoutsAt.TryGetValue(struct (firstRow, cellIndex)) with
                            | true, struct (layoutsAt, layoutCount) when layoutCount = moved.Length ->
                                let mutable at = layoutsAt
                                for layout in moved do
                                    layouts.[at] <- layout
                                    at <- at + 1
                            | _ -> ()
                        rowHeight <- rowHeight + growth
                    | None -> ()
                    rowY <- rowY + rowHeight
                    rowIndex <- rowIndex + 1

                rowY + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale tableBlock, tableWidth
            else
                tableY + WordLayout.Table.emptyTableAdvance settings.LineHeight, 0.0
        if traceTimingEnabled then
            traceSlow (sprintf "table rows=%d" (List.length tableBlock.Rows)) tableStart
        result

    let layoutBlock (index: int) =
        let block = blocks.[index]
        match block with
        | TextBox textBox when textBox.AnchorBlockId.IsSome ->
            // (Placed with the block it is anchored in: see anchoredBoxes.)
            ()
        | TextBox textBox ->
            resetPreviousParagraph ()
            addTextBoxObstacle textBox
        | Table tableBlock ->
            let tableY = y + tableFollowDefaultLineHeightAdjustment ()
            (let struct (obstacles, placed) = withAnchoredBoxes tableBlock.Id tableY floatingObstacles placedFloatingBoxes
             floatingObstacles <- obstacles
             placedFloatingBoxes <- placed)
            resetPreviousParagraph ()
            let blockStart = timingStart ()
            let nextY, tableWidth = layoutTableAt tableBlock tableY
            if tableBlock.Floating then
                let tableLeftOffset = tableLeftOffsetForContext styleContext tableBlock
                let tableAvailableWidth = max 0.0 (maxWidth - tableLeftOffset)
                let tableHorizontalOffset = tableLeftOffset + tableAlignmentOffset tableBlock.Alignment tableAvailableWidth tableWidth
                let tableHeight = max 0.0 (nextY - tableY - WordLayout.Table.afterSpacing)
                floatingObstacles <- (tableHorizontalOffset, tableY, tableWidth + 12.0, tableHeight) :: floatingObstacles
                placedFloatingBoxes <- {
                    Kind = WordLayout.Floating.TableBox tableBlock
                    X = tableHorizontalOffset
                    Y = tableY
                    Width = tableWidth
                    Height = tableHeight
                    Wrap = Some "square"
                    DistanceTop = 0.0
                    DistanceRight = 12.0
                    DistanceBottom = 0.0
                    DistanceLeft = 0.0
                } :: placedFloatingBoxes
            else
                y <- nextY
            if traceTimingEnabled then
                traceSlow (sprintf "block[%d]:table rows=%d" index (List.length tableBlock.Rows)) blockStart

        | _ ->
            let blockStart = timingStart ()
            let para = getParagraph block
            let paragraphStyle = forBlockInContext styleContext block
            let beforeRenderedHintY = y
            y <- applyRenderedPageBreakHint settings y para
            if y <> beforeRenderedHintY then resetPreviousParagraph ()
            if (para.Metadata.PageBreakBefore || paragraphStyle.PageBreakBefore = Some true) && layouts.Count > 0 then
                y <- advanceToNextPage settings y
                resetPreviousParagraph ()
            let effectiveMetadata = WordStyles.effectiveMetadata paragraphStyle para.Metadata
            let spacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle effectiveMetadata
            let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
            let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
            if y > contentTop then
                match previousParagraphContext with
                | Some (previousPara, previousStyle, previousMetadata, previousSpacingAfter) ->
                    y <- collapseParagraphSpacing y previousPara previousStyle previousMetadata previousSpacingAfter para paragraphStyle effectiveMetadata
                | None -> ()
            let initialStart = timingStart ()
            // (With the shapes anchored in the paragraph that text goes around, placed from its top.)
            let struct (obstaclesAtY, placedAtY) = withAnchoredBoxes para.Id y floatingObstacles placedFloatingBoxes
            let initialLayout, initialCounters, initialObstacles, initialPlacedBoxes = layoutParagraphWithMemo memo settings measureText measureRange maxWidth y 0.0 paragraphStyle (effectiveBufferInContext styleContext) para numberedCounters obstaclesAtY placedAtY
            if traceTimingEnabled then traceSlow (sprintf "block[%d]:initial lines=%d" index initialLayout.Lines.Length) initialStart
            let keepWithNextY =
                if WordStyles.keepNext paragraphStyle para.Metadata then
                    let keepStart = timingStart ()
                    let mutable probeY = y + initialLayout.Height
                    let mutable probeIndex = index + 1
                    let mutable groupHeight = paragraphVisibleFitHeight paragraphStyle initialLayout
                    let mutable shouldContinue = true
                    let mutable probedLines = 0
                    let mutable groupIsMoveable = true

                    while shouldContinue do
                        match paragraphBlockAt probeIndex with
                        | Some (nextBlock, nextPara) ->
                            let nextStyle = forBlockInContext styleContext nextBlock
                            if nextPara.Metadata.PageBreakBefore || nextStyle.PageBreakBefore = Some true then
                                groupIsMoveable <- false
                                shouldContinue <- false
                            else
                                let nextLayout, _, _, _ = layoutParagraphWithMemo memo settings measureText measureRange maxWidth probeY 0.0 nextStyle (effectiveBufferInContext styleContext) nextPara numberedCounters initialObstacles initialPlacedBoxes
                                let nextKeepsFurther = WordStyles.keepNext nextStyle nextPara.Metadata
                                let nextContribution =
                                    if nextKeepsFurther then
                                        paragraphVisibleFitHeight nextStyle nextLayout
                                    else
                                        match nextLayout.Lines with
                                        | firstLine :: _ -> max settings.LineHeight firstLine.Height
                                        | [] -> min nextLayout.Height settings.LineHeight
                                groupHeight <- groupHeight + nextContribution
                                probeY <- nextLayout.Y + nextLayout.Height
                                probedLines <- probedLines + nextLayout.Lines.Length
                                if nextPara.Metadata.PageBreakAfter then
                                    groupIsMoveable <- false
                                    shouldContinue <- false
                                else
                                    shouldContinue <- nextKeepsFurther
                                probeIndex <- probeIndex + 1
                        | None -> shouldContinue <- false

                    if traceTimingEnabled then traceSlow (sprintf "block[%d]:keepNextProbe lines=%d" index probedLines) keepStart
                    if groupIsMoveable && probedLines > 0 && shouldMoveKeepGroup y groupHeight then advanceToNextPage settings y else y
                else y
            let initialLinesSpanPages = linesSpanPages settings initialLayout.Lines
            let normalInitialFitHeight = paragraphVisibleFitHeight paragraphStyle initialLayout
            let emptyParagraphWouldMovePage =
                if initialLinesSpanPages then
                    false
                else
                    let pagedY = paginateBlock settings keepWithNextY normalInitialFitHeight
                    let pagedY = paginateBlockWithFootnotes settings footnotesByPage pagedY normalInitialFitHeight
                    abs (pagedY - keepWithNextY) > 0.01
            let initialFitHeight =
                if initialLayout.Lines.IsEmpty && previousBlockIsTextTable index && not (nextBlockIsEmptyParagraph index) && emptyParagraphWouldMovePage then
                    0.0
                else
                    normalInitialFitHeight
            let pagedY = if initialLinesSpanPages then keepWithNextY else paginateBlock settings keepWithNextY initialFitHeight
            let pagedY = if initialLinesSpanPages then pagedY else paginateBlockWithFootnotes settings footnotesByPage pagedY initialFitHeight
            let nextLineCount =
                if not initialLinesSpanPages && shouldProbeNearFooterNextLineCount pagedY initialLayout.Height initialLayout.Lines.Length initialLayout.Metadata.SpacingAfter then
                    match nextParagraphBlock index with
                    | Some (nextBlock, nextPara) ->
                        let nearFooterStart = timingStart ()
                        let nextStyle = forBlockInContext styleContext nextBlock
                        let nextLayout, _, _, _ = layoutParagraphWithMemo memo settings measureText measureRange maxWidth (pagedY + initialLayout.Height) 0.0 nextStyle (effectiveBufferInContext styleContext) nextPara numberedCounters initialObstacles initialPlacedBoxes
                        if traceTimingEnabled then traceSlow (sprintf "block[%d]:nearFooterProbe lines=%d" index nextLayout.Lines.Length) nearFooterStart
                        Some nextLayout.Lines.Length
                    | None -> None
                else None
            let pagedY =
                if not initialLinesSpanPages && pagedY = keepWithNextY && shouldMoveNearFooter pagedY initialLayout.Height initialLayout.Lines.Length initialLayout.Metadata.SpacingAfter nextLineCount then
                    advanceToNextPage settings pagedY
                else pagedY
            let pagedY =
                if WordStyles.keepLines paragraphStyle para.Metadata && pagedY = y && y > settings.PaddingTop && linesSpanPages settings initialLayout.Lines then
                    advanceToNextPage settings y
                else pagedY
            let layout, counters, obstacles, placedBoxes =
                if pagedY = y then
                    initialLayout, initialCounters, initialObstacles, initialPlacedBoxes
                else
                    let relayoutStart = timingStart ()
                    let result =
                        let struct (obstaclesPaged, placedPaged) = withAnchoredBoxes para.Id pagedY floatingObstacles placedFloatingBoxes
                        layoutParagraphWithMemo memo settings measureText measureRange maxWidth pagedY 0.0 paragraphStyle (effectiveBufferInContext styleContext) para numberedCounters obstaclesPaged placedPaged
                    if traceTimingEnabled then traceSlow (sprintf "block[%d]:relayout" index) relayoutStart
                    result
            let layout =
                if pagedY = y then layout
                else suppressLayoutPageTopBeforeSpacing settings paragraphStyle layout
            let layout =
                if layout.Lines.IsEmpty && previousBlockIsTable index && not (previousBlockIsTextTable index) && not emptyParagraphWouldMovePage then
                    { layout with Height = layout.Height + 5.0 }
                else
                    layout
            let layout =
                if layout.Lines.IsEmpty then
                    match previousParagraphContext, nextNonEmptyParagraphBlock index with
                    | Some (_, _, previousMetadata, _), Some (_, nextPara)
                        when previousMetadata.Style = Some "TOC1" && nextPara.Metadata.Style = Some "NormalWeb" ->
                        let trailingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle layout.Metadata
                        { layout with Height = max 0.0 (layout.Height - trailingAfter) }
                    | _ -> layout
                else
                    layout
            let layout =
                match layout.Lines with
                | [line] when paragraphTextIsNonbreakingBlank para ->
                    { layout with Height = layout.Height + line.Height }
                | _ -> layout
            numberedCounters <- counters
            floatingObstacles <- obstacles
            placedFloatingBoxes <- placedBoxes
            addBodyLayout layout
            footnotesByPage <- addLayoutFootnotes settings footnotesByPage layout
            y <- layout.Y + layout.Height
            previousParagraphContext <- Some (para, paragraphStyle, layout.Metadata, WordLayout.Paragraph.spacingAfter paragraphStyle layout.Metadata)
            previousParagraphLayout <- Some layout
            // (A document with sections is laid out section by section - see layoutSections, which starts
            // each on a new page or not as the section says. In this pass a page breaks after a paragraph that says so.)
            if para.Metadata.PageBreakAfter then
                y <- advanceToNextPage settings y
                resetPreviousParagraph ()
            if traceTimingEnabled then traceSlow (sprintf "block[%d]:paragraph total lines=%d" index layout.Lines.Length) blockStart

    let checkpoints = ResizeArray<PassCheckpoint>(blocks.Length + 1)
    let starts = ResizeArray<int>(blocks.Length + 1)
    let mutable blockIndex = 0
    // The blocks of the last pass that stand as they are: those before the
    // first block that differs - short of the blocks whose place depends on
    // what follows them (paragraphs kept with the next, empty paragraphs,
    // which look ahead to the next paragraph with text) - with the state the
    // pass had there. A block looks back at the block before it (a table?),
    // so the pass is the last one again from the second unchanged block on.
    let mutable convergeFrom = System.Int32.MaxValue
    let mutable indexShift = 0
    let mutable unchanged = false
    match lastPass with
    | Some record ->
        let struct (prefix, suffix) = commonBlockEnds record.PassBlocks blocks
        if prefix = blocks.Length && record.PassBlocks.Length = blocks.Length then
            unchanged <- true
            layouts.AddRange record.PassResults
            restoreState record.PassCheckpoints.[blocks.Length]
            hasEndnoteRefs <- record.PassHasEndnotes
            blockIndex <- blocks.Length
        else
            let isParagraph index = (tryGetParagraph blocks.[index]).IsSome
            let isBlank index =
                match tryGetParagraph blocks.[index] with
                | Some para -> paragraphTextIsEmpty para
                | None -> false
            let keepsNext index =
                match tryGetParagraph blocks.[index] with
                | Some para -> WordStyles.keepNext (forBlockInContext styleContext blocks.[index]) para.Metadata
                | None -> false
            let mutable restart = prefix
            // A shape that text goes around is placed with the block it is anchored in:
            // one that changed (or came, or went) has that block laid out again.
            let anchors = HashSet<BlockId>()
            let noteAnchor (block: Block) =
                match wrappingAnchorOf block with
                | Some id -> anchors.Add id |> ignore
                | None -> ()
            for index in prefix .. blocks.Length - suffix - 1 do noteAnchor blocks.[index]
            for index in prefix .. record.PassBlocks.Length - suffix - 1 do noteAnchor record.PassBlocks.[index]
            if anchors.Count > 0 then
                let mutable index = 0
                while index < restart do
                    if anchors.Contains (blockIdOfAny blocks.[index]) then restart <- index
                    index <- index + 1
            let mutable moved = true
            while moved do
                moved <- false
                while restart > 0 && keepsNext (restart - 1) do
                    restart <- restart - 1
                    moved <- true
                let mutable probe = restart - 1
                let mutable blank = -1
                while probe >= 0 && (not (isParagraph probe) || isBlank probe) do
                    if isBlank probe then blank <- probe
                    probe <- probe - 1
                if blank >= 0 then
                    restart <- blank
                    moved <- true
            for index in 0 .. restart - 1 do
                checkpoints.Add record.PassCheckpoints.[index]
                starts.Add record.PassStarts.[index]
            for index in 0 .. record.PassStarts.[restart] - 1 do
                layouts.Add record.PassResults.[index]
            restoreState record.PassCheckpoints.[restart]
            hasEndnoteRefs <- record.PassHasEndnotes
            blockIndex <- restart
            convergeFrom <- blocks.Length - suffix + 1
            indexShift <- record.PassBlocks.Length - blocks.Length
    | None -> ()
    let mutable keptLayouts = layouts.Count > 0
    let keptPrefix = blockIndex
    let mutable keptSuffix = 0
    while blockIndex < blocks.Length do
        let state = currentState ()
        match lastPass with
        | Some record when blockIndex >= convergeFrom
                           && (layouts.Count > 0) = (record.PassStarts.[blockIndex + indexShift] > 0)
                           && samePassState state record.PassCheckpoints.[blockIndex + indexShift] ->
            // From here on this pass is the last one: its layouts stand.
            let from = record.PassStarts.[blockIndex + indexShift]
            let shift = layouts.Count - from
            for index in blockIndex + indexShift .. record.PassBlocks.Length - 1 do
                checkpoints.Add record.PassCheckpoints.[index]
                starts.Add (record.PassStarts.[index] + shift)
            for index in from .. record.PassResults.Length - 1 do
                layouts.Add record.PassResults.[index]
            if from < record.PassResults.Length then keptLayouts <- true
            keptSuffix <- blocks.Length - blockIndex
            restoreState record.PassCheckpoints.[record.PassBlocks.Length]
            blockIndex <- blocks.Length
        | _ ->
            checkpoints.Add state
            starts.Add layouts.Count
            layoutBlock blockIndex
            blockIndex <- blockIndex + 1
            match memo with
            | Some memo when slot <> "" -> memo.MemoBlocksLaidOut <- memo.MemoBlocksLaidOut + 1
            | _ -> ()
    let bodyRecord =
        match lastPass with
        | Some record when unchanged -> (if record.PassKeptPrefix = blocks.Length then record else { record with PassKeptPrefix = blocks.Length; PassKeptSuffix = 0 })
        | _ ->
            checkpoints.Add (currentState ())
            starts.Add layouts.Count
            {
                PassSettings = settings
                PassColumns = None
                PassStartPage = 0
                PassStartY = settings.PaddingTop
                PassPageTop = 0.0
                PassNextPageTop = 0.0
                PassNonUniform = false
                PassBalanced = false
                PassInitial = initialState
                PassStyles = doc.Styles
                PassPageSetup = doc.PageSetup
                PassPackage = doc.Package
                PassBlocks = blocks
                PassCheckpoints = checkpoints.ToArray()
                PassStarts = starts.ToArray()
                PassResults = layouts.ToArray()
                // (Kept layouts that refer to endnotes were found by the pass that made them. An edit that
                // removes the last reference leaves the flag set: the endnote pass then looks, and finds none.)
                PassHasEndnotes = hasEndnoteRefs
                PassKeptPrefix = keptPrefix
                PassKeptSuffix = keptSuffix
            }
    match memo with
    | Some memo when slot <> "" ->
        memo.MemoPasses.[slot] <- bodyRecord
        if keptLayouts then memo.MemoKeptPass <- true
    | _ -> ()
    let bodyCount = layouts.Count

    // Lay out each page's footnotes in the reserved area at the bottom of the
    // page, at a smaller font, tagged with footnoteBlockIdPrefix so Paint can
    // render them and draw the separator. The note id doubles as Word's mark.
    let footnoteParagraphMetadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }
    let footnoteStride = pageHeight settings + pageGap
    for KeyValue (pageIndex, notes) in footnotesByPage do
        if not notes.IsEmpty then
            let pageTop = float pageIndex * footnoteStride
            let pageContentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            layouts.AddRange(
                if slot = "" then footnoteLayoutsForPage settings measureText measureRange maxWidth pageIndex pageContentBottom notes
                else footnoteLayoutsForPageKept slot memo settings measureText measureRange maxWidth pageIndex pageContentBottom notes)

    // Endnotes: collect every endnote in document order and lay them out after
    // the body content (Word stacks endnotes at the end of the document), with a
    // separator drawn above the first by Paint.
    if hasEndnoteRefs then
        let bodyLayouts = ResizeArray<BlockLayout>()
        let mutable bodyBottom = settings.PaddingTop
        for layout in layouts do
            if not (layout.BlockId.StartsWith footnoteBlockIdPrefix) then
                bodyLayouts.Add layout
                bodyBottom <- max bodyBottom (layout.Y + layout.Height)
        let stride = pageHeight settings + pageGap
        let pageAt noteY =
            let pageIndex = int (floor (noteY / stride))
            let pageTop = float pageIndex * stride
            struct (WordLayout.Page.contentTop settings.PaddingTop pageTop,
                    WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop,
                    WordLayout.Page.contentTop settings.PaddingTop (float (pageIndex + 1) * stride))
        layouts.AddRange(endnoteLayoutsAfterBody settings measureText measureRange maxWidth bodyLayouts bodyBottom pageAt)

    let notes = ResizeArray<BlockLayout>()
    for index in bodyCount .. layouts.Count - 1 do
        notes.Add layouts.[index]
    resizeArrayToList layouts, placedFloatingBoxes, bodyRecord, resizeArrayToList notes

let private layoutDocumentWithRangeAndInitial
        (memo: ParagraphLayoutMemo option)
        (settings: LayoutSettings)
        (measureText: TextMeasurer)
        (measureRange: TextRangeMeasurer)
        (doc: Document)
        (initialCounters: Map<string * int, int>)
        (initialObstacles: FloatingObstacle list)
        (initialPlacedBoxes: PlacedFloatingBox list) =
    let layouts, placed, _, _ = layoutDocumentPassKept "" memo settings measureText measureRange doc initialCounters initialObstacles initialPlacedBoxes
    layouts, placed

// Walk a first-pass layout result and shift each placed float's obstacle top
// up to the top of whichever line first sits "inside" the float's effective
// vertical claim. Word's effective claim per paragraph is line + spacing-after
// (the paragraph BLOCK extent), so a heading whose ink ends above the image
// can still claim the float horizontally if the paragraph's spacing-after
// stretches its block bottom into the image span. Preceding paragraphs get
// re-wrapped against that extended top in the second pass.
let private computeAnchorExtendedObstacles
        (availableWidth: float)
        (layouts: BlockLayout list)
        (placed: PlacedFloatingBox list) : FloatingObstacle list * PlacedFloatingBox list =
    // For each line, record (lineTop, paragraphBlockBottom). Using block bottom
    // — which already includes spacing-after — instead of line bottom is what
    // makes a one-line heading whose ink ends above the float top still count
    // as overlapping when the float top sits inside the heading's spacing-after.
    let lineSpans =
        [ for layout in layouts do
            let blockBottom = layout.Y + layout.Height
            for line in layout.Lines do
                yield line.Y, blockBottom ]
    let mutable anyExtended = false
    let extended =
        placed
        |> List.collect (fun box ->
            if not (WordLayout.Floating.participatesInWrap box) then []
            else
            // (A picture text wraps along the outline of: the outline as it is - it is not stretched up to a line.)
            match WordLayout.Floating.outlineObstacles availableWidth box with
            | Some outline -> outline
            | None ->
                let obstacleX, obstacleTop, obstacleWidth, obstacleHeight =
                    WordLayout.Floating.obstacleForBox availableWidth box
                let floatBottom = obstacleTop + obstacleHeight
                if WordLayout.Floating.wrapsTopAndBottom box.Wrap then
                    // For wrapTopAndBottom we never widen the obstacle upward
                    // to the anchor paragraph's first wrap line: the obstacle
                    // already forbids any text on its vertical span (it
                    // claims the full available width), and pulling the top
                    // up would shove preceding/anchor lines below the
                    // picture even when they'd naturally fit above. Just
                    // widen the in-place obstacle horizontally so the second
                    // pass treats it as full-width without re-adding.
                    [ (0.0, obstacleTop, availableWidth, obstacleHeight) ]
                elif
                    match box.Kind with
                    | WordLayout.Floating.ImageBox img ->
                        match img.PositionVRelativeFrom with
                        | Some "page" | Some "margin" | Some "topMargin" | Some "bottomMargin" -> false
                        | Some _ -> true  // paragraph, line, etc. — image position is already paragraph-anchored
                        | None -> false
                    | _ -> false
                then
                    // Paragraph/line-relative posV: the image's Y is already
                    // computed off the containing paragraph's top, so the
                    // first overlapping line is the right wrap anchor without
                    // further upward extension. Keeping the obstacle at its
                    // literal posV offset lets lines above the image flow
                    // full width, matching Word.
                    [ (obstacleX, obstacleTop, obstacleWidth, obstacleHeight) ]
                else
                    let firstOverlap =
                        lineSpans
                        |> List.tryFind (fun (lineTop, blockBottom) ->
                            blockBottom > obstacleTop && lineTop < floatBottom)
                    let extendedTop =
                        match firstOverlap with
                        | Some (lineTop, _) when lineTop < obstacleTop ->
                            anyExtended <- true
                            lineTop
                        | _ -> obstacleTop
                    let extendedHeight = obstacleHeight + (obstacleTop - extendedTop)
                    [ (obstacleX, extendedTop, obstacleWidth, extendedHeight) ])
    if anyExtended then extended, placed else [], []

let private shiftSegmentX dx segment =
    match segment with
    | TextSegment (t, st, x, w, off) -> TextSegment (t, st, x + dx, w, off)
    | ImageSegment (src, x, yOff, w, h, off, fl, wr, rh, cl, ct, cr, cb) ->
        ImageSegment (src, x + dx, yOff, w, h, off, fl, wr, rh, cl, ct, cr, cb)
    | MathSegment (p, st, x, w, asc, desc, off, len) -> MathSegment (p, st, x + dx, w, asc, desc, off, len)

let private shiftSegmentsX dx segments =
    let shifted = ResizeArray<Segment>()
    for segment in segments do
        shifted.Add(shiftSegmentX dx segment)
    resizeArrayToList shifted

let private shiftLinesX dx lines =
    let shifted = ResizeArray<MeasuredLine>()
    for line in lines do
        shifted.Add { line with Segments = shiftSegmentsX dx line.Segments }
    resizeArrayToList shifted

let private shiftBlockX dx (layout: BlockLayout) =
    if dx = 0.0 then layout
    else
        { layout with
            Indent = layout.Indent + dx
            Lines = shiftLinesX dx layout.Lines }

let private shiftBlockY dy (layout: BlockLayout) =
    if dy = 0.0 then layout
    else
        { layout with
            Y = layout.Y + dy
            Lines = shiftLinesY dy layout.Lines }

let private takeLines count (lines: MeasuredLine list) =
    let taken = ResizeArray<MeasuredLine>()
    let mutable remainingCount = count
    let mutable remainingLines: MeasuredLine list = lines
    let mutable taking = remainingCount > 0
    while taking do
        match remainingLines with
        | line :: rest when remainingCount > 0 ->
            taken.Add line
            remainingCount <- remainingCount - 1
            remainingLines <- rest
        | _ -> taking <- false
    resizeArrayToList taken

/// Lay the document body out column by column. Each paragraph is laid out at the
/// current column's width and, if it overflows, split at the column boundary —
/// the remaining text continues in the next column at *that* column's width
/// (so columns of differing widths wrap independently). After `count` columns
/// the flow wraps to the next page. This drives the doc-aware paragraph layout,
/// so styles/numbering/spacing are the engine's, not a re-implementation.
/// Flow `doc`'s blocks into `columns` starting at (startPage, startY). When
/// `balanced`, the section's columns are made roughly equal height (Word
/// balances a column section that ends with a section break); otherwise each
/// column fills to the page bottom and the flow wraps to the next page after
/// `count` columns. Returns the placed layouts.
/// Move a continuation fragment's offsets back into the paragraph's own
/// offset space (it was laid out from a buffer with the leading text removed).
let private shiftLayoutOffsets (delta: int) (layout: BlockLayout) =
    if delta = 0 then layout
    else
        let shiftSegment segment =
            match segment with
            | TextSegment (text, style, x, width, startOffset) -> TextSegment (text, style, x, width, startOffset + delta)
            | ImageSegment (src, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
                ImageSegment (src, x, yOffset, width, height, startOffset + delta, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom)
            | MathSegment (prims, style, x, width, ascent, descent, startOffset, length) ->
                MathSegment (prims, style, x, width, ascent, descent, startOffset + delta, length)
        { layout with
            Lines =
                layout.Lines
                |> List.map (fun line ->
                    { line with
                        StartOffset = line.StartOffset + delta
                        EndOffset = line.EndOffset + delta
                        Segments = line.Segments |> List.map shiftSegment }) }

/// The page a block of a column is laid out on by itself: as wide as the
/// column, and without an end (the column flow places what was laid out).
let private columnStripSettings (settings: LayoutSettings) (columnWidth: float) : LayoutSettings =
    { settings with Width = settings.PaddingLeft + columnWidth + settings.MarginRight; PageHeight = 100000.0 }

/// The rows of a table that is laid out on a column strip (see
/// `columnStripSettings`): the plan of the table there and, row by row, its
/// top and bottom. A table of a multi-column section is laid out so and then
/// placed row by row - in the layout pass and, for its borders, by the paint.
let tableStripRows (memo: ParagraphLayoutMemo option) (initialCounters: Map<string * int, int> voption) (stripSettings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (tableBlock: TableBlock) : TableLayoutPlan * (float * float) array =
    let plan = buildTableLayoutPlanAtWithMemo memo initialCounters stripSettings measureText measureRange doc tableBlock stripSettings.PaddingTop
    plan, (plan.Rows |> List.map (fun row -> row.Y, row.Y + row.Height) |> List.toArray)

/// The row (of the table itself) each paragraph of a table is in - those of the tables in its cells too.
let tableRowOfParagraphs (tableBlock: TableBlock) : System.Collections.Generic.Dictionary<BlockId, int> =
    let rowOf = System.Collections.Generic.Dictionary<BlockId, int>()
    let rec addTable (rowIndex: int) (table: TableBlock) =
        for row in table.Rows do
            for cell in row.Cells do
                for para in cell.Content do
                    if not (rowOf.ContainsKey para.Id) then rowOf.[para.Id] <- rowIndex
                for nested in cell.NestedTables do
                    addTable rowIndex nested
    let mutable rowIndex = 0
    for row in tableBlock.Rows do
        for cell in row.Cells do
            for para in cell.Content do
                if not (rowOf.ContainsKey para.Id) then rowOf.[para.Id] <- rowIndex
            for nested in cell.NestedTables do
                addTable rowIndex nested
        rowIndex <- rowIndex + 1
    rowOf

/// The columns of the section a top-level table is in, with the section's
/// page geometry applied to `settings`: Some when the section has two or
/// more columns. (What the paint needs to find the strip the layout pass
/// laid the table out on.)
let tableColumnSection (settings: LayoutSettings) (doc: Document) (tableBlock: TableBlock) : (LayoutSettings * SectionColumns) option =
    let mutable found = false
    let mutable result = None
    let mutable pending = false
    for block in doc.Blocks do
        if not found then
            match block with
            | Table table when obj.ReferenceEquals(table, tableBlock) || table.Id = tableBlock.Id -> pending <- true
            | _ -> ()
            match tryGetParagraph block with
            | Some para when para.Metadata.SectionBreak.IsSome ->
                if pending then
                    found <- true
                    match para.Metadata.SectionColumns with
                    | Some columns when columns.Count >= 2 ->
                        let sectionSettings =
                            match para.Metadata.SectionPageSetup with
                            | Some g -> { settings with Width = g.PageWidth; PageHeight = g.PageHeight; PaddingLeft = g.MarginLeft; PaddingTop = g.MarginTop; MarginRight = g.MarginRight; MarginBottom = g.MarginBottom }
                            | None -> settings
                        result <- Some (sectionSettings, columns)
                    | _ -> ()
            | _ -> ()
    if not found && pending then
        match doc.PageSetup.Columns with
        | Some columns when columns.Count >= 2 -> result <- Some (settings, columns)
        | _ -> ()
    result

/// The width of the first column of a section's columns (the columns of a
/// section are of one width unless it states their widths).
let sectionColumnWidth (settings: LayoutSettings) (columns: SectionColumns) : float =
    let n = max 1 columns.Count
    let full = settings.Width - settings.PaddingLeft - settings.MarginRight
    if columns.Widths.Length = n then fst columns.Widths.Head
    else (full - float (n - 1) * columns.Gap) / float n

/// The strip a table of a multi-column section is laid out on (None in a section of one column).
let tableColumnStripSettings (settings: LayoutSettings) (doc: Document) (tableBlock: TableBlock) : LayoutSettings option =
    tableColumnSection settings doc tableBlock
    |> Option.map (fun (sectionSettings, columns) -> columnStripSettings sectionSettings (sectionColumnWidth sectionSettings columns))

let private flowColumnsKept (slot: string) (memo: ParagraphLayoutMemo option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (columns: SectionColumns) (pageTopOf: int -> float) (nonUniform: bool) (startPage: int) (startY: float) (balanced: bool) (initialFootnotes: Map<int, (string * string) list>) =
    let styleContext = styleContextForLayout memo doc
    let blocks = List.toArray doc.Blocks
    let initialState = {
        CheckCol = 0; CheckPage = startPage; CheckY = startY; CheckCounters = Map.empty; CheckFootnotes = initialFootnotes
        CheckPrevious = None; CheckPreviousLayout = None; CheckObstacles = []; CheckPlacedBoxes = []
    }
    let lastPass = tryPassRecord memo slot settings (Some columns) startPage startY (pageTopOf startPage) (pageTopOf (startPage + 1)) nonUniform balanced initialState doc
    // The last pass over the very same blocks: it stands as a whole.
    let unchangedPass =
        match lastPass with
        | Some record when record.PassBlocks.Length = blocks.Length ->
            let struct (prefix, _) = commonBlockEnds record.PassBlocks blocks
            if prefix = blocks.Length then Some record else None
        | _ -> None
    let n = max 1 columns.Count
    let full = settings.Width - settings.PaddingLeft - settings.MarginRight
    let widthsArr =
        if columns.Widths.Length = n then columns.Widths |> Array.ofList
        else
            let colWidth = (full - float (n - 1) * columns.Gap) / float n
            Array.init n (fun i -> (colWidth, (if i < n - 1 then columns.Gap else 0.0)))
    let columnWidth col = fst widthsArr.[min col (n - 1)]
    let columnDx col =
        let mutable x = 0.0
        for i in 0 .. col - 1 do
            x <- x + fst widthsArr.[i] + snd widthsArr.[i]
        x
    let contentTopAt page = WordLayout.Page.contentTop settings.PaddingTop (pageTopOf page)
    let contentBottomAt page = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom (pageTopOf page)
    // When pages are non-uniform (multi-section), lay each paragraph out without
    // the engine's own (uniform-stride) pagination — flowColumns controls
    // placement itself, and internal pagination would mis-position content whose
    // absolute Y doesn't match the uniform grid. For single-section columns the
    // grid is uniform, so keep the engine's pagination (and its widow/orphan
    // handling) unchanged.
    let flowSettings = if nonUniform then { settings with PageHeight = 1_000_000.0 } else settings
    // Balanced target height: lay the section out once at column width to get its
    // total height H, then each column targets H/n (on the start page).
    let balancedHeight =
        if balanced && unchangedPass.IsNone then
            let w = columnWidth 0
            let sub, _ = layoutDocumentWithRangeAndInitial memo { settings with Width = settings.PaddingLeft + w + settings.MarginRight; PageHeight = 100000.0 } measureText measureRange doc Map.empty [] []
            // The sub-pass adds its own footnote / endnote paragraphs; they are
            // not part of the text being balanced.
            let sub =
                sub |> List.filter (fun layout ->
                    not (layout.BlockId.StartsWith footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith endnoteBlockIdPrefix))
            let bottom = maxLineBottom settings.PaddingTop sub
            let height = (bottom - settings.PaddingTop) / float n + 2.0
            // (Columns are balanced on the page the section starts on. A section with more text
            // than that page has room for fills its columns page by page instead - it ran on
            // below the page in columns as long as half of it.)
            if startY + height > contentBottomAt startPage + 0.5 then None else Some height
        else None

    let results = ResizeArray<BlockLayout>()
    let mutable col = 0
    let mutable page = startPage
    let mutable y = startY
    let mutable counters : Map<string * int, int> = Map.empty
    // Footnotes referenced so far, by page. Their block sits at the bottom of
    // the page, so each page's columns end above it.
    let mutable footnotesByPage = initialFootnotes
    let pageFootnotes () = footnotesByPage |> Map.tryFind page |> Option.defaultValue []
    let layoutFootnotes (lines: MeasuredLine list) =
        let mutable notes = []
        for line in lines do
            notes <- appendNewNotes notes (footnotesInLine line)
        notes
    let addPageFootnotes notes =
        match notes with
        | [] -> ()
        | _ -> footnotesByPage <- footnotesByPage |> Map.add page (appendNewNotes (pageFootnotes ()) notes)
    // First column-row starts at startY; later pages start at their content top.
    let columnTop () =
        match balancedHeight with
        | Some _ -> startY
        | None -> if page = startPage then startY else contentTopAt page
    let columnBottom () =
        match balancedHeight with
        // The last balanced column has no limit so a small snapping overflow stays
        // in it rather than wrapping back over column 0.
        | Some _ when col >= n - 1 -> System.Double.MaxValue
        | Some bh -> startY + bh
        | None -> contentBottomAt page - footnoteReserveHeight settings (pageFootnotes ())
    let atColumnTop () = abs (y - columnTop ()) < 0.5
    // The paragraph placed just above in this column, for spacing collapse:
    // (paragraph, style, effective metadata, spacing after). None at a column top.
    let mutable previousParagraph: (ParagraphBlock * ParagraphStyle * BlockMetadata * float) option = None
    let advanceColumn () =
        col <- col + 1
        if col >= n then
            col <- 0
            if balancedHeight.IsNone then page <- page + 1
        y <- columnTop ()
        previousParagraph <- None

    let currentState () = {
        CheckCol = col; CheckPage = page; CheckY = y; CheckCounters = counters; CheckFootnotes = footnotesByPage
        CheckPrevious = previousParagraph; CheckPreviousLayout = None; CheckObstacles = []; CheckPlacedBoxes = []
    }
    let restoreState (state: PassCheckpoint) =
        col <- state.CheckCol
        page <- state.CheckPage
        y <- state.CheckY
        counters <- state.CheckCounters
        footnotesByPage <- state.CheckFootnotes
        previousParagraph <- state.CheckPrevious
    // Keep-with-next, as in the single-column pass: a paragraph that keeps with
    // the next one moves to the next column together with the paragraphs it is
    // chained to when the group (those paragraphs and the first line of what
    // follows) does not fit below it but would fit in a fresh column.
    let keepGroupMovesToNextColumn index (style: ParagraphStyle) (paragraphMetadata: BlockMetadata) (layout: BlockLayout) colBottom =
        if not (WordStyles.keepNext style paragraphMetadata) || atColumnTop () then false
        else
            let mutable groupHeight = paragraphVisibleFitHeight style layout
            let mutable probeY = layout.Y + layout.Height
            let mutable probeIndex = index + 1
            let mutable probing = true
            let mutable probed = false
            let mutable movable = true
            while probing do
                let next = if probeIndex < blocks.Length then tryGetParagraph blocks.[probeIndex] else None
                match next with
                | Some nextPara ->
                    let nextStyle = forBlockInContext styleContext blocks.[probeIndex]
                    if nextPara.Metadata.PageBreakBefore || nextStyle.PageBreakBefore = Some true then
                        movable <- false
                        probing <- false
                    else
                        let nextLayout, _, _, _ =
                            layoutParagraphWithMemo memo flowSettings measureText measureRange (columnWidth col) probeY 0.0 nextStyle (effectiveBufferInContext styleContext) nextPara counters [] []
                        let keepsFurther = WordStyles.keepNext nextStyle nextPara.Metadata
                        groupHeight <-
                            groupHeight
                            + (if keepsFurther then paragraphVisibleFitHeight nextStyle nextLayout
                               else
                                   match nextLayout.Lines with
                                   | firstLine :: _ -> max settings.LineHeight firstLine.Height
                                   | [] -> min nextLayout.Height settings.LineHeight)
                        probeY <- nextLayout.Y + nextLayout.Height
                        probed <- true
                        probing <- keepsFurther && not nextPara.Metadata.PageBreakAfter
                        probeIndex <- probeIndex + 1
                | None -> probing <- false
            movable && probed && y + groupHeight > colBottom + 0.5 && groupHeight <= colBottom - columnTop ()

    let flowBlock (blockIndex: int) =
        let block = blocks.[blockIndex]
        match tryGetParagraph block with
        | Some para0 ->
            if (para0.Metadata.PageBreakBefore || para0.Metadata.RenderedPageBreakBefore.IsSome) && not (atColumnTop ()) then
                advanceColumn ()
            let style = forBlockInContext styleContext block
            let mutable keepChecked = false
            let mutable remaining = para0
            // Characters of the paragraph already placed in earlier columns.
            // A continuation is laid out from the rest of the text, so its
            // offsets are moved back by this much: every fragment of a
            // paragraph reports offsets in the paragraph's own text.
            let mutable consumedOffset = 0
            let mutable placed = false
            while not placed do
                let colBottom = columnBottom ()
                let isContinuation = consumedOffset > 0
                // Spacing between this paragraph and the one above it in the
                // column collapses as in the single-column pass.
                let layoutY =
                    match previousParagraph with
                    | Some (previousPara, previousStyle, previousMetadata, previousSpacingAfter) when not isContinuation && not (atColumnTop ()) ->
                        collapseParagraphSpacing y previousPara previousStyle previousMetadata previousSpacingAfter remaining style (WordStyles.effectiveMetadata style remaining.Metadata)
                    | _ -> y
                let rawLayout, newCounters, _, _ =
                    layoutParagraphWithMemo memo flowSettings measureText measureRange (columnWidth col) layoutY 0.0 style (effectiveBufferInContext styleContext) remaining counters [] []
                let moveKeepGroup =
                    if keepChecked || isContinuation then false
                    else
                        keepChecked <- true
                        keepGroupMovesToNextColumn blockIndex style para0.Metadata rawLayout colBottom
                let layout =
                    if isContinuation then
                        // The list marker belongs to the first fragment only.
                        { shiftLayoutOffsets consumedOffset rawLayout with Prefix = ""; PrefixStyle = None }
                    else rawLayout
                // A line fits when it ends above the column bottom, which
                // moves up as the lines placed so far bring footnotes with them.
                let existingNotes = pageFootnotes ()
                let existingReserve = if balancedHeight.IsSome then 0.0 else footnoteReserveHeight settings existingNotes
                let mutable fitNotes = existingNotes
                let mutable fitCount = 0
                let mutable lastFit = None
                let mutable stillFitting = true
                for line in layout.Lines do
                    if stillFitting then
                        let lineNotes = footnotesInLine line
                        let candidateNotes = if lineNotes.IsEmpty then fitNotes else appendNewNotes fitNotes lineNotes
                        let extraReserve =
                            if balancedHeight.IsSome || LanguagePrimitives.PhysicalEquality candidateNotes existingNotes then 0.0
                            else footnoteReserveHeight settings candidateNotes - existingReserve
                        if line.Y + line.Height <= colBottom - extraReserve + 0.5 then
                            fitCount <- fitCount + 1
                            lastFit <- Some line
                            fitNotes <- candidateNotes
                        else
                            stillFitting <- false
                let placeWhole () =
                    results.Add(shiftBlockX (columnDx col) layout)
                    addPageFootnotes (layoutFootnotes layout.Lines)
                    counters <- newCounters
                    y <- layout.Y + layout.Height
                    previousParagraph <- Some (remaining, style, layout.Metadata, WordLayout.Paragraph.spacingAfter style layout.Metadata)
                    placed <- true
                if moveKeepGroup then advanceColumn ()
                elif fitCount = layout.Lines.Length then placeWhole ()
                elif fitCount = 0 then
                    if atColumnTop () then placeWhole ()
                    else advanceColumn ()
                else
                    let fitLines = takeLines fitCount layout.Lines
                    match lastFit with
                    | Some lastFit ->
                        results.Add(shiftBlockX (columnDx col) { layout with Lines = fitLines; Height = lastFit.Y + lastFit.Height - layout.Y })
                        addPageFootnotes (layoutFootnotes fitLines)
                        // lastFit.EndOffset is in paragraph offsets; the buffer
                        // being cut starts at consumedOffset.
                        let cut = lastFit.EndOffset - consumedOffset
                        let continuationMetadata =
                            { remaining.Metadata with
                                SpacingBefore = Some 0.0
                                FirstLineIndent = if remaining.Metadata.ListStyle = NoList then Some 0.0 else remaining.Metadata.FirstLineIndent }
                        remaining <-
                            // The rest of the paragraph: the same object as at the last
                            // pass while it is cut from the same buffer at the same place.
                            let build () = { remaining with Buffer = TextBuffer.delete 0 cut remaining.Buffer; Metadata = continuationMetadata }
                            match memo with
                            | None -> build ()
                            | Some memo ->
                                let key = remaining.Id + "\u0001" + string consumedOffset
                                match memo.MemoContinuations.TryGetValue key with
                                | true, (source, sourceCut, known) when obj.ReferenceEquals(source, remaining.Buffer) && sourceCut = cut
                                                                        && sameOrEqual known.Metadata continuationMetadata -> known
                                | _ ->
                                    let rest = build ()
                                    memo.MemoContinuations.[key] <- (remaining.Buffer, cut, rest)
                                    rest
                        consumedOffset <- lastFit.EndOffset
                        advanceColumn ()
                    | None ->
                        advanceColumn ()
            if para0.Metadata.PageBreakAfter && not (atColumnTop ()) then advanceColumn ()
        | None ->
            previousParagraph <- None
            let sub = { doc with Blocks = [block] }
            // (As wide as the section's first column: a table keeps its width from column to column.)
            let stripSettings = columnStripSettings settings (columnWidth 0)
            let subLayouts, _ =
                layoutDocumentWithRangeAndInitial memo stripSettings measureText measureRange sub counters [] []
            // The sub-pass lays its own footnotes out at the bottom of its
            // 100,000px page; they belong at the bottom of the real page.
            let subLayouts =
                subLayouts
                |> List.filter (fun layout ->
                    not (layout.BlockId.StartsWith footnoteBlockIdPrefix) && not (layout.BlockId.StartsWith endnoteBlockIdPrefix))
            let mutable subTop = System.Double.PositiveInfinity
            let mutable subBottom = System.Double.NegativeInfinity
            for layout in subLayouts do
                if layout.Y < subTop then subTop <- layout.Y
                let layoutBottom = layout.Y + layout.Height
                if layoutBottom > subBottom then subBottom <- layoutBottom
            let placeWhole () =
                let height = subBottom - subTop
                if not (atColumnTop ()) && y + height > columnBottom () + 0.5 then advanceColumn ()
                let dy = y - subTop
                for sl in subLayouts do
                    results.Add(sl |> shiftBlockY dy |> shiftBlockX (columnDx col))
                    addPageFootnotes (layoutFootnotes sl.Lines)
                y <- y + height
            // A table goes from column to column (and page to page) row by row: a row that
            // does not fit under the rows above it starts the next column, below the table's
            // header rows, which are repeated there. (It was placed whole: a table longer than
            // what was left of the column left that empty, and one longer than a column ran
            // on below the page.) A row is not divided: one taller than a column stands at
            // the top of a column of its own.
            let placeRows (tableBlock: TableBlock) =
                let plan, bands = tableStripRows memo (ValueSome counters) stripSettings measureText measureRange sub tableBlock
                let rowOf = tableRowOfParagraphs tableBlock
                // The strip of the plan against the strip the layouts are on (they are the same, up to where the table starts).
                let planned =
                    plan.Rows |> List.tryPick (fun row -> row.Measurement.CellGroups |> List.tryPick (fun group -> List.tryHead group.Layouts))
                let actual = planned |> Option.bind (fun layout -> subLayouts |> List.tryFind (fun candidate -> candidate.BlockId = layout.BlockId))
                match planned, actual with
                | Some planned, Some actual when bands.Length = tableBlock.Rows.Length && bands.Length > 1 ->
                    let delta = actual.Y - planned.Y
                    let rowLayouts = Array.init bands.Length (fun _ -> ResizeArray<BlockLayout>())
                    let mutable lastRow = 0
                    for sl in subLayouts do
                        match rowOf.TryGetValue sl.BlockId with
                        | true, rowIndex when rowIndex < bands.Length -> lastRow <- rowIndex
                        | _ -> ()
                        rowLayouts.[lastRow].Add sl
                    let top rowIndex = fst bands.[rowIndex] + delta
                    let bottom rowIndex = snd bands.[rowIndex] + delta
                    let headerCount = min plan.HeaderRowCount (bands.Length - 1)
                    let addRow rowIndex dy =
                        for sl in rowLayouts.[rowIndex] do
                            results.Add(sl |> shiftBlockY dy |> shiftBlockX (columnDx col))
                            addPageFootnotes (layoutFootnotes sl.Lines)
                    let mutable dy = y - top 0
                    // (Whether a row stands in this column already: the first one of a column is placed whatever its height.)
                    let mutable rowsInColumn = if atColumnTop () then 0 else 1
                    for rowIndex in 0 .. bands.Length - 1 do
                        // (The header rows go with the first row below them.)
                        let needed = if rowIndex < headerCount then bottom headerCount else bottom rowIndex
                        if rowsInColumn > 0 && (rowIndex = 0 || rowIndex >= headerCount) && needed + dy > columnBottom () + 0.5 then
                            advanceColumn ()
                            rowsInColumn <- 0
                            if rowIndex >= headerCount && headerCount > 0 then
                                let headerDy = y - top 0
                                for headerRow in 0 .. headerCount - 1 do
                                    addRow headerRow headerDy
                                y <- y + (top headerCount - top 0)
                            dy <- y - top rowIndex
                        addRow rowIndex dy
                        rowsInColumn <- rowsInColumn + 1
                        y <- bottom rowIndex + dy
                    // (The space after the table.)
                    y <- y + max 0.0 (subBottom - bottom (bands.Length - 1))
                | _ -> placeWhole ()
            match block, subLayouts with
            | _, [] -> ()
            | Table tableBlock, _ when not tableBlock.Floating -> placeRows tableBlock
            | _ -> placeWhole ()
            // A list running through the table carries on after it.
            match block with
            | Table tableBlock -> counters <- listCountersAfterTable counters tableBlock
            | _ -> ()

    let checkpoints = ResizeArray<PassCheckpoint>(blocks.Length + 1)
    let starts = ResizeArray<int>(blocks.Length + 1)
    let mutable blockIndex = 0
    // The blocks of the last pass that stand as they are: those before the
    // first block that differs - short of the paragraphs kept with it, whose
    // place depends on it - with the state the pass had there.
    let mutable convergeFrom = System.Int32.MaxValue
    let mutable indexShift = 0
    let mutable keptEndnotes = false
    match unchangedPass, lastPass with
    | Some record, _ ->
        restoreState record.PassCheckpoints.[blocks.Length]
        blockIndex <- blocks.Length
    | None, Some record when not balanced ->
        let struct (prefix, suffix) = commonBlockEnds record.PassBlocks blocks
        let mutable restart = prefix
        let keepsNext index =
            match tryGetParagraph blocks.[index] with
            | Some para -> WordStyles.keepNext (forBlockInContext styleContext blocks.[index]) para.Metadata
            | None -> false
        while restart > 0 && keepsNext (restart - 1) do
            restart <- restart - 1
        for index in 0 .. restart - 1 do
            checkpoints.Add record.PassCheckpoints.[index]
            starts.Add record.PassStarts.[index]
        for index in 0 .. record.PassStarts.[restart] - 1 do
            results.Add record.PassResults.[index]
        restoreState record.PassCheckpoints.[restart]
        blockIndex <- restart
        convergeFrom <- blocks.Length - suffix
        indexShift <- record.PassBlocks.Length - blocks.Length
        keptEndnotes <- record.PassHasEndnotes
    | _ -> ()
    let freshFrom = results.Count
    let keptPrefix = blockIndex
    let mutable keptSuffix = 0
    let mutable freshTo = -1
    while blockIndex < blocks.Length do
        let state = currentState ()
        match lastPass with
        | Some record when blockIndex >= convergeFrom && samePassState state record.PassCheckpoints.[blockIndex + indexShift] ->
            // From here on this pass is the last one: its layouts stand.
            freshTo <- results.Count
            keptSuffix <- blocks.Length - blockIndex
            let from = record.PassStarts.[blockIndex + indexShift]
            let shift = results.Count - from
            for index in blockIndex + indexShift .. record.PassBlocks.Length - 1 do
                checkpoints.Add record.PassCheckpoints.[index]
                starts.Add (record.PassStarts.[index] + shift)
            for index in from .. record.PassResults.Length - 1 do
                results.Add record.PassResults.[index]
            restoreState record.PassCheckpoints.[record.PassBlocks.Length]
            blockIndex <- blocks.Length
        | _ ->
            checkpoints.Add state
            starts.Add results.Count
            flowBlock blockIndex
            blockIndex <- blockIndex + 1
            match memo with
            | Some memo when slot <> "" -> memo.MemoBlocksLaidOut <- memo.MemoBlocksLaidOut + 1
            | _ -> ()
    match unchangedPass with
    | Some record ->
        match memo with
        | Some memo -> memo.MemoKeptPass <- true
        | None -> ()
        (if record.PassKeptPrefix = blocks.Length then record else { record with PassKeptPrefix = blocks.Length; PassKeptSuffix = 0 }), page, y, footnotesByPage
    | None ->
        if freshTo < 0 then freshTo <- results.Count
        checkpoints.Add (currentState ())
        starts.Add results.Count
        // (Kept layouts that refer to endnotes were found by the pass that made them. An edit that
        // removes the last reference leaves the flag set: the endnote pass then looks, and finds none.)
        let record = {
            PassSettings = settings
            PassColumns = Some columns
            PassStartPage = startPage
            PassStartY = startY
            PassPageTop = pageTopOf startPage
            PassNextPageTop = pageTopOf (startPage + 1)
            PassNonUniform = nonUniform
            PassBalanced = balanced
            PassInitial = initialState
            PassStyles = doc.Styles
            PassPageSetup = doc.PageSetup
            PassPackage = doc.Package
            PassBlocks = blocks
            PassCheckpoints = checkpoints.ToArray()
            PassStarts = starts.ToArray()
            PassResults = results.ToArray()
            PassHasEndnotes = keptEndnotes || resultsHaveEndnotes results freshFrom freshTo
            PassKeptPrefix = keptPrefix
            PassKeptSuffix = keptSuffix
        }
        match memo with
        | Some memo when slot <> "" ->
            memo.MemoPasses.[slot] <- record
            if freshFrom > 0 || freshTo < results.Count then memo.MemoKeptPass <- true
        | _ -> ()
        record, page, y, footnotesByPage

let private flowColumns (memo: ParagraphLayoutMemo option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (columns: SectionColumns) (pageTopOf: int -> float) (nonUniform: bool) (startPage: int) (startY: float) (balanced: bool) (initialFootnotes: Map<int, (string * string) list>) =
    let record, page, y, footnotes = flowColumnsKept "" memo settings measureText measureRange doc columns pageTopOf nonUniform startPage startY balanced initialFootnotes
    List.ofArray record.PassResults, page, y, footnotes

/// Absolute geometry of one rendered page. For single-section documents every
/// page is identical (uniform), matching the legacy stride-based model; with
/// per-section page setup the size/margins can differ page to page.
type PageInfo = {
    Top: float
    Width: float
    Height: float
    MarginTop: float
    MarginBottom: float
    MarginLeft: float
    MarginRight: float
}

/// Split the document into sections at paragraphs carrying a section break;
/// the final section uses the body PageSetup. Each entry is
/// (blocks, columns, geometry, breakType, isFinal).
let private splitIntoSections (doc: Document) =
    let result = ResizeArray<Block list * SectionColumns option * SectionPageSetup option * string option * bool>()
    let mutable current = ResizeArray<Block>()
    for block in doc.Blocks do
        current.Add block
        match tryGetParagraph block with
        | Some para when para.Metadata.SectionBreak.IsSome ->
            result.Add(resizeArrayToList current, para.Metadata.SectionColumns, para.Metadata.SectionPageSetup, para.Metadata.SectionBreak, false)
            current <- ResizeArray<Block>()
        | _ -> ()
    // The final section's start type comes from the body sectPr.
    result.Add(resizeArrayToList current, doc.PageSetup.Columns, None, doc.PageSetup.SectionStart, true)
    resizeArrayToList result

let private columnsOrSingle (columns: SectionColumns option) =
    match columns with
    | Some c when c.Count >= 2 -> c
    | _ -> { Count = 1; Gap = 0.0; Separator = false; Widths = [] }

let private settingsForGeometry (settings: LayoutSettings) (geometry: SectionPageSetup option) =
    match geometry with
    | Some g -> { settings with Width = g.PageWidth; PageHeight = g.PageHeight; PaddingLeft = g.MarginLeft; PaddingTop = g.MarginTop; MarginRight = g.MarginRight; MarginBottom = g.MarginBottom }
    | None -> settings

let private pageInfoOf (s: LayoutSettings) top =
    { Top = top; Width = s.Width; Height = pageHeight s; MarginTop = s.PaddingTop; MarginBottom = s.MarginBottom; MarginLeft = s.PaddingLeft; MarginRight = s.MarginRight }

/// Extend a page list so it reaches down to `bottom` (note areas below the
/// last body line), repeating the geometry of the last page.
let pageInfosCoveringBottom (bottom: float) (pages: PageInfo list) =
    match List.tryLast pages with
    | None -> pages
    | Some last ->
        let extra = ResizeArray<PageInfo>()
        let mutable top = last.Top
        while top + last.Height + pageGap < bottom do
            top <- top + last.Height + pageGap
            extra.Add { last with Top = top }
        if extra.Count = 0 then pages else pages @ resizeArrayToList extra

/// The uniform page list for a single-section document (legacy stride model).
let uniformPageInfosWithLayoutCount (settings: LayoutSettings) (layouts: BlockLayout list) =
    let stride = pageHeight settings + pageGap
    let bottom, layoutCount = maxLineBottomAndLayoutCount 0.0 layouts
    let pageCount = max 1 (int (floor (max 0.0 bottom / stride)) + 1)
    [ for p in 0 .. pageCount - 1 -> pageInfoOf settings (float p * stride) ], layoutCount

let uniformPageInfos (settings: LayoutSettings) (layouts: BlockLayout list) =
    uniformPageInfosWithLayoutCount settings layouts |> fst

/// The layouts of a section's pass at the place the section has in the
/// document: `source` moved down by `offset` (the top of the section's first
/// page). A section is flowed as if its first page were the document's
/// first, so its pass - and the record kept of it - does not depend on what
/// stands before the section: when an edit further up adds or removes a page
/// the sections below keep their passes and are only moved. `shift` moves
/// them sideways as well: a section with a left margin of its own stands
/// that much right (or left) of the document's left margin, which is what
/// painting and hit testing add to a layout's x. The layouts
/// moved to the same place are the same objects from pass to pass (those of
/// the blocks an edit in the section did not touch included); the second
/// result says whether any were kept.
let private sectionLayoutsAt (memo: ParagraphLayoutMemo option) (slot: string) (source: BlockLayout array) (offset: float) (shift: float) : BlockLayout array * bool =
    if offset = 0.0 && shift = 0.0 then source, true
    else
        let move (layout: BlockLayout) = layout |> shiftBlockY offset |> shiftBlockX shift
        let moveAll () = source |> Array.map move
        match memo with
        | None -> moveAll (), false
        | Some memo ->
            let keep placed =
                memo.MemoPlaced.[slot] <- { PlacedSource = source; PlacedOffset = offset; PlacedShift = shift; PlacedLayouts = placed }
            match memo.MemoPlaced.TryGetValue slot with
            | true, known when known.PlacedOffset = offset && known.PlacedShift = shift && obj.ReferenceEquals(known.PlacedSource, source) -> known.PlacedLayouts, true
            | true, known when known.PlacedOffset = offset && known.PlacedShift = shift ->
                // An edit in the section: the layouts it kept at the start and at the end are moved already.
                let before = known.PlacedSource
                let placed = Array.zeroCreate<BlockLayout> source.Length
                let limit = min source.Length before.Length
                let mutable prefix = 0
                while prefix < limit && obj.ReferenceEquals(source.[prefix], before.[prefix]) do
                    placed.[prefix] <- known.PlacedLayouts.[prefix]
                    prefix <- prefix + 1
                let mutable suffix = 0
                while suffix < limit - prefix && obj.ReferenceEquals(source.[source.Length - 1 - suffix], before.[before.Length - 1 - suffix]) do
                    placed.[source.Length - 1 - suffix] <- known.PlacedLayouts.[before.Length - 1 - suffix]
                    suffix <- suffix + 1
                for index in prefix .. source.Length - 1 - suffix do
                    placed.[index] <- move source.[index]
                keep placed
                placed, true
            | _ ->
                let placed = moveAll ()
                keep placed
                placed, false

/// Lay out a multi-section document section by section. A continuous break
/// continues on the same page below the previous section; otherwise the next
/// section starts on a new page (using that section's page size — supporting a
/// landscape page mid-document). Non-final multi-column sections are balanced.
/// Each section is flowed from y = 0 at the top of its first page and its
/// layouts are then moved to where that page stands (see sectionLayoutsAt).
/// Returns the layouts plus the per-page geometry.
let private layoutSections (memo: ParagraphLayoutMemo option) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (headerReferences: (string * string) list array) sections =
    let results = ResizeArray<BlockLayout>()
    // Where each block's layouts start in the results (see LayoutSlices).
    let blockStarts = ResizeArray<int>()
    let mutable keptPrefix = 0
    let mutable keptRun = true
    let mutable keptSuffix = 0
    let mutable hasEndnoteRefs = false
    let pageTops = ResizeArray<float>()    // absolute top of each page index
    let pageSettings = ResizeArray<LayoutSettings>()  // settings owning each page
    pageTops.Add 0.0
    let mutable currentPage = 0
    // Where the next section starts, below the top of the page it starts on.
    let mutable currentYInPage = 0.0
    let mutable started = false
    // Footnotes are per page, and a page can hold several (continuous) sections.
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let recordPage page (s: LayoutSettings) top =
        while pageTops.Count <= page do pageTops.Add 0.0
        while pageSettings.Count <= page do pageSettings.Add s
        pageTops.[page] <- top
        pageSettings.[page] <- s
    // The header each section shows (its own, or the one of the section before it).
    let mutable shownHeader : (string * string) list = []
    sections
    |> List.iteri (fun i (blocks, columns, geometry, breakType, isFinal) ->
        shownHeader <- inheritSectionReferences shownHeader (if i < headerReferences.Length then headerReferences.[i] else [])
        let secSettings = settingsForGeometry settings geometry
        // The body starts below the section's header when that reaches below the top margin.
        let secSettings =
            match WordLayout.Page.referencedHeaderBlocks "default" shownHeader doc with
            | Some headerBlocks ->
                let marginTop, headerDistance =
                    match geometry with
                    | Some g -> g.MarginTop, g.HeaderDistance
                    | None -> doc.PageSetup.MarginTop, doc.PageSetup.HeaderDistance
                let top = WordLayout.Page.bodyTopBelowHeader marginTop headerDistance headerBlocks
                if top = secSettings.PaddingTop then secSettings else { secSettings with PaddingTop = top }
            | None -> secSettings
        let secStride = pageHeight secSettings + pageGap
        // Stitch before this section.
        if not started then
            currentPage <- 0
            recordPage 0 secSettings 0.0
            currentYInPage <- WordLayout.Page.contentTop secSettings.PaddingTop 0.0
            started <- true
        else
            match breakType with
            | Some "continuous" -> ()  // same page, continue at currentY
            | _ ->
                // New page: top = previous page top + previous page height + gap.
                let prevSettings = pageSettings.[currentPage]
                let newTop = pageTops.[currentPage] + pageHeight prevSettings + pageGap
                currentPage <- currentPage + 1
                recordPage currentPage secSettings newTop
                currentYInPage <- WordLayout.Page.contentTop secSettings.PaddingTop 0.0
        // Page tops within this section are uniform (secStride from the section top).
        let sectionBase = currentPage
        let sectionBaseTop = pageTops.[currentPage]
        let pageTopOf p = sectionBaseTop + float (p - sectionBase) * secStride
        let secColumns = columnsOrSingle columns
        let secDoc = { doc with Blocks = blocks; PageSetup = { doc.PageSetup with Columns = columns } }
        // The columns of a section are balanced when the section after it goes on on the same
        // page (a continuous break), as in Word; before a break to a new page the text fills
        // one column after the other.
        let nextContinues =
            match List.tryItem (i + 1) sections with
            | Some (_, _, _, Some "continuous", _) -> true
            | _ -> false
        let balanced = (not isFinal) && secColumns.Count >= 2 && nextContinues
        let slot = match memo with Some _ -> "s" + string i | None -> ""
        // The section's own pages are numbered from 0, the first at y = 0: of what stands
        // before the section its pass is given where on its first page it starts and the
        // footnotes that page has already.
        let ownTopOf p = float p * secStride
        let ownFootnotes =
            match Map.tryFind sectionBase footnotesByPage with
            | Some notes when not notes.IsEmpty -> Map.empty |> Map.add 0 notes
            | _ -> Map.empty
        let record, ownEndPage, ownEndY, secFootnotes = flowColumnsKept slot memo secSettings measureText measureRange secDoc secColumns ownTopOf true 0 currentYInPage balanced ownFootnotes
        for KeyValue (ownPage, notes) in secFootnotes do
            footnotesByPage <- footnotesByPage |> Map.add (sectionBase + ownPage) notes
        let endPage = sectionBase + ownEndPage
        // (On screen a block stands at the document's left margin plus its layout's x: the layouts of a
        // section with a left margin of its own are moved by the difference.)
        let placed, placedBefore = sectionLayoutsAt memo slot record.PassResults sectionBaseTop (secSettings.PaddingLeft - settings.PaddingLeft)
        // (Blocks kept from the last pass, from the start of the document and from its end.
        // A section that was moved has none: its layouts are other objects.)
        let count = record.PassBlocks.Length
        let sectionKeptPrefix = if placedBefore then record.PassKeptPrefix else 0
        let sectionKeptSuffix = if placedBefore then record.PassKeptSuffix else 0
        if keptRun then
            keptPrefix <- keptPrefix + sectionKeptPrefix
            if sectionKeptPrefix < count then keptRun <- false
        if sectionKeptPrefix >= count then keptSuffix <- keptSuffix + count
        else keptSuffix <- sectionKeptSuffix
        for start in 0 .. record.PassBlocks.Length - 1 do
            blockStarts.Add (results.Count + record.PassStarts.[start])
        if record.PassHasEndnotes then hasEndnoteRefs <- true
        results.AddRange placed
        // Record the geometry of every page this section used.
        for p in sectionBase .. endPage do recordPage p secSettings (pageTopOf p)
        currentPage <- endPage
        currentYInPage <- ownEndY - ownTopOf ownEndPage)
    let maxPage = pageTops.Count - 1
    let bodyLayouts = results.ToArray()
    blockStarts.Add bodyLayouts.Length
    let noteLayouts = ResizeArray<BlockLayout>()
    for KeyValue (pageIndex, notes) in footnotesByPage do
        if not notes.IsEmpty && pageIndex <= maxPage && pageIndex < pageSettings.Count then
            let s = pageSettings.[pageIndex]
            let contentBottom = WordLayout.Page.contentBottom (pageHeight s) s.MarginBottom pageTops.[pageIndex]
            let own = footnoteLayoutsForPageKept "s" memo s measureText measureRange (s.Width - s.PaddingLeft - s.MarginRight) pageIndex contentBottom notes
            // (Moved sideways as the text of their section is; the same objects while nothing changed.)
            let shift = s.PaddingLeft - settings.PaddingLeft
            let placed =
                if shift = 0.0 then own
                else
                    match memo with
                    | None -> own |> List.map (shiftBlockX shift)
                    | Some memo ->
                        match memo.MemoPlacedNotes.TryGetValue pageIndex with
                        | true, (source, knownShift, known) when knownShift = shift && obj.ReferenceEquals(source, own) -> known
                        | _ ->
                            let moved = own |> List.map (shiftBlockX shift)
                            memo.MemoPlacedNotes.[pageIndex] <- (own, shift, moved)
                            moved
            noteLayouts.AddRange placed
    results.AddRange noteLayouts
    let pages = [ for p in 0 .. maxPage -> pageInfoOf pageSettings.[p] pageTops.[p] ]
    // Endnotes go after the last section's text, in its page geometry; pages
    // they spill onto repeat the last page.
    let lastSettings = pageSettings.[maxPage]
    let lastStride = pageHeight lastSettings + pageGap
    let pageAt noteY =
        let mutable index = 0
        for p in 0 .. maxPage do
            if pageTops.[p] <= noteY + 0.5 then index <- p
        if index < maxPage then
            let pageSetting = pageSettings.[index]
            struct (WordLayout.Page.contentTop pageSetting.PaddingTop pageTops.[index],
                    WordLayout.Page.contentBottom (pageHeight pageSetting) pageSetting.MarginBottom pageTops.[index],
                    WordLayout.Page.contentTop pageSettings.[index + 1].PaddingTop pageTops.[index + 1])
        else
            let top = pageTops.[maxPage] + floor (max 0.0 (noteY - pageTops.[maxPage]) / lastStride) * lastStride
            struct (WordLayout.Page.contentTop lastSettings.PaddingTop top,
                    WordLayout.Page.contentBottom (pageHeight lastSettings) lastSettings.MarginBottom top,
                    WordLayout.Page.contentTop lastSettings.PaddingTop (top + lastStride))
    let endnotes =
        // (No body layout refers to an endnote: there is nothing to collect, sort or place.)
        if not hasEndnoteRefs then []
        else
            let mutable bodyBottom = 0.0
            for layout in bodyLayouts do
                bodyBottom <- max bodyBottom (layout.Y + layout.Height)
            endnoteLayoutsAfterBody lastSettings measureText measureRange (lastSettings.Width - lastSettings.PaddingLeft - lastSettings.MarginRight) bodyLayouts bodyBottom pageAt
    results.AddRange endnotes
    noteLayouts.AddRange endnotes
    match memo with
    | Some memo ->
        memo.MemoSlices <-
            Some { SliceBlocks = List.toArray doc.Blocks; SliceStarts = blockStarts.ToArray(); SliceLayouts = bodyLayouts; SliceNotes = resizeArrayToList noteLayouts
                   SliceGeneration = 0; SliceBase = 0; SliceKeptPrefix = keptPrefix; SliceKeptSuffix = (if keptRun then 0 else keptSuffix) }
    | None -> ()
    let pages =
        match endnotes with
        | [] -> pages
        | _ -> pageInfosCoveringBottom (maxLineBottom 0.0 endnotes) pages
    resizeArrayToList results, pages

/// The first `limit` blocks of a list of sections: the section the limit
/// falls in is cut short and keeps what it is laid out with (its own page
/// geometry, columns and start), the sections after it are left out.
let private sectionsUpTo (limit: int) (sections: (Block list * SectionColumns option * SectionPageSetup option * string option * bool) list) =
    let kept = ResizeArray<Block list * SectionColumns option * SectionPageSetup option * string option * bool>()
    let mutable left = max 1 limit
    for (blocks, columns, geometry, breakType, isFinal) in sections do
        if left > 0 then
            let count = List.length blocks
            if count <= left then
                kept.Add (blocks, columns, geometry, breakType, isFinal)
                left <- left - count
            else
                kept.Add (List.truncate left blocks, columns, geometry, breakType, isFinal)
                left <- 0
    resizeArrayToList kept

/// Lay out a document, returning the layouts and the per-page geometry.
/// `limit`: the number of top-level blocks to lay out (System.Int32.MaxValue:
/// the whole document) - the start of a long document is laid out and shown
/// before the rest, see IncrementalLayoutState.BlockLimit. What is laid out of
/// a part is laid out as it is in the whole, up to the blocks around the cut
/// (a paragraph kept with the next, a balanced section), which a pass with a
/// larger limit lays out again.
let private layoutDocumentPagesWithMemo (memo: ParagraphLayoutMemo option) (limit: int) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (whole: Document) : BlockLayout list * PageInfo list =
    let allSections = splitIntoSections whole
    let limited = limit < System.Int32.MaxValue
    let doc = if limited then { whole with Blocks = List.truncate (max 1 limit) whole.Blocks } else whole
    let sections = if limited then sectionsUpTo limit allSections else allSections
    match allSections with
    | [ (_, columns, _, _, _) ] ->
        // Single section: keep the existing path; uniform page list.
        let layouts =
            match columns with
            | Some c when c.Count >= 2 ->
                let stride = pageHeight settings + pageGap
                let r, _, _, footnotesByPage = flowColumns memo settings measureText measureRange doc c (fun p -> float p * stride) false 0 (WordLayout.Page.contentTop settings.PaddingTop 0.0) false Map.empty
                let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
                let footnotes =
                    [ for KeyValue (pageIndex, notes) in footnotesByPage do
                        if not notes.IsEmpty then
                            let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom (float pageIndex * stride)
                            yield! footnoteLayoutsForPage settings measureText measureRange maxWidth pageIndex contentBottom notes ]
                // Endnotes after the columns' text, across the full text width.
                let mutable bodyBottom = settings.PaddingTop
                for layout in r do
                    bodyBottom <- max bodyBottom (layout.Y + layout.Height)
                let pageAt noteY =
                    let pageIndex = int (floor (noteY / stride))
                    let pageTop = float pageIndex * stride
                    struct (WordLayout.Page.contentTop settings.PaddingTop pageTop,
                            WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop,
                            WordLayout.Page.contentTop settings.PaddingTop (float (pageIndex + 1) * stride))
                r @ footnotes @ endnoteLayoutsAfterBody settings measureText measureRange maxWidth r bodyBottom pageAt
            | _ ->
                let slot name = match memo with Some _ -> name | None -> ""
                let keepSlices (record: PassRecord) notes =
                    match memo with
                    | Some memo ->
                        memo.MemoSlices <-
                            Some { SliceBlocks = record.PassBlocks; SliceStarts = record.PassStarts; SliceLayouts = record.PassResults; SliceNotes = notes
                                   SliceGeneration = 0; SliceBase = 0; SliceKeptPrefix = record.PassKeptPrefix; SliceKeptSuffix = record.PassKeptSuffix }
                    | None -> ()
                let firstPass, placedAfterFirst, firstRecord, firstNotes =
                    layoutDocumentPassKept (slot "d1") memo settings measureText measureRange doc Map.empty [] []
                let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
                let extendedObstacles, prePlacedBoxes =
                    // (No floating object was placed: none to extend.)
                    if placedAfterFirst.IsEmpty then [], []
                    else computeAnchorExtendedObstacles maxWidth firstPass placedAfterFirst
                if extendedObstacles.IsEmpty then
                    keepSlices firstRecord firstNotes
                    firstPass
                else
                    let secondPass, _, secondRecord, secondNotes =
                        layoutDocumentPassKept (slot "d2") memo settings measureText measureRange doc Map.empty extendedObstacles prePlacedBoxes
                    keepSlices secondRecord secondNotes
                    secondPass
        layouts, uniformPageInfos settings layouts
    | _ ->
        // (What each section says of its header: from the whole document, whatever part of it is laid out.)
        let headerReferences =
            allSections
            |> List.map (fun (blocks, _, _, _, isFinal) ->
                if isFinal then whole.PageSetup.HeaderReferences
                else
                    match List.tryLast blocks |> Option.bind tryGetParagraph with
                    | Some para -> para.Metadata.HeaderReferences
                    | None -> [])
            |> List.toArray
        layoutSections memo settings measureText measureRange doc headerReferences sections

/// True when the document has more than one section, i.e. its pages need not
/// all have the body PageSetup's geometry.
let documentHasSections (doc: Document) =
    doc.Blocks
    |> List.exists (fun block ->
        match tryGetParagraph block with
        | Some para -> para.Metadata.SectionBreak.IsSome
        | None -> false)

/// Index of the page a document Y falls on (the last page starting at or above it).
let pageIndexForYInPages (pages: PageInfo list) (y: float) =
    let mutable index = 0
    let mutable current = 0
    for page in pages do
        if page.Top <= y + 0.5 then index <- current
        current <- current + 1
    index

/// Lay out a document, returning the layouts and the per-page geometry.
let layoutDocumentPages (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) : BlockLayout list * PageInfo list =
    layoutDocumentPagesWithMemo None System.Int32.MaxValue settings measureText measureRange doc

/// The layout of the first `limit` top-level blocks of a document, on its
/// own (see IncrementalLayoutState.BlockLimit, which gives the same).
let layoutDocumentPagesUpTo (limit: int) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) : BlockLayout list * PageInfo list =
    layoutDocumentPagesWithMemo None limit settings measureText measureRange doc

let layoutDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) =
    layoutDocumentPages settings measureText measureRange doc |> fst

let layoutDocument (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) =
    layoutDocumentWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) doc

/// State of an incremental layout: the last result plus the paragraph memo
/// that makes the next pass cheap.
type IncrementalLayoutState = {
    mutable LastDocument: Document option
    mutable LastSettings: LayoutSettings option
    mutable LastLayouts: BlockLayout list option
    mutable LastPages: PageInfo list option
    /// The body layouts of the last pass block by block, when the pass knows them.
    mutable LastSlices: LayoutSlices option
    mutable MeasuredBlockCount: int
    mutable MemoSweepThreshold: int
    /// How many top-level blocks are laid out: System.Int32.MaxValue for the
    /// whole document, less while the start of a long document is shown and
    /// the rest is still to come (a host raises it in slices; the layouts of
    /// the blocks laid out so far stand from one limit to the next, as they do
    /// from one edit to the next). The result with the limit at the whole
    /// document is the layout a pass over the whole document makes.
    mutable BlockLimit: int
    /// The limit the last result was made with, and whether it covered the document.
    mutable LastLimit: int
    mutable LastWasWhole: bool
    /// Counts the passes that made a result (a result handed back as it was is not one).
    mutable LayoutSerial: int
    Memo: ParagraphLayoutMemo
}

let createIncrementalLayoutState () = {
    LastDocument = None
    LastSettings = None
    LastLayouts = None
    LastPages = None
    LastSlices = None
    MeasuredBlockCount = 0
    MemoSweepThreshold = 256
    BlockLimit = System.Int32.MaxValue
    LastLimit = System.Int32.MaxValue
    LastWasWhole = true
    LayoutSerial = 0
    Memo = createParagraphLayoutMemo ()
}

let resetIncrementalLayoutMeasurements (state: IncrementalLayoutState) =
    state.MeasuredBlockCount <- 0

/// Number of paragraphs whose lines had to be broken afresh since the last
/// reset (everything else was reused from the memo).
let incrementalLayoutMeasuredBlockCount (state: IncrementalLayoutState) =
    state.MeasuredBlockCount

/// Number of top-level blocks the last layout pass laid out; the other
/// blocks kept the layouts of the pass before.
let incrementalLayoutBlocksLaidOut (state: IncrementalLayoutState) =
    state.Memo.MemoBlocksLaidOut

/// Whether the last result covers the whole document (see IncrementalLayoutState.BlockLimit).
let incrementalLayoutIsWhole (state: IncrementalLayoutState) =
    state.BlockLimit = System.Int32.MaxValue || (state.LastDocument.IsSome && state.LastWasWhole && state.BlockLimit >= state.LastLimit)

/// Whether the last result is what the state would hand back for the limit it is at.
let incrementalLayoutIsCurrent (state: IncrementalLayoutState) =
    state.LastDocument.IsSome && (state.LastLimit = state.BlockLimit || (state.LastWasWhole && state.BlockLimit >= state.LastLimit))

let clearIncrementalLayoutState (state: IncrementalLayoutState) =
    state.LastDocument <- None
    state.LastSettings <- None
    state.LastLayouts <- None
    state.LastPages <- None
    state.LastSlices <- None
    state.MeasuredBlockCount <- 0
    state.MemoSweepThreshold <- 256
    state.Memo.MemoPasses.Clear()
    state.Memo.MemoPlaced.Clear()
    state.Memo.MemoPlacedNotes.Clear()
    state.Memo.MemoFootnoteAreas.Clear()
    state.Memo.MemoSlices <- None
    state.Memo.MemoEntries.Clear()
    state.Memo.MemoContinuations.Clear()
    state.Memo.MemoStyles <- []
    state.Memo.MemoStyleContext <- None

/// The measurers measure text in some faces otherwise from now on (the host
/// got the files of faces it had drawn in a fallback): `uses style` says
/// whether text of that style is in one of them; all other text measures as
/// it did. What was laid out in those faces is forgotten - the lines and the
/// placements of the paragraphs set in them, and the place they have in the
/// passes that were recorded, so that the next pass lays them out again as it
/// does paragraphs an edit changed - and everything else stands: the lines of
/// the other paragraphs, the layouts before the first paragraph in such a
/// face and after the last one whose place did not move. Returns how many
/// paragraphs were set in the faces.
let invalidateFacesInIncrementalLayout (state: IncrementalLayoutState) (uses: TextBuffer.TextStyle -> bool) : int =
    let memo = state.Memo
    let lineUses (line: MeasuredLine) =
        line.Segments
        |> List.exists (function
            | TextSegment (_, style, _, _, _) -> uses style
            | MathSegment (_, style, _, _, _, _, _, _) -> uses style
            | ImageSegment _ -> false)
    let layoutUses (layout: BlockLayout) =
        (match layout.PrefixStyle with
         | Some style -> uses style
         | None -> false)
        || List.exists lineUses layout.Lines
    let affected = System.Collections.Generic.HashSet<string>()
    for KeyValue (key, entry) in memo.MemoEntries do
        let inFaces =
            uses entry.MemoContentTextStyle
            || (entry.MemoEffectiveBuffer.Runs |> List.exists (function TextBuffer.Text run -> uses run.Style | TextBuffer.Image _ -> false))
            || (match entry.MemoLines with Some (_, lines) -> List.exists lineUses lines | None -> false)
            || (match entry.MemoLinesOlder with Some (_, lines) -> List.exists lineUses lines | None -> false)
            || (match entry.MemoPlacement with Some (_, layout) -> layoutUses layout | None -> false)
            || (match entry.MemoPlacementOlder with Some (_, layout) -> layoutUses layout | None -> false)
            || (match entry.MemoObstaclePlacement with Some (_, _, _, layout, _, _) -> layoutUses layout | None -> false)
        if inFaces then
            let separator = key.IndexOf '\u0001'
            affected.Add (if separator < 0 then key else key.Substring(0, separator)) |> ignore
    // (Every slot of a paragraph that has one in the faces: the whole and the rest of one that breaks over a page.)
    for KeyValue (key, entry) in memo.MemoEntries do
        let separator = key.IndexOf '\u0001'
        if affected.Contains (if separator < 0 then key else key.Substring(0, separator)) then
            entry.MemoLines <- None
            entry.MemoLinesOlder <- None
            entry.MemoPlacement <- None
            entry.MemoPlacementOlder <- None
            entry.MemoObstaclePlacement <- None
    let rec tableAffected (table: TableBlock) =
        table.Rows
        |> List.exists (fun row ->
            row.Cells
            |> List.exists (fun cell ->
                (cell.Content |> List.exists (fun para -> affected.Contains para.Id))
                || List.exists tableAffected cell.NestedTables))
    // (The same block as another object: a pass that compares its blocks with those of the last
    // one takes it for a block that changed, and reads of it what it would have read.)
    let another (block: Block) : Block =
        match block with
        | Paragraph para -> Paragraph { para with Id = para.Id }
        | Heading (level, para) -> Heading (level, { para with Id = para.Id })
        | Table table -> Table { table with Id = table.Id }
        | TextBox textBox -> TextBox { textBox with Id = textBox.Id }
    let mutable changed = 0
    for record in memo.MemoPasses.Values do
        for index in 0 .. record.PassBlocks.Length - 1 do
            let block = record.PassBlocks.[index]
            let byContent =
                match block with
                | Paragraph para
                | Heading (_, para) -> affected.Contains para.Id
                | Table table -> tableAffected table
                | TextBox _ -> false
            let byResult () =
                let mutable found = false
                let mutable at = record.PassStarts.[index]
                while not found && at < record.PassStarts.[index + 1] do
                    if layoutUses record.PassResults.[at] then found <- true
                    at <- at + 1
                found
            if byContent || byResult () then
                record.PassBlocks.[index] <- another block
                changed <- changed + 1
    // What is cheap to make again and was made of text in any face: the note areas.
    memo.MemoFootnoteAreas.Clear()
    memo.MemoPlacedNotes.Clear()
    memo.MemoSlices <- None
    // The last result is not what the next pass gives.
    state.LastDocument <- None
    state.LastLayouts <- None
    state.LastPages <- None
    affected.Count + changed

/// True when two documents lay out (and paint, the selection overlay apart)
/// alike without looking further: they are the same document or differ only
/// in the selection - a caret move, a selection drag. Everything else is the
/// same object in both.
let sameDocumentApartFromSelection (left: Document) (right: Document) =
    obj.ReferenceEquals(left, right)
    || (obj.ReferenceEquals(left.Blocks, right.Blocks)
        && obj.ReferenceEquals(left.PageSetup, right.PageSetup)
        && obj.ReferenceEquals(left.Styles, right.Styles)
        && obj.ReferenceEquals(left.Headers, right.Headers)
        && obj.ReferenceEquals(left.Footers, right.Footers)
        && obj.ReferenceEquals(left.HeaderParts, right.HeaderParts)
        && obj.ReferenceEquals(left.FooterParts, right.FooterParts)
        && obj.ReferenceEquals(left.Package, right.Package))

/// Incremental document layout. Runs the same document pass as
/// layoutDocumentPages (so the result is identical for every document —
/// pagination, numbering, tables, footnotes, floats, sections and columns
/// included) handing it this state's paragraph memo, so only paragraphs whose own
/// inputs changed are measured and paginated again.
let layoutDocumentPagesIncremental (state: IncrementalLayoutState) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) : BlockLayout list * PageInfo list =
    // (A result that covered the whole document stands whatever the limit was raised to since.)
    let sameLimit = state.LastLimit = state.BlockLimit || (state.LastWasWhole && state.BlockLimit >= state.LastLimit)
    match state.LastDocument, state.LastSettings, state.LastLayouts, state.LastPages with
    | Some cachedDoc, Some cachedSettings, Some layouts, Some pages
        when sameLimit && obj.ReferenceEquals(cachedDoc, doc) && cachedSettings = settings -> layouts, pages
    | Some cachedDoc, Some cachedSettings, Some layouts, Some pages
        when sameLimit && sameDocumentApartFromSelection cachedDoc doc && cachedSettings = settings ->
        // The selection has no part in the layout: a caret move keeps it.
        state.LastDocument <- Some doc
        layouts, pages
    | _ ->
        let memo = state.Memo
        if not (obj.ReferenceEquals(memo.MemoStyles, doc.Styles)) then
            // Named (character) styles feed every paragraph's effective runs.
            if memo.MemoStyles <> doc.Styles then
                memo.MemoEntries.Clear()
                memo.MemoContinuations.Clear()
                memo.MemoPasses.Clear()
                memo.MemoPlaced.Clear()
                memo.MemoPlacedNotes.Clear()
                memo.MemoFootnoteAreas.Clear()
            memo.MemoStyles <- doc.Styles
            memo.MemoStyleContext <- None
        memo.Generation <- memo.Generation + 1
        memo.LineBreakCount <- 0
        memo.MemoKeptPass <- false
        memo.MemoBlocksLaidOut <- 0
        memo.MemoSlices <- None
        // A limit at or past the last block is no limit.
        let limit =
            if state.BlockLimit = System.Int32.MaxValue then state.BlockLimit
            else
                let mutable count = 0
                let mutable rest = doc.Blocks
                while count < state.BlockLimit && not rest.IsEmpty do
                    count <- count + 1
                    rest <- rest.Tail
                if rest.IsEmpty then System.Int32.MaxValue else state.BlockLimit
        let layouts, pages = layoutDocumentPagesWithMemo (Some memo) limit settings measureText measureRange doc
        // (The whole document laid out: blocks an edit adds are laid out with it.)
        state.BlockLimit <- limit
        state.LastLimit <- limit
        state.LastWasWhole <- (limit = System.Int32.MaxValue)
        state.LayoutSerial <- state.LayoutSerial + 1
        state.MeasuredBlockCount <- state.MeasuredBlockCount + memo.LineBreakCount
        // (Numbered, with the number of the pass before: what was kept was kept from that one.)
        let before = match state.LastSlices with Some slices -> slices.SliceGeneration | None -> 0
        state.LastSlices <- memo.MemoSlices |> Option.map (fun slices -> { slices with SliceGeneration = nextSliceGeneration (); SliceBase = before })
        // Drop entries of paragraphs that no longer exist once they pile up. (A pass
        // that kept layouts of the last one did not visit the paragraphs they belong
        // to: which entries are stale is known after a pass over the whole document.)
        if memo.MemoEntries.Count > state.MemoSweepThreshold && not memo.MemoKeptPass then
            let stale = ResizeArray<string>()
            for KeyValue (id, entry) in memo.MemoEntries do
                if entry.MemoGeneration <> memo.Generation then stale.Add id
            for id in stale do
                memo.MemoEntries.Remove id |> ignore
            if memo.MemoContinuations.Count > 4096 then memo.MemoContinuations.Clear()
            state.MemoSweepThreshold <- memo.MemoEntries.Count * 2 + 256
        state.LastDocument <- Some doc
        state.LastSettings <- Some settings
        state.LastLayouts <- Some layouts
        state.LastPages <- Some pages
        layouts, pages

let layoutDocumentIncrementalWithRange (state: IncrementalLayoutState) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) =
    layoutDocumentPagesIncremental state settings measureText measureRange doc |> fst

let layoutDocumentIncremental state settings measureText doc =
    layoutDocumentIncrementalWithRange state settings measureText (rangeMeasurerFromTextMeasurer measureText) doc

let hitTest (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (mouseX: float) (mouseY: float) =
    let layouts = layoutDocument settings measureText doc
    hitTestLayouts measureText settings.PaddingLeft layouts mouseX mouseY

let hitTestWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (mouseX: float) (mouseY: float) =
    let layouts = layoutDocumentWithRange settings measureText measureRange doc
    hitTestLayoutsWithRange measureRange settings.PaddingLeft layouts mouseX mouseY

let hitTestForImage (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let layouts = layoutDocument settings measureText doc
    hitTestForImageLayouts settings.PaddingLeft layouts mouseX mouseY

/// Hit tests against the incremental layout: an unchanged document reuses the
/// layout the last render produced instead of laying the document out again.
let hitTestIncrementalWithRange (state: IncrementalLayoutState) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (mouseX: float) (mouseY: float) =
    let layouts = layoutDocumentIncrementalWithRange state settings measureText measureRange doc
    hitTestLayoutsWithRange measureRange settings.PaddingLeft layouts mouseX mouseY

let hitTestForImageIncremental (state: IncrementalLayoutState) (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let layouts = layoutDocumentIncrementalWithRange state settings measureText measureRange doc
    hitTestForImageLayouts settings.PaddingLeft layouts mouseX mouseY
