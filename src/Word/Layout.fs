module RedBat.Word.Layout

open RedBat.Word.Document
open RedBat.Word.TextBuffer
open RedBat.Word.WordStyles

#if FABLE_COMPILER
open Fable.Core

[<Emit("try { if (new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1') console.log(new Date().toISOString() + ' word-layout:' + $0); } catch (_) {}")>]
let private trace (_label: string) : unit = jsNative

[<Emit("(function(){ try { return new URLSearchParams(window.location.search).has('rbPdfExportTrace') || window.localStorage.getItem('rb.word.pdfExportTrace') === '1'; } catch (_) { return false; } })()")>]
let private traceTimingEnabled : bool = jsNative

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

module Eq = RedBat.Word.Math

/// Segment can be text, image, or a laid-out math equation (OMML).
type Segment =
    | TextSegment of text: string * style: TextBuffer.TextStyle * x: float * width: float * startOffset: int
    | ImageSegment of src: string * x: float * yOffset: float * width: float * height: float * startOffset: int * floating: bool * wrap: string option * relativeHeight: string option * cropLeft: string option * cropTop: string option * cropRight: string option * cropBottom: string option
    /// A rendered equation: pre-laid-out primitives (relative to x and the line
    /// baseline), the run style (colour/font), box metrics, and the run's offset span.
    | MathSegment of prims: Eq.MathPrim list * style: TextBuffer.TextStyle * x: float * width: float * ascent: float * descent: float * startOffset: int * length: int

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
        let renderedBreakBefore = lineHasRenderedPageBreakBefore line
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

let rangeMeasurerFromTextMeasurer (measureText: TextMeasurer) : TextRangeMeasurer =
    fun text style start length ->
        let startIndex = max 0 (min text.Length start)
        let endIndex = max startIndex (min text.Length (start + length))
        if endIndex = startIndex then 0.0
        else measureText (text.Substring(startIndex, endIndex - startIndex)) style

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
    let automatic =
        if settings.InTableCell then
            let factor = tableCellLineBoxFactorForTextStyle hasExplicitLineSpacing textStyle
            tableCellLineBoxFontSize fontSize * lineSpacing * factor
        else
            wordLineHeight textStyle fontSize lineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox

    match lineSpacingRule with
    | "exact" -> max 1.0 lineSpacing
    | "atLeast" -> max automatic lineSpacing
    | _ -> automatic

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

type private LineArrayCacheEntry = {
    Lines: MeasuredLine list
    LineArray: MeasuredLine array
    TextCharCount: int
    mutable LastLineSet: bool
    mutable LastOffset: int
    mutable LastLine: MeasuredLine
}

let mutable private lineArrayCache: LineArrayCacheEntry list = []
let mutable private lastLineArrayCacheEntry: LineArrayCacheEntry option = None
let private lineArrayCacheLinesComparer =
    { new System.Collections.Generic.IEqualityComparer<MeasuredLine list> with
        member _.Equals(left, right) = obj.ReferenceEquals(left, right)
        member _.GetHashCode(value) = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value) }
let private lineArrayCacheByLines =
    System.Collections.Generic.Dictionary<MeasuredLine list, LineArrayCacheEntry>(lineArrayCacheLinesComparer)
let private lineArrayCacheThreshold = 32
let private lineArrayCacheMaxEntries = 64
let private lineArrayCacheMaxTotalLines = 200000
let private lineArrayCacheMaxTotalTextChars = 5000000

let private lineListExceeds threshold lines =
    let rec loop count remaining =
        count > threshold ||
            match remaining with
            | [] -> false
            | _ :: tail -> loop (count + 1) tail
    loop 0 lines

let private lineArrayCacheEntryMatches lines (entry: LineArrayCacheEntry) =
    obj.ReferenceEquals(entry.Lines, lines)

let private setLineArrayCache entries =
    lineArrayCache <- entries
    lineArrayCacheByLines.Clear()
    for entry in entries do
        lineArrayCacheByLines.[entry.Lines] <- entry
    match lastLineArrayCacheEntry with
    | Some last ->
        let mutable cached = Unchecked.defaultof<LineArrayCacheEntry>
        if not (lineArrayCacheByLines.TryGetValue(last.Lines, &cached) && obj.ReferenceEquals(cached, last)) then
            lastLineArrayCacheEntry <- None
    | None -> ()

let private tryFindLineArrayCacheEntry lines =
    match lastLineArrayCacheEntry with
    | Some entry when lineArrayCacheEntryMatches lines entry -> Some entry
    | _ ->
        let mutable cached = Unchecked.defaultof<LineArrayCacheEntry>
        if lineArrayCacheByLines.TryGetValue(lines, &cached)
           && lineArrayCacheEntryMatches lines cached then
            lastLineArrayCacheEntry <- Some cached
            Some cached
        else
            None

let private lineArrayCacheOptionTextCharCount (value: string option) =
    match value with
    | Some value when not (isNull value) -> value.Length
    | _ -> 0

let private lineArrayCacheAddCapped value total =
    if total > lineArrayCacheMaxTotalTextChars then
        total
    else
        let value = max 0 value
        let remaining = lineArrayCacheMaxTotalTextChars - total
        if value > remaining then lineArrayCacheMaxTotalTextChars + 1
        else total + value

let private lineArrayCacheTextStyleStringCharCount (style: TextStyle) =
    let mutable total = 0
    let add (value: string option) =
        total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount value) total
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

let private lineArrayCacheMathPrimsTextCharCount (prims: Eq.MathPrim list) =
    let mutable total = 0
    let mutable seen = 0
    let mutable remaining = prims
    while total <= lineArrayCacheMaxTotalTextChars
          && seen <= lineArrayCacheMaxTotalLines
          && not remaining.IsEmpty do
        match remaining.Head with
        | Eq.MGlyph (text, _, _, _, _) when not (isNull text) ->
            total <- lineArrayCacheAddCapped text.Length total
        | _ -> ()
        seen <- seen + 1
        remaining <- remaining.Tail
    if seen > lineArrayCacheMaxTotalLines then
        lineArrayCacheMaxTotalTextChars + 1
    else
        total

let private lineArrayCacheImageTextCharCount
    (src: string)
    (wrap: string option)
    (relativeHeight: string option)
    (cropLeft: string option)
    (cropTop: string option)
    (cropRight: string option)
    (cropBottom: string option) =
    let mutable total = if isNull src then 0 else src.Length
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount wrap) total
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount relativeHeight) total
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount cropLeft) total
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount cropTop) total
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount cropRight) total
    total <- lineArrayCacheAddCapped (lineArrayCacheOptionTextCharCount cropBottom) total
    total

let private lineArrayCacheSegmentTextCharCount segment =
    match segment with
    | TextSegment (text, style, _, _, _) ->
        let mutable total = if isNull text then 0 else text.Length
        total <- lineArrayCacheAddCapped (lineArrayCacheTextStyleStringCharCount style) total
        total
    | ImageSegment (src, _, _, _, _, _, _, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
        lineArrayCacheImageTextCharCount src wrap relativeHeight cropLeft cropTop cropRight cropBottom
    | MathSegment (prims, style, _, _, _, _, _, _) ->
        let mutable total = lineArrayCacheTextStyleStringCharCount style
        total <- lineArrayCacheAddCapped (lineArrayCacheMathPrimsTextCharCount prims) total
        total

let private lineArrayCacheTextCharCount (lines: MeasuredLine list) =
    let limit = lineArrayCacheMaxTotalTextChars + 1
    let mutable total = 0
    let mutable remainingLines = lines
    while total <= lineArrayCacheMaxTotalTextChars
          && not remainingLines.IsEmpty do
        let mutable remainingSegments = remainingLines.Head.Segments
        while total <= lineArrayCacheMaxTotalTextChars
              && not remainingSegments.IsEmpty do
            total <- lineArrayCacheAddCapped (lineArrayCacheSegmentTextCharCount remainingSegments.Head) total
            remainingSegments <- remainingSegments.Tail
        remainingLines <- remainingLines.Tail
    min total limit

let private lineArrayCacheWithoutLines lines =
    let mutable cached = Unchecked.defaultof<LineArrayCacheEntry>
    if lineArrayCacheByLines.TryGetValue(lines, &cached)
       && lineArrayCacheEntryMatches lines cached then
        lineArrayCache
        |> List.filter (lineArrayCacheEntryMatches lines >> not)
    else
        lineArrayCache

let private rememberLineArrayCacheEntry entry =
    if entry.LineArray.Length > lineArrayCacheMaxTotalLines
       || entry.TextCharCount > lineArrayCacheMaxTotalTextChars then
        let entries = lineArrayCacheWithoutLines entry.Lines
        if not (obj.ReferenceEquals(entries, lineArrayCache)) then
            setLineArrayCache entries
    else
        let mutable keptCount = 1
        let mutable keptLines = entry.LineArray.Length
        let mutable keptTextChars = max 1 entry.TextCharCount
        let retained =
            lineArrayCacheWithoutLines entry.Lines
            |> List.choose (fun candidate ->
                let candidateLines = candidate.LineArray.Length
                let candidateTextChars = max 1 candidate.TextCharCount
                if keptCount < lineArrayCacheMaxEntries
                   && keptLines + candidateLines <= lineArrayCacheMaxTotalLines
                   && keptTextChars + candidateTextChars <= lineArrayCacheMaxTotalTextChars then
                    keptCount <- keptCount + 1
                    keptLines <- keptLines + candidateLines
                    keptTextChars <- keptTextChars + candidateTextChars
                    Some candidate
                else
                        None)
        setLineArrayCache (entry :: retained)
        lastLineArrayCacheEntry <- Some entry

let private cachedLineArrayEntry lines =
    match tryFindLineArrayCacheEntry lines with
    | Some entry ->
        entry
    | None ->
        let lineArray = Array.ofList lines
        let entry = {
            Lines = lines
            LineArray = lineArray
            TextCharCount = lineArrayCacheTextCharCount lines
            LastLineSet = false
            LastOffset = 0
            LastLine = Unchecked.defaultof<MeasuredLine>
        }
        rememberLineArrayCacheEntry entry
        entry

let private cachedLineArray lines =
    (cachedLineArrayEntry lines).LineArray

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

let private lineForOffsetCacheEntry layout entry offset =
    if entry.LastLineSet && entry.LastOffset = offset then
        entry.LastLine
    else
        let line = lineForOffsetArray layout entry.LineArray offset
        entry.LastLineSet <- true
        entry.LastOffset <- offset
        entry.LastLine <- line
        line

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
    match lastLineArrayCacheEntry with
    | Some entry when lineArrayCacheEntryMatches layout.Lines entry ->
        lineForOffsetCacheEntry layout entry offset
    | _ ->
        match tryFindLineArrayCacheEntry layout.Lines with
        | Some entry ->
            lineForOffsetCacheEntry layout entry offset
        | None ->
            if lineListExceeds lineArrayCacheThreshold layout.Lines then
                lineForOffsetCacheEntry layout (cachedLineArrayEntry layout.Lines) offset
            else
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

let xForOffsetWithRange (measureRange: TextRangeMeasurer) (line: MeasuredLine) (offset: int) =
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

let xForOffset (measureText: TextMeasurer) (line: MeasuredLine) (offset: int) =
    xForOffsetWithRange (rangeMeasurerFromTextMeasurer measureText) line offset

let private layoutContainsY mouseY (layout: BlockLayout) =
    mouseY >= layout.Y && mouseY < layout.Y + layout.Height

let private hitTestLayoutWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layout: BlockLayout) (mouseX: float) (mouseY: float) =
    let mutable lineOpt = None
    let mutable remaining = layout.Lines
    while lineOpt.IsNone && not remaining.IsEmpty do
        match remaining with
        | line :: tail ->
            if mouseY >= line.Y && mouseY < line.Y + line.Height then
                lineOpt <- Some line
            remaining <- tail
        | [] -> ()

    match lineOpt with
    | None -> Some (layout.BlockId, 0)
    | Some line ->
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

        Some (layout.BlockId, offset)

let hitTestLayoutsWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) =
    match layouts |> List.tryFind (layoutContainsY mouseY) with
    | None -> None
    | Some layout -> hitTestLayoutWithRange measureRange paddingLeft layout mouseX mouseY

let hitTestLayoutArrayWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout array) (mouseX: float) (mouseY: float) =
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

    let mutable result = None
    while result.IsNone && index < layouts.Length && layouts.[index].Y <= mouseY do
        let layout = layouts.[index]
        if layoutContainsY mouseY layout then
            result <- hitTestLayoutWithRange measureRange paddingLeft layout mouseX mouseY
        index <- index + 1
    result

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
        elif isWhitespaceChar text.[index] then
            while index < text.Length && isWhitespaceChar text.[index] && text.[index] <> '\r' && text.[index] <> '\n' do
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
    let placedImageSources = System.Collections.Generic.HashSet<string>()

    let addPlacedImageSource (box: PlacedFloatingBox) =
        match box.Kind with
        | WordLayout.Floating.ImageBox image when image.Src <> "" ->
            placedImageSources.Add image.Src |> ignore
        | _ -> ()

    for box in initialPlacedBoxes do
        addPlacedImageSource box

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
        List.rev gaps

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

            if currentX + imgWidth > currentLineEnd && currentSegments.Count > 0 then
                flushLine ()

            let rawFloatingBox = WordLayout.Floating.imageCanvasBox settings.PaddingLeft contentY currentX y settings.PaddingTop imgWidth imgHeight imgRun
            let floatingBox = if imgRun.AllowOverlap = Some false then WordLayout.Floating.avoidOverlap placedBoxes rawFloatingBox else rawFloatingBox
            let imageX = if imgRun.Floating then floatingBox.X else currentX
            let yOffset = if imgRun.Floating then floatingBox.Y - (contentY + (y - settings.PaddingTop)) else 0.0
            let segment = ImageSegment (imgRun.Src, imageX, yOffset, imgWidth, imgHeight, imgRun.Start, imgRun.Floating, imgRun.Wrap, imgRun.RelativeHeight, imgRun.CropLeft, imgRun.CropTop, imgRun.CropRight, imgRun.CropBottom)
            currentSegments.Add segment
            if not imgRun.Floating then
                currentHasFlowContent <- true
            // The same float may have been pre-placed by an anchor-extended pass
            // earlier in the document layout. Detect that by matching the image
            // source against the existing placed boxes; in that case the
            // obstacle is already in `floatObstacles` with an extended (upward)
            // top, so just emit the ImageSegment and keep wrapping as before.
            let alreadyPlaced =
                imgRun.Floating
                && imgRun.Src <> ""
                && placedImageSources.Contains imgRun.Src
            if imgRun.Floating && WordLayout.Floating.participatesInWrap floatingBox && not alreadyPlaced then
                let obstacleX, obstacleTop, obstacleWidth, obstacleHeight = WordLayout.Floating.obstacleForBox availableWidth floatingBox
                if WordLayout.Floating.wrapsTopAndBottom floatingBox.Wrap then
                    floatObstacles <- (0.0, obstacleTop, availableWidth, obstacleHeight) :: floatObstacles
                    moveToWrappableLine ()
                else
                    floatObstacles <- (obstacleX, obstacleTop, obstacleWidth, obstacleHeight) :: floatObstacles
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
                addPlacedImageSource floatingBox

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
            let lineSpacing = WordLayout.Paragraph.lineSpacing paragraphStyle metadata
            let lineSpacingRule = WordStyles.lineSpacingRule paragraphStyle metadata
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
                        let lineHeight = lineHeightForSettings settings contentTextStyle contentFontSize lineSpacing lineSpacingRule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox
                        let lineScale = lineHeight / contentFontSize
                        let textHeight = contentFontSize * lineScale
                        Some (spacingBefore + borderBefore + max lineHeight textHeight + borderAfter + spacingAfter)
            | _ -> None

let layoutParagraphAtWithStyleAndObstaclesWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    let mutable counters = numberedCounters
    let metadata = WordStyles.effectiveMetadata paragraphStyle para.Metadata
    let indent = WordLayout.Paragraph.leftIndent metadata
    let rightIndent = WordLayout.Paragraph.rightIndent metadata
    let firstLineIndent = WordLayout.Paragraph.firstLineIndent metadata
    let rawSpacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle metadata
    let spacingBefore = rawSpacingBefore
    let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle metadata
    let lineSpacing = WordLayout.Paragraph.lineSpacing paragraphStyle metadata
    let lineSpacingRule = WordStyles.lineSpacingRule paragraphStyle metadata
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
    let buffer = effectiveBufferForParagraph paragraphStyle para.Buffer

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
    let contentFontSize, contentTextStyle = paragraphContentMetrics paragraphStyle metadata buffer
    let lineSettings = { settings with FontSize = contentFontSize; LineHeight = lineHeightForSettings settings contentTextStyle contentFontSize lineSpacing lineSpacingRule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox }
    let lines, updatedObstacles, updatedPlacedBoxes =
        if prefix = "" && not metadata.Rtl then
            match trySingleLineStyledNoObstacles lineSettings measureRange buffer availableWidth metadata.TabStops initialObstacles initialPlacedBoxes with
            | Some lines -> lines, initialObstacles, initialPlacedBoxes
            | None -> layoutLinesStyledWithObstaclesWithRange lineSettings measureRange buffer availableWidth contentY metadata.TabStops initialObstacles initialPlacedBoxes
        else
            layoutLinesStyledWithObstaclesWithRange lineSettings measureRange buffer availableWidth contentY metadata.TabStops initialObstacles initialPlacedBoxes

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
                    max line.Height (contentFontSize * lineSpacing * markerFactor)
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
        if pagedLines.IsEmpty then lineHeightForSettings settings contentTextStyle contentFontSize lineSpacing lineSpacingRule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox
        else
            match tryLastLine pagedLines with
            | Some lastLine -> (lastLine.Y - contentY) + lastLine.Height
            | None -> lineHeightForSettings settings contentTextStyle contentFontSize lineSpacing lineSpacingRule hasExplicitLineSpacing defaultParagraphSpacing tocEntry headingDefaultLineBox

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

let private defaultParagraphFirstLineOffsetInDocument doc =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInDocument doc (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private defaultParagraphFirstLineOffsetInContext context =
    let para = { Id = ""; Buffer = TextBuffer.empty; Metadata = defaultMetadata }
    let paragraphStyle = forBlockInContext context (Paragraph para)
    paragraphStyle.FirstLineIndent |> Option.defaultValue 0.0 |> max 0.0

let private tableLeftOffsetForDocument doc (_tableBlock: TableBlock) =
    defaultParagraphFirstLineOffsetInDocument doc

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

let rec private measureTableRowWithStyleAtDepth nestedTableLayoutDepth (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
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

                    for nestedRow in nestedTable.Rows do
                        let measurement =
                            measureTableRowWithStyleAtDepth (nestedTableLayoutDepth + 1) cellSettings measureText measureRange nestedTable nestedColumnMetrics nestedOuterLeft nestedHorizontalOffset nestedRow nestedY 0.0 paragraphStyleForParagraph effectiveBufferForParagraph
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
            let layout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange cellSettings measureText measureRange cellMaxWidth paraY cellX paragraphStyle effectiveBufferForParagraph para Map.empty [] []
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
        rowHeight <- max rowHeight (contentHeight + textPadding.Top + textPadding.Bottom + verticalMarginExcess + verticalBorderExtent + rowSpacingAdvance)
        let cellLayoutsList = resizeArrayToList cellLayouts
        cellGroups <- { VerticalAlignment = cell.VerticalAlignment; PaddingTop = textPadding.Top; PaddingBottom = textPadding.Bottom; ContentHeight = contentHeight; Layouts = cellLayoutsList } :: cellGroups
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

    { Height = rowHeight; CellGroups = List.rev cellGroups; Traces = List.rev traces }

let internal measureTableRowWithStyleUsingMetrics (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    measureTableRowWithStyleAtDepth 0 settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph

let measureTableRowWithStyle (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (tableMaxWidth: float) (columnWidths: float list) (outerLeft: float) (tableLeftOffset: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowWithStyleUsingMetrics settings measureText measureRange tableBlock columnMetrics outerLeft tableLeftOffset row rowY tableContinuationHeaderHeight paragraphStyleForParagraph effectiveBufferForParagraph

let internal tryMeasureTableRowHeightWithMetrics (settings: LayoutSettings) (measureText: TextMeasurer) (tableBlock: TableBlock) (columnMetrics: TableColumnMetrics) (row: TableRow) (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveTextStyleForParagraph: ParagraphStyle -> TextBuffer.TextStyle -> TextBuffer.TextStyle) =
    let cellSettings = {
        settings with
            InTableCell = true
            MarginBottom = if row.CanSplit = Some false then -1000000.0 else settings.MarginBottom
            TableContinuationHeaderHeight = 0.0
    }
    let numColumns = columnMetrics.Count
    if numColumns = 0 then
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

let internal measureTableRowForContextWithColumnMetrics settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    measureTableRowWithStyleUsingMetrics settings measureText measureRange tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight (forTableCellParagraphInContext context) (effectiveBufferInContext context)

let measureTableRowForContext settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowForContextWithColumnMetrics settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

let measureTableRowForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight =
    let context = WordStyles.styleContext doc
    measureTableRowForContext settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

let internal measureTableRowForContextRowLocalWithColumnMetrics settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row =
    let rowLocalSettings = {
        settings with
            PageHeight = 1000000.0
            MarginBottom = 0.0
            TableContinuationHeaderHeight = 0.0
    }
    measureTableRowForContextWithColumnMetrics rowLocalSettings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row 0.0 0.0

let measureTableRowForContextRowLocal settings measureText measureRange context tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset row =
    let columnMetrics = tableColumnMetrics tableMaxWidth columnWidths
    measureTableRowForContextRowLocalWithColumnMetrics settings measureText measureRange context tableBlock columnMetrics outerLeft tableHorizontalOffset row

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
        for fragment in planRow.Fragments do
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
                        CellHeight = WordLayout.Table.spacedCellPaintHeight cellFragmentHeight cellSpacing + paintTopAdjustment
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
        for fragment in planRow.Fragments do
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
                        CellHeight = WordLayout.Table.spacedCellPaintHeight cellFragmentHeight cellSpacing + paintTopAdjustment
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

let buildTableLayoutPlanForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY =
    let styleContext = WordStyles.styleContext doc
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
        measureTableRowForContextWithColumnMetrics settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight
    let measureRowLocal row =
        measureTableRowForContextRowLocalWithColumnMetrics settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row
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

let buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY =
    let tableLeftOffset = tableLeftOffsetForDocument doc tableBlock
    let maxWidth = max 0.0 (settings.Width - settings.PaddingLeft - settings.MarginRight - tableLeftOffset)
    let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
    let tableMaxWidth = maxWidth + outerLeft + outerRight
    let columnWidths = tableColumnWidths tableMaxWidth tableBlock
    let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
    let tableHorizontalOffset = tableLeftOffset + tableAlignmentOffset tableBlock.Alignment maxWidth tableWidth
    buildTableLayoutPlanForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableHorizontalOffset tableY

/// Layout document blocks. Table cell paragraphs are included so shared hit/cursor logic works inside tables.
let private layoutDocumentWithRangeAndInitial
        (settings: LayoutSettings)
        (measureText: TextMeasurer)
        (measureRange: TextRangeMeasurer)
        (doc: Document)
        (initialObstacles: FloatingObstacle list)
        (initialPlacedBoxes: PlacedFloatingBox list) =
    // Multi-column sections are flowed by columnFlowLayout, which drives this
    // function per column at the column width; here we always use the full width.
    let styleContext = WordStyles.styleContext doc
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let mutable y = settings.PaddingTop
    let layouts = ResizeArray<BlockLayout>()
    let mutable numberedCounters = Map.empty<string * int, int>
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let mutable floatingObstacles: FloatingObstacle list = initialObstacles
    let mutable placedFloatingBoxes: PlacedFloatingBox list = initialPlacedBoxes
    let mutable previousParagraphContext: (ParagraphBlock * ParagraphStyle * BlockMetadata * float) option = None
    let mutable previousParagraphLayout: BlockLayout option = None
    let mutable hasEndnoteRefs = false
    let blocks = doc.Blocks |> List.toArray

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
                let tableLeftOffset = tableLeftOffsetForDocument doc tableBlock
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
                let tablePlan = lazy (buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY)
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

                let measureRow (row: TableRow) rowY tableContinuationHeaderHeight =
                    measureTableRowForContextWithColumnMetrics settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row rowY tableContinuationHeaderHeight

                let measureRowLocal (row: TableRow) =
                    measureTableRowForContextRowLocalWithColumnMetrics settings measureText measureRange styleContext tableBlock columnMetrics outerLeft tableHorizontalOffset row

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
                        addAdjustedLayouts offset queryVAlignOffset group.Layouts
                    rowY <- rowY + rowHeight
                    rowIndex <- rowIndex + 1

                rowY + WordLayout.Table.tableAfterSpacing WordLayout.canvasScale tableBlock, tableWidth
            else
                tableY + WordLayout.Table.emptyTableAdvance settings.LineHeight, 0.0
        if traceTimingEnabled then
            traceSlow (sprintf "table rows=%d" (List.length tableBlock.Rows)) tableStart
        result

    for index in 0 .. blocks.Length - 1 do
        let block = blocks.[index]
        match block with
        | TextBox textBox ->
            resetPreviousParagraph ()
            addTextBoxObstacle textBox
        | Table tableBlock ->
            let tableY = y + tableFollowDefaultLineHeightAdjustment ()
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
                    let stylesMatch = sameParagraphStyle previousPara para
                    let effectivePreviousSpacingAfter =
                        if stylesMatch && contextualSpacingEnabled previousStyle previousMetadata then 0.0 else previousSpacingAfter
                    let effectiveSpacingBefore =
                        if stylesMatch && contextualSpacingEnabled paragraphStyle effectiveMetadata then 0.0 else spacingBefore
                    y <- (y - previousSpacingAfter) + max effectivePreviousSpacingAfter effectiveSpacingBefore - spacingBefore
                    if stylesMatch && sameParagraphBorders previousMetadata effectiveMetadata then
                        let _, previousBorderAfter = paragraphBorderVerticalAdvance previousMetadata
                        let currentBorderBefore, _ = paragraphBorderVerticalAdvance effectiveMetadata
                        y <- y - previousBorderAfter - currentBorderBefore
                | None -> ()
            let initialStart = timingStart ()
            let initialLayout, initialCounters, initialObstacles, initialPlacedBoxes = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth y 0.0 paragraphStyle (effectiveBufferInContext styleContext) para numberedCounters floatingObstacles placedFloatingBoxes
            traceSlow (sprintf "block[%d]:initial lines=%d" index initialLayout.Lines.Length) initialStart
            let keepWithNextY =
                if paragraphStyle.KeepNext = Some true then
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
                                let nextLayout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth probeY 0.0 nextStyle (effectiveBufferInContext styleContext) nextPara numberedCounters initialObstacles initialPlacedBoxes
                                let nextKeepsFurther = nextStyle.KeepNext = Some true
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
                                if nextPara.Metadata.PageBreakAfter || nextPara.Metadata.SectionBreak = Some "nextPage" then
                                    groupIsMoveable <- false
                                    shouldContinue <- false
                                else
                                    shouldContinue <- nextKeepsFurther
                                probeIndex <- probeIndex + 1
                        | None -> shouldContinue <- false

                    traceSlow (sprintf "block[%d]:keepNextProbe lines=%d" index probedLines) keepStart
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
                        let nextLayout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth (pagedY + initialLayout.Height) 0.0 nextStyle (effectiveBufferInContext styleContext) nextPara numberedCounters initialObstacles initialPlacedBoxes
                        traceSlow (sprintf "block[%d]:nearFooterProbe lines=%d" index nextLayout.Lines.Length) nearFooterStart
                        Some nextLayout.Lines.Length
                    | None -> None
                else None
            let pagedY =
                if not initialLinesSpanPages && pagedY = keepWithNextY && shouldMoveNearFooter pagedY initialLayout.Height initialLayout.Lines.Length initialLayout.Metadata.SpacingAfter nextLineCount then
                    advanceToNextPage settings pagedY
                else pagedY
            let pagedY =
                if paragraphStyle.KeepLines = Some true && pagedY = y && y > settings.PaddingTop && linesSpanPages settings initialLayout.Lines then
                    advanceToNextPage settings y
                else pagedY
            let layout, counters, obstacles, placedBoxes =
                if pagedY = y then
                    initialLayout, initialCounters, initialObstacles, initialPlacedBoxes
                else
                    let relayoutStart = timingStart ()
                    let result =
                        layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth pagedY 0.0 paragraphStyle (effectiveBufferInContext styleContext) para numberedCounters floatingObstacles placedFloatingBoxes
                    traceSlow (sprintf "block[%d]:relayout" index) relayoutStart
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
            if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                y <- advanceToNextPage settings y
                resetPreviousParagraph ()
            traceSlow (sprintf "block[%d]:paragraph total lines=%d" index layout.Lines.Length) blockStart

    // Lay out each page's footnotes in the reserved area at the bottom of the
    // page, at a smaller font, tagged with footnoteBlockIdPrefix so Paint can
    // render them and draw the separator. The note id doubles as Word's mark.
    let footnoteParagraphMetadata = { defaultMetadata with SpacingBefore = Some 0.0; SpacingAfter = Some 0.0 }
    let footnoteStride = pageHeight settings + pageGap
    for KeyValue (pageIndex, notes) in footnotesByPage do
        if not notes.IsEmpty then
            let pageTop = float pageIndex * footnoteStride
            let pageContentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            // Lay out the notes at a temporary origin to get their real height
            // (10pt body font, superscript mark number), then place the block so
            // it ends just above the bottom margin, the way Word stacks footnotes.
            let mutable tempY = 0.0
            let laidOut =
                notes |> List.mapi (fun noteIndex (id, text) ->
                    let noteString = sprintf "%s %s" id (text.Trim())
                    let buffer =
                        TextBuffer.ofString noteString
                        |> TextBuffer.setFontFamily 0 noteString.Length footnoteFontFamily
                        |> TextBuffer.setFontSize 0 noteString.Length footnoteFontSize
                        |> TextBuffer.setVertAlign 0 id.Length (Some "superscript")
                    let para = { Id = sprintf "%s%d_%d" footnoteBlockIdPrefix pageIndex noteIndex; Buffer = buffer; Metadata = footnoteParagraphMetadata }
                    let layout, _ = layoutParagraphAtWithRange settings measureText measureRange maxWidth tempY 0.0 para Map.empty
                    tempY <- tempY + layout.Height
                    layout)
            let blockTop = pageContentBottom - footnoteBottomPad - tempY
            for layout in laidOut do
                layouts.Add(moveLayoutToY (blockTop + layout.Y) layout)

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
        bodyLayouts.Sort(System.Comparison<BlockLayout>(fun left right -> compare left.Y right.Y))
        let endnotes = ResizeArray<string * string>()
        let seenEndnotes = System.Collections.Generic.HashSet<string>()
        for layout in bodyLayouts do
            for line in layout.Lines do
                for id, text in endnotesInLine line do
                    if seenEndnotes.Add id then
                        endnotes.Add(id, text)
        let stride = pageHeight settings + pageGap
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
            let para = { Id = sprintf "%s%d" endnoteBlockIdPrefix noteIndex; Buffer = buffer; Metadata = footnoteParagraphMetadata }
            // Lay the note out at a safe Y (page-1 content top) to get its natural
            // height without triggering layoutParagraph's own page-break handling,
            // then place it; if it won't fit on the current page, start it at the
            // top of the next page (Word continues endnotes across pages).
            let probe, _ = layoutParagraphAtWithRange settings measureText measureRange maxWidth settings.PaddingTop 0.0 para Map.empty
            let height = probe.Height
            let pageIndex = int (floor (noteY / stride))
            let pageTop = float pageIndex * stride
            let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
            let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            let startY =
                if noteY > contentTop && noteY + height > contentBottom
                then WordLayout.Page.contentTop settings.PaddingTop (float (pageIndex + 1) * stride)
                else noteY
            layouts.Add(moveLayoutToY startY probe)
            noteY <- startY + height

    resizeArrayToList layouts, placedFloatingBoxes

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
        |> List.choose (fun box ->
            if not (WordLayout.Floating.participatesInWrap box) then None
            else
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
                    Some (0.0, obstacleTop, availableWidth, obstacleHeight)
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
                    Some (obstacleX, obstacleTop, obstacleWidth, obstacleHeight)
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
                    Some (obstacleX, extendedTop, obstacleWidth, extendedHeight))
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
let private flowColumns (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (columns: SectionColumns) (pageTopOf: int -> float) (nonUniform: bool) (startPage: int) (startY: float) (balanced: bool) =
    let styleContext = WordStyles.styleContext doc
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
        if balanced then
            let w = columnWidth 0
            let sub, _ = layoutDocumentWithRangeAndInitial { settings with Width = settings.PaddingLeft + w + settings.MarginRight; PageHeight = 100000.0 } measureText measureRange doc [] []
            let bottom = maxLineBottom settings.PaddingTop sub
            Some ((bottom - settings.PaddingTop) / float n + 2.0)
        else None

    let results = ResizeArray<BlockLayout>()
    let mutable col = 0
    let mutable page = startPage
    let mutable y = startY
    let mutable counters : Map<string * int, int> = Map.empty
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
        | None -> contentBottomAt page
    let atColumnTop () = abs (y - columnTop ()) < 0.5
    let advanceColumn () =
        col <- col + 1
        if col >= n then
            col <- 0
            if balancedHeight.IsNone then page <- page + 1
        y <- columnTop ()

    for block in doc.Blocks do
        match tryGetParagraph block with
        | Some para0 ->
            if (para0.Metadata.PageBreakBefore || para0.Metadata.RenderedPageBreakBefore.IsSome) && not (atColumnTop ()) then
                advanceColumn ()
            let style = forBlockInContext styleContext block
            let mutable remaining = para0
            let mutable placed = false
            while not placed do
                let colBottom = columnBottom ()
                let layout, newCounters, _, _ =
                    layoutParagraphAtWithStyleAndObstaclesWithRange flowSettings measureText measureRange (columnWidth col) y 0.0 style (effectiveBufferInContext styleContext) remaining counters [] []
                let mutable fitCount = 0
                let mutable lastFit = None
                for line in layout.Lines do
                    if line.Y + line.Height <= colBottom + 0.5 then
                        fitCount <- fitCount + 1
                        lastFit <- Some line
                if fitCount = layout.Lines.Length then
                    results.Add(shiftBlockX (columnDx col) layout)
                    counters <- newCounters
                    y <- layout.Y + layout.Height
                    placed <- true
                elif fitCount = 0 then
                    if atColumnTop () then
                        results.Add(shiftBlockX (columnDx col) layout)
                        counters <- newCounters
                        y <- layout.Y + layout.Height
                        placed <- true
                    else advanceColumn ()
                else
                    let fitLines = takeLines fitCount layout.Lines
                    match lastFit with
                    | Some lastFit ->
                        results.Add(shiftBlockX (columnDx col) { layout with Lines = fitLines; Height = lastFit.Y + lastFit.Height - layout.Y })
                        remaining <- { remaining with Buffer = TextBuffer.delete 0 lastFit.EndOffset remaining.Buffer }
                        advanceColumn ()
                    | None ->
                        advanceColumn ()
            if para0.Metadata.PageBreakAfter && not (atColumnTop ()) then advanceColumn ()
        | None ->
            let sub = { doc with Blocks = [block] }
            let subLayouts, _ =
                layoutDocumentWithRangeAndInitial { settings with Width = settings.PaddingLeft + columnWidth col + settings.MarginRight; PageHeight = 100000.0 } measureText measureRange sub [] []
            match subLayouts with
            | [] -> ()
            | _ ->
                let mutable subTop = System.Double.PositiveInfinity
                let mutable subBottom = System.Double.NegativeInfinity
                for layout in subLayouts do
                    if layout.Y < subTop then subTop <- layout.Y
                    let layoutBottom = layout.Y + layout.Height
                    if layoutBottom > subBottom then subBottom <- layoutBottom
                let height = subBottom - subTop
                if not (atColumnTop ()) && y + height > columnBottom () + 0.5 then advanceColumn ()
                let dy = y - subTop
                for sl in subLayouts do
                    results.Add(sl |> shiftBlockY dy |> shiftBlockX (columnDx col))
                y <- y + height
    resizeArrayToList results, page, y

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

/// The uniform page list for a single-section document (legacy stride model).
let uniformPageInfosWithLayoutCount (settings: LayoutSettings) (layouts: BlockLayout list) =
    let stride = pageHeight settings + pageGap
    let bottom, layoutCount = maxLineBottomAndLayoutCount 0.0 layouts
    let pageCount = max 1 (int (floor (max 0.0 bottom / stride)) + 1)
    [ for p in 0 .. pageCount - 1 -> pageInfoOf settings (float p * stride) ], layoutCount

let uniformPageInfos (settings: LayoutSettings) (layouts: BlockLayout list) =
    uniformPageInfosWithLayoutCount settings layouts |> fst

/// Lay out a multi-section document section by section. A continuous break
/// continues on the same page below the previous section; otherwise the next
/// section starts on a new page (using that section's page size — supporting a
/// landscape page mid-document). Non-final multi-column sections are balanced.
/// Returns the layouts plus the per-page geometry.
let private layoutSections (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) sections =
    let results = ResizeArray<BlockLayout>()
    let pageTops = ResizeArray<float>()    // absolute top of each page index
    let pageSettings = ResizeArray<LayoutSettings>()  // settings owning each page
    pageTops.Add 0.0
    let mutable currentPage = 0
    let mutable currentY = 0.0             // absolute Y where the next section starts
    let mutable started = false
    let recordPage page (s: LayoutSettings) top =
        while pageTops.Count <= page do pageTops.Add 0.0
        while pageSettings.Count <= page do pageSettings.Add s
        pageTops.[page] <- top
        pageSettings.[page] <- s
    sections
    |> List.iteri (fun i (blocks, columns, geometry, breakType, isFinal) ->
        let secSettings = settingsForGeometry settings geometry
        let secStride = pageHeight secSettings + pageGap
        // Stitch before this section.
        if not started then
            currentPage <- 0
            recordPage 0 secSettings 0.0
            currentY <- WordLayout.Page.contentTop secSettings.PaddingTop 0.0
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
                currentY <- WordLayout.Page.contentTop secSettings.PaddingTop newTop
        // Page tops within this section are uniform (secStride from the section top).
        let sectionBase = currentPage
        let sectionBaseTop = pageTops.[currentPage]
        let pageTopOf p = sectionBaseTop + float (p - sectionBase) * secStride
        let secColumns = columnsOrSingle columns
        let secDoc = { doc with Blocks = blocks; PageSetup = { doc.PageSetup with Columns = columns } }
        let balanced = (not isFinal) && secColumns.Count >= 2
        let secLayouts, endPage, endY = flowColumns secSettings measureText measureRange secDoc secColumns pageTopOf true currentPage currentY balanced
        results.AddRange secLayouts
        // Record the geometry of every page this section used.
        for p in sectionBase .. endPage do recordPage p secSettings (pageTopOf p)
        currentPage <- endPage
        currentY <- endY)
    let maxPage = pageTops.Count - 1
    let pages = [ for p in 0 .. maxPage -> pageInfoOf pageSettings.[p] pageTops.[p] ]
    resizeArrayToList results, pages

/// Lay out a document, returning the layouts and the per-page geometry.
let layoutDocumentPages (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) : BlockLayout list * PageInfo list =
    match splitIntoSections doc with
    | [ (_, columns, _, _, _) ] ->
        // Single section: keep the existing path; uniform page list.
        let layouts =
            match columns with
            | Some c when c.Count >= 2 ->
                let stride = pageHeight settings + pageGap
                let r, _, _ = flowColumns settings measureText measureRange doc c (fun p -> float p * stride) false 0 (WordLayout.Page.contentTop settings.PaddingTop 0.0) false
                r
            | _ ->
                let firstPass, placedAfterFirst =
                    layoutDocumentWithRangeAndInitial settings measureText measureRange doc [] []
                let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
                let extendedObstacles, prePlacedBoxes =
                    computeAnchorExtendedObstacles maxWidth firstPass placedAfterFirst
                if extendedObstacles.IsEmpty then firstPass
                else
                    let secondPass, _ =
                        layoutDocumentWithRangeAndInitial settings measureText measureRange doc extendedObstacles prePlacedBoxes
                    secondPass
        layouts, uniformPageInfos settings layouts
    | sections ->
        layoutSections settings measureText measureRange doc sections

let layoutDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) =
    layoutDocumentPages settings measureText measureRange doc |> fst

let layoutDocument (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) =
    layoutDocumentWithRange settings measureText (rangeMeasurerFromTextMeasurer measureText) doc

let hitTest (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (mouseX: float) (mouseY: float) =
    let layouts = layoutDocument settings measureText doc
    hitTestLayouts measureText settings.PaddingLeft layouts mouseX mouseY

let hitTestWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) (mouseX: float) (mouseY: float) =
    let layouts = layoutDocumentWithRange settings measureText measureRange doc
    hitTestLayoutsWithRange measureRange settings.PaddingLeft layouts mouseX mouseY

let hitTestForImage (settings: LayoutSettings) (measureText: TextMeasurer) (doc: Document) (mouseX: float) (mouseY: float) : (BlockId * int * float * float * float * float) option =
    let layouts = layoutDocument settings measureText doc
    hitTestForImageLayouts settings.PaddingLeft layouts mouseX mouseY
