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

let private snippet (text: string) =
    let clean = text.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim()
    if clean.Length > 48 then clean.Substring(0, 48) else clean

/// Segment can be text or image.
type Segment =
    | TextSegment of text: string * style: TextBuffer.TextStyle * x: float * width: float * startOffset: int
    | ImageSegment of src: string * x: float * yOffset: float * width: float * height: float * startOffset: int * floating: bool * wrap: string option * relativeHeight: string option * cropLeft: string option * cropTop: string option * cropRight: string option * cropBottom: string option

/// Measured line for layout.
type MeasuredLine = {
    Segments: Segment list
    Y: float
    Width: float
    Height: float
    StartOffset: int
    EndOffset: int
}

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
    PaddingTop: float
    PaddingBottom: float
    ContentY: float
    ContentHeight: float
    VisibleY: float option
    VisibleHeight: float option
    VisibleLineCount: int
    BlockIds: BlockId list
}

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

let private moveLayoutToY newY (layout: BlockLayout) =
    let delta = newY - layout.Y
    { layout with
        Y = layout.Y + delta
        Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + delta }) }

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
    max 24 (int ((settings.Width - settings.PaddingLeft - settings.MarginRight) / 5.5))

let wrapNoteText settings (text: string) =
    let maxChars = noteWrapChars settings
    let words = text.Split([| ' '; '\t'; '\r'; '\n' |], System.StringSplitOptions.RemoveEmptyEntries) |> Array.toList
    let rec loop current lines remaining =
        match remaining with
        | [] ->
            match current with
            | "" -> List.rev lines
            | value -> List.rev (value :: lines)
        | word :: rest ->
            if current = "" then loop word lines rest
            elif current.Length + 1 + word.Length <= maxChars then loop (current + " " + word) lines rest
            else loop word (current :: lines) rest
    match words with
    | [] -> [text]
    | _ -> loop "" [] words

let noteLineCount settings (text: string) =
    wrapNoteText settings text |> List.length

let private footnotesInLine (line: MeasuredLine) =
    line.Segments
    |> List.choose (function
        | TextSegment (_, style, _, _, _) ->
            match style.NoteId, style.NoteType, style.NoteText with
            | Some id, noteType, Some text when noteType |> Option.defaultValue "footnote" = "footnote" && text <> "" -> Some (id, text)
            | _ -> None
        | ImageSegment _ -> None)
    |> List.distinctBy fst

let private footnoteReserveHeight settings footnotes =
    match footnotes with
    | [] -> 0.0
    | notes ->
        let lineCount = notes |> List.sumBy (fun (id, text) -> noteLineCount settings (sprintf "%s. %s" id text))
        20.0 + float lineCount * 13.0

let private addLayoutFootnotes settings (footnotesByPage: Map<int, (string * string) list>) (layout: BlockLayout) =
    layout.Lines
    |> List.fold (fun acc line ->
        let pageIndex = pageIndexForY settings line.Y
        let footnotes =
            (acc |> Map.tryFind pageIndex |> Option.defaultValue []) @ footnotesInLine line
            |> List.distinctBy fst
        acc |> Map.add pageIndex footnotes) footnotesByPage

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
    line.Segments
    |> List.choose (function
        | TextSegment (_, style, _, _, _) -> Some (WordLayout.Text.printFontSize style)
        | ImageSegment (_, _, yOffset, _, height, _, floating, _, _, _, _, _, _) -> if floating then None else Some (height + max 0.0 yOffset))
    |> function
        | [] -> line.Height
        | sizes -> List.max sizes

let private paragraphVisibleFitHeight (paragraphStyle: ParagraphStyle) (layout: BlockLayout) =
    if layout.Metadata.ListStyle = Numbered then
        layout.Height
    else
        let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle layout.Metadata
        match layout.Lines |> List.rev with
        | [] -> max 0.0 (layout.Height - spacingAfter)
        | lastLine :: _ ->
            max 0.0 ((lastLine.Y - layout.Y) + lineVisibleFitHeight lastLine)

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

let private moveWidowParagraphToNextPage settings (lines: MeasuredLine list) =
    match lines with
    | first :: second :: _ ->
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) first.Y
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        let contentBottom = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
        let lineFits (line: MeasuredLine) = line.Y + lineVisibleFitHeight line <= contentBottom + pageFitTolerance
        if first.Y > contentTop && lineFits first && not (lineFits second) then
            let nextY = WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + settings.TableContinuationHeaderHeight
            let delta = nextY - first.Y
            lines |> List.map (fun line -> { line with Y = line.Y + delta })
        else lines
    | _ -> lines

let private paginateLines settings widowControl (lines: MeasuredLine list) =
    let mutable delta = 0.0
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let lines = if widowControl then moveWidowParagraphToNextPage settings lines else lines
    lines
    |> List.map (fun line ->
        let shiftedY = line.Y + delta
        let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) shiftedY
        let contentTop = WordLayout.Page.contentTop settings.PaddingTop pageTop
        let pageIndex = pageIndexForY settings shiftedY
        let lineFootnotes = footnotesInLine line
        let pageFootnotes =
            (footnotesByPage |> Map.tryFind pageIndex |> Option.defaultValue []) @ lineFootnotes
            |> List.distinctBy fst
        let contentBottom =
            WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop
            - footnoteReserveHeight settings pageFootnotes
        if shiftedY > contentTop && shiftedY + lineVisibleFitHeight line > contentBottom + pageFitTolerance then
            let nextY = WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + settings.TableContinuationHeaderHeight
            delta <- delta + (nextY - shiftedY)
            let nextPageIndex = pageIndexForY settings nextY
            let nextPageFootnotes =
                (footnotesByPage |> Map.tryFind nextPageIndex |> Option.defaultValue []) @ lineFootnotes
                |> List.distinctBy fst
            footnotesByPage <- footnotesByPage |> Map.add nextPageIndex nextPageFootnotes
            { line with Y = nextY }
        else
            footnotesByPage <- footnotesByPage |> Map.add pageIndex pageFootnotes
            { line with Y = shiftedY })

let private linesSpanPages settings (lines: MeasuredLine list) =
    match lines with
    | [] -> false
    | first :: _ ->
        let firstPage = WordLayout.Page.pageTopForY (pageHeight settings) first.Y
        lines
        |> List.exists (fun line -> WordLayout.Page.pageTopForY (pageHeight settings) line.Y <> firstPage)

type TextMeasurer = string -> TextBuffer.TextStyle -> float
type TextRangeMeasurer = string -> TextBuffer.TextStyle -> int -> int -> float

let rangeMeasurerFromTextMeasurer (measureText: TextMeasurer) : TextRangeMeasurer =
    fun text style start length ->
        let startIndex = max 0 (min text.Length start)
        let endIndex = max startIndex (min text.Length (start + length))
        if endIndex = startIndex then 0.0
        else measureText (text.Substring(startIndex, endIndex - startIndex)) style

let tableColumnWidths (maxWidth: float) (tableBlock: TableBlock) =
    WordLayout.Table.tableColumnWidths WordLayout.printScale (Some maxWidth) tableBlock

// Word line boxes include font leading beyond the nominal font size.
let private wordLineBoxFactor = 1.20
let private monospaceLineBoxFactor = 1.18
let private tableCellLineBoxFactor = 1.22

let private lineBoxFactorForTextStyle (style: TextBuffer.TextStyle) =
    match style.FontFamily |> Option.map (fun family -> family.Trim().ToLowerInvariant()) with
    | Some "consolas"
    | Some "courier new"
    | Some "courier" -> monospaceLineBoxFactor
    | _ -> wordLineBoxFactor

let private wordLineHeight textStyle fontSize lineSpacing =
    fontSize * lineSpacing * lineBoxFactorForTextStyle textStyle

let private lineHeightForSettings (settings: LayoutSettings) textStyle fontSize lineSpacing lineSpacingRule =
    let automatic = if settings.InTableCell then fontSize * lineSpacing * tableCellLineBoxFactor else wordLineHeight textStyle fontSize lineSpacing
    match lineSpacingRule with
    | "exact" -> max 1.0 lineSpacing
    | "atLeast" -> max automatic lineSpacing
    | _ -> automatic

let private paragraphBorderAdvance (border: ParagraphBorder option) =
    match border with
    | Some value ->
        match value.Style |> Option.map (fun style -> style.ToLowerInvariant()) with
        | Some "none"
        | Some "nil" -> 0.0
        | _ -> (value.Width |> Option.defaultValue 0.5 |> max 0.25) + (value.Space |> Option.defaultValue 0.0)
    | None -> 0.0

let private paragraphBorderVerticalAdvance metadata =
    match metadata.ParagraphBorders with
    | Some borders -> paragraphBorderAdvance borders.Top, paragraphBorderAdvance borders.Bottom
    | None -> 0.0, 0.0

let private paragraphBorderRightAdvance metadata =
    match metadata.ParagraphBorders with
    | Some borders -> paragraphBorderAdvance borders.Right
    | None -> 0.0

let segmentHeight lineHeightFactor = function
    | TextSegment (_, style, _, _, _) -> WordLayout.Text.displayFontSize style * lineHeightFactor
    | ImageSegment (_, _, yOffset, _, height, _, floating, _, _, _, _, _, _) ->
        if floating then 0.0 else height + max 0.0 yOffset

let segmentStart = function
    | TextSegment (_, _, _, _, startOffset) -> startOffset
    | ImageSegment (_, _, _, _, _, startOffset, _, _, _, _, _, _, _) -> startOffset

let segmentEnd = function
    | TextSegment (text, _, _, _, startOffset) -> startOffset + text.Length
    | ImageSegment (_, _, _, _, _, startOffset, _, _, _, _, _, _, _) -> startOffset + 1

let segmentX = function
    | TextSegment (_, _, x, _, _) -> x
    | ImageSegment (_, x, _, _, _, _, _, _, _, _, _, _, _) -> x

let segmentWidth = function
    | TextSegment (_, _, _, width, _) -> width
    | ImageSegment (_, _, _, width, _, _, _, _, _, _, _, _, _) -> width

let alignmentOffset (layout: BlockLayout) (line: MeasuredLine) =
    match layout.Metadata.Alignment with
    | Left -> 0.0
    | Center -> max 0.0 ((layout.AvailableWidth - line.Width) / 2.0)
    | Right -> max 0.0 (layout.AvailableWidth - line.Width)
    | Justify -> 0.0

let emptyLine y height : MeasuredLine = {
    Segments = []
    Y = y
    Width = 0.0
    Height = height
    StartOffset = 0
    EndOffset = 0
}

let lineForOffset (layout: BlockLayout) (offset: int) : MeasuredLine =
    match layout.Lines with
    | [] -> emptyLine layout.Y layout.Height
    | first :: _ when offset <= first.StartOffset -> first
    | lines ->
        match lines |> List.tryFind (fun line -> offset >= line.StartOffset && offset < line.EndOffset) with
        | Some line -> line
        | None ->
            match lines |> List.rev |> List.tryFind (fun line -> offset >= line.EndOffset) with
            | Some line -> line
            | None -> List.last lines

let maxOffsetInLayout (layout: BlockLayout) =
    match layout.Lines with
    | [] -> 0
    | lines -> lines |> List.map (fun l -> l.EndOffset) |> List.max

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
                found <- true
            else
                x <- segmentX segment + segmentWidth segment

    x

let xForOffset (measureText: TextMeasurer) (line: MeasuredLine) (offset: int) =
    xForOffsetWithRange (rangeMeasurerFromTextMeasurer measureText) line offset

let hitTestLayoutsWithRange (measureRange: TextRangeMeasurer) (paddingLeft: float) (layouts: BlockLayout list) (mouseX: float) (mouseY: float) =
    let blockOpt = layouts |> List.tryFind (fun layout ->
        mouseY >= layout.Y && mouseY < layout.Y + layout.Height)

    match blockOpt with
    | None -> None
    | Some layout ->
        let lineOpt = layout.Lines |> List.tryFind (fun line ->
            mouseY >= line.Y && mouseY < line.Y + line.Height)

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
                            let mutable low = 0
                            let mutable high = text.Length
                            while low < high do
                                let mid = (low + high) / 2
                                let textWidth = measureRange (WordLayout.Text.displayText style text) (WordLayout.Text.displayStyle style) 0 mid
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

            Some (layout.BlockId, offset)

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
    result

let imageBoxForOffset (paddingLeft: float) (layouts: BlockLayout list) (blockId: BlockId) (offset: int) : (float * float * float * float) option =
    match layouts |> List.tryFind (fun l -> l.BlockId = blockId) with
    | None -> None
    | Some layout ->
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

type FloatingObstacle = float * float * float * float
type PlacedFloatingBox = WordLayout.Floating.Box

let private isWhitespaceChar (ch: char) =
    ch = ' ' || ch = '\t' || ch = '\r' || ch = '\n'

let private trailingSpaceFitSafety = 1.0
let private lineFitSafety = 0.25
let private tableCellLineFitSafety = 2.0
let private mixedMetricsLineFitSafety = 0.75

let private visibleTokenLength (text: string) =
    let mutable length = text.Length
    while length > 0 && isWhitespaceChar text.[length - 1] do
        length <- length - 1
    length

let private hasMixedTextMetrics (segments: Segment list) =
    let keys =
        segments
        |> List.choose (function
            | TextSegment (_, style, _, _, _) -> Some (style.FontFamily, style.FontSize, style.VertAlign, style.SmallCaps)
            | ImageSegment _ -> None)
        |> List.distinct
    keys.Length > 1

let private wordWrapTokens (text: string) =
    let tokens = ResizeArray<string>()
    let mutable index = 0
    while index < text.Length do
        let start = index
        if text.[index] = '\r' || text.[index] = '\n' then
            if text.[index] = '\r' && index + 1 < text.Length && text.[index + 1] = '\n' then
                index <- index + 2
            else
                index <- index + 1
            tokens.Add(text.Substring(start, index - start))
        elif isWhitespaceChar text.[index] then
            while index < text.Length && isWhitespaceChar text.[index] && text.[index] <> '\r' && text.[index] <> '\n' do
                index <- index + 1
            tokens.Add(text.Substring(start, index - start))
        else
            let mutable keepGoing = true
            while keepGoing && index < text.Length && not (isWhitespaceChar text.[index]) do
                let isHyphenBreak = text.[index] = '-' && index + 1 < text.Length && not (isWhitespaceChar text.[index + 1])
                index <- index + 1
                if isHyphenBreak then
                    keepGoing <- false
            while index < text.Length && isWhitespaceChar text.[index] && text.[index] <> '\r' && text.[index] <> '\n' do
                index <- index + 1
            tokens.Add(text.Substring(start, index - start))
    tokens

/// Layout buffer into lines with styled runs and images.
let layoutLinesStyledWithObstaclesWithRange (settings: LayoutSettings) (measureRange: TextRangeMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) (contentY: float) (rightTabPosition: float option) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    let mutable lines = []
    let mutable currentSegments: Segment list = []
    let mutable currentX = 0.0
    let mutable y = minBodyTop settings 0.0
    let mutable flushedLineCount = 0
    let availableWidth = max 1.0 maxWidth
    let mutable currentLineEnd = availableWidth
    let mutable floatObstacles = initialObstacles
    let mutable placedBoxes = initialPlacedBoxes

    let mergeIntervals intervals =
        let sorted = intervals |> List.sortBy fst
        let rec loop merged remaining =
            match merged, remaining with
            | _, [] -> List.rev merged
            | [], interval :: rest -> loop [interval] rest
            | (startX, endX) :: prior, (nextStart, nextEnd) :: rest when nextStart <= endX ->
                loop ((startX, max endX nextEnd) :: prior) rest
            | _, interval :: rest -> loop (interval :: merged) rest
        loop [] sorted

    let wrapBoundsForY lineY =
        let globalLineY = contentY + (lineY - settings.PaddingTop)
        let active =
            floatObstacles
            |> List.filter (fun (_, top, _, height) -> globalLineY >= top && globalLineY < top + height)
            |> List.map (fun (x, _, width, _) -> x, x + width)
            |> mergeIntervals

        let lineStart =
            active
            |> List.choose (fun (x, right) ->
                if x + (right - x) / 2.0 <= availableWidth / 2.0 then Some right
                else None)
            |> function
                | [] -> 0.0
                | values -> List.max values

        let lineEnd =
            active
            |> List.choose (fun (x, right) ->
                if x + (right - x) / 2.0 > availableWidth / 2.0 && x > lineStart then Some x
                else None)
            |> function
                | [] -> availableWidth
                | values -> List.min values

        lineStart, max lineStart lineEnd

    let moveToWrappableLine () =
        let lineStart, lineEnd = wrapBoundsForY y
        currentX <- lineStart
        currentLineEnd <- lineEnd
        while currentX >= currentLineEnd do
            y <- y + settings.LineHeight
            let lineStart, lineEnd = wrapBoundsForY y
            currentX <- lineStart
            currentLineEnd <- lineEnd

    let flushLine () =
        if not currentSegments.IsEmpty then
            let segments = List.rev currentSegments
            let hasInlineImage =
                segments
                |> List.exists (function
                    | ImageSegment (_, _, _, _, _, _, floating, _, _, _, _, _, _) -> not floating
                    | TextSegment _ -> false)
            let inlineImageDescent = if hasInlineImage then settings.FontSize * 0.27 else 0.0
            let lineHeight =
                (segments |> List.map (segmentHeight (settings.LineHeight / settings.FontSize)) |> List.max) + inlineImageDescent
                |> max settings.LineHeight
            lines <- lines @ [{
                Segments = segments
                Y = y
                Width = currentX
                Height = lineHeight
                StartOffset = segments |> List.head |> segmentStart
                EndOffset = segments |> List.last |> segmentEnd
            }]
            currentSegments <- []
            flushedLineCount <- flushedLineCount + 1
            y <- y + lineHeight
            moveToWrappableLine ()

    let tabWidth () =
        let defaultTab = (floor (currentX / 48.0) + 1.0) * 48.0
        let target = rightTabPosition |> Option.defaultValue defaultTab |> min currentLineEnd
        max 12.0 (target - currentX - 24.0)

    moveToWrappableLine ()

    for run in buffer.Runs do
        match run with
        | TextBuffer.Image imgRun ->
            let imgWidth, imgHeight = WordLayout.Image.printSize imgRun.Width imgRun.Height

            if currentX + imgWidth > currentLineEnd && not currentSegments.IsEmpty then
                flushLine ()

            let rawFloatingBox = WordLayout.Floating.imageCanvasBox settings.PaddingLeft contentY currentX y settings.PaddingTop imgWidth imgHeight imgRun
            let floatingBox = if imgRun.AllowOverlap = Some false then WordLayout.Floating.avoidOverlap placedBoxes rawFloatingBox else rawFloatingBox
            let imageX = if imgRun.Floating then floatingBox.X else currentX
            let yOffset = if imgRun.Floating then floatingBox.Y - (contentY + (y - settings.PaddingTop)) else 0.0
            let segment = ImageSegment (imgRun.Src, imageX, yOffset, imgWidth, imgHeight, imgRun.Start, imgRun.Floating, imgRun.Wrap, imgRun.RelativeHeight, imgRun.CropLeft, imgRun.CropTop, imgRun.CropRight, imgRun.CropBottom)
            currentSegments <- segment :: currentSegments
            if imgRun.Floating && WordLayout.Floating.participatesInWrap floatingBox then
                let obstacleX, obstacleTop, obstacleWidth, obstacleHeight = WordLayout.Floating.obstacleForBox availableWidth floatingBox
                if WordLayout.Floating.wrapsTopAndBottom floatingBox.Wrap then
                    floatObstacles <- (0.0, obstacleTop, availableWidth, obstacleHeight) :: floatObstacles
                    currentX <- availableWidth
                    currentLineEnd <- availableWidth
                else
                    floatObstacles <- (obstacleX, obstacleTop, obstacleWidth, obstacleHeight) :: floatObstacles
                    let lineStart, lineEnd = wrapBoundsForY y
                    currentX <- lineStart
                    currentLineEnd <- lineEnd
            else
                currentX <- currentX + imgWidth
            if imgRun.Floating then placedBoxes <- floatingBox :: placedBoxes

        | TextBuffer.Text textRun ->
            let runText = TextBuffer.slice textRun.Start textRun.Length buffer
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

            for textToAdd in wordWrapTokens runText do

                if String.length textToAdd > 0 then
                    if textToAdd = "\n" || textToAdd = "\r" || textToAdd = "\r\n" then
                        let hadSegments = not currentSegments.IsEmpty
                        flushLine ()
                        if not hadSegments then
                            y <- y + settings.LineHeight
                            moveToWrappableLine ()
                        consumed <- consumed + textToAdd.Length
                    else
                        let mutable remaining = textToAdd
                        while remaining.Length > 0 do
                            if currentSegments.IsEmpty && currentX >= currentLineEnd then
                                moveToWrappableLine ()

                            let width =
                                if remaining.Contains("\t") then tabWidth ()
                                else measureRange displayRunText displayRunStyle consumed remaining.Length
                            let fitWidth =
                                if remaining.Contains("\t") then width
                                else
                                    let visibleLength = visibleTokenLength remaining
                                    let canUseVisibleFit = flushedLineCount > 0 || hasMixedTextMetrics currentSegments
                                    if canUseVisibleFit && visibleLength > 0 && visibleLength < remaining.Length then
                                        let visibleWidth = measureRange displayRunText displayRunStyle consumed visibleLength
                                        if currentLineEnd - (currentX + visibleWidth) >= trailingSpaceFitSafety then visibleWidth else width
                                    else width

                            let fullFitSafety =
                                if settings.InTableCell then
                                    if hasMixedTextMetrics currentSegments then max tableCellLineFitSafety mixedMetricsLineFitSafety else tableCellLineFitSafety
                                elif hasMixedTextMetrics currentSegments then mixedMetricsLineFitSafety
                                else lineFitSafety

                            if currentX + fitWidth > currentLineEnd - fullFitSafety && not currentSegments.IsEmpty then
                                flushLine ()
                            else
                                let segmentText, segmentWidth =
                                    if currentX + fitWidth > currentLineEnd - fullFitSafety && not (remaining.Contains("\t")) then
                                        let maxSegmentWidth = max 1.0 (currentLineEnd - currentX)
                                        let take = fittingLength consumed remaining.Length maxSegmentWidth
                                        remaining.Substring(0, take), measureRange displayRunText displayRunStyle consumed take
                                    else
                                        remaining, width
                                let segment = TextSegment (segmentText, textRun.Style, currentX, segmentWidth, textRun.Start + consumed)
                                currentSegments <- segment :: currentSegments
                                currentX <- currentX + segmentWidth
                                consumed <- consumed + segmentText.Length
                                remaining <- remaining.Substring(segmentText.Length)
                                if remaining.Length > 0 then flushLine ()

    flushLine ()
    lines, floatObstacles, placedBoxes

let layoutLinesStyledWithObstacles (settings: LayoutSettings) (measureText: TextMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) (contentY: float) (rightTabPosition: float option) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    layoutLinesStyledWithObstaclesWithRange settings (rangeMeasurerFromTextMeasurer measureText) buffer maxWidth contentY rightTabPosition initialObstacles initialPlacedBoxes

let layoutLinesStyled (settings: LayoutSettings) (measureText: TextMeasurer) (buffer: TextBuffer.TextBuffer) (maxWidth: float) =
    let lines, _, _ = layoutLinesStyledWithObstacles settings measureText buffer maxWidth 0.0 None [] []
    lines

let layoutParagraphAtWithStyleAndObstaclesWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (maxWidth: float) (y: float) (xOffset: float) (paragraphStyle: ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) (para: ParagraphBlock) (numberedCounters: Map<string * int, int>) (initialObstacles: FloatingObstacle list) (initialPlacedBoxes: PlacedFloatingBox list) =
    let mutable counters = numberedCounters
    let metadata = WordStyles.effectiveMetadata paragraphStyle para.Metadata
    let indent = WordLayout.Paragraph.leftIndent metadata
    let rightIndent = WordLayout.Paragraph.rightIndent metadata
    let firstLineIndent = WordLayout.Paragraph.firstLineIndent metadata
    let rawSpacingBefore = WordLayout.Paragraph.spacingBefore paragraphStyle metadata
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    let spacingBefore =
        if y <= WordLayout.Page.contentTop settings.PaddingTop pageTop + pageFitTolerance then 0.0
        else rawSpacingBefore
    let spacingAfter = WordLayout.Paragraph.spacingAfter paragraphStyle metadata
    let lineSpacing = WordLayout.Paragraph.lineSpacing paragraphStyle metadata
    let lineSpacingRule = WordStyles.lineSpacingRule paragraphStyle metadata
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
    let contentFontSize =
        buffer.Runs
        |> List.choose (function
            | Text run -> Some (WordLayout.Text.displayFontSize run.Style)
            | Image _ -> None)
        |> function
            | [] ->
                metadata.ParagraphMarkFontSize
                |> Option.filter (fun fontSize -> fontSize <= paragraphStyle.Text.FontSize - 2.0)
                |> Option.defaultValue paragraphStyle.Text.FontSize
            | sizes -> List.max sizes
    let lineSettings = { settings with FontSize = contentFontSize; LineHeight = lineHeightForSettings settings paragraphStyle.Text contentFontSize lineSpacing lineSpacingRule }
    let lines, updatedObstacles, updatedPlacedBoxes = layoutLinesStyledWithObstaclesWithRange lineSettings measureRange buffer availableWidth contentY metadata.RightTabPosition initialObstacles initialPlacedBoxes

    let adjustedLines =
        lines |> List.map (fun line ->
            let lineFirstIndent =
                if line = (lines |> List.head) then
                    if isHangingList then 0.0 else firstLineIndent
                else 0.0
            let adjustedSegments =
                line.Segments |> List.map (function
                    | TextSegment (text, style, x, width, startOffset) ->
                        TextSegment (text, style, x + xOffset + indent + markerReserve + lineFirstIndent, width, startOffset)
                    | ImageSegment (src, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom) ->
                        ImageSegment (src, x + xOffset + indent + markerReserve + lineFirstIndent, yOffset, width, height, startOffset, floating, wrap, relativeHeight, cropLeft, cropTop, cropRight, cropBottom))
            { line with
                Y = contentY + (line.Y - settings.PaddingTop)
                Segments = adjustedSegments })

    let lineHeightWithMarker line =
        match prefixStyle |> Option.bind WordLayout.Numbering.markerLineBoxFactor with
        | Some markerFactor when line.StartOffset = 0 ->
            max line.Height (contentFontSize * lineSpacing * markerFactor)
        | _ -> line.Height

    let adjustedLines =
        adjustedLines
        |> List.map (fun line -> { line with Height = lineHeightWithMarker line })

    let pagedLines = paginateLines settings (widowControl paragraphStyle metadata) adjustedLines

    let blockHeight =
        if pagedLines.IsEmpty then lineHeightForSettings settings paragraphStyle.Text contentFontSize lineSpacing lineSpacingRule
        else
            let lastLine = List.last pagedLines
            (lastLine.Y - contentY) + lastLine.Height

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

let measureTableRowWithStyle (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (tableBlock: TableBlock) (tableMaxWidth: float) (columnWidths: float list) (outerLeft: float) (row: TableRow) rowY tableContinuationHeaderHeight (paragraphStyleForParagraph: ParagraphBlock -> ParagraphStyle) (effectiveBufferForParagraph: ParagraphStyle -> TextBuffer -> TextBuffer) =
    let cellSettings = {
        settings with
            InTableCell = true
            MarginBottom = if row.CanSplit = Some false then -1000000.0 else settings.MarginBottom
            TableContinuationHeaderHeight = tableContinuationHeaderHeight
    }
    let numColumns = WordLayout.Table.columnCount tableBlock
    let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
    let outerStartInset = outerLeft + cellSpacing
    let mutable rowHeight = WordLayout.Table.rowMinHeight WordLayout.printScale settings.LineHeight row
    let mutable cellGroups = []
    let mutable traces = []

    for positioned in WordLayout.Table.positionedCells row do
        let cell = positioned.Cell
        let fallbackWidth = tableMaxWidth / float numColumns
        let colWidth = WordLayout.Table.positionedSpacedCellWidth columnWidths cellSpacing fallbackWidth positioned
        let padding = WordLayout.Table.cellPaddingForRow WordLayout.canvasScale tableBlock row cell
        let cellX = WordLayout.Table.positionedSpacedCellX columnWidths cellSpacing positioned + padding.Left - outerStartInset
        let cellMaxWidth = WordLayout.Table.cellContentWidth colWidth padding
        let cellY = rowY + cellSpacing
        let mutable paraY = WordLayout.Table.cellContentStartY cellY padding
        let mutable cellLayouts = []

        let cellParagraphs = cell.Content |> List.toArray
        let mutable pendingEmptyHeight = 0.0
        for paraIndex in 0 .. cellParagraphs.Length - 1 do
            let para = tableCellParagraphForLayout cellParagraphs.[paraIndex]
            let paragraphStyle = paragraphStyleForParagraph para
            let layout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange cellSettings measureText measureRange cellMaxWidth paraY cellX paragraphStyle effectiveBufferForParagraph para Map.empty [] []
            let layout = if layout.Lines.IsEmpty then moveLayoutToY paraY layout else layout
            let layout =
                match layout.Lines with
                | firstLine :: _ when pendingEmptyHeight > 0.0 && pageNumberForY cellSettings layout.Y <> pageNumberForY cellSettings firstLine.Y ->
                    { layout with
                        Height = layout.Height + pendingEmptyHeight
                        Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + pendingEmptyHeight }) }
                | _ -> layout
            cellLayouts <- cellLayouts @ [layout]
            pendingEmptyHeight <- if layout.Lines.IsEmpty then pendingEmptyHeight + layout.Height else 0.0
            paraY <- paraY + layout.Height

        let contentHeight = WordLayout.Table.cellContentHeight cellY padding.Top paraY
        rowHeight <- max rowHeight (contentHeight + padding.Top + padding.Bottom + cellSpacing)
        cellGroups <- cellGroups @ [{ VerticalAlignment = cell.VerticalAlignment; PaddingTop = padding.Top; PaddingBottom = padding.Bottom; ContentHeight = contentHeight; Layouts = cellLayouts }]
        let totalLines = cellLayouts |> List.sumBy (fun layout -> layout.Lines.Length)
        let cellSnippet =
            cell.Content
            |> List.map (fun para -> TextBuffer.toString para.Buffer)
            |> List.tryFind (fun text -> text.Trim() <> "")
            |> Option.map snippet
            |> Option.defaultValue ""
        traces <- traces @ [sprintf "c%d/%d w=%.2f cw=%.2f h=%.2f lines=%d text=%s" positioned.StartColumn positioned.Span colWidth cellMaxWidth contentHeight totalLines cellSnippet]

    { Height = rowHeight; CellGroups = cellGroups; Traces = traces }

let measureTableRowForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row rowY tableContinuationHeaderHeight =
    measureTableRowWithStyle settings measureText measureRange tableBlock tableMaxWidth columnWidths outerLeft row rowY tableContinuationHeaderHeight (forTableCellParagraphInDocument doc) (effectiveBufferInDocument doc)

let measureTableRowForDocumentRowLocal settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row =
    let rowLocalSettings = {
        settings with
            PageHeight = 1000000.0
            MarginBottom = 0.0
            TableContinuationHeaderHeight = 0.0
    }
    measureTableRowForDocument rowLocalSettings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row 0.0 0.0

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
    measurement.CellGroups
    |> List.collect (fun group ->
        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment rowHeight group.PaddingTop group.PaddingBottom group.ContentHeight
        group.Layouts
        |> List.choose (fun layout ->
            if fragment.BlockIds |> List.contains layout.BlockId then
                let shiftedLayout =
                    { layout with
                        Y = layout.Y + offset + contentOriginDelta
                        Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + offset + contentOriginDelta }) }
                let visibleLines =
                    shiftedLayout.Lines
                    |> List.filter (fun line -> line.Y >= fragment.Y && line.Y + line.Height <= fragmentBottom)
                match visibleLines with
                | [] -> None
                | lines ->
                    let y = lines |> List.map _.Y |> List.min
                    let bottom = lines |> List.map (fun line -> line.Y + line.Height) |> List.max
                    Some { shiftedLayout with Lines = lines; Y = y; Height = bottom - y }
            else
                None))

let tableRowFragmentLayouts (fragment: TableRowFragmentPlan) (measurement: TableRowMeasurement) =
    tableRowFragmentLayoutsWithHeight measurement.Height fragment measurement

let tableCellFragmentPlans (plan: TableLayoutPlan) =
    let tableBlock = plan.Table
    let numColumns = WordLayout.Table.columnCount tableBlock
    let resolvedRows =
        WordLayout.Table.resolveTable WordLayout.printScale (Some plan.TableMaxWidth) tableBlock
        |> fun resolved -> resolved.BodyRows |> List.toArray
    plan.Rows
    |> List.collect (fun planRow ->
        let row = tableBlock.Rows.[planRow.RowIndex]
        let positionedCells = WordLayout.Table.positionedCells row |> List.toArray
        let cellGroups = planRow.Measurement.CellGroups |> List.toArray
        let resolvedCells =
            resolvedRows
            |> Array.tryItem planRow.RowIndex
            |> Option.map (fun row -> row.Cells |> List.toArray)
            |> Option.defaultValue [||]
        let cellSpacing = WordLayout.Table.rowCellSpacing WordLayout.canvasScale tableBlock row
        planRow.Fragments
        |> List.collect (fun fragment ->
            let fragmentEnd = fragment.Y + fragment.Height
            let contentOriginDelta = fragment.RowContentY - fragment.RowPlanY
            [0 .. min positionedCells.Length cellGroups.Length - 1]
            |> List.map (fun cellIndex ->
                let positioned = positionedCells.[cellIndex]
                let group = cellGroups.[cellIndex]
                let resolvedCell = resolvedCells |> Array.tryItem cellIndex
                let fallbackWidth = plan.TableMaxWidth / float numColumns
                let cellX = plan.TableX + WordLayout.Table.positionedSpacedCellX plan.ColumnWidths cellSpacing positioned
                let cellWidth = WordLayout.Table.positionedSpacedCellWidth plan.ColumnWidths cellSpacing fallbackWidth positioned
                let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment planRow.Height group.PaddingTop group.PaddingBottom group.ContentHeight
                let visibleLines =
                    group.Layouts
                    |> List.filter (fun layout -> fragment.BlockIds |> List.contains layout.BlockId)
                    |> List.collect (fun layout ->
                        layout.Lines
                        |> List.map (fun line -> { line with Y = line.Y + offset + contentOriginDelta }))
                    |> List.filter (fun line -> line.Y >= fragment.Y && line.Y + line.Height <= fragmentEnd)
                let visibleY = visibleLines |> List.map _.Y |> function | [] -> None | values -> Some (List.min values)
                let visibleEnd = visibleLines |> List.map (fun line -> line.Y + line.Height) |> function | [] -> None | values -> Some (List.max values)
                let visibleHeight =
                    match visibleY, visibleEnd with
                    | Some startY, Some endY -> Some (endY - startY)
                    | _ -> None
                let blockIds =
                    group.Layouts
                    |> List.filter (fun layout -> fragment.BlockIds |> List.contains layout.BlockId)
                    |> List.map _.BlockId
                    |> List.distinct
                {
                    RowIndex = planRow.RowIndex
                    CellIndex = cellIndex
                    StartColumn = resolvedCell |> Option.map _.StartColumn |> Option.defaultValue positioned.StartColumn
                    ColumnSpan = resolvedCell |> Option.map _.ColSpan |> Option.defaultValue positioned.Span
                    RowSpan = resolvedCell |> Option.map _.RowSpan |> Option.defaultValue 1
                    BorderTop = resolvedCell |> Option.bind (fun cell -> cell.Borders.Top)
                    BorderRight = resolvedCell |> Option.bind (fun cell -> cell.Borders.Right)
                    BorderBottom = resolvedCell |> Option.bind (fun cell -> cell.Borders.Bottom)
                    BorderLeft = resolvedCell |> Option.bind (fun cell -> cell.Borders.Left)
                    Page = fragment.Page
                    RowY = planRow.Y
                    RowHeight = planRow.Height
                    FragmentY = fragment.Y
                    FragmentHeight = fragment.Height
                    CellX = cellX
                    CellWidth = cellWidth
                    CellY = fragment.Y + cellSpacing
                    CellHeight = max 0.0 (fragment.Height - cellSpacing)
                    CellSpacing = cellSpacing
                    PaddingTop = group.PaddingTop
                    PaddingBottom = group.PaddingBottom
                    ContentY = fragment.RowContentY + cellSpacing + group.PaddingTop + offset
                    ContentHeight = group.ContentHeight
                    VisibleY = visibleY
                    VisibleHeight = visibleHeight
                    VisibleLineCount = visibleLines.Length
                    BlockIds = blockIds
                })))

let private borderGridStroke priority (stroke: WordLayout.Table.PrintStroke option) =
    stroke
    |> Option.map (fun stroke -> ({
        Color = stroke.Color
        Width = stroke.Width
        Style = stroke.Style
        Priority = priority
    }: RedBat.BorderGrid.Stroke))

let tableCellFragmentBorderCells (cellFragments: TableCellFragmentPlan list) =
    cellFragments
    |> List.map (fun fragment -> ({
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
    }: RedBat.BorderGrid.Cell))

let planTableRowFragments (settings: LayoutSettings) firstTablePage repeatedHeaderHeight rowY rowHeight (measurement: TableRowMeasurement) =
    let rowBottom = rowY + rowHeight
    let lines =
        measurement.CellGroups
        |> List.collect (fun group -> group.Layouts)
        |> List.collect (fun layout -> layout.Lines |> List.map (fun line -> layout.BlockId, line))
        |> List.sortBy (fun (_, line) -> line.Y, line.StartOffset, line.EndOffset)
    let mutable fragments = []
    let mutable consumedLineStart = 0
    let mutable pageTop = WordLayout.Page.pageTopForY (pageHeight settings) rowY
    while pageTop < rowBottom do
        let fragmentTop = max rowY (WordLayout.Page.contentTop settings.PaddingTop pageTop)
        let fragmentBottom = min rowBottom (WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop)
        if fragmentBottom > fragmentTop then
            let fragmentLines =
                lines
                |> List.filter (fun (_, line) -> line.Y < fragmentBottom && line.Y + line.Height > fragmentTop)
            let consumedLineEnd = consumedLineStart + fragmentLines.Length
            let contentTop = fragmentLines |> List.map (fun (_, line) -> line.Y) |> function | [] -> None | values -> Some (List.min values)
            let contentBottom = fragmentLines |> List.map (fun (_, line) -> line.Y + line.Height) |> function | [] -> None | values -> Some (List.max values)
            let blockIds = fragmentLines |> List.map fst |> List.distinct
            let page = pageNumberForY settings fragmentTop
            let repeatedHeaderHeight = if page > firstTablePage then repeatedHeaderHeight else 0.0
            fragments <- fragments @ [{
                Page = page
                RowPlanY = rowY
                RowContentY = rowY
                Y = fragmentTop
                Height = fragmentBottom - fragmentTop
                AvailableHeight = WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop - fragmentTop
                RepeatedHeaderHeight = repeatedHeaderHeight
                ContentTop = contentTop
                ContentBottom = contentBottom
                LineCount = fragmentLines.Length
                ConsumedLineStart = consumedLineStart
                ConsumedLineEnd = consumedLineEnd
                BlockIds = blockIds
                BreakReason = "row-span"
            }]
            consumedLineStart <- consumedLineEnd
        pageTop <- pageTop + pageHeight settings + pageGap
    fragments

let private contentBottomForY settings y =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    WordLayout.Page.contentBottom (pageHeight settings) settings.MarginBottom pageTop

let private nextPageContentY settings y repeatedHeaderHeight =
    let pageTop = WordLayout.Page.pageTopForY (pageHeight settings) y
    WordLayout.Page.contentTop settings.PaddingTop (pageTop + pageHeight settings + pageGap) + repeatedHeaderHeight

let private fragmentPlanForRowTransaction settings firstTablePage headerHeight rowY rowHeight (row: TableRow) (rowLocalMeasurement: TableRowMeasurement) =
    let rowLines =
        rowLocalMeasurement.CellGroups
        |> List.collect (fun group -> group.Layouts)
        |> List.collect (fun layout -> layout.Lines |> List.map (fun line -> layout.BlockId, line))
        |> List.sortBy (fun (_, line) -> line.Y, line.StartOffset, line.EndOffset)
        |> List.toArray
    let mutable fragments = []
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
        let fragmentLines = rowLines.[consumedLineStart .. consumedLineEnd - 1] |> Array.toList
        let contentTop = fragmentLines |> List.map (fun (_, line) -> currentY + line.Y) |> function | [] -> None | values -> Some (List.min values)
        let contentBottom = fragmentLines |> List.map (fun (_, line) -> currentY + line.Y + line.Height) |> function | [] -> None | values -> Some (List.max values)
        let blockIds = fragmentLines |> List.map fst |> List.distinct
        fragments <- fragments @ [{
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
            BlockIds = blockIds
            BreakReason = breakReason
        }]
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
    currentY, fragments

let buildTableLayoutPlanForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableY =
    let outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock |> snd
    let tableX = settings.PaddingLeft - outerLeft
    let tableWidth = WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
    let headerRows = WordLayout.Table.headerRows tableBlock
    let headerRowCount = headerRows |> List.length
    let measureRow row rowY tableContinuationHeaderHeight =
        measureTableRowForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row rowY tableContinuationHeaderHeight
    let measureRowLocal row =
        measureTableRowForDocumentRowLocal settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row
    let headerHeightAt rowY =
        headerRows
        |> List.fold (fun (currentY, totalHeight) headerRow ->
            let header = measureRow headerRow currentY 0.0
            currentY + header.Height, totalHeight + header.Height) (rowY, 0.0)
        |> snd
    let mutable rowY = tableY
    let firstTablePage = pageNumberForY settings tableY
    let mutable rows = []
    for rowIndex in 0 .. tableBlock.Rows.Length - 1 do
        let row = tableBlock.Rows.[rowIndex]
        let continuationHeaderHeight = if row.IsHeader then 0.0 else headerHeightAt rowY
        let rowLocalMeasurement = measureRowLocal row
        let mutable plannedMeasurement = moveTableRowMeasurementToY rowY rowLocalMeasurement
        let mutable plannedY = rowY
        let mutable rowHeight = rowLocalMeasurement.Height
        let pagedRowY = paginateTableRow settings row rowY rowHeight
        let rowContentY =
            if pagedRowY <> rowY && not row.IsHeader && not headerRows.IsEmpty then
                pagedRowY + headerHeightAt pagedRowY
            else
                pagedRowY
        let contentBottom = contentBottomForY settings rowY
        let rowFitsAtCurrentY = rowContentY = rowY && rowY + rowHeight <= contentBottom
        let rowFitsAfterMove = rowContentY <> rowY
        if rowFitsAtCurrentY || rowFitsAfterMove then
            plannedY <- rowContentY
            plannedMeasurement <- moveTableRowMeasurementToY rowContentY rowLocalMeasurement
        else
            plannedMeasurement <- measureRow row rowY continuationHeaderHeight
            rowHeight <- plannedMeasurement.Height
        let repeatedHeaderHeight = if row.IsHeader then 0.0 else headerHeightAt rowY
        let fragments = planTableRowFragments settings (pageNumberForY settings plannedY) repeatedHeaderHeight plannedY rowHeight plannedMeasurement
        rows <- rows @ [{
            RowIndex = rowIndex
            IsHeader = row.IsHeader
            CanSplit = row.CanSplit |> Option.defaultValue true
            Y = plannedY
            Height = rowHeight
            Measurement = plannedMeasurement
            Fragments = fragments
        }]
        rowY <- plannedY + rowHeight
    {
        Table = tableBlock
        TableX = tableX
        TableMaxWidth = tableMaxWidth
        TableWidth = tableWidth
        OuterLeft = outerLeft
        OuterRight = outerRight
        ColumnWidths = columnWidths
        HeaderRowCount = headerRowCount
        Rows = rows
        NextY = rowY + WordLayout.Table.afterSpacing
    }

let buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY =
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
    let tableMaxWidth = maxWidth + outerLeft + outerRight
    let columnWidths = tableColumnWidths tableMaxWidth tableBlock
    buildTableLayoutPlanForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft tableY

/// Layout document blocks. Table cell paragraphs are included so shared hit/cursor logic works inside tables.
let layoutDocumentWithRange (settings: LayoutSettings) (measureText: TextMeasurer) (measureRange: TextRangeMeasurer) (doc: Document) =
    let maxWidth = settings.Width - settings.PaddingLeft - settings.MarginRight
    let mutable y = settings.PaddingTop
    let mutable layouts = []
    let mutable numberedCounters = Map.empty<string * int, int>
    let mutable footnotesByPage: Map<int, (string * string) list> = Map.empty
    let mutable floatingObstacles: FloatingObstacle list = []
    let mutable placedFloatingBoxes: PlacedFloatingBox list = []
    let mutable previousParagraphContext: (ParagraphBlock * ParagraphStyle * BlockMetadata * float) option = None
    let blocks = doc.Blocks |> List.toArray

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

    let sameParagraphBorders (previous: BlockMetadata) (current: BlockMetadata) =
        previous.ParagraphBorders.IsSome && previous.ParagraphBorders = current.ParagraphBorders

    let resetPreviousParagraph () =
        previousParagraphContext <- None

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
                let outerLeft, outerRight = WordLayout.Table.outerCellInsets WordLayout.canvasScale tableBlock
                let tableMaxWidth = maxWidth + outerLeft + outerRight
                let columnWidths = tableColumnWidths tableMaxWidth tableBlock
                let mutable rowY = tableY
                let headerRows = WordLayout.Table.headerRows tableBlock
                let traceRows = traceTimingEnabled && tableBlock.Rows.Length >= 10
                let tracePlanRows =
                    if traceRows then
                        let plan = buildTableLayoutPlanForDocumentAt settings measureText measureRange doc tableBlock tableY
                        plan.Rows |> List.map (fun row -> row.RowIndex, row) |> Map.ofList |> Some
                    else
                        None

                if traceRows then
                    traceTableLayout (sprintf "start rows=%d y=%.2f page=%d widths=%s" tableBlock.Rows.Length tableY (pageNumberForY settings tableY) (columnWidths |> List.map (sprintf "%.2f") |> String.concat ","))

                let measureRow (row: TableRow) rowY tableContinuationHeaderHeight =
                    measureTableRowForDocument settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row rowY tableContinuationHeaderHeight

                let measureRowLocal (row: TableRow) =
                    measureTableRowForDocumentRowLocal settings measureText measureRange doc tableBlock tableMaxWidth columnWidths outerLeft row

                let headerHeightAt rowY =
                    headerRows
                    |> List.fold (fun (currentY, totalHeight) headerRow ->
                        let header = measureRow headerRow currentY 0.0
                        currentY + header.Height, totalHeight + header.Height) (rowY, 0.0)
                    |> snd

                let addHeaderLayoutsAt headerY =
                    let mutable headerRowY = headerY
                    for headerRow in headerRows do
                        let headerMeasurement = measureRow headerRow headerRowY 0.0
                        for group in headerMeasurement.CellGroups do
                            let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment headerMeasurement.Height group.PaddingTop group.PaddingBottom group.ContentHeight
                            let adjustedLayouts =
                                group.Layouts
                                |> List.filter (fun layout -> not layout.Lines.IsEmpty)
                                |> List.map (fun layout ->
                                    { layout with
                                        Y = layout.Y + offset
                                        Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + offset }) })
                            layouts <- layouts @ adjustedLayouts
                            footnotesByPage <- adjustedLayouts |> List.fold (addLayoutFootnotes settings) footnotesByPage
                        headerRowY <- headerRowY + headerMeasurement.Height

                for rowIndex in 0 .. tableBlock.Rows.Length - 1 do
                    let row = tableBlock.Rows.[rowIndex]
                    let continuationHeaderHeight = if row.IsHeader then 0.0 else headerHeightAt rowY
                    let rowLocalMeasurement = measureRowLocal row
                    let mutable measuredRow = rowLocalMeasurement
                    let mutable rowHeight = measuredRow.Height
                    let mutable cellLayoutGroups = moveTableRowMeasurementToY rowY rowLocalMeasurement |> fun measurement -> measurement.CellGroups

                    let pagedRowY = paginateTableRow settings row rowY rowHeight
                    let pagedRowY = paginateTableRowWithFootnotes settings footnotesByPage row pagedRowY rowHeight
                    let rowContentY =
                        if pagedRowY <> rowY && not row.IsHeader && not headerRows.IsEmpty then
                            pagedRowY + headerHeightAt pagedRowY
                        else
                            pagedRowY
                    let contentBottom = contentBottomWithFootnotes settings footnotesByPage rowY
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
                            |> List.tryFind (fun text -> text.Trim() <> "")
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

                    for group in cellLayoutGroups do
                        let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment rowHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                        let adjustedLayouts =
                            group.Layouts
                            |> List.filter (fun layout -> not layout.Lines.IsEmpty)
                            |> List.map (fun layout ->
                                { layout with
                                    Y = layout.Y + offset
                                    Lines = layout.Lines |> List.map (fun line -> { line with Y = line.Y + offset }) })
                        layouts <- layouts @ adjustedLayouts
                        footnotesByPage <- adjustedLayouts |> List.fold (addLayoutFootnotes settings) footnotesByPage
                    rowY <- rowY + rowHeight

                rowY + WordLayout.Table.afterSpacing, WordLayout.Table.tableWidthWithSpacing columnWidths (WordLayout.Table.cellSpacing WordLayout.canvasScale tableBlock)
            else
                tableY + WordLayout.Table.emptyTableAdvance settings.LineHeight, 0.0
        traceSlow (sprintf "table rows=%d" tableBlock.Rows.Length) tableStart
        result

    for index in 0 .. blocks.Length - 1 do
        let block = blocks.[index]
        match block with
        | TextBox textBox ->
            resetPreviousParagraph ()
            addTextBoxObstacle textBox
        | Table tableBlock ->
            resetPreviousParagraph ()
            let blockStart = timingStart ()
            let nextY, tableWidth = layoutTableAt tableBlock y
            if tableBlock.Floating then
                let tableHeight = max 0.0 (nextY - y - WordLayout.Table.afterSpacing)
                floatingObstacles <- (0.0, y, tableWidth + 12.0, tableHeight) :: floatingObstacles
                placedFloatingBoxes <- {
                    Kind = WordLayout.Floating.TableBox tableBlock
                    X = settings.PaddingLeft
                    Y = y
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
            traceSlow (sprintf "block[%d]:table rows=%d" index tableBlock.Rows.Length) blockStart

        | _ ->
            let blockStart = timingStart ()
            let para = getParagraph block
            let paragraphStyle = forBlockInDocument doc block
            let beforeRenderedHintY = y
            y <- applyRenderedPageBreakHint settings y para
            if y <> beforeRenderedHintY then resetPreviousParagraph ()
            if (para.Metadata.PageBreakBefore || paragraphStyle.PageBreakBefore = Some true) && not layouts.IsEmpty then
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
            let initialLayout, initialCounters, initialObstacles, initialPlacedBoxes = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth y 0.0 paragraphStyle (effectiveBufferInDocument doc) para numberedCounters floatingObstacles placedFloatingBoxes
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
                            let nextStyle = forBlockInDocument doc nextBlock
                            if nextPara.Metadata.PageBreakBefore || nextStyle.PageBreakBefore = Some true then
                                groupIsMoveable <- false
                                shouldContinue <- false
                            else
                                let nextLayout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth probeY 0.0 nextStyle (effectiveBufferInDocument doc) nextPara numberedCounters initialObstacles initialPlacedBoxes
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
            let initialFitHeight = paragraphVisibleFitHeight paragraphStyle initialLayout
            let pagedY = if initialLinesSpanPages then keepWithNextY else paginateBlock settings keepWithNextY initialFitHeight
            let pagedY = if initialLinesSpanPages then pagedY else paginateBlockWithFootnotes settings footnotesByPage pagedY initialFitHeight
            let nextLineCount =
                if not initialLinesSpanPages && shouldProbeNearFooterNextLineCount pagedY initialLayout.Height initialLayout.Lines.Length initialLayout.Metadata.SpacingAfter then
                    match nextParagraphBlock index with
                    | Some (nextBlock, nextPara) ->
                        let nearFooterStart = timingStart ()
                        let nextStyle = forBlockInDocument doc nextBlock
                        let nextLayout, _, _, _ = layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth (pagedY + initialLayout.Height) 0.0 nextStyle (effectiveBufferInDocument doc) nextPara numberedCounters initialObstacles initialPlacedBoxes
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
                        layoutParagraphAtWithStyleAndObstaclesWithRange settings measureText measureRange maxWidth pagedY 0.0 paragraphStyle (effectiveBufferInDocument doc) para numberedCounters floatingObstacles placedFloatingBoxes
                    traceSlow (sprintf "block[%d]:relayout" index) relayoutStart
                    result
            numberedCounters <- counters
            floatingObstacles <- obstacles
            placedFloatingBoxes <- placedBoxes
            layouts <- layouts @ [layout]
            footnotesByPage <- addLayoutFootnotes settings footnotesByPage layout
            y <- layout.Y + layout.Height
            previousParagraphContext <- Some (para, paragraphStyle, layout.Metadata, WordLayout.Paragraph.spacingAfter paragraphStyle layout.Metadata)
            if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                y <- advanceToNextPage settings y
                resetPreviousParagraph ()
            traceSlow (sprintf "block[%d]:paragraph total lines=%d" index layout.Lines.Length) blockStart

    layouts

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
