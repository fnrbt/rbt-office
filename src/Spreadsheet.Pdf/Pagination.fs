module RedBat.Spreadsheet.Pdf.Pagination

open RedBat.Spreadsheet.Model

/// 2D page-break computation for spreadsheet printing.
///
/// Excel splits a sheet across pages in *both* directions: rows that don't
/// fit vertically push to a new page below, columns that don't fit
/// horizontally push to a new page to the right. Each page is therefore a
/// (row-range, col-range) slice. This module computes those slices from a
/// `PageSetup` — honoring manual breaks, the print area, repeat
/// rows/cols, and fit-to-page scaling — leaving the actual rendering to
/// the PDF backend.

/// Print-headings chrome metrics, in CSS pixels. When PrintOptions.Headings
/// is on, a strip of this height sits above the body (column letters) and
/// a strip of this width sits to its left (row numbers); both shrink the
/// per-page content area before manual breaks / fit-to-page kick in. Kept
/// here so the PDF emitter and Pagination agree on the same metrics.
let printHeadingsRowHeightPx = 19.0
let printHeadingsColWidthPx = 28.0

let headingsOffsets (ps: PageSetup) : float * float =
    if ps.Options.Headings
    then printHeadingsColWidthPx, printHeadingsRowHeightPx
    else 0.0, 0.0

/// One page in the print layout — an inclusive 0-indexed rectangle.
type PageSlice = {
    RowStart: int
    RowEnd: int
    ColStart: int
    ColEnd: int
}

/// Full pagination for one sheet. Pages are emitted in the "down then
/// over" order Excel uses by default: every row page within the first
/// column slice, then the next column slice, etc.
type Pagination = {
    Pages: PageSlice list
    /// The print bounds — either the explicit `PrintArea` from PageSetup
    /// or the model's used range when none is set.
    PrintBounds: CellRange
    /// Scale (1.0 = no scale). Honors fit-to-page or the manual scale.
    Scale: float
    /// Repeat rows / cols, copied from PageSetup so the emitter doesn't
    /// have to re-resolve them.
    RepeatRows: (int * int) option
    RepeatCols: (int * int) option
}

type private PaginationCache = {
    Cells: Map<CellAddress, CellData>
    RowHeights: Map<int, float>
    ColumnWidths: Map<int, float>
    PageSetup: PageSetup
    CellCount: int
    RowHeightCount: int
    ColumnWidthCount: int
    SourceTextCharCount: int
    Pagination: Pagination
}

let mutable private paginationCache: PaginationCache list = []
let private paginationCacheMaxSources = 4
let private paginationCacheMaxTotalCells = 250000
let private paginationCacheMaxTotalRowHeights = 200000
let private paginationCacheMaxTotalColumnWidths = 200000
let private paginationCacheMaxTotalSourceTextChars = 5000000

let private paginationCacheMatches (model: SpreadsheetModel) (cached: PaginationCache) =
    obj.ReferenceEquals(cached.Cells, model.Cells)
    && obj.ReferenceEquals(cached.RowHeights, model.RowHeights)
    && obj.ReferenceEquals(cached.ColumnWidths, model.ColumnWidths)
    && cached.PageSetup = model.PageSetup

let private optionStringCharCount (value: string option) =
    match value with
    | Some value -> value.Length
    | None -> 0

let private paginationSourceTextAddCapped amount total =
    if amount <= 0 then total
    elif total > paginationCacheMaxTotalSourceTextChars
         || amount > paginationCacheMaxTotalSourceTextChars - total then
        paginationCacheMaxTotalSourceTextChars + 1
    else
        total + amount

let private headerFooterStringCharCount (headerFooter: HeaderFooter) =
    0
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.OddHeader)
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.OddFooter)
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.EvenHeader)
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.EvenFooter)
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.FirstHeader)
    |> paginationSourceTextAddCapped (optionStringCharCount headerFooter.FirstFooter)

let private borderStringCharCount (border: BorderStyle option) =
    match border with
    | Some border -> border.Style.Length + optionStringCharCount border.Color
    | None -> 0

let rec private cellValueStringCharCount = function
    | CellValue.String value
    | CellValue.Error value -> value.Length
    | CellValue.Formula (formula, cachedValue) -> formula.Length + cellValueStringCharCount cachedValue
    | CellValue.Empty
    | CellValue.Number _
    | CellValue.Boolean _ -> 0

let private cellFormatStringCharCount (format: CellFormat) =
    optionStringCharCount format.FontName
    + optionStringCharCount format.TextColor
    + optionStringCharCount format.BgColor
    + optionStringCharCount format.HorizontalAlign
    + optionStringCharCount format.VerticalAlign
    + optionStringCharCount format.NumberFormat
    + borderStringCharCount format.BorderTop
    + borderStringCharCount format.BorderBottom
    + borderStringCharCount format.BorderLeft
    + borderStringCharCount format.BorderRight
    + borderStringCharCount format.BorderDiagonal
    + (match format.DataBar with Some (_, color) -> color.Length | None -> 0)

let private cellCommentStringCharCount (comment: CellComment option) =
    match comment with
    | Some comment -> comment.Text.Length + optionStringCharCount comment.Author
    | None -> 0

let private cellDataStringCharCount (cell: CellData) =
    cellValueStringCharCount cell.Value
    + cellFormatStringCharCount cell.Format
    + cellCommentStringCharCount cell.Comment

let private sourceTextCharCount (pageSetup: PageSetup) (cells: Map<CellAddress, CellData>) =
    let mutable total = headerFooterStringCharCount pageSetup.HeaderFooter
    for KeyValue(_, cell) in cells do
        if total <= paginationCacheMaxTotalSourceTextChars then
            total <- paginationSourceTextAddCapped (cellDataStringCharCount cell) total
    total

let private sourceTextCharCountFor cellCount rowHeightCount columnWidthCount pageSetup cells =
    if cellCount > paginationCacheMaxTotalCells
       || rowHeightCount > paginationCacheMaxTotalRowHeights
       || columnWidthCount > paginationCacheMaxTotalColumnWidths then
        paginationCacheMaxTotalSourceTextChars + 1
    else
        sourceTextCharCount pageSetup cells

let private paginationCacheEntryRetainable (cached: PaginationCache) =
    cached.CellCount <= paginationCacheMaxTotalCells
    && cached.RowHeightCount <= paginationCacheMaxTotalRowHeights
    && cached.ColumnWidthCount <= paginationCacheMaxTotalColumnWidths
    && cached.SourceTextCharCount <= paginationCacheMaxTotalSourceTextChars

let private rememberPaginationCache cached =
    let trim entries =
        let rec loop keptCount keptCells keptRowHeights keptColumnWidths keptTextChars acc remaining =
            match remaining with
            | [] -> List.rev acc
            | candidate :: tail ->
                let candidateCells = max 1 candidate.CellCount
                let candidateRowHeights = max 1 candidate.RowHeightCount
                let candidateColumnWidths = max 1 candidate.ColumnWidthCount
                let candidateTextChars = max 1 candidate.SourceTextCharCount
                if paginationCacheEntryRetainable candidate
                   && keptCount < paginationCacheMaxSources
                   && keptCells + candidateCells <= paginationCacheMaxTotalCells
                   && keptRowHeights + candidateRowHeights <= paginationCacheMaxTotalRowHeights
                   && keptColumnWidths + candidateColumnWidths <= paginationCacheMaxTotalColumnWidths
                   && keptTextChars + candidateTextChars <= paginationCacheMaxTotalSourceTextChars then
                    loop
                        (keptCount + 1)
                        (keptCells + candidateCells)
                        (keptRowHeights + candidateRowHeights)
                        (keptColumnWidths + candidateColumnWidths)
                        (keptTextChars + candidateTextChars)
                        (candidate :: acc)
                        tail
                else
                    loop keptCount keptCells keptRowHeights keptColumnWidths keptTextChars acc tail
        loop 0 0 0 0 0 [] entries
    paginationCache <-
        (cached
         :: (paginationCache
             |> List.filter (fun existing ->
                 not (obj.ReferenceEquals(existing.Cells, cached.Cells)
                      && obj.ReferenceEquals(existing.RowHeights, cached.RowHeights)
                      && obj.ReferenceEquals(existing.ColumnWidths, cached.ColumnWidths)
                      && existing.PageSetup = cached.PageSetup))))
        |> trim

let private colWidthPx (model: SpreadsheetModel) (col: int) : float =
    Map.tryFind col model.ColumnWidths |> Option.defaultValue 100.0

let private rowHeightPx (model: SpreadsheetModel) (row: int) : float =
    Map.tryFind row model.RowHeights |> Option.defaultValue 25.0

let private rangeSum (sizes: int -> float) (startInc: int) (endInc: int) : float =
    let mutable s = 0.0
    for i in startInc .. endInc do s <- s + sizes i
    s

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private breakSet (breaks: int list) =
    let mutable result = Set.empty
    for breakIndex in breaks do
        result <- Set.add (breakIndex - 1) result
    result

/// Bounding box of every non-empty cell. Empty model collapses to A1 so
/// downstream code has a non-degenerate range to work with.
let private usedRange (model: SpreadsheetModel) : CellRange =
    if Map.isEmpty model.Cells then
        { From = { Row = 0; Col = 0 }; To = { Row = 0; Col = 0 } }
    else
        let mutable r1 = System.Int32.MaxValue
        let mutable c1 = System.Int32.MaxValue
        let mutable r2 = System.Int32.MinValue
        let mutable c2 = System.Int32.MinValue
        for KeyValue(addr, _) in model.Cells do
            if addr.Row < r1 then r1 <- addr.Row
            if addr.Row > r2 then r2 <- addr.Row
            if addr.Col < c1 then c1 <- addr.Col
            if addr.Col > c2 then c2 <- addr.Col
        { From = { Row = r1; Col = c1 }; To = { Row = r2; Col = c2 } }

/// Split one axis of a print range into per-page slices. Walks the axis
/// in order, accumulating cell sizes until the next one would overflow
/// `available` — at which point the current page closes and a new one
/// starts. Manual breaks force a close *after* the given 0-indexed slot
/// regardless of accumulated size. A single cell larger than `available`
/// is placed on its own page rather than dropped.
let private axisBreaks
    (startInc: int)
    (endInc: int)
    (manualBreaks0: Set<int>)
    (sizes: int -> float)
    (available: float)
    : (int * int) list =
    if startInc > endInc then [] else
    let result = ResizeArray<int * int>()
    let mutable pageStart = startInc
    let mutable acc = 0.0
    for i in startInc .. endInc do
        let sz = sizes i
        // Manual break sitting *after* index (i - 1) closes the page before
        // we start placing i.
        if i > pageStart && Set.contains (i - 1) manualBreaks0 then
            result.Add(pageStart, i - 1)
            pageStart <- i
            acc <- 0.0
        if acc + sz > available && i > pageStart then
            result.Add(pageStart, i - 1)
            pageStart <- i
            acc <- sz
        else
            acc <- acc + sz
    if pageStart <= endInc then result.Add(pageStart, endInc)
    resizeArrayToList result

/// Convert the model's PageSetup into the full pagination. Repeat rows /
/// cols steal space from every page (so the available area shrinks);
/// `Scale` applies *after* break computation, so a small scale means more
/// content fits and breaks land later — we honor that by inflating the
/// available size by 1/scale before computing the breaks.
let private paginateUncached (model: SpreadsheetModel) : Pagination =
    let ps = model.PageSetup
    let printRange =
        match ps.PrintArea with
        | Some range -> range
        | None -> usedRange model

    let leftPt = ps.Margins.Left * 72.0
    let rightPt = ps.Margins.Right * 72.0
    let topPt = ps.Margins.Top * 72.0
    let bottomPt = ps.Margins.Bottom * 72.0
    let pxPerPt = 96.0 / 72.0
    let printableWidthPx = max 50.0 ((ps.PageWidth - leftPt - rightPt) * pxPerPt)
    let printableHeightPx = max 50.0 ((ps.PageHeight - topPt - bottomPt) * pxPerPt)

    let repeatRowsHeight =
        ps.RepeatRows
        |> Option.map (fun (a, b) -> rangeSum (rowHeightPx model) a b)
        |> Option.defaultValue 0.0
    let repeatColsWidth =
        ps.RepeatCols
        |> Option.map (fun (a, b) -> rangeSum (colWidthPx model) a b)
        |> Option.defaultValue 0.0

    let headingsColPx, headingsRowPx = headingsOffsets ps
    let availableW = max 50.0 (printableWidthPx - repeatColsWidth - headingsColPx)
    let availableH = max 50.0 (printableHeightPx - repeatRowsHeight - headingsRowPx)

    // Resolve scale: explicit fit-to-page wins over the manual scale percent.
    // For fit-to-page we measure the print range against `FitToWidth` ×
    // `FitToHeight` pages of available area. Clamp to ≤ 1 (we shrink to
    // fit but don't blow up to fill — matching Excel).
    let scale =
        if ps.FitToPage then
            let contentW = rangeSum (colWidthPx model) printRange.From.Col printRange.To.Col
            let contentH = rangeSum (rowHeightPx model) printRange.From.Row printRange.To.Row
            let sx =
                if ps.FitToWidth > 0 && contentW > 0.0
                then (availableW * float ps.FitToWidth) / contentW else 1.0
            let sy =
                if ps.FitToHeight > 0 && contentH > 0.0
                then (availableH * float ps.FitToHeight) / contentH else 1.0
            min sx sy |> min 1.0 |> max 0.1
        else
            float ps.Scale / 100.0 |> max 0.1

    // A smaller scale fits more content per page — equivalent to having a
    // larger effective break budget. Inflate the available size before
    // computing breaks rather than scaling every measurement individually.
    let effW = availableW / scale
    let effH = availableH / scale

    let manualRow = breakSet ps.ManualRowBreaks
    let manualCol = breakSet ps.ManualColBreaks

    // Body range excludes the repeat rows / cols — they ride along on every
    // page separately. Without this the first row / column would render
    // twice on page 1 (once as the repeat band, once as the first body row).
    let bodyRowStart =
        match ps.RepeatRows with
        | Some (_, rEnd) when rEnd >= printRange.From.Row -> max printRange.From.Row (rEnd + 1)
        | _ -> printRange.From.Row
    let bodyColStart =
        match ps.RepeatCols with
        | Some (_, cEnd) when cEnd >= printRange.From.Col -> max printRange.From.Col (cEnd + 1)
        | _ -> printRange.From.Col

    let colRanges =
        axisBreaks bodyColStart printRange.To.Col manualCol (colWidthPx model) effW
    let rowRanges =
        axisBreaks bodyRowStart printRange.To.Row manualRow (rowHeightPx model) effH

    // Walk order from PageSetup. DownThenOver (Excel default) emits all
    // row-pages of the first column-slice before the next column-slice;
    // OverThenDown inverts the outer/inner loops.
    let pages =
        let result = ResizeArray<PageSlice>()
        match ps.PageOrder with
        | DownThenOver ->
            for (cs, ce) in colRanges do
                for (rs, re) in rowRanges do
                    result.Add { RowStart = rs; RowEnd = re; ColStart = cs; ColEnd = ce }
        | OverThenDown ->
            for (rs, re) in rowRanges do
                for (cs, ce) in colRanges do
                    result.Add { RowStart = rs; RowEnd = re; ColStart = cs; ColEnd = ce }
        resizeArrayToList result

    let result =
        { Pages = pages
          PrintBounds = printRange
          Scale = scale
          RepeatRows = ps.RepeatRows
          RepeatCols = ps.RepeatCols }
    result

let paginate (model: SpreadsheetModel) : Pagination =
    match paginationCache |> List.tryFind (paginationCacheMatches model) with
    | Some cached ->
        rememberPaginationCache cached
        cached.Pagination
    | None ->
        let result = paginateUncached model
        let cellCount = Map.count model.Cells
        let rowHeightCount = Map.count model.RowHeights
        let columnWidthCount = Map.count model.ColumnWidths
        rememberPaginationCache {
            Cells = model.Cells
            RowHeights = model.RowHeights
            ColumnWidths = model.ColumnWidths
            PageSetup = model.PageSetup
            CellCount = cellCount
            RowHeightCount = rowHeightCount
            ColumnWidthCount = columnWidthCount
            SourceTextCharCount = sourceTextCharCountFor cellCount rowHeightCount columnWidthCount model.PageSetup model.Cells
            Pagination = result
        }
        result
