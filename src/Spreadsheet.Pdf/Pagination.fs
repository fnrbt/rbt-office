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

let private colWidthPx (model: SpreadsheetModel) (col: int) : float =
    Map.tryFind col model.ColumnWidths |> Option.defaultValue 100.0

let private rowHeightPx (model: SpreadsheetModel) (row: int) : float =
    Map.tryFind row model.RowHeights |> Option.defaultValue 25.0

let private rangeSum (sizes: int -> float) (startInc: int) (endInc: int) : float =
    let mutable s = 0.0
    for i in startInc .. endInc do s <- s + sizes i
    s

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
    let result = System.Collections.Generic.List<int * int>()
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
    List.ofSeq result

/// Convert the model's PageSetup into the full pagination. Repeat rows /
/// cols steal space from every page (so the available area shrinks);
/// `Scale` applies *after* break computation, so a small scale means more
/// content fits and breaks land later — we honor that by inflating the
/// available size by 1/scale before computing the breaks.
let paginate (model: SpreadsheetModel) : Pagination =
    let ps = model.PageSetup
    let printRange =
        ps.PrintArea
        |> Option.defaultValue (usedRange model)

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

    let availableW = max 50.0 (printableWidthPx - repeatColsWidth)
    let availableH = max 50.0 (printableHeightPx - repeatRowsHeight)

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

    let manualRow = ps.ManualRowBreaks |> List.map (fun b -> b - 1) |> Set.ofList
    let manualCol = ps.ManualColBreaks |> List.map (fun b -> b - 1) |> Set.ofList

    let colRanges =
        axisBreaks printRange.From.Col printRange.To.Col manualCol (colWidthPx model) effW
    let rowRanges =
        axisBreaks printRange.From.Row printRange.To.Row manualRow (rowHeightPx model) effH

    // Down then over: walk row pages within each column page.
    let pages =
        [ for (cs, ce) in colRanges do
            for (rs, re) in rowRanges ->
                { RowStart = rs; RowEnd = re; ColStart = cs; ColEnd = ce } ]

    { Pages = pages
      PrintBounds = printRange
      Scale = scale
      RepeatRows = ps.RepeatRows
      RepeatCols = ps.RepeatCols }
