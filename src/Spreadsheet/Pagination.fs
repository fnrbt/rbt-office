module Rbt.Spreadsheet.Pdf.Pagination

open Rbt.Spreadsheet.Model

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

/// The paper as it is printed on: (width, height) in points. A setup's
/// `PageWidth` / `PageHeight` are the paper's in portrait where the setup
/// came from a file through the editor, and already turned where a host put
/// the turned size in; a landscape page is the wider way round either way.
let orientedPageSize (ps: PageSetup) : float * float =
    match ps.Orientation with
    | LandscapeOrientation when ps.PageWidth < ps.PageHeight -> ps.PageHeight, ps.PageWidth
    | _ -> ps.PageWidth, ps.PageHeight

/// One page in the print layout — an inclusive 0-indexed rectangle.
type PageSlice = {
    RowStart: int
    RowEnd: int
    ColStart: int
    ColEnd: int
    /// Which of the print areas the page is of (`Pagination.PrintAreas`, 0-based).
    Area: int
    /// The page carries the repeat rows / columns (`Pagination.RepeatRows` / `RepeatCols`) above / beside its own: it
    /// begins past them. A page that begins at or before them has them where they are (vertex42's MortgageCalculator,
    /// rows 1:410 with row 49 to repeat: its first page is rows 1-60, row 49 among them; the pages after begin with
    /// row 49).
    TitleRows: bool
    TitleCols: bool
}

/// Full pagination for one sheet. Pages are emitted in the "down then
/// over" order Excel uses by default: every row page within the first
/// column slice, then the next column slice, etc.
type Pagination = {
    Pages: PageSlice list
    /// The print bounds — either the explicit `PrintArea` from PageSetup
    /// or the model's used range when none is set.
    PrintBounds: CellRange
    /// Every area that is printed, in order: the areas of the sheet's print
    /// area (one, or several - each paginated on its own, its pages following
    /// those of the area before), or the used range when none is set.
    /// `PrintBounds` is the first.
    PrintAreas: CellRange list
    /// Scale (1.0 = no scale). Honors fit-to-page or the manual scale.
    Scale: float
    /// Repeat rows / cols, copied from PageSetup so the emitter doesn't
    /// have to re-resolve them.
    RepeatRows: (int * int) option
    RepeatCols: (int * int) option
}

let private colWidthPx (model: SpreadsheetModel) (col: int) : float =
    Map.tryFind col model.ColumnWidths |> Option.defaultValue model.DefaultColumnWidth

/// (A row a filter hides prints nothing, as one hidden: houseflipping's "2 REPAIR ESTIMATOR", 1,051 rows hidden in its
/// AutoFilter's range, fits 1 x 6 pages on 2 in Excel 2019.)
let private rowHeightPx (model: SpreadsheetModel) (row: int) : float =
    if isFilterHiddenRow row model then 0.0
    else Map.tryFind row model.RowHeights |> Option.defaultValue model.DefaultRowHeight

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

/// Whether a cell makes Excel print as far as it: it holds something (a value, a formula) or shows a fill or a
/// border. A font, a number format, an alignment alone do not (make-print-pages-fixture.py "range": BoldPast and
/// FormatPast print A1:C5, FillPast and BorderPast A1:F12; bramosman's Amortization, formatted down to row 383,
/// prints to its last payment, row 194).
let private printsCell (data: CellData) =
    (match data.Value with CellValue.Empty | CellValue.String "" -> false | _ -> true)
    || data.Format.BgColor.IsSome
    || data.Format.BorderTop.IsSome || data.Format.BorderBottom.IsSome || data.Format.BorderLeft.IsSome
    || data.Format.BorderRight.IsSome || data.Format.BorderDiagonal.IsSome

/// The last column the text of a cell runs into (Excel prints it), or the cell's own column: text wider than its cell
/// runs over the empty cells to its right (`Rbt.Spreadsheet.Layout.textOverflowSpan`), measured in Excel's widths of the
/// cell's font at 100 % (`ExcelText.textWidth`) and scaled as the columns are for print.
let private overflowEnd (model: SpreadsheetModel) (addr: CellAddress) (data: CellData) =
    match Rbt.Spreadsheet.Layout.textOverflowSides data.Format data.Value with
    | Some (_, true) ->
        match unwrapFormulaValue data.Value with
        | CellValue.String text ->
            let format = data.Format
            let points = defaultArg format.FontSize 11.0
            let px = Rbt.Spreadsheet.ExcelText.pixelSize points 1.0
            let family = defaultArg format.FontName "Calibri"
            let screen = Rbt.Spreadsheet.ExcelText.textWidth (fun t -> float t.Length * float px * 0.5) family format.Bold format.Italic px false text
            // (The columns here are as wide as they print: the text is too.)
            let mdw = if model.MaxDigitWidth > 0.0 then model.MaxDigitWidth else 7.0
            let width = screen * (if model.PrintDigitWidth > 0.0 then model.PrintDigitWidth / mdw / 6.25 else 1.0)
            let room = getColumnWidth addr.Col model - Rbt.Spreadsheet.FontMetrics.columnPaddingPx mdw
            let excess = width - room
            if excess <= 0.0 then addr.Col
            else snd (Rbt.Spreadsheet.Layout.textOverflowSpan model addr.Row addr.Col 0.0 excess)
        | _ -> addr.Col
    | _ -> addr.Col

/// What Excel 2019 prints of a sheet without a print area: from A1 - whatever the first row and column with something
/// in them (make-print-pages-fixture.py "range": StartC5's C5:F12 prints A1:F12; acre's MF Rents pages break at
/// columns J, Q, ... from A, not from its first column, B) - to the last row and column with a cell that prints
/// (`printsCell`; a formula that gives "" too: FormulaBlank), a drawing, a merged cell that holds something or shows
/// a fill (MergePast, MergeFill: E10:H14), or text running over from a cell to its left (Overflow: C7's sentence
/// prints to F7, OverflowFar's over three pages; acre-property's Summary). None when nothing prints: Excel has nothing
/// to print ("Empty", acre's "Calculation Modules-->"), and a workbook's PDF leaves the sheet out.
let private usedRange (model: SpreadsheetModel) : CellRange option =
    let mutable r2 = -1
    let mutable c2 = -1
    let cover (row: int) (col: int) =
        if row > r2 then r2 <- row
        if col > c2 then c2 <- col
        if r2 < 0 then r2 <- 0
        if c2 < 0 then c2 <- 0
    // (A row's own fill or border - `<row s customFormat>` - prints the rows down to it and across to the last column
    // with a cell of a format of its own, whatever it is: RowFill, values to C, prints A1:C12; choudhary's Rental
    // Assumptions, a filled row 1 and number formats to U, prints to U. A column's prints neither: ColFill, A1:C5. A
    // cell the sheet has only for its row's style is the row's.)
    let ofRowStyle (addr: CellAddress) (data: CellData) =
        (match data.Value with CellValue.Empty -> true | _ -> false)
        && (match Map.tryFind addr.Row model.RowStyles with Some format -> format = data.Format | None -> false)
    let isPrintingFormat (format: CellFormat) = printsCell { Value = CellValue.Empty; Format = format; Comment = None }
    if model.RowStyles |> Map.exists (fun _ format -> isPrintingFormat format) then
        let mutable lastFormatted = 0
        for KeyValue(addr, data) in model.Cells do
            if addr.Col > lastFormatted && data.Format <> defaultFormat && not (ofRowStyle addr data) then lastFormatted <- addr.Col
        for KeyValue(row, format) in model.RowStyles do
            if isPrintingFormat format then cover row lastFormatted
    for KeyValue(addr, data) in model.Cells do
        if printsCell data && not (ofRowStyle addr data) then
            cover addr.Row (overflowEnd model addr data)
    for merge in model.MergedCells do
        match Map.tryFind { Row = merge.StartRow; Col = merge.StartCol } model.Cells with
        | Some data when printsCell data -> cover merge.EndRow merge.EndCol
        | _ -> ()
    for drawing in model.Drawings do
        let anchor = drawing.Anchor
        cover anchor.FromRow anchor.FromCol
        // (A drawing that ends on the edge of a cell does not reach into it.)
        let lastOf (first: int) (last: int option) (offset: float) =
            match last with
            | Some index -> max first (if offset > 0.0 then index else index - 1)
            | None -> first
        cover (lastOf anchor.FromRow anchor.ToRow anchor.ToOffsetY) (lastOf anchor.FromCol anchor.ToCol anchor.ToOffsetX)
    if r2 < 0 then None else Some { From = { Row = 0; Col = 0 }; To = { Row = r2; Col = c2 } }

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
    (available: int -> float)
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
        if acc + sz > available pageStart && i > pageStart then
            result.Add(pageStart, i - 1)
            pageStart <- i
            acc <- sz
        else
            acc <- acc + sz
    if pageStart <= endInc then result.Add(pageStart, endInc)
    resizeArrayToList result

/// Convert the model's PageSetup into the full pagination. Repeat rows /
/// cols steal space from every page (so the available area shrinks); the
/// rows and columns are laid out at their sizes times the scale.
let paginate (model: SpreadsheetModel) : Pagination =
    let ps = model.PageSetup
    let printRanges =
        match printAreasOf ps with
        | [] -> Option.toList (usedRange model)
        | areas -> areas
    let printRange = match printRanges with first :: _ -> first | [] -> { From = { Row = 0; Col = 0 }; To = { Row = 0; Col = 0 } }

    let leftPt = ps.Margins.Left * 72.0
    let rightPt = ps.Margins.Right * 72.0
    let topPt = ps.Margins.Top * 72.0
    let bottomPt = ps.Margins.Bottom * 72.0
    let pxPerPt = 96.0 / 72.0
    let pageWidth, pageHeight = orientedPageSize ps
    let printableWidthPx = max 50.0 ((pageWidth - leftPt - rightPt) * pxPerPt)
    let printableHeightPx = max 50.0 ((pageHeight - topPt - bottomPt) * pxPerPt)

    let repeatRowsHeight =
        ps.RepeatRows
        |> Option.map (fun (a, b) -> rangeSum (rowHeightPx model) a b)
        |> Option.defaultValue 0.0
    let repeatColsWidth =
        ps.RepeatCols
        |> Option.map (fun (a, b) -> rangeSum (colWidthPx model) a b)
        |> Option.defaultValue 0.0

    let headingsColPx, headingsRowPx = headingsOffsets ps
    let manualRow = breakSet ps.ManualRowBreaks
    let manualCol = breakSet ps.ManualColBreaks

    // The body of each area is the whole area; a page that begins past the repeat rows / cols carries them.
    let bodyOf (printRange: CellRange) = printRange.From.Row, printRange.From.Col
    let titlesBefore (titles: (int * int) option) (pageStart: int) =
        match titles with
        | Some (first, last) when last >= first -> pageStart > last
        | _ -> false

    // At a scale, Excel 2019 sets each row and column at its printed size times the scale in whole dots of 600 dpi
    // and fills a page to a pixel short of its printable width / height (make-print-pages-fixture.py "fit": at 90 %
    // ten of Calibri 11's 430-dot columns fit 511.2 pt, not eleven - 4,257 dots of 4,260; at 55 % 84 rows of 67 dots
    // a page, not 85).
    let dotted (size: float) (scale: float) = System.Math.Floor (size * scale * 6.25 + 0.5) / 6.25
    let rangesAt (scale: float) (printRange: CellRange) =
        let bodyRowStart, bodyColStart = bodyOf printRange
        let availableW (pageStart: int) =
            max 50.0 (printableWidthPx - (if titlesBefore ps.RepeatCols pageStart then dotted repeatColsWidth scale else 0.0) - headingsColPx) - 1.0
        let availableH (pageStart: int) =
            max 50.0 (printableHeightPx - (if titlesBefore ps.RepeatRows pageStart then dotted repeatRowsHeight scale else 0.0) - headingsRowPx) - 1.0
        let colRanges = axisBreaks bodyColStart printRange.To.Col manualCol (fun c -> dotted (colWidthPx model c) scale) availableW
        let rowRanges = axisBreaks bodyRowStart printRange.To.Row manualRow (fun r -> dotted (rowHeightPx model r) scale) availableH
        colRanges, rowRanges

    // Fit to pages: a whole percent, the largest at which every area takes no more than `FitToWidth` pages across and
    // `FitToHeight` down - from the share of a page (a pixel short of it) the area's width / height needs, cut to a
    // whole percent, and down from there while it takes more (Fit10's 99.07 % is 98: its 99 % takes the whole width;
    // FitTall, 130 rows on 2 pages at 72 %). The explicit scale otherwise.
    let scale =
        if ps.FitToPage then
            let share (printRange: CellRange) =
                let bodyRowStart, bodyColStart = bodyOf printRange
                let contentW = rangeSum (colWidthPx model) bodyColStart printRange.To.Col
                let contentH = rangeSum (rowHeightPx model) bodyRowStart printRange.To.Row
                let availableW = max 50.0 (printableWidthPx - repeatColsWidth - headingsColPx) - 1.0
                let availableH = max 50.0 (printableHeightPx - repeatRowsHeight - headingsRowPx) - 1.0
                let sx = if ps.FitToWidth > 0 && contentW > 0.0 then (availableW * float ps.FitToWidth) / contentW else 1.0
                let sy = if ps.FitToHeight > 0 && contentH > 0.0 then (availableH * float ps.FitToHeight) / contentH else 1.0
                min sx sy
            let first = printRanges |> List.map share |> List.fold min 1.0
            let fits (percent: int) =
                printRanges |> List.forall (fun printRange ->
                    let cols, rows = rangesAt (float percent / 100.0) printRange
                    (ps.FitToWidth <= 0 || cols.Length <= ps.FitToWidth) && (ps.FitToHeight <= 0 || rows.Length <= ps.FitToHeight))
            let mutable percent = int (System.Math.Floor (first * 100.0 + 1e-9)) |> min 100 |> max 10
            while percent > 10 && not (fits percent) do
                percent <- percent - 1
            float percent / 100.0
        else
            float ps.Scale / 100.0 |> max 0.1

    let pagesOf (area: int) (printRange: CellRange) =
        let colRanges, rowRanges = rangesAt scale printRange
        let slice rs re cs ce =
            { RowStart = rs; RowEnd = re; ColStart = cs; ColEnd = ce; Area = area
              TitleRows = titlesBefore ps.RepeatRows rs; TitleCols = titlesBefore ps.RepeatCols cs }
        // Walk order from PageSetup. DownThenOver (Excel default) emits all
        // row-pages of the first column-slice before the next column-slice;
        // OverThenDown inverts the outer/inner loops.
        let result = ResizeArray<PageSlice>()
        match ps.PageOrder with
        | DownThenOver ->
            for (cs, ce) in colRanges do
                for (rs, re) in rowRanges do
                    result.Add (slice rs re cs ce)
        | OverThenDown ->
            for (rs, re) in rowRanges do
                for (cs, ce) in colRanges do
                    result.Add (slice rs re cs ce)
        resizeArrayToList result
    let pages = printRanges |> List.mapi pagesOf |> List.concat
    // Trailing pages with no printing cell, overflow, merge or drawing are omitted. An inherited column fill/border
    // also holds a page when a visible, participating row on that page is in the same absolute 16-row block as a
    // used cell reaching that column. The style alone never widens the used range or creates a page.
    let pages =
        // (Each cell where it is first - most last pages have one; the text running over from cells, measured, only
        // when that finds nothing: a sheet of 846,241 cells and 3,018 pages.)
        let spans = ResizeArray<int * int * int * int>()
        let shown (row: int) (col: int) = rowHeightPx model row > 0.0 && colWidthPx model col > 0.0
        let printing = ResizeArray<CellAddress * CellData>()
        for KeyValue(addr, data) in model.Cells do
            if printsCell data && shown addr.Row addr.Col then
                printing.Add((addr, data))
                spans.Add((addr.Row, addr.Row, addr.Col, addr.Col))
        for merge in model.MergedCells do
            match Map.tryFind { Row = merge.StartRow; Col = merge.StartCol } model.Cells with
            | Some data when printsCell data -> spans.Add((merge.StartRow, merge.EndRow, merge.StartCol, merge.EndCol))
            | _ -> ()
        for drawing in model.Drawings do
            let anchor = drawing.Anchor
            spans.Add((anchor.FromRow, defaultArg anchor.ToRow anchor.FromRow, anchor.FromCol, defaultArg anchor.ToCol anchor.FromCol))
        let overflows = lazy (
            let runs = ResizeArray<int * int * int * int>()
            for (addr, data) in printing do
                let last = overflowEnd model addr data
                if last > addr.Col then runs.Add((addr.Row, addr.Row, addr.Col, last))
            runs)
        let inheritedBlocks =
            model.ColumnStyles
            |> List.filter (fun style -> printsCell { Value = CellValue.Empty; Format = style.Format; Comment = None })
        let blockExtents = System.Collections.Generic.Dictionary<int, int>()
        let participatingRows = System.Collections.Generic.Dictionary<int, int>()
        if not inheritedBlocks.IsEmpty then
            let participate row =
                if rowHeightPx model row > 0.0
                   && (match Map.tryFind row model.RowStyles with
                       | Some format -> printsCell { Value = CellValue.Empty; Format = format; Comment = None }
                       | None -> true) then
                    let block = row / 16
                    let bit = 1 <<< (row % 16)
                    match participatingRows.TryGetValue block with
                    | true, bits -> participatingRows.[block] <- bits ||| bit
                    | _ -> participatingRows.[block] <- bit
            for KeyValue(addr, data) in model.Cells do
                // An empty cell generated solely by its row's style is not a used cell of its own.
                if data.Value <> CellValue.Empty
                   || data.Format <> defaultFormat && Map.tryFind addr.Row model.RowStyles <> Some data.Format then
                    let block = addr.Row / 16
                    match blockExtents.TryGetValue block with
                    | true, last when last >= addr.Col -> ()
                    | _ -> blockExtents.[block] <- addr.Col
                    if colWidthPx model addr.Col > 0.0 then participate addr.Row
            for KeyValue(row, _) in model.RowPoints do participate row
            for KeyValue(row, _) in model.RowHeights do participate row
            for KeyValue(row, format) in model.RowStyles do
                if printsCell { Value = CellValue.Empty; Format = format; Comment = None } then participate row
        let holdsInherited (page: PageSlice) =
            let rowInPage block =
                let firstRow = max page.RowStart (block * 16)
                let lastRow = min page.RowEnd (block * 16 + 15)
                let pageBits = ((1 <<< (lastRow - firstRow + 1)) - 1) <<< (firstRow % 16)
                (match participatingRows.TryGetValue block with
                 | true, bits -> bits &&& pageBits <> 0
                 | _ -> false)
                || (model.MergedCells |> List.exists (fun merge ->
                    let mutable row = max firstRow merge.StartRow
                    let last = min lastRow merge.EndRow
                    let mutable shown = false
                    while not shown && row <= last do
                        shown <- rowHeightPx model row > 0.0
                                 && (match Map.tryFind row model.RowStyles with
                                     | Some format -> printsCell { Value = CellValue.Empty; Format = format; Comment = None }
                                     | None -> true)
                        row <- row + 1
                    shown && colWidthPx model merge.StartCol > 0.0))
            inheritedBlocks |> List.exists (fun style ->
                let mutable visibleCol = max page.ColStart style.First
                let lastCol = min page.ColEnd style.Last
                while visibleCol <= lastCol && colWidthPx model visibleCol <= 0.0 do
                    visibleCol <- visibleCol + 1
                if visibleCol > lastCol then false
                else
                    let mutable block = page.RowStart / 16
                    let mutable held = false
                    while not held && block <= page.RowEnd / 16 do
                        if rowInPage block then
                            held <-
                                (match blockExtents.TryGetValue block with
                                 | true, extent -> extent >= visibleCol
                                 | _ -> false)
                                || (model.MergedCells |> List.exists (fun merge ->
                                    merge.StartRow / 16 <= block && merge.EndRow / 16 >= block && merge.EndCol >= visibleCol))
                        block <- block + 1
                    held)
        let meets (page: PageSlice) (r0, r1, c0, c1) = r0 <= page.RowEnd && r1 >= page.RowStart && c0 <= page.ColEnd && c1 >= page.ColStart
        let hasSomething (page: PageSlice) =
            spans |> Seq.exists (meets page) || overflows.Value |> Seq.exists (meets page) || holdsInherited page
        let rec trim (reversed: PageSlice list) =
            match reversed with
            | last :: earlier when not earlier.IsEmpty && not (hasSomething last) -> trim earlier
            | _ -> reversed
        if pages.Length <= 1 then pages else pages |> List.rev |> trim |> List.rev

    let result =
        { Pages = pages
          PrintBounds = printRange
          PrintAreas = printRanges
          Scale = scale
          RepeatRows = ps.RepeatRows
          RepeatCols = ps.RepeatCols }
    result

/// Excel 2019 re-fits `bestFit` columns when Page Break Preview lays out a sheet. Only numbers participate;
/// if none needs more than its stored width, the sheet keeps every width. If one does, every participating
/// column gets at least one pixel beyond its stored width. The needed width is the larger of its font at
/// 100 % and 60 % zoom (the latter is often Small Fonts), including that font's digit-based padding.
/// Measured on bestfit-sweep, bestfit-figures (15 Normal fonts), and bestfit-chars in Excel 2019.
let withBestFitWidths (model: SpreadsheetModel) : SpreadsheetModel =
    if model.BestFitColumns.IsEmpty || Map.isEmpty model.Cells then model
    else
        let needed = System.Collections.Generic.Dictionary<int, float>()
        let faceOf (name: string) =
            if name.EndsWith(" MT", System.StringComparison.OrdinalIgnoreCase) then name.Substring(0, name.Length - 3) else name
        let zooms = [| 1.0; 0.6 |]
        let normalFace = faceOf model.NormalFontName
        let normalDigits =
            zooms |> Array.map (fun zoom ->
                let px = Rbt.Spreadsheet.ExcelText.pixelSize model.NormalFontSize zoom
                let small = Rbt.Spreadsheet.ExcelText.isSmallFonts model.NormalFontSize zoom
                let mutable widest = 0
                for digit in int '0' .. int '9' do
                    match Rbt.Spreadsheet.ExcelText.advance normalFace false false px small digit with
                    | ValueSome width -> widest <- max widest width
                    | ValueNone -> ()
                widest)
        let fittedWidth (format: CellFormat) (text: string) (figures: int) =
            let face = faceOf (defaultArg format.FontName model.NormalFontName)
            let points = defaultArg format.FontSize model.NormalFontSize
            let mutable widest = 0.0
            let mutable screenFit = 0.0
            for pass = 0 to zooms.Length - 1 do
                let zoom = zooms.[pass]
                let px = Rbt.Spreadsheet.ExcelText.pixelSize points zoom
                let small = Rbt.Spreadsheet.ExcelText.isSmallFonts points zoom
                let advance = Rbt.Spreadsheet.ExcelText.advance face format.Bold format.Italic px small
                let mutable mdw = 0
                for digit in int '0' .. int '9' do
                    match advance digit with
                    | ValueSome width -> mdw <- max mdw width
                    | ValueNone -> ()
                if mdw > 0 then
                    let mutable extent = 0
                    let mutable measured = true
                    if figures > 0 then extent <- figures * mdw
                    else
                        for ch in text do
                            match advance (int ch) with
                            | ValueSome width -> extent <- extent + width
                            | ValueNone -> measured <- false
                    if measured then
                        // The long-line adjustment is in the measuring font's pixels, before undoing the zoom.
                        // Its 4-digit step uses Normal's digit width at that zoom, not the cell's font:
                        // a 60%-zoom Small Fonts extent of 60/80/100 in Calibri 11 gains 1/2/3 pixels.
                        let longExtra =
                            let normal = normalDigits.[pass]
                            if normal > 0 then max 0 (extent / (4 * normal) - 2) else 0
                        let padding = 2.0 + 2.0 * System.Math.Ceiling(float mdw / 4.0)
                        let boldSmall = if small && format.Bold then 2.0 else 0.0
                        let width = 1.0 + System.Math.Floor((float (extent + longExtra) + padding + boldSmall) / zoom + 0.5)
                        if pass = 0 then screenFit <- width
                        widest <- max widest width
            struct (screenFit, widest)
        let mutable grow = false
        for KeyValue(addr, data) in model.Cells do
            if Set.contains addr.Col model.BestFitColumns && getColumnWidth addr.Col model > 0.0 then
                match unwrapFormulaValue data.Value with
                | Number number ->
                    let format = data.Format
                    let code = defaultArg format.NumberFormat "General"
                    let general = code.Equals("General", System.StringComparison.OrdinalIgnoreCase) || System.String.IsNullOrWhiteSpace code
                    let shown = Rbt.Spreadsheet.CellText.formatNumberWithCodeIn model.Date1904 code number
                    if not (System.String.IsNullOrEmpty shown) then
                        let figures, text =
                            if general then
                                let mutable count = 0
                                let mutable i = 0
                                while i < shown.Length && shown.[i] <> '.' && shown.[i] <> 'E' && shown.[i] <> 'e' do
                                    if System.Char.IsDigit shown.[i] then count <- count + 1
                                    i <- i + 1
                                max 1 count, ""
                            else
                                let lead, trail = Rbt.Spreadsheet.CellText.widthPads format data.Value
                                if lead.Length + trail.Length <= shown.Length then
                                    0, lead + shown.Substring(lead.Length, shown.Length - lead.Length - trail.Length) + trail
                                else 0, shown
                        let struct (screenFit, width) = fittedWidth format text figures
                        if width > 0.0 then
                            // Excel triggers the sheet-wide refit only when the 100%-zoom measurement
                            // *before* its one-pixel column bias exceeds the stored width. At Calibri 12,
                            // General 2010 triggers at 37 px, not at 38, although its 60% fit is 51 px.
                            if screenFit - 1.0 > getColumnWidth addr.Col model then grow <- true
                            match needed.TryGetValue addr.Col with
                            | true, previous when previous >= width -> ()
                            | _ -> needed.[addr.Col] <- width
                | _ -> ()
        if not grow then model
        else
            let mutable widths = model.ColumnWidths
            for KeyValue(col, width) in needed do
                let stored = getColumnWidth col model
                widths <- Map.add col (max (stored + 1.0) width) widths
            { model with ColumnWidths = widths }

/// The width a column prints at, in pixels: Excel 2019's PDF sets a column of `width` pixels on the screen
/// (`FontMetrics.columnWidthPx`, whole pixels of the Normal font's maximum digit width `mdw`) at `printDigit` / `mdw`
/// dots of 600 dpi a pixel, `printDigit` the widest digit of the Normal font in print (`ExcelText.printDigitWidth`),
/// rounded to whole dots: 64 px of Calibri 11 (7 / 47) print at 430 dots (68.8 px), 64 px of Calibri 12 (8 / 51) at
/// 408 (65.28 px), 12 px of Arial 10 (7 / 46) at 79 (make-print-pages-fixture.py "widths": 144 columns of 0.1 to 25
/// characters in Calibri 11 and 12, Arial 10 and 12, every one within a dot of Excel's PDF).
let printColumnWidthPx (mdw: float) (printDigit: float) (width: float) : float =
    if width <= 0.0 then width
    else
        let mdw = if mdw > 0.0 then mdw else 7.0
        let digit = if printDigit > 0.0 then printDigit else mdw * 47.0 / 7.0
        System.Math.Floor (width * digit / mdw + 0.5) / 6.25

/// The height a row of `px` pixels on the screen prints at, in pixels: Excel 2019's PDF sets it at `scale` dots of
/// 600 dpi a pixel (`ExcelText.printRowScale`: 6.05 for a Normal font of Calibri 11), cut down to whole dots - a row
/// of 20 px (15 pt) prints at 121 dots (14.52 pt), one of 546 px (409.5 pt) at 3303. (A height of a fraction of a
/// pixel more than a whole one prints as that one: 0.14 of a dot is let in.) make-print-pages-fixture.py "heights":
/// rows of 7.5 to 409.5 pt in Calibri 11 and 12, Arial 10 and 12, Cambria 14 within a dot of Excel's PDF where the
/// height is whole pixels.
let printRowHeightPx (scale: float) (px: float) : float =
    if px <= 0.0 then px
    else System.Math.Floor (px * (if scale > 0.0 then scale else 6.05) + 0.14) / 6.25

/// A row's height in pixels unrounded: its points x 4/3 while they are the `px` it shows at (`FontMetrics.pointsToRowPx`),
/// else `px` (heights 7.5 to 409.5 pt in Excel 2019's PDF: 20 pt prints as 26.67 px would, not as the 26 shown).
let exactRowPx (points: float) (px: float) : float =
    if points > 0.0 && px > 0.0 && Rbt.Spreadsheet.FontMetrics.pointsToRowPx points = px then points * 4.0 / 3.0 else px

/// The sheet with its rows and columns at the sizes they print at.
let withPrintSizes (model: SpreadsheetModel) : SpreadsheetModel =
    let model = withBestFitWidths model
    { model with
        DefaultRowHeight = printRowHeightPx model.PrintRowScale (exactRowPx model.DefaultRowPoints model.DefaultRowHeight)
        DefaultColumnWidth = printColumnWidthPx model.MaxDigitWidth model.PrintDigitWidth model.DefaultColumnWidth
        ColumnWidths = (if Map.isEmpty model.ColumnWidths then model.ColumnWidths else model.ColumnWidths |> Map.map (fun _ width -> printColumnWidthPx model.MaxDigitWidth model.PrintDigitWidth width))
        RowHeights = (if Map.isEmpty model.RowHeights then model.RowHeights else model.RowHeights |> Map.map (fun row px -> printRowHeightPx model.PrintRowScale (exactRowPx (defaultArg (Map.tryFind row model.RowPoints) 0.0) px))) }

/// The pages the sheet prints on: `paginate` of the sheet at its print
/// sizes - what the PDF export lays out, and what a host's print preview is told.
let paginatePrinted (model: SpreadsheetModel) : Pagination =
    paginate (withPrintSizes model)
