module RedBat.Spreadsheet.Paint

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.CellText

type Color = string

type TextAlign =
    | AlignLeft
    | AlignCenter
    | AlignRight

type VerticalAlign =
    | VAlignTop
    | VAlignMiddle
    | VAlignBottom

type PaintText = {
    Text: string
    /// The full cell (or header) rectangle. Per-line offset is derived from
    /// LineIndex/LineCount and font metrics in the renderer — Paint.fs does
    /// not pack baseline math into the rect.
    Rect: Rect
    Color: Color
    /// Font size in points, as written in the source XLSX. Each renderer
    /// converts to its native unit (px for raster, pt for PDF, css-pt for DOM).
    FontSize: float
    FontName: string option
    Bold: bool
    Italic: bool
    /// Render a horizontal stroke under the text (XLSX `<u/>` font flag).
    Underline: bool
    /// Render two parallel strokes for `<u val="double"/>` instead of one.
    UnderlineDouble: bool
    /// Render a horizontal stroke through the text (XLSX `<strike/>` flag).
    Strikethrough: bool
    Align: TextAlign
    /// Vertical alignment within Rect. Excel "General" = VAlignBottom.
    VerticalAlign: VerticalAlign
    /// 0-based line index within a wrapped block; 0 if not wrapped.
    LineIndex: int
    /// Total lines in the wrapped block; 1 if not wrapped.
    LineCount: int
    /// Signed pixel offset the renderer adds to the computed baseline.
    /// Used for superscript (negative — shift up) and subscript (positive
    /// — shift down). Zero for normal text.
    BaselineShift: float
    /// Cell-text rotation in degrees. 0 = horizontal (default), 1..90 =
    /// CCW (text rises left-to-right), 91..180 = CW (encoded as
    /// 90 + |angle|), 255 = stacked (each glyph on its own line). The
    /// renderer rotates the glyph stream around the cell's anchor; the
    /// rect itself is unrotated.
    Rotation: int
}

type PaintItem =
    | FillRect of Rect * Color
    | StrokeRect of Rect * Color * float
    | DrawText of PaintText
    | GridLine of float * float * float * float * Color * float

type ViewportPaintItem = {
    Zone: ViewportZone
    Item: PaintItem
}

/// Excel's General horizontal alignment puts numbers/booleans on the right,
/// strings on the left, and errors in the center. Explicit alignment on
/// CellFormat overrides this default.
let rec private generalAlign (value: CellValue): TextAlign =
    match value with
    | Number _ -> AlignRight
    | Boolean _ -> AlignCenter
    | Error _ -> AlignCenter
    | Formula (_, cached) -> generalAlign cached
    | CellValue.Empty -> AlignLeft
    | _ -> AlignLeft

let private alignFromFormat (format: CellFormat) (value: CellValue): TextAlign =
    match format.HorizontalAlign with
    | Some "center" -> AlignCenter
    // "Center Across Selection" — the text is centred (across a run of
    // centerContinuous cells; the run rect is computed at paint time).
    | Some "centerContinuous" -> AlignCenter
    | Some "right" -> AlignRight
    | Some "left" -> AlignLeft
    | _ ->
        // A `@` (Text) number format makes Excel treat the value as text, so
        // even a number is General-aligned left rather than right.
        match format.NumberFormat with
        | Some nf when nf.Trim() = "@" -> AlignLeft
        | _ -> generalAlign value

let private isSelected (selection: Selection) row col =
    match selection with
    | SingleCell (r, c) -> r = row && c = col
    | Range (sr, sc, er, ec) ->
        row >= sr && row <= er && col >= sc && col <= ec

let private buildHeaderPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield FillRect (rect, "#e8e8e8")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
            | ColumnHeader (col, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = columnToLetter col
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    FontName = None
                    Bold = false
                    Italic = false
                    Underline = false
                    UnderlineDouble = false
                    Strikethrough = false
                    Align = AlignCenter
                    VerticalAlign = VAlignMiddle
                    LineIndex = 0
                    LineCount = 1
                    BaselineShift = 0.0
                    Rotation = 0
                }
            | RowHeader (row, rect) ->
                yield FillRect (rect, "#f2f2f2")
                yield StrokeRect (rect, "#a6a6a6", 1.0)
                yield DrawText {
                    Text = string (row + 1)
                    Rect = rect
                    Color = "#444444"
                    FontSize = 11.0
                    FontName = None
                    Bold = false
                    Italic = false
                    Underline = false
                    UnderlineDouble = false
                    Strikethrough = false
                    Align = AlignCenter
                    VerticalAlign = VAlignMiddle
                    LineIndex = 0
                    LineCount = 1
                    BaselineShift = 0.0
                    Rotation = 0
                }
    ]

/// Translate an OOXML border style name into an approximate stroke width in
/// CSS pixels. Excel has many variants (hair, thin, medium, thick, double,
/// dashed/dotted in three weights, slantDashDot, etc.); we collapse them into
/// thin/medium/thick stroke widths and a default black color.
let private borderStrokeWidth (style: string) : float =
    match style.ToLowerInvariant() with
    | "hair" -> 0.5
    | "thin" | "dashed" | "dotted" | "dashdot" | "dashdotdot" -> 1.0
    | "medium" | "mediumdashed" | "mediumdotted" | "mediumdashdot" | "mediumdashdotdot" -> 1.75
    | "thick" -> 3.0
    | "double" -> 2.0
    | "slantdashdot" -> 1.5
    | _ -> 1.0

/// Border-collapse priority for `<border style>` values. NOT the same as
/// stroke width — Excel's collapse rule follows the OOXML hierarchy where
/// `double` outranks `thick` even though `thick` paints fatter, because
/// `double` carries more visual semantics. Order: None < Hair < Dotted <
/// Dashed-family < Thin < Medium < Thick < Double. Equal priority resolves
/// in the caller's tiebreak.
let private borderPriority (style: string) : int =
    match style.ToLowerInvariant() with
    | "" -> 0
    | "hair" -> 1
    | "dotted" -> 2
    | "dashed" -> 3
    | "dashdot" -> 4
    | "dashdotdot" -> 5
    | "mediumdashed" -> 6
    | "mediumdashdot" -> 7
    | "mediumdashdotdot" -> 8
    | "slantdashdot" -> 9
    | "thin" -> 10
    | "medium" -> 11
    | "thick" -> 12
    | "double" -> 13
    | _ -> 10  // unknown → treat as thin

let private borderColor (side: BorderStyle) : string =
    side.Color |> Option.defaultValue "#000000"

/// Multi-pass cell rendering. Drawing fill+stroke per cell in a single loop
/// lets the next cell's fill overwrite the previous cell's right/bottom edge
/// — invisible against white but obvious when a cell has a non-default fill
/// (Table styles, fills baked into cellXf). Splitting into three passes
/// (fills → strokes/borders → text) keeps all borders on top of all fills.
///
/// Items are tagged with the ViewportZone of the cell they came from so the
/// interactive renderer can scroll the body without sliding frozen
/// rows/cols. The static (PNG/PDF) renderer ignores the tag — at
/// ScrollX/Y=0 every zone paints at its absolute position anyway, so the
/// only visible effect of frozen panes there is the divider lines emitted
/// below.
let private buildCellPaintItems (model: SpreadsheetModel) (viewport: Viewport): (ViewportZone * PaintItem) list =
    let cells = visibleCells model viewport
    let cellState (cell: _) =
        let addr = cell.Address
        let cellData = getCell addr model
        let value = getCellValue addr model
        let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
        let selected = isSelected model.Selection addr.Row addr.Col
        let background =
            if selected then "#E8F5E9"
            else format.BgColor |> Option.defaultValue "#ffffff"
        let zone = zoneForCell model addr.Row addr.Col
        cell, format, value, selected, background, zone

    let states = cells |> List.map cellState

    [
        // Pass 1: all backgrounds. Adjacent cells share an edge, so when a
        // cell with a non-default fill is followed by another cell, the
        // second cell's fill would overdraw the first cell's right/bottom
        // stroke if we drew them in the same per-cell loop.
        for (cell, _format, _value, _selected, background, zone) in states do
            yield zone, FillRect (cell.Rect, background)

        // Pass 1b: conditional-formatting data bars. Drawn on top of the
        // cell background but under borders / text — a horizontal bar
        // from the cell's left edge spanning fraction × width. Excel
        // insets the bar ~2 px vertically; mirror that so the gridline
        // stays visible above and below.
        for (cell, format, _value, _selected, _background, zone) in states do
            match format.DataBar with
            | Some (frac, color) when frac > 0.0 ->
                let inset = 2.0
                let barW = (cell.Rect.Width - 1.0) * frac
                if barW > 0.5 then
                    yield zone, FillRect (
                        { X = cell.Rect.X
                          Y = cell.Rect.Y + inset
                          Width = barW
                          Height = max 1.0 (cell.Rect.Height - inset * 2.0) },
                        color)
            | _ -> ()

        // Pass 2: per-cell strokes + cellXf borders, drawn on top of every
        // fill. Match Excel's *interactive* view: a light gray gridline
        // around every cell — not the near-black grid that print mode emits
        // with PrintGridlines on. Our reference PNGs come from print, but the
        // renderer's target is the on-screen view.
        // Emit a styled border side as one (single solid) or two (double)
        // GridLine items. "double" in Excel is two ~0.75 px parallel rules
        // straddling the cell edge with a small gap; emit two 1 px strokes
        // offset by ±1 px from the edge so the gap is visible at 96 dpi.
        // Other styles flow through the existing single-stroke path with
        // the borderStrokeWidth weight.
        let emitBorderSide (isHorizontal: bool) (side: BorderStyle) (x1: float, y1: float) (x2: float, y2: float) =
            let color = borderColor side
            let w = borderStrokeWidth side.Style
            let style = side.Style.ToLowerInvariant()
            // Excel's dashed / dotted / dash-dot variants render as a
            // pattern of on/off chunks. We approximate by chunking the
            // line into short solid segments separated by gaps — the
            // PaintItem.GridLine plumbing carries only width, not dash
            // info, and emitting one segment per dash keeps every
            // downstream renderer (PNG / PDF / browser) consistent.
            let dashChunks onLen offLen =
                let segs = System.Collections.Generic.List<PaintItem>()
                if isHorizontal then
                    let mutable x = x1
                    while x < x2 do
                        let segEnd = min x2 (x + onLen)
                        segs.Add(GridLine (x, y1, segEnd, y2, color, w))
                        x <- segEnd + offLen
                else
                    let mutable y = y1
                    while y < y2 do
                        let segEnd = min y2 (y + onLen)
                        segs.Add(GridLine (x1, y, x2, segEnd, color, w))
                        y <- segEnd + offLen
                List.ofSeq segs
            if style = "double" then
                // Excel renders double borders straddling the cell edge
                // (one stroke just outside, one just inside). When the
                // adjacent cell has a wider border (e.g. `thick`, 3 px),
                // both straddling strokes get overpainted by the
                // neighbour's stroke. The caller compensates by passing
                // a `doubleInset` adjusted pair of y / x values — see
                // the per-cell emission loop for the direction-aware
                // offsets.
                let off = 1.0
                if isHorizontal then
                    [ GridLine (x1, y1 - off, x2, y2 - off, color, 1.0)
                      GridLine (x1, y1 + off, x2, y2 + off, color, 1.0) ]
                else
                    [ GridLine (x1 - off, y1, x2 - off, y2, color, 1.0)
                      GridLine (x1 + off, y1, x2 + off, y2, color, 1.0) ]
            elif style = "dotted" || style = "mediumdotted" then
                dashChunks 1.0 1.5
            elif style = "dashed" || style = "mediumdashed" then
                dashChunks 3.0 2.5
            elif style = "dashdot" || style = "mediumdashdot" then
                // 3-on, 1.5-off, 1-on, 1.5-off pattern: emit two sub-passes
                let mainDash = dashChunks 3.0 4.0  // 3 + 1 + ~ offset for dot
                let dotPass =
                    let segs = System.Collections.Generic.List<PaintItem>()
                    if isHorizontal then
                        let mutable x = x1 + 4.0
                        while x < x2 do
                            let segEnd = min x2 (x + 1.0)
                            segs.Add(GridLine (x, y1, segEnd, y2, color, w))
                            x <- x + 8.0
                    else
                        let mutable y = y1 + 4.0
                        while y < y2 do
                            let segEnd = min y2 (y + 1.0)
                            segs.Add(GridLine (x1, y, x2, segEnd, color, w))
                            y <- y + 8.0
                    List.ofSeq segs
                mainDash @ dotPass
            elif style = "dashdotdot" || style = "mediumdashdotdot" then
                dashChunks 3.0 5.0  // simplification; main dash with bigger gap
            elif style = "slantdashdot" then
                dashChunks 3.0 3.0
            elif style = "hair" then
                // Hair = very thin dotted. Simulate with tiny dots.
                dashChunks 0.5 1.0
            else
                [ GridLine (x1, y1, x2, y2, color, w) ]

        // Border collapse: when two adjacent cells both declare a border
        // on their shared edge, Excel draws the heavier one and skips the
        // lighter. Same-weight ties resolve in favour of the cell that
        // "owns" the edge first in the natural emission order — for
        // vertical edges that's the upper cell's bottom; for horizontal
        // edges that's the left cell's right. Without this collapse,
        // double + adjacent thin produced a doubled-line artifact and
        // dashed borders got overpainted by the neighbour's solid edge.
        let borderByAddr : Map<int * int, CellFormat> =
            states
            |> List.map (fun (c, f, _, _, _, _) -> (c.Address.Row, c.Address.Col), f)
            |> Map.ofList
        let neighborSide (row: int) (col: int) (side: CellFormat -> BorderStyle option) : BorderStyle option =
            Map.tryFind (row, col) borderByAddr |> Option.bind side
        // Use OOXML border priority (not raw stroke width) for collapse:
        // `double` outranks `thick` in Excel even though it paints thinner.
        // Order: None < Hair < Dotted < Dashed-family < Thin < Medium <
        // Thick < Double.
        let weightOf (side: BorderStyle option) : int =
            match side with
            | None -> 0
            | Some s -> borderPriority s.Style
        // Gridlines (the always-on light cell outline) FIRST, as their own
        // sub-pass, so explicit borders always draw on top. Excel hides the
        // light gridline on any edge that carries an explicit border, so we
        // draw the grey edge only when neither this cell nor its neighbour has
        // a border there. Without this the grey gridline bleeds through a
        // double border's gap (and gets half-covered by neighbouring grid
        // strokes). Selected cells keep a full highlight box.
        for (cell, format, _value, selected, _background, zone) in states do
            let rect = cell.Rect
            let r, c = cell.Address.Row, cell.Address.Col
            if selected then
                yield zone, StrokeRect (rect, "#217346", 2.0)
            // Suppress the light grey screen gridlines when the sheet has
            // `showGridLines="0"` (polished workbooks rely on explicit borders).
            elif model.ShowGridLines then
                let bordered (mine: BorderStyle option) nr nc nside =
                    mine.IsSome || (neighborSide nr nc nside).IsSome
                let g = "#a6a6a6"
                if not (bordered format.BorderTop (r - 1) c (fun f -> f.BorderBottom)) then
                    yield zone, GridLine (rect.X, rect.Y, rect.X + rect.Width, rect.Y, g, 1.0)
                if not (bordered format.BorderBottom (r + 1) c (fun f -> f.BorderTop)) then
                    yield zone, GridLine (rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0)
                if not (bordered format.BorderLeft r (c - 1) (fun f -> f.BorderRight)) then
                    yield zone, GridLine (rect.X, rect.Y, rect.X, rect.Y + rect.Height, g, 1.0)
                if not (bordered format.BorderRight r (c + 1) (fun f -> f.BorderLeft)) then
                    yield zone, GridLine (rect.X + rect.Width, rect.Y, rect.X + rect.Width, rect.Y + rect.Height, g, 1.0)
        for (cell, format, _value, _selected, _background, zone) in states do
            let rect = cell.Rect
            let r, c = cell.Address.Row, cell.Address.Col
            // Merged cells: Excel splits a box border across the merge's
            // constituent cells (top-left carries left+top, top-right the
            // right edge, bottom-left/-right the bottom edge, etc.).
            // visibleCells yields only the merge primary — with the full
            // merge Rect — and skips the trailing cells, so we gather each
            // edge's style from the merge perimeter directly off the model
            // rather than from the primary's own format. For a non-merged
            // cell every edge list is the single cell, so this reduces to
            // the previous `format.BorderX` behavior exactly.
            let mr0, mc0, mr1, mc1 =
                match cell.Merge with
                | Some m -> m.StartRow, m.StartCol, m.EndRow, m.EndCol
                | None -> r, c, r, c
            let fmtAt row col =
                getCell { Row = row; Col = col } model
                |> Option.map (fun cd -> cd.Format)
                |> Option.defaultValue defaultFormat
            // Heaviest border style found along an edge of the merge.
            let edgeSide (edge: (int * int) list) (pick: CellFormat -> BorderStyle option) =
                edge
                |> List.choose (fun (rr, cc) -> pick (fmtAt rr cc))
                |> List.sortByDescending (fun s -> borderPriority s.Style)
                |> List.tryHead
            let bTop = edgeSide [ for cc in mc0 .. mc1 -> mr0, cc ] (fun f -> f.BorderTop)
            let bBottom = edgeSide [ for cc in mc0 .. mc1 -> mr1, cc ] (fun f -> f.BorderBottom)
            let bLeft = edgeSide [ for rr in mr0 .. mr1 -> rr, mc0 ] (fun f -> f.BorderLeft)
            let bRight = edgeSide [ for rr in mr0 .. mr1 -> rr, mc1 ] (fun f -> f.BorderRight)
            // Top border: skip if the cell above already drew a heavier
            // (or equal-weight) bottom edge. The above-bottom owns ties.
            // Higher priority wins (double > thick > medium > thin > dashed).
            match bTop with
            | Some side ->
                let aboveBottom = neighborSide (mr0 - 1) mc0 (fun f -> f.BorderBottom)
                let mine = weightOf (Some side)
                let neighbor = weightOf aboveBottom
                if mine > neighbor then
                    for item in emitBorderSide true side (rect.X, rect.Y) (rect.X + rect.Width, rect.Y) do
                        yield zone, item
            | None -> ()
            // Bottom border: draws unless the cell below has a strictly
            // heavier top border.
            match bBottom with
            | Some side ->
                let belowTop = neighborSide (mr1 + 1) mc0 (fun f -> f.BorderTop)
                let mine = weightOf (Some side)
                let neighbor = weightOf belowTop
                if mine >= neighbor then
                    let y = rect.Y + rect.Height
                    for item in emitBorderSide true side (rect.X, y) (rect.X + rect.Width, y) do
                        yield zone, item
            | None -> ()
            // Left border: skip if the cell to the left already drew a
            // heavier-or-equal right edge.
            match bLeft with
            | Some side ->
                let leftRight = neighborSide mr0 (mc0 - 1) (fun f -> f.BorderRight)
                let mine = weightOf (Some side)
                let neighbor = weightOf leftRight
                if mine > neighbor then
                    for item in emitBorderSide false side (rect.X, rect.Y) (rect.X, rect.Y + rect.Height) do
                        yield zone, item
            | None -> ()
            // Right border: draws unless the cell to the right has a
            // strictly heavier left border.
            match bRight with
            | Some side ->
                let rightLeft = neighborSide mr0 (mc1 + 1) (fun f -> f.BorderLeft)
                let mine = weightOf (Some side)
                let neighbor = weightOf rightLeft
                if mine >= neighbor then
                    let x = rect.X + rect.Width
                    for item in emitBorderSide false side (x, rect.Y) (x, rect.Y + rect.Height) do
                        yield zone, item
            | None -> ()
            // Diagonal strokes: down = top-left → bottom-right, up =
            // bottom-left → top-right. Cells set one or both via the
            // `diagonalUp` / `diagonalDown` border attributes. Style +
            // color come from the shared `<diagonal>` child, so both
            // strokes share the same look when both flags are set.
            match format.BorderDiagonal with
            | Some side ->
                let color = borderColor side
                let w = borderStrokeWidth side.Style
                if format.DiagonalDown then
                    yield zone, GridLine (rect.X, rect.Y, rect.X + rect.Width, rect.Y + rect.Height, color, w)
                if format.DiagonalUp then
                    yield zone, GridLine (rect.X, rect.Y + rect.Height, rect.X + rect.Width, rect.Y, color, w)
            | None -> ()

        // Pass 3: text on top of all strokes/borders.
        for (cell, format, value, _selected, _background, zone) in states do
            let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
            // Build the real-font measurer first so cellDisplayText can size
            // its ##### overflow run against the glyph the renderer will
            // actually draw, not the mdwPx approximation.
            let measure =
                viewport.Measurer
                |> Option.map (fun m ->
                    fun (s: string) -> m s fontPt format.FontName format.Bold format.Italic)
            let displayText = cellDisplayText measure cell.Rect.Width format value
            // Cells in a hidden column (0 width) or hidden row (0 height) paint
            // nothing — Excel renders no text there, and without this guard the
            // hidden cell's text overflows into the next visible column (and a
            // 0-width cell wraps its text to one glyph per line).
            if displayText <> "" && cell.Rect.Width > 0.0 && cell.Rect.Height > 0.0 then
                // Indent past the conditional-formatting icon when present
                // so the icon and the cell text don't overlap (Excel does
                // the same — the icon occupies a ~16 px gutter).
                let iconGutter = if format.IconSet.IsSome then 16.0 else 0.0
                // Excel's `<alignment indent="N"/>` adds N units of leading
                // padding (≈ 3 character widths per unit). It applies from
                // the LEFT for left/distributed alignment, from the RIGHT
                // for right align; ignored for center alignment.
                let indentPx =
                    if format.Indent <= 0 then 0.0
                    else
                        let unit = mdwPx fontPt * 3.0
                        float format.Indent * unit
                let horizForIndent =
                    format.HorizontalAlign |> Option.map (fun s -> s.ToLowerInvariant()) |> Option.defaultValue ""
                let leftIndent =
                    if indentPx > 0.0 && (horizForIndent = "left" || horizForIndent = "distributed") then indentPx
                    else 0.0
                let rightIndent =
                    if indentPx > 0.0 && horizForIndent = "right" then indentPx
                    else 0.0
                // "Center Across Selection" (`horizontal="centerContinuous"`):
                // centre the text across the contiguous run of centerContinuous
                // cells to the right that carry no content of their own — Excel
                // does this without merging. Expand the base rect over the run.
                let baseRect =
                    if horizForIndent = "centercontinuous" then
                        let row = cell.Address.Row
                        let mutable lastCol = cell.Address.Col
                        let mutable c2 = cell.Address.Col + 1
                        let mutable go = true
                        while go && c2 < model.ColCount do
                            let naddr = { Row = row; Col = c2 }
                            let nfmt = getCell naddr model |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
                            let nHoriz = nfmt.HorizontalAlign |> Option.map (fun s -> s.ToLowerInvariant()) |> Option.defaultValue ""
                            let nEmpty = match getCellValue naddr model with CellValue.Empty -> true | _ -> false
                            if nHoriz = "centercontinuous" && nEmpty then lastCol <- c2; c2 <- c2 + 1
                            else go <- false
                        if lastCol > cell.Address.Col then
                            let r = cellRect model viewport row lastCol
                            { cell.Rect with Width = (r.X + r.Width) - cell.Rect.X }
                        else cell.Rect
                    else cell.Rect
                let textRect =
                    { baseRect with
                        X = baseRect.X + 4.0 + iconGutter + leftIndent
                        Width = max 0.0 (baseRect.Width - 8.0 - iconGutter - leftIndent - rightIndent) }
                // ShrinkToFit (single-line text only — wrap-mode escapes via
                // line breaks). When the text overflows the cell width and
                // shrinkToFit is set, scale the font size down until it fits,
                // clamped at 1 pt. We need to re-measure with the smaller
                // size so the rest of the rendering uses the actual width.
                let fontPt, measure =
                    if not format.ShrinkToFit || format.WrapText then fontPt, measure
                    else
                        match measure with
                        | None -> fontPt, measure
                        | Some m ->
                            let width = m displayText
                            if width <= textRect.Width || textRect.Width <= 0.0 then fontPt, measure
                            else
                                let ratio = max 0.1 (textRect.Width / width)
                                let shrunk = max 1.0 (fontPt * ratio)
                                let shrunkMeasure =
                                    viewport.Measurer
                                    |> Option.map (fun mm ->
                                        fun (s: string) -> mm s shrunk format.FontName format.Bold format.Italic)
                                shrunk, shrunkMeasure
                // Excel number-format section colour overrides the cell's
                // font colour: `#,##0.00;[Red]-#,##0.00` paints negatives red
                // even when the cellXf has no explicit fontColor.
                let sectionColor =
                    match format.NumberFormat, value with
                    | Some code, Number n -> numberFormatColorForValue code n
                    | Some code, Formula (_, Number n) -> numberFormatColorForValue code n
                    | _ -> None
                let color =
                    sectionColor
                    |> Option.orElse format.TextColor
                    |> Option.defaultValue "#000000"
                let align = alignFromFormat format value
                // Wrapping and manual `\n` line breaks only apply when wrapText
                // is ON. With wrapText OFF, Excel renders on a single line and
                // ignores any embedded break (shown as a space), so the row is
                // not multi-line.
                let lines =
                    if format.WrapText then
                        wrapLines measure textRect.Width fontPt displayText
                    else [ displayText.Replace("\n", "") ]
                let vAlign =
                    match format.VerticalAlign |> Option.map (fun s -> s.ToLowerInvariant()) with
                    | Some "top" -> VAlignTop
                    | Some "center" | Some "middle" -> VAlignMiddle
                    | _ -> VAlignBottom
                let lineCount = lines.Length
                let richRuns = Map.tryFind cell.Address model.RichRuns
                match richRuns, viewport.Measurer with
                | Some runs, Some m when not (List.isEmpty runs) ->
                    // Rich-text rendering. For single-line cells we emit one
                    // DrawText per run. For wrapped/multi-line cells we walk
                    // a position cursor through the concatenated source text
                    // and chunk runs that straddle line boundaries — each
                    // line gets per-run DrawTexts at its own y, with the
                    // breaking whitespace dropped by wrapLines absorbed into
                    // the cursor (so a run that ends with " " before a line
                    // break still aligns with the next line's start).
                    let resolveRun (ovr: RunOverride) =
                        let bold = ovr.Bold |> Option.defaultValue format.Bold
                        let italic = ovr.Italic |> Option.defaultValue format.Italic
                        let underline = ovr.Underline |> Option.defaultValue format.Underline
                        let underlineDouble = ovr.UnderlineDouble |> Option.defaultValue format.UnderlineDouble
                        let strike = ovr.Strikethrough |> Option.defaultValue format.Strikethrough
                        let fName = ovr.FontName |> Option.orElse format.FontName
                        let baseSize = ovr.FontSize |> Option.defaultValue fontPt
                        let runColor = ovr.Color |> Option.defaultValue color
                        // vertAlign: shrink to ~60% of the cell font (Excel
                        // uses ~58% empirically, close enough) and shift
                        // the baseline. Negative shift = up (super); we
                        // express shift in pt at the cell's base size so
                        // the renderer's pt→px scaling stays consistent.
                        let fSize, baselineShift =
                            match ovr.VertAlign with
                            | Some VertAlignSuperscript -> baseSize * 0.6, -fontPt * 0.35
                            | Some VertAlignSubscript -> baseSize * 0.6, fontPt * 0.15
                            | _ -> baseSize, 0.0
                        bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, baselineShift
                    // Mutable list of remaining runs. We consume from the
                    // head; if a wrap break splits a run, the head is
                    // replaced with the unconsumed suffix.
                    let mutable remaining: (string * RunOverride) list = runs
                    let displayConcat = runs |> List.map fst |> String.concat ""
                    let mutable docPos = 0
                    for i, line in List.indexed lines do
                        let slices = System.Collections.Generic.List<string * RunOverride>()
                        let mutable lineRem = line.Length
                        while lineRem > 0 && not (List.isEmpty remaining) do
                            let (runText, runProps) = List.head remaining
                            let take = min lineRem runText.Length
                            slices.Add(runText.Substring(0, take), runProps)
                            lineRem <- lineRem - take
                            if take = runText.Length then
                                remaining <- List.tail remaining
                            else
                                // Suffix of the current run stays at head.
                                remaining <- (runText.Substring(take), runProps) :: List.tail remaining
                        docPos <- docPos + line.Length
                        // Skip the breaking space wrapLines drops between
                        // non-final lines. Also consume the matching space
                        // from the run queue head so positions stay aligned.
                        if i < lines.Length - 1
                           && docPos < displayConcat.Length
                           && displayConcat.[docPos] = ' ' then
                            docPos <- docPos + 1
                            match remaining with
                            | (rt, rp) :: rest when rt.StartsWith " " ->
                                if rt.Length > 1 then
                                    remaining <- (rt.Substring(1), rp) :: rest
                                else
                                    remaining <- rest
                            | _ -> ()
                        // Measure this line's segments and emit.
                        let metrics =
                            slices
                            |> Seq.map (fun (t, ovr) ->
                                let bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, baselineShift = resolveRun ovr
                                let width = m t fSize fName bold italic
                                t, ovr, bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, width, baselineShift)
                            |> Seq.toList
                        let totalWidth = metrics |> List.sumBy (fun (_, _, _, _, _, _, _, _, _, _, w, _) -> w)
                        let baseLineX =
                            match align with
                            | AlignLeft -> textRect.X
                            | AlignCenter -> textRect.X + (textRect.Width - totalWidth) / 2.0
                            | AlignRight -> textRect.X + textRect.Width - totalWidth
                        let mutable cursor = baseLineX
                        for (t, _ovr, bold, italic, underline, underlineDouble, strike, fName, fSize, runColor, width, baselineShift) in metrics do
                            yield zone, DrawText {
                                Text = t
                                Rect = { textRect with X = cursor; Width = width }
                                Color = runColor
                                FontSize = fSize
                                FontName = fName
                                Bold = bold
                                Italic = italic
                                Underline = underline
                                UnderlineDouble = underlineDouble
                                Strikethrough = strike
                                Align = AlignLeft
                                VerticalAlign = vAlign
                                LineIndex = i
                                LineCount = lineCount
                                BaselineShift = baselineShift
                                Rotation = format.TextRotation
                            }
                            cursor <- cursor + width
                | _ ->
                    // Plain (or wrapped) cell text — one DrawText per line.
                    for i, line in List.indexed lines do
                        yield zone, DrawText {
                            Text = line
                            Rect = textRect
                            Color = color
                            FontSize = fontPt
                            FontName = format.FontName
                            Bold = format.Bold
                            Italic = format.Italic
                            Underline = format.Underline
                            UnderlineDouble = format.UnderlineDouble
                            Strikethrough = format.Strikethrough
                            Align = align
                            VerticalAlign = vAlign
                            LineIndex = i
                            LineCount = lineCount
                            BaselineShift = 0.0
                            Rotation = format.TextRotation
                        }

        // Pass 4: auto-filter dropdown triangles on table header cells. Excel
        // renders these as a small downward triangle on the right edge of each
        // header cell whose column falls within the table's <autoFilter>
        // range. The triangle is drawn over the header's accent fill — same
        // color as the header text (white on a Medium-style accent).
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.FilterButton then
                let color = format.TextColor |> Option.defaultValue "#FFFFFF"
                // 7-px-wide, 4-px-tall triangle centered in a 16-px right-edge
                // button area. Drawn as five horizontal stroked lines (no
                // primitive for filled polygons; this is cheap enough).
                let buttonWidth = 16.0
                let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
                let triHalfWidth = 3.5
                let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
                for i in 0 .. 3 do
                    let shrink = float i * (triHalfWidth / 4.0)
                    let y = triTopY + float i
                    yield zone, GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0)

        // Pass 5: comment-indicator red triangle in the top-right corner.
        // Excel uses a 5-px right-triangle filled with #FF0000 — we
        // approximate with horizontal strokes of decreasing width since
        // the renderer has no fill-polygon primitive. The triangle's
        // hypotenuse goes from (right-5, top) down to (right, top+5).
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.HasComment then
                let color = "#C00000"  // Excel's "note marker" red
                let triSize = 5.0
                let rightX = cell.Rect.X + cell.Rect.Width
                let topY = cell.Rect.Y
                // Row k from top: from x = rightX - (triSize - k) to x = rightX.
                for k in 0 .. int triSize - 1 do
                    let leftX = rightX - (triSize - float k)
                    yield zone, GridLine (leftX, topY + float k + 0.5, rightX, topY + float k + 0.5, color, 1.0)

        // Pass 6: data-validation list dropdown arrow. Excel only shows
        // this when the cell is selected, but for static export we
        // always paint it so the visual cue carries through to PNG /
        // PDF snapshots. Drawn as a small 7-px-wide triangle just
        // inside the cell's right edge, vertically centered, in gray.
        // FilterButton already owns the dropdown styling for table
        // header cells; skip those to avoid doubled triangles.
        for (cell, format, _value, _selected, _background, zone) in states do
            if format.HasListDropdown && not format.FilterButton then
                let color = "#666666"
                let buttonWidth = 16.0
                let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
                let triHalfWidth = 3.5
                let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
                for i in 0 .. 3 do
                    let shrink = float i * (triHalfWidth / 4.0)
                    let y = triTopY + float i
                    yield zone, GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0)

        // Pass 7: conditional-formatting icon-set indicator. The renderer
        // has no icon artwork or circle primitive, so approximate with a
        // small filled square in a low→high color (red → amber → green
        // for a 3-set; interpolated for 4/5-sets). It conveys the same
        // low/mid/high signal the real icons do. Drawn in the 16 px
        // gutter the text was already indented past.
        for (cell, format, _value, _selected, _background, zone) in states do
            match format.IconSet with
            | Some (idx, count) when count > 1 ->
                // 0 = worst (red), count-1 = best (green), midpoints amber.
                let t = float idx / float (count - 1)
                let color =
                    if t <= 0.5 then
                        // red → amber over the first half
                        let u = t * 2.0
                        sprintf "#FF%02X00" (int (0.0 * (1.0 - u) + 192.0 * u))
                    else
                        // amber → green over the second half
                        let u = (t - 0.5) * 2.0
                        sprintf "#%02X%02X%02X"
                            (int (255.0 * (1.0 - u)))
                            (int (192.0 * (1.0 - u) + 176.0 * u))
                            (int (80.0 * u))
                let sz = 9.0
                let iconX = cell.Rect.X + 3.0
                let iconY = cell.Rect.Y + cell.Rect.Height / 2.0 - sz / 2.0
                yield zone, FillRect ({ X = iconX; Y = iconY; Width = sz; Height = sz }, color)
            | _ -> ()
    ]

/// Freeze-divider strokes drawn on top of the cell layer. Excel renders
/// this as a single ~1 px line slightly darker than the regular gridline,
/// at the split's right (cols) / bottom (rows) edge. Matches the line
/// Excel draws in both interactive and print-snapshot views. Coordinates
/// are in the same space as cellRect (post-HeaderHeight / -RowHeaderWidth
/// offset) so the line lands flush with the bottom/right of the last
/// frozen row/column.
let private buildFreezeDividerItems (model: SpreadsheetModel) (viewport: Viewport) : PaintItem list =
    // Slightly darker than the regular gridline. Excel uses a fixed
    // medium-gray (~#595959 in print mode).
    let color = "#595959"
    let w = 1.0
    [
        if model.FrozenRows > 0 then
            let y =
                viewport.HeaderHeight
                + ([ 0 .. model.FrozenRows - 1 ] |> List.sumBy (fun r -> getRowHeight r model))
            let totalWidth =
                viewport.RowHeaderWidth
                + ([ 0 .. model.ColCount - 1 ] |> List.sumBy (fun c -> getColumnWidth c model))
            yield GridLine (viewport.RowHeaderWidth, y, totalWidth, y, color, w)
        if model.FrozenCols > 0 then
            let x =
                viewport.RowHeaderWidth
                + ([ 0 .. model.FrozenCols - 1 ] |> List.sumBy (fun c -> getColumnWidth c model))
            let totalHeight =
                viewport.HeaderHeight
                + ([ 0 .. model.RowCount - 1 ] |> List.sumBy (fun r -> getRowHeight r model))
            yield GridLine (x, viewport.HeaderHeight, x, totalHeight, color, w)
    ]

let buildPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    [
        yield FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff")
        yield! buildHeaderPaintItems model viewport
        // Static (PNG/PDF) path: drop the zone tag; everything paints at
        // absolute cell coordinates because ScrollX/Y are zero.
        for (_zone, item) in buildCellPaintItems model viewport do
            yield item
        // Freeze divider lines on top of cell strokes but under any
        // selection / DrawText emitted later. The cell-item zone tag is
        // irrelevant here; we draw the divider in the model coordinate
        // system that's already baked into the cell rects above.
        yield! buildFreezeDividerItems model viewport
    ]

let private translatePaintItem viewport zone item =
    let tx = toViewportRect viewport zone
    match item with
    | FillRect (rect, color) -> FillRect (tx rect, color)
    | StrokeRect (rect, color, width) -> StrokeRect (tx rect, color, width)
    | DrawText text -> DrawText { text with Rect = tx text.Rect }
    | GridLine (x1, y1, x2, y2, color, width) ->
        let p1 = tx { X = x1; Y = y1; Width = 0.0; Height = 0.0 }
        let p2 = tx { X = x2; Y = y2; Width = 0.0; Height = 0.0 }
        GridLine (p1.X, p1.Y, p2.X, p2.Y, color, width)

let buildViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport): ViewportPaintItem list =
    [
        yield { Zone = CornerZone
                Item = FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff") }

        for header in headers model viewport do
            match header with
            | CornerHeader rect ->
                yield { Zone = CornerZone; Item = FillRect (rect, "#e8e8e8") }
                yield { Zone = CornerZone; Item = StrokeRect (rect, "#a6a6a6", 1.0) }
            | ColumnHeader (col, rect) ->
                let zone = ColumnHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = columnToLetter col
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            FontName = None
                            Bold = false
                            Italic = false
                            Underline = false
                            UnderlineDouble = false
                            Strikethrough = false
                            Align = AlignCenter
                            VerticalAlign = VAlignMiddle
                            LineIndex = 0
                            LineCount = 1
                            BaselineShift = 0.0
                            Rotation = 0
                        })
                }
            | RowHeader (row, rect) ->
                let zone = RowHeaderZone
                yield { Zone = zone; Item = translatePaintItem viewport zone (FillRect (rect, "#f2f2f2")) }
                yield { Zone = zone; Item = translatePaintItem viewport zone (StrokeRect (rect, "#a6a6a6", 1.0)) }
                yield {
                    Zone = zone
                    Item =
                        translatePaintItem viewport zone (DrawText {
                            Text = string (row + 1)
                            Rect = rect
                            Color = "#444444"
                            FontSize = 11.0
                            FontName = None
                            Bold = false
                            Italic = false
                            Underline = false
                            UnderlineDouble = false
                            Strikethrough = false
                            Align = AlignCenter
                            VerticalAlign = VAlignMiddle
                            LineIndex = 0
                            LineCount = 1
                            BaselineShift = 0.0
                            Rotation = 0
                        })
                }

        for (zone, item) in buildCellPaintItems model viewport do
            yield { Zone = zone; Item = translatePaintItem viewport zone item }

        // Freeze dividers paint with the same zone as the body they border,
        // so they scroll with the body horizontally (row divider) or
        // vertically (col divider) and disappear under the frozen band when
        // the body scrolls past them. Excel does the same: the divider
        // moves with the scrolling region, not the frozen one.
        for item in buildFreezeDividerItems model viewport do
            let zone =
                match item with
                | GridLine (x1, _, x2, _, _, _) when x1 = x2 ->
                    // Vertical line (cols divider) — scrolls vertically with
                    // body but its X is pinned (no horizontal scroll).
                    FrozenRowsZone
                | _ ->
                    // Horizontal line (rows divider) — scrolls horizontally
                    // with body, pinned vertically.
                    FrozenColsZone
            yield { Zone = zone; Item = translatePaintItem viewport zone item }
    ]
