module Rbt.Spreadsheet.Paint

open System.Collections.Generic
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Layout
open Rbt.Spreadsheet.CellText

type Color = string

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
    /// Where the line's baseline is, in pixels below the top of `Rect`, as Excel 2019 prints it (`Paint`: cell text);
    /// negative - placed by the font's metrics (its Windows ascent and descent at its size) and `VerticalAlign`, and
    /// -`Baseline` lower, from the grid's lines; 0 - the renderer places it by the font's metrics and `VerticalAlign`.
    Baseline: float
    /// Cell-text rotation in degrees. 0 = horizontal (default), 1..90 =
    /// CCW (text rises left-to-right), 91..180 = CW (encoded as
    /// 90 + |angle|), 255 = stacked (each glyph on its own line). The
    /// renderer rotates the glyph stream around the cell's anchor; the
    /// rect itself is unrotated.
    Rotation: int
    /// What the text is clipped to, when it is wider than its cell: the cell
    /// itself, or the run of empty cells it overflows into (Excel's text
    /// overflow). None: the text fits, or is not clipped (rotated text).
    Clip: Rect option
}

/// A single segment of a vector path. Coordinates are in the same px space
/// as everything else in PaintItem. Mirrors the renderer / canvas / PDF path
/// primitives so a chart scene can be expressed once and drawn by all three.
type VecSeg =
    | VMove of float * float
    | VLine of float * float
    | VCubic of float * float * float * float * float * float
    | VClose

type PaintItem =
    | FillRect of Rect * Color
    | StrokeRect of Rect * Color * float
    | DrawText of PaintText
    | GridLine of float * float * float * float * Color * float
    /// Arbitrary vector path with an optional fill and an optional
    /// (color, width) stroke — the primitive charts (and future drawings)
    /// emit. At least one of fill/stroke is expected.
    | VecPath of VecSeg list * fill: Color option * stroke: (Color * float) option
    /// A nested scene transformed by a 2×3 affine (a, b, c, d, e, f), i.e.
    /// x' = a·x + c·y + e, y' = b·x + d·y + f. Used to place/scale a chart
    /// scene into its anchor rect.
    | VecGroup of (float * float * float * float * float * float) * PaintItem list
    /// Items cut off at a rectangle: a chart across a page break, on each of its pages.
    | Clipped of Rect * PaintItem list
    /// A picture stretched over a rectangle. A renderer that cannot decode
    /// images draws the frame `drawingFrameItems` describes instead.
    | DrawImage of Rect * PaintImage

/// The data of a picture to paint.
and PaintImage = {
    /// "image/png", "image/jpeg", ...
    ContentType: string
    Data: byte[]
    /// What to say when the picture cannot be shown.
    Label: string
}

type ViewportPaintItem = {
    Zone: ViewportZone
    Item: PaintItem
}

/// A hand-built vector scene used to prove the VecPath/VecGroup primitives
/// render identically through the native, canvas, and PDF paths (the
/// foundation a chart adapter would emit). Exercises a filled triangle, a
/// stroked polyline, a filled+stroked box, and a `“data bar”` strip — all
/// inside a translated group plus standalone axes.
let vectorTestScene : PaintItem list =
    [ // Plot frame (axes) as a stroked path, no group.
      VecPath ([ VMove (40.0, 20.0); VLine (40.0, 200.0); VLine (320.0, 200.0) ], None, Some ("#808080", 1.0))
      // Everything else placed by a group translated to the plot origin.
      VecGroup ((1.0, 0.0, 0.0, 1.0, 40.0, 200.0),
        [ // Filled area (semi-triangle) under a curve — fill only.
          VecPath ([ VMove (0.0, 0.0); VLine (0.0, -60.0); VCubic (40.0, -120.0, 90.0, -40.0, 130.0, -90.0); VLine (130.0, 0.0); VClose ],
                   Some "#BDD7EE", None)
          // Line on top of the area — stroke only.
          VecPath ([ VMove (0.0, -60.0); VCubic (40.0, -120.0, 90.0, -40.0, 130.0, -90.0) ], None, Some ("#1F497D", 2.0))
          // Three column bars — filled + thin stroke.
          VecPath ([ VMove (160.0, 0.0); VLine (190.0, 0.0); VLine (190.0, -70.0); VLine (160.0, -70.0); VClose ], Some "#70AD47", Some ("#000000", 0.75))
          VecPath ([ VMove (200.0, 0.0); VLine (230.0, 0.0); VLine (230.0, -110.0); VLine (200.0, -110.0); VClose ], Some "#70AD47", Some ("#000000", 0.75))
          VecPath ([ VMove (240.0, 0.0); VLine (270.0, 0.0); VLine (270.0, -45.0); VLine (240.0, -45.0); VClose ], Some "#70AD47", Some ("#000000", 0.75)) ]) ]

let private isSelected (selection: Selection) row col =
    match selection with
    | SingleCell (r, c) -> r = row && c = col
    | Range (sr, sc, er, ec) ->
        row >= sr && row <= er && col >= sc && col <= ec

// ---- the look of a cell ------------------------------------------------------------

/// How a cell looks: what every way of drawing it draws it with - the canvas, PDF and the native
/// hosts (`buildCellPaintItems`) and the web grid's cell nodes (`Spreadsheet.drawCell`, which maps
/// it to CSS). Worked out once, here, so that they cannot disagree about a cell; how each draws it
/// (its fonts, its paddings) is its own.
type CellLook = {
    /// The cell's format with the look of its table's style under it and its conditional formats
    /// over it (and `FilterButton` on a header cell of a table's auto-filter).
    Format: CellFormat
    /// What the style of its table gives it (`Layout.tableEntryAt`).
    Table: TableLookEntry option
    /// The value it shows: the sheet's view options applied (formulas as text, no zeros).
    Value: CellValue
    /// Its text: the value under its number format (`CellText.cellDisplayTextIn`: General numbers
    /// to the digits that fit, #### for a number of a format code that does not fit); "" where a
    /// data bar or an icon set shows no values. A line break of text that is not wrapped is not
    /// shown (Excel draws such text on one line).
    Text: string
    /// Its fill: its own, else its table's; a conditional format's over both.
    Fill: Color option
    /// The colour of its text: a conditional format's, else the colour of its number format's
    /// section (`[Red]`), else its own, else its table's. None: the default, black.
    TextColor: Color option
    /// In points.
    FontSize: float
    FontName: string option
    Bold: bool
    Italic: bool
    Underline: bool
    UnderlineDouble: bool
    Strikethrough: bool
    /// Where the text sits across the cell (`Layout.cellTextAlign`).
    Align: TextAlign
    /// The format's horizontal alignment, lower-cased: "" for General.
    Horizontal: string
    VerticalAlign: VerticalAlign
    Wrap: bool
    /// What the indent adds at the left / at the right of the text, in pixels: at the side the text
    /// is aligned to (none for centred text).
    IndentLeft: float
    IndentRight: float
    /// `CellFormat.TextRotation` (0: none).
    Rotation: int
    /// One line of text, drawn smaller where it is wider than its cell (shrink to fit; not for
    /// wrapped or turned text).
    Shrink: bool
    /// Centred across the selection: the last column of the run of empty cells it is centred over;
    /// -1 when it is not.
    AcrossTo: int
    /// A conditional format's data bar.
    DataBar: DataBarDraw option
    /// A conditional format's icon (of a set of more than one).
    Icon: IconDraw option
}

let private lowerAlign (align: string option) : string =
    match align with
    | None -> ""
    | Some ("left" | "center" | "right" | "top" | "bottom" | "general" | "fill" | "justify" | "distributed" as known) -> known
    | Some other -> other.ToLowerInvariant()

/// What one level of indent is, in pixels, where the workbook's is not known: 9 px, three spaces of Calibri 11
/// (`SpreadsheetModel.IndentUnit`).
let indentUnit (_fontPt: float) : float = 9.0

/// The look of a cell with nothing in it, no table and no conditional format.
let private emptyLook : CellLook =
    { Format = defaultFormat; Table = None; Value = CellValue.Empty; Text = ""; Fill = None; TextColor = None
      FontSize = 11.0; FontName = None; Bold = false; Italic = false; Underline = false; UnderlineDouble = false; Strikethrough = false
      Align = AlignLeft; Horizontal = ""; VerticalAlign = VAlignBottom; Wrap = false; IndentLeft = 0.0; IndentRight = 0.0
      Rotation = 0; Shrink = false; AcrossTo = -1; DataBar = None; Icon = None }

/// The look of the cell (row, col) of a sheet (`CellLook`), which holds `data` (`getCell`: what is
/// stored). `measure` measures its text where the painter knows its fonts (None: by Excel's digit
/// width, `mdwPx`); `width` is the cell's width in pixels (a merged cell's whole width), what a
/// number has to fit.
/// (`fit`: the text measured as Excel draws it, in screen pixels at a zoom, `width` too: a number
/// fitted to its cell as Excel fits it, `CellText.NumberFit`.)
let cellLookFitted (fit: NumberFit option) (model: SpreadsheetModel) (row: int) (col: int) (data: CellData option) (measure: TextMeasurer option) (width: float) : CellLook =
    if data.IsNone && model.Tables.IsEmpty && model.Pivots.IsEmpty && model.ConditionalFormats.IsEmpty then emptyLook
    else
        let addr = { Row = row; Col = col }
        let own, stored =
            match data with
            | Some data -> data.Format, data.Value
            | None -> defaultFormat, CellValue.Empty
        let value = viewedValue model stored
        // Under it the look of its table's style; a header cell of the table's filter has its button.
        let table = if model.Tables.IsEmpty && model.Pivots.IsEmpty then None else styleEntryAt model row col
        let withTable =
            match table with
            | None -> own
            | Some entry ->
                let styled = formatWithTableEntry entry own
                let filterButton =
                    model.Tables |> List.exists (fun table ->
                        row >= table.StartRow - 1 && row < table.StartRow - 1 + table.HeaderRowCount
                        && col >= table.StartCol - 1 && col <= table.EndCol - 1
                        && (match table.AutoFilter with
                            | Some (firstCol, lastCol) -> col + 1 >= firstCol && col + 1 <= lastCol
                            | None -> false))
                if filterButton && not styled.FilterButton then { styled with FilterButton = true } else styled
        // Over it its conditional formats.
        let effect =
            if model.ConditionalFormats.IsEmpty then None
            else Rbt.Spreadsheet.ConditionalFormatting.effectAt (Rbt.Spreadsheet.ConditionalFormatting.contextFor model) addr
        let format =
            match effect with
            | Some effect -> Rbt.Spreadsheet.ConditionalFormatting.applyEffect effect withTable
            | None -> withTable
        let fontPt = defaultArg format.FontSize 11.0
        let horizontal = lowerAlign format.HorizontalAlign
        let rotation = format.TextRotation
        let wrap = format.WrapText
        let icon = match format.IconSet with Some icon when icon.Count > 1 -> Some icon | _ -> None
        // "Center Across Selection": centred over the run of center-across cells to its right that
        // hold nothing, as Excel does without merging them.
        let acrossTo =
            if horizontal <> "centercontinuous" then -1
            else
                let mutable last = col
                let mutable next = col + 1
                let mutable go = true
                while go && next < model.ColCount do
                    match getCell { Row = row; Col = next } model with
                    | Some cell when lowerAlign cell.Format.HorizontalAlign = "centercontinuous"
                                     && (match cell.Value with CellValue.Empty -> true | _ -> false) ->
                        last <- next
                        next <- next + 1
                    | _ -> go <- false
                last
        // (A number centred across the selection fits in the width of the run, as Excel shows it.)
        let fitWidth =
            if acrossTo <= col then width
            else
                width + (match fit with Some fit -> fit.Zoom | None -> 1.0)
                        * (seq { col + 1 .. acrossTo } |> Seq.sumBy (fun c -> model.ColumnWidths |> Map.tryFind c |> Option.defaultValue model.DefaultColumnWidth))
        // The text. (A number shrunk to fit or turned is not fitted to the cell's width: shrinking
        // is what makes it fit.)
        let text =
            match value with
            | CellValue.Empty -> ""
            | CellValue.String text -> text
            | _ ->
                let measureText =
                    match measure with
                    | Some m -> Some (fun (text: string) -> m text fontPt format.FontName format.Bold format.Italic)
                    | None -> None
                // (Italic, and bold in Small Fonts, have less room: `ExcelText.overhang`, `NumberFit.Overhang`.)
                let fit = fit |> Option.map (fun fit -> { fit with Overhang = ExcelText.overhang (ExcelText.pixelSize fontPt fit.Zoom) (ExcelText.isSmallFonts fontPt fit.Zoom) format.Bold format.Italic })
                cellDisplayTextFitted fit model.Date1904 measureText (if format.ShrinkToFit || rotation <> 0 then 1.0e6 else fitWidth) format value
        let text =
            match format.DataBar, format.IconSet with
            | Some bar, _ when bar.HideValue -> ""
            | _, Some icon when icon.HideValue -> ""
            | _ -> if not wrap && text.IndexOf '\n' >= 0 then text.Replace("\n", "") else text
        let textColor =
            match effect with
            | Some { TextColor = Some color } -> Some color
            | _ ->
                let section =
                    match format.NumberFormat with
                    | Some code ->
                        match unwrapFormulaValue value with
                        | CellValue.Number n -> numberFormatColorForValue code n
                        | _ -> None
                    | None -> None
                match section with
                | Some _ -> section
                | None -> format.TextColor
        let align = cellTextAlign format value
        let indentLeft, indentRight =
            if format.Indent <= 0 then 0.0, 0.0
            else
                let indent = float format.Indent * model.IndentUnit
                match horizontal with
                | "left" | "distributed" -> indent, 0.0
                | "right" -> 0.0, indent
                | "" | "general" ->
                    match align with
                    | AlignLeft -> indent, 0.0
                    | AlignRight -> 0.0, indent
                    | AlignCenter -> 0.0, 0.0
                | _ -> 0.0, 0.0
        let verticalAlign =
            match lowerAlign format.VerticalAlign with
            | "top" -> VAlignTop
            | "center" | "middle" -> VAlignMiddle
            | _ -> VAlignBottom
        { Format = format; Table = table; Value = value; Text = text; Fill = format.BgColor; TextColor = textColor
          FontSize = fontPt; FontName = format.FontName; Bold = format.Bold; Italic = format.Italic
          Underline = format.Underline; UnderlineDouble = format.UnderlineDouble; Strikethrough = format.Strikethrough
          Align = align; Horizontal = horizontal; VerticalAlign = verticalAlign; Wrap = wrap
          IndentLeft = indentLeft; IndentRight = indentRight; Rotation = rotation
          Shrink = format.ShrinkToFit && not wrap && rotation = 0; AcrossTo = acrossTo
          DataBar = format.DataBar; Icon = icon }

/// `cellLookFitted` with the text's width estimated (`CellText.cellDisplayTextIn`).
let cellLook (model: SpreadsheetModel) (row: int) (col: int) (data: CellData option) (measure: TextMeasurer option) (width: float) : CellLook =
    cellLookFitted None model row col data measure width

let private isWhiteFill color =
    color = "#ffffff" || color = "#FFFFFF" || color = "white"

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
                    Baseline = 0.0
                    Rotation = 0
                    Clip = None
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
                    Baseline = 0.0
                    Rotation = 0
                    Clip = None
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
    | "medium" | "mediumdashed" | "mediumdotted" | "mediumdashdot" | "mediumdashdotdot" -> 2.0
    | "thick" -> 3.0
    | "double" -> 2.0
    | "slantdashdot" -> 1.5
    | _ -> 1.0

/// The colour of the painter's grid lines.
let private gridLineColor = "#a6a6a6"

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
/// How far a line of text stands from its cell's sides in print (the PDF this paints) under a Normal font of Calibri
/// 11: 16/600 in, 2.56 px - checked in Excel 2019's PDF of generated cells (make-print-text-fixture.py: left and right
/// aligned text in 26 columns of Calibri 11 and bold, Arial 10, Times New Roman 12, 1.93 pt either side). On screen
/// it is 2 px left and 3 right. Other Normal fonts: `printPadXOf`.
[<Literal>]
let printPadX = 2.56

/// How far a line of text stands from its cell's sides in print: ceil(d / 4) + 4 dots of 600 dpi, d the Normal font's
/// widest digit in print (`SpreadsheetModel.PrintDigitWidth`; whatever the cell's own font) - 16 dots for Calibri 11
/// (d 47) and Arial 10 (46), 14 for Calibri 9 / Arial 8, 17 for Calibri 12, 18 for Arial 12, 19 for Calibri 14, 21
/// for Cambria 14 (the rowunit, heights and widths fixtures of make-print-pages-fixture.py: 33 Normal fonts, the
/// text's origin in Excel 2019's PDF to the dot).
let printPadXOf (model: SpreadsheetModel) =
    let digit = if model.PrintDigitWidth > 0.0 then model.PrintDigitWidth else 47.0
    (System.Math.Ceiling (digit / 4.0) + 4.0) * 96.0 / 600.0

/// What a wrapped line's room is less than its printed column: 28/600 in, 4.48 px (the same fixture's paragraphs,
/// broken in the font's advances at its size: 102 of 108 columns as Excel's PDF; on screen 5 px, measured in GDI's).
[<Literal>]
let printWrapMargin = 4.48

/// The largest descent (`FontMetrics.lineOf`, screen pixels) of the text at the bottom of its cells, by row: the cells
/// of a row so aligned meet at one baseline (Excel 2019's PDF). Worked out once for a sheet's cells (the last asked
/// for: a PDF paints them page by page).
let private descentsCache : (obj * Dictionary<int, float>) option ref = ref None
let private rowDescentsOf (model: SpreadsheetModel) : int -> float =
    let descents =
        match descentsCache.Value with
        | Some (cells, known) when obj.ReferenceEquals(cells, model.Cells) -> known
        | _ ->
            let known = Dictionary<int, float>()
            // (A cell merged down over other rows holds its text at its block's top left.)
            let multiRowStarts = HashSet<CellAddress>()
            for merge in model.MergedCells do
                if merge.EndRow > merge.StartRow then multiRowStarts.Add { Row = merge.StartRow; Col = merge.StartCol } |> ignore
            let lines = Dictionary<string, float>()
            let descentOf (face: string option) (points: float) (bold: bool) (italic: bool) =
                let key = (defaultArg face "") + "|" + string points + "|" + string bold + string italic
                match lines.TryGetValue key with
                | true, d -> d
                | _ ->
                    let struct (_, d) = FontMetrics.lineOf face points bold italic
                    lines.[key] <- d
                    d
            for KeyValue(addr, data) in model.Cells do
                let format = data.Format
                if (match format.VerticalAlign with None | Some "bottom" -> true | Some _ -> false) && format.TextRotation = 0
                   && (match unwrapFormulaValue data.Value with CellValue.Empty | CellValue.String "" -> false | _ -> true)
                   && not (multiRowStarts.Count > 0 && multiRowStarts.Contains addr) then
                    let points = defaultArg format.FontSize 11.0
                    let mutable most = descentOf format.FontName points format.Bold format.Italic
                    match Map.tryFind addr model.RichRuns with
                    | Some runs ->
                        for _, run in runs do
                            most <- max most (descentOf (run.FontName |> Option.orElse format.FontName) (defaultArg run.FontSize points)
                                                (defaultArg run.Bold format.Bold) (defaultArg run.Italic format.Italic))
                    | None -> ()
                    match known.TryGetValue addr.Row with
                    | true, d when d >= most -> ()
                    | _ -> known.[addr.Row] <- most
            descentsCache.Value <- Some (box model.Cells, known)
            known
    fun row -> match descents.TryGetValue row with | true, d -> d | _ -> 0.0

/// Place the same 16-pixel paths the browser's icon SVG uses, without
/// reparsing SVG (or rasterizing it at PDF resolution).
let iconPaintItems (family: string) (index: int) (rect: Rect) (align: VerticalAlign) : PaintItem list =
    let top =
        match align with
        | VAlignTop -> 3.0
        | VAlignMiddle -> System.Math.Floor((rect.Height - 16.0) / 2.0) + 1.0
        | VAlignBottom -> rect.Height - 17.0
    let x, y = rect.X + 2.0, rect.Y + top
    let toSeg = function
        | Rbt.Spreadsheet.ConditionalFormatting.IconMove (a,b) -> VMove (a,b)
        | Rbt.Spreadsheet.ConditionalFormatting.IconLine (a,b) -> VLine (a,b)
        | Rbt.Spreadsheet.ConditionalFormatting.IconCurve (a,b,c,d,e,f) -> VCubic (a,b,c,d,e,f)
        | Rbt.Spreadsheet.ConditionalFormatting.IconClose -> VClose
    let parts =
        [for part in Rbt.Spreadsheet.ConditionalFormatting.iconArtwork family index do
            let mark = VecPath (part.Segments |> List.map toSeg, part.Fill, part.Stroke)
            yield if part.ClipLeft then Clipped ({ X = 0.0; Y = 0.0; Width = 8.0; Height = 16.0 }, [mark]) else mark]
    [VecGroup ((1.0,0.0,0.0,1.0,x,y),parts)]

/// Convert a sparkline's shared cell-relative marks into renderer-neutral
/// paths. Geometry is calculated once regardless of the target renderer.
let sparklinePaintItems (group: SparklineGroup) (values: float option array) (scale: float * float) (rect: Rect) : PaintItem list =
    let dot x y r color =
        let k = 0.5522847498307936 * r
        VecPath ([VMove(x+r,y); VCubic(x+r,y+k,x+k,y+r,x,y+r); VCubic(x-k,y+r,x-r,y+k,x-r,y)
                  VCubic(x-r,y-k,x-k,y-r,x,y-r); VCubic(x+k,y-r,x+r,y-k,x+r,y); VClose],Some color,None)
    let parts =
        [for mark in Rbt.Spreadsheet.Sparklines.marks group values scale rect.Width rect.Height do
            match mark with
            | Rbt.Spreadsheet.Sparklines.SparkPath (points,color,width) ->
                match points with
                | (x,y)::tail -> yield VecPath (VMove(x,y)::[for a,b in tail -> VLine(a,b)],None,Some(color,width))
                | [] -> ()
            | Rbt.Spreadsheet.Sparklines.SparkBar (x,y,w,h,color) ->
                yield FillRect ({ X=x; Y=y; Width=w; Height=h },color)
            | Rbt.Spreadsheet.Sparklines.SparkDot (x,y,r,color) -> yield dot x y r color]
    if List.isEmpty parts then [] else [VecGroup ((1.0,0.0,0.0,1.0,rect.X,rect.Y),parts)]

/// Sparse sparkline destinations are not necessarily cells in model.Cells.
/// Iterate their location references rather than the populated-cell list.
let sparklineCellItems (source: Rbt.Spreadsheet.Sparklines.Source) (model: SpreadsheetModel) (viewport: Viewport) : (ViewportZone * PaintItem) list =
    if model.Sparklines.IsEmpty then []
    else
        let visible = visibleRange model viewport
        [for group in model.Sparklines do
            for addr, _ in group.Items do
                if addr.Row >= visible.StartRow && addr.Row < visible.EndRowExclusive
                   && addr.Col >= visible.StartCol && addr.Col < visible.EndColExclusive then
                    match Rbt.Spreadsheet.Sparklines.inCell source [group] addr with
                    | Some (_,values,scale) ->
                        let rect = cellRect model viewport addr.Row addr.Col
                        for item in sparklinePaintItems group values scale rect do
                            yield zoneForCell model addr.Row addr.Col,item
                    | None -> ()]

let private buildCellPaintItems (model: SpreadsheetModel) (viewport: Viewport): (ViewportZone * PaintItem) list =
    let visible = visibleRange model viewport
    let padX = printPadXOf model
    let rowDescent = rowDescentsOf model
    let cells = visibleCells model viewport
    // Selection is rendered as a separate overlay (see `selectionPaintItems`)
    // so cursor moves don't invalidate the cached grid paint items.
    let cellState (cell: CellGeometry) =
        let addr = cell.Address
        // Its look (`cellLook`): its format with its table's look under it and its conditional formats over it.
        let look = cellLook model addr.Row addr.Col (SheetOps.shownCell model addr) viewport.Measurer cell.Rect.Width
        cell, look, zoneForCell model addr.Row addr.Col

    let states = cells |> List.map cellState

    let items = [
        // Pass 1: all backgrounds. Adjacent cells share an edge, so when a
        // cell with a non-default fill is followed by another cell, the
        // second cell's fill would overdraw the first cell's right/bottom
        // stroke if we drew them in the same per-cell loop.
        for (cell, look, zone) in states do
            match look.Fill with
            | Some fill when not (isWhiteFill fill) -> yield zone, FillRect (cell.Rect, fill)
            | _ -> ()

        // Pass 1b: conditional-formatting data bars. Drawn on top of the
        // cell background but under borders / text — a horizontal bar
        // from the cell's left edge spanning fraction × width. Excel
        // insets the bar ~2 px vertically; mirror that so the gridline
        // stays visible above and below.
        for (cell, look, zone) in states do
            let format = look.Format
            match format.DataBar with
            | Some bar ->
                let inset = 2.0
                let width = cell.Rect.Width - 1.0
                let top = cell.Rect.Y + inset
                let height = max 1.0 (cell.Rect.Height - inset * 2.0)
                let left = cell.Rect.X + width * bar.From
                let barW = width * (bar.To - bar.From)
                if barW > 0.5 then
                    if bar.Gradient then
                        // The fill fades to white along the bar: strips of
                        // the colour, lighter one after the other.
                        let strips = 8
                        for strip in 0 .. strips - 1 do
                            let lightness = (float strip / float strips) * 0.85
                            let place = if bar.FadeRight then strip else strips - 1 - strip
                            yield zone, FillRect (
                                { X = left + barW * float place / float strips
                                  Y = top
                                  Width = barW / float strips + 0.5
                                  Height = height },
                                Rbt.Spreadsheet.ConditionalFormatting.blendColors bar.Color "#FFFFFF" lightness)
                    else
                        yield zone, FillRect ({ X = left; Y = top; Width = barW; Height = height }, bar.Color)
                    match bar.Border with
                    | Some line -> yield zone, StrokeRect ({ X = left + 0.5; Y = top + 0.5; Width = max 1.0 (barW - 1.0); Height = max 1.0 (height - 1.0) }, line, 1.0)
                    | None -> ()
                match bar.Axis with
                | Some (at, color) ->
                    // A dashed line down the cell.
                    let x = cell.Rect.X + width * at
                    let mutable y = cell.Rect.Y
                    while y < cell.Rect.Y + cell.Rect.Height do
                        yield zone, FillRect ({ X = x; Y = y; Width = 1.0; Height = min 2.0 (cell.Rect.Y + cell.Rect.Height - y) }, color)
                        y <- y + 3.0
                | None -> ()
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
                // A solid line runs half its width past both ends, so that two
                // of them meeting at a corner leave no notch there (an outline
                // is continuous through its corners).
                let cap = if w > 1.0 then w / 2.0 else 0.0
                if isHorizontal then [ GridLine (x1 - cap, y1, x2 + cap, y2, color, w) ]
                else [ GridLine (x1, y1 - cap, x2, y2 + cap, color, w) ]

        // The edges: one line per edge two cells share, the one the web grid
        // draws too (Layout.edgeAbove / edgeLeft: the stronger border of the
        // two cells, each cell's own line else its table's; else a grid line,
        // which the sheet's option or a fill of either cell hides). Each cell
        // draws the edge above it and the edge to its left, along a merged
        // cell edge by edge of its cells; the cells of the last row / column
        // in view also the edges below / to their right. Grid lines first,
        // so that borders are drawn over them.
        let lastRow = visible.EndRowExclusive - 1
        let lastCol = visible.EndColExclusive - 1
        let gridItems = ResizeArray<ViewportZone * PaintItem>()
        let borderItems = ResizeArray<ViewportZone * PaintItem>()
        let edgeItems (zone: ViewportZone) (isHorizontal: bool) (x: float, y: float) (runs: (float * float * SharedEdge) list) =
            for offset, length, edge in runs do
                if length > 0.0 then
                    let x1, y1 = if isHorizontal then x + offset, y else x, y + offset
                    let x2, y2 = if isHorizontal then x1 + length, y1 else x1, y1 + length
                    match edge with
                    | EdgeGrid true -> gridItems.Add (zone, GridLine (x1, y1, x2, y2, gridLineColor, 1.0))
                    | EdgeBorder side -> for item in emitBorderSide isHorizontal side (x1, y1) (x2, y2) do borderItems.Add (zone, item)
                    | _ -> ()
        for (cell, _look, zone) in states do
            let rect = cell.Rect
            if rect.Width > 0.0 && rect.Height > 0.0 then
                let r0, c0, r1, c1 =
                    match cell.Merge with
                    | Some m -> m.StartRow, m.StartCol, m.EndRow, m.EndCol
                    | None -> cell.Address.Row, cell.Address.Col, cell.Address.Row, cell.Address.Col
                let columnStart (c: int) = columnLeft model c
                let rowStart (r: int) = rowTop model r
                edgeItems zone true (rect.X, rect.Y) (edgeRuns c0 c1 columnStart (fun c -> edgeAbove model r0 c))
                edgeItems zone false (rect.X, rect.Y) (edgeRuns r0 r1 rowStart (fun r -> edgeLeft model r c0))
                if r1 >= lastRow then
                    edgeItems zone true (rect.X, rect.Y + rect.Height) (edgeRuns c0 c1 columnStart (fun c -> edgeBelow model r1 c))
                if c1 >= lastCol then
                    edgeItems zone false (rect.X + rect.Width, rect.Y) (edgeRuns r0 r1 rowStart (fun r -> edgeRight model r c1))
        yield! gridItems
        yield! borderItems
        for (cell, look, zone) in states do
            let rect = cell.Rect
            let format = look.Format
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
        for (cell, look, zone) in states do
            let format = look.Format
            let value = look.Value
            let fontPt = look.FontSize
            let measure =
                viewport.Measurer
                |> Option.map (fun m ->
                    fun (s: string) -> m s fontPt format.FontName format.Bold format.Italic)
            // (Its text, number-formatted to fit, "" where a data bar or an icon set hides it: `cellLook`.)
            let displayText = look.Text
            // Cells in a hidden column (0 width) or hidden row (0 height) paint
            // nothing — Excel renders no text there, and without this guard the
            // hidden cell's text overflows into the next visible column (and a
            // 0-width cell wraps its text to one glyph per line).
            if displayText <> "" && cell.Rect.Width > 0.0 && cell.Rect.Height > 0.0 then
                // Indent past the conditional-formatting icon when present
                // so the icon and the cell text don't overlap (Excel does
                // the same — the icon occupies a ~16 px gutter).
                let iconGutter = if look.Icon.IsSome then 16.0 else 0.0
                // Excel's `<alignment indent="N"/>` adds N units of padding at
                // the side the text is aligned to (`cellLook`).
                let leftIndent = look.IndentLeft
                let rightIndent = look.IndentRight
                let horizForIndent = look.Horizontal
                // "Center Across Selection" (`horizontal="centerContinuous"`):
                // centred across the run of empty center-across cells to its
                // right (`cellLook`). Expand the base rect over the run.
                let baseRect =
                    if look.AcrossTo > cell.Address.Col then
                        let r = cellRect model viewport cell.Address.Row look.AcrossTo
                        { cell.Rect with Width = (r.X + r.Width) - cell.Rect.X }
                    else cell.Rect
                // (A format's `_x` at the text's start / end is drawn as a space and leaves the width of x in Excel's
                // PDF - `"$"#,##0_)` stands a ")" in from the right: the side the text is aligned to is set in by
                // the difference, `CellText.widthPads`.)
                let padLeft, padRight =
                    match measure with
                    | Some m when not (displayText |> Seq.forall (fun c -> c = '#')) ->
                        let lead, trail = widthPads format value
                        let extra (pads: string) = pads |> Seq.sumBy (fun c -> m (string c) - m " ")
                        (if look.Align = AlignLeft && lead <> "" then extra lead else 0.0),
                        (if look.Align = AlignRight && trail <> "" then extra trail else 0.0)
                    | _ -> 0.0, 0.0
                let textRect =
                    { baseRect with
                        X = baseRect.X + padX + iconGutter + leftIndent + padLeft
                        Width = max 0.0 (baseRect.Width - 2.0 * padX - iconGutter - leftIndent - rightIndent - padLeft - padRight) }
                // ShrinkToFit (single-line text only — wrap-mode escapes via
                // line breaks). When the text overflows the cell width and
                // shrinkToFit is set, scale the font size down until it fits,
                // clamped at 1 pt. We need to re-measure with the smaller
                // size so the rest of the rendering uses the actual width.
                let fontPt, measure =
                    if not look.Shrink then fontPt, measure
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
                // (A conditional format's colour, else the number format's section colour - `[Red]` -
                // else the cell's own, else its table's: `cellLook`.)
                let color = look.TextColor |> Option.defaultValue "#000000"
                let align = look.Align
                // Wrapping and manual `\n` line breaks only apply when wrapText
                // is ON (unwrapped, the look's text has none).
                let lines =
                    if look.Wrap then
                        match measure with
                        | Some m when displayText <> "" ->
                            excelWrapLines m (baseRect.Width - printWrapMargin - iconGutter - leftIndent - rightIndent) displayText
                        | _ -> wrapLines measure textRect.Width fontPt displayText
                    else [ displayText ]
                let vAlign = look.VerticalAlign
                let lineCount = lines.Length
                // Where Excel 2019 prints each line (make-print-place-fixture.py "Place", the rowunit fixtures: 24 Normal
                // fonts and 7 cell fonts in rows of 15 - 45 pt, from the grid's lines): at the bottom and in the middle laid
                // out in the pixels of the screen and scaled as the row is (`PrintRowScale` dots a pixel) - at the bottom,
                // the baseline (the row's largest descent of its text at the bottom + 1) px above the cell's bottom, every
                // such cell of the row on that baseline, each line a line of its font (`FontMetrics.lineOf`) above the
                // one below. At the top and in the middle, whatever the Normal font, by the font's own ascent and descent
                // at its size: at the top 4.5 dots lower than that (Calibri 11 90.5 dots below the top, Arial 10 79.5,
                // Segoe UI 10 91.5, under Calibri 11 and Segoe UI 10 alike), in the middle centred. Within a pixel of
                // Excel's PDF.
                let lineBaseline =
                    let rowScale = (if model.PrintRowScale > 0.0 then model.PrintRowScale else 6.05) / 6.25
                    let struct (ownLine, ownDescent) = FontMetrics.lineOf format.FontName fontPt format.Bold format.Italic
                    let height = textRect.Height
                    fun (i: int) ->
                        if format.TextRotation <> 0 || height <= 0.0 then 0.0
                        else
                            let below = float (lineCount - 1 - i) * ownLine * rowScale
                            match vAlign with
                            | VAlignBottom -> height - (max ownDescent (rowDescent cell.Address.Row) + 1.0) * rowScale - below
                            | VAlignTop -> -0.72
                            | VAlignMiddle -> -1e-6
                let richRuns = Map.tryFind cell.Address model.RichRuns
                // Text overflow: a line of text wider than its cell is drawn
                // over the empty cells beside it and cut off at the first
                // cell that holds something (at its own edge when that is
                // the neighbour).
                let textClip =
                    // (Wrapped lines are cut at the cell's edges: the lines past a row's bottom are not printed -
                    // Excel's PDF of make-print-text-fixture.py's 300-pt rows.)
                    if format.WrapText && format.TextRotation = 0 then Some baseRect
                    elif format.WrapText || format.TextRotation <> 0 || horizForIndent = "centercontinuous" || horizForIndent = "fill" then None
                    else
                        let width =
                            match richRuns, viewport.Measurer with
                            | Some runs, Some m when not (List.isEmpty runs) ->
                                runs |> List.sumBy (fun (text, ovr) ->
                                    m text (ovr.FontSize |> Option.defaultValue fontPt) (ovr.FontName |> Option.orElse format.FontName)
                                        (ovr.Bold |> Option.defaultValue format.Bold) (ovr.Italic |> Option.defaultValue format.Italic))
                            | _ ->
                                // (Not measured when even the widest glyphs would fit the cell.)
                                if float lines.Head.Length * fontPt * 1.4 <= textRect.Width then 0.0
                                else
                                    match measure with
                                    | Some m -> m lines.Head
                                    | None -> float lines.Head.Length * fontPt * 0.75
                        let excess = width - textRect.Width
                        if excess <= 0.0 then None
                        else
                            let isMerged = not model.MergedCells.IsEmpty && (findMergedCell cell.Address model).IsSome
                            match (if isMerged then None else textOverflowSides format value) with
                            | None -> Some cell.Rect
                            | Some (toLeft, toRight) ->
                                // How far the text reaches past the cell's edge (the
                                // cell's own padding on that side is its to use);
                                // centred text needs half of the excess on each side.
                                let needLeft = if toLeft then (if toRight then excess / 2.0 else excess) - padX else 0.0
                                let needRight = if toRight then (if toLeft then excess / 2.0 else excess) - padX else 0.0
                                let first, last = textOverflowSpan model cell.Address.Row cell.Address.Col needLeft needRight
                                if first = cell.Address.Col && last = cell.Address.Col then Some cell.Rect
                                else
                                    let leftEdge = cell.Rect.X - (columnLeft model cell.Address.Col - columnLeft model first)
                                    let rightEdge = cell.Rect.X + cell.Rect.Width + (columnLeft model (last + 1) - columnLeft model (cell.Address.Col + 1))
                                    Some { cell.Rect with X = leftEdge; Width = rightEdge - leftEdge }
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
                                Baseline = lineBaseline i
                                Rotation = format.TextRotation
                                Clip = textClip
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
                            Baseline = lineBaseline i
                            Rotation = format.TextRotation
                            Clip = textClip
                        }

        // Pass 4: auto-filter dropdown triangles on table header cells. Excel
        // renders these as a small downward triangle on the right edge of each
        // header cell whose column falls within the table's <autoFilter>
        // range. The triangle is drawn over the header's accent fill — same
        // color as the header text (white on a Medium-style accent).
        for (cell, look, zone) in states do
            let format = look.Format
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
        for (cell, look, zone) in states do
            let format = look.Format
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
        let hasValidations = not model.Validations.IsEmpty
        for (cell, look, zone) in states do
            let format = look.Format
            if (format.HasListDropdown || (hasValidations && showsListDropdown cell.Address model)) && not format.FilterButton then
                let color = "#666666"
                let buttonWidth = 16.0
                let centerX = cell.Rect.X + cell.Rect.Width - buttonWidth / 2.0
                let triHalfWidth = 3.5
                let triTopY = cell.Rect.Y + cell.Rect.Height / 2.0 - 2.0
                for i in 0 .. 3 do
                    let shrink = float i * (triHalfWidth / 4.0)
                    let y = triTopY + float i
                    yield zone, GridLine (centerX - triHalfWidth + shrink, y, centerX + triHalfWidth - shrink, y, color, 1.0)

        // Pass 7: Excel's 16-pixel icon, 2 px inside the left cell edge.
        for (cell, look, zone) in states do
            match look.Format.IconSet with
            | Some icon when icon.Count > 1 && cell.Rect.Width > 0.0 && cell.Rect.Height > 0.0 ->
                for part in iconPaintItems icon.Family icon.Index cell.Rect look.VerticalAlign do
                    yield zone, part
            | _ -> ()
    ]
    // Text that runs over the cells beside its own (its clip is that run of
    // cells) hides the grid lines it crosses, as in Excel: the vertical grid
    // lines inside such a clip, within its row, are not drawn. Borders stay.
    let overflows = ResizeArray<Rect>()
    for _, item in items do
        match item with
        | DrawText { Clip = Some clip; Rect = rect } when clip.X < rect.X - 6.0 || clip.X + clip.Width > rect.X + rect.Width + 6.0 -> overflows.Add clip
        | _ -> ()
    if overflows.Count = 0 then items
    else
        items |> List.filter (fun (_, item) ->
            match item with
            | GridLine (x1, y1, x2, y2, color, width) when x1 = x2 && width = 1.0 && color = gridLineColor ->
                let mutable crossed = false
                for clip in overflows do
                    if x1 > clip.X + 0.5 && x1 < clip.X + clip.Width - 0.5 && y1 >= clip.Y - 0.5 && y2 <= clip.Y + clip.Height + 0.5 then crossed <- true
                not crossed
            | _ -> true)

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
                + frozenRowsHeight model
            let totalWidth =
                viewport.RowHeaderWidth
                + columnLeft model model.ColCount
            yield GridLine (viewport.RowHeaderWidth, y, totalWidth, y, color, w)
        if model.FrozenCols > 0 then
            let x =
                viewport.RowHeaderWidth
                + frozenColsWidth model
            let totalHeight =
                viewport.HeaderHeight
                + rowTop model model.RowCount
            yield GridLine (x, viewport.HeaderHeight, x, totalHeight, color, w)
    ]

/// Selection-only paint items, drawn on top of the cached grid. Separated so
/// that cursor moves (which only change `model.Selection`) don't invalidate
/// the cell paint items.
let selectionPaintItems (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    let model = model |> applyFilterRowHeights
    let r1, c1, r2, c2 =
        match model.Selection with
        | SingleCell (r, c) -> r, c, r, c
        | Range (sr, sc, er, ec) ->
            let r1 = min sr er
            let r2 = max sr er
            let c1 = min sc ec
            let c2 = max sc ec
            r1, c1, r2, c2
    let bounds = {
        X = columnX model viewport c1
        Y = rowY model viewport r1
        Width =
            ([ c1 .. c2 ] |> List.sumBy (fun c -> getColumnWidth c model))
        Height =
            ([ r1 .. r2 ] |> List.sumBy (fun r -> getRowHeight r model))
    }
    [
        // Light mint range fill ONLY for multi-cell selections (Excel doesn't
        // tint the active cell). The outer green outline is drawn for both.
        if r1 <> r2 || c1 <> c2 then
            yield FillRect (bounds, "#E8F5E9")
        yield StrokeRect (bounds, "#217346", 2.0)
    ]

/// A labelled frame standing for a drawing object that is not painted as
/// itself: a chart or another kept object, or a picture on a renderer that
/// cannot decode images.
let drawingFrameItems (rect: Rect) (label: string) : PaintItem list =
    [ FillRect (rect, "#f6f8fa")
      StrokeRect (rect, "#8a94a6", 1.0)
      DrawText {
          Text = label
          Rect = rect
          Color = "#4a5568"
          FontSize = 10.0
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
          Baseline = 0.0
          Rotation = 0
          Clip = None
      } ]

/// The sheet's drawing objects, back to front, at sheet coordinates (the same
/// space as the cell items: headers included, scroll not applied). Pictures
/// are `DrawImage`; charts and other kept objects a labelled frame.
/// A scene (`Charts.objectScene`) as paint items, for the renderers that are not SVG: PDF. Everything is
/// flattened to the scene's own coordinates, because text inside a `VecGroup` is not painted. A text
/// keeps a quarter turn; other turns of text, clips, dashes and gradients (the first stop) are not kept.
let scenePaintItems (scene: Rbt.Spreadsheet.Charts.Scene) : PaintItem list =
    let items = ResizeArray<PaintItem>()
    let rgb (color: string) (at: int) = System.Convert.ToInt32 (color.Substring (at, 2), 16)
    /// A colour as an opaque one on white paper.
    let paper (color: string) (opacity: float) =
        if opacity >= 1.0 || color.Length <> 7 then color
        else
            let blend (at: int) = int (System.Math.Round (255.0 + (float (rgb color at) - 255.0) * opacity))
            sprintf "#%02X%02X%02X" (blend 1) (blend 3) (blend 5)
    let solid (fill: string option) (opacity: float) =
        match fill with
        | Some "none" | None -> None
        | Some color when color.StartsWith "gradient(" ->
            // gradient(angle;offset color opacity;...)
            let stops = color.Substring(9).TrimEnd(')').Split(';')
            if stops.Length > 1 then (match stops.[1].Split(' ') with [| _; first; _ |] -> Some (paper first opacity) | _ -> None) else None
        | Some color -> Some (paper color opacity)
    let outline (stroke: string option) (width: float) = stroke |> Option.filter (fun color -> color <> "none" && width > 0.0) |> Option.map (fun color -> color, width)
    // (A path with neither would stay the renderer's current path, and be painted with the next.)
    let addPath segments fill stroke = if Option.isSome fill || Option.isSome stroke then items.Add (VecPath (segments, fill, stroke))
    // (The typeface of the text being walked: a chart's own, inside a `PFace`.)
    let face = ref "Calibri"
    let rec walk (a, b, c, d, e, f) (item: Rbt.Spreadsheet.Charts.Primitive) =
        let at x y = a * x + c * y + e, b * x + d * y + f
        let move x y = VMove (at x y)
        let line x y = VLine (at x y)
        let curve x1 y1 x2 y2 x y =
            let p1x, p1y = at x1 y1
            let p2x, p2y = at x2 y2
            let px, py = at x y
            VCubic (p1x, p1y, p2x, p2y, px, py)
        let text (x: float) (y: float) (value: string) (size: float) (color: string) (anchor: string) (bold: bool) (italic: bool) (rotation: int) =
            let px, py = at x y
            let wide = 2000.0
            let left = (match anchor with "middle" -> px - wide / 2.0 | "end" -> px - wide | _ -> px)
            items.Add (DrawText {
                Text = value
                // (A turned text starts at the bottom left of its rectangle: a rectangle as long as it, about its middle.)
                Rect = (if rotation = 0 then { X = left; Y = py - size; Width = wide; Height = size * 2.0 }
                        else
                            let long = Rbt.Spreadsheet.Charts.runWidth size bold value
                            { X = px - size / 2.0 - 4.0; Y = py - long / 2.0; Width = size * 1.2; Height = long + 2.0 })
                Color = color
                FontSize = size * 0.75
                FontName = Some face.Value
                Bold = bold
                Italic = italic
                Underline = false
                UnderlineDouble = false
                Strikethrough = false
                Align = (match anchor with "middle" -> AlignCenter | "end" -> AlignRight | _ -> AlignLeft)
                VerticalAlign = VAlignMiddle
                LineIndex = 0
                LineCount = 1
                BaselineShift = 0.0
                Baseline = 0.0
                Rotation = rotation
                Clip = None })
        match item with
        | Rbt.Spreadsheet.Charts.PRect (x, y, width, height, fill, opacity, stroke, strokeWidth, radius) ->
            let r = min radius (min width height / 2.0)
            let k = r * 0.4477
            let segments =
                if r <= 0.0 then [ move x y; line (x + width) y; line (x + width) (y + height); line x (y + height); VClose ]
                else
                    [ move (x + r) y; line (x + width - r) y; curve (x + width - k) y (x + width) (y + k) (x + width) (y + r)
                      line (x + width) (y + height - r); curve (x + width) (y + height - k) (x + width - k) (y + height) (x + width - r) (y + height)
                      line (x + r) (y + height); curve (x + k) (y + height) x (y + height - k) x (y + height - r)
                      line x (y + r); curve x (y + k) (x + k) y (x + r) y; VClose ]
            addPath segments (solid fill opacity) (outline stroke strokeWidth)
        | Rbt.Spreadsheet.Charts.PEllipse (cx, cy, rx, ry, fill, opacity, stroke, strokeWidth) ->
            let kx, ky = rx * 0.5523, ry * 0.5523
            let segments =
                [ move (cx + rx) cy; curve (cx + rx) (cy + ky) (cx + kx) (cy + ry) cx (cy + ry)
                  curve (cx - kx) (cy + ry) (cx - rx) (cy + ky) (cx - rx) cy
                  curve (cx - rx) (cy - ky) (cx - kx) (cy - ry) cx (cy - ry)
                  curve (cx + kx) (cy - ry) (cx + rx) (cy - ky) (cx + rx) cy; VClose ]
            addPath segments (solid fill opacity) (outline stroke strokeWidth)
        | Rbt.Spreadsheet.Charts.PLine (x1, y1, x2, y2, stroke, strokeWidth, _) ->
            addPath [ move x1 y1; line x2 y2 ] None (outline (Some stroke) strokeWidth)
        | Rbt.Spreadsheet.Charts.PPolyline (points, fill, opacity, stroke, strokeWidth, _) ->
            match points with
            | [] -> ()
            | (x, y) :: rest ->
                let color = solid fill opacity
                let segments = move x y :: (rest |> List.map (fun (px, py) -> line px py))
                addPath (if color.IsSome then segments @ [ VClose ] else segments) color (outline stroke strokeWidth)
        | Rbt.Spreadsheet.Charts.PPath (path, fill, opacity, stroke, strokeWidth) ->
            let segments =
                Rbt.Spreadsheet.Charts.pathSegments path
                |> List.map (function
                    | Rbt.Spreadsheet.Charts.MoveTo (x, y) -> move x y
                    | Rbt.Spreadsheet.Charts.LineTo (x, y) -> line x y
                    | Rbt.Spreadsheet.Charts.CubicTo (x1, y1, x2, y2, x, y) -> curve x1 y1 x2 y2 x y
                    | Rbt.Spreadsheet.Charts.ClosePath -> VClose)
            addPath segments (solid fill opacity) (outline stroke strokeWidth)
        | Rbt.Spreadsheet.Charts.PText (x, y, value, size, color, anchor, bold, italic, rotate) ->
            text x y value size color anchor bold italic (if rotate = -90.0 || rotate = 270.0 then 90 elif rotate = 90.0 then 180 else 0)
        | Rbt.Spreadsheet.Charts.PTextRuns (x, y, anchor, runs) ->
            let width = runs |> List.sumBy (fun (value, size, _, bold, _) -> Rbt.Spreadsheet.Charts.runWidth size bold value)
            let mutable left = (match anchor with "middle" -> x - width / 2.0 | "end" -> x - width | _ -> x)
            for value, size, color, bold, italic in runs do
                text left y value size color "start" bold italic 0
                left <- left + Rbt.Spreadsheet.Charts.runWidth size bold value
        | Rbt.Spreadsheet.Charts.PImage (x, y, width, height, href) ->
            let px, py = at x y
            let comma = href.IndexOf ','
            if href.StartsWith "data:" && comma > 0 then
                let contentType = href.Substring(5, comma - 5).Split(';').[0]
                items.Add (DrawImage ({ X = px; Y = py; Width = width * a; Height = height * d },
                                      { ContentType = contentType; Data = System.Convert.FromBase64String (href.Substring (comma + 1)); Label = "" }))
        | Rbt.Spreadsheet.Charts.PGroup (x, y, _, _, children) ->
            let shift = (a, b, c, d, a * x + c * y + e, b * x + d * y + f)
            children |> List.iter (walk shift)
        | Rbt.Spreadsheet.Charts.PTransform (a2, b2, c2, d2, e2, f2, children) ->
            let composed = (a * a2 + c * b2, b * a2 + d * b2, a * c2 + c * d2, b * c2 + d * d2, a * e2 + c * f2 + e, b * e2 + d * f2 + f)
            children |> List.iter (walk composed)
        // (A shadow is not printed: what casts it is.)
        | Rbt.Spreadsheet.Charts.PShadow (_, _, _, _, _, _, _, _, _, children) -> children |> List.iter (walk (a, b, c, d, e, f))
        | Rbt.Spreadsheet.Charts.PFace (inner, children) ->
            let outer = face.Value
            face.Value <- inner
            children |> List.iter (walk (a, b, c, d, e, f))
            face.Value <- outer
    scene.Items |> List.iter (walk (1.0, 0.0, 0.0, 1.0, 0.0, 0.0))
    List.ofSeq items

/// The charts and shapes of a sheet as the overlays of a PDF: anchored to their first cell, each drawn
/// by `sceneOf` (its anchor's XML, its width and height) in the size it has on the sheet.
let drawingOverlays (sceneOf: string -> float -> float -> Rbt.Spreadsheet.Charts.Scene option) (model: SpreadsheetModel) : (int * int * float * float * float * float * PaintItem list) list =
    [ for drawing in model.Drawings do
        match drawing.Content with
        | KeptObject (_, xml) ->
            let rect = drawingRect model drawing
            let items =
                match sceneOf xml rect.Width rect.Height with
                | Some scene -> scenePaintItems scene
                | None -> drawingFrameItems { X = 0.0; Y = 0.0; Width = rect.Width; Height = rect.Height } (drawingLabel drawing)
            yield (drawing.Anchor.FromRow, drawing.Anchor.FromCol, drawing.Anchor.FromOffsetX, drawing.Anchor.FromOffsetY, rect.Width, rect.Height, items)
        | _ -> () ]

/// Excel snapshots a sparkline in its screen cell, then stretches that
/// picture to the cell's printed width/height. Resolve the marks in screen
/// pixels first, and apply the same scale when placing the PDF overlay.
let sparklineOverlays (source: Rbt.Spreadsheet.Sparklines.Source) (screen: SpreadsheetModel) (printed: SpreadsheetModel) : (int * int * float * float * float * float * PaintItem list) list =
    [for group in screen.Sparklines do
        for addr, _ in group.Items do
            match Rbt.Spreadsheet.Sparklines.inCell source [group] addr with
            | Some (_,values,scale) ->
                let sw, sh = getColumnWidth addr.Col screen, getRowHeight addr.Row screen
                let pw, ph = getColumnWidth addr.Col printed, getRowHeight addr.Row printed
                if sw > 0.0 && sh > 0.0 && pw > 0.0 && ph > 0.0 then
                    let rect = { X = 0.0; Y = 0.0; Width = sw; Height = sh }
                    let artwork = sparklinePaintItems group values scale rect
                    if not artwork.IsEmpty then
                        yield addr.Row,addr.Col,0.0,0.0,pw,ph,
                              [VecGroup ((pw/sw,0.0,0.0,ph/sh,0.0,0.0),artwork)]
            | None -> ()]

let buildDrawingPaintItems (model: SpreadsheetModel) (viewport: Viewport) : PaintItem list =
    if model.Drawings.IsEmpty then []
    else
        [ for drawing in model.Drawings do
            let rect = drawingRect model drawing
            if rect.Width >= 1.0 && rect.Height >= 1.0 then
                let placed = { rect with X = rect.X + viewport.RowHeaderWidth; Y = rect.Y + viewport.HeaderHeight }
                match drawing.Content with
                | Picture (contentType, _, data) ->
                    yield DrawImage (placed, { ContentType = contentType; Data = data; Label = drawingLabel drawing })
                | KeptObject _ -> yield! drawingFrameItems placed (drawingLabel drawing) ]

let buildPaintItemsWithSource (source: Rbt.Spreadsheet.Sparklines.Source) (model: SpreadsheetModel) (viewport: Viewport): PaintItem list =
    let model = model |> applyFilterRowHeights
    [
        yield FillRect ({ X = 0.0; Y = 0.0; Width = viewport.Width; Height = viewport.Height }, "#ffffff")
        yield! buildHeaderPaintItems model viewport
        // Static (PNG/PDF) path: drop the zone tag; everything paints at
        // absolute cell coordinates because ScrollX/Y are zero.
        for (_zone, item) in buildCellPaintItems model viewport do
            yield item
        for (_zone, item) in sparklineCellItems source model viewport do
            yield item
        // Pictures and placeholders for kept objects float over the cells.
        yield! buildDrawingPaintItems model viewport
        // Freeze divider lines on top of cell strokes but under any
        // selection / DrawText emitted later. The cell-item zone tag is
        // irrelevant here; we draw the divider in the model coordinate
        // system that's already baked into the cell rects above.
        yield! buildFreezeDividerItems model viewport
    ]

let buildPaintItems (model: SpreadsheetModel) (viewport: Viewport) =
    buildPaintItemsWithSource (Rbt.Spreadsheet.Sparklines.modelSource model) model viewport


let private translatePaintItem viewport zone item =
    let tx = toViewportRect viewport zone
    match item with
    | FillRect (rect, color) -> FillRect (tx rect, color)
    | StrokeRect (rect, color, width) -> StrokeRect (tx rect, color, width)
    | DrawText text -> DrawText { text with Rect = tx text.Rect; Clip = text.Clip |> Option.map tx }
    | GridLine (x1, y1, x2, y2, color, width) ->
        let p1 = tx { X = x1; Y = y1; Width = 0.0; Height = 0.0 }
        let p2 = tx { X = x2; Y = y2; Width = 0.0; Height = 0.0 }
        GridLine (p1.X, p1.Y, p2.X, p2.Y, color, width)
    | VecPath (segs, fill, stroke) ->
        // The zone offset is a pure translation; shift every path point by it.
        let o = tx { X = 0.0; Y = 0.0; Width = 0.0; Height = 0.0 }
        let mv x y = x + o.X, y + o.Y
        let segs' =
            segs |> List.map (function
                | VMove (x, y) -> let a, b = mv x y in VMove (a, b)
                | VLine (x, y) -> let a, b = mv x y in VLine (a, b)
                | VCubic (x1, y1, x2, y2, x, y) ->
                    let a1, b1 = mv x1 y1
                    let a2, b2 = mv x2 y2
                    let a, b = mv x y
                    VCubic (a1, b1, a2, b2, a, b)
                | VClose -> VClose)
        VecPath (segs', fill, stroke)
    | VecGroup ((a, b, c, d, e, f), children) ->
        // Fold the zone offset into the group's translation component;
        // children stay in group-local coordinates.
        let o = tx { X = 0.0; Y = 0.0; Width = 0.0; Height = 0.0 }
        VecGroup ((a, b, c, d, e + o.X, f + o.Y), children)
    | DrawImage (rect, image) -> DrawImage (tx rect, image)
    | Clipped _ -> item

let buildViewportPaintItemsWithSource (source: Rbt.Spreadsheet.Sparklines.Source) (model: SpreadsheetModel) (viewport: Viewport): ViewportPaintItem list =
    let model = model |> applyFilterRowHeights
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
                            Baseline = 0.0
                            Rotation = 0
                            Clip = None
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
                            Baseline = 0.0
                            Rotation = 0
                            Clip = None
                        })
                }

        for (zone, item) in buildCellPaintItems model viewport do
            yield { Zone = zone; Item = translatePaintItem viewport zone item }
        for (zone, item) in sparklineCellItems source model viewport do
            yield { Zone = zone; Item = translatePaintItem viewport zone item }

        // Drawing objects float over the cells and scroll with the body.
        for item in buildDrawingPaintItems model viewport do
            yield { Zone = BodyZone; Item = translatePaintItem viewport BodyZone item }

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

let buildViewportPaintItems (model: SpreadsheetModel) (viewport: Viewport) =
    buildViewportPaintItemsWithSource (Rbt.Spreadsheet.Sparklines.modelSource model) model viewport
