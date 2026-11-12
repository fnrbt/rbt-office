module Rbt.SpreadsheetNative.GridRender

open Rbt.Renderer
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Layout
open Rbt.Spreadsheet.Paint

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private resizeArrayToList (items: System.Collections.Generic.List<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private trimBounds (value: string) =
    let mutable start = 0
    let mutable last = value.Length - 1
    while start <= last && System.Char.IsWhiteSpace value.[start] do
        start <- start + 1
    while last >= start && System.Char.IsWhiteSpace value.[last] do
        last <- last - 1
    start, last - start + 1


let private parseHexColor (value: string) =
    let value =
        match value.Trim().ToLowerInvariant() with
        | "black" -> "#000000"
        | "white" -> "#ffffff"
        | "red" -> "#ff0000"
        | "blue" -> "#0000ff"
        | "green" -> "#008000"
        | "darkgray" | "darkgrey" -> "#808080"
        | "lightgray" | "lightgrey" -> "#c0c0c0"
        | other -> other
    let hexValue (c: char) =
        if c >= '0' && c <= '9' then int c - int '0'
        elif c >= 'a' && c <= 'f' then int c - int 'a' + 10
        elif c >= 'A' && c <= 'F' then int c - int 'A' + 10
        else -1
    let parseByteAt offset =
        let hi = hexValue value.[offset]
        let lo = hexValue value.[offset + 1]
        if hi < 0 || lo < 0 then None else Some (byte ((hi <<< 4) ||| lo))
    let parseShortAt offset =
        let v = hexValue value.[offset]
        if v < 0 then None else Some (byte ((v <<< 4) ||| v))
    let start0, length0 = trimBounds value
    let start, length =
        if length0 > 0 && value.[start0] = '#' then start0 + 1, length0 - 1
        else start0, length0
    match length with
    | 3 ->
        match parseShortAt start, parseShortAt (start + 1), parseShortAt (start + 2) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 6 ->
        match parseByteAt start, parseByteAt (start + 2), parseByteAt (start + 4) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 8 ->
        match parseByteAt start, parseByteAt (start + 2), parseByteAt (start + 4), parseByteAt (start + 6) with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | _ -> Color.black
    | _ -> Color.black

/// Some fonts set OS/2 fsSelection bit 7 (USE_TYPO_METRICS) to opt into
/// typographic ascender/descender for line stacking. When that bit is
/// set we honor it; otherwise we use the GDI win metrics Excel anchors
/// its cell-text box layout to.
let private useTypoMetrics (font: Rbt.Renderer.Text.FontData) =
    match font.OtFont.OS2 with
    | Some os2 -> (os2.FsSelection &&& 128us) <> 0us
    | None -> false

/// OS/2 winAscent scaled to pixels — the "WIN ascent" GDI/Excel uses as
/// the line top. Derived here from the (public) FontData.OtFont rather
/// than in Rbt.Renderer.Text so the general-purpose renderer doesn't carry a
/// GDI-specific opinion. USE_TYPO_METRICS fonts get the typo ascender +
/// half the typo line gap instead; with no OS/2 table at all we fall
/// back to the hhea ascender.
let private winAscent (font: Rbt.Renderer.Text.FontData) (fontSize: float) : float =
    let unitsPerEm = float font.OtFont.Head.UnitsPerEm
    let units =
        match font.OtFont.OS2 with
        | Some os2 when not (useTypoMetrics font) -> float os2.WinAscent
        | Some os2 -> float os2.TypoAscender + float os2.TypoLineGap / 2.0
        | None -> float font.OtFont.Hhea.Ascender
    units * fontSize / unitsPerEm

/// OS/2 winDescent as a positive pixel distance (baseline → em bottom).
/// Mirrors winAscent's USE_TYPO_METRICS / no-OS/2 fallbacks.
let private winDescent (font: Rbt.Renderer.Text.FontData) (fontSize: float) : float =
    let unitsPerEm = float font.OtFont.Head.UnitsPerEm
    let units =
        match font.OtFont.OS2 with
        | Some os2 when not (useTypoMetrics font) -> float os2.WinDescent
        | Some os2 -> float -os2.TypoDescender + float os2.TypoLineGap / 2.0
        | None -> - (float font.OtFont.Hhea.Descender)
    units * fontSize / unitsPerEm

/// winAscent + winDescent — the GDI line-box height.
let private winLineHeight (font: Rbt.Renderer.Text.FontData) (fontSize: float) : float =
    winAscent font fontSize + winDescent font fontSize

let private toRendererCommands (text: PaintText) =
    // Excel font sizes (XLSX <sz val="11"/>) are in points; the Renderer takes
    // pixel units. Convert pt→px so a Calibri 11-pt cell renders at Excel's
    // glyph height. Row/column heights are pre-converted via 96/72 upstream.
    let primary = Rbt.OfficeNative.Fonts.resolve text.FontName text.Bold text.Italic
    let pxSize = text.FontSize * 96.0 / 72.0

    let segments = Rbt.OfficeNative.Fonts.segmentByFontCoverage primary text.Bold text.Text

    let segmentLayouts = System.Collections.Generic.List<_>(segments.Count)
    let mutable totalWidth = 0.0
    for f, s in segments do
        let layout = Rbt.Renderer.Text.layout f pxSize s
        let width = Rbt.Renderer.Text.measureWidth f pxSize s
        segmentLayouts.Add(f, layout, width)
        totalWidth <- totalWidth + width
    // Cell-text rotation (Excel `<alignment textRotation>`): 1..90 = degrees
    // CCW, 91..180 = (value-90) degrees CW, 255 = stacked (handled as flat
    // here). The PDF renderer already rotates; the native/interactive path
    // laid every cell flat. We lay the text out horizontally from a leading-
    // edge anchor, then rotate the whole glyph stream around that anchor via
    // a Cmd.Group transform — mirroring the (Excel-validated) PDF path: anchor
    // at left + 4 px, baseline at the cell bottom for CCW (text rises) or at
    // the cell top + ascent for CW (text descends, staying inside the cell).
    let isRotated = text.Rotation > 0 && text.Rotation <> 255
    let alignedX =
        if isRotated then text.Rect.X + 4.0
        else
            match text.Align with
            | AlignLeft -> text.Rect.X
            | AlignCenter -> text.Rect.X + (text.Rect.Width - totalWidth) / 2.0
            | AlignRight -> text.Rect.X + text.Rect.Width - totalWidth
    // Baseline from OS/2 winAscent/winDescent of the PRIMARY font. Mixing
    // ascents across fallback fonts would jitter the baseline; sticking to
    // the primary's ascent matches Excel/GDI's font-linking behavior, which
    // keeps the baseline anchored to the primary font's metrics.
    let ascent = winAscent primary pxSize
    let lineHeight = winLineHeight primary pxSize
    let blockHeight = lineHeight * float text.LineCount
    // Excel quirk: vertical="top" is ignored on rows that aren't taller than
    // about one line-height. pandas/openpyxl emit `<alignment vertical="top"/>`
    // on every header cell, but Excel renders default-height rows as bottom
    // (Excel "General") regardless. Mirror that — when the cell has at most
    // ~half a line of vertical slack, collapse top → bottom. Tall wrap-text
    // rows still respect the hint and place text near the cell top.
    let slack = text.Rect.Height - blockHeight
    let effectiveVAlign =
        match text.VerticalAlign with
        | VAlignTop when slack < lineHeight * 0.5 -> VAlignBottom
        | va -> va
    let blockTopY =
        match effectiveVAlign with
        | VAlignTop -> text.Rect.Y
        | VAlignMiddle -> text.Rect.Y + (text.Rect.Height - blockHeight) / 2.0
        | VAlignBottom -> text.Rect.Y + text.Rect.Height - blockHeight
    // Super/subscript shifts the baseline relative to the line's normal
    // baseline. PaintText.BaselineShift is in pt at the cell's base size;
    // convert to px the same way pt → px conversion happens elsewhere.
    let baselineShiftPx = text.BaselineShift * 96.0 / 72.0
    // Rotated text is laid out at a provisional bottom-left position; the
    // Group transform below rotates it and then aligns its rotated bounding
    // box within the cell, so this exact baseline only needs to be consistent.
    let baseline =
        if isRotated then text.Rect.Y + text.Rect.Height - 2.0
        else blockTopY + float text.LineIndex * lineHeight + ascent + baselineShiftPx
    let color = Solid (parseHexColor text.Color)
    let mutable cursor = alignedX
    let innerCommands = System.Collections.Generic.List<Cmd>(segmentLayouts.Count + 3)
    for _, layout, width in segmentLayouts do
        innerCommands.Add(Cmd.DrawGlyphs(vec2 cursor baseline, layout, color))
        cursor <- cursor + width
    let lineStyle = {
        Width = px 1.0
        Cap = Butt
        Join = Miter
        Paint = color
        DashArray = []
        DashOffset = px 0.0
        MiterLimit = 4.0
    }
    if text.Underline && totalWidth > 0.0 then
        // Single underline stroke spans all glyph segments. Position at
        // ~1 px below the baseline (matches Excel/GDI; OpenType's
        // OS/2.sxUnderlinePosition isn't surfaced by our font loader,
        // so use a Calibri-tuned constant). Stroke thickness 1 px.
        let underlineY = baseline + max 1.0 (pxSize * 0.08)
        innerCommands.Add(Cmd.StrokePath(
            [ MoveTo (vec2 alignedX underlineY)
              LineTo (vec2 (alignedX + totalWidth) underlineY) ],
            lineStyle))
        if text.UnderlineDouble then
            // Double underline: a second stroke ~2 px below the first.
            let underlineY2 = underlineY + max 2.0 (pxSize * 0.12)
            innerCommands.Add(Cmd.StrokePath(
                [ MoveTo (vec2 alignedX underlineY2)
                  LineTo (vec2 (alignedX + totalWidth) underlineY2) ],
                lineStyle))
    if text.Strikethrough && totalWidth > 0.0 then
        // Strikethrough at ~30% above baseline — visually centers on
        // the lowercase x-height. (OpenType yStrikeoutPosition would
        // be more accurate; using a Calibri-tuned constant to match
        // Excel/GDI.)
        let strikeY = baseline - pxSize * 0.28
        innerCommands.Add(Cmd.StrokePath(
            [ MoveTo (vec2 alignedX strikeY)
              LineTo (vec2 (alignedX + totalWidth) strikeY) ],
            lineStyle))
    let inner = resizeArrayToList innerCommands
    if not isRotated then inner
    else
        // Screen space is y-down so Matrix3x2.rotate(θ) turns clockwise;
        // Excel's 1..90 is CCW (θ negative) and 91..180 is CW (θ positive).
        let theta =
            if text.Rotation <= 90 then -(float text.Rotation) * System.Math.PI / 180.0
            else float (text.Rotation - 90) * System.Math.PI / 180.0
        // Rotate about the run's anchor, then translate so the rotated bounding
        // box is aligned within the cell the way Excel does (derived from a
        // 13-angle × 3-length sweep, tools/excel-parity/build-rotation-sweep.vbs):
        // RIGHT-align the box when 90 ≤ textRotation < 180 (CCW-vertical and
        // CW-angled), LEFT-align otherwise; always bottom-align. Bottom-align
        // also keeps vertical placement matching (rot 90 at y 88–122 vs 88–121).
        let descent = winDescent primary pxSize
        let cosT, sinT = cos theta, sin theta
        let rotPt x y =
            let dx, dy = x - alignedX, y - baseline
            alignedX + dx * cosT - dy * sinT, baseline + dx * sinT + dy * cosT
        let x1, y1 = rotPt alignedX (baseline - ascent)
        let x2, y2 = rotPt (alignedX + totalWidth) (baseline - ascent)
        let x3, y3 = rotPt alignedX (baseline + descent)
        let x4, y4 = rotPt (alignedX + totalWidth) (baseline + descent)
        let minX = min (min x1 x2) (min x3 x4)
        let maxX = max (max x1 x2) (max x3 x4)
        let maxY = max (max y1 y2) (max y3 y4)
        let rightAlign = text.Rotation >= 90 && text.Rotation < 180
        let shiftX =
            if rightAlign then (text.Rect.X + text.Rect.Width - 2.0) - maxX
            else (text.Rect.X + 2.0) - minX
        let shiftY = (text.Rect.Y + text.Rect.Height - 2.0) - maxY
        let rot =
            Matrix3x2.multiply
                (Matrix3x2.multiply (Matrix3x2.translate -alignedX -baseline) (Matrix3x2.rotate theta))
                (Matrix3x2.translate alignedX baseline)
        let xform = Matrix3x2.multiply rot (Matrix3x2.translate shiftX shiftY)
        [ Cmd.Group(xform, 1.0, inner) ]

let rec private appendPaintItem (commands: System.Collections.Generic.List<Cmd>) item =
    match item with
    | FillRect (rect, color) ->
        commands.Add(Cmd.FillRect(Rbt.Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color))
    | StrokeRect (rect, color, width) ->
        commands.Add(Cmd.StrokeRect(Rbt.Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color, px width))
    | DrawText text ->
        // Text wider than its cell is cut off at the cells it may run over.
        match text.Clip with
        | Some clip when clip.Width > 0.0 && clip.Height > 0.0 ->
            let x, y, r, b = clip.X, clip.Y, clip.X + clip.Width, clip.Y + clip.Height
            let path = [ MoveTo (vec2 x y); LineTo (vec2 r y); LineTo (vec2 r b); LineTo (vec2 x b); Close ]
            commands.Add(Cmd.Clip(path, NonZero, List.ofSeq (toRendererCommands text)))
        | None -> commands.AddRange(toRendererCommands text)
        | Some _ -> ()
    | GridLine (x1, y1, x2, y2, color, width) ->
        // Pixel-snap axis-aligned strokes so they cover WHOLE pixels — like
        // Excel's GDI — instead of straddling the grid. A stroke centred on an
        // integer coordinate with an odd width spreads half a pixel either
        // side, so its edges anti-alias to ~50% grey (the "half the line is
        // lighter" artefact). Rounding the width and centring odd widths on a
        // pixel centre (n+0.5), even widths on a boundary (n), keeps the line
        // crisp. Diagonals (neither axis constant) are left unsnapped.
        let w = max 1.0 (System.Math.Round width)
        let snap (c: float) =
            if int w % 2 = 1 then floor c + 0.5 else System.Math.Round c
        let sx1, sy1, sx2, sy2 =
            if abs (y1 - y2) < 0.01 then let sy = snap y1 in x1, sy, x2, sy
            elif abs (x1 - x2) < 0.01 then let sx = snap x1 in sx, y1, sx, y2
            else x1, y1, x2, y2
        let style = {
            // Butt caps: with pixel-snapped strokes the box sides meet crisply
            // at the shared corner pixel. Square caps were tried but their
            // half-width extension lands off the pixel grid and anti-aliases to
            // a grey blob at the corner ("sticks out like a ball").
            Width = px w
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = []
            DashOffset = px 0.0
            MiterLimit = 4.0
        }
        commands.Add(Cmd.StrokePath([MoveTo (vec2 sx1 sy1); LineTo (vec2 sx2 sy2)], style))
    | VecPath (segs, fill, stroke) ->
        let pathSegments = System.Collections.Generic.List<_>()
        for seg in segs do
            match seg with
            | VMove (x, y) -> pathSegments.Add(MoveTo (vec2 x y))
            | VLine (x, y) -> pathSegments.Add(LineTo (vec2 x y))
            | VCubic (x1, y1, x2, y2, x, y) -> pathSegments.Add(CubicTo (vec2 x1 y1, vec2 x2 y2, vec2 x y))
            | VClose -> pathSegments.Add Close
        let ps = resizeArrayToList pathSegments
        match fill with
        | Some c -> commands.Add(Cmd.FillPath(ps, Solid (parseHexColor c), NonZero))
        | None -> ()
        match stroke with
        | Some (c, w) ->
            commands.Add(Cmd.StrokePath(ps, { Width = px w; Cap = Butt; Join = Miter
                                              Paint = Solid (parseHexColor c)
                                              DashArray = []; DashOffset = px 0.0; MiterLimit = 4.0 }))
        | None -> ()
    | Clipped (clip, children) ->
        if clip.Width > 0.0 && clip.Height > 0.0 then
            let x, y, r, b = clip.X, clip.Y, clip.X + clip.Width, clip.Y + clip.Height
            let path = [ MoveTo (vec2 x y); LineTo (vec2 r y); LineTo (vec2 r b); LineTo (vec2 x b); Close ]
            let childCommands = System.Collections.Generic.List<Cmd>()
            for child in children do appendPaintItem childCommands child
            commands.Add(Cmd.Clip(path, NonZero, resizeArrayToList childCommands))
    | VecGroup ((a, b, c, d, e, f), children) ->
        let xform = { M11 = a; M12 = b; M21 = c; M22 = d; DX = e; DY = f }
        let childCommands = System.Collections.Generic.List<Cmd>()
        for child in children do
            appendPaintItem childCommands child
        commands.Add(Cmd.Group(xform, 1.0, resizeArrayToList childCommands))
    | DrawImage (rect, image) ->
        if rect.Width > 0.0 && rect.Height > 0.0 then
            let fb = Rbt.OfficeNative.Images.decode image.Data
            commands.Add(Cmd.DrawImage(fb, Rbt.Renderer.Types.rect 0.0 0.0 (float fb.W) (float fb.H), Rbt.Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, true))


let paintItemsToCommands (paintItems: PaintItem list) =
    let commands = System.Collections.Generic.List<Cmd>()
    for item in paintItems do
        appendPaintItem commands item
    resizeArrayToList commands

let measureText
    (text: string) (fontPt: float)
    (fontName: string option) (bold: bool) (italic: bool) : float =
    Rbt.OfficeNative.Fonts.measure text (fontPt * 96.0 / 72.0) fontName bold italic

let private viewportTextMeasurer = Some measureText

let viewport scrollX scrollY width height : Viewport =
    // Match Excel's column/row header pixel sizes at 96 DPI / 100% zoom so
    // parity screenshots align cell-for-cell with our PNG renders. Measured
    // against artifacts/excel/pandas-test1-sheet1-headings.png: the column
    // header strip is 20 px tall (Calibri 11), the row header strip is 28 px
    // wide for two-digit row numbers, and each row in the default 15-pt
    // height renders at 21 source px.
    let headerHeight = 20.0
    let rowHeaderWidth = 28.0
    { ScrollX = scrollX
      ScrollY = scrollY
      Width = width
      Height = height
      HeaderHeight = headerHeight
      RowHeaderWidth = rowHeaderWidth
      Measurer = viewportTextMeasurer }

let viewportOf (model: SpreadsheetModel) width height : Viewport =
    viewport 0.0 0.0 width height

let commands (model: SpreadsheetModel) (viewport: Viewport) : Cmd list =
    let grouped = buildViewportPaintItems model viewport |> List.groupBy (fun item -> item.Zone) |> Map.ofList
    [ for zone in [ CornerZone; BodyZone; FrozenRowsZone; FrozenColsZone; FrozenCornerZone; ColumnHeaderZone; RowHeaderZone ] do
        match Map.tryFind zone grouped with
        | None -> ()
        | Some items ->
            let clip = viewportClipFor model viewport zone
            if clip.Width > 0.0 && clip.Height > 0.0 then
                let x, y, r, b = clip.X, clip.Y, clip.X + clip.Width, clip.Y + clip.Height
                let path = [ MoveTo (vec2 x y); LineTo (vec2 r y); LineTo (vec2 r b); LineTo (vec2 x b); Close ]
                yield Cmd.Clip(path, NonZero, items |> List.map (fun item -> item.Item) |> paintItemsToCommands) ]

let renderSceneCommands width height items =
    Cmd.FillRect(Rbt.Renderer.Types.rect 0.0 0.0 width height, Color.white) :: paintItemsToCommands items
