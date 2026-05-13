module Composit.Ui.SpreadsheetGrid

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private defaultFont () = Composit.Ui.Layout.font()

/// Cache of font variants keyed by (familyKey, bold, italic). familyKey is the
/// normalized lower-case font name we use to pick a TTF set ("calibri",
/// "arial", "cambria", "times"); unknown families fall back to "calibri".
let private fontCache = System.Collections.Generic.Dictionary<string * bool * bool, Renderer.Text.FontData option>()

let private fontDirs () =
    let cwd = System.IO.Directory.GetCurrentDirectory()
    [
        System.Environment.GetEnvironmentVariable("RB_WORD_FONT_DIR")
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "artifacts/fonts/windows"))
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../../artifacts/fonts/windows"))
        // Sibling office repo's font cache, when running from office-spreadsheet/.
        System.IO.Path.GetFullPath(System.IO.Path.Combine(cwd, "../office/artifacts/fonts/windows"))
    ]
    |> List.choose (fun dir ->
        if System.String.IsNullOrWhiteSpace dir then None
        elif System.IO.Directory.Exists dir then Some dir
        else None)

/// Map an XLSX font `name val="..."` to a normalized family key. Anything we
/// don't have a TTF for falls back to "calibri" — the renderer keeps working
/// even when we encounter Asian/script fonts we can't substitute properly.
let private familyKey (name: string option) =
    match name with
    | None -> "calibri"
    | Some n ->
        let lower = n.Trim().ToLowerInvariant()
        if lower.StartsWith "arial" then "arial"
        elif lower.StartsWith "calibri" then "calibri"
        elif lower.StartsWith "cambria" then "cambria"
        elif lower.StartsWith "times" then "times"
        else "calibri"

let private fontFiles family bold italic =
    match family with
    | "arial" ->
        if bold && italic then ["arialbi.ttf"; "arialbd.ttf"; "arial.ttf"]
        elif bold then ["arialbd.ttf"; "arial.ttf"]
        elif italic then ["ariali.ttf"; "arial.ttf"]
        else ["arial.ttf"]
    | "cambria" ->
        // Cambria ships only regular/bold/italic — no bold-italic TTF.
        if bold && italic then ["cambriab.ttf"; "cambriai.ttf"; "cambria.ttc"]
        elif bold then ["cambriab.ttf"; "cambria.ttc"]
        elif italic then ["cambriai.ttf"; "cambria.ttc"]
        else ["cambria.ttc"]
    | "times" ->
        if bold && italic then ["timesbi.ttf"; "timesbd.ttf"; "times.ttf"]
        elif bold then ["timesbd.ttf"; "times.ttf"]
        elif italic then ["timesi.ttf"; "times.ttf"]
        else ["times.ttf"]
    | _ ->
        if bold && italic then ["calibriz.ttf"; "calibrib.ttf"; "calibri.ttf"]
        elif bold then ["calibrib.ttf"; "calibri.ttf"]
        elif italic then ["calibrii.ttf"; "calibri.ttf"]
        else ["calibri.ttf"]

let private loadVariant family bold italic =
    let key = (family, bold, italic)
    match fontCache.TryGetValue key with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontFiles family bold italic
            |> List.tryPick (fun name ->
                fontDirs ()
                |> List.tryPick (fun dir ->
                    let path = System.IO.Path.Combine(dir, name)
                    if System.IO.File.Exists path then
                        try Some (Renderer.Text.loadFontFile path) with _ -> None
                    else None))
        fontCache.[key] <- loaded
        loaded

let private rendererFont (fontName: string option) (bold: bool) (italic: bool) =
    let family = familyKey fontName
    match loadVariant family bold italic with
    | Some f -> f
    | None ->
        // Family lookup failed — try Calibri before falling back to the
        // built-in default so we don't lose bold/italic for Arial/Times etc.
        loadVariant "calibri" bold italic
        |> Option.defaultWith defaultFont

let private parseHexColor (value: string) =
    let value = value.Trim()
    let parseByte (s: string) =
        match System.Byte.TryParse(s, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
        | true, value -> Some value
        | false, _ -> None
    let s = if value.StartsWith("#") then value.Substring(1) else value
    match s.Length with
    | 3 ->
        let expand (c: char) = string c + string c
        match parseByte (expand s.[0]), parseByte (expand s.[1]), parseByte (expand s.[2]) with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 6 ->
        match parseByte s.[0..1], parseByte s.[2..3], parseByte s.[4..5] with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    | 8 ->
        match parseByte s.[0..1], parseByte s.[2..3], parseByte s.[4..5], parseByte s.[6..7] with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | _ -> Color.black
    | _ -> Color.black

let private toRendererCommands (text: PaintText) =
    // Excel font sizes (XLSX <sz val="11"/>) are in points; the Renderer takes
    // pixel units. Convert pt→px so a Calibri 11-pt cell renders at Excel's
    // glyph height. Row/column heights are pre-converted via 96/72 upstream.
    let font = rendererFont text.FontName text.Bold text.Italic
    let pxSize = text.FontSize * 96.0 / 72.0
    let layout = Renderer.Text.layout font pxSize text.Text
    let measured = Renderer.Text.measureWidth font pxSize text.Text
    let alignedX =
        match text.Align with
        | AlignLeft -> text.Rect.X
        | AlignCenter -> text.Rect.X + (text.Rect.Width - measured) / 2.0
        | AlignRight -> text.Rect.X + text.Rect.Width - measured
    // Baseline from OS/2 winAscent/winDescent. These (not hhea/typo) are the
    // metrics Excel + GDI use for box layout: `winAscent` is em-top→baseline,
    // `winDescent` is baseline→em-bottom (positive), `winLineHeight` is the
    // single-line vertical reservation. hhea.ascender is smaller (cap+accent
    // height for typographic stacking) — using it would land the baseline ~3
    // px above where Excel renders. blockTopY positions the line block:
    //   top    → em-top of first line at cell.top
    //   middle → block centered in cell
    //   bottom → em-bottom of last line at cell.bottom (descender flush)
    // Excel's "General" (no explicit vertical align) is VAlignBottom.
    let ascent = Renderer.Text.winAscent font pxSize
    let lineHeight = Renderer.Text.winLineHeight font pxSize
    let blockHeight = lineHeight * float text.LineCount
    let blockTopY =
        match text.VerticalAlign with
        | VAlignTop -> text.Rect.Y
        | VAlignMiddle -> text.Rect.Y + (text.Rect.Height - blockHeight) / 2.0
        | VAlignBottom -> text.Rect.Y + text.Rect.Height - blockHeight
    let baseline = blockTopY + float text.LineIndex * lineHeight + ascent
    [ Cmd.DrawGlyphs(vec2 alignedX baseline, layout, Solid (parseHexColor text.Color)) ]

let private renderPaintItem = function
    | FillRect (rect, color) ->
        [ Cmd.FillRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color) ]
    | StrokeRect (rect, color, width) ->
        [ Cmd.StrokeRect(Renderer.Types.rect rect.X rect.Y rect.Width rect.Height, parseHexColor color, px width) ]
    | DrawText text -> toRendererCommands text
    | GridLine (x1, y1, x2, y2, color, width) ->
        let style = {
            Width = px width
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = []
            DashOffset = px 0.0
            MiterLimit = 4.0
        }
        [ Cmd.StrokePath([MoveTo (vec2 x1 y1); LineTo (vec2 x2 y2)], style) ]

/// Measure a text run in pixels using the real loaded font. Plugged into
/// Viewport.Measurer so Paint.wrapLines can break lines on actual glyph
/// width instead of mdwPx × charCount. Critical for non-Calibri-narrow
/// fonts where the approximation drifts (e.g. Times Roman fits more
/// characters per line than mdwPx assumes).
let private measureText
    (text: string) (fontPt: float)
    (fontName: string option) (bold: bool) (italic: bool) : float =
    let font = rendererFont fontName bold italic
    let pxSize = fontPt * 96.0 / 72.0
    Renderer.Text.measureWidth font pxSize text

let private viewportOf (model: SpreadsheetModel) width height : Viewport =
    // Match Excel's column/row header pixel sizes at 96 DPI / 100% zoom so
    // parity screenshots align cell-for-cell with our PNG renders. Measured
    // against artifacts/excel/pandas-test1-sheet1-headings.png: the column
    // header strip is 20 px tall (Calibri 11), the row header strip is 28 px
    // wide for two-digit row numbers, and each row in the default 15-pt
    // height renders at 21 source px.
    let headerHeight = 20.0
    let rowHeaderWidth = 28.0
    { ScrollX = 0.0
      ScrollY = 0.0
      Width = width
      Height = height
      HeaderHeight = headerHeight
      RowHeaderWidth = rowHeaderWidth
      Measurer = Some measureText }

let private cellCommands (model: SpreadsheetModel) width height =
    let viewport = viewportOf model width height
    buildPaintItems model viewport
    |> List.collect renderPaintItem

let widget width height (model: SpreadsheetModel) =
    canvas width height (fun w h -> cellCommands model w h)
