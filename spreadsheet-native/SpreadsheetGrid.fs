module Composit.Ui.SpreadsheetGrid

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private defaultFont () = Composit.Ui.Layout.font()

/// Cache of bold/italic font variants keyed by (bold, italic). Calibri's
/// bold/italic TTFs are the same set Word's native runner loads; falling back
/// to the default font keeps rendering working when the variant isn't found.
let private fontCache = System.Collections.Generic.Dictionary<bool * bool, Renderer.Text.FontData option>()

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

let private fontFiles bold italic =
    if bold && italic then ["calibriz.ttf"; "calibrib.ttf"; "calibri.ttf"]
    elif bold then ["calibrib.ttf"; "calibri.ttf"]
    elif italic then ["calibrii.ttf"; "calibri.ttf"]
    else ["calibri.ttf"]

let private loadVariant bold italic =
    let key = (bold, italic)
    match fontCache.TryGetValue key with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontFiles bold italic
            |> List.tryPick (fun name ->
                fontDirs ()
                |> List.tryPick (fun dir ->
                    let path = System.IO.Path.Combine(dir, name)
                    if System.IO.File.Exists path then
                        try Some (Renderer.Text.loadFontFile path) with _ -> None
                    else None))
        fontCache.[key] <- loaded
        loaded

let private rendererFont (bold: bool) (italic: bool) =
    loadVariant bold italic |> Option.defaultWith defaultFont

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
    let font = rendererFont text.Bold text.Italic
    let layout = Renderer.Text.layout font text.FontSize text.Text
    let measured = Renderer.Text.measureWidth font text.FontSize text.Text
    let alignedX =
        match text.Align with
        | AlignLeft -> text.Rect.X
        | AlignCenter -> text.Rect.X + (text.Rect.Width - measured) / 2.0
        | AlignRight -> text.Rect.X + text.Rect.Width - measured
    let baseline = text.Rect.Y + (text.Rect.Height + text.FontSize) / 2.0 - 2.0
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

let private viewportOf (model: SpreadsheetModel) width height : Viewport =
    let headerHeight = 22.0
    let rowHeaderWidth = 48.0
    { ScrollX = 0.0
      ScrollY = 0.0
      Width = width
      Height = height
      HeaderHeight = headerHeight
      RowHeaderWidth = rowHeaderWidth }

let private cellCommands (model: SpreadsheetModel) width height =
    let viewport = viewportOf model width height
    buildPaintItems model viewport
    |> List.collect renderPaintItem

let widget width height (model: SpreadsheetModel) =
    canvas width height (fun w h -> cellCommands model w h)
