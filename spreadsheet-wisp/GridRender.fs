module Rbt.SpreadsheetWisp.GridRender

open System
open System.Collections.Generic
open System.IO
open Rbt.Renderer

module M = Rbt.Spreadsheet.Model
module L = Rbt.Spreadsheet.Layout
module P = Rbt.Spreadsheet.Paint

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v
let private pointsToPixels points = points * 96.0 / 72.0

let private parseHexColor (value: string) =
    let value =
        match value.Trim().ToLowerInvariant() with
        | "" -> "#000000"
        | "black" -> "#000000"
        | "white" -> "#ffffff"
        | "red" -> "#ff0000"
        | "blue" -> "#0000ff"
        | "green" -> "#008000"
        | "darkgray" | "darkgrey" -> "#808080"
        | "lightgray" | "lightgrey" -> "#c0c0c0"
        | other -> other
    let parseByte (s: string) =
        match Byte.TryParse(s, Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture) with
        | true, v -> Some v
        | _ -> None
    let fromParts r g b a =
        match parseByte r, parseByte g, parseByte b, a |> Option.bind parseByte with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | Some r, Some g, Some b, None -> Color.rgb r g b
        | _ -> Color.black
    let s = if value.StartsWith "#" then value.Substring 1 else value
    match s.Length with
    | 3 ->
        let e (c: char) = string c + string c
        fromParts (e s.[0]) (e s.[1]) (e s.[2]) None
    | 6 -> fromParts s.[0..1] s.[2..3] s.[4..5] None
    | 8 -> fromParts s.[0..1] s.[2..3] s.[4..5] (Some s.[6..7])
    | _ -> Color.black

let mutable private defaultFont : Rbt.Renderer.Text.FontData option = None
let private rendererFont () =
    match defaultFont with
    | Some f -> f
    | None ->
        let f = Rbt.Renderer.Text.loadFontFile "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
        defaultFont <- Some f
        f

let private fontCache = Dictionary<string, Rbt.Renderer.Text.FontData option>()
let private existingDir path = if Directory.Exists path then Some path else None
let private fontDirs () =
    [ Environment.GetEnvironmentVariable "RB_WORD_FONT_DIR"
      Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "artifacts/fonts/windows"))
      "/usr/share/fonts/truetype/liberation" ]
    |> List.choose (fun p -> if String.IsNullOrWhiteSpace p then None else existingDir p)

let private fontFilesFor (family: string option) (bold: bool) (italic: bool) =
    match family |> Option.map (fun v -> v.Trim().ToLowerInvariant()) with
    | Some "calibri" ->
        if bold then [ "calibrib.ttf"; "LiberationSans-Bold.ttf" ]
        elif italic then [ "calibrii.ttf"; "LiberationSans-Italic.ttf" ]
        else [ "calibri.ttf"; "LiberationSans-Regular.ttf" ]
    | Some "arial" | Some "helvetica" | Some "segoe ui" ->
        if bold then [ "arialbd.ttf"; "LiberationSans-Bold.ttf" ]
        elif italic then [ "ariali.ttf"; "LiberationSans-Italic.ttf" ]
        else [ "arial.ttf"; "LiberationSans-Regular.ttf" ]
    | Some "times new roman" | Some "cambria" | Some "georgia" ->
        if bold then [ "timesbd.ttf"; "LiberationSerif-Bold.ttf" ]
        elif italic then [ "timesi.ttf"; "LiberationSerif-Italic.ttf" ]
        else [ "times.ttf"; "LiberationSerif-Regular.ttf" ]
    | Some "consolas" | Some "courier new" ->
        if bold then [ "consolab.ttf"; "LiberationMono-Bold.ttf" ]
        elif italic then [ "consolai.ttf"; "LiberationMono-Italic.ttf" ]
        else [ "consola.ttf"; "LiberationMono-Regular.ttf" ]
    | _ -> []

let private tryLoadFont (family: string option) (bold: bool) (italic: bool) =
    let key = sprintf "%A|%b|%b" family bold italic
    match fontCache.TryGetValue key with
    | true, cached -> cached
    | _ ->
        let loaded =
            fontFilesFor family bold italic
            |> List.tryPick (fun fileName ->
                fontDirs () |> List.tryPick (fun dir ->
                    let path = Path.Combine(dir, fileName)
                    if File.Exists path then
                        try Some(Rbt.Renderer.Text.loadFontFile path) with _ -> None
                    else None))
        fontCache.[key] <- loaded
        loaded

let fontFor (family: string option) bold italic =
    tryLoadFont family bold italic |> Option.defaultWith rendererFont

let measureText (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool) =
    Rbt.Renderer.Text.measureWidth (fontFor fontName bold italic) (pointsToPixels fontPt) text

let viewport (scrollX: float) (scrollY: float) (width: float) (height: float) : L.Viewport =
    { L.ScrollX = scrollX
      ScrollY = scrollY
      Width = width
      Height = height
      HeaderHeight = 20.0
      RowHeaderWidth = 28.0
      Measurer = Some measureText }

let private strokeStyle color width =
    { Width = px width
      Cap = Butt
      Join = Miter
      Paint = Solid(parseHexColor color)
      DashArray = []
      DashOffset = 0.0<px>
      MiterLimit = 4.0 }

let private rpath x y w h : Rbt.Renderer.Types.PathSeg list =
    [ MoveTo(vec2 x y)
      LineTo(vec2 (x + w) y)
      LineTo(vec2 (x + w) (y + h))
      LineTo(vec2 x (y + h))
      Close ]

let private pathSegs (segs: P.VecSeg list) =
    segs
    |> List.map (function
        | P.VMove (x, y) -> MoveTo(vec2 x y)
        | P.VLine (x, y) -> LineTo(vec2 x y)
        | P.VCubic (x1, y1, x2, y2, x, y) -> CubicTo(vec2 x1 y1, vec2 x2 y2, vec2 x y)
        | P.VClose -> Close)

let private placeholder (r: L.Rect) =
    [ Cmd.FillRect(rect r.X r.Y (max 6.0 r.Width) (max 6.0 r.Height), Color.rgba 255uy 0uy 255uy 96uy)
      Cmd.Line(vec2 r.X r.Y, vec2 (r.X + max 6.0 r.Width) (r.Y + max 6.0 r.Height), Color.rgb 160uy 0uy 160uy, px 1.0)
      Cmd.Line(vec2 (r.X + max 6.0 r.Width) r.Y, vec2 r.X (r.Y + max 6.0 r.Height), Color.rgb 160uy 0uy 160uy, px 1.0) ]

let private baseline (font: Rbt.Renderer.Text.FontData) fontPx (text: P.PaintText) =
    let lineHeight = max (fontPx * 1.25) (fontPx + 2.0)
    let blockHeight = lineHeight * float text.LineCount
    let top =
        match text.VerticalAlign with
        | P.VAlignTop -> text.Rect.Y + 2.0
        | P.VAlignMiddle -> text.Rect.Y + (text.Rect.Height - blockHeight) / 2.0
        | P.VAlignBottom -> text.Rect.Y + text.Rect.Height - blockHeight - 2.0
    top + float text.LineIndex * lineHeight + Rbt.Renderer.Text.ascender font fontPx + text.BaselineShift * 96.0 / 72.0

let private alignedX font fontPx (text: P.PaintText) =
    let width = Rbt.Renderer.Text.measureWidth font fontPx text.Text
    match text.Align with
    | P.AlignLeft -> text.Rect.X
    | P.AlignCenter -> text.Rect.X + max 0.0 ((text.Rect.Width - width) / 2.0)
    | P.AlignRight -> text.Rect.X + max 0.0 (text.Rect.Width - width)

let private textCommands (text: P.PaintText) =
    if String.IsNullOrEmpty text.Text || text.Rect.Width <= 0.0 || text.Rect.Height <= 0.0 then []
    else
        let font = fontFor text.FontName text.Bold text.Italic
        let fontPx = pointsToPixels text.FontSize
        let x = alignedX font fontPx text
        let y = baseline font fontPx text
        let glyphs = Cmd.DrawGlyphs(vec2 x y, Rbt.Renderer.Text.layout font fontPx text.Text, Solid(parseHexColor text.Color))
        let width = Rbt.Renderer.Text.measureWidth font fontPx text.Text
        let decoY = y + 1.5
        let strikeY = y - fontPx * 0.32
        let deco =
            [ if text.Underline then
                  yield Cmd.Line(vec2 x decoY, vec2 (x + width) decoY, parseHexColor text.Color, px 1.0)
                  if text.UnderlineDouble then
                      yield Cmd.Line(vec2 x (decoY + 2.0), vec2 (x + width) (decoY + 2.0), parseHexColor text.Color, px 1.0)
              if text.Strikethrough then
                  yield Cmd.Line(vec2 x strikeY, vec2 (x + width) strikeY, parseHexColor text.Color, px 1.0) ]
        let cmds = glyphs :: deco
        match text.Rotation with
        | 0 -> cmds
        | 255 ->
            // Stacked text is rare; draw a conspicuous but readable fallback instead of dropping it.
            placeholder text.Rect @ cmds
        | r ->
            let degrees = if r <= 90 then float -r else float (r - 90)
            let m =
                Matrix3x2.multiply
                    (Matrix3x2.translate x y)
                    (Matrix3x2.multiply (Matrix3x2.rotate (degrees * Math.PI / 180.0)) (Matrix3x2.translate -x -y))
            [ Cmd.Group(m, 1.0, cmds) ]

/// The decoded pixels of a picture, kept with its data (a picture is drawn at every paint).
let private decodedPictures = System.Runtime.CompilerServices.ConditionalWeakTable<byte[], Framebuffer option ref>()

let private pictureFramebuffer (data: byte[]) : Framebuffer option =
    if isNull data then None
    else
        let entry =
            decodedPictures.GetValue(data, fun bytes ->
                ref (Rbt.Spreadsheet.Png.tryDecode bytes
                     |> Option.map (fun image -> Framebuffer.ofArray (Rbt.Spreadsheet.Png.toPremultipliedBgra image) image.Width image.Height)))
        entry.Value

let rec private itemCommands (item: P.PaintItem) =
    match item with
    | P.FillRect (r, color) ->
        [ Cmd.FillRect(rect r.X r.Y r.Width r.Height, parseHexColor color) ]
    | P.StrokeRect (r, color, width) ->
        [ Cmd.StrokeRect(rect r.X r.Y r.Width r.Height, parseHexColor color, px width) ]
    | P.GridLine (x1, y1, x2, y2, color, width) ->
        [ Cmd.Line(vec2 x1 y1, vec2 x2 y2, parseHexColor color, px width) ]
    | P.DrawText text ->
        // Text wider than its cell is cut off at the cells it may run over.
        match text.Clip, textCommands text with
        | Some clip, (_ :: _ as cmds) when clip.Width > 0.0 && clip.Height > 0.0 ->
            [ Cmd.Clip(rpath clip.X clip.Y clip.Width clip.Height, NonZero, cmds) ]
        | _, cmds -> cmds
    | P.VecPath (segs, fill, stroke) ->
        let path = pathSegs segs
        [ match fill with
          | Some color -> yield Cmd.FillPath(path, Solid(parseHexColor color), NonZero)
          | None -> ()
          match stroke with
          | Some (color, width) -> yield Cmd.StrokePath(path, strokeStyle color width)
          | None -> ()
          if fill.IsNone && stroke.IsNone then
              yield! placeholder { X = 0.0; Y = 0.0; Width = 24.0; Height = 24.0 } ]
    | P.VecGroup ((a, b, c, d, e, f), children) ->
        [ Cmd.Group({ M11 = a; M12 = b; M21 = c; M22 = d; DX = e; DY = f }, 1.0, children |> List.collect itemCommands) ]
    // A picture is drawn from its pixels (PNG); a labelled frame stands for a
    // picture in a format that is not decoded here, as it does for a chart.
    | P.DrawImage (r, image) ->
        match pictureFramebuffer image.Data with
        | Some fb when r.Width > 0.0 && r.Height > 0.0 ->
            [ Cmd.DrawImage(fb, rect 0.0 0.0 (float fb.W) (float fb.H), rect r.X r.Y r.Width r.Height, true) ]
        | _ -> P.drawingFrameItems r image.Label |> List.collect itemCommands

let private clipCommands (clip: L.Rect) (cmds: Cmd list) =
    if clip.Width <= 0.0 || clip.Height <= 0.0 then []
    else [ Cmd.Clip(rpath clip.X clip.Y clip.Width clip.Height, NonZero, cmds) ]

let commands (model: M.SpreadsheetModel) (viewport: L.Viewport) : Cmd list =
    let items = P.buildViewportPaintItems model viewport
    let grouped =
        items
        |> List.groupBy (fun i -> i.Zone)
        |> Map.ofList
    let order =
        [ L.CornerZone; L.BodyZone; L.FrozenRowsZone; L.FrozenColsZone; L.FrozenCornerZone; L.ColumnHeaderZone; L.RowHeaderZone ]
    [
        for zone in order do
            match Map.tryFind zone grouped with
            | Some zoneItems ->
                let cmds = zoneItems |> List.collect (fun i -> itemCommands i.Item)
                yield! clipCommands (L.viewportClipFor model viewport zone) cmds
            | None -> ()
    ]

let render (model: M.SpreadsheetModel) (scrollX: float) (scrollY: float) (width: float) (height: float) =
    commands model (viewport scrollX scrollY width height)
