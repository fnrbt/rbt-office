module Rbt.WordNative.Render

open System
open System.Collections.Generic
open Rbt.Renderer
open Rbt.Word.Document
module WordLayout = Rbt.Word.Layout
module WordPaint = Rbt.Word.Paint
module Fonts = Rbt.OfficeNative.Fonts

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v
let private pointsToPixels points = points * 96.0 / 72.0
let layoutSettings (doc: Document) = WordLayout.documentSettings 16.0 24.0 3.5 doc
let scale width (doc: Document) = max 0.001 width / max 1.0 doc.PageSetup.PageWidth
let measureText (text: string) (style: Rbt.Word.TextBuffer.TextStyle) =
    Fonts.measure text (pointsToPixels style.FontSize) style.FontFamily style.Bold style.Italic + Rbt.Word.TextBuffer.characterSpacingAdvance text style
let measureTextWithFont font (text: string) (style: Rbt.Word.TextBuffer.TextStyle) =
    Rbt.Renderer.Text.measureWidth font (pointsToPixels style.FontSize) text + Rbt.Word.TextBuffer.characterSpacingAdvance text style
let private fontForPaint (style: WordPaint.TextPaintStyle) = Fonts.resolve style.FontFamily style.Bold style.Italic

let private parseHexColor (value: string) =
    let value =
        match value.Trim().ToLowerInvariant() with
        | "black" -> "#000000"
        | "blue" -> "#0000FF"
        | "cyan" -> "#00FFFF"
        | "green" -> "#00FF00"
        | "magenta" -> "#FF00FF"
        | "red" -> "#FF0000"
        | "yellow" -> "#FFFF00"
        | "white" -> "#FFFFFF"
        | "darkblue" -> "#000080"
        | "darkcyan" -> "#008080"
        | "darkgreen" -> "#008000"
        | "darkmagenta" -> "#800080"
        | "darkred" -> "#800000"
        | "darkyellow" -> "#808000"
        | "darkgray" -> "#808080"
        | "lightgray" -> "#C0C0C0"
        | _ -> value
    let parseByte (s: string) =
        match System.Byte.TryParse(s, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
        | true, value -> Some value
        | false, _ -> None
    let colorFromParts r g b =
        match parseByte r, parseByte g, parseByte b with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    let colorFromPartsWithAlpha r g b a =
        match parseByte r, parseByte g, parseByte b, parseByte a with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | _ -> Color.black
    let s = if value.StartsWith("#") then value.Substring(1) else value

    match s.Length with
    | 3 ->
        let expand (c: char) = string c + string c
        colorFromParts (expand s.[0]) (expand s.[1]) (expand s.[2])
    | 6 ->
        colorFromParts s.[0..1] s.[2..3] s.[4..5]
    | 8 ->
        colorFromPartsWithAlpha s.[0..1] s.[2..3] s.[4..5] s.[6..7]
    | _ when value.StartsWith("rgba(") ->
        let parts =
            value.Substring(5, value.Length - 6).Split(',')
            |> Array.map (fun p -> p.Trim())

        if parts.Length = 4 then
            let alpha = float parts.[3]
            Color.rgba (byte parts.[0]) (byte parts.[1]) (byte parts.[2]) (byte (max 0.0 (min 255.0 (alpha * 255.0))))
        else
            Color.black
    | _ -> Color.black

let rec toRendererCommandWith (fontFor: WordPaint.TextPaintStyle -> Rbt.Renderer.Text.FontData) (resolveImage: string -> Framebuffer option) = function
    | WordPaint.DrawText (text, style, x, baseline, _) ->
        let font = fontFor style
        let size = pointsToPixels style.FontSize
        let mutable origin = x
        [ for face, run in Fonts.segmentByFontCoverage font style.Bold text do
            yield Cmd.DrawGlyphs(vec2 origin baseline, Rbt.Renderer.Text.layout face size run, Solid (parseHexColor style.Color))
            origin <- origin + Rbt.Renderer.Text.measureWidth face size run ]

    | WordPaint.DrawUnderline (color, _lineStyle, x, y, width) ->
        [ Cmd.Line(vec2 x y, vec2 (x + width) y, parseHexColor color, px 1.0) ]

    | WordPaint.DrawImage (src, x, y, width, height, cropLeft, cropTop, cropRight, cropBottom) ->
        match resolveImage src with
        | Some fb ->
            let crop value =
                match value with
                | Some text ->
                    match Double.TryParse((text: string), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
                    | true, value -> max 0.0 (min 100000.0 value)
                    | _ -> 0.0
                | None -> 0.0
            let left, top, right, bottom = crop cropLeft, crop cropTop, crop cropRight, crop cropBottom
            let source =
                if left + right >= 100000.0 || top + bottom >= 100000.0 then
                    rect 0.0 0.0 (float fb.W) (float fb.H)
                else
                    rect (float fb.W * left / 100000.0) (float fb.H * top / 100000.0)
                        (float fb.W * (100000.0 - left - right) / 100000.0)
                        (float fb.H * (100000.0 - top - bottom) / 100000.0)
            [ Cmd.DrawImage(fb, source, rect x y width height, true) ]
        | None ->
            // The browser path can draw data URLs directly. Widgets need a decoded framebuffer.
            [ Cmd.StrokeRect(rect x y width height, Color.rgb 180uy 180uy 180uy, px 1.0) ]

    | WordPaint.FillRect (color, x, y, width, height) ->
        [ Cmd.FillRect(rect x y width height, parseHexColor color) ]

    | WordPaint.StrokeRect (color, lineWidth, x, y, width, height) ->
        [ Cmd.StrokeRect(rect x y width height, parseHexColor color, px lineWidth) ]

    | WordPaint.StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        let dashArray =
            match dashStyle with
            | 1 -> [px lineWidth; px lineWidth]
            | 2 -> [px (lineWidth * 3.0); px (lineWidth * 1.5)]
            | 3 -> [px (lineWidth * 3.0); px (lineWidth * 1.5); px lineWidth; px (lineWidth * 1.5)]
            | _ -> []
        let style = {
            Width = px lineWidth
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = dashArray
            DashOffset = px dashOffset
            MiterLimit = 4.0
        }
        [ Cmd.StrokePath([MoveTo (vec2 x1 y1); LineTo (vec2 x2 y2)], style) ]

    | WordPaint.DrawShape (fill, stroke, strokeWidth, (bx, by, bw, bh), segs) ->
        let rpath =
            segs |> List.map (function
                | PMove (x, y) -> MoveTo (vec2 x y)
                | PLine (x, y) -> LineTo (vec2 x y)
                | PQuad (cx, cy, x, y) -> QuadTo (vec2 cx cy, vec2 x y)
                | PCubic (a, b, c, d, x, y) -> CubicTo (vec2 a b, vec2 c d, vec2 x y)
                | PClose -> Close)
        let fillCmds =
            match fill with
            | WordPaint.FillNone -> []
            | WordPaint.FillSolid c -> [ Cmd.FillPath(rpath, Solid (parseHexColor c), NonZero) ]
            | WordPaint.FillGradient g ->
                let stops = g.Stops |> List.map (fun s -> { Offset = s.Offset; StopColor = parseHexColor s.Color })
                let cx, cy = bx + bw / 2.0, by + bh / 2.0
                let paint =
                    if g.Radial then RadialGradient(vec2 cx cy, px (max bw bh / 2.0), stops, Pad)
                    else
                        let rad = g.Angle * System.Math.PI / 180.0
                        let dx, dy = cos rad * bw / 2.0, sin rad * bh / 2.0
                        LinearGradient(vec2 (cx - dx) (cy - dy), vec2 (cx + dx) (cy + dy), stops, Pad)
                [ Cmd.FillPath(rpath, paint, NonZero) ]
        let strokeCmds =
            match stroke with
            | Some c ->
                let style = { Width = px strokeWidth; Cap = Butt; Join = Miter; Paint = Solid (parseHexColor c); DashArray = []; DashOffset = px 0.0; MiterLimit = 4.0 }
                [ Cmd.StrokePath(rpath, style) ]
            | None -> []
        fillCmds @ strokeCmds

    | WordPaint.DrawGroup (clip, rotation, centerX, centerY, opacity, _, inner) ->
        // A text box or shape with its text: clipped to its box, turned about its centre, translucent.
        let innerCmds = inner |> List.collect (toRendererCommandWith fontFor resolveImage)
        let clipped =
            match clip with
            | Some (x, y, width, height) ->
                [ Cmd.Clip([ MoveTo (vec2 x y); LineTo (vec2 (x + width) y); LineTo (vec2 (x + width) (y + height)); LineTo (vec2 x (y + height)); Close ], NonZero, innerCmds) ]
            | None -> innerCmds
        if abs rotation < 0.01 && opacity >= 0.999 then clipped
        else
            let angle = rotation * System.Math.PI / 180.0
            let c, s = cos angle, sin angle
            let xform : Matrix3x2 =
                { M11 = c; M12 = s; M21 = -s; M22 = c
                  DX = centerX - centerX * c + centerY * s; DY = centerY - centerX * s - centerY * c }
            [ Cmd.Group(xform, opacity, clipped) ]
    | WordPaint.DrawMirrored (flipH, flipV, centerX, centerY, _, inner) ->
        // A flipped picture: mirrored about its middle.
        let innerCmds = inner |> List.collect (toRendererCommandWith fontFor resolveImage)
        let sx, sy = (if flipH then -1.0 else 1.0), (if flipV then -1.0 else 1.0)
        let xform : Matrix3x2 = { M11 = sx; M12 = 0.0; M21 = 0.0; M22 = sy; DX = centerX - sx * centerX; DY = centerY - sy * centerY }
        [ Cmd.Group(xform, 1.0, innerCmds) ]

let toRendererCommand resolveImage command = toRendererCommandWith fontForPaint resolveImage command

/// Decode embedded data URLs once; external paths and URLs are not fetched.
let createImageResolver () =
    let cache = Dictionary<string, Framebuffer option>()
    fun (src: string) ->
        match cache.TryGetValue src with
        | true, result -> result
        | _ ->
            let result =
                try
                    let comma = src.IndexOf(',')
                    if not (src.StartsWith("data:image/", StringComparison.OrdinalIgnoreCase)) || comma < 0 || not (src.Substring(0, comma).EndsWith(";base64", StringComparison.OrdinalIgnoreCase)) then None
                    else
                        let bytes = Convert.FromBase64String(src.Substring(comma + 1))
                        Some (Rbt.OfficeNative.Images.decode bytes)
                with _ -> None
            cache.[src] <- result
            result

let viewTransform width doc commands =
    let s = scale width doc
    if abs (s - 1.0) < 0.000001 then commands
    else [ Cmd.Group({ M11 = s; M12 = 0.0; M21 = 0.0; M22 = s; DX = 0.0; DY = 0.0 }, 1.0, commands) ]
