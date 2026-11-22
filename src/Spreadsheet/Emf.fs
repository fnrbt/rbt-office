/// An EMF picture (Enhanced Metafile) as SVG, for the pictures Excel keeps of ActiveX controls (the picture an
/// ActiveX control is shown by until it is activated: xl/media/imageN.emf beside xl/activeX/activeXN.xml) and
/// other small metafiles. The records such a picture is made of are drawn - the mapping (window and viewport
/// origin / extent), solid brushes and pens, rectangles, ellipses, polygons and lines, fills by a brush
/// (`BITBLT` with no source), bitmaps (`BITBLT` / `STRETCHDIBITS` of a DIB, as an embedded BMP), and text
/// (`EXTTEXTOUTW` in the selected font, colour and alignment); clipping, ROPs other than copying, paths and
/// world transforms other than scaling and moving are not.
module Rbt.Spreadsheet.Emf

open System

let private u32 (d: byte[]) (o: int) = if o + 4 <= d.Length then int (BitConverter.ToUInt32(d, o)) else 0
let private i32 (d: byte[]) (o: int) = if o + 4 <= d.Length then BitConverter.ToInt32(d, o) else 0
let private i16 (d: byte[]) (o: int) = if o + 2 <= d.Length then int (BitConverter.ToInt16(d, o)) else 0
let private f32 (d: byte[]) (o: int) = if o + 4 <= d.Length then float (BitConverter.ToSingle(d, o)) else 0.0

let private colorOf (ref: int) = sprintf "#%02X%02X%02X" (ref &&& 0xFF) ((ref >>> 8) &&& 0xFF) ((ref >>> 16) &&& 0xFF)
let private num (v: float) = (Math.Round(v, 2)).ToString(Globalization.CultureInfo.InvariantCulture)
let private esc (text: string) = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

type private GdiObject =
    | Brush of string option
    | Pen of string option * float
    | Font of face: string * height: float * weight: int * italic: bool

type private Dc = {
    MapMode: int
    WinOrg: float * float
    WinExt: float * float
    VpOrg: float * float
    VpExt: float * float
    Brush: string option
    Pen: string option * float
    Font: string * float * int * bool
    TextColor: string
    TextAlign: int
    Transform: float * float * float * float * float * float
    Position: float * float
}

/// A DIB of a record (its BITMAPINFO and bits) as a BMP file.
let private bmpOf (d: byte[]) (start: int) (offBmi: int) (cbBmi: int) (offBits: int) (cbBits: int) : byte[] option =
    if cbBmi <= 0 || cbBits <= 0 || start + offBmi + cbBmi > d.Length || start + offBits + cbBits > d.Length then None
    else
        let size = 14 + cbBmi + cbBits
        let file = Array.zeroCreate<byte> size
        file.[0] <- byte 'B'; file.[1] <- byte 'M'
        let put (o: int) (v: int) = Array.blit (BitConverter.GetBytes v) 0 file o 4
        put 2 size
        put 10 (14 + cbBmi)
        Array.blit d (start + offBmi) file 14 cbBmi
        Array.blit d (start + offBits) file (14 + cbBmi) cbBits
        Some file

/// The SVG of an EMF picture drawn in `width` x `height` pixels; None for what is not an EMF.
let svg (data: byte[]) (width: float) (height: float) : string option =
    if data.Length < 88 || u32 data 0 <> 1 || u32 data 40 <> 0x464D4520 then None
    else
        let left, top, right, bottom = i32 data 8, i32 data 12, i32 data 16, i32 data 20
        let boundsW = float (max 1 (right - left + 1))
        let boundsH = float (max 1 (bottom - top + 1))
        let out = Text.StringBuilder()
        let objects = Collections.Generic.Dictionary<int, GdiObject>()
        let stack = Collections.Generic.Stack<Dc>()
        let mutable dc =
            { MapMode = 1; WinOrg = 0.0, 0.0; WinExt = 1.0, 1.0; VpOrg = 0.0, 0.0; VpExt = 1.0, 1.0
              Brush = Some "#FFFFFF"; Pen = Some "#000000", 1.0; Font = "Arial", 12.0, 400, false
              TextColor = "#000000"; TextAlign = 0; Transform = 1.0, 0.0, 0.0, 1.0, 0.0, 0.0; Position = 0.0, 0.0 }
        // Logical to device: the world transform, then the window to the viewport (MM_ISOTROPIC / ANISOTROPIC).
        let device (x: float) (y: float) =
            let m11, m12, m21, m22, dx, dy = dc.Transform
            let x, y = x * m11 + y * m21 + dx, x * m12 + y * m22 + dy
            if dc.MapMode = 7 || dc.MapMode = 8 then
                let wx, wy = dc.WinOrg
                let wex, wey = dc.WinExt
                let vx, vy = dc.VpOrg
                let vex, vey = dc.VpExt
                (x - wx) * (if wex = 0.0 then 1.0 else vex / wex) + vx, (y - wy) * (if wey = 0.0 then 1.0 else vey / wey) + vy
            else x, y
        let scaleY () =
            let _, wey = dc.WinExt
            let _, vey = dc.VpExt
            if (dc.MapMode = 7 || dc.MapMode = 8) && wey <> 0.0 then abs (vey / wey) else 1.0
        let fillAttr () = match dc.Brush with Some c -> sprintf "fill=\"%s\"" c | None -> "fill=\"none\""
        let strokeAttr () =
            match dc.Pen with
            | Some c, w -> sprintf "stroke=\"%s\" stroke-width=\"%s\"" c (num (max 1.0 (w * scaleY ())))
            | None, _ -> "stroke=\"none\""
        let points (o: int) (count: int) (short: bool) =
            [ for i in 0 .. count - 1 ->
                let x, y = if short then float (i16 data (o + 4 * i)), float (i16 data (o + 4 * i + 2)) else float (i32 data (o + 8 * i)), float (i32 data (o + 8 * i + 4))
                let px, py = device x y
                num px + "," + num py ]
            |> String.concat " "
        let rect (l: int) (t: int) (r: int) (b: int) =
            let x0, y0 = device (float l) (float t)
            let x1, y1 = device (float r) (float b)
            min x0 x1, min y0 y1, abs (x1 - x0), abs (y1 - y0)
        let image (bmp: byte[] option) (x: float) (y: float) (w: float) (h: float) =
            match bmp with
            | Some file ->
                out.AppendFormat("<image x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" preserveAspectRatio=\"none\" href=\"data:image/bmp;base64,{4}\"/>",
                                 num x, num y, num w, num h, Convert.ToBase64String file) |> ignore
            | None -> ()
        let mutable o = 0
        let mutable stop = false
        while not stop && o + 8 <= data.Length do
            let kind = u32 data o
            let size = u32 data (o + 4)
            if size < 8 || o + size > data.Length then stop <- true
            else
                let p = o + 8
                match kind with
                | 9 -> dc <- { dc with WinExt = float (i32 data p), float (i32 data (p + 4)) }
                | 10 -> dc <- { dc with WinOrg = float (i32 data p), float (i32 data (p + 4)) }
                | 11 -> dc <- { dc with VpExt = float (i32 data p), float (i32 data (p + 4)) }
                | 12 -> dc <- { dc with VpOrg = float (i32 data p), float (i32 data (p + 4)) }
                | 14 -> stop <- true
                | 17 -> dc <- { dc with MapMode = i32 data p }
                | 22 -> dc <- { dc with TextAlign = i32 data p }
                | 24 -> dc <- { dc with TextColor = colorOf (i32 data p) }
                | 27 -> dc <- { dc with Position = float (i32 data p), float (i32 data (p + 4)) }
                | 33 -> stack.Push dc
                | 34 ->
                    let back = -(i32 data p)
                    for _ in 1 .. back - 1 do (if stack.Count > 1 then stack.Pop() |> ignore)
                    if stack.Count > 0 then dc <- stack.Pop()
                | 35 -> dc <- { dc with Transform = f32 data p, f32 data (p + 4), f32 data (p + 8), f32 data (p + 12), f32 data (p + 16), f32 data (p + 20) }
                | 37 ->
                    let index = u32 data p
                    if index &&& 0x80000000 <> 0 then
                        // (Stock objects: white, light grey, grey, dark grey, black, no brush; white, black, no pen.)
                        match index &&& 0x7FFFFFFF with
                        | 0 -> dc <- { dc with Brush = Some "#FFFFFF" }
                        | 1 -> dc <- { dc with Brush = Some "#C0C0C0" }
                        | 2 -> dc <- { dc with Brush = Some "#808080" }
                        | 3 -> dc <- { dc with Brush = Some "#404040" }
                        | 4 -> dc <- { dc with Brush = Some "#000000" }
                        | 5 -> dc <- { dc with Brush = None }
                        | 6 -> dc <- { dc with Pen = Some "#FFFFFF", 1.0 }
                        | 7 -> dc <- { dc with Pen = Some "#000000", 1.0 }
                        | 8 -> dc <- { dc with Pen = None, 0.0 }
                        | _ -> ()
                    else
                        match objects.TryGetValue index with
                        | true, Brush color -> dc <- { dc with Brush = color }
                        | true, Pen (color, w) -> dc <- { dc with Pen = (color, w) }
                        | true, Font (face, h, weight, italic) -> dc <- { dc with Font = (face, h, weight, italic) }
                        | _ -> ()
                | 38 ->
                    // (CREATEPEN: style, width, colour.)
                    let style = u32 data (p + 4)
                    objects.[u32 data p] <- Pen ((if style &&& 0xF = 5 then None else Some (colorOf (i32 data (p + 16)))), float (i32 data (p + 8)))
                | 95 ->
                    // (EXTCREATEPEN: its LOGPEN-like part after the DIB offsets.)
                    let style = u32 data (p + 20)
                    objects.[u32 data p] <- Pen ((if style &&& 0xF = 5 then None else Some (colorOf (i32 data (p + 32)))), float (i32 data (p + 24)))
                | 39 ->
                    let style = u32 data (p + 4)
                    objects.[u32 data p] <- Brush (if style = 1 then None else Some (colorOf (i32 data (p + 8))))
                | 40 -> objects.Remove (u32 data p) |> ignore
                | 82 ->
                    let height = float (abs (i32 data (p + 4)))
                    let weight = i32 data (p + 20)
                    let italic = data.[p + 24] <> 0uy
                    let face =
                        let raw = Text.Encoding.Unicode.GetString(data, p + 32, min 64 (max 0 (data.Length - p - 32)))
                        match raw.IndexOf '\000' with -1 -> raw | at -> raw.Substring(0, at)
                    objects.[u32 data p] <- Font (face, height, weight, italic)
                | 42 | 43 ->
                    let x, y, w, h = rect (i32 data p) (i32 data (p + 4)) (i32 data (p + 8)) (i32 data (p + 12))
                    if kind = 43 then
                        out.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" {4} {5}/>", num x, num y, num w, num h, fillAttr (), strokeAttr ()) |> ignore
                    else
                        out.AppendFormat("<ellipse cx=\"{0}\" cy=\"{1}\" rx=\"{2}\" ry=\"{3}\" {4} {5}/>", num (x + w / 2.0), num (y + h / 2.0), num (w / 2.0), num (h / 2.0), fillAttr (), strokeAttr ()) |> ignore
                | 3 | 4 | 86 | 87 ->
                    let count = u32 data (p + 16)
                    let short = kind >= 86
                    let element = if kind = 3 || kind = 86 then "polygon" else "polyline"
                    let fill = if element = "polygon" then fillAttr () else "fill=\"none\""
                    out.AppendFormat("<{0} points=\"{1}\" {2} {3}/>", element, points (p + 20) count short, fill, strokeAttr ()) |> ignore
                | 8 | 91 ->
                    let polys = u32 data (p + 16)
                    let short = kind = 91
                    let mutable at = p + 24 + 4 * polys
                    let path = Text.StringBuilder()
                    for i in 0 .. polys - 1 do
                        let count = u32 data (p + 24 + 4 * i)
                        path.Append("M").Append(points at count short).Append("Z") |> ignore
                        at <- at + count * (if short then 4 else 8)
                    out.AppendFormat("<path d=\"{0}\" {1} {2}/>", path.ToString(), fillAttr (), strokeAttr ()) |> ignore
                | 54 ->
                    let x0, y0 = dc.Position
                    let x1, y1 = float (i32 data p), float (i32 data (p + 4))
                    let a, b = device x0 y0
                    let c, d = device x1 y1
                    out.AppendFormat("<line x1=\"{0}\" y1=\"{1}\" x2=\"{2}\" y2=\"{3}\" {4}/>", num a, num b, num c, num d, strokeAttr ()) |> ignore
                    dc <- { dc with Position = x1, y1 }
                | 76 ->
                    // BITBLT: destination x, y, cx, cy after the bounds; the raster operation; a source DIB or none.
                    let x, y, w, h = rect (i32 data (p + 16)) (i32 data (p + 20)) (i32 data (p + 16) + i32 data (p + 24)) (i32 data (p + 20) + i32 data (p + 28))
                    let cbBmi = u32 data (p + 76)
                    if cbBmi = 0 then
                        match dc.Brush with
                        | Some color -> out.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" fill=\"{4}\"/>", num x, num y, num w, num h, color) |> ignore
                        | None -> ()
                    else image (bmpOf data o (u32 data (p + 72)) cbBmi (u32 data (p + 80)) (u32 data (p + 84))) x y w h
                | 81 ->
                    // STRETCHDIBITS: destination x, y; the DIB's offsets; destination cx, cy at the end.
                    let xDest, yDest = i32 data (p + 16), i32 data (p + 20)
                    let cxDest, cyDest = i32 data (p + 64), i32 data (p + 68)
                    let x, y, w, h = rect xDest yDest (xDest + cxDest) (yDest + cyDest)
                    image (bmpOf data o (u32 data (p + 40)) (u32 data (p + 44)) (u32 data (p + 48)) (u32 data (p + 52))) x y w h
                | 84 ->
                    // EXTTEXTOUTW: the reference point, the characters (from the record's start).
                    let refX, refY = float (i32 data (p + 28)), float (i32 data (p + 32))
                    let count = u32 data (p + 36)
                    let offset = u32 data (p + 40)
                    if count > 0 && o + offset + 2 * count <= data.Length then
                        let text = Text.Encoding.Unicode.GetString(data, o + offset, 2 * count)
                        let x, y = device refX refY
                        let face, height, weight, italic = dc.Font
                        let px = height * scaleY ()
                        let anchor = match dc.TextAlign &&& 6 with 6 -> "middle" | 2 -> "end" | _ -> "start"
                        // (TA_BASELINE 24, TA_BOTTOM 8, else the top: the text hangs below its point.)
                        let baseline =
                            match dc.TextAlign &&& 24 with
                            | 24 -> "alphabetic"
                            | 8 -> "text-after-edge"
                            | _ -> "text-before-edge"
                        out.AppendFormat("<text x=\"{0}\" y=\"{1}\" font-family=\"'{2}', sans-serif\" font-size=\"{3}px\" fill=\"{4}\"{5}{6} text-anchor=\"{7}\" dominant-baseline=\"{8}\" xml:space=\"preserve\">{9}</text>",
                                         num x, num y, esc face, num px, dc.TextColor,
                                         (if weight >= 600 then " font-weight=\"bold\"" else ""), (if italic then " font-style=\"italic\"" else ""),
                                         anchor, baseline, esc text) |> ignore
                | _ -> ()
                o <- o + size
        Some (sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%s\" height=\"%s\" viewBox=\"%d %d %s %s\" preserveAspectRatio=\"none\" style=\"display:block;position:absolute;left:0;top:0\">%s</svg>"
                    (num width) (num height) left top (num boundsW) (num boundsH) (out.ToString()))
