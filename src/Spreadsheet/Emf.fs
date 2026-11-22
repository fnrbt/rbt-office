/// An EMF picture (Enhanced Metafile) as SVG, for the pictures Excel keeps of ActiveX controls (the picture an
/// ActiveX control is shown by until it is activated: xl/media/imageN.emf beside xl/activeX/activeXN.xml) and
/// other small metafiles. The records such a picture is made of are drawn - the mapping (window and viewport
/// origin / extent), solid brushes and pens, rectangles, ellipses, polygons and lines, fills by a brush
/// (`BITBLT` with no source: a solid brush, or a monochrome pattern brush - `CREATEMONOBRUSH` - in the text and
/// background colours, as Excel 2019 draws a scroll bar's track and a pressed toggle button), bitmaps (`BITBLT` /
/// `STRETCHDIBITS` of a DIB: copied, or a mask `SRCAND` then the picture `SRCPAINT` - an option button's circle -
/// drawn where the mask lets it through), text (`EXTTEXTOUTW` in the selected font, colour and alignment, each
/// character at its own advance, the rectangle filled with the background colour first when the record says so
/// and the text cut at it), clipping (`INTERSECTCLIPRECT`, `EXTSELECTCLIPRGN` of rectangles, saved and restored
/// with the DC) and paths (`BEGINPATH` .. `ENDPATH` of moves, lines, polylines, Béziers and rectangles, then
/// filled, stroked or both, or made the clip). Other ROPs and world transforms other than scaling and moving are
/// not drawn. synthetic-activex.xlsx (build-activex.vbs) against Excel 2019's pictures of the controls.
module Rbt.Spreadsheet.Emf

open System

let private u32 (d: byte[]) (o: int) = if o + 4 <= d.Length then int (BitConverter.ToUInt32(d, o)) else 0
let private i32 (d: byte[]) (o: int) = if o + 4 <= d.Length then BitConverter.ToInt32(d, o) else 0
let private i16 (d: byte[]) (o: int) = if o + 2 <= d.Length then int (BitConverter.ToInt16(d, o)) else 0
let private u16 (d: byte[]) (o: int) = if o + 2 <= d.Length then int (BitConverter.ToUInt16(d, o)) else 0
let private f32 (d: byte[]) (o: int) = if o + 4 <= d.Length then float (BitConverter.ToSingle(d, o)) else 0.0

let private colorOf (ref: int) = sprintf "#%02X%02X%02X" (ref &&& 0xFF) ((ref >>> 8) &&& 0xFF) ((ref >>> 16) &&& 0xFF)
let private num (v: float) = (Math.Round(v, 2)).ToString(Globalization.CultureInfo.InvariantCulture)
let private esc (text: string) = text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

/// A monochrome pattern: its width, height and bits (true: a 1, drawn in the background colour), top row first.
type private Pattern = int * int * bool[][]

type private GdiObject =
    | Brush of string option
    | PatternBrush of Pattern
    | Pen of string option * float
    | Font of face: string * height: float * weight: int * italic: bool

type private Dc = {
    MapMode: int
    WinOrg: float * float
    WinExt: float * float
    VpOrg: float * float
    VpExt: float * float
    Brush: string option
    Pattern: Pattern option
    Pen: string option * float
    Font: string * float * int * bool
    TextColor: string
    BkColor: string
    TextAlign: int
    PolyFill: int
    Transform: float * float * float * float * float * float
    Position: float * float
    /// The clip: device rectangles (x, y, w, h) whose union is drawn in, or an SVG path; None for none.
    Clip: string option
}

/// The SVG pictures of a page: their patterns' and clips' ids are numbered on from one picture to the next.
let mutable private nextId = 0

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

/// The pixels of a DIB of a record, top row first: Some colour ("#RRGGBB"), for 1-, 24- and 32-bit DIBs (a 1-bit DIB's
/// pixels by its two palette entries); None for other DIBs.
let private pixelsOf (d: byte[]) (start: int) (offBmi: int) (offBits: int) : (int * int * string[][]) option =
    let bmi = start + offBmi
    let w, h, bpp = i32 d (bmi + 4), i32 d (bmi + 8), u16 d (bmi + 14)
    let rows = abs h
    if w <= 0 || rows = 0 || w > 512 || rows > 512 then None
    else
        let stride = ((w * bpp + 31) / 32) * 4
        let bits = start + offBits
        let palette (i: int) = let o = bmi + u32 d bmi + 4 * i in sprintf "#%02X%02X%02X" d.[o + 2] d.[o + 1] d.[o]
        let row (y: int) = bits + stride * (if h > 0 then rows - 1 - y else y)
        if bits + stride * rows > d.Length then None
        else
            match bpp with
            | 24 | 32 ->
                let step = bpp / 8
                Some (w, rows, Array.init rows (fun y -> let r = row y in Array.init w (fun x -> let o = r + step * x in sprintf "#%02X%02X%02X" d.[o + 2] d.[o + 1] d.[o])))
            | 1 ->
                let zero, one = palette 0, palette 1
                Some (w, rows, Array.init rows (fun y -> let r = row y in Array.init w (fun x -> if (int d.[r + x / 8] >>> (7 - x % 8)) &&& 1 = 1 then one else zero)))
            | _ -> None

/// The SVG of an EMF picture drawn in `width` x `height` pixels; None for what is not an EMF.
let svg (data: byte[]) (width: float) (height: float) : string option =
    if data.Length < 88 || u32 data 0 <> 1 || u32 data 40 <> 0x464D4520 then None
    else
        let left, top, right, bottom = i32 data 8, i32 data 12, i32 data 16, i32 data 20
        let boundsW = float (max 1 (right - left + 1))
        let boundsH = float (max 1 (bottom - top + 1))
        let out = Text.StringBuilder()
        let defs = Text.StringBuilder()
        nextId <- nextId + 1
        let prefix = sprintf "emf%d_" nextId
        let mutable ids = 0
        let newId () = ids <- ids + 1; prefix + string ids
        let clips = Collections.Generic.Dictionary<string, string>()
        let objects = Collections.Generic.Dictionary<int, GdiObject>()
        let stack = Collections.Generic.Stack<Dc>()
        let mutable dc =
            { MapMode = 1; WinOrg = 0.0, 0.0; WinExt = 1.0, 1.0; VpOrg = 0.0, 0.0; VpExt = 1.0, 1.0
              Brush = Some "#FFFFFF"; Pattern = None; Pen = Some "#000000", 1.0; Font = "Arial", 12.0, 400, false
              TextColor = "#000000"; BkColor = "#FFFFFF"; TextAlign = 0; PolyFill = 1
              Transform = 1.0, 0.0, 0.0, 1.0, 0.0, 0.0; Position = 0.0, 0.0; Clip = None }
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
        let scaleX () =
            let wex, _ = dc.WinExt
            let vex, _ = dc.VpExt
            if (dc.MapMode = 7 || dc.MapMode = 8) && wex <> 0.0 then abs (vex / wex) else 1.0
        let scaleY () =
            let _, wey = dc.WinExt
            let _, vey = dc.VpExt
            if (dc.MapMode = 7 || dc.MapMode = 8) && wey <> 0.0 then abs (vey / wey) else 1.0
        // (An element in the clip of the DC: one <clipPath> for each clip drawn in.)
        let emit (element: string) =
            match dc.Clip with
            | None -> out.Append element |> ignore
            | Some path ->
                let id =
                    match clips.TryGetValue path with
                    | true, id -> id
                    | _ ->
                        let id = newId ()
                        defs.AppendFormat("<clipPath id=\"{0}\" clipPathUnits=\"userSpaceOnUse\"><path d=\"{1}\"/></clipPath>", id, path) |> ignore
                        clips.[path] <- id
                        id
                out.AppendFormat("<g clip-path=\"url(#{0})\">{1}</g>", id, element) |> ignore
        let rectPath (x: float) (y: float) (w: float) (h: float) = sprintf "M%s %sh%sv%sh%sZ" (num x) (num y) (num w) (num h) (num -w)
        // (The clip's rectangles kept as they are written: an intersection with a rectangle is worked out on
        // a clip of one rectangle, a clip of several is cut by the rectangle's own clip around it.)
        let intersect (x: float) (y: float) (w: float) (h: float) =
            match dc.Clip with
            | None -> dc <- { dc with Clip = Some (rectPath x y w h) }
            | Some path ->
                let m = Text.RegularExpressions.Regex.Match(path, "^M(-?[\\d.]+) (-?[\\d.]+)h(-?[\\d.]+)v(-?[\\d.]+)h-?[\\d.]+Z$")
                if m.Success then
                    let v (i: int) = Double.Parse(m.Groups.[i].Value, Globalization.CultureInfo.InvariantCulture)
                    let x0, y0 = max x (v 1), max y (v 2)
                    let x1, y1 = min (x + w) (v 1 + v 3), min (y + h) (v 2 + v 4)
                    dc <- { dc with Clip = Some (rectPath x0 y0 (max 0.0 (x1 - x0)) (max 0.0 (y1 - y0))) }
                else dc <- { dc with Clip = Some (rectPath x y w h) }
        let paint (fill: string option) =
            (match fill with Some c -> sprintf "fill=\"%s\"" c | None -> "fill=\"none\"")
        let fillAttr () =
            match dc.Brush, dc.Pattern with
            | Some c, _ -> sprintf "fill=\"%s\"" c
            | None, Some (w, h, bits) ->
                // (A monochrome pattern: its 0s in the text colour, its 1s in the background colour, from the device's origin.)
                let id = newId ()
                defs.AppendFormat("<pattern id=\"{0}\" patternUnits=\"userSpaceOnUse\" x=\"0\" y=\"0\" width=\"{1}\" height=\"{2}\"><rect width=\"{1}\" height=\"{2}\" fill=\"{3}\"/>", id, w, h, dc.TextColor) |> ignore
                for y in 0 .. h - 1 do
                    for x in 0 .. w - 1 do
                        if bits.[y].[x] then defs.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"1\" height=\"1\" fill=\"{2}\"/>", x, y, dc.BkColor) |> ignore
                defs.Append("</pattern>") |> ignore
                sprintf "fill=\"url(#%s)\"" id
            | None, None -> "fill=\"none\""
        let strokeAttr () =
            match dc.Pen with
            | Some c, w -> sprintf "stroke=\"%s\" stroke-width=\"%s\"" c (num (max 1.0 (w * scaleY ())))
            | None, _ -> "stroke=\"none\""
        let pointList (o: int) (count: int) (short: bool) =
            [ for i in 0 .. count - 1 ->
                let x, y = if short then float (i16 data (o + 4 * i)), float (i16 data (o + 4 * i + 2)) else float (i32 data (o + 8 * i)), float (i32 data (o + 8 * i + 4))
                device x y ]
        let points (o: int) (count: int) (short: bool) =
            pointList o count short |> List.map (fun (px, py) -> num px + "," + num py) |> String.concat " "
        let rect (l: int) (t: int) (r: int) (b: int) =
            let x0, y0 = device (float l) (float t)
            let x1, y1 = device (float r) (float b)
            min x0 x1, min y0 y1, abs (x1 - x0), abs (y1 - y0)
        let image (bmp: byte[] option) (x: float) (y: float) (w: float) (h: float) =
            match bmp with
            | Some file ->
                emit (sprintf "<image x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" preserveAspectRatio=\"none\" href=\"data:image/bmp;base64,%s\"/>"
                                 (num x) (num y) (num w) (num h) (Convert.ToBase64String file))
            | None -> ()
        // A picture drawn pixel by pixel where `shown` lets it through (runs of a colour as rectangles).
        let pixelRects (pixels: int * int * string[][]) (x: float) (y: float) (w: float) (h: float) (shown: int -> int -> bool) =
            let pw, ph, rows = pixels
            let sx, sy = w / float pw, h / float ph
            let sb = Text.StringBuilder()
            for row in 0 .. ph - 1 do
                let mutable col = 0
                while col < pw do
                    if shown col row then
                        let color = rows.[row].[col]
                        let mutable last = col
                        while last + 1 < pw && shown (last + 1) row && rows.[row].[last + 1] = color do last <- last + 1
                        sb.AppendFormat("<rect x=\"{0}\" y=\"{1}\" width=\"{2}\" height=\"{3}\" fill=\"{4}\"/>", num (x + float col * sx), num (y + float row * sy), num (float (last - col + 1) * sx), num sy, color) |> ignore
                        col <- last + 1
                    else col <- col + 1
            if sb.Length > 0 then emit ("<g shape-rendering=\"crispEdges\">" + sb.ToString() + "</g>")
        // The mask a SRCAND copy left (where its 0s are, the next SRCPAINT copy at the same place is drawn).
        let mutable mask : ((float * float * float * float) * (int * int * string[][])) option = None
        let blit (rop: int) (record: int) (offBmi: int) (cbBmi: int) (offBits: int) (cbBits: int) (x: float) (y: float) (w: float) (h: float) =
            match rop with
            | 0x8800C6 ->
                // SRCAND of a mask: kept for the picture painted over it.
                mask <- pixelsOf data record offBmi offBits |> Option.map (fun pixels -> (x, y, w, h), pixels)
            | 0xEE0086 ->
                // SRCPAINT: the picture where the mask before it is 0 (else where it is not black).
                match pixelsOf data record offBmi offBits with
                | Some pixels ->
                    let shown =
                        match mask with
                        | Some ((mx, my, mw, mh), (pw, ph, bits)) when abs (mx - x) < 0.5 && abs (my - y) < 0.5 && abs (mw - w) < 0.5 && abs (mh - h) < 0.5 ->
                            let _, _, rows = pixels
                            fun col row -> row < ph && col < pw && bits.[row].[col] = "#000000" && row < rows.Length
                        | _ ->
                            let _, _, rows = pixels
                            fun col row -> rows.[row].[col] <> "#000000"
                    pixelRects pixels x y w h shown
                | None -> image (bmpOf data record offBmi cbBmi offBits cbBits) x y w h
                mask <- None
            | _ -> image (bmpOf data record offBmi cbBmi offBits cbBits) x y w h
        // A path being made (BEGINPATH .. ENDPATH) and the one made.
        let mutable inPath = false
        let path = Text.StringBuilder()
        let moveTo (x: float) (y: float) = let px, py = device x y in path.AppendFormat("M{0} {1}", num px, num py) |> ignore
        let lineTo (x: float) (y: float) = let px, py = device x y in path.AppendFormat("L{0} {1}", num px, num py) |> ignore
        let fillRule () = if dc.PolyFill = 2 then "nonzero" else "evenodd"
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
                | 19 -> dc <- { dc with PolyFill = i32 data p }
                | 22 -> dc <- { dc with TextAlign = i32 data p }
                | 24 -> dc <- { dc with TextColor = colorOf (i32 data p) }
                | 25 -> dc <- { dc with BkColor = colorOf (i32 data p) }
                | 27 ->
                    dc <- { dc with Position = float (i32 data p), float (i32 data (p + 4)) }
                    if inPath then moveTo (float (i32 data p)) (float (i32 data (p + 4)))
                | 30 ->
                    // INTERSECTCLIPRECT (logical units).
                    let x, y, w, h = rect (i32 data p) (i32 data (p + 4)) (i32 data (p + 8)) (i32 data (p + 12))
                    intersect x y w h
                | 33 -> stack.Push dc
                | 34 ->
                    let back = -(i32 data p)
                    for _ in 1 .. back - 1 do (if stack.Count > 1 then stack.Pop() |> ignore)
                    if stack.Count > 0 then dc <- stack.Pop()
                | 35 -> dc <- { dc with Transform = f32 data p, f32 data (p + 4), f32 data (p + 8), f32 data (p + 12), f32 data (p + 16), f32 data (p + 20) }
                | 36 ->
                    // MODIFYWORLDTRANSFORM: identity (1), or the transform before (2) / after (3) the current one.
                    let a11, a12, a21, a22, adx, ady = f32 data p, f32 data (p + 4), f32 data (p + 8), f32 data (p + 12), f32 data (p + 16), f32 data (p + 20)
                    let b11, b12, b21, b22, bdx, bdy = dc.Transform
                    let times (x11, x12, x21, x22, xdx, xdy) (y11, y12, y21, y22, ydx, ydy) =
                        x11 * y11 + x12 * y21, x11 * y12 + x12 * y22, x21 * y11 + x22 * y21, x21 * y12 + x22 * y22,
                        xdx * y11 + xdy * y21 + ydx, xdx * y12 + xdy * y22 + ydy
                    match i32 data (p + 24) with
                    | 1 -> dc <- { dc with Transform = 1.0, 0.0, 0.0, 1.0, 0.0, 0.0 }
                    | 2 -> dc <- { dc with Transform = times (a11, a12, a21, a22, adx, ady) (b11, b12, b21, b22, bdx, bdy) }
                    | 3 -> dc <- { dc with Transform = times (b11, b12, b21, b22, bdx, bdy) (a11, a12, a21, a22, adx, ady) }
                    | _ -> ()
                | 37 ->
                    let index = u32 data p
                    if index &&& 0x80000000 <> 0 then
                        // (Stock objects: white, light grey, grey, dark grey, black, no brush; white, black, no pen.)
                        match index &&& 0x7FFFFFFF with
                        | 0 -> dc <- { dc with Brush = Some "#FFFFFF"; Pattern = None }
                        | 1 -> dc <- { dc with Brush = Some "#C0C0C0"; Pattern = None }
                        | 2 -> dc <- { dc with Brush = Some "#808080"; Pattern = None }
                        | 3 -> dc <- { dc with Brush = Some "#404040"; Pattern = None }
                        | 4 -> dc <- { dc with Brush = Some "#000000"; Pattern = None }
                        | 5 -> dc <- { dc with Brush = None; Pattern = None }
                        | 6 -> dc <- { dc with Pen = Some "#FFFFFF", 1.0 }
                        | 7 -> dc <- { dc with Pen = Some "#000000", 1.0 }
                        | 8 -> dc <- { dc with Pen = None, 0.0 }
                        | _ -> ()
                    else
                        match objects.TryGetValue index with
                        | true, Brush color -> dc <- { dc with Brush = color; Pattern = None }
                        | true, PatternBrush pattern -> dc <- { dc with Brush = None; Pattern = Some pattern }
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
                | 93 | 94 ->
                    // CREATEMONOBRUSH / CREATEDIBPATTERNBRUSHPT: a pattern of a monochrome DIB.
                    // (Its bits as they are - the DIB's colours, if any, are not used: a 0 is drawn in the text colour, a
                    // 1 in the background colour.)
                    let bmi = o + u32 data (p + 8)
                    let bits = o + u32 data (p + 16)
                    let w, h, bpp = i32 data (bmi + 4), i32 data (bmi + 8), u16 data (bmi + 14)
                    let rows = abs h
                    let stride = ((w * bpp + 31) / 32) * 4
                    if bpp = 1 && w > 0 && w <= 64 && rows > 0 && rows <= 64 && bits + stride * rows <= data.Length then
                        let row (y: int) = bits + stride * (if h > 0 then rows - 1 - y else y)
                        objects.[u32 data p] <- PatternBrush (w, rows, Array.init rows (fun y -> Array.init w (fun x -> (int data.[row y + x / 8] >>> (7 - x % 8)) &&& 1 = 1)))
                    else objects.[u32 data p] <- Brush (Some "#C0C0C0")
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
                    if inPath && kind = 43 then
                        let l, t, r, b = float (i32 data p), float (i32 data (p + 4)), float (i32 data (p + 8)), float (i32 data (p + 12))
                        moveTo l t; lineTo r t; lineTo r b; lineTo l b; path.Append "Z" |> ignore
                    else
                        let x, y, w, h = rect (i32 data p) (i32 data (p + 4)) (i32 data (p + 8)) (i32 data (p + 12))
                        if kind = 43 then
                            emit (sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" %s %s/>" (num x) (num y) (num w) (num h) (fillAttr ()) (strokeAttr ()))
                        else
                            emit (sprintf "<ellipse cx=\"%s\" cy=\"%s\" rx=\"%s\" ry=\"%s\" %s %s/>" (num (x + w / 2.0)) (num (y + h / 2.0)) (num (w / 2.0)) (num (h / 2.0)) (fillAttr ()) (strokeAttr ()))
                | 3 | 4 | 86 | 87 ->
                    let count = u32 data (p + 16)
                    let short = kind >= 86
                    if inPath then
                        pointList (p + 20) count short |> List.iteri (fun i (px, py) -> path.AppendFormat("{0}{1} {2}", (if i = 0 then "M" else "L"), num px, num py) |> ignore)
                        if kind = 3 || kind = 86 then path.Append "Z" |> ignore
                    else
                        let element = if kind = 3 || kind = 86 then "polygon" else "polyline"
                        let fill = if element = "polygon" then fillAttr () else "fill=\"none\""
                        emit (sprintf "<%s points=\"%s\" %s %s/>" element (points (p + 20) count short) fill (strokeAttr ()))
                | 6 | 89 ->
                    // POLYLINETO: lines on from the current position.
                    let count = u32 data (p + 16)
                    let pts = pointList (p + 20) count (kind = 89)
                    if inPath then pts |> List.iter (fun (px, py) -> path.AppendFormat("L{0} {1}", num px, num py) |> ignore)
                    else
                        let x0, y0 = device (fst dc.Position) (snd dc.Position)
                        emit (sprintf "<polyline points=\"%s\" fill=\"none\" %s/>" ((num x0 + "," + num y0) :: (pts |> List.map (fun (px, py) -> num px + "," + num py)) |> String.concat " ") (strokeAttr ()))
                    match List.tryLast pts with
                    | Some _ ->
                        let last = p + 20 + (count - 1) * (if kind = 89 then 4 else 8)
                        let lx, ly = if kind = 89 then float (i16 data last), float (i16 data (last + 2)) else float (i32 data last), float (i32 data (last + 4))
                        dc <- { dc with Position = lx, ly }
                    | None -> ()
                | 2 | 5 | 85 | 88 ->
                    // POLYBEZIER (from its first point) / POLYBEZIERTO (from the current position).
                    let count = u32 data (p + 16)
                    let short = kind = 85 || kind = 88
                    let pts = pointList (p + 20) count short
                    let d = Text.StringBuilder()
                    let pts =
                        if kind = 2 || kind = 85 then
                            match pts with
                            | (x0, y0) :: rest -> d.AppendFormat("M{0} {1}", num x0, num y0) |> ignore; rest
                            | [] -> []
                        else
                            (if not inPath then let x0, y0 = device (fst dc.Position) (snd dc.Position) in d.AppendFormat("M{0} {1}", num x0, num y0) |> ignore)
                            pts
                    pts |> List.chunkBySize 3 |> List.iter (fun chunk ->
                        match chunk with
                        | [ (x1, y1); (x2, y2); (x3, y3) ] -> d.AppendFormat("C{0} {1} {2} {3} {4} {5}", num x1, num y1, num x2, num y2, num x3, num y3) |> ignore
                        | _ -> ())
                    if inPath then path.Append (d.ToString()) |> ignore
                    else emit (sprintf "<path d=\"%s\" fill=\"none\" %s/>" (d.ToString()) (strokeAttr ()))
                | 7 | 8 | 90 | 91 ->
                    let polys = u32 data (p + 16)
                    let short = kind = 90 || kind = 91
                    let closed = kind = 8 || kind = 91
                    let mutable at = p + 24 + 4 * polys
                    let d = Text.StringBuilder()
                    for i in 0 .. polys - 1 do
                        let count = u32 data (p + 24 + 4 * i)
                        d.Append("M").Append(points at count short) |> ignore
                        if closed then d.Append("Z") |> ignore
                        at <- at + count * (if short then 4 else 8)
                    if inPath then path.Append(d.ToString().Replace(",", " ")) |> ignore
                    elif closed then emit (sprintf "<path d=\"%s\" fill-rule=\"%s\" %s %s/>" (d.ToString()) (fillRule ()) (fillAttr ()) (strokeAttr ()))
                    else emit (sprintf "<path d=\"%s\" fill=\"none\" %s/>" (d.ToString()) (strokeAttr ()))
                | 54 ->
                    let x0, y0 = dc.Position
                    let x1, y1 = float (i32 data p), float (i32 data (p + 4))
                    if inPath then lineTo x1 y1
                    else
                        let a, b = device x0 y0
                        let c, d = device x1 y1
                        emit (sprintf "<line x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\" %s/>" (num a) (num b) (num c) (num d) (strokeAttr ()))
                    dc <- { dc with Position = x1, y1 }
                | 59 -> inPath <- true; path.Clear() |> ignore
                | 60 -> inPath <- false
                | 61 -> if inPath then path.Append "Z" |> ignore
                | 68 -> inPath <- false; path.Clear() |> ignore
                | 62 | 63 | 64 ->
                    // FILLPATH / STROKEANDFILLPATH / STROKEPATH: the path made, then gone.
                    if path.Length > 0 then
                        let fill = if kind = 64 then "fill=\"none\"" else fillAttr ()
                        let stroke = if kind = 62 then "stroke=\"none\"" else strokeAttr ()
                        emit (sprintf "<path d=\"%s\" fill-rule=\"%s\" %s %s/>" (path.ToString()) (fillRule ()) fill stroke)
                    path.Clear() |> ignore
                | 67 ->
                    // SELECTCLIPPATH: the path as the clip (RGN_COPY or AND; an AND of a clip of rectangles keeps the path).
                    if path.Length > 0 then dc <- { dc with Clip = Some (path.ToString()) }
                    path.Clear() |> ignore
                | 75 ->
                    // EXTSELECTCLIPRGN: no region with RGN_COPY: no clip; else the region's rectangles (device units).
                    let cb, mode = u32 data p, u32 data (p + 4)
                    if cb < 32 then (if mode = 5 then dc <- { dc with Clip = None })
                    else
                        let count = u32 data (p + 8 + 8)
                        let rects =
                            [ for k in 0 .. count - 1 ->
                                let r = p + 8 + 32 + 16 * k
                                float (i32 data r), float (i32 data (r + 4)), float (i32 data (r + 8) - i32 data r), float (i32 data (r + 12) - i32 data (r + 4)) ]
                        match mode, rects with
                        | (1 | 5), [ (x, y, w, h) ] -> (if mode = 5 then dc <- { dc with Clip = None }); intersect x y w h
                        | 5, _ -> dc <- { dc with Clip = Some (rects |> List.map (fun (x, y, w, h) -> rectPath x y w h) |> String.concat "") }
                        | _ -> ()
                | 76 ->
                    // BITBLT: destination x, y, cx, cy after the bounds; the raster operation; a source DIB or none.
                    let x, y, w, h = rect (i32 data (p + 16)) (i32 data (p + 20)) (i32 data (p + 16) + i32 data (p + 24)) (i32 data (p + 20) + i32 data (p + 28))
                    // (The source's BITMAPINFO and bits by offsets after its colour and usage: none, a fill by the brush.)
                    let cbBmi = u32 data (p + 80)
                    if cbBmi = 0 then
                        if w > 0.0 && h > 0.0 then
                            match dc.Brush, dc.Pattern with
                            | None, None -> ()
                            | _ -> emit (sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" %s/>" (num x) (num y) (num w) (num h) (fillAttr ()))
                    else blit (u32 data (p + 32)) o (u32 data (p + 76)) cbBmi (u32 data (p + 84)) (u32 data (p + 88)) x y w h
                | 81 ->
                    // STRETCHDIBITS: destination x, y; the DIB's offsets; the raster operation; destination cx, cy at the end.
                    let xDest, yDest = i32 data (p + 16), i32 data (p + 20)
                    let cxDest, cyDest = i32 data (p + 64), i32 data (p + 68)
                    let x, y, w, h = rect xDest yDest (xDest + cxDest) (yDest + cyDest)
                    blit (u32 data (p + 60)) o (u32 data (p + 40)) (u32 data (p + 44)) (u32 data (p + 48)) (u32 data (p + 52)) x y w h
                | 84 ->
                    // EXTTEXTOUTW: the reference point, the characters (from the record's start), the options and their
                    // rectangle, each character's advance.
                    let refX, refY = float (i32 data (p + 28)), float (i32 data (p + 32))
                    let count = u32 data (p + 36)
                    let offset = u32 data (p + 40)
                    let options = u32 data (p + 44)
                    let rx, ry, rw, rh = rect (i32 data (p + 48)) (i32 data (p + 52)) (i32 data (p + 56)) (i32 data (p + 60))
                    let offDx = u32 data (p + 64)
                    let saved = dc
                    // (ETO_OPAQUE: the rectangle filled with the background colour; ETO_CLIPPED: the text cut at it.)
                    if options &&& 2 <> 0 && rw > 0.0 && rh > 0.0 then
                        emit (sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" fill=\"%s\"/>" (num rx) (num ry) (num rw) (num rh) dc.BkColor)
                    if options &&& 4 <> 0 && rw > 0.0 && rh > 0.0 then intersect rx ry rw rh
                    if count > 0 && o + offset + 2 * count <= data.Length then
                        let text = Text.Encoding.Unicode.GetString(data, o + offset, 2 * count)
                        let x, y = device refX refY
                        let face, height, weight, italic = dc.Font
                        let px = height * scaleY ()
                        let anchor = match dc.TextAlign &&& 6 with 6 -> "middle" | 2 -> "end" | _ -> "start"
                        // (TA_BASELINE 24; TA_BOTTOM 8 - the baseline the font's GDI descent above the point; else the
                        // top - the baseline its GDI ascent, the font's Windows ascender, below: Calibri 15 px 14 px down.)
                        let ascent, descent =
                            match face.ToLowerInvariant() with
                            | "calibri" -> 1950.0 / 2048.0, 550.0 / 2048.0
                            | "tahoma" -> 2049.0 / 2048.0, 423.0 / 2048.0
                            | "segoe ui" -> 2210.0 / 2048.0, 514.0 / 2048.0
                            | "arial" -> 1854.0 / 2048.0, 434.0 / 2048.0
                            | "times new roman" -> 1825.0 / 2048.0, 443.0 / 2048.0
                            | "verdana" -> 2059.0 / 2048.0, 430.0 / 2048.0
                            | _ -> 0.92, 0.22
                        let y =
                            match dc.TextAlign &&& 24 with
                            | 24 -> y
                            | 8 -> y - Math.Round (px * descent)
                            | _ -> y + Math.Round (px * ascent)
                        let baseline = "alphabetic"
                        // (Each character at its advance from the record, left-aligned text.)
                        let xs =
                            if offDx > 0 && anchor = "start" && o + offDx + 4 * count <= data.Length then
                                let mutable at = x
                                [ for k in 0 .. count - 1 ->
                                    let here = at
                                    at <- at + float (i32 data (o + offDx + 4 * k)) * scaleX ()
                                    num here ] |> String.concat " "
                            else num x
                        emit (sprintf "<text x=\"%s\" y=\"%s\" font-family=\"'%s', sans-serif\" font-size=\"%spx\" fill=\"%s\"%s%s text-anchor=\"%s\" dominant-baseline=\"%s\" xml:space=\"preserve\">%s</text>"
                                         xs (num y) (esc face) (num px) dc.TextColor
                                         (if weight >= 600 then " font-weight=\"bold\"" else "") (if italic then " font-style=\"italic\"" else "")
                                         anchor baseline (esc text))
                    dc <- { dc with Clip = saved.Clip }
                | _ -> ()
                o <- o + size
        Some (sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%s\" height=\"%s\" viewBox=\"%d %d %s %s\" preserveAspectRatio=\"none\" style=\"display:block;position:absolute;left:0;top:0\">%s%s</svg>"
                    (num width) (num height) left top (num boundsW) (num boundsH)
                    (if defs.Length > 0 then "<defs>" + defs.ToString() + "</defs>" else "") (out.ToString()))
