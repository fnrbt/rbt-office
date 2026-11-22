/// Charts and shapes of a sheet's drawing, drawn from the parts of the file.
///
/// A chart of a file is a part of its own (`xl/charts/chartN.xml`, a
/// `c:chartSpace`) that a drawing anchor names; the editor keeps the anchor
/// and the part as they are and writes them back untouched. This module
/// reads such a part into a model (`Chart`), lays it out for a rectangle
/// (`layoutChart`: a `Scene` of rectangles, lines, paths and texts, in the
/// pixels of that rectangle) and writes a scene as SVG (`sceneSvg`). The
/// model and the layout know nothing of the browser: a PDF writer can draw
/// the same scene.
///
/// Values are read live from the cells a series names (`c:f`), through the
/// `ChartData` a host gives; where a reference cannot be read - another
/// workbook, a name, a sheet that is not there - the values the file cached
/// (`c:numCache` / `c:strCache`) are drawn.
///
/// Drawn: bar and column charts (clustered, stacked, 100 % stacked), line
/// charts (markers), pie and doughnut charts, area charts (standard,
/// stacked), scatter charts, and combinations of them with a secondary
/// axis; the title, the legend, the axes with their number formats and an
/// automatic scale, grid lines, the series' colours (`c:spPr`, the theme),
/// data labels. `parseChart` gives None for anything else (3-D, radar,
/// stock, surface, bubble, chartEx): the caller shows a placeholder.
///
/// Shapes (`xdr:sp`), connectors and group shapes (`xdr:grpSp`) are drawn by
/// `objectScene`: fill, outline, preset geometry (rectangles, rounded
/// rectangles, ellipses, lines; any other preset as its rectangle) and text
/// with its formatting; charts and pictures inside a group at their place in it.
module Rbt.Spreadsheet.Charts

open System
open Rbt.Spreadsheet.Model

// ---------------------------------------------------------------------------
// XML
// ---------------------------------------------------------------------------

let private localName (name: string) =
    match name.IndexOf ':' with
    | -1 -> name
    | at -> name.Substring(at + 1)

let private kids (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    match node with
    | XmlDsl.Element (_, _, children, _) -> children
    | _ -> []

let private nameOf (node: XmlDsl.XmlNode) : string =
    match node with
    | XmlDsl.Element (name, _, _, _) -> localName name
    | _ -> ""

let private childrenNamed (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    kids node |> List.filter (fun child -> nameOf child = name)

let private child (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    kids node |> List.tryFind (fun child -> nameOf child = name)

let private attr (name: string) (node: XmlDsl.XmlNode) : string option =
    match node with
    | XmlDsl.Element (_, attrs, _, _) -> Map.tryFind name attrs
    | _ -> None

/// The `val` of a child element (`<c:barDir val="col"/>`).
let private valOf (name: string) (node: XmlDsl.XmlNode) : string option =
    child name node |> Option.bind (attr "val")

let private path (names: string list) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    names |> List.fold (fun (at: XmlDsl.XmlNode option) name -> at |> Option.bind (child name)) (Some node)

let rec private descendants (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    kids node |> List.collect (fun child -> (if nameOf child = name then [ child ] else []) @ descendants name child)

let rec private innerText (node: XmlDsl.XmlNode) : string =
    match node with
    | XmlDsl.Text text -> text
    | XmlDsl.Element (_, _, children, _) -> children |> List.map innerText |> String.concat ""
    | _ -> ""

let private number (text: string) : float option =
    match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, value when not (Double.IsNaN value) && not (Double.IsInfinity value) -> Some value
    | _ -> None

let private isOn (value: string option) (otherwise: bool) =
    match value with
    | Some "0" | Some "false" -> false
    | Some _ -> true
    | None -> otherwise

/// (The parser drops the white space at the ends of a text unless its element says
/// `xml:space="preserve"`: in DrawingML the space at the end of a run is text.)
let private spacesKept (xml: string) : string =
    xml.Replace("<a:t>", "<a:t xml:space=\"preserve\">")

/// `mc:AlternateContent` stands for its fallback (or its first choice).
let rec private flattened (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    if nameOf node = "AlternateContent" then
        match child "Fallback" node, child "Choice" node with
        // (A choice that holds a frame - a slicer, a chart of a newer kind - is that object: its
        // fallback is a text box that says the object is not there.)
        | _, Some choice when not (descendants "graphicFrame" choice).IsEmpty -> kids choice |> List.collect flattened
        | Some fallback, _ when not (kids fallback).IsEmpty -> kids fallback |> List.collect flattened
        | _, Some choice -> kids choice |> List.collect flattened
        | _ -> []
    else [ node ]

/// A number as text: two decimals at most, a point for the decimals.
let private fmtNumber (value: float) : string =
    let rounded = Math.Round (value * 100.0) / 100.0
    if Double.IsNaN rounded || Double.IsInfinity rounded then "0" else rounded.ToString(Globalization.CultureInfo.InvariantCulture)

// ---------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------

/// The colours of the workbook's theme by name (`dk1`, `lt1`, `dk2`, `lt2`,
/// `accent1` .. `accent6`, `hlink`, `folHlink`), as `RRGGBB`.
type ThemeColors = Map<string, string>

/// Office's own theme (2013 and later), for a workbook without one.
let defaultTheme : ThemeColors =
    Map.ofList [ "dk1", "000000"; "lt1", "FFFFFF"; "dk2", "44546A"; "lt2", "E7E6E6"
                 "accent1", "4472C4"; "accent2", "ED7D31"; "accent3", "A5A5A5"; "accent4", "FFC000"; "accent5", "5B9BD5"; "accent6", "70AD47"
                 "hlink", "0563C1"; "folHlink", "954F72" ]

/// The formatting schemes (`a:fmtScheme`) of the themes read, by the key a theme's colours name theirs
/// under (`#fmtScheme`): a shape's style reference (`fillRef idx`, `lnRef idx`) is one of its fills and lines.
let private formatSchemes = Collections.Generic.Dictionary<string, XmlDsl.XmlNode>()

/// The colours of a theme part (`a:clrScheme`), and its formatting scheme (under `#fmtScheme`).
let parseTheme (xml: string) : ThemeColors =
    try
        let root = XmlDsl.parse xml
        let colors =
            match descendants "clrScheme" root |> List.tryHead with
            | None -> defaultTheme
            | Some scheme ->
                kids scheme
                |> List.fold (fun (colors: ThemeColors) entry ->
                    let color =
                        match child "srgbClr" entry |> Option.bind (attr "val"), child "sysClr" entry |> Option.bind (attr "lastClr") with
                        | Some rgb, _ -> Some rgb
                        | None, Some rgb -> Some rgb
                        | _ -> None
                    match color with
                    | Some rgb when nameOf entry <> "" -> Map.add (nameOf entry) (rgb.ToUpperInvariant()) colors
                    | _ -> colors) defaultTheme
        // (The theme's fonts, which "+mn-lt" / "+mj-lt" name - and a chart's text that names none is in the minor.)
        let font (name: string) = descendants name root |> List.tryHead |> Option.bind (child "latin") |> Option.bind (attr "typeface") |> Option.filter (fun face -> face <> "")
        let colors =
            [ "#minorFont", font "minorFont"; "#majorFont", font "majorFont" ]
            |> List.fold (fun (colors: ThemeColors) (key, face) -> match face with Some face -> Map.add key face colors | None -> colors) colors
        match descendants "fmtScheme" root |> List.tryHead with
        | Some scheme ->
            let key = string (formatSchemes.Count + 1)
            formatSchemes.[key] <- scheme
            Map.add "#fmtScheme" key colors
        | None -> colors
    with _ -> defaultTheme

/// Office's own formatting scheme (2013 and later): the fills and lines of a theme that has none.
let private defaultFormatScheme =
    lazy (
        XmlDsl.parse (
            "<a:fmtScheme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office\"><a:fillStyleLst>"
            + "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"110000\"/><a:satMod val=\"105000\"/><a:tint val=\"67000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"103000\"/><a:tint val=\"73000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"105000\"/><a:satMod val=\"109000\"/><a:tint val=\"81000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:satMod val=\"103000\"/><a:lumMod val=\"102000\"/><a:tint val=\"94000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:satMod val=\"110000\"/><a:lumMod val=\"100000\"/><a:shade val=\"100000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:lumMod val=\"99000\"/><a:satMod val=\"120000\"/><a:shade val=\"78000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "</a:fillStyleLst><a:lnStyleLst>"
            + "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            + "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            + "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>"
            + "</a:lnStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>"
            + "<a:solidFill><a:schemeClr val=\"phClr\"><a:tint val=\"95000\"/><a:satMod val=\"170000\"/></a:schemeClr></a:solidFill>"
            + "<a:gradFill rotWithShape=\"1\"><a:gsLst><a:gs pos=\"0\"><a:schemeClr val=\"phClr\"><a:tint val=\"93000\"/><a:satMod val=\"150000\"/><a:shade val=\"98000\"/><a:lumMod val=\"102000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"50000\"><a:schemeClr val=\"phClr\"><a:tint val=\"98000\"/><a:satMod val=\"130000\"/><a:shade val=\"90000\"/><a:lumMod val=\"103000\"/></a:schemeClr></a:gs>"
            + "<a:gs pos=\"100000\"><a:schemeClr val=\"phClr\"><a:shade val=\"63000\"/><a:satMod val=\"120000\"/></a:schemeClr></a:gs></a:gsLst><a:lin ang=\"5400000\" scaled=\"0\"/></a:gradFill>"
            + "</a:bgFillStyleLst></a:fmtScheme>"))

/// The fill (`fillRef`: 1.. the scheme's fills, 1001.. its background fills) or line (`lnRef`: 1.. its
/// lines) a style reference of index `idx` names, from the theme's formatting scheme.
let private schemeEntry (theme: ThemeColors) (lists: string list) (idx: int) : XmlDsl.XmlNode option =
    let scheme =
        match Map.tryFind "#fmtScheme" theme |> Option.bind (fun key -> match formatSchemes.TryGetValue key with | true, node -> Some node | _ -> None) with
        | Some node -> node
        | None -> defaultFormatScheme.Value
    let listName, index =
        match lists with
        | [ fills; backgrounds ] -> (if idx >= 1001 then backgrounds, idx - 1001 else fills, idx - 1)
        | _ -> List.head lists, idx - 1
    descendants listName scheme |> List.tryHead |> Option.bind (fun list -> kids list |> List.filter (fun node -> nameOf node <> "") |> List.tryItem index)

let private hex2 (value: float) =
    let n = int (Math.Round (max 0.0 (min 255.0 value)))
    let digits = "0123456789ABCDEF"
    string digits.[n / 16] + string digits.[n % 16]

let private rgbOf (hex: string) : float * float * float =
    let hex = hex.TrimStart '#'
    if hex.Length < 6 then 0.0, 0.0, 0.0
    else
        let part (at: int) = float (Convert.ToInt32(hex.Substring(at, 2), 16))
        part (hex.Length - 6), part (hex.Length - 4), part (hex.Length - 2)

let private hslOf (r: float, g: float, b: float) : float * float * float =
    let r, g, b = r / 255.0, g / 255.0, b / 255.0
    let hi, lo = max r (max g b), min r (min g b)
    let l = (hi + lo) / 2.0
    if hi = lo then 0.0, 0.0, l
    else
        let d = hi - lo
        let s = if l > 0.5 then d / (2.0 - hi - lo) else d / (hi + lo)
        let h =
            if hi = r then (g - b) / d + (if g < b then 6.0 else 0.0)
            elif hi = g then (b - r) / d + 2.0
            else (r - g) / d + 4.0
        h / 6.0, s, l

let private rgbOfHsl (h: float, s: float, l: float) : float * float * float =
    if s = 0.0 then l * 255.0, l * 255.0, l * 255.0
    else
        let q = if l < 0.5 then l * (1.0 + s) else l + s - l * s
        let p = 2.0 * l - q
        let channel (t: float) =
            let t = if t < 0.0 then t + 1.0 elif t > 1.0 then t - 1.0 else t
            if t < 1.0 / 6.0 then p + (q - p) * 6.0 * t
            elif t < 0.5 then q
            elif t < 2.0 / 3.0 then p + (q - p) * (2.0 / 3.0 - t) * 6.0
            else p
        channel (h + 1.0 / 3.0) * 255.0, channel h * 255.0, channel (h - 1.0 / 3.0) * 255.0

/// A DrawingML colour element (`a:srgbClr`, `a:schemeClr`, `a:sysClr`,
/// `a:prstClr`) with its modifiers (`lumMod`, `lumOff`, `tint`, `shade`), as
/// `#RRGGBB`, and its opacity (`alpha`).
/// A channel (0..255) mixed with `toward` (0 black, 1 white) by `share`, in linear light.
let private linearMix (channel: float) (share: float) (toward: float) : float =
    let c = max 0.0 (min 255.0 channel) / 255.0
    let linear = if c <= 0.04045 then c / 12.92 else Math.Pow ((c + 0.055) / 1.055, 2.4)
    let mixed = linear * (1.0 - share) + toward * share
    (if mixed <= 0.0031308 then mixed * 12.92 else 1.055 * Math.Pow (mixed, 1.0 / 2.4) - 0.055) * 255.0

let rec private colorOfNodeWith (theme: ThemeColors) (placeholder: XmlDsl.XmlNode option) (node: XmlDsl.XmlNode) : (string * float) option =
    let baseColor =
        match nameOf node with
        // (The colour of a theme's style: the one the shape's style reference gives, with its modifiers.)
        | "schemeClr" when attr "val" node = Some "phClr" && placeholder.IsSome ->
            placeholder |> Option.bind (colorOfNodeWith theme None) |> Option.map (fun (hex, _) -> hex.TrimStart '#')
        | "srgbClr" -> attr "val" node
        | "sysClr" -> attr "lastClr" node |> Option.orElse (match attr "val" node with Some "window" -> Some "FFFFFF" | Some _ -> Some "000000" | None -> None)
        | "schemeClr" ->
            let name =
                match attr "val" node with
                | Some "tx1" -> "dk1" | Some "bg1" -> "lt1" | Some "tx2" -> "dk2" | Some "bg2" -> "lt2"
                | Some other -> other
                | None -> "dk1"
            Map.tryFind name theme |> Option.orElse (Map.tryFind name defaultTheme)
        | "prstClr" ->
            (match attr "val" node with
             | Some "white" -> Some "FFFFFF" | Some "black" -> Some "000000" | Some "red" -> Some "FF0000" | Some "green" -> Some "008000"
             | Some "blue" -> Some "0000FF" | Some "yellow" -> Some "FFFF00" | Some "gray" | Some "grey" -> Some "808080"
             | _ -> Some "000000")
        | _ -> None
    baseColor
    |> Option.map (fun hex ->
        let mutable r, g, b = rgbOf hex
        let mutable alpha = 1.0
        for modifier in kids node do
            let amount = attr "val" modifier |> Option.bind number |> Option.map (fun value -> value / 100000.0) |> Option.defaultValue 1.0
            match nameOf modifier with
            | "lumMod" -> (let h, s, l = hslOf (r, g, b) in let r2, g2, b2 = rgbOfHsl (h, s, min 1.0 (l * amount)) in r <- r2; g <- g2; b <- b2)
            | "lumOff" -> (let h, s, l = hslOf (r, g, b) in let r2, g2, b2 = rgbOfHsl (h, s, min 1.0 (max 0.0 (l + amount))) in r <- r2; g <- g2; b <- b2)
            // (Tint and shade mix with white and black in linear light, as Office does: a 50 % tint of
            // black is C0C0C0-ish in Excel 2019, not 808080.)
            | "tint" -> (r <- linearMix r (1.0 - amount) 1.0; g <- linearMix g (1.0 - amount) 1.0; b <- linearMix b (1.0 - amount) 1.0)
            | "shade" -> (r <- linearMix r (1.0 - amount) 0.0; g <- linearMix g (1.0 - amount) 0.0; b <- linearMix b (1.0 - amount) 0.0)
            | "satMod" -> (let h, s, l = hslOf (r, g, b) in let r2, g2, b2 = rgbOfHsl (h, min 1.0 (s * amount), l) in r <- r2; g <- g2; b <- b2)
            | "alpha" -> alpha <- amount
            | _ -> ()
        "#" + hex2 r + hex2 g + hex2 b, alpha)

let private colorOfNode (theme: ThemeColors) (node: XmlDsl.XmlNode) : (string * float) option = colorOfNodeWith theme None node

/// How a shape is filled: nothing stated, no fill, or a colour - `#RRGGBB`, or for a
/// gradient of a shape or a series what `gradientOf` reads (a line and a text take its first colour).
type Fill =
    | FillAuto
    | FillNone
    | FillColor of color: string * opacity: float

/// A line: None when the element says nothing of it.
type Stroke = {
    Paint: Fill
    /// In pixels.
    Width: float
    Dash: string option
}

/// A fill of a primitive is a colour (`#RRGGBB`) or a gradient written as
/// `gradient(angle;offset color opacity;...)` - `angle` in degrees clockwise from
/// left-to-right, or `radial`; offsets 0 .. 1. `gradientOf` reads one back: None for a plain colour.
let gradientOf (fill: string) : (float option * (float * string * float) list) option =
    if not (fill.StartsWith "gradient(") || not (fill.EndsWith ")") then None
    else
        let parts = fill.Substring(9, fill.Length - 10).Split(';')
        if parts.Length < 3 then None
        else
            let angle = if parts.[0] = "radial" then None else Some (number parts.[0] |> Option.defaultValue 0.0)
            let stops =
                parts |> Array.skip 1 |> Array.toList
                |> List.choose (fun stop ->
                    match stop.Split(' ') with
                    | [| offset; color; opacity |] -> Some (number offset |> Option.defaultValue 0.0, color, number opacity |> Option.defaultValue 1.0)
                    | _ -> None)
            if stops.Length < 2 then None else Some (angle, stops)

/// One colour for a fill: the colour itself, or the mean of a gradient's colours
/// (for a renderer without gradients).
let solidOf (fill: string) : string =
    match gradientOf fill with
    | None -> fill
    | Some (_, stops) ->
        let n = float stops.Length
        let r, g, b = stops |> List.fold (fun (r, g, b) (_, color: string, _) -> let cr, cg, cb = rgbOf color in r + cr / n, g + cg / n, b + cb / n) (0.0, 0.0, 0.0)
        "#" + hex2 r + hex2 g + hex2 b

/// The fill an element states. `gradients`: a gradient of two or more stops is a gradient
/// (a shape, a series, a chart's background); otherwise its first colour (a line, a text).
let private fillWithPlaceholder (placeholder: XmlDsl.XmlNode option) (gradients: bool) (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill =
    let colorIn (node: XmlDsl.XmlNode) = kids node |> List.tryPick (colorOfNodeWith theme placeholder)
    match child "noFill" holder, child "solidFill" holder, child "gradFill" holder, child "pattFill" holder with
    | Some _, _, _, _ -> FillNone
    | _, Some solid, _, _ -> (match colorIn solid with Some (color, alpha) -> FillColor (color, alpha) | None -> FillAuto)
    | _, _, Some gradient, _ ->
        let stops =
            descendants "gs" gradient
            |> List.choose (fun stop ->
                match colorIn stop, attr "pos" stop |> Option.bind number with
                | Some (color, alpha), Some position -> Some (position / 100000.0, color, alpha)
                | Some (color, alpha), None -> Some (0.0, color, alpha)
                | None, _ -> None)
            |> List.sortBy (fun (position, _, _) -> position)
        (match stops with
         | [] -> FillAuto
         | [ (_, color, alpha) ] -> FillColor (color, alpha)
         | (_, color, alpha) :: _ when not gradients -> FillColor (color, alpha)
         | _ ->
            let direction =
                match child "lin" gradient with
                | Some linear -> fmtNumber ((attr "ang" linear |> Option.bind number |> Option.defaultValue 0.0) / 60000.0)
                | None -> if (child "path" gradient).IsSome then "radial" else "90"
            let text = stops |> List.map (fun (position, color, alpha) -> fmtNumber position + " " + color + " " + fmtNumber alpha) |> String.concat ";"
            FillColor ("gradient(" + direction + ";" + text + ")", 1.0))
    | _, _, _, Some pattern -> (match child "fgClr" pattern |> Option.bind colorIn with Some (color, alpha) -> FillColor (color, alpha) | None -> FillAuto)
    | _ -> FillAuto

let private fillWith (gradients: bool) (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill = fillWithPlaceholder None gradients theme holder

/// The fill a shape's style reference names (`<a:fillRef idx>` and its colour): none for 0, else the
/// theme's fill of that index in the reference's colour (Excel 2019: an Office 2007 theme's `idx="2"` in
/// `dk1` is a grey gradient, not black).
let private styleFill (theme: ThemeColors) (reference: XmlDsl.XmlNode) : Fill =
    match attr "idx" reference |> Option.bind number |> Option.map int with
    | Some 0 -> FillNone
    | Some idx ->
        let color = kids reference |> List.tryFind (fun node -> nameOf node <> "")
        match schemeEntry theme [ "fillStyleLst"; "bgFillStyleLst" ] idx with
        | Some entry -> fillWithPlaceholder color true theme (XmlDsl.Element ("spPr", Map.empty, [ entry ], XmlDsl.Span.none))
        | None -> (match color |> Option.bind (colorOfNode theme) with Some (hex, alpha) -> FillColor (hex, alpha) | None -> FillAuto)
    | None -> FillAuto

/// How many pixels wide Excel 2019 draws a line of `emu` (at 100 %): whole pixels, at least one - 0.25 .. 1 pt
/// one, 1.5 pt two, 2.25 pt three, 3 pt four, 4.5 pt six (synthetic-chart-lines "Widths", `Chart.Export`).
let private linePixels (emu: float) : float = max 1.0 (Math.Floor (emu / 9525.0 + 0.5))

/// The line a shape's style reference names (`<a:lnRef idx>`): none for 0, else the theme's line of that
/// index in the reference's colour - its colour and its width in pixels.
let private styleLine (theme: ThemeColors) (reference: XmlDsl.XmlNode) : (string * float) option =
    let color = kids reference |> List.tryFind (fun node -> nameOf node <> "")
    match attr "idx" reference |> Option.bind number |> Option.map int with
    | Some 0 -> None
    | idx ->
        let entry = idx |> Option.bind (schemeEntry theme [ "lnStyleLst" ])
        let width = entry |> Option.bind (attr "w") |> Option.bind number |> Option.map linePixels |> Option.defaultValue 1.0
        match entry |> Option.map (fillWithPlaceholder color false theme) with
        | Some (FillColor (hex, _)) -> Some (hex, width)
        | Some FillNone -> None
        | _ -> color |> Option.bind (colorOfNode theme) |> Option.map (fun (hex, _) -> hex, width)

let private fillOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill = fillWith false theme holder

/// The fill of a shape, a series, a background: a gradient is one.
let private shapeFillOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill = fillWith true theme holder

let private strokeOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Stroke option =
    child "ln" holder
    |> Option.map (fun line ->
        { Paint = fillOf theme line
          Width = (match attr "w" line |> Option.bind number with Some emu -> linePixels emu | None -> 1.0)
          Dash = (match valOf "prstDash" line with Some "solid" | None -> None | Some other -> Some other) })

/// The font size (points), the colour and the weight a text properties
/// element (`c:txPr`, `a:rPr`, `a:defRPr`) states.
type TextStyle = {
    Size: float option
    Color: string option
    Bold: bool option
    Italic: bool option
    /// The typeface (`a:latin`; a theme's "+mn-lt" / "+mj-lt" its minor / major font).
    Face: string option
}

let private noTextStyle = { Size = None; Color = None; Bold = None; Italic = None; Face = None }

let private runStyle (theme: ThemeColors) (props: XmlDsl.XmlNode) : TextStyle =
    { Size = attr "sz" props |> Option.bind number |> Option.map (fun value -> value / 100.0)
      Color = (match fillOf theme props with FillColor (color, _) -> Some color | _ -> None)
      Bold = attr "b" props |> Option.map (fun value -> value = "1" || value = "true")
      Italic = attr "i" props |> Option.map (fun value -> value = "1" || value = "true")
      Face =
        child "latin" props |> Option.bind (attr "typeface")
        |> Option.bind (fun face ->
            if face.StartsWith "+mn-" then Map.tryFind "#minorFont" theme
            elif face.StartsWith "+mj-" then Map.tryFind "#majorFont" theme
            elif face = "" || face.StartsWith "+" then None
            else Some face) }

let private over (first: TextStyle) (second: TextStyle) : TextStyle =
    { Size = first.Size |> Option.orElse second.Size; Color = first.Color |> Option.orElse second.Color
      Bold = first.Bold |> Option.orElse second.Bold; Italic = first.Italic |> Option.orElse second.Italic
      Face = first.Face |> Option.orElse second.Face }

/// The style of `c:txPr` (its paragraph's default run properties).
let private textPropsOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : TextStyle =
    match child "txPr" holder |> Option.bind (path [ "p"; "pPr"; "defRPr" ]) with
    | Some props -> runStyle theme props
    | None -> noTextStyle

// ---------------------------------------------------------------------------
// The model of a chart
// ---------------------------------------------------------------------------

/// The cells a series reads, and what the file cached of them.
type DataRef = {
    /// `Sheet1!$B$2:$B$7` (None: literal values).
    Formula: string option
    Numbers: Map<int, float>
    Strings: Map<int, string>
    Count: int
    FormatCode: string option
    /// Its formula differs from the one in the original chart part (including
    /// a chartEx defined name's original target): the old cache's point count
    /// must not limit the new live range.
    FollowedEdit: bool
    /// A pivot chart's range can include the grand total, which was never a
    /// point of the chart. Only those charts use the cache's point count to
    /// bound live data; ordinary charts can grow after save and reopen.
    LimitToCachedPoints: bool
    /// The outer levels of a multi-level category cache (`c:multiLvlStrCache`), from the inner out: each a level's
    /// labels by the point each group begins at.
    Outer: Map<int, string> list
}

type LabelSpec = {
    ShowValue: bool
    ShowPercent: bool
    ShowCategory: bool
    ShowSeries: bool
    /// "ctr" | "inEnd" | "inBase" | "outEnd" | "bestFit" | "t" | "b" | "l" | "r"
    Position: string option
    FormatCode: string option
    Style: TextStyle
    /// What stands between the parts of a label (`c:separator`).
    Separator: string option
    /// The points whose label is taken away (`c:dLbl` with `c:delete`).
    Deleted: Set<int>
    /// A point's label of its own (`c:dLbl`): what it shows, its look.
    Points: Map<int, LabelSpec>
    /// A label moved from its place (`c:dLbl`'s manual layout): by these fractions of the chart's width and height.
    Offset: (float * float) option
}

type Series = {
    Order: int
    Index: int
    Name: DataRef option
    Categories: DataRef option
    Values: DataRef option
    /// The x values of a scatter series.
    XValues: DataRef option
    Fill: Fill
    Line: Stroke option
    /// "none" | "auto" | "circle" | "square" | "diamond" | "triangle" | ...
    MarkerSymbol: string option
    MarkerSize: float
    MarkerFill: Fill
    MarkerLine: Stroke option
    Smooth: bool
    /// The fills of single points (`c:dPt`), by index.
    PointFills: Map<int, Fill>
    /// The outlines of single points that state their look (`c:dPt` `c:spPr`), by index: in place of the
    /// series' - None where the point's look has no `a:ln` (Excel 2019 then draws none).
    PointLines: Map<int, Stroke option>
    Labels: LabelSpec option
    /// The sizes of the bubbles of a bubble chart's series.
    BubbleSizes: DataRef option
    /// A pie slice pulled out, in per cent of the radius.
    Explosion: float
}

/// How a chartEx histogram cuts its values into bins (`cx:binning`): set by a width or a count, else Excel's own
/// rule; values below an underflow / above an overflow in a bin of their own; bins closed on the right (`(a, b]`,
/// the first `[a, b]`) unless `intervalClosed="l"`.
type Binning = {
    ClosedLeft: bool
    Size: float option
    Count: int option
    Underflow: float option
    Overflow: float option
}

/// What a chartEx box and whisker series shows (`cx:statistics`, `cx:visibility`).
type BoxOptions = {
    /// Quartiles of the median included (`quartileMethod="inclusive"`), else excluded.
    Inclusive: bool
    MeanLine: bool
    MeanMarker: bool
    /// The points between the whiskers' ends (`nonoutliers`).
    InnerPoints: bool
    Outliers: bool
}

/// The figures of a box: whiskers' ends, quartiles, median, mean, the points inside and outside the whiskers.
type BoxFigures = {
    Low: float
    Q1: float
    Median: float
    Q3: float
    High: float
    Mean: float
    Inner: float list
    Outside: float list
}

type PlotKind =
    | BarPlot
    | LinePlot
    | PiePlot
    | DoughnutPlot
    | AreaPlot
    | ScatterPlot
    | RadarPlot
    | BubblePlot
    | StockPlot
    /// A chartEx waterfall: the points that are totals (from zero), and whether lines join the bars.
    | WaterfallPlot of subtotals: int list * connectors: bool
    /// A chartEx funnel: bars centred on each other, one per category, top to bottom.
    | FunnelPlot
    /// A chartEx histogram (`clusteredColumn`): its values cut into bins - by its categories where it has them -,
    /// a bar a bin; a Pareto chart (`paretoLine` beside it) the bins from the most to the least with their
    /// running share on a second axis.
    | HistogramPlot of binning: Binning * pareto: bool
    /// A chartEx box and whisker chart: a box a series a category.
    | BoxWhiskerPlot of BoxOptions
    /// A box and whisker chart's boxes as laid out: the figures of each series' boxes, a category each.
    | BoxesPlot of figures: BoxFigures option [] [] * options: BoxOptions
    /// A chartEx treemap: its categories (levels of them) as rectangles of their values' sizes; how a parent's label
    /// stands (`parentLabelLayout`: "overlapping" | "banner" | "none").
    | TreemapPlot of parentLabels: string
    /// A chartEx sunburst: its categories' levels as rings, the outer level's parts within their parent's.
    | SunburstPlot
    /// A chartEx chart of a kind that is not drawn (`layoutId`: "regionMap", ...): a placeholder.
    | UndrawnPlot of layoutId: string

type Plot = {
    Kind: PlotKind
    /// Bars that run to the right (`barDir="bar"`).
    Horizontal: bool
    /// "clustered" | "stacked" | "percentStacked" | "standard"
    Grouping: string
    Series: Series list
    GapWidth: float
    Overlap: float
    HoleSize: float
    FirstSliceAngle: float
    VaryColors: bool
    AxisIds: string list
    Labels: LabelSpec option
    /// A line chart draws markers (`c:marker val="1"`), a scatter chart lines and / or markers.
    Markers: bool
    ScatterStyle: string
    /// "standard" | "marker" | "filled" (a radar chart).
    RadarStyle: string
    /// The lines from the highest to the lowest value of a category (`c:hiLowLines`: a stock chart, a line chart).
    HiLowLines: Stroke option
    /// The bars between the first and the last series (`c:upDownBars`): the fill of a rise and of a fall.
    UpDownBars: (Fill * Fill) option
    /// The largest bubble, in per cent of the default (`c:bubbleScale`).
    BubbleScale: float
}

type Axis = {
    Id: string
    /// "cat" | "val" | "date" | "ser"
    Kind: string
    /// A value axis crossing on the categories (`crossBetween` midCat): the points are at the ends of the
    /// plot, on the ticks, not in the middle of bands.
    OnTicks: bool
    /// "b" | "l" | "t" | "r"
    Position: string
    Deleted: bool
    Min: float option
    Max: float option
    MajorUnit: float option
    Reversed: bool
    FormatCode: string option
    SourceLinked: bool
    Gridlines: Stroke option
    Title: string option
    TitleStyle: TextStyle
    /// The title's runs, a list per paragraph, each in its own look (over the title's): empty unless rich text.
    TitleRuns: (string * TextStyle) list list
    LabelStyle: TextStyle
    Line: Stroke option
    LabelsShown: bool
    CrossesId: string option
    /// The base of a logarithmic scale (`c:logBase`).
    LogBase: float option
    /// What the labels are divided by (`c:dispUnits`: 1000 for thousands), and the label that says so.
    DisplayUnit: float
    DisplayUnitLabel: string option
    /// A date axis' unit of its points (`c:baseTimeUnit`: "days" | "months" | "years") and of its major step.
    BaseTimeUnit: string option
    MajorTimeUnit: string option
    /// The labels turned (`a:bodyPr rot`), in degrees; None where Excel chooses (no `rot`, or -60000000).
    LabelRotation: float option
    /// Every how many categories a label is (`c:tickLblSkip`); None where Excel chooses.
    LabelSkip: int option
}

/// Where the file puts a title, a legend or the plot area (`c:manualLayout`): fractions of the
/// chart's width and height.
type ManualLayout = {
    /// The plot area without its axes' labels (`layoutTarget="inner"`).
    Inner: bool
    /// x / y are from the chart's left / top edge (`edge`); otherwise an offset from where the
    /// element would be on its own.
    XEdge: bool
    YEdge: bool
    X: float option
    Y: float option
    W: float option
    H: float option
}

type Legend = {
    /// "r" | "l" | "t" | "b" | "tr"
    Position: string
    Style: TextStyle
    Layout: ManualLayout option
    /// Drawn over the plot area: no room is made for it.
    Overlay: bool
    /// The entries taken out of the legend (`c:legendEntry` with `c:delete`), by index.
    Deleted: Set<int>
    /// An entry's text of its own look (`c:legendEntry` with `c:txPr`), by index.
    EntryStyles: Map<int, TextStyle>
    /// The legend box's own fill and outline (`c:spPr`).
    Fill: Fill
    Border: Stroke option
}

type Chart = {
    Title: string option
    TitleStyle: TextStyle
    /// The title's runs, a list per paragraph, each in its own look (over the title's): empty unless rich text.
    TitleRuns: (string * TextStyle) list list
    AutoTitleDeleted: bool
    Plots: Plot list
    Axes: Axis list
    Legend: Legend option
    Background: Fill
    Border: Stroke option
    PlotBackground: Fill
    /// The plot area's outline (`c:plotArea` `c:spPr` `a:ln`).
    PlotBorder: Stroke option
    TextStyle: TextStyle
    RoundedCorners: bool
    Date1904: bool
    TitleLayout: ManualLayout option
    /// The title is drawn over the plot area: no room is made for it.
    TitleOverlay: bool
    PlotLayout: ManualLayout option
    /// `c:style` (1 .. 48): which colours the series get that state none.
    Style: int
    /// The shapes drawn over the chart (`c:userShapes`: the drawing part the chart's relationships name), as the
    /// part's XML; None when it has none.
    UserShapes: string option
    /// One of the newer kinds (a chartEx part's), laid out as Excel lays those out.
    ChartEx: bool
}

// ---------------------------------------------------------------------------
// Reading a chart part
// ---------------------------------------------------------------------------

let private dataRefOf (holder: XmlDsl.XmlNode) : DataRef option =
    let cached (cache: XmlDsl.XmlNode) (formula: string option) =
        let mutable numbers = Map.empty<int, float>
        let mutable strings = Map.empty<int, string>
        let mutable count = 0
        let points (node: XmlDsl.XmlNode) =
            for point in childrenNamed "pt" node do
                match attr "idx" point |> Option.bind number with
                | Some index ->
                    let text = child "v" point |> Option.map innerText |> Option.defaultValue ""
                    strings <- Map.add (int index) text strings
                    (match number text with Some value -> numbers <- Map.add (int index) value numbers | None -> ())
                    if int index + 1 > count then count <- int index + 1
                | None -> ()
        // (A multi-level category cache: its innermost level names the points.)
        let levels = childrenNamed "lvl" cache
        (match levels with
         | [] -> points cache
         | level :: _ -> points level)
        (match valOf "ptCount" cache |> Option.bind number with Some stated when int stated > count -> count <- int stated | _ -> ())
        let outer =
            levels |> List.skip (min 1 levels.Length)
            |> List.map (fun level ->
                childrenNamed "pt" level
                |> List.choose (fun point -> attr "idx" point |> Option.bind number |> Option.map (fun index -> int index, (child "v" point |> Option.map innerText |> Option.defaultValue "")))
                |> Map.ofList)
        { Formula = formula; Numbers = numbers; Strings = strings; Count = count
          FormatCode = child "formatCode" cache |> Option.map innerText |> Option.filter (fun code -> code <> "")
          FollowedEdit = false; LimitToCachedPoints = false; Outer = outer }
    let reference (node: XmlDsl.XmlNode) =
        let formula = child "f" node |> Option.map innerText |> Option.filter (fun text -> text.Trim() <> "")
        match kids node |> List.tryFind (fun item -> (nameOf item).EndsWith "Cache") with
        | Some cache -> cached cache formula
        | None -> { Formula = formula; Numbers = Map.empty; Strings = Map.empty; Count = 0; FormatCode = None; FollowedEdit = false; LimitToCachedPoints = false; Outer = [] }
    kids holder
    |> List.tryPick (fun item ->
        match nameOf item with
        | "numRef" | "strRef" | "multiLvlStrRef" -> Some (reference item)
        | "numLit" | "strLit" -> Some (cached item None)
        | "v" -> Some { Formula = None; Numbers = Map.empty; Strings = Map.ofList [ 0, innerText item ]; Count = 1; FormatCode = None; FollowedEdit = false; LimitToCachedPoints = false; Outer = [] }
        | _ -> None)

let private labelSpecOf (theme: ThemeColors) (labels: XmlDsl.XmlNode) (parent: LabelSpec option) : LabelSpec =
    let flag (name: string) (inherited: LabelSpec -> bool) =
        match valOf name labels with
        | Some _ -> isOn (valOf name labels) false
        | None -> parent |> Option.map inherited |> Option.defaultValue false
    let offset =
        path [ "layout"; "manualLayout" ] labels
        |> Option.bind (fun manual ->
            match valOf "x" manual |> Option.bind number, valOf "y" manual |> Option.bind number with
            | None, None -> None
            | x, y -> Some (defaultArg x 0.0, defaultArg y 0.0))
    { ShowValue = flag "showVal" (fun spec -> spec.ShowValue)
      ShowPercent = flag "showPercent" (fun spec -> spec.ShowPercent)
      ShowCategory = flag "showCatName" (fun spec -> spec.ShowCategory)
      ShowSeries = flag "showSerName" (fun spec -> spec.ShowSeries)
      Position = valOf "dLblPos" labels |> Option.orElse (parent |> Option.bind (fun spec -> spec.Position))
      FormatCode =
        (match child "numFmt" labels with
         | Some format -> if isOn (attr "sourceLinked" format) false then None else attr "formatCode" format
         | None -> parent |> Option.bind (fun spec -> spec.FormatCode))
      Style = over (textPropsOf theme labels) (parent |> Option.map (fun spec -> spec.Style) |> Option.defaultValue noTextStyle)
      Separator = child "separator" labels |> Option.map innerText |> Option.orElse (parent |> Option.bind (fun spec -> spec.Separator))
      Deleted = Set.empty
      Points = Map.empty
      Offset = offset }

let private labelsOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : LabelSpec option =
    child "dLbls" holder
    |> Option.bind (fun labels ->
        if isOn (valOf "delete" labels) false then None
        else
            let spec = labelSpecOf theme labels None
            let pointLabels = childrenNamed "dLbl" labels
            let deleted =
                pointLabels
                |> List.choose (fun label -> if isOn (valOf "delete" label) false then valOf "idx" label |> Option.bind number |> Option.map int else None)
                |> Set.ofList
            let points =
                pointLabels
                |> List.choose (fun label ->
                    match valOf "idx" label |> Option.bind number with
                    | Some index when not (isOn (valOf "delete" label) false) -> Some (int index, labelSpecOf theme label (Some spec))
                    | _ -> None)
                |> Map.ofList
            let spec = { spec with Deleted = deleted; Points = points }
            let shows (spec: LabelSpec) = spec.ShowValue || spec.ShowPercent || spec.ShowCategory || spec.ShowSeries
            if shows spec || (points |> Map.exists (fun _ point -> shows point)) then Some spec else None)

/// `c:layout` of a title, a legend or the plot area.
let private layoutOf (holder: XmlDsl.XmlNode) : ManualLayout option =
    path [ "layout"; "manualLayout" ] holder
    |> Option.map (fun manual ->
        let value (name: string) = valOf name manual |> Option.bind number
        { Inner = valOf "layoutTarget" manual = Some "inner"
          XEdge = valOf "xMode" manual = Some "edge"
          YEdge = valOf "yMode" manual = Some "edge"
          X = value "x"; Y = value "y"; W = value "w"; H = value "h" })

let private seriesOf (theme: ThemeColors) (node: XmlDsl.XmlNode) : Series =
    let shape = child "spPr" node
    let marker = child "marker" node
    let markerShape = marker |> Option.bind (child "spPr")
    { Order = valOf "order" node |> Option.bind number |> Option.map int |> Option.defaultValue 0
      Index = valOf "idx" node |> Option.bind number |> Option.map int |> Option.defaultValue 0
      Name = child "tx" node |> Option.bind dataRefOf
      Categories = child "cat" node |> Option.bind dataRefOf
      Values = (match child "val" node with Some values -> dataRefOf values | None -> child "yVal" node |> Option.bind dataRefOf)
      XValues = child "xVal" node |> Option.bind dataRefOf
      Fill = shape |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
      Line = shape |> Option.bind (strokeOf theme)
      MarkerSymbol = marker |> Option.bind (valOf "symbol")
      // (0: the chart style's - `markerOf`.)
      MarkerSize = marker |> Option.bind (valOf "size") |> Option.bind number |> Option.defaultValue 0.0
      MarkerFill = markerShape |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
      MarkerLine = markerShape |> Option.bind (strokeOf theme)
      // (A line or scatter series that does not say is smoothed: Excel 2019 draws synthetic-chart-looks' "Markers" as curves,
      // Series.Smooth True.)
      Smooth = isOn (valOf "smooth" node) true
      PointFills =
        childrenNamed "dPt" node
        |> List.choose (fun point ->
            match valOf "idx" point |> Option.bind number, child "spPr" point |> Option.map (shapeFillOf theme) with
            | Some index, Some fill when fill <> FillAuto -> Some (int index, fill)
            | _ -> None)
        |> Map.ofList
      PointLines =
        childrenNamed "dPt" node
        |> List.choose (fun point ->
            match valOf "idx" point |> Option.bind number, child "spPr" point with
            | Some index, Some shape -> Some (int index, strokeOf theme shape)
            | _ -> None)
        |> Map.ofList
      Labels = labelsOf theme node
      BubbleSizes = child "bubbleSize" node |> Option.bind dataRefOf
      Explosion = valOf "explosion" node |> Option.bind number |> Option.defaultValue 0.0 }

let private plotOf (theme: ThemeColors) (node: XmlDsl.XmlNode) : Plot option =
    let kind =
        match nameOf node with
        // (A 3-D chart is drawn flat: the same bars, lines and slices without their depth.)
        | "barChart" | "bar3DChart" -> Some BarPlot
        | "lineChart" | "line3DChart" -> Some LinePlot
        | "pieChart" | "pie3DChart" | "ofPieChart" -> Some PiePlot
        | "doughnutChart" -> Some DoughnutPlot
        | "areaChart" | "area3DChart" -> Some AreaPlot
        | "scatterChart" -> Some ScatterPlot
        | "radarChart" -> Some RadarPlot
        | "bubbleChart" -> Some BubblePlot
        | "stockChart" -> Some StockPlot
        | _ -> None
    kind
    |> Option.map (fun kind ->
        { Kind = kind
          Horizontal = valOf "barDir" node = Some "bar"
          Grouping = valOf "grouping" node |> Option.defaultValue (if kind = BarPlot then "clustered" else "standard")
          Series = childrenNamed "ser" node |> List.map (seriesOf theme) |> List.sortBy (fun series -> series.Order)
          GapWidth = valOf "gapWidth" node |> Option.bind number |> Option.defaultValue 150.0
          Overlap = valOf "overlap" node |> Option.bind number |> Option.defaultValue 0.0
          HoleSize = valOf "holeSize" node |> Option.bind number |> Option.defaultValue (if kind = DoughnutPlot then 50.0 else 0.0)
          FirstSliceAngle = valOf "firstSliceAng" node |> Option.bind number |> Option.defaultValue 0.0
          VaryColors = isOn (valOf "varyColors" node) (kind = PiePlot || kind = DoughnutPlot)
          AxisIds = childrenNamed "axId" node |> List.choose (attr "val")
          Labels = labelsOf theme node
          Markers = isOn (valOf "marker" node) false
          ScatterStyle = valOf "scatterStyle" node |> Option.defaultValue "lineMarker"
          RadarStyle = valOf "radarStyle" node |> Option.defaultValue "standard"
          HiLowLines =
            child "hiLowLines" node
            |> Option.map (fun lines ->
                match child "spPr" lines |> Option.bind (strokeOf theme) with
                | Some stroke -> stroke
                | None -> { Paint = FillColor ("#595959", 1.0); Width = 1.0; Dash = None })
          UpDownBars =
            child "upDownBars" node
            |> Option.map (fun bars ->
                let fillAt (name: string) (otherwise: string) =
                    match child name bars |> Option.bind (child "spPr") |> Option.map (shapeFillOf theme) with
                    | Some (FillColor (color, opacity)) -> FillColor (color, opacity)
                    | Some FillNone -> FillNone
                    | _ -> FillColor (otherwise, 1.0)
                fillAt "upBars" "#FFFFFF", fillAt "downBars" "#404040")
          BubbleScale = valOf "bubbleScale" node |> Option.bind number |> Option.defaultValue 100.0 })

/// The text of a title (`c:title`): its paragraphs, one per line.
let private titleText (title: XmlDsl.XmlNode) : string option =
    match path [ "tx"; "rich" ] title with
    | Some rich ->
        let lines = childrenNamed "p" rich |> List.map (fun paragraph -> descendants "t" paragraph |> List.map innerText |> String.concat "")
        let text = String.concat "\n" lines
        if text.Trim() = "" then None else Some text
    | None ->
        child "tx" title |> Option.bind dataRefOf |> Option.bind (fun source -> Map.tryFind 0 source.Strings) |> Option.filter (fun text -> text.Trim() <> "")

let private titleStyle (theme: ThemeColors) (title: XmlDsl.XmlNode) : TextStyle =
    let fromRich =
        match path [ "tx"; "rich" ] title with
        | Some rich ->
            let paragraph = child "p" rich
            let run = paragraph |> Option.bind (child "r") |> Option.bind (child "rPr") |> Option.map (runStyle theme) |> Option.defaultValue noTextStyle
            let defaults = paragraph |> Option.bind (path [ "pPr"; "defRPr" ]) |> Option.map (runStyle theme) |> Option.defaultValue noTextStyle
            over run defaults
        | None -> noTextStyle
    over fromRich (textPropsOf theme title)

/// The runs of a title of rich text (`c:tx` `c:rich`), a list per paragraph: each run's look over its paragraph's
/// (`a:defRPr`) and the title's (`c:txPr`); empty for a title of a cell or of no text.
let private titleRuns (theme: ThemeColors) (title: XmlDsl.XmlNode) : (string * TextStyle) list list =
    match path [ "tx"; "rich" ] title with
    | Some rich ->
        let outer = textPropsOf theme title
        childrenNamed "p" rich
        |> List.map (fun paragraph ->
            let defaults = paragraph |> path [ "pPr"; "defRPr" ] |> Option.map (runStyle theme) |> Option.defaultValue noTextStyle
            kids paragraph
            |> List.choose (fun run ->
                match nameOf run with
                | "r" | "fld" ->
                    let own = child "rPr" run |> Option.map (runStyle theme) |> Option.defaultValue noTextStyle
                    Some ((child "t" run |> Option.map innerText |> Option.defaultValue ""), over own (over defaults outer))
                | _ -> None))
    | None -> []

let private axisOf (theme: ThemeColors) (plain: string) (node: XmlDsl.XmlNode) : Axis option =
    let kind =
        match nameOf node with
        | "catAx" -> Some "cat"
        | "valAx" -> Some "val"
        | "dateAx" -> Some "date"
        | "serAx" -> Some "ser"
        | _ -> None
    kind
    |> Option.map (fun kind ->
        let scaling = child "scaling" node
        let format = child "numFmt" node
        { Id = valOf "axId" node |> Option.defaultValue ""
          Kind = kind
          Position = valOf "axPos" node |> Option.defaultValue (if kind = "val" then "l" else "b")
          Deleted = isOn (valOf "delete" node) false
          Min = scaling |> Option.bind (valOf "min") |> Option.bind number
          Max = scaling |> Option.bind (valOf "max") |> Option.bind number
          MajorUnit = valOf "majorUnit" node |> Option.bind number
          Reversed = (scaling |> Option.bind (valOf "orientation")) = Some "maxMin"
          FormatCode = format |> Option.bind (attr "formatCode")
          SourceLinked = format |> Option.map (fun item -> isOn (attr "sourceLinked" item) false) |> Option.defaultValue true
          Gridlines =
            child "majorGridlines" node
            |> Option.map (fun lines ->
                match child "spPr" lines |> Option.bind (strokeOf theme) with
                | Some stroke -> stroke
                // (Without a line of their own, Excel 2019 draws them black: the files it writes
                // always give one, as light grey. `plain`: a part of a chart style's, below.)
                | None -> { Paint = FillColor (plain, 1.0); Width = 1.0; Dash = None })
          Title = child "title" node |> Option.bind titleText
          TitleStyle = child "title" node |> Option.map (titleStyle theme) |> Option.defaultValue noTextStyle
          TitleRuns = child "title" node |> Option.map (titleRuns theme) |> Option.defaultValue []
          LabelStyle = textPropsOf theme node
          // (An axis without a look of its own is drawn as a black line, as Excel does.)
          Line =
            match child "spPr" node with
            | Some shape -> strokeOf theme shape
            | None -> Some { Paint = FillColor (plain, 1.0); Width = 1.0; Dash = None }
          LabelsShown = valOf "tickLblPos" node <> Some "none"
          CrossesId = valOf "crossAx" node
          OnTicks = valOf "crossBetween" node = Some "midCat"
          LogBase = scaling |> Option.bind (valOf "logBase") |> Option.bind number |> Option.filter (fun value -> value > 1.0)
          DisplayUnit =
            (match child "dispUnits" node with
             | Some units ->
                (match valOf "builtInUnit" units, valOf "custUnit" units |> Option.bind number with
                 | Some "hundreds", _ -> 1e2 | Some "thousands", _ -> 1e3 | Some "tenThousands", _ -> 1e4 | Some "hundredThousands", _ -> 1e5
                 | Some "millions", _ -> 1e6 | Some "tenMillions", _ -> 1e7 | Some "hundredMillions", _ -> 1e8 | Some "billions", _ -> 1e9 | Some "trillions", _ -> 1e12
                 | _, Some custom when custom > 0.0 -> custom
                 | _ -> 1.0)
             | None -> 1.0)
          DisplayUnitLabel =
            child "dispUnits" node
            |> Option.bind (fun units ->
                child "dispUnitsLbl" units
                |> Option.map (fun label ->
                    match titleText label, valOf "builtInUnit" units with
                    | Some text, _ -> text
                    | None, Some "hundreds" -> "Hundreds" | None, Some "thousands" -> "Thousands" | None, Some "tenThousands" -> "x 10000" | None, Some "hundredThousands" -> "x 100000"
                    | None, Some "millions" -> "Millions" | None, Some "tenMillions" -> "x 10000000" | None, Some "hundredMillions" -> "x 100000000"
                    | None, Some "billions" -> "Billions" | None, Some "trillions" -> "Trillions"
                    | None, _ -> "")
                |> Option.filter (fun text -> text <> ""))
          BaseTimeUnit = valOf "baseTimeUnit" node
          MajorTimeUnit = valOf "majorTimeUnit" node
          LabelRotation =
            child "txPr" node |> Option.bind (child "bodyPr") |> Option.bind (attr "rot") |> Option.bind number
            |> Option.filter (fun rot -> rot > -60000000.0) |> Option.map (fun rot -> rot / 60000.0)
          LabelSkip = valOf "tickLblSkip" node |> Option.bind number |> Option.map int |> Option.filter (fun skip -> skip >= 1) })

// ---------------------------------------------------------------------------
// The newer kinds (`cx:chartSpace` of a chartEx part)
// ---------------------------------------------------------------------------

/// What a kind of chartEx series (`layoutId`) is called.
let chartExKindName (layoutId: string) : string =
    match layoutId with
    | "regionMap" -> "Map chart"
    | "sunburst" -> "Sunburst chart"
    | "treemap" -> "Treemap chart"
    | "boxWhisker" -> "Box and whisker chart"
    | "clusteredColumn" -> "Histogram"
    | "paretoLine" -> "Pareto chart"
    | "waterfall" -> "Waterfall chart"
    | "funnel" -> "Funnel chart"
    | "" -> "Chart"
    | other -> other + " chart"

let private chartExSpace = "http://schemas.microsoft.com/office/drawing/2014/chartex"

/// Whether a chart part is one of the newer kinds (`cx:chartSpace`: waterfall, funnel, treemap,
/// sunburst, histogram, box and whisker, map), told by its root's start tag - such a part can hold
/// megabytes of data (a map chart's geography), and is not read to find that out.
let isChartEx (xml: string) : bool =
    match (try XmlDsl.rootStartTag xml with _ -> None) with
    | Some (XmlDsl.Element (name, attrs, _, _), _) ->
        let colon = name.IndexOf ':'
        let declaration = if colon < 0 then "xmlns" else "xmlns:" + name.Substring(0, colon)
        localName name = "chartSpace" && Map.tryFind declaration attrs = Some chartExSpace
    | _ -> false

/// The start of a chartEx part as far as a chart that is not drawn needs it, read with the lexer
/// without going on into the rest (a map's geography): the kind of its first series
/// (`layoutId`), the chart's title element and that series' name (`cx:tx`).
let private chartExOutline (xml: string) : string option * XmlDsl.XmlNode option * XmlDsl.XmlNode option =
    let lexer = XmlDsl.XmlLexer(xml, 0, xml.Length)
    let names = Collections.Generic.HashSet<string>()
    let mutable kind = None
    let mutable title = None
    let mutable seriesName = None
    let mutable plotSeen = false
    let mutable going = true
    while going do
        match lexer.Next () with
        | XmlDsl.XmlToken.StartTag ->
            if kind.IsNone then
                if lexer.LocalNameIs "plotArea" then plotSeen <- true
                elif lexer.LocalNameIs "title" && not plotSeen && title.IsNone then title <- XmlDsl.tryReadElement lexer names
                elif lexer.LocalNameIs "series" then
                    let layoutId = lexer.Attribute "layoutId"
                    kind <- Some (if isNull layoutId then "" else layoutId)
                    if lexer.SelfClosing then going <- false
            elif lexer.LocalNameIs "tx" && seriesName.IsNone then
                seriesName <- XmlDsl.tryReadElement lexer names
                going <- false
            // (What follows a series' name - its look, its data, its geography - is not needed.)
            else going <- false
        | XmlDsl.XmlToken.EndTag when kind.IsSome -> going <- false
        | XmlDsl.XmlToken.EndOfText | XmlDsl.XmlToken.Malformed -> going <- false
        | _ -> ()
    kind, title, seriesName

/// The text of a chartEx title or series name (`cx:tx`: `cx:txData` with its cached `cx:v`, or `cx:rich`).
let private chartExText (tx: XmlDsl.XmlNode) : string option =
    match child "rich" tx with
    | Some rich ->
        let lines = childrenNamed "p" rich |> List.map (fun paragraph -> descendants "t" paragraph |> List.map innerText |> String.concat "")
        Some (String.concat "\n" lines)
    | None -> child "txData" tx |> Option.bind (child "v") |> Option.map innerText
    |> Option.filter (fun text -> text.Trim() <> "")

/// What a chartEx part's data dimension reads (`cx:strDim` / `cx:numDim`): its formula - a defined
/// name (`_xlchart.v1.0`) is what `names` says it refers to - and the points cached in `cx:lvl`.
let private chartExDim (names: string -> string option) (dim: XmlDsl.XmlNode) : DataRef =
    let formula =
        child "f" dim |> Option.map innerText |> Option.map (fun text -> text.Trim()) |> Option.filter (fun text -> text <> "")
        |> Option.bind (fun text -> if text.Contains "!" then Some text else names text)
        |> Option.map (fun text -> text.TrimStart('='))
    // (Levels of categories - a treemap's, a sunburst's -: the first names the points, the others are the levels
    // above it, from the inner out.)
    let levels = childrenNamed "lvl" dim
    let level = levels |> List.tryHead
    let outer =
        levels |> List.skip (min 1 levels.Length)
        |> List.map (fun level ->
            childrenNamed "pt" level
            |> List.choose (fun point -> attr "idx" point |> Option.bind number |> Option.map (fun index -> int index, innerText point))
            |> Map.ofList)
    let mutable numbers = Map.empty<int, float>
    let mutable strings = Map.empty<int, string>
    let mutable count = 0
    match level with
    | Some level ->
        for point in childrenNamed "pt" level do
            match attr "idx" point |> Option.bind number with
            | Some index ->
                let text = innerText point
                strings <- Map.add (int index) text strings
                (match number text with Some value -> numbers <- Map.add (int index) value numbers | None -> ())
                count <- max count (int index + 1)
            | None -> ()
        (match attr "ptCount" level |> Option.bind number with Some stated -> count <- max count (int stated) | None -> ())
    | None -> ()
    { Formula = formula; Numbers = numbers; Strings = strings; Count = count
      FormatCode = level |> Option.bind (attr "formatCode") |> Option.filter (fun code -> code <> "")
      FollowedEdit = false; LimitToCachedPoints = false; Outer = outer }

/// A chartEx chart of a kind that is not drawn: a placeholder that says what it is, with its title
/// (or, where the title is the automatic one, its series' name).
let private undrawnChartEx (layoutId: string) (title: XmlDsl.XmlNode option) (seriesName: XmlDsl.XmlNode option) : Chart =
    let titleText =
        match title |> Option.bind (child "tx") |> Option.bind chartExText with
        | Some text -> Some text
        | None -> seriesName |> Option.bind chartExText
    { Title = titleText; TitleStyle = noTextStyle; TitleRuns = []; AutoTitleDeleted = titleText.IsNone
      Plots =
        [ { Kind = UndrawnPlot layoutId; Horizontal = false; Grouping = "standard"; Series = []; GapWidth = 0.0; Overlap = 0.0; HoleSize = 0.0
            FirstSliceAngle = 0.0; VaryColors = false; AxisIds = []; Labels = None; Markers = false; ScatterStyle = ""; RadarStyle = ""
            HiLowLines = None; UpDownBars = None; BubbleScale = 100.0 } ]
      Axes = []; Legend = None; Background = FillAuto; Border = None; PlotBackground = FillNone; PlotBorder = None; TextStyle = noTextStyle
      RoundedCorners = false; Date1904 = false; TitleLayout = None; TitleOverlay = false; PlotLayout = None; Style = 2; UserShapes = None; ChartEx = true }

/// A chartEx part read: a waterfall or a funnel as a chart of one plot, any other kind a chart that
/// is drawn as a placeholder. `names` resolves the defined names its data refers to.
let private parseChartEx (names: string -> string option) (theme: ThemeColors) (xml: string) : Chart option =
    match chartExOutline xml with
    | None, _, _ -> None
    | Some layoutId, title, seriesName when not (List.contains layoutId [ "waterfall"; "funnel"; "clusteredColumn"; "boxWhisker"; "treemap"; "sunburst" ]) ->
        Some (undrawnChartEx layoutId title seriesName)
    | Some layoutId, title, seriesName ->
        let root = XmlDsl.parse (spacesKept xml)
        match child "chart" root with
        | None -> None
        | Some chart ->
            let dataSets =
                child "chartData" root |> Option.map (childrenNamed "data") |> Option.defaultValue []
                |> List.map (fun data -> attr "id" data |> Option.defaultValue "0", data)
            let area = child "plotArea" chart
            let region = area |> Option.bind (child "plotAreaRegion")
            let seriesNodes = region |> Option.map (childrenNamed "series") |> Option.defaultValue [] |> List.filter (fun node -> attr "layoutId" node = Some layoutId)
            let series =
                seriesNodes
                |> List.mapi (fun order node ->
                    let data = valOf "dataId" node |> Option.bind (fun id -> dataSets |> List.tryFind (fun (known, _) -> known = id)) |> Option.map snd
                    let dim (element: string) (kinds: string list) =
                        data |> Option.bind (fun data -> childrenNamed element data |> List.tryFind (fun dim -> List.contains (defaultArg (attr "type" dim) "") kinds))
                        |> Option.map (chartExDim names)
                    let labels =
                        child "dataLabels" node
                        |> Option.map (fun labels ->
                            let visibility = child "visibility" labels
                            let shown (name: string) (otherwise: bool) = isOn (visibility |> Option.bind (attr name)) otherwise
                            { ShowValue = shown "value" true
                              ShowPercent = false
                              ShowCategory = shown "categoryName" false
                              ShowSeries = shown "seriesName" false
                              Position = attr "pos" labels
                              FormatCode = child "numFmt" labels |> Option.bind (fun format -> if isOn (attr "sourceLinked" format) false then None else attr "formatCode" format)
                              Style = textPropsOf theme labels
                              Separator = child "separator" labels |> Option.map innerText
                              Deleted = childrenNamed "dataLabelHidden" labels |> List.choose (fun hidden -> attr "idx" hidden |> Option.bind number |> Option.map int) |> Set.ofList
                              Points = Map.empty; Offset = None })
                        |> Option.filter (fun spec -> spec.ShowValue || spec.ShowCategory || spec.ShowSeries)
                    let shape = child "spPr" node
                    { Order = order; Index = order
                      Name =
                        child "tx" node |> Option.bind (child "txData")
                        |> Option.map (fun txData ->
                            let reference = chartExDim names txData
                            match child "v" txData |> Option.map innerText with
                            | Some text -> { reference with Strings = Map.ofList [ 0, text ]; Count = max 1 reference.Count }
                            | None -> reference)
                      Categories = dim "strDim" [ "cat" ] |> Option.orElse (dim "numDim" [ "cat" ])
                      Values = dim "numDim" [ "val"; "size"; "" ]
                      XValues = None
                      Fill = shape |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                      Line = shape |> Option.bind (strokeOf theme)
                      MarkerSymbol = None; MarkerSize = 0.0; MarkerFill = FillAuto; MarkerLine = None; Smooth = false
                      PointFills =
                        childrenNamed "dataPt" node
                        |> List.choose (fun point ->
                            match attr "idx" point |> Option.bind number, child "spPr" point |> Option.map (shapeFillOf theme) with
                            | Some index, Some fill when fill <> FillAuto -> Some (int index, fill)
                            | _ -> None)
                        |> Map.ofList
                      PointLines =
                        childrenNamed "dataPt" node
                        |> List.choose (fun point ->
                            match attr "idx" point |> Option.bind number, child "spPr" point with
                            | Some index, Some shape -> Some (int index, strokeOf theme shape)
                            | _ -> None)
                        |> Map.ofList
                      Labels = labels
                      BubbleSizes = None
                      Explosion = 0.0 })
            let first = seriesNodes |> List.tryHead
            let subtotals =
                first |> Option.bind (path [ "layoutPr"; "subtotals" ])
                |> Option.map (fun totals -> childrenNamed "idx" totals |> List.choose (attr "val" >> Option.bind number >> Option.map int))
                |> Option.defaultValue []
            let connectors = first |> Option.bind (path [ "layoutPr"; "visibility" ]) |> Option.bind (attr "connectorLines") |> fun value -> isOn value true
            let kind =
                match layoutId with
                | "funnel" -> FunnelPlot
                | "clusteredColumn" ->
                    let binning = first |> Option.bind (path [ "layoutPr"; "binning" ])
                    let bound (name: string) = binning |> Option.bind (attr name) |> Option.bind number
                    HistogramPlot
                        ({ ClosedLeft = (binning |> Option.bind (attr "intervalClosed")) = Some "l"
                           Size = binning |> Option.bind (child "binSize") |> Option.bind (attr "val") |> Option.bind number |> Option.filter (fun size -> size > 0.0)
                           Count = binning |> Option.bind (child "binCount") |> Option.bind (attr "val") |> Option.bind number |> Option.map int |> Option.filter (fun count -> count > 0)
                           Underflow = bound "underflow"; Overflow = bound "overflow" },
                         region |> Option.map (childrenNamed "series") |> Option.defaultValue [] |> List.exists (fun node -> attr "layoutId" node = Some "paretoLine"))
                | "boxWhisker" ->
                    let layout = first |> Option.bind (child "layoutPr")
                    let visibility = layout |> Option.bind (child "visibility")
                    // (What a series that does not say shows: Excel 2019's boxes made through COM - the mean's mark,
                    // the points inside the whiskers and outside them, no line through the means.)
                    let shown (name: string) (otherwise: bool) = isOn (visibility |> Option.bind (attr name)) otherwise
                    BoxWhiskerPlot
                        { Inclusive = (layout |> Option.bind (child "statistics") |> Option.bind (attr "quartileMethod")) = Some "inclusive"
                          MeanLine = shown "meanLine" false; MeanMarker = shown "meanMarker" true
                          InnerPoints = shown "nonoutliers" true; Outliers = shown "outliers" true }
                | "treemap" ->
                    TreemapPlot (first |> Option.bind (path [ "layoutPr"; "parentLabelLayout" ]) |> Option.bind (attr "val") |> Option.defaultValue "overlapping")
                | "sunburst" -> SunburstPlot
                | _ -> WaterfallPlot (subtotals, connectors)
            let axes =
                area |> Option.map (childrenNamed "axis") |> Option.defaultValue []
                |> List.map (fun axis ->
                    let isValue = (child "valScaling" axis).IsSome
                    let scaling = child "valScaling" axis |> Option.orElse (child "catScaling" axis)
                    let scaled (name: string) = scaling |> Option.bind (attr name) |> Option.bind number
                    let format = child "numFmt" axis
                    { Id = attr "id" axis |> Option.defaultValue ""
                      Kind = (if isValue then "val" else "cat")
                      OnTicks = false
                      // (A Pareto chart's second value axis - its running share - at the right.)
                      Position = (match layoutId, isValue with "funnel", false -> "l" | "funnel", true -> "b" | _, true when attr "id" axis = Some "2" -> "r" | _, true -> "l" | _ -> "b")
                      Deleted = isOn (attr "hidden" axis) false
                      Min = scaled "min"; Max = scaled "max"; MajorUnit = scaled "majorUnit"
                      Reversed = false
                      // (`cx:units unit="percentage"`: a Pareto chart's running share, 0 % .. 100 %.)
                      FormatCode = format |> Option.bind (attr "formatCode") |> Option.orElse (if (child "units" axis |> Option.bind (attr "unit")) = Some "percentage" then Some "0%" else None)
                      SourceLinked = format |> Option.map (fun item -> isOn (attr "sourceLinked" item) false) |> Option.defaultValue (child "units" axis |> Option.bind (attr "unit") <> Some "percentage")
                      Gridlines =
                        child "majorGridlines" axis
                        |> Option.map (fun lines ->
                            match child "spPr" lines |> Option.bind (strokeOf theme) with
                            | Some stroke -> stroke
                            | None -> { Paint = FillColor ("#D9D9D9", 1.0); Width = 1.0; Dash = None })
                      Title = child "title" axis |> Option.bind (child "tx") |> Option.bind chartExText
                      TitleStyle = noTextStyle
                      TitleRuns = []
                      LabelStyle = textPropsOf theme axis
                      Line = child "spPr" axis |> Option.bind (strokeOf theme)
                      LabelsShown = (child "tickLabels" axis).IsSome
                      CrossesId = None
                      LogBase = None
                      DisplayUnit = 1.0
                      DisplayUnitLabel = None
                      BaseTimeUnit = None; MajorTimeUnit = None; LabelRotation = None; LabelSkip = None })
            let titleNode = child "title" chart
            let titleText =
                match titleNode |> Option.bind (child "tx") |> Option.bind chartExText with
                | Some text -> Some text
                // (A title that says nothing is "Chart Title": Excel 2019's chartEx charts made through COM.)
                | None when titleNode.IsSome -> Some "Chart Title"
                | None -> None
            let shape = child "spPr" root
            let gap =
                area |> Option.map (childrenNamed "axis") |> Option.defaultValue []
                |> List.tryPick (fun axis -> child "catScaling" axis |> Option.bind (attr "gapWidth") |> Option.bind number)
                |> Option.defaultValue (match layoutId with "funnel" -> 0.06 | "clusteredColumn" -> 0.0 | "boxWhisker" -> 1.0 | _ -> 0.5)
            Some
                { Title = titleText
                  TitleStyle = titleNode |> Option.map (titleStyle theme) |> Option.defaultValue noTextStyle
                  TitleRuns = []
                  AutoTitleDeleted = titleText.IsNone
                  Plots =
                    [ { Kind = kind
                        Horizontal = (layoutId = "funnel"); Grouping = "standard"; Series = series
                        GapWidth = gap * 100.0; Overlap = 0.0; HoleSize = 0.0; FirstSliceAngle = 0.0; VaryColors = false
                        AxisIds = axes |> List.map (fun axis -> axis.Id); Labels = None; Markers = false; ScatterStyle = ""; RadarStyle = ""
                        HiLowLines = None; UpDownBars = None; BubbleScale = 100.0 } ]
                  Axes = axes
                  Legend =
                    child "legend" chart
                    |> Option.map (fun legend ->
                        { Position = attr "pos" legend |> Option.defaultValue "t"; Style = textPropsOf theme legend; Layout = None
                          Overlay = isOn (attr "overlay" legend) false; Deleted = Set.empty; EntryStyles = Map.empty; Fill = FillNone; Border = None })
                  Background = shape |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                  // (A chart that states no look of its own: the frame of Excel 2019's chartEx style, #D9D9D9.)
                  Border = (match shape with Some shape -> strokeOf theme shape | None -> Some { Paint = FillColor ("#D9D9D9", 1.0); Width = 1.0; Dash = None })
                  PlotBackground = area |> Option.bind (child "spPr") |> Option.map (shapeFillOf theme) |> Option.defaultValue FillNone
                  PlotBorder = None
                  TextStyle = textPropsOf theme root
                  RoundedCorners = false
                  Date1904 = false
                  TitleLayout = None
                  TitleOverlay = titleNode |> Option.map (fun node -> isOn (attr "overlay" node) false) |> Option.defaultValue false
                  PlotLayout = None
                  Style = 2
                  UserShapes = None
                  ChartEx = true }

/// A chart part of the classic kind (`c:chartSpace`) read; None for a part this module does not
/// draw (no plot of a kind it knows, a part that is not a chart).
let private parseClassicChart (theme: ThemeColors) (xml: string) : Chart option =
    try
        let root = XmlDsl.parse (spacesKept xml)
        let space = if nameOf root = "chartSpace" then Some root else descendants "chartSpace" root |> List.tryHead
        match space |> Option.bind (child "chart") with
        | None -> None
        | Some chart ->
            let space = space.Value
            match child "plotArea" chart with
            | None -> None
            | Some area ->
                let items = kids area |> List.collect flattened
                let known = items |> List.choose (plotOf theme)
                let unknown =
                    items |> List.exists (fun item ->
                        let name = nameOf item
                        name.EndsWith "Chart" && (plotOf theme item).IsNone)
                if known.IsEmpty || unknown then None
                else
                    // (A part that names a chart style - `c:style`, as Excel's own old API writes one: Charts.Add, ChartObjects.Add -
                    // and no look of its own: Excel 2019 draws its lines and frame #898989, its title 18 pt bold, its text 10 pt
                    // - legacy-chart-probe.vbs, synthetic-chartsheets.xlsx. One that names no style - openpyxl's - black lines.)
                    let styled = kids space |> List.collect flattened |> List.exists (fun item -> nameOf item = "style")
                    let plain = if styled then "#898989" else "#000000"
                    let shape = child "spPr" space
                    let pivot = (child "pivotSource" space).IsSome
                    let limit (reference: DataRef option) =
                        if pivot then reference |> Option.map (fun value -> { value with LimitToCachedPoints = true }) else reference
                    let plots =
                        if not pivot then known
                        else
                            known |> List.map (fun plot ->
                                let series =
                                    plot.Series |> List.map (fun series ->
                                        // (A pivot chart's series is named as the pivot table names its field - the file's
                                        // "Total", not the "Sum of Sales" of the cell: dashboard-ecommerce-sales' "Top 5".)
                                        { series with
                                            Name = series.Name |> Option.map (fun name -> if name.Count > 0 then { name with Formula = None } else name)
                                            Categories = limit series.Categories
                                            Values = limit series.Values; XValues = limit series.XValues
                                            BubbleSizes = limit series.BubbleSizes })
                                { plot with Series = series })
                    Some { Title = child "title" chart |> Option.bind titleText
                           TitleStyle =
                             child "title" chart |> Option.map (titleStyle theme) |> Option.defaultValue noTextStyle
                             |> fun style -> if styled && (child "txPr" space).IsNone then { style with Size = style.Size |> Option.orElse (Some 18.0) } else style
                           TitleRuns = child "title" chart |> Option.map (titleRuns theme) |> Option.defaultValue []
                           AutoTitleDeleted = isOn (valOf "autoTitleDeleted" chart) false || (child "title" chart).IsNone
                           Plots = plots
                           Axes = items |> List.choose (axisOf theme plain)
                           Legend =
                             child "legend" chart
                             |> Option.map (fun legend ->
                                { Position = valOf "legendPos" legend |> Option.defaultValue "r"; Style = textPropsOf theme legend
                                  // (A legend or a title that does not say is laid over the plot, as Excel 2019 does.)
                                  Layout = layoutOf legend; Overlay = isOn (valOf "overlay" legend) true
                                  Deleted =
                                    childrenNamed "legendEntry" legend
                                    |> List.choose (fun entry -> if isOn (valOf "delete" entry) false then valOf "idx" entry |> Option.bind number |> Option.map int else None)
                                    |> Set.ofList
                                  Fill = child "spPr" legend |> Option.map (shapeFillOf theme) |> Option.defaultValue FillNone
                                  Border = child "spPr" legend |> Option.bind (strokeOf theme)
                                  EntryStyles =
                                    childrenNamed "legendEntry" legend
                                    |> List.choose (fun entry ->
                                        match valOf "idx" entry |> Option.bind number, child "txPr" entry with
                                        | Some index, Some _ -> Some (int index, textPropsOf theme entry)
                                        | _ -> None)
                                    |> Map.ofList })
                           Background = shape |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                           // (A chart without a look of its own has Excel's grey frame.)
                           Border =
                             match shape with
                             | Some shape -> strokeOf theme shape
                             | None -> Some { Paint = FillColor ((if styled then "#898989" else "#868686"), 1.0); Width = 1.0; Dash = None }
                           PlotBackground = child "spPr" area |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                           PlotBorder = child "spPr" area |> Option.bind (strokeOf theme)
                           TextStyle = textPropsOf theme space
                           // (Rounded unless the file says not: Excel writes `roundedCorners val="0"`.)
                           RoundedCorners = isOn (valOf "roundedCorners" space) true
                           Date1904 = isOn (valOf "date1904" space) false
                           TitleLayout = child "title" chart |> Option.bind layoutOf
                           TitleOverlay = child "title" chart |> Option.map (fun title -> isOn (valOf "overlay" title) true) |> Option.defaultValue false
                           PlotLayout = layoutOf area
                           Style =
                             // (`c:style` of the part, or the one its `mc:AlternateContent` falls back to.)
                             kids space |> List.collect flattened |> List.tryPick (fun item -> if nameOf item = "style" then attr "val" item |> Option.bind number |> Option.map int else None)
                             |> Option.defaultValue 2
                           UserShapes = None
                           ChartEx = false }
    with _ -> None


/// `parseChart`, with what the defined names a chartEx part reads its data through refer to.
let parseChartWith (names: string -> string option) (theme: ThemeColors) (xml: string) : Chart option =
    if isChartEx xml then (try parseChartEx names theme xml with _ -> None)
    else parseClassicChart theme xml
/// Mark only references whose formulas changed since the imported chart. Cached point
/// counts still trim unchanged pivot-chart ranges (which include grand totals), while
/// inserted points and changed defined-name ranges use the full live range.
let markChangedReferences (original: Chart) (current: Chart) : Chart =
    let moved before after =
        match before, after with
        | Some oldRef, Some newRef when oldRef.Formula <> newRef.Formula ->
            Some { newRef with FollowedEdit = true }
        | _ -> after
    let series before after =
        { after with
            Name = moved before.Name after.Name
            Categories = moved before.Categories after.Categories
            Values = moved before.Values after.Values
            XValues = moved before.XValues after.XValues
            BubbleSizes = moved before.BubbleSizes after.BubbleSizes }
    let plot before after =
        if List.length before.Series <> List.length after.Series then after
        else { after with Series = List.map2 series before.Series after.Series }
    if List.length original.Plots <> List.length current.Plots then current
    else { current with Plots = List.map2 plot original.Plots current.Plots }

/// The chart of a chart part, or None for a part this module does not draw: a chartEx part reads
/// its data through defined names, which this does not resolve (`parseChartWith` does).
let parseChart (theme: ThemeColors) (xml: string) : Chart option = parseChartWith (fun _ -> None) theme xml

// ---------------------------------------------------------------------------
// The scene: what a chart or a shape is drawn with
// ---------------------------------------------------------------------------

type Primitive =
    | PRect of x: float * y: float * width: float * height: float * fill: string option * opacity: float * stroke: string option * strokeWidth: float * radius: float
    | PEllipse of cx: float * cy: float * rx: float * ry: float * fill: string option * opacity: float * stroke: string option * strokeWidth: float
    | PLine of x1: float * y1: float * x2: float * y2: float * stroke: string * strokeWidth: float * dash: string option
    /// Points joined by lines; closed and filled when `fill` is given.
    | PPolyline of points: (float * float) list * fill: string option * opacity: float * stroke: string option * strokeWidth: float * dash: string option
    /// An SVG path (`d`): the wedges of a pie.
    | PPath of d: string * fill: string option * opacity: float * stroke: string option * strokeWidth: float
    /// `anchor`: "start" | "middle" | "end"; `y` is the middle of the line; `rotate` in degrees around (x, y).
    | PText of x: float * y: float * text: string * size: float * color: string * anchor: string * bold: bool * italic: bool * rotate: float
    /// A line of text in runs, each (text, size, colour, bold, italic), one after the other from
    /// (x, y) - or centred on it, or ending at it (`anchor`).
    | PTextRuns of x: float * y: float * anchor: string * runs: (string * float * string * bool * bool) list
    /// A picture (a data URL) in a rectangle.
    | PImage of x: float * y: float * width: float * height: float * href: string
    /// A scene inside a rectangle of this one (a chart in a group), clipped to it.
    | PGroup of x: float * y: float * width: float * height: float * items: Primitive list
    /// Primitives under a transformation (a rotated or flipped shape): x' = a x + c y + e, y' = b x + d y + f.
    | PTransform of a: float * b: float * c: float * d: float * e: float * f: float * items: Primitive list
    /// Primitives whose text is in a typeface of its own (a chart's `a:latin`), not Calibri.
    | PFace of face: string * items: Primitive list

type Scene = {
    Width: float
    Height: float
    Items: Primitive list
}

let private fmt (value: float) : string = fmtNumber value

let private xmlText (text: string) : string =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

let private dashArray (dash: string option) (width: float) : string =
    match dash with
    | Some "dash" | Some "lgDash" | Some "sysDash" -> sprintf " stroke-dasharray=\"%s %s\"" (fmt (width * 4.0)) (fmt (width * 3.0))
    | Some "dot" | Some "sysDot" -> sprintf " stroke-dasharray=\"%s %s\"" (fmt width) (fmt (width * 2.0))
    | Some "dashDot" | Some "lgDashDot" | Some "sysDashDot" -> sprintf " stroke-dasharray=\"%s %s %s %s\"" (fmt (width * 4.0)) (fmt (width * 2.0)) (fmt width) (fmt (width * 2.0))
    | _ -> ""

let mutable private clipSerial = 0

/// The attributes of a fill and an outline; a gradient is defined in `defs` and named.
let private paint (defs: ResizeArray<string>) (fill: string option) (opacity: float) (stroke: string option) (strokeWidth: float) : string =
    let fillText =
        match fill with
        | None -> " fill=\"none\""
        | Some color ->
            match gradientOf color with
            | None -> sprintf " fill=\"%s\"" color
            | Some (angle, stops) ->
                clipSerial <- clipSerial + 1
                let id = sprintf "rbt-chart-gradient-%d" clipSerial
                let stopText =
                    stops
                    |> List.map (fun (offset, stopColor, stopOpacity) ->
                        sprintf "<stop offset=\"%s\" stop-color=\"%s\"%s/>" (fmt offset) stopColor (if stopOpacity < 0.999 then sprintf " stop-opacity=\"%s\"" (fmt stopOpacity) else ""))
                    |> String.concat ""
                (match angle with
                 | Some degrees ->
                    // (The line of the gradient through the middle of the shape's box, at its angle.)
                    let radians = degrees * Math.PI / 180.0
                    let dx, dy = Math.Cos radians / 2.0, Math.Sin radians / 2.0
                    defs.Add (sprintf "<linearGradient id=\"%s\" x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\">%s</linearGradient>" id (fmt (0.5 - dx)) (fmt (0.5 - dy)) (fmt (0.5 + dx)) (fmt (0.5 + dy)) stopText)
                 | None -> defs.Add (sprintf "<radialGradient id=\"%s\" cx=\"0.5\" cy=\"0.5\" r=\"0.71\">%s</radialGradient>" id stopText))
                sprintf " fill=\"url(#%s)\"" id
    fillText
    + (if opacity < 0.999 && fill.IsSome then sprintf " fill-opacity=\"%s\"" (fmt opacity) else "")
    + (match stroke with Some color when strokeWidth > 0.0 -> sprintf " stroke=\"%s\" stroke-width=\"%s\"" (solidOf color) (fmt strokeWidth) | _ -> "")

let rec private primitiveSvg (defs: ResizeArray<string>) (item: Primitive) : string =
    match item with
    | PRect (x, y, width, height, fill, opacity, stroke, strokeWidth, radius) ->
        sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\"%s%s/>" (fmt x) (fmt y) (fmt (max 0.0 width)) (fmt (max 0.0 height))
            (if radius > 0.0 then sprintf " rx=\"%s\"" (fmt radius) else "") (paint defs fill opacity stroke strokeWidth)
    | PEllipse (cx, cy, rx, ry, fill, opacity, stroke, strokeWidth) ->
        sprintf "<ellipse cx=\"%s\" cy=\"%s\" rx=\"%s\" ry=\"%s\"%s/>" (fmt cx) (fmt cy) (fmt (max 0.0 rx)) (fmt (max 0.0 ry)) (paint defs fill opacity stroke strokeWidth)
    | PLine (x1, y1, x2, y2, stroke, strokeWidth, dash) ->
        sprintf "<line x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\" stroke=\"%s\" stroke-width=\"%s\"%s/>" (fmt x1) (fmt y1) (fmt x2) (fmt y2) (solidOf stroke) (fmt strokeWidth) (dashArray dash strokeWidth)
    | PPolyline (points, fill, opacity, stroke, strokeWidth, dash) ->
        let text = points |> List.map (fun (x, y) -> fmt x + "," + fmt y) |> String.concat " "
        sprintf "<%s points=\"%s\"%s stroke-linejoin=\"round\" stroke-linecap=\"round\"%s/>" (if fill.IsSome then "polygon" else "polyline") text (paint defs fill opacity stroke strokeWidth) (dashArray dash strokeWidth)
    | PPath (d, fill, opacity, stroke, strokeWidth) ->
        sprintf "<path d=\"%s\"%s stroke-linejoin=\"round\"/>" d (paint defs fill opacity stroke strokeWidth)
    | PText (x, y, text, size, color, anchor, bold, italic, rotate) ->
        sprintf "<text x=\"%s\" y=\"%s\" font-size=\"%s\" fill=\"%s\" text-anchor=\"%s\" dominant-baseline=\"central\"%s%s%s>%s</text>"
            (fmt x) (fmt y) (fmt size) (solidOf color) anchor (if bold then " font-weight=\"bold\"" else "") (if italic then " font-style=\"italic\"" else "")
            (if rotate <> 0.0 then sprintf " transform=\"rotate(%s %s %s)\"" (fmt rotate) (fmt x) (fmt y) else "") (xmlText text)
    | PTextRuns (x, y, anchor, runs) ->
        let spans =
            runs
            |> List.map (fun (text, size, color, bold, italic) ->
                sprintf "<tspan font-size=\"%s\" fill=\"%s\"%s%s>%s</tspan>" (fmt size) (solidOf color) (if bold then " font-weight=\"bold\"" else "") (if italic then " font-style=\"italic\"" else "") (xmlText text))
            |> String.concat ""
        sprintf "<text x=\"%s\" y=\"%s\" text-anchor=\"%s\" dominant-baseline=\"central\" xml:space=\"preserve\" style=\"white-space:pre\">%s</text>" (fmt x) (fmt y) anchor spans
    | PImage (x, y, width, height, href) ->
        sprintf "<image x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" preserveAspectRatio=\"none\" href=\"%s\"/>" (fmt x) (fmt y) (fmt width) (fmt height) (xmlText href)
    | PGroup (x, y, width, height, items) ->
        clipSerial <- clipSerial + 1
        let id = sprintf "rbt-chart-clip-%d" clipSerial
        sprintf "<clipPath id=\"%s\"><rect x=\"0\" y=\"0\" width=\"%s\" height=\"%s\"/></clipPath><g transform=\"translate(%s %s)\" clip-path=\"url(#%s)\">%s</g>"
            id (fmt width) (fmt height) (fmt x) (fmt y) id (items |> List.map (primitiveSvg defs) |> String.concat "")
    | PTransform (a, b, c, d, e, f, items) ->
        // (Four decimals for the matrix: two would turn a shape by a visibly different angle.)
        let precise (value: float) = (Math.Round (value * 10000.0) / 10000.0).ToString(Globalization.CultureInfo.InvariantCulture)
        sprintf "<g transform=\"matrix(%s %s %s %s %s %s)\">%s</g>" (precise a) (precise b) (precise c) (precise d) (fmt e) (fmt f) (items |> List.map (primitiveSvg defs) |> String.concat "")
    | PFace (face, items) ->
        sprintf "<g font-family=\"%s, Calibri, Carlito, Arial, sans-serif\">%s</g>" (xmlText ("'" + face.Replace("'", "") + "'")) (items |> List.map (primitiveSvg defs) |> String.concat "")

/// A scene as an `<svg>` element of its size (the font is Calibri, as Excel's charts).
let sceneSvg (scene: Scene) : string =
    let defs = ResizeArray<string>()
    let body = scene.Items |> List.map (primitiveSvg defs) |> String.concat ""
    sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%s\" height=\"%s\" viewBox=\"0 0 %s %s\" font-family=\"Calibri, Carlito, Arial, sans-serif\" style=\"display:block;overflow:hidden\">%s%s</svg>"
        (fmt scene.Width) (fmt scene.Height) (fmt scene.Width) (fmt scene.Height) (if defs.Count = 0 then "" else "<defs>" + String.concat "" defs + "</defs>") body

// ---------------------------------------------------------------------------
// Paths: what the `d` of a `PPath` is for a renderer without SVG
// ---------------------------------------------------------------------------

/// A segment of a path, in absolute coordinates.
type PathSegment =
    | MoveTo of x: float * y: float
    | LineTo of x: float * y: float
    | CubicTo of x1: float * y1: float * x2: float * y2: float * x: float * y: float
    | ClosePath

/// An elliptical arc from (x0, y0) to (x, y) as cubic curves (the arc of an SVG path: radii,
/// the larger of the two arcs or not, clockwise or not; no rotation of the ellipse).
let private arcCubics (x0: float) (y0: float) (rx: float) (ry: float) (large: bool) (sweep: bool) (x: float) (y: float) : PathSegment list =
    if rx <= 0.0 || ry <= 0.0 || (x0 = x && y0 = y) then [ LineTo (x, y) ]
    else
        // (The centre, as SVG's implementation notes work it out.)
        let dx, dy = (x0 - x) / 2.0, (y0 - y) / 2.0
        let lambda = dx * dx / (rx * rx) + dy * dy / (ry * ry)
        let rx, ry = (if lambda > 1.0 then rx * Math.Sqrt lambda, ry * Math.Sqrt lambda else rx, ry)
        let numerator = rx * rx * ry * ry - rx * rx * dy * dy - ry * ry * dx * dx
        let denominator = rx * rx * dy * dy + ry * ry * dx * dx
        let factor = (if large = sweep then -1.0 else 1.0) * Math.Sqrt (max 0.0 (if denominator = 0.0 then 0.0 else numerator / denominator))
        let cxp, cyp = factor * rx * dy / ry, -factor * ry * dx / rx
        let cx, cy = cxp + (x0 + x) / 2.0, cyp + (y0 + y) / 2.0
        let angleOf (px: float) (py: float) = Math.Atan2 ((py - cy) / ry, (px - cx) / rx)
        let start = angleOf x0 y0
        let mutable delta = angleOf x y - start
        if sweep && delta < 0.0 then delta <- delta + 2.0 * Math.PI
        if not sweep && delta > 0.0 then delta <- delta - 2.0 * Math.PI
        let pieces = max 1 (int (Math.Ceiling (abs delta / (Math.PI / 2.0) - 1e-9)))
        let step = delta / float pieces
        let k = 4.0 / 3.0 * Math.Tan (step / 4.0)
        [ for index in 0 .. pieces - 1 do
            let a0 = start + step * float index
            let a1 = a0 + step
            let p0x, p0y = cx + rx * Math.Cos a0, cy + ry * Math.Sin a0
            let p1x, p1y = cx + rx * Math.Cos a1, cy + ry * Math.Sin a1
            yield CubicTo (p0x - k * rx * Math.Sin a0, p0y + k * ry * Math.Cos a0, p1x + k * rx * Math.Sin a1, p1y - k * ry * Math.Cos a1, p1x, p1y) ]

/// The segments of the `d` of a `PPath` (the commands this module writes: M, L, C, Q, A, Z -
/// absolute), arcs and quadratic curves as cubic ones: what a renderer without SVG draws.
let pathSegments (d: string) : PathSegment list =
    let tokens = ResizeArray<string>()
    let mutable at = 0
    while at < d.Length do
        let ch = d.[at]
        if Char.IsLetter ch && ch <> 'e' && ch <> 'E' then
            tokens.Add (string ch)
            at <- at + 1
        elif ch = ' ' || ch = ',' || ch = '\n' || ch = '\t' || ch = '\r' then at <- at + 1
        else
            let start = at
            at <- at + 1
            while at < d.Length && (Char.IsDigit d.[at] || d.[at] = '.' || d.[at] = 'e' || d.[at] = 'E' || ((d.[at] = '-' || d.[at] = '+') && (d.[at - 1] = 'e' || d.[at - 1] = 'E'))) do
                at <- at + 1
            tokens.Add (d.Substring(start, at - start))
    let result = ResizeArray<PathSegment>()
    let mutable index = 0
    let mutable command = ""
    let mutable curX, curY = 0.0, 0.0
    let mutable startX, startY = 0.0, 0.0
    let next () =
        let value = if index < tokens.Count then number tokens.[index] |> Option.defaultValue 0.0 else 0.0
        index <- index + 1
        value
    let mutable guard = 0
    while index < tokens.Count && guard < 100000 do
        guard <- guard + 1
        let token = tokens.[index]
        if token.Length = 1 && Char.IsLetter token.[0] then
            command <- token
            index <- index + 1
        match command with
        | "M" ->
            let x = next ()
            let y = next ()
            result.Add (MoveTo (x, y))
            curX <- x; curY <- y; startX <- x; startY <- y
            command <- "L"
        | "L" ->
            let x = next ()
            let y = next ()
            result.Add (LineTo (x, y))
            curX <- x; curY <- y
        | "C" ->
            let x1 = next ()
            let y1 = next ()
            let x2 = next ()
            let y2 = next ()
            let x = next ()
            let y = next ()
            result.Add (CubicTo (x1, y1, x2, y2, x, y))
            curX <- x; curY <- y
        | "Q" ->
            let x1 = next ()
            let y1 = next ()
            let x = next ()
            let y = next ()
            result.Add (CubicTo (curX + 2.0 / 3.0 * (x1 - curX), curY + 2.0 / 3.0 * (y1 - curY), x + 2.0 / 3.0 * (x1 - x), y + 2.0 / 3.0 * (y1 - y), x, y))
            curX <- x; curY <- y
        | "A" ->
            let rx = next ()
            let ry = next ()
            next () |> ignore
            let large = next () <> 0.0
            let sweep = next () <> 0.0
            let x = next ()
            let y = next ()
            for segment in arcCubics curX curY rx ry large sweep x y do result.Add segment
            curX <- x; curY <- y
        | "Z" | "z" ->
            result.Add ClosePath
            curX <- startX; curY <- startY
            command <- ""
        | _ -> index <- index + 1
    List.ofSeq result

// ---------------------------------------------------------------------------
// Data
// ---------------------------------------------------------------------------

/// What a chart reads of the workbook.
type ChartData = {
    /// The cells of a block of a sheet (by name), row by row: None when the sheet is not there.
    Cells: string -> int -> int -> int -> int -> CellData option [] [] option
    /// What a cell shows (its value in its number format).
    Display: CellData -> string
    /// A number in a number format.
    Format: string -> float -> string
}

/// `'My Sheet'!$A$1:$B$5` (also one cell, several areas in parentheses): sheet, first row / column, last row / column.
let private parseAreas (formula: string) : (string * int * int * int * int) list option =
    let text = formula.Trim()
    let text = if text.StartsWith "(" && text.EndsWith ")" then text.Substring(1, text.Length - 2) else text
    // (Commas inside a quoted sheet name do not separate areas.)
    let parts = ResizeArray<string>()
    let mutable quoted = false
    let mutable start = 0
    for at in 0 .. text.Length - 1 do
        if text.[at] = '\'' then quoted <- not quoted
        elif text.[at] = ',' && not quoted then
            parts.Add (text.Substring(start, at - start))
            start <- at + 1
    parts.Add (text.Substring start)
    let cellOf (reference: string) : (int * int) option =
        let reference = reference.Replace("$", "")
        let mutable at = 0
        let mutable col = 0
        while at < reference.Length && Char.IsLetter reference.[at] do
            col <- col * 26 + (int (Char.ToUpperInvariant reference.[at]) - int 'A' + 1)
            at <- at + 1
        match Int32.TryParse (reference.Substring at) with
        | true, row when col > 0 && row > 0 && at > 0 -> Some (row - 1, col - 1)
        | _ -> None
    let area (part: string) =
        let part = part.Trim()
        match part.LastIndexOf '!' with
        | -1 -> None
        | bang ->
            let sheet = part.Substring(0, bang)
            let sheet = if sheet.StartsWith "'" && sheet.EndsWith "'" && sheet.Length >= 2 then sheet.Substring(1, sheet.Length - 2).Replace("''", "'") else sheet
            let range = part.Substring(bang + 1)
            // (A second sheet name after the colon - `Sheet1!A1:Sheet1!A5` - is dropped.)
            let ends = range.Split(':') |> Array.map (fun item -> match item.LastIndexOf '!' with -1 -> item | cut -> item.Substring(cut + 1))
            match cellOf ends.[0], cellOf ends.[ends.Length - 1] with
            | Some (r0, c0), Some (r1, c1) when sheet <> "" && not (sheet.Contains "[") -> Some (sheet, min r0 r1, min c0 c1, max r0 r1, max c0 c1)
            | _ -> None
    let areas = parts |> Seq.map area |> Seq.toList
    if areas.IsEmpty || areas |> List.exists Option.isNone then None else Some (areas |> List.map Option.get)

/// The cells a reference names, one after the other (down a column, along a row; several areas in their order).
let private cellsOf (data: ChartData) (source: DataRef) : CellData option list option =
    // (Multi-level categories - a range of more than one column - are the file's levels: its cells are no list of
    // the points' names. phpsp-32chartreadwrite's years and quarters.)
    if not source.Outer.IsEmpty then None
    else
    source.Formula
    |> Option.bind parseAreas
    |> Option.bind (fun areas ->
        let blocks = areas |> List.map (fun (sheet, r0, c0, r1, c1) -> data.Cells sheet r0 c0 r1 c1)
        if blocks |> List.exists Option.isNone then None
        else
            let cells = blocks |> List.collect (fun block -> block.Value |> Array.toList |> List.collect Array.toList)
            // (As many points as the file cached: the range a pivot chart names runs on over the
            // pivot table's grand total - 13 cells for 12 points -, which is no point of the chart. A range
            // longer still has the rows of an outer field among the points - the years of a pivot table of
            // years and months, their subtotals: the points are the file's, as Excel draws them from the pivot
            // table - dashboard-sales-analysis' "Sales By Chain".)
            let pivoted = source.LimitToCachedPoints && not source.FollowedEdit && source.Count > 0
            if pivoted && cells.Length > source.Count + 1 then None
            else Some (if pivoted && source.Count < cells.Length then List.truncate source.Count cells else cells))

/// The x values of a scatter series as Excel 2019 plots them: the numbers of its cells; but where a cell holds
/// text, or the range is of more than one column and row (`c:multiLvlStrRef`), the points' places 1, 2, 3 ...
/// (checked: synthetic-chart-parity "ScatterText", an x axis 0 .. 150 for 120 points; vertex42's mortgage
/// calculator's interest-rate history against its No. / Date columns).
let private xValuesOf (data: ChartData) (source: DataRef option) (count: int) (numbers: DataRef option -> float option []) (texts: DataRef option -> string []) : float option [] * bool =
    match source with
    | None -> [||], false
    | Some reference ->
        let blocky =
            reference.Formula |> Option.bind parseAreas
            |> Option.map (List.exists (fun (_, r0, c0, r1, c1) -> r1 > r0 && c1 > c0)) |> Option.defaultValue false
        let values = numbers source
        let words = texts source
        let text = blocky || (words |> Array.mapi (fun index word -> word.Trim() <> "" && (index >= values.Length || values.[index].IsNone)) |> Array.exists id)
        if text then Array.init count (fun index -> Some (float (index + 1))), true
        else values, false

let rec private numberOfValue (value: CellValue) : float option =
    match value with
    | CellValue.Number n -> Some n
    | CellValue.Formula (_, result) -> numberOfValue result
    | CellValue.Boolean _ -> None
    | _ -> None

/// The numbers of a series: live when its cells can be read, else the file's
/// cache only if its reference has not changed since that cache was written.
let private numbersOf (data: ChartData) (source: DataRef option) : float option [] =
    match source with
    | None -> [||]
    | Some source ->
        match cellsOf data source with
        | Some cells -> cells |> List.map (fun cell -> cell |> Option.bind (fun cell -> numberOfValue cell.Value)) |> List.toArray
        | None when source.FollowedEdit -> [||]
        | None -> Array.init source.Count (fun index -> Map.tryFind index source.Numbers)

/// The texts of categories (or of a series name).
let private textsOf (data: ChartData) (source: DataRef option) : string [] =
    match source with
    | None -> [||]
    | Some source ->
        match cellsOf data source with
        | Some cells -> cells |> List.map (fun cell -> match cell with Some cell -> data.Display cell | None -> "") |> List.toArray
        | None when source.FollowedEdit -> [||]
        | None ->
            Array.init source.Count (fun index ->
                match Map.tryFind index source.Strings, Map.tryFind index source.Numbers, source.FormatCode with
                | _, Some value, Some code when code <> "General" -> data.Format code value
                | Some text, _, _ -> text
                | None, _, _ -> "")

let private formatCodeOf (data: ChartData) (source: DataRef option) : string option =
    match source with
    | None -> None
    | Some source ->
        let live =
            cellsOf data source
            |> Option.bind (fun cells -> cells |> List.tryPick (fun cell -> cell |> Option.bind (fun cell -> cell.Format.NumberFormat)))
        live |> Option.orElse source.FormatCode |> Option.filter (fun code -> code <> "" && code <> "General")

// ---------------------------------------------------------------------------
// Layout
// ---------------------------------------------------------------------------

let private accents = [| "accent1"; "accent2"; "accent3"; "accent4"; "accent5"; "accent6" |]

/// The colour Excel gives the n-th series (or point) that states none: the
/// theme's accents in turn, then darker and lighter rounds of them.
let private autoColor (theme: ThemeColors) (index: int) : string =
    let name = accents.[((index % 6) + 6) % 6]
    let hex = Map.tryFind name theme |> Option.orElse (Map.tryFind name defaultTheme) |> Option.defaultValue "4472C4"
    let round = index / 6
    if round = 0 then "#" + hex
    else
        let h, s, l = hslOf (rgbOf hex)
        // (The second six lighter - 60 % of the lightness and 40 % more: Excel 2019 draws series 7 and 8 of
        // synthetic-chart-looks "Markers" #8FA2D4 and #F1A78A, accents 1 and 2 lighter; later sixes not measured.)
        let l = if round = 1 then l * 0.6 + 0.4 elif round % 2 = 1 then l * 0.6 else min 0.9 (l + (1.0 - l) * 0.4)
        let r, g, b = rgbOfHsl (h, s, l)
        "#" + hex2 r + hex2 g + hex2 b

/// The colour of the n-th of `count` series (or points) that state none, under a chart style of the accents in turn,
/// as Excel 2019 draws a part that has no colours of its own: up to six the accents; more, each round of six a shade or
/// a tint (in linear light) of them - round p of a chart of n is (p - q / 2) x 1.4 / (q + 2) from the accent, q = n / 6
/// (whole): below zero a shade, above a tint (synthetic-chart-colors.xlsx through COM, 7 - 36 series: 7 are 0.767 shades
/// then 0.767 tints - #3B64AD, #8FA2D4 -; 12 are 0.65 shades then the accents; 30 are 0.5, 0.7, 0.9 shades, 0.9, 0.7
/// tints - each to a unit of 255).
let private autoColorOf (theme: ThemeColors) (index: int) (count: int) : string =
    if count <= 6 then autoColor theme index
    else
        let name = accents.[((index % 6) + 6) % 6]
        let hex = Map.tryFind name theme |> Option.orElse (Map.tryFind name defaultTheme) |> Option.defaultValue "4472C4"
        let q = float (count / 6)
        let x = (float (max 0 index / 6) - q / 2.0) * 1.4 / (q + 2.0)
        let toLinear (c: float) = let c = c / 255.0 in if c <= 0.04045 then c / 12.92 else ((c + 0.055) / 1.055) ** 2.4
        let toSrgb (c: float) = 255.0 * (if c <= 0.0031308 then c * 12.92 else 1.055 * (c ** (1.0 / 2.4)) - 0.055)
        let vary (c: float) =
            let linear = toLinear c
            toSrgb (if x < 0.0 then linear * (1.0 + x) else 1.0 - (1.0 - linear) * (1.0 - x))
        let r, g, b = rgbOf hex
        "#" + hex2 (vary r) + hex2 (vary g) + hex2 (vary b)

/// The colour of the n-th of `count` series (or points) under a chart style (`c:style`, 1 .. 48,
/// in rows of eight): the first of a row is shades of grey, the second the accents in turn, the
/// others one accent from dark to light.
let private styleColor (theme: ThemeColors) (style: int) (index: int) (count: int) : string =
    let kind = ((max 1 style - 1) % 8)
    if kind = 1 then autoColorOf theme index count
    else
        let share = if count <= 1 then 0.5 else float (((index % count) + count) % count) / float (count - 1)
        if kind = 0 then
            let level = 64.0 + share * 140.0
            "#" + hex2 level + hex2 level + hex2 level
        else
            let name = accents.[kind - 2]
            let hex = Map.tryFind name theme |> Option.orElse (Map.tryFind name defaultTheme) |> Option.defaultValue "4472C4"
            let r, g, b = rgbOf hex
            // (From a shade of the accent, through the accent, to a tint of it.)
            let mix (channel: float) =
                if share < 0.5 then channel * (0.55 + share * 0.9)
                else channel + (255.0 - channel) * ((share - 0.5) * 1.1)
            "#" + hex2 (mix r) + hex2 (mix g) + hex2 (mix b)

/// The advance widths of Calibri for the characters 32 .. 126, in thousandths of the font size.
let private calibriWidths : float [] =
    [| 226.0; 326.0; 401.0; 498.0; 507.0; 715.0; 682.0; 221.0; 303.0; 303.0; 498.0; 498.0; 250.0; 306.0; 252.0; 386.0
       507.0; 507.0; 507.0; 507.0; 507.0; 507.0; 507.0; 507.0; 507.0; 507.0; 268.0; 268.0; 498.0; 498.0; 498.0; 463.0
       894.0; 579.0; 544.0; 533.0; 615.0; 488.0; 459.0; 631.0; 623.0; 252.0; 319.0; 520.0; 420.0; 855.0; 646.0; 662.0
       517.0; 673.0; 543.0; 459.0; 487.0; 642.0; 567.0; 890.0; 519.0; 487.0; 468.0; 307.0; 386.0; 307.0; 498.0; 498.0
       291.0; 479.0; 525.0; 423.0; 525.0; 498.0; 305.0; 471.0; 525.0; 229.0; 239.0; 455.0; 229.0; 799.0; 525.0; 527.0
       525.0; 525.0; 349.0; 391.0; 335.0; 525.0; 452.0; 715.0; 433.0; 453.0; 395.0; 314.0; 460.0; 314.0; 498.0 |]

/// The width of a text in Calibri, in pixels (`size` in pixels; bold is a little wider).
let private calibriWidth (size: float) (bold: bool) (text: string) : float =
    let mutable units = 0.0
    for ch in text do
        let code = int ch
        units <- units + (if code >= 32 && code <= 126 then calibriWidths.[code - 32]
                          elif code >= 0x2E80 then 1000.0     // (CJK: a full square.)
                          else 520.0)
    units * size / 1000.0 * (if bold then 1.04 else 1.0)

/// The width a text is given in a chart's layout, in pixels: Calibri's, and a little more -
/// where Calibri is not installed the browser's font is wider.
/// The width of a run in Calibri at `size` pixels, as lines of shapes are broken by it.
let runWidth (size: float) (bold: bool) (text: string) : float = calibriWidth size bold text

let private textWidth (size: float) (text: string) : float =
    calibriWidth size false text * 1.08

/// Points to pixels.
let private px (points: float) = points * 96.0 / 72.0

/// The width of a text in a chart's typeface (None: Calibri) at `size` pixels, as GDI draws it where the
/// tables know the face (`ExcelText`), else Calibri's and a little more. A heavier weight of a family the
/// tables have (Segoe UI Black, Semibold) is its bold, a tenth wider for Black.
let private faceWidth (face: string option) (size: float) (bold: bool) (text: string) : float =
    let family = defaultArg face "Calibri"
    let family, bold, wider =
        if family.StartsWith "Segoe UI " then "Segoe UI", true, (if family.EndsWith "Black" then 1.1 else 1.0)
        else family, bold, 1.0
    if size > 0.0 && ExcelText.knowsFamily family then
        let pixels = max 1 (int (Math.Round size))
        ExcelText.textWidth (fun ch -> calibriWidth (float pixels) bold ch) family bold false pixels false text * size / float pixels * wider
    else calibriWidth size bold text * 1.08 * wider

let private lineExtents = GdiAdvances.lineExtents |> Array.map (fun (family, extent, units) -> family, float extent / float units) |> dict

/// A line of text in a typeface (None: Calibri) at `size` pixels, as Excel 2019 spaces its charts' text: the face's
/// own line as GDI has it, tmHeight and tmExternalLeading (Calibri 1.2207 em, Garamond 1.125, Arial 1.1499, Segoe UI
/// 1.3301, Tw Cen MT 1.0889: synthetic charts of each face made in Excel, `chart-axis-face-probe.vbs`); a face the
/// tables do not know as Calibri; a weight of a family (Segoe UI Semibold) as the family.
let private lineIn (face: string option) (size: float) =
    let family = (defaultArg face "Calibri").Trim().ToLowerInvariant()
    let known (name: string) = match lineExtents.TryGetValue name with | true, ratio -> Some ratio | _ -> None
    let ratio =
        known family
        |> Option.orElse ([ " black"; " heavy"; " semibold"; " bold"; " light" ] |> List.tryFind family.EndsWith |> Option.bind (fun weight -> known (family.Substring(0, family.Length - weight.Length))))
        |> Option.defaultValue (2500.0 / 2048.0)
    size * ratio

/// The scale of a value axis over lo .. hi: its ends and the step of its ticks, as Excel 2019
/// chooses them (each rule checked on charts made in Excel, `tools/excel-parity/chart-axis-probe.vbs`):
/// - zero is an end unless the values are close together far from it - their spread no more than a
///   sixth of the largest: then the low end is the lowest less half the spread (50 .. 60: from 45);
/// - 5 % of the spread is added beyond the largest (below the lowest, when that is under zero);
/// - the step is the smallest of 1, 2, 5 x 10^n that cuts the axis into no more than `most` intervals
///   (at most 10: `axisIntervals`), the ends whole steps;
/// - nothing to show (no values, or only zeros) is 0 .. 1.
let private valueScale (lo: float) (hi: float) (fixedMin: float option) (fixedMax: float option) (fixedStep: float option) (most: int) : float * float * float =
    let lo, hi = (if lo > hi then hi, lo else lo, hi)
    let most = max 1 (min 10 most)
    let empty = lo = 0.0 && hi = 0.0
    let lo, hi = (if lo = hi && not empty then (if lo > 0.0 then 0.0, hi else lo, 0.0) else lo, hi)
    let spread = hi - lo
    // (The ends before they are whole steps.)
    let low, high =
        if empty then 0.0, 1.0
        elif lo >= 0.0 then (if spread / hi > 1.0 / 6.0 then 0.0 else lo - spread / 2.0), hi + spread * 0.05
        elif hi <= 0.0 then lo - spread * 0.05, (if spread / -lo > 1.0 / 6.0 then 0.0 else hi + spread / 2.0)
        else lo - spread * 0.05, hi + spread * 0.05
    let low = defaultArg fixedMin low
    let high = defaultArg fixedMax high
    let wholeLow (step: float) = match fixedMin with Some value -> value | None -> Math.Floor (low / step + 1e-9) * step
    let wholeHigh (step: float) = match fixedMax with Some value -> value | None -> Math.Ceiling (high / step - 1e-9) * step
    let step =
        match fixedStep with
        | Some step when step > 0.0 -> step
        | _ ->
            let range = if high > low then high - low else 1.0
            let power = Math.Pow(10.0, Math.Floor (Math.Log10 (range / 10.0)))
            let ladder = [ for scale in 0 .. 12 do for unit in [ 1.0; 2.0; 5.0 ] -> unit * power * Math.Pow(10.0, float scale) ]
            ladder |> List.tryFind (fun step -> (wholeHigh step - wholeLow step) / step <= float most + 1e-9) |> Option.defaultValue (List.last ladder)
    let minimum = wholeLow step
    let maximum = wholeHigh step
    let maximum = if maximum <= minimum then minimum + step else maximum
    // (Never more ticks than can be read.)
    let step = if (maximum - minimum) / step > 40.0 then (maximum - minimum) / 10.0 else step
    minimum, maximum, step

/// How many intervals Excel lets a value axis have (`valueScale`): an axis up the side as many as lines of its
/// labels (in their face: `lineIn`) fit in the plot area's height, less two; one along the bottom one for every 7.5
/// widths of the face's widest digit (Calibri: 3.8 times the labels' size) of the plot area's width (both with the
/// axis' labels; sizes in pixels). Checked in Excel 2019 on column and bar charts of fonts of 8 - 24 pt, to a quarter
/// of a point, and of 10 faces at 9, 10, 12 pt (15,000 column and 6,600 bar charts, `chart-axis-face-probe.vbs`).
let private axisIntervals (alongBottom: bool) (room: float) (labelSize: float) (face: string option) : int =
    if labelSize <= 0.0 then 10
    elif alongBottom then
        // (One interval for 7.5 of the face's widest digit: Calibri's 0.5068 em, 3.8 x the size; Garamond's 0.4688, Verdana's
        // 0.6357 - 6,649 of 6,651 bar charts of 10 faces.)
        let digit = face |> Option.bind (fun family -> ExcelText.digitEm (family.Trim())) |> Option.defaultValue (1038.0 / 2048.0)
        int (Math.Floor (room / (7.5 * digit * labelSize)))
    else int (Math.Floor (room / lineIn face labelSize)) - 2

let private generalNumber (value: float) : string =
    let rounded = Math.Round (value * 1e10) / 1e10
    if rounded = Math.Floor rounded && abs rounded < 1e15 then (int64 rounded).ToString(Globalization.CultureInfo.InvariantCulture)
    else rounded.ToString(Globalization.CultureInfo.InvariantCulture)

let private strokeColor (stroke: Stroke option) : (string * float * string option) option =
    match stroke with
    | Some { Paint = FillColor (color, _); Width = width; Dash = dash } -> Some (color, width, dash)
    | _ -> None

/// The scale of an axis: its ends, the step of its ticks, and the base when it is logarithmic
/// (the ends and the step are then exponents of the base).
type private AxisScale = {
    Min: float
    Max: float
    Step: float
    Log: float option
}

/// Where a value is between the ends of a scale: 0 .. 1.
let private shareOn (scale: AxisScale) (value: float) : float =
    let at =
        match scale.Log with
        | Some logBase -> if value > 0.0 then Math.Log value / Math.Log logBase else scale.Min
        | None -> value
    if scale.Max > scale.Min then max 0.0 (min 1.0 ((at - scale.Min) / (scale.Max - scale.Min))) else 0.0

/// The values of the ticks of a scale (what their labels say).
let private ticksOf (scale: AxisScale) : float list =
    [ let mutable value = scale.Min
      let mutable guard = 0
      while value <= scale.Max + scale.Step * 1e-6 && guard < 60 do
          let tick = if abs value < scale.Step * 1e-9 then 0.0 else value
          yield (match scale.Log with Some logBase -> Math.Pow (logBase, tick) | None -> tick)
          value <- value + scale.Step
          guard <- guard + 1 ]

/// What one series draws with: its fill, its line and whether the line is drawn at all.
type private SeriesLook = {
    Color: string
    Opacity: float
    Filled: bool
    LineColor: string option
    LineWidth: float
    LineDash: string option
}

/// A series laid out: its name, its values and what it is drawn with.
type private SeriesData = {
    Source: Series
    Plot: Plot
    /// The index among all the series of the chart (for the automatic colours).
    Ordinal: int
    Label: string
    Values: float option []
    XValues: float option []
    Sizes: float option []
    Look: SeriesLook
    FormatCode: string option
    /// The number format of the x values (a scatter chart's), when they are numbers.
    XFormatCode: string option
    /// On the secondary value axis.
    Secondary: bool
}

/// A rectangle a manual layout gives in a chart of `width` x `height`, when it states all of it.
let private layoutRect (layout: ManualLayout option) (width: float) (height: float) : (float * float * float * float) option =
    match layout with
    | Some { XEdge = true; YEdge = true; X = Some x; Y = Some y; W = Some w; H = Some h } when w > 0.0 && h > 0.0 ->
        let left, top = max 0.0 (x * width), max 0.0 (y * height)
        Some (left, top, min (width - left) (w * width), min (height - top) (h * height))
    | _ -> None

/// A placeholder for a chartEx chart of a kind that is not drawn: what it is, with its title.
let private layoutUndrawnChartEx (layoutId: string) (chart: Chart) (width: float) (height: float) : Scene =
    let items = ResizeArray<Primitive>()
    let add (item: Primitive) = items.Add item
    let fitted (size: float) (room: float) (text: string) =
        if room <= 0.0 || textWidth size text <= room || text.Length <= 2 then text
        else
            let keep = max 1 (int (float text.Length * room / textWidth size text) - 1)
            text.Substring(0, min text.Length keep) + "…"
    add (PRect (0.5, 0.5, max 0.0 (width - 1.0), max 0.0 (height - 1.0), Some "#F6F8FA", 1.0, Some "#8A94A6", 1.0, 0.0))
    let lines = (match chart.Title with Some title -> title.Split('\n') |> Array.toList | None -> [])
    let titleSize = 14.0
    let block = float lines.Length * titleSize * 1.3 + 16.0
    let mutable y = height / 2.0 - block / 2.0 + titleSize * 0.65
    for line in lines do
        add (PText (width / 2.0, y, fitted titleSize (width - 16.0) line, titleSize, "#2D3748", "middle", true, false, 0.0))
        y <- y + titleSize * 1.3
    add (PText (width / 2.0, y + 2.0, fitted 11.0 (width - 16.0) (chartExKindName layoutId + " (not drawn)"), 11.0, "#718096", "middle", false, false, 0.0))
    { Width = width; Height = height; Items = List.ofSeq items }

/// The scene of a chartEx waterfall or funnel.
let private layoutPlottedChartEx (theme: ThemeColors) (data: ChartData) (chart: Chart) (plot: Plot) (width: float) (height: float) : Scene =
    let items = ResizeArray<Primitive>()
    let add (item: Primitive) = items.Add item
    let textColor (style: TextStyle) = (over style chart.TextStyle).Color |> Option.defaultValue "#595959"
    let textSize (style: TextStyle) (otherwise: float) = px ((over style chart.TextStyle).Size |> Option.defaultValue otherwise)
    let isBold (style: TextStyle) = (over style chart.TextStyle).Bold |> Option.defaultValue false
    let fitted (size: float) (room: float) (text: string) =
        if room <= 0.0 || textWidth size text <= room || text.Length <= 2 then text
        else
            let keep = max 1 (int (float text.Length * room / textWidth size text) - 1)
            text.Substring(0, min text.Length keep) + "…"
    (match chart.Background with
     | FillNone -> ()
     | FillColor (color, opacity) -> add (PRect (0.0, 0.0, width, height, Some color, opacity, None, 0.0, 0.0))
     | FillAuto -> add (PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, 0.0)))
    (match strokeColor chart.Border with
     | Some (color, strokeWidth, _) -> add (PRect (strokeWidth / 2.0, strokeWidth / 2.0, width - strokeWidth, height - strokeWidth, None, 1.0, Some color, strokeWidth, 0.0))
     | None -> ())
    let pad = 8.0
    let mutable top = pad
    let mutable bottom = height - pad
    let mutable left = pad
    let mutable right = width - pad
    (match (if chart.AutoTitleDeleted then None else chart.Title) with
     | Some title ->
        let size = textSize chart.TitleStyle 14.0
        let lines = title.Split('\n')
        for index in 0 .. lines.Length - 1 do
            add (PText (width / 2.0, top + size * 0.6 + float index * size * 1.2, fitted size (width - 2.0 * pad) lines.[index], size, textColor chart.TitleStyle, "middle", isBold chart.TitleStyle, false, 0.0))
        if not chart.TitleOverlay then top <- top + float lines.Length * size * 1.2 + 4.0
     | None -> ())
    (match plot.Series |> List.tryHead with
     | None -> ()
     | Some series ->
        let categories = textsOf data series.Categories
        let values = numbersOf data series.Values
        let count = max categories.Length values.Length
        let value (index: int) = if index < values.Length then defaultArg values.[index] 0.0 else 0.0
        let category (index: int) = if index < categories.Length then categories.[index] else ""
        let sourceFormat = formatCodeOf data series.Values
        let labelSpec = series.Labels
        let labelSize = labelSpec |> Option.map (fun spec -> textSize spec.Style 9.0) |> Option.defaultValue (px 9.0)
        let labelText (index: int) (amount: float) =
            match labelSpec with
            | Some spec when not (spec.Deleted.Contains index) ->
                let number =
                    match spec.FormatCode |> Option.orElse sourceFormat with
                    | Some code -> data.Format code amount
                    | None -> generalNumber amount
                let seriesName = series.Name |> Option.bind (fun name -> textsOf data (Some name) |> Array.tryHead) |> Option.defaultValue ""
                let parts =
                    [ if spec.ShowSeries && seriesName <> "" then yield seriesName
                      if spec.ShowCategory then yield category index
                      if spec.ShowValue then yield number ]
                match parts with
                | [] -> None
                | _ -> Some (String.concat (defaultArg spec.Separator ", ") parts)
            | _ -> None
        let fillOfPoint (index: int) (automatic: string) : (string * float) option =
            match Map.tryFind index series.PointFills with
            | Some (FillColor (color, opacity)) -> Some (color, opacity)
            | Some FillNone -> None
            | _ ->
                match series.Fill with
                | FillColor (color, opacity) -> Some (color, opacity)
                | FillNone -> None
                | FillAuto -> Some (automatic, 1.0)
        let outline = strokeColor series.Line
        /// A label's colour inside a bar: the label's own, else white on a dark bar.
        let insideColor (fill: (string * float) option) =
            match labelSpec |> Option.bind (fun spec -> spec.Style.Color) with
            | Some color -> color
            | None ->
                match fill with
                | Some (color, _) ->
                    let r, g, b = rgbOf (solidOf color)
                    if 0.299 * r + 0.587 * g + 0.114 * b < 140.0 then "#FFFFFF" else "#404040"
                | None -> "#404040"
        let labelColor = labelSpec |> Option.map (fun spec -> textColor spec.Style) |> Option.defaultValue "#595959"
        let catAxis = chart.Axes |> List.tryFind (fun axis -> axis.Kind = "cat")
        let valAxis = chart.Axes |> List.tryFind (fun axis -> axis.Kind = "val")
        let shows (axis: Axis option) = match axis with Some axis -> not axis.Deleted && axis.LabelsShown | None -> false
        let axisSize (axis: Axis option) = axis |> Option.map (fun axis -> textSize axis.LabelStyle 9.0) |> Option.defaultValue (px 9.0)
        let axisColor (axis: Axis option) = axis |> Option.map (fun axis -> textColor axis.LabelStyle) |> Option.defaultValue "#595959"
        match plot.Kind with
        | WaterfallPlot (subtotals, connectors) ->
            let isTotal (index: int) = List.contains index subtotals
            // Increase, decrease, total: the theme's first three accents.
            let automatic (kind: int) = autoColor theme kind
            (match chart.Legend with
             | Some legend ->
                let size = textSize legend.Style 9.0
                let color = textColor legend.Style
                let entries = [ "Increase", 0; "Decrease", 1; "Total", 2 ]
                let swatch (kind: int) = match series.Fill with FillColor (fill, _) -> fill | _ -> automatic kind
                // (A square 5.25 pt across, the text 5.25 pt after it, 6.5 pt to the next: Excel 2019's legends of
                // synthetic-chartex-zoo's waterfalls and treemaps.)
                let entryWidth (text: string) = px 5.25 + px 5.25 + textWidth size text + px 6.5
                let draw (x: float) (y: float) (text: string) (kind: int) =
                    add (PRect (x, y - px 5.25 / 2.0, px 5.25, px 5.25, Some (swatch kind), 1.0, None, 0.0, 0.0))
                    add (PText (x + px 5.25 + px 5.25, y, text, size, color, "start", false, false, 0.0))
                match legend.Position with
                | "l" | "r" ->
                    let column = entries |> List.map (fst >> entryWidth) |> List.max
                    let x = if legend.Position = "l" then left else right - column
                    let start = (top + bottom) / 2.0 - float entries.Length * size * 1.4 / 2.0
                    entries |> List.iteri (fun index (text, kind) -> draw x (start + float index * size * 1.4 + size * 0.7) text kind)
                    if not legend.Overlay then (if legend.Position = "l" then left <- left + column + 6.0 else right <- right - column - 6.0)
                | position ->
                    // (Excel 2019 puts a legend at the top 10.5 px lower than below the title and the plot 26.5 px
                    // further down: measured on synthetic-chartex.xlsx, synthetic-chartex-zoo and the e-commerce dashboard.)
                    let total = (entries |> List.sumBy (fst >> entryWidth)) - px 6.5
                    let y = if position = "b" then bottom - size * 0.6 else top + size * 0.6 + 10.5
                    let mutable x = (width - total) / 2.0
                    for text, kind in entries do
                        draw x y text kind
                        x <- x + entryWidth text
                    if not legend.Overlay then (if position = "b" then bottom <- bottom - size * 1.6 else top <- top + size * 1.6 + 26.5)
             | None -> ())
            // Where each bar starts and ends: a total from zero, every other point on from the last.
            let starts = Array.zeroCreate<float> count
            let ends = Array.zeroCreate<float> count
            let mutable running = 0.0
            for index in 0 .. count - 1 do
                let amount = value index
                if isTotal index then
                    starts.[index] <- 0.0
                    ends.[index] <- amount
                    running <- amount
                else
                    starts.[index] <- running
                    ends.[index] <- running + amount
                    running <- running + amount
            let lo = Array.append starts ends |> Array.fold min 0.0
            let hi = Array.append starts ends |> Array.fold max 0.0
            let minimum, maximum, step =
                valueScale lo hi (valAxis |> Option.bind (fun axis -> axis.Min)) (valAxis |> Option.bind (fun axis -> axis.Max)) (valAxis |> Option.bind (fun axis -> axis.MajorUnit)) 10
            let ticks = ticksOf { Min = minimum; Max = maximum; Step = step; Log = None }
            let tickText (tick: float) =
                match valAxis with
                | Some { FormatCode = Some code; SourceLinked = false } when code <> "General" -> data.Format code tick
                | _ -> match sourceFormat with Some code -> data.Format code tick | None -> generalNumber tick
            let valSize, catSize = axisSize valAxis, axisSize catAxis
            if shows valAxis then left <- left + (ticks |> List.map (tickText >> textWidth valSize) |> List.fold max 0.0) + 6.0
            if shows catAxis then bottom <- bottom - catSize * 1.6 - 3.5
            // (Labels above the bars have the room the axis leaves over the highest; with no legend
            // above, as much room as a label.)
            if labelSpec.IsSome && not (match chart.Legend with Some legend -> legend.Position = "t" && not legend.Overlay | None -> false) then top <- top + labelSize * 1.2
            let plotTop, plotBottom = top, max (top + 1.0) bottom
            let yOf (amount: float) = plotBottom - (amount - minimum) / (maximum - minimum) * (plotBottom - plotTop)
            (match valAxis |> Option.bind (fun axis -> axis.Gridlines) |> strokeColor with
             | Some (color, lineWidth, dash) -> for tick in ticks do add (PLine (left, yOf tick, right, yOf tick, color, lineWidth, dash))
             | None -> ())
            if shows valAxis then
                for tick in ticks do add (PText (left - 6.0, yOf tick, tickText tick, valSize, axisColor valAxis, "end", false, false, 0.0))
            let band = (right - left) / float (max 1 count)
            let barWidth = band / (1.0 + max 0.0 plot.GapWidth / 100.0)
            for index in 0 .. count - 1 do
                let x = left + band * float index + (band - barWidth) / 2.0
                let kind = if isTotal index then 2 elif value index >= 0.0 then 0 else 1
                let fill = fillOfPoint index (automatic kind)
                let high, low = max starts.[index] ends.[index], min starts.[index] ends.[index]
                let y0, y1 = yOf high, yOf low
                add (PRect (x, y0, barWidth, max 1.0 (y1 - y0), fill |> Option.map fst, (match fill with Some (_, opacity) -> opacity | None -> 1.0),
                            (outline |> Option.map (fun (color, _, _) -> color)), (match outline with Some (_, lineWidth, _) -> lineWidth | None -> 0.0), 0.0))
                // (The lines that join the bars: light grey, as the gridlines - Excel 2019's.)
                if connectors && index < count - 1 then
                    add (PLine (x + barWidth, yOf ends.[index], x + band, yOf ends.[index], "#D9D9D9", 0.75, None))
                match labelText index (value index) with
                | Some text ->
                    let position = labelSpec |> Option.bind (fun spec -> spec.Position) |> Option.defaultValue "outEnd"
                    let rising = kind <> 1
                    let y =
                        match position with
                        | "ctr" -> (y0 + y1) / 2.0
                        | "inEnd" -> if rising then y0 + labelSize * 0.7 else y1 - labelSize * 0.7
                        | "inBase" -> if rising then y1 - labelSize * 0.7 else y0 + labelSize * 0.7
                        | _ -> if rising then y0 - labelSize * 0.7 else y1 + labelSize * 0.7
                    let color = if position = "outEnd" || position = "" then labelColor else insideColor fill
                    add (PText (x + barWidth / 2.0, y, fitted labelSize (band + 4.0) text, labelSize, color, "middle", false, false, 0.0))
                | None -> ()
            if minimum < 0.0 && maximum > 0.0 then add (PLine (left, yOf 0.0, right, yOf 0.0, "#D9D9D9", 0.75, None))
            if shows catAxis then
                for index in 0 .. count - 1 do
                    add (PText (left + band * (float index + 0.5), plotBottom + catSize * 0.9 + 6.0, fitted catSize (band - 2.0) (category index), catSize, axisColor catAxis, "middle", false, false, 0.0))
        | _ ->
            // A funnel: one bar a category, top to bottom, centred, as long as its value is of the largest.
            let catSize = axisSize catAxis
            if shows catAxis then
                left <- left + (Seq.init count category |> Seq.map (textWidth catSize) |> Seq.fold max 0.0) + 8.0
            let band = (bottom - top) / float (max 1 count)
            let barHeight = band / (1.0 + max 0.0 plot.GapWidth / 100.0)
            let largest = Seq.init count (value >> abs) |> Seq.fold max 0.0
            let largest = if largest > 0.0 then largest else 1.0
            let centre = (left + right) / 2.0
            let automatic = autoColor theme 0
            for index in 0 .. count - 1 do
                let length = (right - left) * abs (value index) / largest
                let y = top + band * float index + (band - barHeight) / 2.0
                let fill = fillOfPoint index automatic
                add (PRect (centre - length / 2.0, y, length, barHeight, fill |> Option.map fst, (match fill with Some (_, opacity) -> opacity | None -> 1.0),
                            (outline |> Option.map (fun (color, _, _) -> color)), (match outline with Some (_, lineWidth, _) -> lineWidth | None -> 0.0), 0.0))
                if shows catAxis then
                    add (PText (left - 8.0, y + barHeight / 2.0, category index, catSize, axisColor catAxis, "end", false, false, 0.0))
                match labelText index (value index) with
                | Some text ->
                    let inside = textWidth labelSize text <= length
                    let x, anchor, color = if inside then centre, "middle", insideColor fill else centre + length / 2.0 + 4.0, "start", labelColor
                    add (PText (x, y + barHeight / 2.0, text, labelSize, color, anchor, false, false, 0.0))
                | None -> ())
    { Width = width; Height = height; Items = List.ofSeq items }

/// The colour of the n-th point of a chartEx chart that states none: its colours part's (Excel's own, `colors*.xml`
/// id 10) - the six accents, then again under each variation in turn: darker to 60 %, lighter (80 % + 20 %), 80 %,
/// 60 % + 40 %, 50 %, 70 % + 30 %, 70 %, 50 % + 50 % of their luminance (synthetic-chartex-zoo "TreemapFlat": the
/// seventh #264478).
let chartExColor (theme: ThemeColors) (index: int) : string =
    let name = accents.[((index % 6) + 6) % 6]
    let hex = Map.tryFind name theme |> Option.orElse (Map.tryFind name defaultTheme) |> Option.defaultValue "4472C4"
    let variations = [| 1.0, 0.0; 0.6, 0.0; 0.8, 0.2; 0.8, 0.0; 0.6, 0.4; 0.5, 0.0; 0.7, 0.3; 0.7, 0.0; 0.5, 0.5 |]
    let modulate, offset = variations.[(max 0 index / 6) % variations.Length]
    if modulate = 1.0 then "#" + hex
    else
        let h, s, l = hslOf (rgbOf hex)
        let r, g, b = rgbOfHsl (h, s, min 1.0 (l * modulate + offset))
        "#" + hex2 r + hex2 g + hex2 b

/// A node of a treemap's or a sunburst's categories: its name, its value (its own, or its children's together), the
/// place of its first row among the data (which colour it has) and its children.
type HierarchyNode = { Name: string; Value: float; First: int; Children: HierarchyNode list }

/// The rows of a hierarchy chart's categories - each the path from its outer level in, and its value - grouped into
/// nodes in the order their names first come.
let rec private hierarchyNodes (rows: (int * string list * float) list) : HierarchyNode list =
    let order = ResizeArray<string>()
    let groups = Collections.Generic.Dictionary<string, ResizeArray<int * string list * float>>()
    for (first, path, value) in rows do
        match path with
        | head :: _ ->
            if not (groups.ContainsKey head) then (order.Add head; groups.[head] <- ResizeArray())
            groups.[head].Add (first, path, value)
        | [] -> ()
    [ for name in order ->
        let members = List.ofSeq groups.[name]
        let deeper = members |> List.choose (fun (first, path, value) -> match path with _ :: (_ :: _ as rest) -> Some (first, rest, value) | _ -> None)
        { Name = name; Value = members |> List.sumBy (fun (_, _, value) -> value)
          First = members |> List.map (fun (first, _, _) -> first) |> List.min
          Children = hierarchyNodes deeper } ]

/// What a treemap's or a sunburst's series holds: each row's categories (from the outer level in: the columns of its
/// range left to right, or the levels its part caches) and its value; rows of no value above zero left out.
let private hierarchyRows (data: ChartData) (series: Series) : (int * string list * float) list =
    let values = numbersOf data series.Values
    let paths : string list [] =
        match series.Categories with
        | None -> Array.init values.Length (fun index -> [ string (index + 1) ])
        | Some categories ->
            let columns =
                categories.Formula |> Option.bind parseAreas
                |> Option.bind (fun areas ->
                    match areas with
                    | [ sheet, r0, c0, r1, c1 ] -> data.Cells sheet r0 c0 r1 c1 |> Option.map (Array.map (Array.map (fun cell -> match cell with Some cell -> data.Display cell | None -> "")))
                    | _ -> None)
            match columns with
            | Some rows -> rows |> Array.map Array.toList
            | None ->
                // (The part's cache: its first level names the points, the others are the levels above, from the inner out.)
                let levels = (categories.Outer |> List.rev) @ [ categories.Strings ]
                Array.init (max categories.Count values.Length) (fun index -> levels |> List.map (fun level -> Map.tryFind index level |> Option.defaultValue ""))
    // (A path ends at its first empty level.)
    let trimmed (path: string list) = path |> List.map (fun name -> name.Trim()) |> List.takeWhile (fun name -> name <> "")
    [ for index in 0 .. values.Length - 1 do
        match values.[index] with
        | Some value when value > 0.0 ->
            let path = if index < paths.Length then trimmed paths.[index] else []
            yield index, (if path.IsEmpty then [ string (index + 1) ] else path), value
        | _ -> () ]

/// Rectangles for values, largest first, in a rectangle, as Excel 2019 lays out a treemap: a strip at a time along
/// the shorter side (a column at the left of a rectangle at least as wide as high, else a row at the top), as many in
/// it as bring the strip's smallest rectangle's width over height nearer 1 - not its ratio of longer to shorter side:
/// 160 of 160 flat treemaps through COM as Excel (`chartex-treemap-probe.vbs`: 40 sets of 3 - 14 values in 4 chart
/// sizes; the usual squarify rule 128).
let squarify (values: float list) (x: float) (y: float) (w: float) (h: float) : (float * float * float * float) list =
    let total = List.sum values
    if total <= 0.0 || values.IsEmpty then values |> List.map (fun _ -> x, y, 0.0, 0.0)
    else
        let scale = w * h / total
        let rec go (items: float list) (x: float) (y: float) (w: float) (h: float) : (float * float * float * float) list =
            match items with
            | [] -> []
            | [ _ ] -> [ x, y, w, h ]
            | _ ->
                let column = w >= h
                let side = if column then h else w
                /// How far from 1 the width over height of the strip's last (smallest) rectangle is.
                let badness (row: float list) =
                    let thick = List.sum row / side
                    let length = List.last row / thick
                    abs ((if column then thick / length else length / thick) - 1.0)
                let rec grow (row: float list) (rest: float list) =
                    match rest with
                    | next :: more when badness (row @ [ next ]) <= badness row -> grow (row @ [ next ]) more
                    | _ -> row, rest
                let row, rest = grow [ List.head items ] (List.tail items)
                let thick = if side > 0.0 then List.sum row / side else 0.0
                let mutable offset = 0.0
                let placed =
                    row |> List.map (fun area ->
                        let length = if thick > 0.0 then area / thick else 0.0
                        let rect = if column then x, y + offset, thick, length else x + offset, y, length, thick
                        offset <- offset + length
                        rect)
                placed @ (if column then go rest (x + thick) y (w - thick) h else go rest x (y + thick) w (h - thick))
        go (values |> List.map (fun value -> value * scale)) x y w h

/// The scene of a chartEx treemap or sunburst: its title, its legend (the outer level's names), its rectangles or
/// rings - each outer node in its colour (`chartExColor`, by the order its first row comes in), its children in its,
/// largest first; 1.5 pt of white round each (the chartEx style's) - and its labels.
let private layoutHierarchyChartEx (theme: ThemeColors) (data: ChartData) (chart: Chart) (plot: Plot) (width: float) (height: float) : Scene =
    let items = ResizeArray<Primitive>()
    let add (item: Primitive) = items.Add item
    let face (style: TextStyle) = (over style chart.TextStyle).Face |> Option.orElse (Map.tryFind "#minorFont" theme)
    let sizeOf (style: TextStyle) (otherwise: float) = px ((over style chart.TextStyle).Size |> Option.defaultValue otherwise)
    let colorOf (style: TextStyle) (otherwise: string) = (over style chart.TextStyle).Color |> Option.defaultValue otherwise
    let text (style: TextStyle) (item: Primitive) = match face style with Some name when name <> "Calibri" -> add (PFace (name, [ item ])) | _ -> add item
    let widthOf (style: TextStyle) (size: float) (bold: bool) (value: string) = faceWidth (face style) size bold value
    (match chart.Background with
     | FillColor (color, opacity) -> add (PRect (0.0, 0.0, width, height, Some color, opacity, None, 0.0, 0.0))
     | FillAuto -> add (PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, 0.0))
     | FillNone -> ())
    (match strokeColor chart.Border with
     | Some (color, strokeWidth, _) -> add (PRect (strokeWidth / 2.0, strokeWidth / 2.0, width - strokeWidth, height - strokeWidth, None, 1.0, Some color, strokeWidth, 0.0))
     | None -> ())
    // (Measured on Excel 2019's pictures of synthetic-chartex-zoo at 360 x 216 pt: the title's box 4.5 pt from the top,
    // a legend's line 8.5 pt below it, the plot 10.5 pt below that - 3 pt below the title without a legend -, 6 pt in
    // from the other edges.)
    let mutable top = px 6.0
    (match (if chart.AutoTitleDeleted then None else chart.Title) with
     | Some title ->
        let size = sizeOf chart.TitleStyle 14.0
        let bold = chart.TitleStyle.Bold |> Option.defaultValue false
        let line = lineIn (face chart.TitleStyle) size
        let mutable y = px 4.5 + px 1.5
        for part in title.Split('\n') do
            text chart.TitleStyle (PText (width / 2.0, y + line / 2.0, part, size, colorOf chart.TitleStyle "#595959", "middle", bold, false, 0.0))
            y <- y + line
        if not chart.TitleOverlay then top <- y + px 1.5 + px 3.0
     | None -> ())
    let rows = plot.Series |> List.tryHead |> Option.map (hierarchyRows data) |> Option.defaultValue []
    let nodes = hierarchyNodes rows
    let colorAt (node: HierarchyNode) =
        let series = plot.Series |> List.tryHead
        match series |> Option.bind (fun series -> Map.tryFind node.First series.PointFills) with
        | Some (FillColor (color, _)) -> color
        | _ ->
            match series |> Option.map (fun series -> series.Fill) with
            | Some (FillColor (color, _)) -> color
            | _ -> chartExColor theme (nodes |> List.findIndex (fun known -> known.First = node.First))
    (match chart.Legend with
     | Some legend when not nodes.IsEmpty && (legend.Position = "t" || legend.Position = "b") ->
        let size = sizeOf legend.Style 9.0
        let line = lineIn (face legend.Style) size
        let swatch, after, between = px 5.25, px 5.25, px 6.5
        let entryWidth (node: HierarchyNode) = swatch + after + widthOf legend.Style size false node.Name
        let total = (nodes |> List.sumBy entryWidth) + between * float (nodes.Length - 1)
        let middle = if legend.Position = "b" then height - px 6.0 - line / 2.0 else top + px 5.5 + line / 2.0
        let mutable x = (width - total) / 2.0
        for node in nodes do
            add (PRect (x, middle - swatch / 2.0, swatch, swatch, Some (colorAt node), 1.0, None, 0.0, 0.0))
            text legend.Style (PText (x + swatch + after, middle, node.Name, size, colorOf legend.Style "#595959", "start", false, false, 0.0))
            x <- x + entryWidth node + between
        if not legend.Overlay && legend.Position = "t" then top <- middle + line / 2.0 + px 10.5
     | _ -> ())
    let bottom = height - px 6.0 - (match chart.Legend with Some legend when legend.Position = "b" && not legend.Overlay -> lineIn (face legend.Style) (sizeOf legend.Style 9.0) + px 6.0 | _ -> 0.0)
    let left, right = px 5.25, width - px 6.0
    let series = plot.Series |> List.tryHead
    let outline = series |> Option.bind (fun series -> strokeColor series.Line) |> Option.defaultValue ("#FFFFFF", 2.0, None)
    let lineColor, lineWidth, _ = outline
    let labels = series |> Option.bind (fun series -> series.Labels)
    let labelStyle = labels |> Option.map (fun spec -> spec.Style) |> Option.defaultValue noTextStyle
    let labelSize = sizeOf labelStyle 9.0
    let labelText (node: HierarchyNode) =
        match labels with
        | Some spec ->
            let number = match spec.FormatCode |> Option.orElse (series |> Option.bind (fun series -> formatCodeOf data series.Values)) with Some code -> data.Format code node.Value | None -> generalNumber node.Value
            [ if spec.ShowCategory then yield node.Name
              if spec.ShowValue then yield number ]
            |> function [] -> None | parts -> Some (String.concat (defaultArg spec.Separator ", ") parts)
        | None -> None
    let labelLine = lineIn (face labelStyle) labelSize
    let labelColor = colorOf labelStyle "#FFFFFF"
    /// A label's lines no wider than `room`: broken at its spaces, a word too long for a line at its letters.
    let brokenTo (room: float) (label: string) : string list =
        let fits (value: string) = widthOf labelStyle labelSize false value <= room
        let rec letters (word: string) =
            if word = "" then []
            else
                let mutable cut = word.Length
                while cut > 1 && not (fits (word.Substring(0, cut))) do cut <- cut - 1
                word.Substring(0, cut) :: letters (word.Substring cut)
        label.Split(' ')
        |> Array.fold (fun (lines: string list) word ->
            match lines with
            | last :: rest when fits (last + " " + word) -> (last + " " + word) :: rest
            | _ -> (List.rev (letters word)) @ lines) []
        |> List.rev
    match plot.Kind with
    | SunburstPlot ->
        // Rings: the hole and each level as wide as the outer radius over the levels and one (synthetic-chartex-zoo: 121 px
        // in 3, 2, 1 levels: a hole of 30, 40, 61), from twelve o'clock clockwise, largest first.
        let rec depth (node: HierarchyNode) = 1 + (node.Children |> List.fold (fun most child -> max most (depth child)) 0)
        let levels = nodes |> List.fold (fun most node -> max most (depth node)) 0
        let radius = max 0.0 (min (right - left) (bottom - top) / 2.0)
        let cx, cy = (left + right) / 2.0, (top + bottom) / 2.0
        let ring = radius / float (levels + 1)
        let total = nodes |> List.sumBy (fun node -> node.Value)
        let point (r: float) (angle: float) = cx + r * Math.Sin angle, cy - r * Math.Cos angle
        let n (value: float) = string (Math.Round (value, 2))
        let rec draw (siblings: HierarchyNode list) (level: int) (start: float) (span: float) (color: string option) =
            let whole = siblings |> List.sumBy (fun node -> node.Value)
            let mutable at = start
            for node in siblings |> List.sortBy (fun node -> -node.Value) do
                let sweep = if whole > 0.0 then span * node.Value / whole else 0.0
                let inner, outer = ring * float (level + 1), ring * float (level + 2)
                let x0, y0 = point outer at
                let x1, y1 = point outer (at + sweep)
                let x2, y2 = point inner (at + sweep)
                let x3, y3 = point inner at
                let large = if sweep > Math.PI then "1" else "0"
                let d =
                    if sweep >= 2.0 * Math.PI - 1e-9 then
                        let mx, my = point outer (at + Math.PI)
                        let ix, iy = point inner (at + Math.PI)
                        sprintf "M %s %s A %s %s 0 1 1 %s %s A %s %s 0 1 1 %s %s Z M %s %s A %s %s 0 1 0 %s %s A %s %s 0 1 0 %s %s Z"
                            (n x0) (n y0) (n outer) (n outer) (n mx) (n my) (n outer) (n outer) (n x0) (n y0) (n x3) (n y3) (n inner) (n inner) (n ix) (n iy) (n inner) (n inner) (n x3) (n y3)
                    else
                        sprintf "M %s %s A %s %s 0 %s 1 %s %s L %s %s A %s %s 0 %s 0 %s %s Z" (n x0) (n y0) (n outer) (n outer) large (n x1) (n y1) (n x2) (n y2) (n inner) (n inner) large (n x3) (n y3)
                let fill = (match color with Some given -> given | None -> colorAt node)
                add (PPath (d, Some fill, 1.0, Some lineColor, lineWidth))
                // (A label along its part's middle radius, reading outwards - turned over on the left half -, broken to
                // the ring's width, its lines side by side across the part: Excel 2019's "Tokyo", "Franc / e".)
                (match labelText node with
                 | Some label when sweep > 0.0 ->
                    let middle = at + sweep / 2.0
                    let lx, ly = point ((inner + outer) / 2.0) middle
                    let lines = brokenTo (ring + 2.0) label
                    let arc = sweep * (inner + outer) / 2.0
                    if not lines.IsEmpty && float lines.Length * labelLine <= arc then
                        let degrees = middle * 180.0 / Math.PI
                        let turn = if degrees > 180.0 then degrees - 270.0 else degrees - 90.0
                        let across = turn * Math.PI / 180.0
                        lines |> List.iteri (fun index part ->
                            let shift = (float index - float (lines.Length - 1) / 2.0) * labelLine
                            text labelStyle (PText (lx - shift * Math.Sin across, ly + shift * Math.Cos across, part, labelSize, labelColor, "middle", false, false, turn)))
                 | _ -> ())
                draw node.Children (level + 1) at sweep (Some fill)
                at <- at + sweep
        if total > 0.0 then draw nodes 0 0.0 (2.0 * Math.PI) None
    | TreemapPlot parentLabels ->
        // Rectangles: the outer level's in the plot, each node's children in its rectangle, largest first (`squarify`).
        // Labels (Excel 2019's pictures of synthetic-chartex-looks): a tile's at its bottom left, a parent's name at its
        // top left over its children ("overlapping"), or in a grey banner (#A6A6A6) a line high above them ("banner"),
        // 4.5 pt in from the sides, 5.25 pt from the top, 3.75 pt from the bottom.
        let padX, padY = px 4.5, px 3.75
        let labelAt (x: float) (y: float) (w: float) (h: float) (top: bool) (label: string) =
            if labelLine <= h - padY && widthOf labelStyle labelSize false label <= w - 2.0 * padX then
                let middle = if top then y + px 5.25 + labelLine / 2.0 else y + h - padY - labelLine / 2.0
                text labelStyle (PText (x + padX, middle, label, labelSize, labelColor, "start", false, false, 0.0))
        // (Only the outer level's names stand as a parent's: Excel 2019's "Treemap3" - Asia over Tokyo, Osaka, Delhi; no
        // Japan, no India.)
        let rec place (siblings: HierarchyNode list) (x: float) (y: float) (w: float) (h: float) (color: string option) =
            let sorted = siblings |> List.sortBy (fun node -> -node.Value)
            let rects = squarify (sorted |> List.map (fun node -> node.Value)) x y w h
            for node, (rx, ry, rw, rh) in List.zip sorted rects do
                let fill = (match color with Some given -> given | None -> colorAt node)
                if node.Children.IsEmpty then
                    add (PRect (rx, ry, rw, rh, Some fill, 1.0, Some lineColor, lineWidth, 0.0))
                    labelText node |> Option.iter (labelAt rx ry rw rh false)
                else
                    match (if color.IsNone then labels else None), parentLabels with
                    | Some _, "banner" ->
                        let banner = min rh (labelLine + 2.0 * padY)
                        add (PRect (rx, ry, rw, banner, Some "#A6A6A6", 1.0, Some lineColor, lineWidth, 0.0))
                        labelAt rx ry rw banner true node.Name
                        place node.Children rx (ry + banner) rw (rh - banner) (Some fill)
                    | Some _, "none" | None, _ -> place node.Children rx ry rw rh (Some fill)
                    | Some _, _ ->
                        place node.Children rx ry rw rh (Some fill)
                        labelAt rx ry rw rh true node.Name
        place nodes left top (right - left) (bottom - top) None
    | _ -> ()
    { Width = width; Height = height; Items = List.ofSeq items }

/// The scene of a chartEx chart: a waterfall, a funnel, a treemap, a sunburst, or a placeholder that says what a
/// chart of another kind is, with its title.
let private layoutChartEx (theme: ThemeColors) (data: ChartData) (chart: Chart) (plot: Plot) (width: float) (height: float) : Scene =
    match plot.Kind with
    | UndrawnPlot layoutId -> layoutUndrawnChartEx layoutId chart width height
    | TreemapPlot _ | SunburstPlot -> layoutHierarchyChartEx theme data chart plot width height
    | _ -> layoutPlottedChartEx theme data chart plot width height

/// A smoothed line through points (`c:smooth`): a cubic Bézier curve between each two, its handles a
/// sixth of the way to the points around them (Catmull-Rom), as Excel draws one.
let private smoothPath (points: (float * float) list) : string =
    let p = List.toArray points
    let n (value: float) = string (Math.Round (value, 2))
    let sb = System.Text.StringBuilder()
    sb.Append("M ").Append(n (fst p.[0])).Append(" ").Append(n (snd p.[0])) |> ignore
    for i in 0 .. p.Length - 2 do
        let x0, y0 = p.[max 0 (i - 1)]
        let x1, y1 = p.[i]
        let x2, y2 = p.[i + 1]
        let x3, y3 = p.[min (p.Length - 1) (i + 2)]
        sb.Append(" C ").Append(n (x1 + (x2 - x0) / 6.0)).Append(" ").Append(n (y1 + (y2 - y0) / 6.0))
          .Append(" ").Append(n (x2 - (x3 - x1) / 6.0)).Append(" ").Append(n (y2 - (y3 - y1) / 6.0))
          .Append(" ").Append(n x2).Append(" ").Append(n y2) |> ignore
    sb.ToString()

/// A legend entry's key: a line of the series' width (pixels) with its marker (symbol, size, fill, outline) in
/// the middle, or a box with an outline (colour, width).
type private KeyLook = {
    Line: float option
    Marker: (string * float * string option * string option) option
    Outline: (string * float) option
}

/// How a category axis' labels stand: in a line, broken at spaces (each label's lines), or turned (degrees).
type private LabelPlan =
    | Flat
    | Wrapped of string list []
    | Turned of float

/// The scene of a chart of the classic kinds.
let private layoutClassicChart (theme: ThemeColors) (data: ChartData) (chart: Chart) (width: float) (height: float) : Scene =
    let items = ResizeArray<Primitive>()
    let add (item: Primitive) = items.Add item
    // (Text that names no colour is black, as Excel draws it: the charts it writes name theirs.)
    let textColor (style: TextStyle) = over style chart.TextStyle |> fun merged -> merged.Color |> Option.defaultValue "#000000"
    let textSize (style: TextStyle) (otherwise: float) = px (over style chart.TextStyle |> fun merged -> merged.Size |> Option.defaultValue otherwise)
    let faceOf (style: TextStyle) = (over style chart.TextStyle).Face |> Option.orElse (Map.tryFind "#minorFont" theme)
    /// A text in its typeface: Calibri as it is, another in a `PFace`.
    let addText (style: TextStyle) (item: Primitive) =
        match faceOf style with
        | Some face when face <> "Calibri" ->
            // (A heavy weight of a family - Segoe UI Black - is the family's bold where the weight is not there.)
            let heavy = [ " Black"; " Heavy"; " Semibold"; " Bold" ] |> List.tryFind (fun weight -> face.EndsWith weight)
            match heavy, item with
            | Some weight, PText (x, y, text, size, color, anchor, _, italic, turn) ->
                add (PFace (face.Substring(0, face.Length - weight.Length), [ PText (x, y, text, size, color, anchor, true, italic, turn) ]))
            | _ -> add (PFace (face, [ item ]))
        | _ -> add item
    let widthIn (style: TextStyle) (size: float) (bold: bool) (text: string) = faceWidth (faceOf style) size bold text
    /// A line of a text of this look at `size` pixels: its typeface's.
    let lineFor (style: TextStyle) (size: float) = lineIn (faceOf style) size
    /// Whether a title's runs differ in their looks (else it is drawn as one text of the title's look).
    let mixedRuns (runs: (string * TextStyle) list list) =
        runs |> List.exists (fun line -> (line |> List.map snd |> List.distinct |> List.length) > 1)
    /// A line of runs, each in its own look (`otherwise`: what a run names not, in pixels), its middle at `y` - or,
    /// turned up (-90), at `x` - from `at` (the line's start, its middle or its end by `anchor`); baselines meet.
    let drawRuns (runs: (string * TextStyle) list) (otherwiseSize: float) (otherwiseBold: bool) (x: float) (y: float) (anchor: string) (turn: float) =
        let pieces =
            runs |> List.map (fun (text, style) ->
                let size = textSize style otherwiseSize
                let bold = style.Bold |> Option.defaultValue otherwiseBold
                text, style, size, bold, widthIn style size bold text)
        let total = pieces |> List.sumBy (fun (_, _, _, _, wide) -> wide)
        let largest = pieces |> List.fold (fun most (_, _, size, _, _) -> max most size) 0.0
        let start = match anchor with "middle" -> -total / 2.0 | "end" -> -total | _ -> 0.0
        let mutable along = start
        for text, style, size, bold, wide in pieces do
            // (Texts are drawn about their middle: a smaller run is moved down to the larger's baseline.)
            let drop = (largest - size) * 0.28
            let item =
                if turn = 0.0 then PText (x + along, y + drop, text, size, textColor style, "start", bold, (style.Italic |> Option.defaultValue false), 0.0)
                else PText (x + drop, y - along, text, size, textColor style, "start", bold, (style.Italic |> Option.defaultValue false), turn)
            addText style item
            along <- along + wide

    // ---- the chart's own rectangle ----
    // (Rounded corners - `c:roundedCorners` on, or not written - are 10 px round, the chart's fill too: Excel 2019's
    // pictures of synthetic-chart-looks "RoundedBare", "LegendRight".)
    let corner = if chart.RoundedCorners then 10.0 else 0.0
    (match chart.Background with
     | FillNone -> ()
     | FillColor (color, opacity) -> add (PRect (0.0, 0.0, width, height, Some color, opacity, None, 0.0, corner))
     | FillAuto -> add (PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, corner)))
    (match strokeColor chart.Border with
     | Some (color, strokeWidth, _) -> add (PRect (strokeWidth / 2.0, strokeWidth / 2.0, width - strokeWidth, height - strokeWidth, None, 1.0, Some color, strokeWidth, corner))
     | None -> ())

    // ---- the series ----
    let axisById (id: string) = chart.Axes |> List.tryFind (fun axis -> axis.Id = id)
    let valueAxes = chart.Axes |> List.filter (fun axis -> axis.Kind = "val")
    let pointPlot (plot: Plot) = plot.Kind = ScatterPlot || plot.Kind = BubblePlot
    let primaryValueAxis =
        // (The first value axis of the part that a plot uses - Excel's primary axes are the first written, whatever
        // plot comes first: dashboard-ecommerce-sales' area of profit before its columns of sales has the sales'
        // axis at the left; on a scatter chart the first plot's at the side.)
        let ofPlot (plot: Plot) = plot.AxisIds |> List.choose axisById |> List.filter (fun axis -> axis.Kind = "val")
        let used = chart.Plots |> List.collect (fun plot -> plot.AxisIds) |> Set.ofList
        match chart.Plots with
        | first :: _ when not (pointPlot first) && (valueAxes |> List.exists (fun axis -> used.Contains axis.Id)) ->
            valueAxes |> List.tryFind (fun axis -> used.Contains axis.Id)
        | first :: _ ->
            (match ofPlot first with
             | [ only ] -> Some only
             | several when pointPlot first -> several |> List.tryFind (fun axis -> axis.Position = "l" || axis.Position = "r") |> Option.orElse (List.tryLast several)
             | several -> List.tryHead several)
        | [] -> None
    let seriesCount = chart.Plots |> List.sumBy (fun plot -> plot.Series.Length)
    let mutable ordinal = 0
    let series : SeriesData list =
        chart.Plots
        |> List.collect (fun plot ->
            let plotValueAxis = plot.AxisIds |> List.choose axisById |> List.filter (fun axis -> axis.Kind = "val" && (not (pointPlot plot) || axis.Position = "l" || axis.Position = "r")) |> List.tryHead
            let secondary =
                match plotValueAxis, primaryValueAxis with
                | Some mine, Some first -> mine.Id <> first.Id
                | _ -> false
            plot.Series
            |> List.map (fun source ->
                let at = ordinal
                ordinal <- ordinal + 1
                let auto = styleColor theme chart.Style at seriesCount
                let lineOnly = plot.Kind = LinePlot || plot.Kind = ScatterPlot || plot.Kind = StockPlot || (plot.Kind = RadarPlot && plot.RadarStyle <> "filled")
                let look =
                    if lineOnly then
                        match source.Line with
                        | Some { Paint = FillNone } -> { Color = auto; Opacity = 1.0; Filled = false; LineColor = None; LineWidth = 0.0; LineDash = None }
                        | Some { Paint = FillColor (color, opacity); Width = strokeWidth; Dash = dash } ->
                            { Color = color; Opacity = opacity; Filled = false; LineColor = Some color; LineWidth = strokeWidth; LineDash = dash }
                        | Some { Width = strokeWidth; Dash = dash } -> { Color = auto; Opacity = 1.0; Filled = false; LineColor = Some auto; LineWidth = strokeWidth; LineDash = dash }
                        | None ->
                            // (A series that states no line: the chart style's - 1.5 pt for styles 1 - 8, 3.5 pt for
                            // 25 - 32, else 2.5 pt; Excel 2019, synthetic-chart-lines "Plain2" / "Plain34" and every
                            // style set through COM.)
                            let points = if chart.Style <= 8 then 1.5 elif chart.Style >= 25 && chart.Style <= 32 then 3.5 else 2.5
                            { Color = auto; Opacity = 1.0; Filled = false; LineColor = Some auto; LineWidth = linePixels (points * 12700.0); LineDash = None }
                    else
                        let color, opacity, filled =
                            match source.Fill with
                            | FillColor (color, opacity) -> color, opacity, true
                            | FillNone -> auto, 1.0, false
                            | FillAuto -> auto, (if plot.Kind = BubblePlot || plot.Kind = RadarPlot then 0.75 else 1.0), true
                        let lineColor, lineWidth, dash =
                            match strokeColor source.Line with
                            | Some (color, strokeWidth, dash) -> Some color, strokeWidth, dash
                            | None -> None, 0.0, None
                        { Color = color; Opacity = opacity; Filled = filled; LineColor = lineColor; LineWidth = lineWidth; LineDash = dash }
                let name = textsOf data source.Name |> String.concat " "
                { Source = source; Plot = plot; Ordinal = at
                  Label = (if name.Trim() = "" then sprintf "Series%d" (at + 1) else name)
                  Values = numbersOf data source.Values
                  XValues = (let values = numbersOf data source.Values in fst (xValuesOf data source.XValues values.Length (numbersOf data) (textsOf data)))
                  Sizes = numbersOf data source.BubbleSizes
                  Look = look
                  FormatCode = formatCodeOf data source.Values
                  XFormatCode = (if snd (xValuesOf data source.XValues 0 (numbersOf data) (textsOf data)) then None else formatCodeOf data source.XValues)
                  Secondary = secondary }))
    let firstPlot = chart.Plots |> List.tryHead
    let isPie = match firstPlot with Some plot -> plot.Kind = PiePlot || plot.Kind = DoughnutPlot | None -> false
    let isScatter = match firstPlot with Some plot -> pointPlot plot | None -> false
    let isRadar = match firstPlot with Some plot -> plot.Kind = RadarPlot | None -> false
    let categories : string [] =
        let stated = series |> List.tryPick (fun item -> match textsOf data item.Source.Categories with [||] -> None | texts -> Some texts)
        let count = series |> List.fold (fun most item -> max most item.Values.Length) 0
        match stated with
        | Some texts -> Array.init (max count texts.Length) (fun index -> if index < texts.Length then texts.[index] else "")
        | None -> Array.init count (fun index -> string (index + 1))
    let pointCount = max 1 categories.Length
    let pointColor (item: SeriesData) (index: int) : string * float =
        match Map.tryFind index item.Source.PointFills with
        | Some (FillColor (color, opacity)) -> color, opacity
        | _ -> if item.Plot.VaryColors then styleColor theme chart.Style index pointCount, 1.0 else item.Look.Color, item.Look.Opacity

    // ---- title ----
    let title =
        match chart.Title with
        | Some text -> Some text
        | None when not chart.AutoTitleDeleted && series.Length = 1 -> Some series.Head.Label
        | None -> None
    // (Excel 2019 lays a chart out from a margin: the plot area's box - with its axes' labels - 11 pt in from
    // the chart's edges, a title or a legend 6 pt in, and 11 pt between them and the plot area.)
    let edge, gap = px 6.0, px 11.0
    // (A chartEx chart - a histogram, a box and whisker chart - is laid out tighter: its title 4.5 pt from the top, its
    // plot area's box 6 pt in from the chart's edges and 8 pt below a title, its labels' GDI widths 6 pt from a value
    // axis and 6.5 pt below the categories' - Excel 2019's pictures of synthetic-chartex-zoo, 360 x 216 pt, and of
    // build-pareto-layout.vbs' Pareto charts of labels 1 to 7 figures wide: their plots from 36, 42, 49 .. 67 px.)
    let margin = if chart.ChartEx then px 6.0 else gap
    let gap = if chart.ChartEx then px 8.0 else gap
    let edge = if chart.ChartEx then px 4.5 else edge
    let valueGap, categoryGap = (if chart.ChartEx then px 6.0, px 6.5 else px 3.85, px 1.5)
    let mutable top = margin
    (match title with
     | Some text ->
        // (A title of its own text that names no size or weight: 10 pt bold, as Excel draws it; one made
        // from the series, 14 pt.)
        let size = textSize chart.TitleStyle (if chart.Title.IsSome then 10.0 else 14.0)
        let bold = chart.TitleStyle.Bold |> Option.orElse chart.TextStyle.Bold |> Option.defaultValue chart.Title.IsSome
        let lines = text.Split('\n')
        // (A title of runs of their own looks - phpsp-32's - is drawn run by run, a line as high as its largest.)
        let runs = if chart.Title.IsSome && mixedRuns chart.TitleRuns && chart.TitleRuns.Length = lines.Length then Some chart.TitleRuns else None
        let lineHeights =
            match runs with
            | Some runs -> runs |> List.map (fun line -> line |> List.fold (fun most (_, style) -> max most (lineFor style (textSize style 10.0))) (lineFor chart.TitleStyle size)) |> Array.ofList
            | None -> Array.create lines.Length (lineFor chart.TitleStyle size)
        // (Its box: its lines and 1.5 pt above and below them.)
        let blockHeight = Array.sum lineHeights + px 3.0
        let widthOfLine (index: int) =
            match runs with
            | Some runs -> runs.[index] |> List.sumBy (fun (text, style) -> let runSize = textSize style 10.0 in widthIn style runSize (style.Bold |> Option.defaultValue bold) text)
            | None -> widthIn chart.TitleStyle size bold lines.[index]
        let widest = [ 0 .. lines.Length - 1 ] |> List.fold (fun most index -> max most (widthOfLine index)) 0.0
        // Where the file puts it (`c:layout`): from the chart's edges, or moved from its own place.
        let x, anchor =
            match chart.TitleLayout with
            | Some { XEdge = true; X = Some x } -> x * width + 4.0, "start"
            | Some { X = Some x } -> width / 2.0 + x * width, "middle"
            | _ -> width / 2.0, "middle"
        // (A title is kept inside the chart.)
        let x = if anchor = "start" then max 2.0 (min x (width - widest - 2.0)) else x
        let y =
            match chart.TitleLayout with
            | Some { YEdge = true; Y = Some y } -> y * height + 3.0
            | Some { Y = Some y } -> edge + y * height
            | _ -> edge
        let y = max 0.0 (min y (height - blockHeight))
        let mutable lineTop = y + px 1.5
        lines |> Array.iteri (fun index line ->
            let middle = lineTop + lineHeights.[index] / 2.0
            (match runs with
             | Some runs -> drawRuns runs.[index] 10.0 bold x middle anchor 0.0
             | None -> addText chart.TitleStyle (PText (x, middle, line, size, textColor chart.TitleStyle, anchor, bold, (chart.TitleStyle.Italic |> Option.defaultValue false), 0.0)))
            lineTop <- lineTop + lineHeights.[index])
        // (Room is made for a title at the top that is not laid over the plot.)
        if not chart.TitleOverlay && y < height * 0.4 then top <- max top (y + blockHeight + gap)
     | None -> ())

    /// A marker of a series' points (also in its legend key).
    let drawMarker (symbol: string) (x: float) (y: float) (size: float) (fill: string option) (stroke: string option) =
        match symbol with
        | "square" -> add (PRect (x - size / 2.0, y - size / 2.0, size, size, fill, 1.0, stroke, 0.75, 0.0))
        | "diamond" -> add (PPolyline ([ x, y - size / 1.6; x + size / 1.6, y; x, y + size / 1.6; x - size / 1.6, y ], (fill |> Option.orElse (Some "#FFFFFF")), 1.0, stroke, 0.75, None))
        | "triangle" -> add (PPolyline ([ x, y - size / 1.6; x + size / 1.6, y + size / 2.0; x - size / 1.6, y + size / 2.0 ], (fill |> Option.orElse (Some "#FFFFFF")), 1.0, stroke, 0.75, None))
        | "dash" -> add (PLine (x - size / 2.0, y, x + size / 2.0, y, (stroke |> Option.orElse fill |> Option.defaultValue "#595959"), 1.5, None))
        | "x" | "star" | "plus" ->
            let color = stroke |> Option.orElse fill |> Option.defaultValue "#595959"
            if symbol <> "x" then
                add (PLine (x - size / 2.0, y, x + size / 2.0, y, color, 1.0, None))
                add (PLine (x, y - size / 2.0, x, y + size / 2.0, color, 1.0, None))
            if symbol <> "plus" then
                add (PLine (x - size / 2.0, y - size / 2.0, x + size / 2.0, y + size / 2.0, color, 1.0, None))
                add (PLine (x - size / 2.0, y + size / 2.0, x + size / 2.0, y - size / 2.0, color, 1.0, None))
        | _ -> add (PEllipse (x, y, size / 2.0, size / 2.0, fill, 1.0, stroke, 0.75))
    /// A series' marker: its symbol and size (pixels), None for none. A series that names no symbol (or "auto") where
    /// the plot draws markers has Excel's next of diamond, square, triangle, x, star, circle, plus, dash; one that names
    /// no size the chart style's: 6 pt for styles 1 - 8, 10 pt for 25 - 32, else 8 pt (synthetic-chart-looks "Markers"
    /// under every style through COM, `chart-marker-style-probe.vbs`).
    let markerOf (item: SeriesData) (shown: bool) : (string * float) option =
        let symbol =
            match item.Source.MarkerSymbol with
            | Some "auto" -> [| "diamond"; "square"; "triangle"; "x"; "star"; "circle"; "plus"; "dash" |].[item.Ordinal % 8]
            | None when shown -> [| "diamond"; "square"; "triangle"; "x"; "star"; "circle"; "plus"; "dash" |].[item.Ordinal % 8]
            | Some symbol -> symbol
            | None -> "none"
        if symbol = "none" then None
        else
            let points =
                if item.Source.MarkerSize > 0.0 then item.Source.MarkerSize
                elif chart.Style <= 8 then 6.0
                elif chart.Style >= 25 && chart.Style <= 32 then 10.0
                else 8.0
            Some (symbol, max 3.0 (points * 4.0 / 3.0))
    // ---- legend ----
    let legendEntries : (string * string * KeyLook * TextStyle) list =
        // (text, colour, its key, the look of its text)
        let styleAt (index: int) =
            match chart.Legend with
            | Some legend -> (match Map.tryFind index legend.EntryStyles with Some own -> over own legend.Style | None -> legend.Style)
            | None -> noTextStyle
        let all =
            if isPie then
                match series with
                | first :: _ ->
                    categories |> Array.toList |> List.mapi (fun index text ->
                        let outline =
                            match (match Map.tryFind index first.Source.PointLines with Some own -> own | None -> first.Source.Line) |> strokeColor with
                            | Some (color, strokeWidth, _) -> Some (color, strokeWidth)
                            | None -> None
                        text, fst (pointColor first index), { Line = None; Marker = None; Outline = outline }, styleAt index)
                | [] -> []
            else
                series |> List.mapi (fun index item ->
                    let lined = item.Plot.Kind = LinePlot || item.Plot.Kind = ScatterPlot || item.Plot.Kind = StockPlot || (item.Plot.Kind = RadarPlot && item.Plot.RadarStyle <> "filled")
                    let key =
                        if lined then
                            let shown =
                                match item.Plot.Kind with
                                | LinePlot -> item.Plot.Markers
                                | ScatterPlot -> item.Plot.ScatterStyle.EndsWith "Marker" || item.Plot.ScatterStyle = "marker"
                                | _ -> false
                            let marker =
                                markerOf item shown |> Option.map (fun (symbol, size) ->
                                    let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                                    let stroke = (match strokeColor item.Source.MarkerLine with Some (color, _, _) -> Some color | None -> Some item.Look.Color)
                                    symbol, size, fill, stroke)
                            { Line = (if item.Look.LineColor.IsSome && item.Look.LineWidth > 0.0 then Some item.Look.LineWidth else None); Marker = marker; Outline = None }
                        else
                            { Line = None; Marker = None
                              Outline = (match item.Look.LineColor with Some color when item.Look.LineWidth > 0.0 -> Some (color, item.Look.LineWidth) | _ -> None) }
                    item.Label, (match item.Look.LineColor with Some color when not item.Look.Filled -> color | _ -> item.Look.Color), key, styleAt index)
        match chart.Legend with
        | Some legend when not legend.Deleted.IsEmpty -> all |> List.indexed |> List.filter (fun (index, _) -> not (Set.contains index legend.Deleted)) |> List.map snd
        | _ -> all
    // (A chartEx chart with a second value axis - a Pareto chart's running share - keeps 12 pt at its sides, not 6:
    // build-pareto-layout.vbs, Excel 2019's pictures of Pareto charts whose labels are 1 to 7 figures wide - the left
    // labels from 17 px, the right ones to 462 of 480, whatever their widths.)
    let sideMargin =
        if chart.ChartEx && chart.Plots |> List.exists (fun plot -> plot.AxisIds |> List.contains "2") then px 12.0 else margin
    let mutable left = sideMargin
    let mutable right = width - sideMargin
    let mutable bottom = height - margin
    (match chart.Legend with
     | Some legend when not legendEntries.IsEmpty ->
        // (Excel 2019's legend, checked on charts made in Excel: an entry is a line of its text; entries one under
        // another are 5.85 pt apart, the box 3 pt round them; a legend at the side is centred on the plot area's box,
        // one at the bottom or the top on its width; a box of the file's own size shares its height out evenly.)
        let sizeOf (style: TextStyle) = textSize style 10.0
        let size = sizeOf legend.Style
        // (A square key 0.55 of the text's size - 4.94 pt at 9 pt, Excel 2019's LegendKey through COM -; an entry as
        // long as its text by the face's design advances, as Excel lays chart text out: synthetic-chart-looks
        // "LegendBottom", "Sales" 24.7 px of Calibri 9 pt, GDI's 30.)
        let swatch = size * 0.55
        let entryTextWidth (style: TextStyle) (textSize: float) (bold: bool) (text: string) =
            match ExcelText.idealWidth (defaultArg (faceOf style) "Calibri") bold (style.Italic |> Option.defaultValue false) textSize text with
            | Some ideal -> ideal
            | None -> widthIn style textSize bold text
        // (The legend's own fill and outline, `c:spPr`, under its entries.)
        let legendAt = items.Count
        let boxed (boxX: float) (boxY: float) (boxW: float) (boxH: float) =
            let fill = match legend.Fill with FillColor (color, opacity) -> Some (color, opacity) | _ -> None
            match fill, strokeColor legend.Border with
            | None, None -> ()
            | fill, border ->
                let strokeWidth = match border with Some (_, strokeWidth, _) -> strokeWidth | None -> 0.0
                items.Insert(legendAt, PRect (boxX, boxY, boxW, boxH, fill |> Option.map fst, (match fill with Some (_, opacity) -> opacity | None -> 1.0),
                                              border |> Option.map (fun (color, _, _) -> color), strokeWidth, 0.0))
        // (Excel 2019's keys, synthetic-chart-lines "Widths" / "Keys9" / "Keys14": in a legend with a line every key
        // is 26 px long - a line of the series' width with its marker in the middle, a box as high as a swatch -
        // whatever the size of the text; else a square swatch.)
        let lineKeys = legendEntries |> List.exists (fun (_, _, key, _) -> key.Line.IsSome || key.Marker.IsSome)
        let keyWidth = if lineKeys then 26.0 else swatch
        let entryWidth (text: string, _: string, _: KeyLook, style: TextStyle) = keyWidth + px 3.0 + entryTextWidth style (sizeOf style) (style.Bold |> Option.defaultValue false) text
        let pitch = lineFor legend.Style size + px 5.85
        let padding = px 3.0
        let drawEntryIn (wrap: bool) (x: float) (y: float) (room: float) (text: string, entryColor: string, key: KeyLook, style: TextStyle) =
            let textSize = sizeOf style
            let bold = style.Bold |> Option.defaultValue false
            let textRoom = room - keyWidth - px 3.0
            // (In a legend box of the file's own size a text too long for it is broken at its spaces, its key by
            // the first line, the lines about the entry's middle - dashboard-ecommerce-sales' "KI Adjustable- /
            // Height Table", "GBC Premium / Transparent / Covers"; elsewhere it ends in an ellipsis.)
            let lines =
                if wrap && textRoom > 12.0 && entryTextWidth style textSize bold text > textRoom + 2.0 && text.Trim().Contains " " then
                    let broken = ResizeArray<string>()
                    let mutable current = ""
                    for word in text.Split(' ') |> Array.filter (fun word -> word <> "") do
                        let joined = if current = "" then word else current + " " + word
                        if current = "" || entryTextWidth style textSize bold joined <= textRoom then current <- joined
                        else
                            broken.Add current
                            current <- word
                    broken.Add current
                    List.ofSeq broken
                else [ text ]
            let y = y - float (lines.Length - 1) * lineFor style textSize / 2.0
            match key.Line, key.Marker with
            | None, None ->
                let outline, outlineWidth = match key.Outline with Some (color, strokeWidth) -> Some color, strokeWidth | None -> None, 0.0
                add (PRect (x, y - swatch / 2.0, keyWidth, swatch, Some entryColor, 1.0, outline, outlineWidth, 0.0))
            | line, marker ->
                (match line with Some strokeWidth -> add (PLine (x, y, x + keyWidth, y, entryColor, strokeWidth, None)) | None -> ())
                (match marker with Some (symbol, markerSize, fill, stroke) -> drawMarker symbol (x + keyWidth / 2.0) y (min markerSize (swatch * 1.4)) fill stroke | None -> ())
            lines |> List.iteri (fun index text ->
                // (A text longer than its place ends in an ellipsis.)
                let shownText =
                    if textRoom > 12.0 && entryTextWidth style textSize bold text > textRoom + 2.0 && text.Length > 3 then
                        let mutable keep = text.Length - 1
                        while keep > 1 && entryTextWidth style textSize bold (text.Substring(0, keep) + "…") > textRoom do keep <- keep - 1
                        text.Substring(0, keep) + "…"
                    else text
                addText style (PText (x + keyWidth + px 3.0, y + float index * lineFor style textSize, shownText, textSize, textColor style, "start", bold, (style.Italic |> Option.defaultValue false), 0.0)))
        let drawEntry = drawEntryIn false
        let horizontal = legend.Position = "b" || legend.Position = "t"
        // The rectangle the file gives the legend (`c:layout`), or the side of the chart it is at.
        let box =
            match legend.Layout with
            | Some { XEdge = true; YEdge = true; X = Some x; Y = Some y; W = w; H = h } ->
                let boxLeft, boxTop = max 0.0 (min (width - 10.0) (x * width)), max 0.0 (min (height - 10.0) (y * height))
                Some (boxLeft, boxTop, (w |> Option.map (fun value -> min (width - boxLeft) (value * width))), (h |> Option.map (fun value -> min (height - boxTop) (value * height))))
            | _ -> None
        if horizontal then
            // Rows of entries, centred, 5.85 pt apart (Excel 2019's "LegendBottom": keys at 183, 226, 266 px).
            let between = px 5.85
            let room = match box with Some (_, _, Some boxWidth, _) -> boxWidth - 2.0 * padding | _ -> right - left
            let rows = ResizeArray<ResizeArray<string * string * KeyLook * TextStyle>>()
            let mutable current = ResizeArray<string * string * KeyLook * TextStyle>()
            let mutable used = 0.0
            for entry in legendEntries do
                let entryW = entryWidth entry
                if used + entryW > room && current.Count > 0 then
                    rows.Add current
                    current <- ResizeArray<string * string * KeyLook * TextStyle>()
                    used <- 0.0
                current.Add entry
                used <- used + entryW + between
            if current.Count > 0 then rows.Add current
            let rows = rows |> Seq.truncate 6 |> Seq.toList
            let blockHeight = float rows.Length * pitch
            let boxLeft, boxTop, boxWidth, boxHeight =
                match box with
                | Some (boxLeft, boxTop, boxWidth, boxHeight) -> boxLeft, boxTop, (boxWidth |> Option.defaultValue (width - boxLeft)), (boxHeight |> Option.defaultValue blockHeight)
                | None -> left, (if legend.Position = "b" then height - edge - blockHeight else (if top > gap then top - gap + edge else edge)), right - left, blockHeight
            let startY = boxTop + max 0.0 ((boxHeight - blockHeight) / 2.0)
            let totalOf (row: ResizeArray<string * string * KeyLook * TextStyle>) = (row |> Seq.sumBy (fun entry -> entryWidth entry + between)) - between
            (match box with
             | Some _ -> boxed boxLeft boxTop boxWidth boxHeight
             | None ->
                let widestRow = rows |> List.fold (fun most row -> max most (totalOf row)) 0.0
                boxed (max 4.0 (boxLeft + (boxWidth - widestRow) / 2.0) - padding) startY (widestRow + 2.0 * padding) blockHeight)
            rows |> List.iteri (fun rowIndex row ->
                let total = totalOf row
                let mutable x = max 4.0 (boxLeft + (boxWidth - total) / 2.0)
                for entry in row do
                    drawEntry x (startY + pitch * (float rowIndex + 0.5)) (entryWidth entry) entry
                    x <- x + entryWidth entry + between)
            if not legend.Overlay then
                if legend.Position = "b" then bottom <- min bottom (boxTop - gap)
                else top <- max top (boxTop + max blockHeight boxHeight + gap)
        else
            let position = legend.Position
            let widest = legendEntries |> List.fold (fun most entry -> max most (entryWidth entry)) 0.0
            let count = legendEntries.Length
            let boxLeft, boxTop, boxWidth, boxHeight, spread =
                match box with
                | Some (boxLeft, boxTop, boxWidth, boxHeight) ->
                    boxLeft, boxTop, (boxWidth |> Option.defaultValue (min (widest + 2.0 * padding) (width - boxLeft))), (boxHeight |> Option.defaultValue (bottom - boxTop)), boxHeight.IsSome
                | None ->
                    let blockWidth = min (width * 0.4) (widest + 2.0 * padding)
                    let blockHeight = min (bottom - top) (2.0 * padding + float count * lineFor legend.Style size + float (count - 1) * px 5.85)
                    let blockTop = if position = "tr" then top else (top + bottom) / 2.0 - blockHeight / 2.0
                    (if position = "l" then edge else width - edge - blockWidth), blockTop, blockWidth, blockHeight, false
            // (Every entry is there: the lines move together where the legend is short.)
            let rowHeight = if spread then boxHeight / float count else max (size * 1.05) (min pitch ((boxHeight - 2.0 * padding + px 5.85) / float count))
            let shown = legendEntries |> List.truncate (max 1 (int ((boxHeight + 1.0) / max 1.0 (min rowHeight (size * 1.05)))))
            let startY = if spread then boxTop else boxTop + padding - px 5.85 / 2.0
            boxed boxLeft boxTop boxWidth boxHeight
            // (The entries 1.5 pt further in than the box's padding: Excel 2019's "LegendRight", keys 7 px from its box.)
            shown |> List.iteri (fun index entry -> drawEntryIn spread (boxLeft + padding + px 1.5) (startY + rowHeight * (float index + 0.5)) (boxWidth - 2.0 * padding - px 1.5) entry)
            if not legend.Overlay then
                if position = "l" then left <- max left (boxLeft + boxWidth + gap) else right <- min right (boxLeft - gap)
     | _ -> ())

    let labelStyleOf (item: SeriesData) = item.Source.Labels |> Option.orElse item.Plot.Labels
    let labelText (item: SeriesData) (spec: LabelSpec) (index: int) (value: float) (share: float option) : string =
        [ if spec.ShowSeries then yield item.Label
          if spec.ShowCategory && index < categories.Length then yield categories.[index]
          if spec.ShowValue then
              yield (match spec.FormatCode |> Option.orElse item.FormatCode with
                     | Some code -> (try data.Format code value with _ -> generalNumber value)
                     | None -> generalNumber value)
          if spec.ShowPercent then
              match share with
              | Some part ->
                  // (The labels' own number format is the per cent's when only the per cent is shown.)
                  yield (match spec.FormatCode with
                         | Some code when code.Contains "%" && not spec.ShowValue -> (try data.Format code part with _ -> generalNumber (Math.Round (part * 100.0)) + "%")
                         | _ -> generalNumber (Math.Round (part * 100.0)) + "%")
              | None -> () ]
        // (Without a separator of its own, a pie's label is a line for each part - Excel 2019's
        // "Studio\n14%" -, another chart's parts are on one line.)
        |> String.concat (spec.Separator |> Option.defaultValue (if item.Plot.Kind = PiePlot || item.Plot.Kind = DoughnutPlot then "\n" else "; "))
    let labelled (item: SeriesData) (index: int) : LabelSpec option =
        labelStyleOf item
        |> Option.filter (fun spec -> not (Set.contains index spec.Deleted))
        |> Option.map (fun spec -> Map.tryFind index spec.Points |> Option.defaultValue spec)
        |> Option.filter (fun spec -> spec.ShowValue || spec.ShowPercent || spec.ShowCategory || spec.ShowSeries)
    /// A label's lines: its parts, each broken at its spaces to the width Excel gives a label, a fifth of the
    /// chart's (checked: synthetic-chart-parity "PieLabels": "Two / BR/One / BA / 14%" in a chart of 225 pt).
    let wrapLabel (style: TextStyle) (size: float) (bold: bool) (text: string) : string list =
        let room = width / 5.0 - px 6.0
        [ for part in text.Split('\n') do
            let words = ResizeArray<string>()
            let mutable start = 0
            for at in 0 .. part.Length - 1 do
                if part.[at] = ' ' then
                    words.Add (part.Substring(start, at - start + 1))
                    start <- at + 1
            if start < part.Length then words.Add (part.Substring start)
            let mutable line = ""
            for word in words do
                if line <> "" && widthIn style size bold ((line + word).TrimEnd()) > room then
                    yield line.TrimEnd()
                    line <- word
                else line <- line + word
            yield line.TrimEnd() ]
    let plotRect = layoutRect chart.PlotLayout width height

    if isPie then
        // ---- pie, doughnut ----
        let rings = series |> List.filter (fun item -> item.Plot.Kind = PiePlot || item.Plot.Kind = DoughnutPlot)
        let rings = (match firstPlot with Some plot when plot.Kind = PiePlot -> List.truncate 1 rings | _ -> rings)
        // (Labels that may stand outside a pie - at its end, or where they fit best, Excel's choice for a pie's -
        // and any labels of a doughnut make Excel 2019 draw it an eighth of the chart's shorter side smaller,
        // whatever their size: 153 pt for 178 in a chart of 300 x 200 pt, 196.75 for 228 in 400 x 250 and 250 x
        // 300, labels of 8, 10 and 14 pt (pie-probe through COM); a pie of 174.9 pt in 203, 225 pt high.)
        let outside =
            rings |> List.exists (fun item ->
                match labelStyleOf item with
                | Some spec when spec.ShowValue || spec.ShowPercent || spec.ShowCategory || spec.ShowSeries ->
                    item.Plot.Kind = DoughnutPlot || spec.Position = Some "outEnd" || spec.Position = Some "bestFit" || spec.Position.IsNone
                | _ -> false)
        let cx, cy, radius =
            match plotRect with
            | Some (x, y, w, h) -> x + w / 2.0, y + h / 2.0, max 4.0 (min w h / 2.0)
            | None -> (left + right) / 2.0, (top + bottom) / 2.0, max 4.0 (min (right - left) (bottom - top) / 2.0 - (if outside then min width height / 16.0 else 0.0))
        let exploded = rings |> List.fold (fun most item -> max most item.Source.Explosion) 0.0
        let radius = if exploded > 0.0 then radius / (1.0 + min 1.0 (exploded / 100.0)) else radius
        let hole = (match firstPlot with Some plot when plot.Kind = DoughnutPlot -> radius * (max 10.0 (min 90.0 plot.HoleSize)) / 100.0 | _ -> 0.0)
        let ringWidth = if rings.IsEmpty then 0.0 else (radius - hole) / float rings.Length
        rings |> List.iteri (fun ringIndex item ->
            let outer = (if hole > 0.0 then hole + ringWidth * float (ringIndex + 1) else radius)
            let inner = (if hole > 0.0 then hole + ringWidth * float ringIndex else 0.0)
            let values = item.Values |> Array.map (fun value -> match value with Some v when v > 0.0 -> v | _ -> 0.0)
            let total = Array.sum values
            if total > 0.0 then
                let mutable angle = item.Plot.FirstSliceAngle * Math.PI / 180.0
                // (A slice's outline, as Excel 2019 draws synthetic-chart-lines "PiePoints" / "PiePlain": the line of
                // the point's own look where it has one (none when that look states no line), else the series';
                // stating none, the chart style's - a pixel of white for styles 9 - 16, of a shade of the slice for
                // 33 - 40, else none.)
                let separator (index: int) (color: string) =
                    let stated = match Map.tryFind index item.Source.PointLines with Some own -> own | None -> item.Source.Line
                    match stated, strokeColor stated with
                    | Some { Paint = FillNone }, _ -> None, 0.0
                    | _, Some (lineColor, strokeWidth, _) -> Some lineColor, strokeWidth
                    | _ when chart.Style >= 9 && chart.Style <= 16 -> Some "#FFFFFF", 1.0
                    | _ when chart.Style >= 33 && chart.Style <= 40 -> (let r, g, b = rgbOf color in Some ("#" + hex2 (r * 0.75) + hex2 (g * 0.75) + hex2 (b * 0.75))), 1.0
                    | _ -> None, 0.0
                values |> Array.iteri (fun index value ->
                    if value > 0.0 then
                        let sweep = value / total * 2.0 * Math.PI
                        let color, opacity = pointColor item index
                        let middle = angle + sweep / 2.0
                        // (A slice pulled out moves along its middle.)
                        let shift = radius * (min 100.0 item.Source.Explosion) / 100.0
                        let cx, cy = cx + shift * Math.Sin middle, cy - shift * Math.Cos middle
                        // (Angles run clockwise from twelve o'clock.)
                        let point (r: float) (a: float) = cx + r * Math.Sin a, cy - r * Math.Cos a
                        let fromAngle, toAngle = angle, angle + sweep
                        if sweep >= 2.0 * Math.PI - 1e-9 then
                            if inner > 0.0 then
                                let half = fromAngle + Math.PI
                                let (ax, ay), (bx, by) = point outer fromAngle, point outer half
                                let (ix, iy), (jx, jy) = point inner fromAngle, point inner half
                                add (PPath (sprintf "M%s %sA%s %s 0 1 1 %s %sA%s %s 0 1 1 %s %sZM%s %sA%s %s 0 1 0 %s %sA%s %s 0 1 0 %s %sZ"
                                                (fmt ax) (fmt ay) (fmt outer) (fmt outer) (fmt bx) (fmt by) (fmt outer) (fmt outer) (fmt ax) (fmt ay)
                                                (fmt ix) (fmt iy) (fmt inner) (fmt inner) (fmt jx) (fmt jy) (fmt inner) (fmt inner) (fmt ix) (fmt iy), Some color, opacity, None, 0.0))
                            else add (PEllipse (cx, cy, outer, outer, Some color, opacity, None, 0.0))
                        else
                            let large = if sweep > Math.PI then 1 else 0
                            let (ax, ay), (bx, by) = point outer fromAngle, point outer toAngle
                            let d =
                                if inner > 0.0 then
                                    let (ix, iy), (jx, jy) = point inner toAngle, point inner fromAngle
                                    sprintf "M%s %sA%s %s 0 %d 1 %s %sL%s %sA%s %s 0 %d 0 %s %sZ" (fmt ax) (fmt ay) (fmt outer) (fmt outer) large (fmt bx) (fmt by) (fmt ix) (fmt iy) (fmt inner) (fmt inner) large (fmt jx) (fmt jy)
                                else sprintf "M%s %sL%s %sA%s %s 0 %d 1 %s %sZ" (fmt cx) (fmt cy) (fmt ax) (fmt ay) (fmt outer) (fmt outer) large (fmt bx) (fmt by)
                            let outline, outlineWidth = separator index color
                            add (PPath (d, Some color, opacity, outline, outlineWidth))
                        (match labelled item index with
                         | Some spec ->
                            let size = textSize spec.Style 10.0
                            let bold = spec.Style.Bold |> Option.defaultValue false
                            let lines = wrapLabel spec.Style size bold (labelText item spec index value (Some (value / total)))
                            let lineHeight = lineFor spec.Style size
                            let boxWidth = (lines |> List.fold (fun most line -> max most (widthIn spec.Style size bold line)) 0.0) + px 6.0
                            let boxHeight = float lines.Length * lineHeight + px 3.0
                            // (Where a label stands: in the slice, or outside its end - the box just clear of the
                            // circle; where it fits best is in the slice when the slice holds it.)
                            let reach = abs (Math.Sin middle) * boxWidth / 2.0 + abs (Math.Cos middle) * boxHeight / 2.0
                            let fitsInside =
                                let at = outer * 0.6
                                sweep * at >= min boxWidth boxHeight + px 2.0 && at + reach <= outer && (sweep >= Math.PI || 2.0 * at * Math.Sin (sweep / 2.0) >= max boxWidth boxHeight * 0.8)
                            // (A label moved by hand is moved from where the series' labels stand: A.CRE's "Occupied
                            // Units" of best fit, moved a little, is outside its 75 % slice as the others are.)
                            let position = if spec.Offset.IsSome then (labelStyleOf item |> Option.bind (fun series -> series.Position) |> Option.orElse spec.Position) else spec.Position
                            let outsideLabel = position = Some "outEnd" || (position = Some "bestFit" && not fitsInside)
                            let at =
                                if outsideLabel then outer + px 2.0 + reach
                                elif inner > 0.0 then (inner + outer) / 2.0
                                elif position = Some "inEnd" then outer - reach - px 2.0
                                elif position = Some "ctr" then outer * 0.5
                                else outer * 0.6
                            let x, y = point at middle
                            let x, y = (match spec.Offset with Some (dx, dy) -> x + dx * width, y + dy * height | None -> x, y)
                            // (Kept inside the chart.)
                            let x = max (boxWidth / 2.0) (min (width - boxWidth / 2.0) x)
                            let y = max (boxHeight / 2.0) (min (height - boxHeight / 2.0) y)
                            let color = spec.Style.Color |> Option.orElse chart.TextStyle.Color |> Option.defaultValue (if outsideLabel then "#404040" else "#FFFFFF")
                            lines |> List.iteri (fun lineIndex line ->
                                addText spec.Style (PText (x, y - boxHeight / 2.0 + px 1.5 + (float lineIndex + 0.5) * lineHeight, line, size, color, "middle", bold, (spec.Style.Italic |> Option.defaultValue false), 0.0)))
                         | _ -> ())
                        angle <- angle + sweep))
    elif isRadar then
        // ---- radar: the categories around a centre, the values along the spokes ----
        let count = pointCount
        let lo, hi =
            let values = series |> List.collect (fun item -> item.Values |> Array.choose id |> Array.toList)
            if values.IsEmpty then 0.0, 1.0 else List.min values, List.max values
        let vMin, vMax, vStep = valueScale lo hi (primaryValueAxis |> Option.bind (fun a -> a.Min)) (primaryValueAxis |> Option.bind (fun a -> a.Max)) (primaryValueAxis |> Option.bind (fun a -> a.MajorUnit)) 10
        let scale = { Min = vMin; Max = vMax; Step = vStep; Log = None }
        let categoryAxis = chart.Axes |> List.tryFind (fun axis -> axis.Kind = "cat" || axis.Kind = "date")
        let labelSize = textSize (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let labelColor = textColor (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle)
        let widestLabel = categories |> Array.fold (fun most text -> max most (textWidth labelSize text)) 0.0
        let cx, cy, radius =
            match plotRect with
            | Some (x, y, w, h) -> x + w / 2.0, y + h / 2.0, max 4.0 (min w h / 2.0)
            | None ->
                (left + right) / 2.0, (top + bottom) / 2.0,
                max 4.0 (min ((right - left) / 2.0 - min widestLabel ((right - left) * 0.25) - 6.0) ((bottom - top) / 2.0 - labelSize * 1.4))
        let spoke (index: int) = float index * 2.0 * Math.PI / float count
        let pointAt (index: int) (share: float) = cx + radius * share * Math.Sin (spoke index), cy - radius * share * Math.Cos (spoke index)
        // The rings of the value axis and the spokes.
        let gridColor, gridWidth =
            match primaryValueAxis |> Option.bind (fun axis -> strokeColor axis.Gridlines) with
            | Some (color, strokeWidth, _) -> color, strokeWidth
            | None -> "#D9D9D9", 1.0
        let valueLabelSize = textSize (primaryValueAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let format = series |> List.tryPick (fun item -> item.FormatCode)
        for tick in ticksOf scale do
            let share = shareOn scale tick
            if share > 0.0 then add (PPolyline ([ for index in 0 .. count -> pointAt (index % count) share ], None, 1.0, Some gridColor, gridWidth, None))
            if (match primaryValueAxis with Some axis -> not axis.Deleted && axis.LabelsShown | None -> true) then
                let text = (match format with Some code -> (try data.Format code tick with _ -> generalNumber tick) | None -> generalNumber tick)
                add (PText (cx - 4.0, cy - radius * share, text, valueLabelSize, labelColor, "end", false, false, 0.0))
        for index in 0 .. count - 1 do
            let x, y = pointAt index 1.0
            add (PLine (cx, cy, x, y, gridColor, gridWidth, None))
            let lx, ly = cx + (radius + labelSize * 0.9) * Math.Sin (spoke index), cy - (radius + labelSize * 0.9) * Math.Cos (spoke index)
            let anchor = if Math.Sin (spoke index) > 0.2 then "start" elif Math.Sin (spoke index) < -0.2 then "end" else "middle"
            add (PText (lx, ly, categories.[index], labelSize, labelColor, anchor, false, false, 0.0))
        for item in series do
            let points = [ for index in 0 .. count - 1 do
                            match (if index < item.Values.Length then item.Values.[index] else None) with
                            | Some value -> yield pointAt index (shareOn scale value)
                            | None -> () ]
            if points.Length > 1 then
                let closed = points @ [ points.Head ]
                if item.Plot.RadarStyle = "filled" then
                    add (PPolyline (points, (if item.Look.Filled then Some item.Look.Color else None), item.Look.Opacity, Some (item.Look.LineColor |> Option.defaultValue item.Look.Color), max 1.0 item.Look.LineWidth, None))
                else
                    (match item.Look.LineColor with
                     | Some color -> add (PPolyline (closed, None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash))
                     | None -> ())
                    if item.Plot.RadarStyle = "marker" && item.Source.MarkerSymbol <> Some "none" then
                        for x, y in points do add (PEllipse (x, y, 3.0, 3.0, Some item.Look.Color, 1.0, Some item.Look.Color, 0.75))
    else
        // ---- charts on axes ----
        let horizontal = series |> List.exists (fun item -> item.Plot.Kind = BarPlot && item.Plot.Horizontal)
        let categoryAxis =
            chart.Axes |> List.tryFind (fun axis -> (axis.Kind = "cat" || axis.Kind = "date") && not axis.Deleted)
            |> Option.orElse (chart.Axes |> List.tryFind (fun axis -> axis.Kind = "cat" || axis.Kind = "date"))
        let secondaryAxis =
            match primaryValueAxis with
            | Some first -> valueAxes |> List.tryFind (fun axis -> axis.Id <> first.Id && series |> List.exists (fun item -> item.Secondary && List.contains axis.Id item.Plot.AxisIds))
            | None -> None
        let xAxis = if isScatter then valueAxes |> List.tryFind (fun axis -> axis.Position = "b" || axis.Position = "t") else None
        // A date axis (`c:dateAx`) over categories that are dates: a band for every day, month or year from the
        // first date to the last (`c:baseTimeUnit`, else the finest the dates need), each point in the band of
        // its date, a label every major unit (`c:majorUnit` of `c:majorTimeUnit`) from the first, in the axis'
        // number format - as Excel 2019 draws vertex42's mortgage balance (checked: synthetic-chart-parity
        // "DateLine", the axis from the first date to the last).
        let dated : (float option [] * string * int) option =
            match categoryAxis with
            | Some axis when axis.Kind = "date" && not isScatter && not horizontal ->
                let dates = series |> List.tryPick (fun item -> match item.Source.Categories with Some source -> Some (numbersOf data (Some source)) | None -> None)
                match dates with
                | Some dates when dates |> Array.exists (fun date -> match date with Some value -> value >= 1.0 | None -> false) ->
                    let days = dates |> Array.choose id |> Array.filter (fun value -> value >= 1.0) |> Array.map (fun value -> DateTime(1899, 12, 30).AddDays(Math.Floor value))
                    let unit =
                        match axis.BaseTimeUnit with
                        | Some unit -> unit
                        | None ->
                            if days |> Array.forall (fun day -> day.Day = 1 && day.Month = days.[0].Month) then "years"
                            elif days |> Array.forall (fun day -> day.Day = days.[0].Day) then "months"
                            else "days"
                    Some (dates, unit, 0)
                | _ -> None
            | _ -> None
        let unitOf (unit: string) (serial: float) =
            let day = DateTime(1899, 12, 30).AddDays(Math.Floor serial)
            match unit with
            | "years" -> day.Year
            | "months" -> day.Year * 12 + day.Month - 1
            | _ -> int (Math.Floor serial)
        let series, categories =
            match dated with
            | Some (dates, unit, _) ->
                let units = dates |> Array.map (Option.filter (fun value -> value >= 1.0) >> Option.map (unitOf unit))
                let first = units |> Array.choose id |> Array.min
                let last = units |> Array.choose id |> Array.max
                let bands = last - first + 1
                let moved (values: float option []) =
                    let placed = Array.create bands None
                    values |> Array.iteri (fun index value ->
                        match (if index < units.Length then units.[index] else None), value with
                        | Some at, Some v -> placed.[at - first] <- Some v
                        | _ -> ())
                    placed
                let axis = categoryAxis.Value
                let step = max 1 (int (defaultArg axis.MajorUnit 1.0))
                let stepUnits =
                    match unit, defaultArg axis.MajorTimeUnit unit with
                    | "months", "years" -> step * 12
                    | "days", "months" -> step * 30
                    | "days", "years" -> step * 365
                    | _ -> step
                let code = (match axis.FormatCode with Some code when not axis.SourceLinked && code <> "General" -> code | _ -> "m/d/yyyy")
                let labelOf (band: int) =
                    let serial =
                        match unit with
                        | "years" -> (DateTime(first + band, 1, 1) - DateTime(1899, 12, 30)).TotalDays
                        | "months" -> let m = first + band in (DateTime(m / 12, m % 12 + 1, 1) - DateTime(1899, 12, 30)).TotalDays
                        | _ -> float (first + band)
                    try data.Format code serial with _ -> generalNumber serial
                (series |> List.map (fun item -> { item with Values = moved item.Values })),
                Array.init bands (fun band -> if band % stepUnits = 0 then labelOf band else "")
            | None -> series, categories
        let count = max 1 categories.Length
        // What a series reaches on its value axis, stacked ones by their sums.
        let stackedOf (plot: Plot) = plot.Grouping = "stacked" || plot.Grouping = "percentStacked"
        let groupTotals (members: SeriesData list) (percent: bool) : (float [] * float []) =
            // (the sum of the positive values and of the negative ones, per category)
            let positive = Array.zeroCreate<float> count
            let negative = Array.zeroCreate<float> count
            for item in members do
                item.Values |> Array.iteri (fun index value ->
                    match value with
                    | Some v when index < count -> if v >= 0.0 then positive.[index] <- positive.[index] + v else negative.[index] <- negative.[index] + v
                    | _ -> ())
            if percent then
                for index in 0 .. count - 1 do
                    let whole = positive.[index] - negative.[index]
                    if whole > 0.0 then
                        positive.[index] <- positive.[index] / whole
                        negative.[index] <- negative.[index] / whole
            positive, negative
        let extent (secondary: bool) (positiveOnly: bool) : (float * float) option =
            let mutable lo = Double.PositiveInfinity
            let mutable hi = Double.NegativeInfinity
            let see (value: float) = if not positiveOnly || value > 0.0 then ((if value < lo then lo <- value); (if value > hi then hi <- value))
            for plot in chart.Plots do
                let members = series |> List.filter (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary)
                if not members.IsEmpty then
                    if stackedOf plot then
                        let positive, negative = groupTotals members (plot.Grouping = "percentStacked")
                        positive |> Array.iter see
                        negative |> Array.iter see
                    else
                        match plot.Kind with
                        | BoxesPlot (figures, options) ->
                            // (A box and whisker chart's axis spans the whiskers, and the points shown beyond them.)
                            for boxes in figures do
                                for box in boxes do
                                    match box with
                                    | Some box ->
                                        see box.Low
                                        see box.High
                                        if options.Outliers then box.Outside |> List.iter see
                                    | None -> ()
                        | _ ->
                            for item in members do
                                for value in item.Values do (match value with Some v -> see v | None -> ())
            if lo > hi then None else Some (lo, hi)
        /// A logarithmic scale over lo .. hi (both above zero): whole powers of the base.
        let logScale (logBase: float) (lo: float) (hi: float) (axis: Axis) : AxisScale =
            let exponent (value: float) = Math.Log value / Math.Log logBase
            let low = (match axis.Min with Some value when value > 0.0 -> exponent value | _ -> Math.Floor (exponent lo + 1e-9))
            let high = (match axis.Max with Some value when value > 0.0 -> exponent value | _ -> Math.Ceiling (exponent hi - 1e-9))
            let high = if high <= low then low + 1.0 else high
            { Min = low; Max = high; Step = max 1.0 (Math.Ceiling ((high - low) / 12.0)); Log = Some logBase }
        let scaleFor (axis: Axis option) (secondary: bool) (most: int) : AxisScale =
            let percent = chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            match axis with
            | Some ({ LogBase = Some logBase } as logAxis) when not percent ->
                (match extent secondary true with
                 | Some (lo, hi) -> logScale logBase lo hi logAxis
                 | None -> { Min = 0.0; Max = 1.0; Step = 1.0; Log = Some logBase })
            | _ ->
                let fixedOf (pick: Axis -> float option) = axis |> Option.bind pick
                match extent secondary false with
                | Some (lo, _) when percent ->
                    { Min = (if lo < 0.0 then -1.0 else 0.0); Max = 1.0; Step = 0.1; Log = None }
                | found ->
                    // (No values at all: as values from 0 to 1 - Excel 2019's 0 .. 1.2 on a short axis.)
                    let lo, hi = defaultArg found (0.0, 1.0)
                    let minimum, maximum, step = valueScale lo hi (fixedOf (fun a -> a.Min)) (fixedOf (fun a -> a.Max)) (fixedOf (fun a -> a.MajorUnit)) most
                    { Min = minimum; Max = maximum; Step = step; Log = None }
        let xFormat = series |> List.tryPick (fun item -> item.XFormatCode)
        let axisFormat (axis: Axis option) (secondary: bool) (isX: bool) : float -> string =
            let percent = not isX && chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            // (An axis linked to its source has the format of the values it shows: the x values' along the bottom
            // of a scatter chart, the y values' up its side.)
            let linked = if isX then xFormat else series |> List.tryPick (fun item -> if item.Secondary = secondary then item.FormatCode else None)
            let code =
                match axis with
                | Some axis when not axis.SourceLinked -> axis.FormatCode
                | Some axis -> linked |> Option.orElse axis.FormatCode
                | None -> linked
            let code = code |> Option.filter (fun code -> code <> "General")
            // (The labels of an axis in thousands, millions ... say the value divided by that.)
            let unit = match axis with Some axis when axis.DisplayUnit > 0.0 -> axis.DisplayUnit | _ -> 1.0
            match code with
            | _ when percent -> (fun value -> generalNumber (Math.Round (value * 100.0)) + "%")
            | Some code -> (fun value -> try data.Format code (value / unit) with _ -> generalNumber (value / unit))
            | None -> (fun value -> generalNumber (value / unit))
        let shown (axis: Axis option) = match axis with Some axis -> not axis.Deleted && axis.LabelsShown | None -> false
        let styleOf (axis: Axis option) = axis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle
        let valueStyle, categoryStyle = styleOf primaryValueAxis, styleOf categoryAxis
        let xStyle = styleOf xAxis
        let valueSize = textSize valueStyle 10.0
        let categorySize = textSize categoryStyle 10.0
        let valueColor = textColor valueStyle
        let categoryColor = textColor categoryStyle
        let widest (style: TextStyle) (labels: string list) (size: float) = labels |> List.fold (fun most text -> max most (widthIn style size false text)) 0.0
        // How the labels of a category axis along the bottom stand, as Excel 2019 sets them (checked on 13,000
        // column charts made in Excel - labels of numbers, "Cat 1", "Category 1", "Q1 2000" and of mixed lengths,
        // 6 - 100 categories, 8 - 11 pt, 90 - 760 pt wide; tools/excel-parity chart-probe.vbs, skip-probe): every
        // s-th label, s the least for which the labels stand in s bands - in a line when the widest is no wider
        // than s bands; else broken at spaces when the widest word is not; else (where the file leaves the angle to
        // Excel) turned -45 degrees when s bands are at least sqrt 2 lines, -90 when a line. `c:tickLblSkip` sets s.
        let categoryPlan (style: TextStyle) (size: float) (texts: string []) (band: float) (auto: bool) (skip: int option) : int * LabelPlan =
            let measure = widthIn style size false
            let widestOf (items: string []) = items |> Array.fold (fun most text -> max most (measure text)) 0.0
            let widestText = widestOf texts
            let widestWord = widestOf (texts |> Array.collect (fun text -> text.Split(' ')))
            let spaced = texts |> Array.exists (fun text -> text.Trim().Contains " ")
            let line = lineFor style size
            let breakAt (room: float) (text: string) : string list =
                let lines = ResizeArray<string>()
                let mutable current = ""
                for word in text.Split(' ') do
                    let joined = if current = "" then word else current + " " + word
                    if current = "" || measure joined <= room then current <- joined
                    else
                        lines.Add current
                        current <- word
                lines.Add current
                List.ofSeq lines
            let planAt (every: int) : LabelPlan option =
                let room = float every * band
                // (Turned, a label has the plot's width shared among the labels shown - every s-th of n, ceil (n / s) of them:
                // Excel 2019 shows every 12th of 60 "Category n" in 60.56 pt, not every 11th, whose 6 labels would have 10.09
                // pt each, less than a line - synthetic charts of 24 and 60 categories, 90 - 760 pt, at 9 pt.)
                let shared = band * float texts.Length / Math.Ceiling (float texts.Length / float every)
                if widestText <= room then Some Flat
                elif spaced && widestWord <= room then Some (Wrapped (texts |> Array.map (breakAt room)))
                elif auto && shared >= sqrt 2.0 * line then Some (Turned -45.0)
                elif auto && shared >= line then Some (Turned -90.0)
                else None
            match skip with
            | Some every -> every, (planAt every |> Option.defaultValue Flat)
            | None when band <= 0.0 || texts.Length = 0 -> 1, Flat
            | None ->
                let rec find (every: int) =
                    if every >= texts.Length then every, (planAt every |> Option.defaultValue Flat)
                    else
                        match planAt every with
                        | Some plan -> every, plan
                        | None -> find (every + 1)
                find 1
        /// The height the category labels take below the plot area, for every s-th of them laid out as planned.
        let planRoom (style: TextStyle) (size: float) (texts: string []) (every: int, plan: LabelPlan) : float =
            let shownTexts = texts |> Array.indexed |> Array.filter (fun (index, _) -> index % every = 0) |> Array.map snd
            let longest = shownTexts |> Array.fold (fun most text -> max most (widthIn style size false text)) 0.0
            let line = lineFor style size
            match plan with
            | Flat -> line + categoryGap
            | Wrapped lines ->
                let most = lines |> Array.indexed |> Array.filter (fun (index, _) -> index % every = 0) |> Array.fold (fun most (_, broken) -> max most broken.Length) 1
                float most * line + categoryGap
            | Turned angle ->
                let turn = abs angle * Math.PI / 180.0
                (if abs angle >= 90.0 then longest else (longest + line) * Math.Sin turn) + categoryGap
        // Axis titles (and what says that an axis is in thousands, millions ...).
        let axisTitle (axis: Axis option) =
            axis |> Option.bind (fun axis ->
                if axis.Deleted then None
                else
                    let style = over axis.TitleStyle axis.LabelStyle
                    match axis.Title, axis.DisplayUnitLabel with
                    | Some text, Some units -> Some (text + " (" + units + ")", textSize style 10.0, textColor style, style, false)
                    | Some text, None -> Some (text, textSize style 10.0, textColor style, style, false)
                    | None, Some units -> Some (units, textSize style 10.0, textColor style, style, true)
                    | None, None -> None)
        /// An axis title of one line of runs of their own looks, drawn run by run (else None: one text).
        let titleRunsOf (axis: Axis option) =
            match axis with
            | Some axis when axis.DisplayUnitLabel.IsNone && axis.TitleRuns.Length = 1 && mixedRuns axis.TitleRuns ->
                Some (axis.TitleRuns.Head |> List.map (fun (text, style) -> text, over style axis.LabelStyle))
            | _ -> None
        let titleHeight (size: float) (axis: Axis option) =
            match titleRunsOf axis with
            | Some runs -> runs |> List.fold (fun most (_, style) -> max most (textSize style 10.0)) size
            | None -> size
        let sideAxis, baseAxis = (if horizontal then categoryAxis, primaryValueAxis elif isScatter then primaryValueAxis, xAxis else primaryValueAxis, categoryAxis)
        // (A plot area the file places leaves the titles of its axes where they are by themselves. What says the
        // labels are in thousands, millions ... stands at the far end of its axis - the top of one up the side,
        // the right of one along the bottom: dashboard-sales-analysis' "Thousands".)
        (match axisTitle sideAxis with
         | Some (text, size, color, style, unitsOnly) ->
            let y =
                if unitsOnly then (match plotRect with Some (_, y, _, _) -> y | None -> top) + widthIn style size false text / 2.0
                else (match plotRect with Some (_, y, _, h) -> y + h / 2.0 | None -> (top + bottom) / 2.0)
            let high = titleHeight size sideAxis
            (match titleRunsOf sideAxis with
             | Some runs -> drawRuns runs 10.0 false (left + high * 0.6) y "middle" -90.0
             | None -> addText style (PText (left + size * 0.6, y, text, size, color, "middle", false, false, -90.0)))
            left <- left + high * 1.5
         | None -> ())
        (match axisTitle baseAxis with
         | Some (text, size, color, style, unitsOnly) ->
            let x =
                if unitsOnly then (match plotRect with Some (x, _, w, _) -> x + w | None -> right) - widthIn style size false text / 2.0
                else (match plotRect with Some (x, _, w, _) -> x + w / 2.0 | None -> (left + right) / 2.0)
            let high = titleHeight size baseAxis
            (match titleRunsOf baseAxis with
             | Some runs -> drawRuns runs 10.0 false x (bottom - high * 0.6) "middle" 0.0
             | None -> addText style (PText (x, bottom - size * 0.6, text, size, color, "middle", false, false, 0.0)))
            bottom <- bottom - high * 1.5
         | None -> ())
        (match (if secondaryAxis.IsSome then axisTitle secondaryAxis else None) with
         | Some (text, size, color, style, _) ->
            addText style (PText (right - size * 0.6, (match plotRect with Some (_, y, _, h) -> y + h / 2.0 | None -> (top + bottom) / 2.0), text, size, color, "middle", false, false, -90.0))
            right <- right - size * 1.5
         | None -> ())
        // The labels along the bottom: a line of them, or - turned - as long as the longest.
        let categoryLabels = categories |> Array.toList
        let turned = match categoryAxis with Some axis when not isScatter && not horizontal -> axis.LabelRotation |> Option.filter (fun rot -> abs rot > 45.0) | _ -> None
        let baseSize = if horizontal then valueSize elif isScatter then textSize xStyle 10.0 else categorySize
        let baseStyle = if horizontal then valueStyle elif isScatter then xStyle else categoryStyle
        let baseShown = if horizontal then shown primaryValueAxis elif isScatter then shown xAxis else shown categoryAxis
        let baseRoom =
            if not baseShown then 4.0
            else
                match turned with
                | Some _ -> widest categoryStyle categoryLabels categorySize + px 4.0
                | None -> lineFor baseStyle baseSize + categoryGap
        // The plot area's box with its axes' labels: where the file puts it, or what the title, the legend and the
        // axes' titles leave. Excel cuts a value axis into as many intervals as its labels allow in it (`axisIntervals`).
        let boxWidth, boxHeight =
            match chart.PlotLayout, plotRect with
            | Some { Inner = true }, Some (_, _, w, h) -> w + px 20.0, h + baseRoom
            | _, Some (_, _, w, h) -> w, h
            | _ -> right - left, bottom - top
        let primaryScale = scaleFor primaryValueAxis false (axisIntervals horizontal (if horizontal then boxWidth else boxHeight) valueSize (faceOf valueStyle))
        let secondaryScale = scaleFor secondaryAxis true (axisIntervals false boxHeight valueSize (faceOf valueStyle))
        let xScale : AxisScale =
            if not isScatter then { Min = 0.0; Max = 1.0; Step = 1.0; Log = None }
            else
                // (The x values of the points drawn: those with a y value.)
                let values =
                    series |> List.collect (fun item ->
                        item.XValues |> Array.mapi (fun index x -> if index < item.Values.Length && item.Values.[index].IsSome then x else None)
                        |> Array.choose id |> Array.toList)
                let most = axisIntervals true boxWidth (textSize xStyle 10.0) (faceOf xStyle)
                match values, xAxis with
                | [], _ -> { Min = 0.0; Max = float (max 1 count); Step = 1.0; Log = None }
                | _, Some ({ LogBase = Some logBase } as logAxis) when values |> List.exists (fun value -> value > 0.0) ->
                    let positive = values |> List.filter (fun value -> value > 0.0)
                    logScale logBase (List.min positive) (List.max positive) logAxis
                | _ ->
                    let minimum, maximum, step = valueScale (List.min values) (List.max values) (xAxis |> Option.bind (fun a -> a.Min)) (xAxis |> Option.bind (fun a -> a.Max)) (xAxis |> Option.bind (fun a -> a.MajorUnit)) most
                    { Min = minimum; Max = maximum; Step = step; Log = None }
        let valueLabels = ticksOf primaryScale |> List.map (axisFormat primaryValueAxis false false)
        let secondaryLabels = if secondaryAxis.IsSome then ticksOf secondaryScale |> List.map (axisFormat secondaryAxis true false) else []
        // The room the tick labels take.
        let xLabels = if isScatter then ticksOf xScale |> List.map (axisFormat xAxis false true) else []
        let sideLabels, baseLabels = (if horizontal then categoryLabels, valueLabels elif isScatter then valueLabels, xLabels else valueLabels, categoryLabels)
        let sideSize, sideStyle = (if horizontal then categorySize, categoryStyle else valueSize, valueStyle)
        let sideShown = if horizontal then shown categoryAxis else shown primaryValueAxis
        // (A multi-level category axis - a pivot chart of months in years, of managers in states: a row (a column,
        // turned, for bars that run to the right) of labels for each outer level, each centred on its group, beyond
        // the points' own, lines between the groups - dashboard-sales-analysis' "Sales By Chain" / "By Manager".)
        let outerLevels =
            if isScatter || dated.IsSome || not (shown categoryAxis) then []
            else series |> List.tryPick (fun item -> match item.Source.Categories with Some source when not source.Outer.IsEmpty -> Some source.Outer | _ -> None) |> Option.defaultValue []
        let outerRow = lineFor categoryStyle categorySize + px 3.0
        let sideRoom = if sideShown then min (width * 0.4) (widest sideStyle sideLabels sideSize) + valueGap + (if horizontal then float outerLevels.Length * outerRow else 0.0) else 4.0
        let secondaryRoom = if secondaryAxis.IsSome && shown secondaryAxis then widest valueStyle secondaryLabels valueSize + valueGap else 0.0
        // (The last label of a value axis along the bottom is centred on the plot's edge: half of it is beyond.)
        let overhang =
            if baseShown && (horizontal || isScatter) && secondaryRoom = 0.0 then
                (match List.tryLast baseLabels with Some text -> widthIn (if isScatter then xStyle else valueStyle) baseSize false text / 2.0 | None -> 0.0)
            else 0.0
        // (Points on the ticks - a value axis crossing `midCat`, no bars: the first and the last at the plot's ends.)
        let onTicks = not horizontal && count > 1 && chart.Plots |> List.forall (fun plot -> plot.Kind <> BarPlot) && valueAxes |> List.exists (fun axis -> axis.OnTicks)
        // (The category labels along the bottom are laid out for the plot area's width: they take the height
        // that leaves. On the ticks a label is as far from the next as the points are.)
        let planFor (plotLeft: float) (plotRight: float) =
            match categoryAxis with
            | Some axis when not horizontal && not isScatter && shown categoryAxis && dated.IsNone && turned.IsNone ->
                Some (categoryPlan categoryStyle categorySize categories ((plotRight - plotLeft) / float (if onTicks then count - 1 else count)) axis.LabelRotation.IsNone axis.LabelSkip)
            | _ -> None
        // (On the ticks the first and the last category labels are centred on the plot's ends, and half of each stays in
        // the plot area's box: Excel 2019's inside plot of synthetic-chart-looks "AreaTicks" / "LineTicks" is 31.22 pt from
        // the left, half of "January 2024" past the box's 7, and half of "June 2024" short of its right, 18.54 pt.)
        let tickHalves =
            if onTicks && not isScatter && shown categoryAxis && dated.IsNone && turned.IsNone && categories.Length > 0 then
                widthIn categoryStyle categorySize false categories.[0] / 2.0, widthIn categoryStyle categorySize false categories.[categories.Length - 1] / 2.0
            else 0.0, 0.0
        let plotLeft, plotTop, plotRight, plotBottom, labelPlan =
            match chart.PlotLayout, plotRect with
            | Some { Inner = true }, Some (x, y, w, h) -> x, y, x + w, y + h, planFor x (x + w)
            | _ ->
                // (A rectangle for the plot area with its labels: the labels are inside it.)
                let left, top, right, bottom = (match plotRect with Some (x, y, w, h) -> max 2.0 x, max 2.0 y, min (width - 2.0) (x + w), min (height - 2.0) (y + h) | None -> left, top, right, bottom)
                let plotLeft = max (left + sideRoom) (left + fst tickHalves)
                let plotRight = max (plotLeft + 10.0) (right - max (secondaryRoom + (if baseShown && not horizontal && not isScatter then 0.0 else overhang)) (snd tickHalves))
                let plan = planFor plotLeft plotRight
                let room = (match plan with Some planned -> planRoom categoryStyle categorySize categories planned | None -> baseRoom) + (if horizontal then 0.0 else float outerLevels.Length * outerRow)
                let plotBottom = max (top + 10.0) (bottom - room)
                plotLeft, top, plotRight, plotBottom, plan
        let plotWidth, plotHeight = max 1.0 (plotRight - plotLeft), max 1.0 (plotBottom - plotTop)
        (match chart.PlotBackground, strokeColor chart.PlotBorder with
         | FillColor (color, opacity), border -> add (PRect (plotLeft, plotTop, plotWidth, plotHeight, Some color, opacity, border |> Option.map (fun (c, _, _) -> c), (match border with Some (_, w, _) -> w | None -> 0.0), 0.0))
         | _, Some (color, strokeWidth, _) -> add (PRect (plotLeft, plotTop, plotWidth, plotHeight, None, 1.0, Some color, strokeWidth, 0.0))
         | _ -> ())
        // Where a value is on its value axis (along the plot's height, or its width for bars that run to the right).
        let reversedOf (axis: Axis option) = axis |> Option.map (fun axis -> axis.Reversed) |> Option.defaultValue false
        let along (scale: AxisScale) (reversed: bool) (value: float) : float =
            let share = shareOn scale value
            let share = if reversed then 1.0 - share else share
            if horizontal then plotLeft + share * plotWidth else plotBottom - share * plotHeight
        let primaryReversed, secondaryReversed = reversedOf primaryValueAxis, reversedOf secondaryAxis
        let valueAt (secondary: bool) (value: float) = if secondary then along secondaryScale secondaryReversed value else along primaryScale primaryReversed value
        /// Where a bar or an area begins: at zero, or at the lower end of a logarithmic scale.
        let floorOf (secondary: bool) : float =
            let scale = if secondary then secondaryScale else primaryScale
            match scale.Log with
            | Some logBase -> Math.Pow (logBase, scale.Min)
            | None -> max scale.Min (min scale.Max 0.0)
        let xReversed = reversedOf xAxis
        let xAt (value: float) : float =
            let share = shareOn xScale value
            plotLeft + (if xReversed then 1.0 - share else share) * plotWidth
        // Where a category's band begins, and its width (bars that run to the right list their categories upwards).
        let categoryReversed = reversedOf categoryAxis
        let band = (if horizontal then plotHeight else plotWidth) / float count
        let bandStart (index: int) : float =
            if horizontal then (if categoryReversed then plotTop + float index * band else plotBottom - float (index + 1) * band)
            else (if categoryReversed then plotRight - float (index + 1) * band else plotLeft + float index * band)
        // ---- grid lines ----
        /// Where the points of a category are.
        let pointAt (index: int) : float = if onTicks then bandStart 0 + plotWidth * float index / float (count - 1) else bandStart index + band / 2.0
        let gridOf (axis: Axis option) = axis |> Option.bind (fun axis -> strokeColor axis.Gridlines)
        (match gridOf primaryValueAxis with
         | Some (color, strokeWidth, dash) ->
            for value in ticksOf primaryScale do
                let at = along primaryScale primaryReversed value
                if horizontal then add (PLine (at, plotTop, at, plotBottom, color, strokeWidth, dash))
                else add (PLine (plotLeft, at, plotRight, at, color, strokeWidth, dash))
         | None -> ())
        (match (if isScatter then gridOf xAxis else None) with
         | Some (color, strokeWidth, dash) ->
            for value in ticksOf xScale do
                let at = xAt value
                add (PLine (at, plotTop, at, plotBottom, color, strokeWidth, dash))
         | None -> ())
        // ---- the series, plot by plot (areas and bars under lines) ----
        let pendingLabels = ResizeArray<Primitive>()
        let labelAt (item: SeriesData) (index: int) (value: float) (x: float) (y: float) (anchor: string) (inside: bool) =
            match labelled item index with
            | Some spec ->
                let size = textSize spec.Style 9.0
                pendingLabels.Add (PText (x, y, labelText item spec index value None, size, (spec.Style.Color |> Option.defaultValue (if inside then "#FFFFFF" else "#404040")), anchor, (spec.Style.Bold |> Option.defaultValue false), false, 0.0))
            | None -> ()
        let order (plot: Plot) = match plot.Kind with AreaPlot -> 0 | BarPlot -> 1 | StockPlot -> 2 | LinePlot -> 3 | BubblePlot -> 4 | ScatterPlot -> 5 | _ -> 6
        for plot in chart.Plots |> List.sortBy order do
            let members = series |> List.filter (fun item -> obj.ReferenceEquals(item.Plot, plot))
            let stacked = stackedOf plot
            let percent = plot.Grouping = "percentStacked"
            let wholes =
                if percent then
                    let positive = Array.zeroCreate<float> count
                    for item in members do
                        item.Values |> Array.iteri (fun index value -> match value with Some v when index < count -> positive.[index] <- positive.[index] + abs v | _ -> ())
                    positive
                else [||]
            let scaled (index: int) (value: float) = if percent then (if wholes.[index] > 0.0 then value / wholes.[index] else 0.0) else value
            match plot.Kind with
            | BarPlot ->
                let n = max 1 members.Length
                let overlap = if stacked then 1.0 else max -1.0 (min 1.0 (plot.Overlap / 100.0))
                // (The gap between clusters is a share of one bar, as Excel 2019 draws it: two series with a gap of
                // 182 % are bars of a band / 3.82 - dashboard-sales-analysis' "Sales By Manager", 7.5 px of 29.6.)
                let cluster = if stacked then 1.0 else float n - float (n - 1) * overlap
                let barWidth = band / (cluster + max 0.0 plot.GapWidth / 100.0)
                let groupWidth = barWidth * cluster
                let positive = Array.zeroCreate<float> count
                let negative = Array.zeroCreate<float> count
                members |> List.iteri (fun memberIndex item ->
                    for index in 0 .. count - 1 do
                        match (if index < item.Values.Length then item.Values.[index] else None) with
                        | Some raw ->
                            let value = scaled index raw
                            let from, upto =
                                if stacked then
                                    if value >= 0.0 then (let start = positive.[index] in positive.[index] <- start + value; start, start + value)
                                    else (let start = negative.[index] in negative.[index] <- start + value; start, start + value)
                                else floorOf item.Secondary, value
                            let a, b = valueAt item.Secondary from, valueAt item.Secondary upto
                            let offset = (band - groupWidth) / 2.0 + (if stacked then 0.0 else float memberIndex * barWidth * (1.0 - overlap))
                            // (Bars that run to the right list their series upwards too.)
                            let offset = if horizontal && not stacked && not categoryReversed then (band - groupWidth) / 2.0 + float (n - 1 - memberIndex) * barWidth * (1.0 - overlap) else offset
                            let start = bandStart index + offset
                            let color, opacity = pointColor item index
                            let fill = if item.Look.Filled || (Map.containsKey index item.Source.PointFills) then Some color else None
                            if horizontal then add (PRect (min a b, start, abs (b - a), barWidth, fill, opacity, item.Look.LineColor, item.Look.LineWidth, 0.0))
                            else add (PRect (start, min a b, barWidth, abs (b - a), fill, opacity, item.Look.LineColor, item.Look.LineWidth, 0.0))
                            let position = labelStyleOf item |> Option.bind (fun spec -> spec.Position) |> Option.defaultValue (if stacked then "ctr" else "outEnd")
                            let inside = position = "ctr" || position = "inEnd" || position = "inBase"
                            if horizontal then
                                let x = (match position with "ctr" -> (a + b) / 2.0 | "inEnd" -> b - (if b >= a then 4.0 else -4.0) | "inBase" -> a + (if b >= a then 4.0 else -4.0) | _ -> b + (if b >= a then 4.0 else -4.0))
                                let anchor = (match position with "ctr" -> "middle" | "inEnd" -> (if b >= a then "end" else "start") | _ -> (if b >= a then "start" else "end"))
                                labelAt item index raw x (start + barWidth / 2.0) anchor inside
                            else
                                let y = (match position with "ctr" -> (a + b) / 2.0 | "inEnd" -> b + (if b <= a then 8.0 else -8.0) | "inBase" -> a - (if b <= a then 8.0 else -8.0) | _ -> b - (if b <= a then 8.0 else -8.0))
                                labelAt item index raw (start + barWidth / 2.0) y "middle" inside
                        | None -> ())
            | LinePlot | AreaPlot | StockPlot ->
                let sums = Array.zeroCreate<float> count
                // (What every series of the plot has at a category: for the lines between the highest and the lowest.)
                let reached = Array.init count (fun _ -> ResizeArray<float>())
                for item in members do
                    let points = ResizeArray<float * float * int * float>()
                    let floors = ResizeArray<float * float>()
                    let runs = ResizeArray<(float * float) list>()
                    let mutable run : (float * float) list = []
                    for index in 0 .. count - 1 do
                        match (if index < item.Values.Length then item.Values.[index] else None) with
                        | Some raw ->
                            let value = scaled index raw
                            let basis = if stacked then sums.[index] else floorOf item.Secondary
                            let total = if stacked then basis + value else value
                            if stacked then sums.[index] <- total
                            let x = pointAt index
                            let y = valueAt item.Secondary total
                            points.Add ((x, y, index, raw))
                            reached.[index].Add y
                            floors.Add ((x, valueAt item.Secondary basis))
                            run <- (x, y) :: run
                        | None ->
                            if not run.IsEmpty then runs.Add (List.rev run)
                            run <- []
                    if not run.IsEmpty then runs.Add (List.rev run)
                    if plot.Kind = AreaPlot && points.Count > 0 then
                        let upper = points |> Seq.map (fun (x, y, _, _) -> x, y) |> Seq.toList
                        let lower = floors |> Seq.toList |> List.rev
                        add (PPolyline (upper @ lower, (if item.Look.Filled then Some item.Look.Color else None), item.Look.Opacity, item.Look.LineColor, item.Look.LineWidth, None))
                    elif plot.Kind = LinePlot || plot.Kind = StockPlot then
                        // (The series of a stock chart are drawn as lines only where the file gives them one.)
                        (match item.Look.LineColor with
                         | Some color when plot.Kind = LinePlot || (match item.Source.Line with Some { Paint = FillColor _ } -> true | _ -> false) ->
                            for line in runs do
                                if line.Length > 2 && item.Source.Smooth then add (PPath (smoothPath line, None, 1.0, Some color, item.Look.LineWidth))
                                elif line.Length > 1 then add (PPolyline (line, None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash))
                         | _ -> ())
                        // Markers: the chart's switch, unless the series says none.
                        match markerOf item (plot.Kind = LinePlot && plot.Markers) with
                        | Some (symbol, size) ->
                            // (A marker of c:size n is n pt across: 7 px for 5, chart-zoo's scatter in Excel 2019.)
                            let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                            let stroke = (match strokeColor item.Source.MarkerLine with Some (color, _, _) -> Some color | None -> Some item.Look.Color)
                            for (x, y, _, _) in points do drawMarker symbol x y size fill stroke
                        | None -> ()
                    for (x, y, index, raw) in points do
                        let position = labelStyleOf item |> Option.bind (fun spec -> spec.Position) |> Option.defaultValue (if plot.Kind = AreaPlot then "ctr" else "r")
                        (match position with
                         | "t" -> labelAt item index raw x (y - 9.0) "middle" false
                         | "b" -> labelAt item index raw x (y + 9.0) "middle" false
                         | "l" -> labelAt item index raw (x - 6.0) y "end" false
                         | "ctr" -> labelAt item index raw x y "middle" false
                         | _ -> labelAt item index raw (x + 6.0) y "start" false)
                // The lines from the highest to the lowest value, the bars from the first series to the last.
                let hiLow = if plot.Kind = StockPlot then (plot.HiLowLines |> Option.orElse (Some { Paint = FillColor ("#595959", 1.0); Width = 1.0; Dash = None })) else plot.HiLowLines
                (match strokeColor hiLow with
                 | Some (color, strokeWidth, dash) ->
                    for index in 0 .. count - 1 do
                        if reached.[index].Count > 1 then
                            let x = pointAt index
                            add (PLine (x, Seq.min reached.[index], x, Seq.max reached.[index], color, strokeWidth, dash))
                 | None -> ())
                (match plot.UpDownBars, members with
                 | Some (rise, fall), first :: _ when members.Length > 1 ->
                    let last = List.last members
                    let barWidth = band / 2.5
                    for index in 0 .. count - 1 do
                        match (if index < first.Values.Length then first.Values.[index] else None), (if index < last.Values.Length then last.Values.[index] else None) with
                        | Some opening, Some closing ->
                            let a, b = valueAt first.Secondary opening, valueAt last.Secondary closing
                            let fill = (match (if closing >= opening then rise else fall) with FillColor (color, _) -> Some color | _ -> None)
                            add (PRect (bandStart index + (band - barWidth) / 2.0, min a b, barWidth, max 1.0 (abs (b - a)), fill, 1.0, Some "#404040", 0.75, 0.0))
                        | _ -> ()
                 | None, _ :: _ when plot.Kind = StockPlot ->
                    // (Without bars the last series - the close - is a tick to the right of the line.)
                    let last = List.last members
                    if last.Source.MarkerSymbol.IsNone || last.Source.MarkerSymbol = Some "none" then
                        for index in 0 .. count - 1 do
                            match (if index < last.Values.Length then last.Values.[index] else None) with
                            | Some closing ->
                                let x, y = pointAt index, valueAt last.Secondary closing
                                add (PLine (x, y, x + min 6.0 (band / 3.0), y, "#595959", 1.5, None))
                            | None -> ()
                 | _ -> ())
            | ScatterPlot | BubblePlot ->
                // (The largest bubble is a quarter of the shorter side of the plot, at 100 %.)
                let largest = members |> List.fold (fun most item -> item.Sizes |> Array.fold (fun most size -> match size with Some value -> max most value | None -> most) most) 0.0
                let bubbleRadius (size: float) =
                    if largest <= 0.0 || size <= 0.0 then 0.0
                    else min plotWidth plotHeight * 0.125 * (max 10.0 (min 300.0 plot.BubbleScale) / 100.0) * Math.Sqrt (size / largest)
                for item in members do
                    let points =
                        [ for index in 0 .. item.Values.Length - 1 do
                            let x = (if index < item.XValues.Length then item.XValues.[index] else Some (float (index + 1)))
                            match x, item.Values.[index] with
                            | Some x, Some y when xScale.Max > xScale.Min -> yield xAt x, valueAt item.Secondary y, index, y
                            | _ -> () ]
                    if plot.Kind = BubblePlot then
                        for (x, y, index, _) in points do
                            let radius = (match (if index < item.Sizes.Length then item.Sizes.[index] else None) with Some size -> bubbleRadius size | None -> 4.0)
                            let color, opacity = pointColor item index
                            if radius > 0.0 then add (PEllipse (x, y, radius, radius, (if item.Look.Filled then Some color else None), opacity, item.Look.LineColor, item.Look.LineWidth))
                    else
                        let lined = plot.ScatterStyle.StartsWith "line" || plot.ScatterStyle.StartsWith "smooth"
                        (match item.Look.LineColor with
                         | Some color when lined && points.Length > 1 && (item.Source.Line.IsSome || plot.ScatterStyle <> "lineMarker" || item.Source.MarkerSymbol = Some "none") ->
                            add (PPolyline (points |> List.map (fun (x, y, _, _) -> x, y), None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash))
                         | _ -> ())
                        match markerOf item (plot.ScatterStyle.EndsWith "Marker" || plot.ScatterStyle = "marker") with
                        | Some (symbol, size) ->
                            let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                            let stroke = (match strokeColor item.Source.MarkerLine with Some (color, _, _) -> Some color | None -> Some item.Look.Color)
                            for (x, y, _, _) in points do drawMarker symbol x y (max 4.0 size) fill stroke
                        | None -> ()
                    for (x, y, index, raw) in points do labelAt item index raw (x + 6.0) y "start" false
            | BoxesPlot (figures, options) ->
                // A box a series a category: from the first quartile to the third, the median across it, whiskers to
                // the farthest points within 1.5 times the box's height of it, capped; the mean's mark (x), the points
                // inside the whiskers and outside them - their outlines and the box's the series' colour darker (80 %:
                // Excel 2019's #365B9D round #4472C4). The boxes of a category stand side by side in the band less its
                // gap, a tenth of a box apart.
                let n = max 1 members.Length
                let group = band / (1.0 + max 0.0 plot.GapWidth / 100.0)
                let boxWidth = group / (float n + float (n - 1) * 0.1)
                let darker (color: string) = let r, g, b = rgbOf (solidOf color) in "#" + hex2 (r * 0.8) + hex2 (g * 0.8) + hex2 (b * 0.8)
                members |> List.iteri (fun memberIndex item ->
                    let color = item.Look.Color
                    let edge = darker color
                    let means = ResizeArray<float * float>()
                    let boxes = if memberIndex < figures.Length then figures.[memberIndex] else [||]
                    for index in 0 .. min (count - 1) (boxes.Length - 1) do
                        match boxes.[index] with
                        | Some box ->
                            let left = bandStart index + (band - group) / 2.0 + float memberIndex * boxWidth * 1.1
                            let centre = left + boxWidth / 2.0
                            let y (value: float) = valueAt item.Secondary value
                            let cap = boxWidth / 4.0
                            add (PLine (centre, y box.Q3, centre, y box.High, edge, 1.0, None))
                            add (PLine (centre - cap, y box.High, centre + cap, y box.High, edge, 1.0, None))
                            add (PLine (centre, y box.Q1, centre, y box.Low, edge, 1.0, None))
                            add (PLine (centre - cap, y box.Low, centre + cap, y box.Low, edge, 1.0, None))
                            add (PRect (left, y box.Q3, boxWidth, max 0.0 (y box.Q1 - y box.Q3), (if item.Look.Filled then Some color else None), item.Look.Opacity, Some edge, 1.0, 0.0))
                            add (PLine (left, y box.Median, left + boxWidth, y box.Median, edge, 1.0, None))
                            let dot (value: float) = add (PEllipse (centre, y value, 2.5, 2.5, Some color, 1.0, Some edge, 1.0))
                            if options.InnerPoints then for value in box.Inner do dot value
                            if options.Outliers then for value in box.Outside do dot value
                            if options.MeanMarker then drawMarker "x" centre (y box.Mean) 7.0 None (Some edge)
                            means.Add (centre, y box.Mean)
                        | None -> ()
                    if options.MeanLine && means.Count > 1 then add (PPolyline (List.ofSeq means, None, 1.0, Some edge, 1.0, None)))
            | _ -> ()
        // ---- axes: lines and labels ----
        let axisLine (axis: Axis option) = axis |> Option.bind (fun axis -> if axis.Deleted then None else strokeColor axis.Line)
        let zeroAt = along primaryScale primaryReversed (floorOf false)
        (match axisLine categoryAxis with
         | Some (color, strokeWidth, _) when not isScatter ->
            if horizontal then add (PLine (zeroAt, plotTop, zeroAt, plotBottom, color, strokeWidth, None))
            else add (PLine (plotLeft, zeroAt, plotRight, zeroAt, color, strokeWidth, None))
         | _ -> ())
        (match axisLine primaryValueAxis with
         | Some (color, strokeWidth, _) ->
            if horizontal then add (PLine (plotLeft, plotBottom, plotRight, plotBottom, color, strokeWidth, None))
            else add (PLine (plotLeft, plotTop, plotLeft, plotBottom, color, strokeWidth, None))
         | None -> ())
        (match (if isScatter then axisLine xAxis else None) with
         | Some (color, strokeWidth, _) -> add (PLine (plotLeft, plotBottom, plotRight, plotBottom, color, strokeWidth, None))
         | None -> ())
        // (Every n-th label where they would run into each other.)
        let thinned (style: TextStyle) (labels: string list) (spacing: float) (alongAxis: bool) (size: float) : int =
            let need = if alongAxis then widest style labels size + 8.0 else size * 1.15
            if spacing <= 0.0 then 1 else max 1 (int (Math.Ceiling (need / spacing)))
        let below = plotBottom + categoryGap + lineFor baseStyle baseSize / 2.0
        if shown primaryValueAxis then
            let every = thinned valueStyle valueLabels ((if horizontal then plotWidth else plotHeight) / float (max 1 (valueLabels.Length - 1))) horizontal valueSize
            List.zip (ticksOf primaryScale) valueLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then
                    let at = along primaryScale primaryReversed value
                    if horizontal then addText valueStyle (PText (at, below, text, valueSize, valueColor, "middle", false, false, 0.0))
                    else addText valueStyle (PText (plotLeft - valueGap, at, text, valueSize, valueColor, "end", false, false, 0.0)))
        if secondaryAxis.IsSome && shown secondaryAxis then
            let style = styleOf secondaryAxis
            let every = thinned style secondaryLabels (plotHeight / float (max 1 (secondaryLabels.Length - 1))) false valueSize
            List.zip (ticksOf secondaryScale) secondaryLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then addText style (PText (plotRight + valueGap, along secondaryScale secondaryReversed value, text, valueSize, valueColor, "start", false, false, 0.0)))
        if isScatter then
            if shown xAxis then
                let size = textSize xStyle 10.0
                let every = thinned xStyle xLabels (plotWidth / float (max 1 (xLabels.Length - 1))) true size
                List.zip (ticksOf xScale) xLabels |> List.iteri (fun index (value, text) ->
                    if index % every = 0 then addText xStyle (PText (xAt value, below, text, size, textColor xStyle, "middle", false, false, 0.0)))
        elif shown categoryAxis then
            // (As many labels as fit: every n-th one when they are crowded.)
            if horizontal then
                let every = max 1 (int (Math.Ceiling ((categorySize * 1.15) / band)))
                // (The labels have the room left of the plot to the chart's edge, whatever placed the plot: dashboard-sales-analysis'
                // "Womens", "Accessories" whole in Excel.)
                let outerWidth = float outerLevels.Length * outerRow
                let room = max 10.0 (plotLeft - valueGap - 2.0 - outerWidth)
                if not outerLevels.IsEmpty then
                    let lineColor, lineWidth = (match axisLine categoryAxis with Some (color, strokeWidth, _) -> color, strokeWidth | None -> "#D9D9D9", 1.0)
                    let labelsLeft = plotLeft - valueGap - min room (widest categoryStyle categoryLabels categorySize)
                    outerLevels |> List.iteri (fun level groups ->
                        let starts = groups |> Map.toList |> List.filter (fun (index, _) -> index < count)
                        let x = labelsLeft - float (level + 1) * outerRow
                        starts |> List.iteri (fun at (first, text) ->
                            let last = (if at + 1 < starts.Length then fst starts.[at + 1] else count) - 1
                            let y0, y1 = bandStart first + band, bandStart last
                            add (PLine (x, y0, plotLeft, y0, lineColor, lineWidth, None))
                            if at + 1 = starts.Length then add (PLine (x, y1, plotLeft, y1, lineColor, lineWidth, None))
                            if text <> "" then addText categoryStyle (PText (x + outerRow / 2.0, (y0 + y1) / 2.0, text, categorySize, categoryColor, "middle", false, false, -90.0))))
                categories |> Array.iteri (fun index text ->
                    if index % every = 0 then
                        let long = widthIn categoryStyle categorySize false text
                        let text = if long > room && text.Length > 4 then text.Substring(0, max 1 (int (float text.Length * room / long) - 1)) + "…" else text
                        addText categoryStyle (PText (plotLeft - valueGap, bandStart index + band / 2.0, text, categorySize, categoryColor, "end", false, false, 0.0)))
            else
                match turned with
                | Some rotation ->
                    // (Turned labels: their ends at the axis.)
                    let pitch = lineFor categoryStyle categorySize
                    let every = if dated.IsSome then 1 else max 1 (int (Math.Ceiling (pitch / band)))
                    categories |> Array.iteri (fun index text ->
                        if index % every = 0 && text <> "" then
                            addText categoryStyle (PText (pointAt index, plotBottom + px 3.0, text, categorySize, categoryColor, (if rotation < 0.0 then "end" else "start"), false, false, rotation)))
                | None ->
                    let line = lineFor categoryStyle categorySize
                    match labelPlan with
                    | Some (every, Wrapped lines) ->
                        lines |> Array.iteri (fun index broken ->
                            if index % every = 0 then
                                broken |> List.iteri (fun at text ->
                                    if text <> "" then addText categoryStyle (PText (pointAt index, plotBottom + categoryGap + line / 2.0 + float at * line, text, categorySize, categoryColor, "middle", false, false, 0.0))))
                    | Some (every, Turned angle) ->
                        // (Turned labels: their ends at the axis, below the middle of their band.)
                        categories |> Array.iteri (fun index text ->
                            if index % every = 0 && text <> "" then
                                addText categoryStyle (PText (pointAt index, plotBottom + px 3.0, text, categorySize, categoryColor, "end", false, false, angle)))
                    | Some (every, Flat) ->
                        categories |> Array.iteri (fun index text ->
                            if index % every = 0 && text <> "" then addText categoryStyle (PText (pointAt index, below, text, categorySize, categoryColor, "middle", false, false, 0.0)))
                    | None ->
                        // (A date axis: a label for every major unit, as placed.)
                        categories |> Array.iteri (fun index text ->
                            if text <> "" then addText categoryStyle (PText (pointAt index, below, text, categorySize, categoryColor, "middle", false, false, 0.0)))
                    // The outer levels, each group's label centred on its points, lines between the groups.
                    if not outerLevels.IsEmpty then
                        let innerRoom = (match labelPlan with Some planned -> planRoom categoryStyle categorySize categories planned | None -> baseRoom)
                        let lineColor, lineWidth = (match axisLine categoryAxis with Some (color, strokeWidth, _) -> color, strokeWidth | None -> "#D9D9D9", 1.0)
                        let depth = innerRoom + float outerLevels.Length * outerRow
                        outerLevels |> List.iteri (fun level groups ->
                            let starts = groups |> Map.toList |> List.filter (fun (index, _) -> index < count)
                            let top = plotBottom + innerRoom + float level * outerRow
                            starts |> List.iteri (fun at (first, text) ->
                                let last = (if at + 1 < starts.Length then fst starts.[at + 1] else count) - 1
                                let x0, x1 = bandStart first, bandStart last + band
                                add (PLine (x0, plotBottom, x0, plotBottom + depth, lineColor, lineWidth, None))
                                if at + 1 = starts.Length then add (PLine (x1, plotBottom, x1, plotBottom + depth, lineColor, lineWidth, None))
                                if text <> "" then addText categoryStyle (PText ((x0 + x1) / 2.0, top + outerRow / 2.0, text, categorySize, categoryColor, "middle", false, false, 0.0))))
        for label in pendingLabels do add label
    { Width = width; Height = height; Items = List.ofSeq items }

/// The scene of a chart in a rectangle of `width` x `height` pixels.
/// A number to `digits` significant digits.
let private roundSignificant (digits: int) (value: float) : float =
    if value = 0.0 || Double.IsNaN value || Double.IsInfinity value then value
    else
        let scale = Math.Pow (10.0, float (digits - 1 - int (Math.Floor (Math.Log10 (abs value)))))
        Math.Round (value * scale, MidpointRounding.AwayFromZero) / scale

/// The bins of a chartEx histogram as Excel 2019 makes them, each its label and how many values it holds (by
/// category: the sum of each category's values, in the order the categories come). Values are cut from the least
/// (or the underflow) in bins of the set width, of the set count, or - set neither - of Scott's width, 3.5 times
/// the values' standard deviation (of a sample) over the cube root of their count, to two significant digits;
/// as many as reach the greatest (or the overflow). Checked: build-chartex-zoo.vbs, 8 data sets of 5 - 200 values
/// (12.654 is 13 wide: [11.87, 24.87] .. (76.87, 89.87]; 183.1 is 180; 0.304 is 0.3).
let histogramBins (binning: Binning) (values: float list) (categories: string list option) : (string * float) list =
    match categories with
    | Some names ->
        let order = ResizeArray<string>()
        let sums = Collections.Generic.Dictionary<string, float>()
        names |> List.iteri (fun index name ->
            let value = if index < values.Length then values.[index] else 0.0
            if not (sums.ContainsKey name) then (order.Add name; sums.[name] <- 0.0)
            sums.[name] <- sums.[name] + value)
        order |> Seq.map (fun name -> name, sums.[name]) |> List.ofSeq
    | None when values.IsEmpty -> []
    | None ->
        let tidy (value: float) = generalNumber (roundSignificant 12 value)
        let lowest, highest = List.min values, List.max values
        let start = defaultArg binning.Underflow lowest
        let stop = defaultArg binning.Overflow highest
        let n = float values.Length
        let width =
            match binning.Size, binning.Count with
            | Some size, _ -> size
            | None, Some count -> (stop - start) / float count
            | None, None ->
                let mean = List.average values
                let deviation = if values.Length > 1 then sqrt ((values |> List.sumBy (fun value -> (value - mean) ** 2.0)) / (n - 1.0)) else 0.0
                roundSignificant 2 (3.5 * deviation / Math.Pow (n, 1.0 / 3.0))
        let width = if width > 0.0 && not (Double.IsNaN width) then width else max 1.0 (stop - start)
        let count =
            match binning.Count with
            | Some count when binning.Size.IsNone -> count
            | _ -> max 1 (int (Math.Ceiling ((stop - start) / width - 1e-9)))
        // (The last bin ends at the overflow: Excel 2019's "(700, 800]" of 150 wide from 100.)
        let edge (index: int) = roundSignificant 12 (if binning.Overflow.IsSome then min stop (start + width * float index) else start + width * float index)
        let counts = Array.zeroCreate<float> count
        let mutable under, over = 0.0, 0.0
        for value in values do
            match binning.Underflow, binning.Overflow with
            | Some limit, _ when (if binning.ClosedLeft then value < limit else value <= limit) -> under <- under + 1.0
            | _, Some limit when (if binning.ClosedLeft then value >= limit else value > limit) -> over <- over + 1.0
            | _ ->
                let at = (value - start) / width
                let index = if binning.ClosedLeft then int (Math.Floor (at + 1e-9)) else int (Math.Ceiling (at - 1e-9)) - 1
                let index = max 0 (min (count - 1) index)
                counts.[index] <- counts.[index] + 1.0
        let label (index: int) =
            let a, b = tidy (edge index), tidy (edge (index + 1))
            if binning.ClosedLeft then (if index = count - 1 && binning.Overflow.IsNone then sprintf "[%s, %s]" a b else sprintf "[%s, %s)" a b)
            else (if index = 0 && binning.Underflow.IsNone then sprintf "[%s, %s]" a b else sprintf "(%s, %s]" a b)
        [ if binning.Underflow.IsSome then yield ((if binning.ClosedLeft then "< " else "≤ ") + tidy start), under
          for index in 0 .. count - 1 do yield label index, counts.[index]
          if binning.Overflow.IsSome then yield ((if binning.ClosedLeft then "≥ " else "> ") + tidy stop), over ]

/// A quartile of sorted values: QUARTILE.INC's (inclusive) or QUARTILE.EXC's place among them, between two
/// values in proportion - as Excel 2019's boxes (1 .. 9, 30: 2.75 .. 8.25 exclusive, 3.25 .. 7.75 inclusive).
let private quartileOf (inclusive: bool) (sorted: float []) (share: float) : float =
    let n = sorted.Length
    let place = if inclusive then share * float (n - 1) + 1.0 else share * float (n + 1)
    let place = max 1.0 (min (float n) place)
    let below = int (Math.Floor place)
    if below >= n then sorted.[n - 1] else sorted.[below - 1] + (place - float below) * (sorted.[below] - sorted.[below - 1])

/// A box's figures: the quartiles, the whiskers to the farthest values within 1.5 times the box's height of it, the
/// values inside and outside them, the mean.
let boxFiguresOf (inclusive: bool) (values: float list) : BoxFigures option =
    if values.IsEmpty then None
    else
        let sorted = values |> List.sort |> List.toArray
        let q1, median, q3 = quartileOf inclusive sorted 0.25, quartileOf true sorted 0.5, quartileOf inclusive sorted 0.75
        let reach = 1.5 * (q3 - q1)
        let inside, outside = sorted |> Array.toList |> List.partition (fun value -> value >= q1 - reach && value <= q3 + reach)
        let low, high = (match inside with [] -> q1, q3 | _ -> min q1 (List.min inside), max q3 (List.max inside))
        Some { Low = low; Q1 = q1; Median = median; Q3 = q3; High = high; Mean = List.average values; Inner = inside; Outside = outside }

/// A chartEx histogram, Pareto or box and whisker chart as a chart of the classic kind, its bins or boxes worked
/// out from its data, in the look of Excel 2019's chartEx charts: text #595959, the title 14 pt, axes' labels 9 pt,
/// axes' lines and gridlines #D9D9D9.
let private classicOfChartEx (theme: ThemeColors) (data: ChartData) (chart: Chart) (plot: Plot) : Chart =
    let literal (texts: string list) (numbers: float option list) : DataRef =
        { Formula = None
          Strings = texts |> List.mapi (fun index text -> index, text) |> Map.ofList
          Numbers = numbers |> List.mapi (fun index value -> index, value) |> List.choose (fun (index, value) -> value |> Option.map (fun v -> index, v)) |> Map.ofList
          Count = max texts.Length numbers.Length; FormatCode = None; FollowedEdit = false; LimitToCachedPoints = false; Outer = [] }
    let grey = { Paint = FillColor ("#D9D9D9", 1.0); Width = 1.0; Dash = None }
    let text (style: TextStyle) (size: float) = over style { noTextStyle with Size = Some size }
    let axes =
        chart.Axes |> List.map (fun axis ->
            { axis with
                LabelStyle = text axis.LabelStyle 9.0
                Line = axis.Line |> Option.orElse (Some grey)
                Gridlines = axis.Gridlines |> Option.map (fun lines -> match lines.Paint with FillColor _ -> lines | _ -> grey) })
    let numbers (series: Series) = numbersOf data series.Values |> Array.toList |> List.choose id
    let named (series: Series) = textsOf data series.Categories |> Array.toList |> function [] -> None | names -> Some names
    let plots =
        match plot.Kind with
        | HistogramPlot (binning, pareto) ->
            let series = plot.Series |> List.head
            // (Categories that are text: a bar each, by its values' sum; numbers: binned.)
            let isNumber (name: string) = fst (Double.TryParse(name, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture))
            let categories = named series |> Option.filter (List.exists (fun name -> name.Trim() <> "" && not (isNumber name)))
            let values = match categories with Some _ -> numbersOf data series.Values |> Array.toList |> List.map (Option.defaultValue 0.0) | None -> numbers series
            let bins = histogramBins binning values categories
            // (A Pareto chart's bars from the most to the least, its line their running share.)
            let bins = if pareto then bins |> List.sortBy (fun (_, count) -> -count) else bins
            let total = bins |> List.sumBy snd
            let bars =
                { series with
                    Categories = Some (literal (List.map fst bins) [])
                    Values = Some (literal [] (bins |> List.map (snd >> Some)))
                    // (The bars touch, a pixel of white between them.)
                    Line = series.Line |> Option.orElse (Some { Paint = FillColor ("#FFFFFF", 1.0); Width = 1.0; Dash = None }) }
            let barPlot = { plot with Kind = BarPlot; Grouping = "clustered"; Series = [ bars ]; AxisIds = [ "0"; "1" ] }
            let running = bins |> List.scan (fun sum (_, count) -> sum + count) 0.0 |> List.tail |> List.map (fun sum -> Some (if total > 0.0 then sum / total else 0.0))
            let line =
                { bars with
                    Order = 1; Index = 1; Name = None
                    Values = Some (literal [] running)
                    Fill = FillAuto
                    Line = Some { Paint = FillColor (autoColor theme 1, 1.0); Width = linePixels 28575.0; Dash = None }
                    MarkerSymbol = Some "none"; Smooth = false; Labels = None; PointFills = Map.empty; PointLines = Map.empty }
            if pareto then [ barPlot; { plot with Kind = LinePlot; Grouping = "standard"; Series = [ line ]; AxisIds = [ "0"; "2" ] } ] else [ barPlot ]
        | BoxWhiskerPlot options ->
            // (A box a category a series: the categories in the order they come, else one, "1".)
            let groups =
                plot.Series |> List.map (fun series ->
                    let values = numbersOf data series.Values
                    match named series with
                    | Some names -> names |> List.mapi (fun index name -> name, (if index < values.Length then values.[index] else None))
                    | None -> values |> Array.toList |> List.map (fun value -> "1", value))
            let categories = groups |> List.collect (List.map fst) |> List.filter (fun name -> name <> "") |> List.distinct
            let categories = if categories.IsEmpty then [ "1" ] else categories
            let figures =
                groups |> List.map (fun pairs ->
                    categories |> List.map (fun name -> pairs |> List.filter (fun (known, _) -> known = name) |> List.choose snd |> boxFiguresOf options.Inclusive) |> List.toArray)
                |> List.toArray
            let series =
                List.zip plot.Series (Array.toList figures) |> List.map (fun (series, boxes) ->
                    { series with Categories = Some (literal categories []); Values = Some (literal [] (boxes |> Array.toList |> List.map (Option.map (fun box -> box.Median)))) })
            [ { plot with Kind = BoxesPlot (figures, options); Grouping = "clustered"; Series = series; AxisIds = [ "0"; "1" ] } ]
        | _ -> [ plot ]
    { chart with
        Plots = plots
        Axes = axes
        TitleStyle = over (text chart.TitleStyle 14.0) { noTextStyle with Bold = Some false }
        TextStyle = over chart.TextStyle { noTextStyle with Color = Some "#595959" }
        RoundedCorners = false
        Style = 2 }

/// Whether a plot is one of a chartEx part's (a waterfall, a funnel, a histogram ...): such a part keeps no copy of
/// its data.
let isChartExPlot (plot: Plot) : bool =
    match plot.Kind with
    | WaterfallPlot _ | FunnelPlot | HistogramPlot _ | BoxWhiskerPlot _ | TreemapPlot _ | SunburstPlot | UndrawnPlot _ -> true
    | _ -> false

let layoutChart (theme: ThemeColors) (data: ChartData) (chart: Chart) (width: float) (height: float) : Scene =
    match chart.Plots |> List.tryFind isChartExPlot with
    | Some ({ Kind = HistogramPlot _ | BoxWhiskerPlot _ } as plot) -> layoutClassicChart theme data (classicOfChartEx theme data chart plot) width height
    | Some plot -> layoutChartEx theme data chart plot width height
    | None -> layoutClassicChart theme data chart width height

// ---------------------------------------------------------------------------
// Shapes, groups, and what an anchor of a drawing holds
// ---------------------------------------------------------------------------

/// What an anchor's objects need of the package they come from.
type DrawingParts = {
    Theme: ThemeColors
    /// The chart of a relationship id of the drawing (`<c:chart r:id>`), read.
    Chart: string -> Chart option
    /// The picture of a relationship id (`<a:blip r:embed>`), as a data URL.
    Picture: string -> string option
    /// A slicer or a timeline the drawing shows by name (`<sle:slicer name>`: kind "slicer";
    /// `<tsle:timeslicer name>`: "timeslicer") drawn in a rectangle of that width and height, reading
    /// the cells it needs through the data (`Slicers.drawer`); None when the workbook has none of that name.
    Slicer: string -> string -> ChartData -> float -> float -> Primitive list option
    /// A linked picture of a range (Excel's camera, `<a14:cameraTool cellRange>`) inside a group: the range's cells as
    /// they are now, drawn in a rectangle of that width and height, as a data URL; None when it cannot be read.
    Camera: string -> float -> float -> string option
}

let private emuPx (emu: float) = emu / 9525.0

/// The rectangle `a:xfrm` gives (EMU), with what a group maps it to.
let private frameOf (shape: XmlDsl.XmlNode) : (float * float * float * float) option =
    let transform = child "spPr" shape |> Option.bind (child "xfrm") |> Option.orElse (child "grpSpPr" shape |> Option.bind (child "xfrm")) |> Option.orElse (child "xfrm" shape)
    transform
    |> Option.bind (fun xfrm ->
        match child "off" xfrm, child "ext" xfrm with
        | Some off, Some ext ->
            match attr "x" off |> Option.bind number, attr "y" off |> Option.bind number, attr "cx" ext |> Option.bind number, attr "cy" ext |> Option.bind number with
            | Some x, Some y, Some w, Some h -> Some (x, y, w, h)
            | _ -> None
        | _ -> None)

// ---------------------------------------------------------------------------
// Custom geometry (`a:custGeom`) and the preset shapes that are outlines of points
// ---------------------------------------------------------------------------

/// The guides of a geometry (`a:avLst`, `a:gdLst`: `<a:gd name fmla>`) worked out for a shape of
/// `width` x `height` (in the units of the shape: EMU), with the names every shape has (w, h, ss, ...).
let private guidesOf (lists: XmlDsl.XmlNode list) (width: float) (height: float) : Collections.Generic.Dictionary<string, float> =
    let values = Collections.Generic.Dictionary<string, float>()
    let short, long = min width height, max width height
    for name, value in
        [ "w", width; "h", height; "l", 0.0; "t", 0.0; "r", width; "b", height
          "hc", width / 2.0; "vc", height / 2.0; "ss", short; "ls", long
          "wd2", width / 2.0; "wd3", width / 3.0; "wd4", width / 4.0; "wd5", width / 5.0; "wd6", width / 6.0; "wd8", width / 8.0; "wd10", width / 10.0; "wd12", width / 12.0; "wd32", width / 32.0
          "hd2", height / 2.0; "hd3", height / 3.0; "hd4", height / 4.0; "hd5", height / 5.0; "hd6", height / 6.0; "hd8", height / 8.0; "hd10", height / 10.0; "hd12", height / 12.0; "hd32", height / 32.0
          "ssd2", short / 2.0; "ssd4", short / 4.0; "ssd6", short / 6.0; "ssd8", short / 8.0; "ssd16", short / 16.0; "ssd32", short / 32.0
          "cd2", 10800000.0; "cd4", 5400000.0; "cd8", 2700000.0; "3cd4", 16200000.0; "3cd8", 8100000.0; "5cd8", 13500000.0; "7cd8", 18900000.0 ] do
        values.[name] <- value
    let valueOf (token: string) : float =
        match Portable.Dict.tryFind token values with
        | ValueSome value -> value
        | ValueNone -> number token |> Option.defaultValue 0.0
    // (Angles of the formulas are in 60000ths of a degree.)
    let radians (angle: float) = angle / 60000.0 * Math.PI / 180.0
    for list in lists do
        for guide in childrenNamed "gd" list do
            match attr "name" guide, attr "fmla" guide with
            | Some name, Some formula ->
                let parts = formula.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
                let arg (index: int) = if index < parts.Length then valueOf parts.[index] else 0.0
                let x, y, z = arg 1, arg 2, arg 3
                let value =
                    if parts.Length = 0 then 0.0
                    else
                        match parts.[0] with
                        | "val" -> x
                        | "*/" -> if z = 0.0 then 0.0 else x * y / z
                        | "+-" -> x + y - z
                        | "+/" -> if z = 0.0 then 0.0 else (x + y) / z
                        | "?:" -> if x > 0.0 then y else z
                        | "abs" -> abs x
                        | "at2" -> Math.Atan2 (y, x) * 180.0 / Math.PI * 60000.0
                        | "cat2" -> x * Math.Cos (Math.Atan2 (z, y))
                        | "sat2" -> x * Math.Sin (Math.Atan2 (z, y))
                        | "cos" -> x * Math.Cos (radians y)
                        | "sin" -> x * Math.Sin (radians y)
                        | "tan" -> x * Math.Tan (radians y)
                        | "max" -> max x y
                        | "min" -> min x y
                        | "mod" -> Math.Sqrt (x * x + y * y + z * z)
                        | "pin" -> if y < x then x elif y > z then z else y
                        | "sqrt" -> Math.Sqrt (max 0.0 x)
                        | _ -> 0.0
                values.[name] <- (if Double.IsNaN value || Double.IsInfinity value then 0.0 else value)
            | _ -> ()
    values

/// The paths of a custom geometry for a rectangle of the scene: each the `d` of a `PPath`, and
/// whether it is filled and whether it is outlined.
let private customPaths (geometry: XmlDsl.XmlNode) (x: float) (y: float) (width: float) (height: float) : (string * bool * bool) list =
    // (The guides are worked out in EMU, as the file's numbers are.)
    let emuWidth, emuHeight = width * 9525.0, height * 9525.0
    let guides = guidesOf ([ child "avLst" geometry; child "gdLst" geometry ] |> List.choose id) emuWidth emuHeight
    let valueOf (text: string option) : float =
        match text with
        | Some token -> (match guides.TryGetValue token with | true, value -> value | false, _ -> number token |> Option.defaultValue 0.0)
        | None -> 0.0
    match child "pathLst" geometry with
    | None -> []
    | Some list ->
        childrenNamed "path" list
        |> List.choose (fun pathNode ->
            let pathWidth = attr "w" pathNode |> Option.bind number |> Option.filter (fun value -> value > 0.0) |> Option.defaultValue emuWidth
            let pathHeight = attr "h" pathNode |> Option.bind number |> Option.filter (fun value -> value > 0.0) |> Option.defaultValue emuHeight
            let scaleX, scaleY = (if pathWidth > 0.0 then width / pathWidth else 0.0), (if pathHeight > 0.0 then height / pathHeight else 0.0)
            let d = Text.StringBuilder()
            let mutable curX, curY = 0.0, 0.0
            let point (node: XmlDsl.XmlNode) = valueOf (attr "x" node), valueOf (attr "y" node)
            let write (px: float, py: float) = d.Append(fmt (x + px * scaleX)).Append(' ').Append(fmt (y + py * scaleY)) |> ignore
            for command in kids pathNode do
                let points = childrenNamed "pt" command |> List.map point
                match nameOf command, points with
                | "moveTo", [ p ] -> d.Append('M') |> ignore; write p; curX <- fst p; curY <- snd p
                | "lnTo", [ p ] -> d.Append('L') |> ignore; write p; curX <- fst p; curY <- snd p
                | "cubicBezTo", [ a; b; c ] ->
                    d.Append('C') |> ignore; write a; d.Append(' ') |> ignore; write b; d.Append(' ') |> ignore; write c
                    curX <- fst c; curY <- snd c
                | "quadBezTo", [ a; b ] ->
                    d.Append('Q') |> ignore; write a; d.Append(' ') |> ignore; write b
                    curX <- fst b; curY <- snd b
                | "arcTo", _ ->
                    // An arc of an ellipse of radii wR, hR from the current point, which is at the
                    // angle stAng on it, over swAng (clockwise when positive).
                    let radiusX, radiusY = valueOf (attr "wR" command), valueOf (attr "hR" command)
                    let start = valueOf (attr "stAng" command) / 60000.0 * Math.PI / 180.0
                    let sweep = valueOf (attr "swAng" command) / 60000.0 * Math.PI / 180.0
                    let centreX, centreY = curX - radiusX * Math.Cos start, curY - radiusY * Math.Sin start
                    // (A whole turn is two halves: an arc that ends where it begins draws nothing.)
                    let pieces = if abs sweep > Math.PI * 1.5 then 2 else 1
                    for piece in 1 .. pieces do
                        let angle = start + sweep * float piece / float pieces
                        let endX, endY = centreX + radiusX * Math.Cos angle, centreY + radiusY * Math.Sin angle
                        d.Append('A').Append(fmt (radiusX * scaleX)).Append(' ').Append(fmt (radiusY * scaleY)).Append(" 0 ")
                         .Append(if abs sweep / float pieces > Math.PI then "1 " else "0 ").Append(if sweep > 0.0 then "1 " else "0 ") |> ignore
                        write (endX, endY)
                        curX <- endX; curY <- endY
                | "close", _ -> d.Append('Z') |> ignore
                | _ -> ()
            if d.Length = 0 then None
            else Some (d.ToString(), attr "fill" pathNode <> Some "none", isOn (attr "stroke" pathNode) true))

/// The outline of a preset shape that is a polygon, as points of the unit square - with the
/// shape's first adjustment (0 .. 1 of the shorter side, where the preset has one). None for a
/// preset that is not one of these.
let private presetPolygon (geometry: string) (adjust: float option) (width: float) (height: float) : (float * float) list option =
    let short = min width height
    // (An adjustment is a share of the shorter side, along the width or the height.)
    let alongX (share: float) = if width > 0.0 then min 1.0 (share * short / width) else 0.0
    let alongY (share: float) = if height > 0.0 then min 1.0 (share * short / height) else 0.0
    let adj (otherwise: float) = adjust |> Option.defaultValue otherwise
    match geometry with
    | "rtTriangle" -> Some [ 0.0, 0.0; 1.0, 1.0; 0.0, 1.0 ]
    | "parallelogram" | "flowChartInputOutput" -> (let a = alongX (adj 0.25) in Some [ a, 0.0; 1.0, 0.0; 1.0 - a, 1.0; 0.0, 1.0 ])
    | "trapezoid" -> (let a = alongX (adj 0.25) in Some [ a, 0.0; 1.0 - a, 0.0; 1.0, 1.0; 0.0, 1.0 ])
    | "pentagon" -> Some [ 0.5, 0.0; 1.0, 0.38; 0.81, 1.0; 0.19, 1.0; 0.0, 0.38 ]
    | "hexagon" -> (let a = alongX (adj 0.25) in Some [ a, 0.0; 1.0 - a, 0.0; 1.0, 0.5; 1.0 - a, 1.0; a, 1.0; 0.0, 0.5 ])
    | "octagon" -> (let a, b = alongX (adj 0.29289), alongY (adj 0.29289) in Some [ a, 0.0; 1.0 - a, 0.0; 1.0, b; 1.0, 1.0 - b; 1.0 - a, 1.0; a, 1.0; 0.0, 1.0 - b; 0.0, b ])
    | "homePlate" -> (let a = alongX (adj 0.5) in Some [ 0.0, 0.0; 1.0 - a, 0.0; 1.0, 0.5; 1.0 - a, 1.0; 0.0, 1.0 ])
    | "chevron" -> (let a = alongX (adj 0.5) in Some [ 0.0, 0.0; 1.0 - a, 0.0; 1.0, 0.5; 1.0 - a, 1.0; 0.0, 1.0; a, 0.5 ])
    | "rightArrow" -> (let a = alongX (adj 0.5) in Some [ 0.0, 0.25; 1.0 - a, 0.25; 1.0 - a, 0.0; 1.0, 0.5; 1.0 - a, 1.0; 1.0 - a, 0.75; 0.0, 0.75 ])
    | "leftArrow" -> (let a = alongX (adj 0.5) in Some [ 1.0, 0.25; a, 0.25; a, 0.0; 0.0, 0.5; a, 1.0; a, 0.75; 1.0, 0.75 ])
    | "upArrow" -> (let a = alongY (adj 0.5) in Some [ 0.25, 1.0; 0.25, a; 0.0, a; 0.5, 0.0; 1.0, a; 0.75, a; 0.75, 1.0 ])
    | "downArrow" -> (let a = alongY (adj 0.5) in Some [ 0.25, 0.0; 0.25, 1.0 - a; 0.0, 1.0 - a; 0.5, 1.0; 1.0, 1.0 - a; 0.75, 1.0 - a; 0.75, 0.0 ])
    | "leftRightArrow" -> (let a = alongX (adj 0.5) in Some [ 0.0, 0.5; a, 0.0; a, 0.25; 1.0 - a, 0.25; 1.0 - a, 0.0; 1.0, 0.5; 1.0 - a, 1.0; 1.0 - a, 0.75; a, 0.75; a, 1.0 ])
    | "plus" -> (let a, b = alongX (adj 0.25), alongY (adj 0.25) in Some [ a, 0.0; 1.0 - a, 0.0; 1.0 - a, b; 1.0, b; 1.0, 1.0 - b; 1.0 - a, 1.0 - b; 1.0 - a, 1.0; a, 1.0; a, 1.0 - b; 0.0, 1.0 - b; 0.0, b; a, b ])
    | "star5" ->
        Some [ for index in 0 .. 9 do
                let angle = -Math.PI / 2.0 + float index * Math.PI / 5.0
                let radius = if index % 2 = 0 then 0.5 else 0.19
                yield 0.5 + radius * Math.Cos angle * 1.05, 0.55 + radius * Math.Sin angle * 1.1 ]
    | "snip1Rect" -> (let a, b = alongX (adj 0.16667), alongY (adj 0.16667) in Some [ 0.0, 0.0; 1.0 - a, 0.0; 1.0, b; 1.0, 1.0; 0.0, 1.0 ])
    | "snip2DiagRect" -> (let a, b = alongX 0.16667, alongY 0.16667 in Some [ 0.0, 0.0; 1.0 - a, 0.0; 1.0, b; 1.0, 1.0; a, 1.0; 0.0, 1.0 - b ])
    | "flowChartManualInput" -> Some [ 0.0, 0.2; 1.0, 0.0; 1.0, 1.0; 0.0, 1.0 ]
    | "flowChartPreparation" -> Some [ 0.2, 0.0; 0.8, 0.0; 1.0, 0.5; 0.8, 1.0; 0.2, 1.0; 0.0, 0.5 ]
    | "flowChartMerge" -> Some [ 0.0, 0.0; 1.0, 0.0; 0.5, 1.0 ]
    | "flowChartExtract" -> Some [ 0.5, 0.0; 1.0, 1.0; 0.0, 1.0 ]
    | _ -> None

let private shapeItems (parts: DrawingParts) (data: ChartData) (shape: XmlDsl.XmlNode) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    let props = child "spPr" shape
    let style = child "style" shape
    let styleColor (name: string) = style |> Option.bind (child name) |> Option.bind (fun reference -> kids reference |> List.tryPick (colorOfNode parts.Theme))
    let isConnector = nameOf shape = "cxnSp"
    let geometry = props |> Option.bind (child "prstGeom") |> Option.bind (attr "prst") |> Option.defaultValue (if isConnector then "line" else "rect")
    let fill =
        match props |> Option.map (shapeFillOf parts.Theme) |> Option.defaultValue FillAuto with
        | FillColor (color, opacity) -> Some (color, opacity)
        | FillNone -> None
        | FillAuto ->
            if isConnector then None
            else
                match style |> Option.bind (child "fillRef") |> Option.map (styleFill parts.Theme) with
                | Some (FillColor (color, opacity)) -> Some (color, opacity)
                | Some FillNone -> None
                | _ -> styleColor "fillRef"
    let styleStroke () = style |> Option.bind (child "lnRef") |> Option.bind (styleLine parts.Theme)
    let stroke =
        match props |> Option.bind (strokeOf parts.Theme) with
        | Some { Paint = FillColor (color, _); Width = strokeWidth } -> Some (color, strokeWidth)
        | Some { Paint = FillNone } -> None
        | Some { Width = strokeWidth } -> styleStroke () |> Option.map (fun (color, _) -> color, strokeWidth)
        | None -> styleStroke ()
    let fillColor, opacity = (match fill with Some (color, opacity) -> Some color, opacity | None -> None, 1.0)
    let strokeColor, strokeWidth = (match stroke with Some (color, strokeWidth) -> Some color, strokeWidth | None -> None, 0.0)
    let flipH = props |> Option.bind (child "xfrm") |> Option.bind (attr "flipH") |> Option.exists (fun value -> value = "1" || value = "true")
    let flipV = props |> Option.bind (child "xfrm") |> Option.bind (attr "flipV") |> Option.exists (fun value -> value = "1" || value = "true")
    let outline =
        match geometry with
        | "line" | "straightConnector1" | "bentConnector2" | "bentConnector3" | "curvedConnector3" ->
            let x1, x2 = (if flipH then x + width, x else x, x + width)
            let y1, y2 = (if flipV then y + height, y else y, y + height)
            [ PLine (x1, y1, x2, y2, (strokeColor |> Option.defaultValue "#000000"), max 1.0 strokeWidth, None) ]
        | "ellipse" -> [ PEllipse (x + width / 2.0, y + height / 2.0, width / 2.0, height / 2.0, fillColor, opacity, strokeColor, strokeWidth) ]
        | "roundRect" | "round1Rect" | "round2SameRect" | "round2DiagRect" | "flowChartAlternateProcess" | "flowChartTerminator" ->
            // (The corner: a sixth of the shorter side unless the shape's adjustment says otherwise.)
            let adjust =
                props |> Option.bind (path [ "prstGeom"; "avLst"; "gd" ]) |> Option.bind (attr "fmla")
                |> Option.bind (fun formula -> if formula.StartsWith "val " then number (formula.Substring 4) else None)
                |> Option.defaultValue 16667.0
            let radius = (if geometry = "flowChartTerminator" then min width height / 2.0 else min width height * (min 50000.0 adjust) / 100000.0)
            [ PRect (x, y, width, height, fillColor, opacity, strokeColor, strokeWidth, radius) ]
        | "triangle" -> [ PPolyline ([ x + width / 2.0, y; x + width, y + height; x, y + height ], (fillColor |> Option.orElse (Some "none")), opacity, strokeColor, strokeWidth, None) ]
        | "leftBrace" | "rightBrace" ->
            // A curly brace, opening right ({) or left (}): its curls a twelfth of the shorter side
            // (adj1 8333), its point half way down (adj2 50000) unless the shape says otherwise. A line,
            // not an area: Excel draws its outline only.
            let guide (name: string) (otherwise: float) =
                props |> Option.bind (path [ "prstGeom"; "avLst" ])
                |> Option.bind (fun list -> kids list |> List.tryFind (fun gd -> attr "name" gd = Some name))
                |> Option.bind (attr "fmla")
                |> Option.bind (fun formula -> if formula.StartsWith "val " then number (formula.Substring 4) else None)
                |> Option.defaultValue otherwise
            let curl = min (min width height * guide "adj1" 8333.0 / 100000.0) (height / 4.0)
            let point = height * guide "adj2" 50000.0 / 100000.0
            let n (value: float) = string (System.Math.Round (value, 2))
            let left = geometry = "leftBrace"
            let outer, inner, mid = (if left then x + width, x, x + width / 2.0 else x, x + width, x + width / 2.0)
            let top, centre, bottom = y, y + point, y + height
            let d =
                "M " + n outer + " " + n top + " Q " + n mid + " " + n top + " " + n mid + " " + n (top + curl)
                + " L " + n mid + " " + n (centre - curl) + " Q " + n mid + " " + n centre + " " + n inner + " " + n centre
                + " Q " + n mid + " " + n centre + " " + n mid + " " + n (centre + curl)
                + " L " + n mid + " " + n (bottom - curl) + " Q " + n mid + " " + n bottom + " " + n outer + " " + n bottom
            [ PPath (d, None, opacity, (strokeColor |> Option.orElse fillColor), max 0.75 strokeWidth) ]
        | "diamond" | "flowChartDecision" -> [ PPolyline ([ x + width / 2.0, y; x + width, y + height / 2.0; x + width / 2.0, y + height; x, y + height / 2.0 ], (fillColor |> Option.orElse (Some "none")), opacity, strokeColor, strokeWidth, None) ]
        | _ ->
            let adjust = props |> Option.bind (path [ "prstGeom"; "avLst"; "gd" ]) |> Option.bind (attr "fmla") |> Option.bind (fun formula -> if formula.StartsWith "val " then number (formula.Substring 4) else None)
            match props |> Option.bind (child "custGeom"), presetPolygon geometry (adjust |> Option.map (fun value -> value / 100000.0)) width height with
            | Some custom, _ when not (customPaths custom x y width height).IsEmpty ->
                customPaths custom x y width height
                |> List.map (fun (d, filled, outlined) -> PPath (d, (if filled then fillColor else None), opacity, (if outlined then strokeColor else None), strokeWidth))
            | _, Some points ->
                [ PPolyline (points @ [ List.head points ] |> List.map (fun (px, py) -> x + px * width, y + py * height), (fillColor |> Option.orElse (Some "none")), opacity, strokeColor, strokeWidth, None) ]
            | _ -> if fillColor.IsNone && strokeColor.IsNone then [] else [ PRect (x, y, width, height, fillColor, opacity, strokeColor, strokeWidth, 0.0) ]
    let centreX, centreY = x + width / 2.0, y + height / 2.0
    // (A line was drawn from the flipped corners already. The text of a flipped shape stays readable.)
    let outline =
        if isConnector || geometry = "line" then outline
        else
            let outline = if flipH then [ PTransform (-1.0, 0.0, 0.0, 1.0, 2.0 * centreX, 0.0, outline) ] else outline
            if flipV then [ PTransform (1.0, 0.0, 0.0, -1.0, 0.0, 2.0 * centreY, outline) ] else outline
    // A linked shape (`textlink`) shows its cell as the sheet has it now; its field holds the text of the last save.
    let linkedText =
        attr "textlink" shape
        |> Option.filter (fun link -> link <> "")
        |> Option.bind parseAreas
        |> Option.bind List.tryHead
        |> Option.bind (fun (sheet, row, col, _, _) -> data.Cells sheet row col row col)
        |> Option.bind (fun block -> block |> Array.tryHead |> Option.bind Array.tryHead |> Option.flatten)
        |> Option.map data.Display
    // The text: paragraphs of runs, each run with its own size, colour and weight; a paragraph is one line.
    let text =
        match child "txBody" shape with
        | None -> []
        | Some body ->
            let bodyProps = child "bodyPr" body
            let inset (name: string) (otherwise: float) = bodyProps |> Option.bind (attr name) |> Option.bind number |> Option.map emuPx |> Option.defaultValue otherwise
            let padLeft, padRight, padTop, padBottom = inset "lIns" 9.6, inset "rIns" 9.6, inset "tIns" 4.8, inset "bIns" 4.8
            let fontColor = styleColor "fontRef" |> Option.map fst
            let lines =
                childrenNamed "p" body
                |> List.map (fun paragraph ->
                    let paragraphProps = child "pPr" paragraph
                    let align = paragraphProps |> Option.bind (attr "algn") |> Option.defaultValue "l"
                    let defaults = paragraphProps |> Option.bind (child "defRPr") |> Option.map (runStyle parts.Theme) |> Option.defaultValue noTextStyle
                    let runs =
                        kids paragraph
                        |> List.choose (fun run ->
                            match nameOf run with
                            | "r" | "fld" ->
                                let style = over (child "rPr" run |> Option.map (runStyle parts.Theme) |> Option.defaultValue noTextStyle) defaults
                                let cached = child "t" run |> Option.map innerText |> Option.defaultValue ""
                                let linked = (attr "type" run |> Option.map (fun kind -> kind.ToLowerInvariant ())) = Some "txlink"
                                Some (Some ((if linked then linkedText |> Option.defaultValue cached else cached), style))
                            | "br" -> Some None
                            | _ -> None)
                    let endStyle = child "endParaRPr" paragraph |> Option.map (runStyle parts.Theme) |> Option.defaultValue noTextStyle
                    // (A line break makes a line of its own of what follows.)
                    let broken = ResizeArray<(string * TextStyle) list>([ [] ])
                    for run in runs do
                        match run with
                        | Some piece -> broken.[broken.Count - 1] <- broken.[broken.Count - 1] @ [ piece ]
                        | None -> broken.Add []
                    broken |> Seq.map (fun line -> align, line, over endStyle defaults) |> Seq.toList)
                |> List.concat
            let sizeOf (style: TextStyle) = px (style.Size |> Option.defaultValue 11.0)
            let wraps = (bodyProps |> Option.bind (attr "wrap")) <> Some "none"
            // (A run that names no typeface is in the theme's minor font - the e-commerce dashboard's "Most Sold Items" in
            // Tw Cen MT, two lines in Excel 2019; lines are broken by the width the text has in it.)
            let faceIn (style: TextStyle) = style.Face |> Option.orElse (Map.tryFind "#minorFont" parts.Theme)
            // (Measured by the face's design advances, as shapes' text is laid out: `ExcelText.idealWidth`.)
            let measured (style: TextStyle) (text: string) =
                let bold = style.Bold |> Option.defaultValue false
                match ExcelText.idealWidth (defaultArg (faceIn style) "Calibri") bold (style.Italic |> Option.defaultValue false) (sizeOf style) text with
                | Some ideal -> ideal
                | None -> faceWidth (faceIn style) (sizeOf style) bold text
            let room = max 10.0 (width - padLeft - padRight)
            // Each paragraph in lines that fit the shape: broken between words.
            let wrapped : (string * (string * TextStyle) list * TextStyle) list =
                lines
                |> List.collect (fun (align, runs, fallback) ->
                    let whole = runs |> List.sumBy (fun (text, style) -> measured style text)
                    if not wraps || whole <= room || runs.IsEmpty then [ align, runs, fallback ]
                    else
                        // (Words keep the space after them.)
                        let words =
                            runs
                            |> List.collect (fun (text, style) ->
                                let pieces = ResizeArray<string>()
                                let mutable start = 0
                                for at in 0 .. text.Length - 1 do
                                    if text.[at] = ' ' then
                                        pieces.Add (text.Substring(start, at - start + 1))
                                        start <- at + 1
                                if start < text.Length then pieces.Add (text.Substring start)
                                pieces |> Seq.map (fun piece -> piece, style) |> Seq.toList)
                        let made = ResizeArray<(string * TextStyle) list>()
                        let mutable current : (string * TextStyle) list = []
                        let mutable used = 0.0
                        for word, style in words do
                            let wide = measured style (word.TrimEnd())
                            if used + wide > room && not current.IsEmpty then
                                made.Add (List.rev current)
                                current <- []
                                used <- 0.0
                            current <- (word, style) :: current
                            used <- used + measured style word
                        if not current.IsEmpty then made.Add (List.rev current)
                        made |> Seq.map (fun line -> align, line, fallback) |> Seq.toList)
            let lineHeight (_, runs: (string * TextStyle) list, fallback: TextStyle) =
                (match runs with [] -> sizeOf fallback | _ -> runs |> List.fold (fun most (_, style) -> max most (sizeOf style)) 0.0) * 1.2
            let total = wrapped |> List.sumBy lineHeight
            let anchor = bodyProps |> Option.bind (attr "anchor") |> Option.defaultValue "t"
            let mutable at =
                match anchor with
                | "ctr" -> y + max padTop ((height - total) / 2.0)
                | "b" -> y + max padTop (height - padBottom - total)
                | _ -> y + padTop
            [ for line in wrapped do
                let align, runs, _ = line
                let high = lineHeight line
                if not runs.IsEmpty then
                    let spans =
                        runs |> List.map (fun (text, style) ->
                            text, sizeOf style, (style.Color |> Option.orElse fontColor |> Option.defaultValue "#000000"), (style.Bold |> Option.defaultValue false), (style.Italic |> Option.defaultValue false))
                    let item =
                        match align with
                        | "ctr" -> PTextRuns (x + width / 2.0, at + high / 2.0, "middle", spans)
                        | "r" -> PTextRuns (x + width - padRight, at + high / 2.0, "end", spans)
                        | _ -> PTextRuns (x + padLeft, at + high / 2.0, "start", spans)
                    // (A line of one typeface other than Calibri is drawn in it.)
                    yield (match runs |> List.map (snd >> faceIn) |> List.distinct with
                           | [ Some face ] when face <> "Calibri" -> PFace (face, [ item ])
                           | _ -> item)
                at <- at + high ]
    match props |> Option.bind (child "xfrm") |> Option.bind (attr "rot") |> Option.bind number with
    | Some turn when turn <> 0.0 ->
        let angle = turn / 60000.0 * Math.PI / 180.0
        let cos, sin = Math.Cos angle, Math.Sin angle
        [ PTransform (cos, sin, -sin, cos, centreX - centreX * cos + centreY * sin, centreY - centreX * sin - centreY * cos, outline @ text) ]
    | _ -> outline @ text

/// What an element of a drawing draws in a rectangle (pixels): a shape, a
/// connector, a picture, a frame with a chart, a group of these. `charts`
/// counts the charts that were drawn, `unknown` what could not be.
let rec private elementItems (parts: DrawingParts) (data: ChartData) (charts: int ref) (unknown: int ref) (element: XmlDsl.XmlNode) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    match nameOf element with
    | "sp" | "cxnSp" -> shapeItems parts data element x y width height
    | "pic" when not (descendants "cameraTool" element).IsEmpty ->
        (match descendants "cameraTool" element |> List.tryHead |> Option.bind (attr "cellRange") |> Option.bind (fun reference -> parts.Camera reference width height) with
         | Some href -> [ PImage (x, y, width, height, href) ]
         | None -> unknown.Value <- unknown.Value + 1; [])
    | "pic" ->
        (match descendants "blip" element |> List.tryHead |> Option.bind (attr "r:embed") |> Option.bind parts.Picture with
         | Some href -> [ PImage (x, y, width, height, href) ]
         | None -> unknown.Value <- unknown.Value + 1; [])
    | "graphicFrame" ->
        // (A slicer or a timeline: what the workbook says of the one of that name.)
        let control =
            descendants "graphicData" element
            |> List.tryHead
            |> Option.bind (fun graphic ->
                kids graphic |> List.tryPick (fun item ->
                    match nameOf item with
                    | "slicer" | "timeslicer" as kind -> attr "name" item |> Option.map (fun name -> kind, name)
                    | _ -> None))
            |> Option.bind (fun (kind, name) -> parts.Slicer kind name data width height)
        match control with
        | Some items ->
            charts.Value <- charts.Value + 1
            [ PGroup (x, y, width, height, items) ]
        | None ->
            (match descendants "chart" element |> List.tryPick (attr "r:id") |> Option.bind parts.Chart with
             | Some chart ->
                charts.Value <- charts.Value + 1
                // The shapes over the chart (`c:userShapes`): each anchored to fractions of the chart's rectangle - from
                // a corner to a corner (`cdr:relSizeAnchor`), or from a corner by its size (`cdr:absSizeAnchor`).
                let shapes =
                    match chart.UserShapes |> Option.bind (fun xml -> try Some (XmlDsl.parse (spacesKept xml)) with _ -> None) with
                    | None -> []
                    | Some root ->
                        let point (node: XmlDsl.XmlNode option) =
                            node |> Option.map (fun corner ->
                                (child "x" corner |> Option.bind (innerText >> number) |> Option.defaultValue 0.0),
                                (child "y" corner |> Option.bind (innerText >> number) |> Option.defaultValue 0.0))
                        kids root
                        |> List.collect (fun anchor ->
                            let rectangle =
                                match nameOf anchor, point (child "from" anchor) with
                                | "relSizeAnchor", Some (x0, y0) ->
                                    point (child "to" anchor) |> Option.map (fun (x1, y1) -> x0 * width, y0 * height, (x1 - x0) * width, (y1 - y0) * height)
                                | "absSizeAnchor", Some (x0, y0) ->
                                    child "ext" anchor |> Option.bind (fun ext ->
                                        match attr "cx" ext |> Option.bind number, attr "cy" ext |> Option.bind number with
                                        | Some cx, Some cy -> Some (x0 * width, y0 * height, emuPx cx, emuPx cy)
                                        | _ -> None)
                                | _ -> None
                            match rectangle with
                            | Some (shapeX, shapeY, shapeW, shapeH) ->
                                kids anchor
                                |> List.collect flattened
                                |> List.filter (fun item -> match nameOf item with "sp" | "cxnSp" | "grpSp" -> true | _ -> false)
                                |> List.collect (fun item -> elementItems parts data (ref 0) (ref 0) item shapeX shapeY shapeW shapeH)
                            | None -> [])
                [ PGroup (x, y, width, height, (layoutChart parts.Theme data chart width height).Items @ shapes) ]
             | None ->
                // (A slicer, a chart of a kind that is not drawn: a box that says what it is.)
                unknown.Value <- unknown.Value + 1
                let name = descendants "cNvPr" element |> List.tryPick (attr "name") |> Option.defaultValue ""
                [ PRect (x + 0.5, y + 0.5, max 0.0 (width - 1.0), max 0.0 (height - 1.0), Some "#F6F8FA", 1.0, Some "#8A94A6", 1.0, 0.0)
                  PText (x + width / 2.0, y + height / 2.0, name, 12.0, "#4A5568", "middle", false, false, 0.0) ])
    | "grpSp" ->
        // The group's own rectangle (`chOff` / `chExt`) is mapped onto the one it is given.
        let transform = child "grpSpPr" element |> Option.bind (child "xfrm")
        let numbers (name: string) (first: string) (second: string) =
            transform |> Option.bind (child name) |> Option.bind (fun node -> match attr first node |> Option.bind number, attr second node |> Option.bind number with | Some a, Some b -> Some (a, b) | _ -> None)
        let childOff = numbers "chOff" "x" "y" |> Option.orElse (numbers "off" "x" "y") |> Option.defaultValue (0.0, 0.0)
        let childExt = numbers "chExt" "cx" "cy" |> Option.orElse (numbers "ext" "cx" "cy") |> Option.defaultValue (width * 9525.0, height * 9525.0)
        let scaleX = if fst childExt > 0.0 then width / fst childExt else 1.0
        let scaleY = if snd childExt > 0.0 then height / snd childExt else 1.0
        kids element
        |> List.collect flattened
        |> List.collect (fun item ->
            match nameOf item with
            | "sp" | "cxnSp" | "pic" | "graphicFrame" | "grpSp" ->
                (match frameOf item with
                 | Some (itemX, itemY, itemW, itemH) ->
                    elementItems parts data charts unknown item (x + (itemX - fst childOff) * scaleX) (y + (itemY - snd childOff) * scaleY) (itemW * scaleX) (itemH * scaleY)
                 | None -> unknown.Value <- unknown.Value + 1; [])
            | _ -> [])
    | _ -> []

/// The scene of an anchor of a drawing (its XML as the file has it:
/// `xdr:twoCellAnchor` ...) for the rectangle it has on the sheet - a chart,
/// a shape, a connector, a group of charts, shapes and pictures. None when
/// the anchor holds nothing this module draws (a chart of a kind it does not
/// know, a slicer, a control): the caller shows a placeholder.
let objectScene (parts: DrawingParts) (data: ChartData) (anchorXml: string) (width: float) (height: float) : Scene option =
    try
        let anchor = XmlDsl.parse (spacesKept anchorXml)
        let charts = ref 0
        let unknown = ref 0
        let content =
            kids anchor
            |> List.collect flattened
            |> List.tryFind (fun item -> match nameOf item with "sp" | "cxnSp" | "pic" | "graphicFrame" | "grpSp" -> true | _ -> false)
        // (Excel does not draw an object marked hidden - `cNvPr hidden="1"`: houseflippingspreadsheet's
        // check marks and "15%" that a macro shows -, nor anything of a shape with no fill, no line and no
        // text: a placeholder for either would hide the cells under it.)
        let hidden (element: XmlDsl.XmlNode) =
            descendants "cNvPr" element |> List.tryHead |> Option.bind (attr "hidden") |> Option.exists (fun value -> value = "1" || value = "true")
        let blankShape (element: XmlDsl.XmlNode) =
            nameOf element = "sp"
            && (descendants "t" element |> List.forall (fun text -> (innerText text).Trim() = ""))
            && (match child "spPr" element with
                | Some properties ->
                    (child "noFill" properties).IsSome
                    && (match child "ln" properties with Some line -> (child "noFill" line).IsSome | None -> false)
                | None -> false)
        match content with
        | None -> None
        | Some element when hidden element || blankShape element -> Some { Width = width; Height = height; Items = [] }
        | Some element ->
            let items = elementItems parts data charts unknown element 0.0 0.0 width height
            // (A frame whose chart is not drawn, a group of nothing known: a placeholder says what it is.)
            if items.IsEmpty || (nameOf element = "graphicFrame" && charts.Value = 0) || (nameOf element = "grpSp" && charts.Value = 0 && unknown.Value > 0 && items.Length <= unknown.Value * 2) then None
            else Some { Width = width; Height = height; Items = items }
    with _ -> None

/// A chart part's shapes (`c:userShapes`): the XML of the drawing part its relationships (`partText` of a path)
/// name as its `chartUserShapes`, in the chart.
let withUserShapes (partText: string -> string option) (chartPath: string) (chart: Chart) : Chart =
    let found =
        partText (PartPath.relationshipsOf chartPath)
        |> Option.bind (fun xml -> try Some (XmlDsl.parse xml) with _ -> None)
        |> Option.bind (fun rels ->
            kids rels
            |> List.tryFind (fun rel -> (attr "Type" rel |> Option.defaultValue "").EndsWith "/chartUserShapes" && attr "TargetMode" rel <> Some "External")
            |> Option.bind (attr "Target"))
        |> Option.bind (fun target -> partText (ZipInterop.resolveTarget (PartPath.directory chartPath) target))
    match found with
    | Some xml -> { chart with UserShapes = Some xml }
    | None -> chart

/// The cells and sheets the charts of an anchor read (`c:f`), by sheet
/// name: what a host watches to draw a chart again when they change.
let chartSheets (chart: Chart) : string list =
    chart.Plots
    |> List.collect (fun plot -> plot.Series)
    |> List.collect (fun series -> [ series.Name; series.Categories; series.Values; series.XValues ] |> List.choose id)
    |> List.choose (fun source -> source.Formula |> Option.bind parseAreas)
    |> List.collect (fun areas -> areas |> List.map (fun (sheet, _, _, _, _) -> sheet))
    |> List.distinct

// ---------------------------------------------------------------------------
// Chart sheets
// ---------------------------------------------------------------------------

/// How a chart sheet is shown (its part's `sheetView`) and the page it prints on (`pageSetup`, `pageMargins`).
type ChartSheetView = {
    /// The chart's page fitted to the window (`zoomToFit`, what Excel writes for a chart sheet by default).
    ZoomToFit: bool
    /// The zoom otherwise (`zoomScale`), in per cent.
    ZoomScale: float
    /// The page less its margins, in points.
    PageWidth: float
    PageHeight: float
}

/// A chart sheet's view from its part: a page of Letter (`paperSize`), landscape unless it says portrait (a chart
/// sheet's default), less margins of 0.7 in left and right and 0.75 in top and bottom unless it names its own.
let chartSheetView (xml: string) : ChartSheetView =
    let root = try Some (XmlDsl.parse xml) with _ -> None
    let find (name: string) = root |> Option.bind (fun node -> descendants name node |> List.tryHead)
    let view = find "sheetView"
    let setup = find "pageSetup"
    let margins = find "pageMargins"
    let margin (name: string) (otherwise: float) = margins |> Option.bind (attr name) |> Option.bind number |> Option.defaultValue otherwise
    let paperWidth, paperHeight = paperSizePoints (setup |> Option.bind (attr "paperSize") |> Option.bind number |> Option.map int |> Option.defaultValue 1)
    let portrait = (setup |> Option.bind (attr "orientation")) = Some "portrait"
    let pageWidth, pageHeight = if portrait then paperWidth, paperHeight else paperHeight, paperWidth
    { ZoomToFit = isOn (view |> Option.bind (attr "zoomToFit")) false
      ZoomScale = view |> Option.bind (attr "zoomScale") |> Option.bind number |> Option.filter (fun zoom -> zoom >= 10.0) |> Option.defaultValue 100.0
      PageWidth = max 1.0 (pageWidth - 72.0 * (margin "left" 0.7 + margin "right" 0.7))
      PageHeight = max 1.0 (pageHeight - 72.0 * (margin "top" 0.75 + margin "bottom" 0.75)) }

/// Where Excel 2019 shows a chart sheet's chart of `chartWidth` x `chartHeight` px (at 100 %) in a window of
/// `windowWidth` x `windowHeight` px: the zoom, the chart's left and top. Fitted to the window (`zoomToFit`), the zoom
/// is the whole per cent at which the page less its margins fits it; the page is centred in the window, the chart in
/// the page (chartsheet-probe.ps1: Letter landscape 117 %, portrait 86 %, A4 121 %, margins of 2 in 149 % in a
/// window of 1898 x 787 px - each the page's fit, not the chart's -; the chart at 415 px of 1898, 5 px down).
let chartSheetPlacement (view: ChartSheetView) (windowWidth: float) (windowHeight: float) (chartWidth: float) (chartHeight: float) : float * float * float =
    let pageWidth, pageHeight = view.PageWidth * 4.0 / 3.0, view.PageHeight * 4.0 / 3.0
    let zoom =
        if view.ZoomToFit then max 0.1 (Math.Floor (100.0 * min (windowWidth / pageWidth) (windowHeight / pageHeight) + 1e-9) / 100.0)
        else view.ZoomScale / 100.0
    let left = max 0.0 ((windowWidth - pageWidth * zoom) / 2.0) + (pageWidth - chartWidth) * zoom / 2.0
    let top = max 0.0 ((windowHeight - pageHeight * zoom) / 2.0) + (pageHeight - chartHeight) * zoom / 2.0
    zoom, max 0.0 left, max 0.0 top
