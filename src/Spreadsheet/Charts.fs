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
    | XmlDsl.Element (_, _, children) -> children
    | _ -> []

let private nameOf (node: XmlDsl.XmlNode) : string =
    match node with
    | XmlDsl.Element (name, _, _) -> localName name
    | _ -> ""

let private childrenNamed (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    kids node |> List.filter (fun child -> nameOf child = name)

let private child (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    kids node |> List.tryFind (fun child -> nameOf child = name)

let private attr (name: string) (node: XmlDsl.XmlNode) : string option =
    match node with
    | XmlDsl.Element (_, attrs, _) -> Map.tryFind name attrs
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
    | XmlDsl.Element (_, _, children) -> children |> List.map innerText |> String.concat ""
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

/// The colours of a theme part (`a:clrScheme`).
let parseTheme (xml: string) : ThemeColors =
    try
        match descendants "clrScheme" (XmlDsl.parse xml) |> List.tryHead with
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
    with _ -> defaultTheme

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
let private colorOfNode (theme: ThemeColors) (node: XmlDsl.XmlNode) : (string * float) option =
    let baseColor =
        match nameOf node with
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
            | "tint" -> (r <- r + (255.0 - r) * (1.0 - amount); g <- g + (255.0 - g) * (1.0 - amount); b <- b + (255.0 - b) * (1.0 - amount))
            | "shade" -> (r <- r * amount; g <- g * amount; b <- b * amount)
            | "alpha" -> alpha <- amount
            | _ -> ()
        "#" + hex2 r + hex2 g + hex2 b, alpha)

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
let private fillWith (gradients: bool) (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill =
    let colorIn (node: XmlDsl.XmlNode) = kids node |> List.tryPick (colorOfNode theme)
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

let private fillOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill = fillWith false theme holder

/// The fill of a shape, a series, a background: a gradient is one.
let private shapeFillOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill = fillWith true theme holder

let private strokeOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Stroke option =
    child "ln" holder
    |> Option.map (fun line ->
        { Paint = fillOf theme line
          Width = (match attr "w" line |> Option.bind number with Some emu -> max 0.75 (emu / 9525.0) | None -> 1.0)
          Dash = (match valOf "prstDash" line with Some "solid" | None -> None | Some other -> Some other) })

/// The font size (points), the colour and the weight a text properties
/// element (`c:txPr`, `a:rPr`, `a:defRPr`) states.
type TextStyle = {
    Size: float option
    Color: string option
    Bold: bool option
    Italic: bool option
}

let private noTextStyle = { Size = None; Color = None; Bold = None; Italic = None }

let private runStyle (theme: ThemeColors) (props: XmlDsl.XmlNode) : TextStyle =
    { Size = attr "sz" props |> Option.bind number |> Option.map (fun value -> value / 100.0)
      Color = (match fillOf theme props with FillColor (color, _) -> Some color | _ -> None)
      Bold = attr "b" props |> Option.map (fun value -> value = "1" || value = "true")
      Italic = attr "i" props |> Option.map (fun value -> value = "1" || value = "true") }

let private over (first: TextStyle) (second: TextStyle) : TextStyle =
    { Size = first.Size |> Option.orElse second.Size; Color = first.Color |> Option.orElse second.Color
      Bold = first.Bold |> Option.orElse second.Bold; Italic = first.Italic |> Option.orElse second.Italic }

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
    Labels: LabelSpec option
    /// The sizes of the bubbles of a bubble chart's series.
    BubbleSizes: DataRef option
    /// A pie slice pulled out, in per cent of the radius.
    Explosion: float
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
    LabelStyle: TextStyle
    Line: Stroke option
    LabelsShown: bool
    CrossesId: string option
    /// The base of a logarithmic scale (`c:logBase`).
    LogBase: float option
    /// What the labels are divided by (`c:dispUnits`: 1000 for thousands), and the label that says so.
    DisplayUnit: float
    DisplayUnitLabel: string option
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
}

type Chart = {
    Title: string option
    TitleStyle: TextStyle
    AutoTitleDeleted: bool
    Plots: Plot list
    Axes: Axis list
    Legend: Legend option
    Background: Fill
    Border: Stroke option
    PlotBackground: Fill
    TextStyle: TextStyle
    RoundedCorners: bool
    Date1904: bool
    TitleLayout: ManualLayout option
    /// The title is drawn over the plot area: no room is made for it.
    TitleOverlay: bool
    PlotLayout: ManualLayout option
    /// `c:style` (1 .. 48): which colours the series get that state none.
    Style: int
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
        (match childrenNamed "lvl" cache with
         | [] -> points cache
         | level :: _ -> points level)
        (match valOf "ptCount" cache |> Option.bind number with Some stated when int stated > count -> count <- int stated | _ -> ())
        { Formula = formula; Numbers = numbers; Strings = strings; Count = count
          FormatCode = child "formatCode" cache |> Option.map innerText |> Option.filter (fun code -> code <> "") }
    let reference (node: XmlDsl.XmlNode) =
        let formula = child "f" node |> Option.map innerText |> Option.filter (fun text -> text.Trim() <> "")
        match kids node |> List.tryFind (fun item -> (nameOf item).EndsWith "Cache") with
        | Some cache -> cached cache formula
        | None -> { Formula = formula; Numbers = Map.empty; Strings = Map.empty; Count = 0; FormatCode = None }
    kids holder
    |> List.tryPick (fun item ->
        match nameOf item with
        | "numRef" | "strRef" | "multiLvlStrRef" -> Some (reference item)
        | "numLit" | "strLit" -> Some (cached item None)
        | "v" -> Some { Formula = None; Numbers = Map.empty; Strings = Map.ofList [ 0, innerText item ]; Count = 1; FormatCode = None }
        | _ -> None)

let private labelsOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : LabelSpec option =
    child "dLbls" holder
    |> Option.bind (fun labels ->
        if isOn (valOf "delete" labels) false then None
        else
            let spec =
                { ShowValue = isOn (valOf "showVal" labels) false
                  ShowPercent = isOn (valOf "showPercent" labels) false
                  ShowCategory = isOn (valOf "showCatName" labels) false
                  ShowSeries = isOn (valOf "showSerName" labels) false
                  Position = valOf "dLblPos" labels
                  FormatCode = child "numFmt" labels |> Option.bind (fun format -> if isOn (attr "sourceLinked" format) false then None else attr "formatCode" format)
                  Style = textPropsOf theme labels
                  Separator = child "separator" labels |> Option.map innerText
                  Deleted =
                    childrenNamed "dLbl" labels
                    |> List.choose (fun label -> if isOn (valOf "delete" label) false then valOf "idx" label |> Option.bind number |> Option.map int else None)
                    |> Set.ofList }
            if spec.ShowValue || spec.ShowPercent || spec.ShowCategory || spec.ShowSeries then Some spec else None)

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
      MarkerSize = marker |> Option.bind (valOf "size") |> Option.bind number |> Option.defaultValue 5.0
      MarkerFill = markerShape |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
      MarkerLine = markerShape |> Option.bind (strokeOf theme)
      Smooth = isOn (valOf "smooth" node) false
      PointFills =
        childrenNamed "dPt" node
        |> List.choose (fun point ->
            match valOf "idx" point |> Option.bind number, child "spPr" point |> Option.map (shapeFillOf theme) with
            | Some index, Some fill when fill <> FillAuto -> Some (int index, fill)
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

let private axisOf (theme: ThemeColors) (node: XmlDsl.XmlNode) : Axis option =
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
                | None -> { Paint = FillColor ("#D9D9D9", 1.0); Width = 1.0; Dash = None })
          Title = child "title" node |> Option.bind titleText
          TitleStyle = child "title" node |> Option.map (titleStyle theme) |> Option.defaultValue noTextStyle
          LabelStyle = textPropsOf theme node
          Line = child "spPr" node |> Option.bind (strokeOf theme)
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
                |> Option.filter (fun text -> text <> "")) })

/// The chart of a chart part, or None for a part this module does not draw
/// (no plot of a kind it knows, a 3-D chart, a part that is not a chart).
/// Whether a chart part is one of the newer kinds (`cx:chartSpace`: waterfall, funnel, treemap,
/// sunburst, histogram, box and whisker, map), told by its first tag - such a part can hold
/// megabytes of data (a map chart's geography), and is not read to find that out.
let isChartEx (xml: string) : bool =
    let head = if xml.Length > 4096 then xml.Substring(0, 4096) else xml
    let mutable at = head.IndexOf '<'
    while at >= 0 && at + 1 < head.Length && (head.[at + 1] = '?' || head.[at + 1] = '!') do
        at <- head.IndexOf('<', at + 1)
    if at < 0 then false
    else
        let close = head.IndexOf('>', at)
        let tag = if close < 0 then head.Substring at else head.Substring(at, close - at)
        let name = (tag.Substring 1).Split([| ' '; '\t'; '\r'; '\n'; '/' |]).[0]
        let space = "http://schemas.microsoft.com/office/drawing/2014/chartex"
        match name.IndexOf ':' with
        | -1 -> tag.Contains ("xmlns=\"" + space + "\"") || tag.Contains ("xmlns='" + space + "'")
        | colon ->
            let prefix = name.Substring(0, colon)
            tag.Contains ("xmlns:" + prefix + "=\"" + space + "\"") || tag.Contains ("xmlns:" + prefix + "='" + space + "'")

let parseChart (theme: ThemeColors) (xml: string) : Chart option =
    if isChartEx xml then None else
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
                    let shape = child "spPr" space
                    Some { Title = child "title" chart |> Option.bind titleText
                           TitleStyle = child "title" chart |> Option.map (titleStyle theme) |> Option.defaultValue noTextStyle
                           AutoTitleDeleted = isOn (valOf "autoTitleDeleted" chart) false || (child "title" chart).IsNone
                           Plots = known
                           Axes = items |> List.choose (axisOf theme)
                           Legend =
                             child "legend" chart
                             |> Option.map (fun legend ->
                                { Position = valOf "legendPos" legend |> Option.defaultValue "r"; Style = textPropsOf theme legend
                                  Layout = layoutOf legend; Overlay = isOn (valOf "overlay" legend) false
                                  Deleted =
                                    childrenNamed "legendEntry" legend
                                    |> List.choose (fun entry -> if isOn (valOf "delete" entry) false then valOf "idx" entry |> Option.bind number |> Option.map int else None)
                                    |> Set.ofList })
                           Background = shape |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                           Border = shape |> Option.bind (strokeOf theme)
                           PlotBackground = child "spPr" area |> Option.map (shapeFillOf theme) |> Option.defaultValue FillAuto
                           TextStyle = textPropsOf theme space
                           RoundedCorners = isOn (valOf "roundedCorners" space) false
                           Date1904 = isOn (valOf "date1904" space) false
                           TitleLayout = child "title" chart |> Option.bind layoutOf
                           TitleOverlay = child "title" chart |> Option.map (fun title -> isOn (valOf "overlay" title) false) |> Option.defaultValue false
                           PlotLayout = layoutOf area
                           Style =
                             // (`c:style` of the part, or the one its `mc:AlternateContent` falls back to.)
                             kids space |> List.collect flattened |> List.tryPick (fun item -> if nameOf item = "style" then attr "val" item |> Option.bind number |> Option.map int else None)
                             |> Option.defaultValue 2 }
    with _ -> None

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
    source.Formula
    |> Option.bind parseAreas
    |> Option.bind (fun areas ->
        let blocks = areas |> List.map (fun (sheet, r0, c0, r1, c1) -> data.Cells sheet r0 c0 r1 c1)
        if blocks |> List.exists Option.isNone then None
        else
            let cells = blocks |> List.collect (fun block -> block.Value |> Array.toList |> List.collect Array.toList)
            // (As many points as the file cached: the range a pivot chart names runs on over the
            // pivot table's grand total - 13 cells for 12 points -, which is no point of the chart.)
            Some (if source.Count > 0 && source.Count < cells.Length then List.truncate source.Count cells else cells))

let rec private numberOfValue (value: CellValue) : float option =
    match value with
    | CellValue.Number n -> Some n
    | CellValue.Formula (_, result) -> numberOfValue result
    | CellValue.Boolean _ -> None
    | _ -> None

/// The numbers of a series: live when its cells can be read, else the file's cache. A point without a number is None.
let private numbersOf (data: ChartData) (source: DataRef option) : float option [] =
    match source with
    | None -> [||]
    | Some source ->
        match cellsOf data source with
        | Some cells -> cells |> List.map (fun cell -> cell |> Option.bind (fun cell -> numberOfValue cell.Value)) |> List.toArray
        | None -> Array.init source.Count (fun index -> Map.tryFind index source.Numbers)

/// The texts of categories (or of a series name).
let private textsOf (data: ChartData) (source: DataRef option) : string [] =
    match source with
    | None -> [||]
    | Some source ->
        match cellsOf data source with
        | Some cells -> cells |> List.map (fun cell -> match cell with Some cell -> data.Display cell | None -> "") |> List.toArray
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
        let l = if round % 2 = 1 then l * 0.6 else min 0.9 (l + (1.0 - l) * 0.4)
        let r, g, b = rgbOfHsl (h, s, l)
        "#" + hex2 r + hex2 g + hex2 b

/// The colour of the n-th of `count` series (or points) under a chart style (`c:style`, 1 .. 48,
/// in rows of eight): the first of a row is shades of grey, the second the accents in turn, the
/// others one accent from dark to light.
let private styleColor (theme: ThemeColors) (style: int) (index: int) (count: int) : string =
    let kind = ((max 1 style - 1) % 8)
    if kind = 1 then autoColor theme index
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

/// The scale of a value axis over lo .. hi: its ends and the step of its
/// ticks, as Excel chooses them (zero is an end unless the values are
/// close together far from it; room above the largest value).
let private valueScale (lo: float) (hi: float) (fixedMin: float option) (fixedMax: float option) (fixedStep: float option) : float * float * float =
    let lo, hi = (if lo > hi then hi, lo else lo, hi)
    let lo, hi = (if lo = hi then (if lo = 0.0 then 0.0, 1.0 elif lo > 0.0 then 0.0, hi else lo, 0.0) else lo, hi)
    let low = if lo >= 0.0 then (if (hi - lo) / hi > 0.1667 then 0.0 else lo) else lo
    let high = if hi <= 0.0 then (if (hi - lo) / (-lo) > 0.1667 then 0.0 else hi) else hi
    let low = defaultArg fixedMin low
    let high = defaultArg fixedMax high
    let range = if high > low then high - low else 1.0
    let step =
        match fixedStep with
        | Some step when step > 0.0 -> step
        | _ ->
            let rough = range / 10.0
            let power = Math.Pow(10.0, Math.Floor (Math.Log10 rough))
            let unit = rough / power
            (if unit <= 1.0 then 1.0 elif unit <= 2.0 then 2.0 elif unit <= 5.0 then 5.0 else 10.0) * power
    let minimum =
        match fixedMin with
        | Some value -> value
        | None -> if low = 0.0 then 0.0 else Math.Floor ((low - (if lo < 0.0 then range * 0.05 else 0.0)) / step) * step
    let maximum =
        match fixedMax with
        | Some value -> value
        | None -> if high = 0.0 then 0.0 else Math.Ceiling ((high + range * 0.05) / step) * step
    let maximum = if maximum <= minimum then minimum + step else maximum
    // (Never more ticks than can be read.)
    let step = if (maximum - minimum) / step > 40.0 then (maximum - minimum) / 10.0 else step
    minimum, maximum, step

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

/// The scene of a chart in a rectangle of `width` x `height` pixels.
let layoutChart (theme: ThemeColors) (data: ChartData) (chart: Chart) (width: float) (height: float) : Scene =
    let items = ResizeArray<Primitive>()
    let add (item: Primitive) = items.Add item
    let textColor (style: TextStyle) = over style chart.TextStyle |> fun merged -> merged.Color |> Option.defaultValue "#595959"
    let textSize (style: TextStyle) (otherwise: float) = px (over style chart.TextStyle |> fun merged -> merged.Size |> Option.defaultValue otherwise)

    // ---- the chart's own rectangle ----
    (match chart.Background with
     | FillNone -> ()
     | FillColor (color, opacity) -> add (PRect (0.0, 0.0, width, height, Some color, opacity, None, 0.0, 0.0))
     | FillAuto -> add (PRect (0.0, 0.0, width, height, Some "#FFFFFF", 1.0, None, 0.0, 0.0)))
    (match strokeColor chart.Border with
     | Some (color, strokeWidth, _) -> add (PRect (strokeWidth / 2.0, strokeWidth / 2.0, width - strokeWidth, height - strokeWidth, None, 1.0, Some color, strokeWidth, (if chart.RoundedCorners then 8.0 else 0.0)))
     | None -> ())

    // ---- the series ----
    let axisById (id: string) = chart.Axes |> List.tryFind (fun axis -> axis.Id = id)
    let valueAxes = chart.Axes |> List.filter (fun axis -> axis.Kind = "val")
    let pointPlot (plot: Plot) = plot.Kind = ScatterPlot || plot.Kind = BubblePlot
    let primaryValueAxis =
        // (The value axis of the first plot; on a scatter chart the one at the side.)
        let ofPlot (plot: Plot) = plot.AxisIds |> List.choose axisById |> List.filter (fun axis -> axis.Kind = "val")
        match chart.Plots with
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
                        | None -> { Color = auto; Opacity = 1.0; Filled = false; LineColor = Some auto; LineWidth = 2.25; LineDash = None }
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
                  XValues = numbersOf data source.XValues
                  Sizes = numbersOf data source.BubbleSizes
                  Look = look
                  FormatCode = formatCodeOf data source.Values
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
    let mutable top = 6.0
    (match title with
     | Some text ->
        let size = textSize chart.TitleStyle 14.0
        let bold = chart.TitleStyle.Bold |> Option.defaultValue false
        let lines = text.Split('\n')
        let blockHeight = float lines.Length * size * 1.25
        let widest = lines |> Array.fold (fun most line -> max most (calibriWidth size bold line)) 0.0
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
            | Some { Y = Some y } -> top + y * height
            | _ -> top
        let y = max 0.0 (min y (height - blockHeight))
        lines |> Array.iteri (fun index line ->
            add (PText (x, y + size * 0.7 + float index * size * 1.25, line, size, textColor chart.TitleStyle, anchor, bold, (chart.TitleStyle.Italic |> Option.defaultValue false), 0.0)))
        // (Room is made for a title at the top that is not laid over the plot.)
        if not chart.TitleOverlay && y < height * 0.4 then top <- max top (y + blockHeight + 8.0)
     | None -> ())

    // ---- legend ----
    let legendEntries : (string * string * bool) list =
        // (text, colour, a line rather than a box)
        let all =
            if isPie then
                match series with
                | first :: _ -> categories |> Array.toList |> List.mapi (fun index text -> text, fst (pointColor first index), false)
                | [] -> []
            else
                series |> List.map (fun item ->
                    item.Label, (match item.Look.LineColor with Some color when not item.Look.Filled -> color | _ -> item.Look.Color),
                    (item.Plot.Kind = LinePlot || item.Plot.Kind = ScatterPlot || item.Plot.Kind = StockPlot || (item.Plot.Kind = RadarPlot && item.Plot.RadarStyle <> "filled")))
        match chart.Legend with
        | Some legend when not legend.Deleted.IsEmpty -> all |> List.indexed |> List.filter (fun (index, _) -> not (Set.contains index legend.Deleted)) |> List.map snd
        | _ -> all
    let mutable left = 8.0
    let mutable right = width - 8.0
    let mutable bottom = height - 8.0
    (match chart.Legend with
     | Some legend when not legendEntries.IsEmpty ->
        let size = textSize legend.Style 9.0
        let color = textColor legend.Style
        let swatch = size * 0.75
        let keyWidth (line: bool) = if line then swatch * 1.6 else swatch
        let entryWidth (text: string, _: string, line: bool) = keyWidth line + 5.0 + textWidth size text + 14.0
        let drawEntry (x: float) (y: float) (room: float) (text: string, entryColor: string, line: bool) =
            if line then add (PLine (x, y, x + swatch * 1.6, y, entryColor, 2.25, None))
            else add (PRect (x, y - swatch / 2.0, swatch, swatch, Some entryColor, 1.0, None, 0.0, 0.0))
            // (A text longer than its place ends in an ellipsis.)
            let textRoom = room - keyWidth line - 5.0
            let shownText =
                if textRoom > 12.0 && textWidth size text > textRoom + 14.0 && text.Length > 3 then
                    let mutable keep = text.Length - 1
                    while keep > 1 && textWidth size (text.Substring(0, keep) + "…") > textRoom + 10.0 do keep <- keep - 1
                    text.Substring(0, keep) + "…"
                else text
            add (PText (x + keyWidth line + 5.0, y, shownText, size, color, "start", false, false, 0.0))
        let horizontal = legend.Position = "b" || legend.Position = "t"
        // The rectangle the file gives the legend (`c:layout`), or the side of the chart it is at.
        let box =
            match legend.Layout with
            | Some { XEdge = true; YEdge = true; X = Some x; Y = Some y; W = w; H = h } ->
                let boxLeft, boxTop = max 0.0 (min (width - 10.0) (x * width)), max 0.0 (min (height - 10.0) (y * height))
                Some (boxLeft, boxTop, (w |> Option.map (fun value -> min (width - boxLeft) (value * width))), (h |> Option.map (fun value -> min (height - boxTop) (value * height))))
            | _ -> None
        if horizontal then
            // Rows of entries, centred.
            let rowHeight = size * 1.5
            let room = match box with Some (_, _, Some boxWidth, _) -> boxWidth | _ -> width - 16.0
            let rows = ResizeArray<ResizeArray<string * string * bool>>()
            let mutable current = ResizeArray<string * string * bool>()
            let mutable used = 0.0
            for entry in legendEntries do
                let entryW = entryWidth entry
                if used + entryW > room + 14.0 && current.Count > 0 then
                    rows.Add current
                    current <- ResizeArray<string * string * bool>()
                    used <- 0.0
                current.Add entry
                used <- used + entryW
            if current.Count > 0 then rows.Add current
            let rows = rows |> Seq.truncate 6 |> Seq.toList
            let blockHeight = float rows.Length * rowHeight
            let boxLeft, boxTop, boxWidth, boxHeight =
                match box with
                | Some (boxLeft, boxTop, boxWidth, boxHeight) -> boxLeft, boxTop, (boxWidth |> Option.defaultValue (width - boxLeft)), (boxHeight |> Option.defaultValue blockHeight)
                | None -> 0.0, (if legend.Position = "b" then bottom - blockHeight else top), width, blockHeight
            let startY = boxTop + max 0.0 ((boxHeight - blockHeight) / 2.0)
            rows |> List.iteri (fun rowIndex row ->
                let total = (row |> Seq.sumBy entryWidth) - 14.0
                let mutable x = max 4.0 (boxLeft + (boxWidth - total) / 2.0)
                for entry in row do
                    drawEntry x (startY + rowHeight * (float rowIndex + 0.5)) (entryWidth entry - 14.0) entry
                    x <- x + entryWidth entry)
            if not legend.Overlay then
                if legend.Position = "b" then bottom <- min bottom (boxTop - 4.0)
                else top <- max top (boxTop + max blockHeight boxHeight + 4.0)
        else
            let position = legend.Position
            let widest = legendEntries |> List.fold (fun most entry -> max most (entryWidth entry)) 0.0
            let boxLeft, boxTop, boxWidth, boxHeight =
                match box with
                | Some (boxLeft, boxTop, boxWidth, boxHeight) ->
                    boxLeft, boxTop, (boxWidth |> Option.defaultValue (min widest (width - boxLeft))), (boxHeight |> Option.defaultValue (bottom - boxTop))
                | None ->
                    let blockWidth = min (width * 0.4) widest
                    (if position = "l" then left else right - blockWidth + 4.0), top, blockWidth, bottom - top
            // (Every entry is there: the lines move together where the legend is short.)
            let rowHeight = max (size * 1.05) (min (size * 1.6) (boxHeight / float legendEntries.Length))
            let shown = legendEntries |> List.truncate (max 1 (int ((boxHeight + 1.0) / rowHeight)))
            let blockHeight = float shown.Length * rowHeight
            let startY = if position = "tr" && box.IsNone then boxTop else boxTop + max 0.0 ((boxHeight - blockHeight) / 2.0)
            shown |> List.iteri (fun index entry -> drawEntry (boxLeft + 2.0) (startY + rowHeight * (float index + 0.5)) (boxWidth - 4.0) entry)
            if not legend.Overlay then
                if position = "l" then left <- max left (boxLeft + boxWidth + 6.0) else right <- min right (boxLeft - 2.0)
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
        |> String.concat (spec.Separator |> Option.defaultValue "; ")
    let labelled (item: SeriesData) (index: int) : LabelSpec option =
        labelStyleOf item |> Option.filter (fun spec -> not (Set.contains index spec.Deleted))
    let plotRect = layoutRect chart.PlotLayout width height

    if isPie then
        // ---- pie, doughnut ----
        let rings = series |> List.filter (fun item -> item.Plot.Kind = PiePlot || item.Plot.Kind = DoughnutPlot)
        let rings = (match firstPlot with Some plot when plot.Kind = PiePlot -> List.truncate 1 rings | _ -> rings)
        let outside = rings |> List.exists (fun item -> match labelStyleOf item with Some spec -> spec.Position = Some "outEnd" | None -> false)
        let cx, cy, radius =
            match plotRect with
            | Some (x, y, w, h) -> x + w / 2.0, y + h / 2.0, max 4.0 (min w h / 2.0)
            | None ->
                let room = if outside then 18.0 else 4.0
                (left + right) / 2.0, (top + bottom) / 2.0, max 4.0 (min (right - left) (bottom - top) / 2.0 - room)
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
                let separator = (match strokeColor item.Source.Line with Some (color, strokeWidth, _) -> Some color, strokeWidth | None -> Some "#FFFFFF", (if values.Length > 1 then 1.0 else 0.0))
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
                            add (PPath (d, Some color, opacity, fst separator, snd separator))
                        (match labelled item index with
                         | Some spec when sweep > 0.12 ->
                            let size = textSize spec.Style 9.0
                            let outsideLabel = spec.Position = Some "outEnd"
                            let at =
                                if outsideLabel then outer + size * 0.9
                                elif inner > 0.0 then (inner + outer) / 2.0
                                elif spec.Position = Some "inEnd" then outer * 0.82
                                elif spec.Position = Some "ctr" then outer * 0.5
                                else outer * 0.68
                            let x, y = point at middle
                            let anchor = if not outsideLabel then "middle" elif Math.Sin middle > 0.3 then "start" elif Math.Sin middle < -0.3 then "end" else "middle"
                            add (PText (x, y, labelText item spec index value (Some (value / total)), size, (spec.Style.Color |> Option.defaultValue (if outsideLabel then "#404040" else "#FFFFFF")), anchor, (spec.Style.Bold |> Option.defaultValue false), false, 0.0))
                         | _ -> ())
                        angle <- angle + sweep))
    elif isRadar then
        // ---- radar: the categories around a centre, the values along the spokes ----
        let count = pointCount
        let lo, hi =
            let values = series |> List.collect (fun item -> item.Values |> Array.choose id |> Array.toList)
            if values.IsEmpty then 0.0, 1.0 else List.min values, List.max values
        let vMin, vMax, vStep = valueScale lo hi (primaryValueAxis |> Option.bind (fun a -> a.Min)) (primaryValueAxis |> Option.bind (fun a -> a.Max)) (primaryValueAxis |> Option.bind (fun a -> a.MajorUnit))
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
        let count = pointCount
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
        let scaleFor (axis: Axis option) (secondary: bool) : AxisScale =
            let percent = chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            match axis with
            | Some ({ LogBase = Some logBase } as logAxis) when not percent ->
                (match extent secondary true with
                 | Some (lo, hi) -> logScale logBase lo hi logAxis
                 | None -> { Min = 0.0; Max = 1.0; Step = 1.0; Log = Some logBase })
            | _ ->
                match extent secondary false with
                | None -> { Min = 0.0; Max = 1.0; Step = 0.2; Log = None }
                | Some (lo, hi) ->
                    if percent then { Min = (if lo < 0.0 then -1.0 else 0.0); Max = 1.0; Step = 0.1; Log = None }
                    else
                        let minimum, maximum, step = valueScale lo hi (axis |> Option.bind (fun a -> a.Min)) (axis |> Option.bind (fun a -> a.Max)) (axis |> Option.bind (fun a -> a.MajorUnit))
                        { Min = minimum; Max = maximum; Step = step; Log = None }
        let primaryScale = scaleFor primaryValueAxis false
        let secondaryScale = scaleFor secondaryAxis true
        let xScale : AxisScale =
            if not isScatter then { Min = 0.0; Max = 1.0; Step = 1.0; Log = None }
            else
                let values = series |> List.collect (fun item -> item.XValues |> Array.choose id |> Array.toList)
                match values, xAxis with
                | [], _ -> { Min = 0.0; Max = float (max 1 count); Step = 1.0; Log = None }
                | _, Some ({ LogBase = Some logBase } as logAxis) when values |> List.exists (fun value -> value > 0.0) ->
                    let positive = values |> List.filter (fun value -> value > 0.0)
                    logScale logBase (List.min positive) (List.max positive) logAxis
                | _ ->
                    let minimum, maximum, step = valueScale (List.min values) (List.max values) (xAxis |> Option.bind (fun a -> a.Min)) (xAxis |> Option.bind (fun a -> a.Max)) (xAxis |> Option.bind (fun a -> a.MajorUnit))
                    { Min = minimum; Max = maximum; Step = step; Log = None }
        let axisFormat (axis: Axis option) (secondary: bool) : float -> string =
            let percent = chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            let linked = series |> List.tryPick (fun item -> if item.Secondary = secondary then item.FormatCode else None)
            let code =
                match axis with
                | Some axis when not axis.SourceLinked -> axis.FormatCode |> Option.filter (fun code -> code <> "General")
                | Some axis -> linked |> Option.orElse (axis.FormatCode |> Option.filter (fun code -> code <> "General"))
                | None -> linked
            // (The labels of an axis in thousands, millions ... say the value divided by that.)
            let unit = match axis with Some axis when axis.DisplayUnit > 0.0 -> axis.DisplayUnit | _ -> 1.0
            match code with
            | _ when percent -> (fun value -> generalNumber (Math.Round (value * 100.0)) + "%")
            | Some code -> (fun value -> try data.Format code (value / unit) with _ -> generalNumber (value / unit))
            | None -> (fun value -> generalNumber (value / unit))
        let valueLabels = ticksOf primaryScale |> List.map (axisFormat primaryValueAxis false)
        let secondaryLabels = if secondaryAxis.IsSome then ticksOf secondaryScale |> List.map (axisFormat secondaryAxis true) else []
        let shown (axis: Axis option) = match axis with Some axis -> not axis.Deleted && axis.LabelsShown | None -> false
        let valueSize = textSize (primaryValueAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let categorySize = textSize (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let valueColor = textColor (primaryValueAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle)
        let categoryColor = textColor (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle)
        let widest (labels: string list) (size: float) = labels |> List.fold (fun most text -> max most (textWidth size text)) 0.0
        // Axis titles (and what says that an axis is in thousands, millions ...).
        let axisTitle (axis: Axis option) =
            axis |> Option.bind (fun axis ->
                if axis.Deleted then None
                else
                    match axis.Title, axis.DisplayUnitLabel with
                    | Some text, Some units -> Some (text + " (" + units + ")", textSize axis.TitleStyle 10.0, textColor axis.TitleStyle)
                    | Some text, None -> Some (text, textSize axis.TitleStyle 10.0, textColor axis.TitleStyle)
                    | None, Some units -> Some (units, textSize axis.TitleStyle 10.0, textColor axis.TitleStyle)
                    | None, None -> None)
        let sideAxis, baseAxis = (if horizontal then categoryAxis, primaryValueAxis elif isScatter then primaryValueAxis, xAxis else primaryValueAxis, categoryAxis)
        // (A plot area the file places leaves the titles of its axes where they are by themselves.)
        (match axisTitle sideAxis with
         | Some (text, size, color) ->
            add (PText (left + size * 0.6, (match plotRect with Some (_, y, _, h) -> y + h / 2.0 | None -> (top + bottom) / 2.0), text, size, color, "middle", false, false, -90.0))
            left <- left + size * 1.5
         | None -> ())
        (match axisTitle baseAxis with
         | Some (text, size, color) ->
            add (PText ((match plotRect with Some (x, _, w, _) -> x + w / 2.0 | None -> (left + right) / 2.0), bottom - size * 0.6, text, size, color, "middle", false, false, 0.0))
            bottom <- bottom - size * 1.5
         | None -> ())
        (match (if secondaryAxis.IsSome then axisTitle secondaryAxis else None) with
         | Some (text, size, color) ->
            add (PText (right - size * 0.6, (match plotRect with Some (_, y, _, h) -> y + h / 2.0 | None -> (top + bottom) / 2.0), text, size, color, "middle", false, false, -90.0))
            right <- right - size * 1.5
         | None -> ())
        // The room the tick labels take.
        let categoryLabels = categories |> Array.toList
        let xLabels = if isScatter then ticksOf xScale |> List.map (axisFormat xAxis false) else []
        let sideLabels, baseLabels = (if horizontal then categoryLabels, valueLabels elif isScatter then valueLabels, xLabels else valueLabels, categoryLabels)
        let sideSize, baseSize = (if horizontal then categorySize, valueSize else valueSize, categorySize)
        let sideShown, baseShown = (if horizontal then shown categoryAxis, shown primaryValueAxis elif isScatter then shown primaryValueAxis, shown xAxis else shown primaryValueAxis, shown categoryAxis)
        let sideRoom = if sideShown then min (width * 0.4) (widest sideLabels sideSize) + 8.0 else 4.0
        let baseRoom = if baseShown then baseSize * 1.25 + 6.0 else 4.0
        let secondaryRoom = if secondaryAxis.IsSome && shown secondaryAxis then widest secondaryLabels valueSize + 8.0 else 0.0
        // (The last label of a value axis along the bottom is centred on the plot's edge: half of it is beyond.)
        let overhang =
            if baseShown && (horizontal || isScatter) && secondaryRoom = 0.0 then
                (match List.tryLast baseLabels with Some text -> textWidth baseSize text / 2.0 | None -> 0.0)
            else 0.0
        let plotLeft, plotTop, plotRight, plotBottom =
            match chart.PlotLayout, plotRect with
            | Some { Inner = true }, Some (x, y, w, h) -> x, y, x + w, y + h
            | _ ->
                // (A rectangle for the plot area with its labels: the labels are inside it.)
                let left, top, right, bottom = (match plotRect with Some (x, y, w, h) -> max 2.0 x, max 2.0 y, min (width - 2.0) (x + w), min (height - 2.0) (y + h) | None -> left, top, right, bottom)
                let plotLeft = left + sideRoom
                let plotRight = max (plotLeft + 10.0) (right - secondaryRoom - (if baseShown && not horizontal && not isScatter then 0.0 else max 6.0 overhang) - 4.0)
                let plotTop = top + (if title.IsNone && plotRect.IsNone then 6.0 else 0.0)
                let plotBottom = max (plotTop + 10.0) (bottom - baseRoom)
                plotLeft, plotTop, plotRight, plotBottom
        let plotWidth, plotHeight = max 1.0 (plotRight - plotLeft), max 1.0 (plotBottom - plotTop)
        (match chart.PlotBackground with
         | FillColor (color, opacity) -> add (PRect (plotLeft, plotTop, plotWidth, plotHeight, Some color, opacity, None, 0.0, 0.0))
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
        let onTicks = not horizontal && count > 1 && chart.Plots |> List.forall (fun plot -> plot.Kind <> BarPlot) && valueAxes |> List.exists (fun axis -> axis.OnTicks)
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
                let groupWidth = band / (1.0 + max 0.0 plot.GapWidth / 100.0)
                let barWidth = if stacked then groupWidth else groupWidth / (float n - float (n - 1) * overlap)
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
                            for line in runs do (if line.Length > 1 then add (PPolyline (line, None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash)))
                         | _ -> ())
                        // Markers: the chart's switch, unless the series says none.
                        let symbol = item.Source.MarkerSymbol |> Option.defaultValue (if plot.Kind = LinePlot && plot.Markers then "circle" else "none")
                        if symbol <> "none" then
                            let size = max 3.0 (item.Source.MarkerSize * 1.1)
                            let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                            let stroke = (match strokeColor item.Source.MarkerLine with Some (color, _, _) -> Some color | None -> Some item.Look.Color)
                            for (x, y, _, _) in points do drawMarker (if symbol = "auto" then "circle" else symbol) x y size fill stroke
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
                        let symbol = item.Source.MarkerSymbol |> Option.defaultValue (if plot.ScatterStyle.EndsWith "Marker" || plot.ScatterStyle = "marker" then "circle" else "none")
                        if symbol <> "none" then
                            let size = max 4.0 (item.Source.MarkerSize * 1.1)
                            let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                            for (x, y, _, _) in points do drawMarker (if symbol = "auto" then "circle" else symbol) x y size fill (Some item.Look.Color)
                    for (x, y, index, raw) in points do labelAt item index raw (x + 6.0) y "start" false
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
        let thinned (labels: string list) (spacing: float) (alongAxis: bool) (size: float) : int =
            let need = if alongAxis then widest labels size + 8.0 else size * 1.15
            if spacing <= 0.0 then 1 else max 1 (int (Math.Ceiling (need / spacing)))
        if shown primaryValueAxis then
            let every = thinned valueLabels ((if horizontal then plotWidth else plotHeight) / float (max 1 (valueLabels.Length - 1))) horizontal valueSize
            List.zip (ticksOf primaryScale) valueLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then
                    let at = along primaryScale primaryReversed value
                    if horizontal then add (PText (at, plotBottom + baseSize * 0.85, text, valueSize, valueColor, "middle", false, false, 0.0))
                    else add (PText (plotLeft - 6.0, at, text, valueSize, valueColor, "end", false, false, 0.0)))
        if secondaryAxis.IsSome && shown secondaryAxis then
            let every = thinned secondaryLabels (plotHeight / float (max 1 (secondaryLabels.Length - 1))) false valueSize
            List.zip (ticksOf secondaryScale) secondaryLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then add (PText (plotRight + 6.0, along secondaryScale secondaryReversed value, text, valueSize, valueColor, "start", false, false, 0.0)))
        if isScatter then
            if shown xAxis then
                let every = thinned xLabels (plotWidth / float (max 1 (xLabels.Length - 1))) true valueSize
                List.zip (ticksOf xScale) xLabels |> List.iteri (fun index (value, text) ->
                    if index % every = 0 then add (PText (xAt value, plotBottom + baseSize * 0.85, text, valueSize, valueColor, "middle", false, false, 0.0)))
        elif shown categoryAxis then
            // (As many labels as fit: every n-th one when they are crowded.)
            let widestLabel = widest categoryLabels categorySize
            if horizontal then
                let every = max 1 (int (Math.Ceiling ((categorySize * 1.15) / band)))
                // (The labels have the room left of the plot, whatever placed the plot.)
                let room = max 10.0 (plotLeft - left - 8.0)
                categories |> Array.iteri (fun index text ->
                    if index % every = 0 then
                        let text = if textWidth categorySize text > room && text.Length > 4 then text.Substring(0, max 1 (int (float text.Length * room / textWidth categorySize text) - 1)) + "…" else text
                        add (PText (plotLeft - 6.0, bandStart index + band / 2.0, text, categorySize, categoryColor, "end", false, false, 0.0)))
            else
                let every = max 1 (int (Math.Ceiling ((widestLabel + 6.0) / band)))
                categories |> Array.iteri (fun index text ->
                    if index % every = 0 then add (PText (pointAt index, plotBottom + baseSize * 0.85, text, categorySize, categoryColor, "middle", false, false, 0.0)))
        for label in pendingLabels do add label
    { Width = width; Height = height; Items = List.ofSeq items }

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
        match values.TryGetValue token with
        | true, value -> value
        | false, _ -> number token |> Option.defaultValue 0.0
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
        | FillAuto -> if isConnector then None else styleColor "fillRef"
    let stroke =
        match props |> Option.bind (strokeOf parts.Theme) with
        | Some { Paint = FillColor (color, _); Width = strokeWidth } -> Some (color, strokeWidth)
        | Some { Paint = FillNone } -> None
        | Some { Width = strokeWidth } -> styleColor "lnRef" |> Option.map (fun (color, _) -> color, strokeWidth)
        | None -> styleColor "lnRef" |> Option.map (fun (color, _) -> color, 1.0)
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
            // (Lines are broken where Excel breaks them: by the width the text has in Calibri.)
            let measured (style: TextStyle) (text: string) = calibriWidth (sizeOf style) (style.Bold |> Option.defaultValue false) text
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
                    yield (match align with
                           | "ctr" -> PTextRuns (x + width / 2.0, at + high / 2.0, "middle", spans)
                           | "r" -> PTextRuns (x + width - padRight, at + high / 2.0, "end", spans)
                           | _ -> PTextRuns (x + padLeft, at + high / 2.0, "start", spans))
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
    | "pic" ->
        (match descendants "blip" element |> List.tryHead |> Option.bind (attr "r:embed") |> Option.bind parts.Picture with
         | Some href -> [ PImage (x, y, width, height, href) ]
         | None -> unknown.Value <- unknown.Value + 1; [])
    | "graphicFrame" ->
        (match descendants "chart" element |> List.tryPick (attr "r:id") |> Option.bind parts.Chart with
         | Some chart ->
            charts.Value <- charts.Value + 1
            [ PGroup (x, y, width, height, (layoutChart parts.Theme data chart width height).Items) ]
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
        match content with
        | None -> None
        | Some element ->
            let items = elementItems parts data charts unknown element 0.0 0.0 width height
            // (A frame whose chart is not drawn, a group of nothing known: a placeholder says what it is.)
            if items.IsEmpty || (nameOf element = "graphicFrame" && charts.Value = 0) || (nameOf element = "grpSp" && charts.Value = 0 && unknown.Value > 0 && items.Length <= unknown.Value * 2) then None
            else Some { Width = width; Height = height; Items = items }
    with _ -> None

/// The cells and sheets the charts of an anchor read (`c:f`), by sheet
/// name: what a host watches to draw a chart again when they change.
let chartSheets (chart: Chart) : string list =
    chart.Plots
    |> List.collect (fun plot -> plot.Series)
    |> List.collect (fun series -> [ series.Name; series.Categories; series.Values; series.XValues ] |> List.choose id)
    |> List.choose (fun source -> source.Formula |> Option.bind parseAreas)
    |> List.collect (fun areas -> areas |> List.map (fun (sheet, _, _, _, _) -> sheet))
    |> List.distinct
