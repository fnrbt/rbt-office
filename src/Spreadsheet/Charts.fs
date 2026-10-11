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

/// How a shape is filled: nothing stated, no fill, or a colour (a gradient is its first stop).
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

let private fillOf (theme: ThemeColors) (holder: XmlDsl.XmlNode) : Fill =
    let colorIn (node: XmlDsl.XmlNode) = kids node |> List.tryPick (colorOfNode theme)
    match child "noFill" holder, child "solidFill" holder, child "gradFill" holder, child "pattFill" holder with
    | Some _, _, _, _ -> FillNone
    | _, Some solid, _, _ -> (match colorIn solid with Some (color, alpha) -> FillColor (color, alpha) | None -> FillAuto)
    | _, _, Some gradient, _ ->
        (match descendants "gs" gradient |> List.tryPick colorIn with
         | Some (color, alpha) -> FillColor (color, alpha)
         | None -> FillAuto)
    | _, _, _, Some pattern -> (match child "fgClr" pattern |> Option.bind colorIn with Some (color, alpha) -> FillColor (color, alpha) | None -> FillAuto)
    | _ -> FillAuto

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
}

type PlotKind =
    | BarPlot
    | LinePlot
    | PiePlot
    | DoughnutPlot
    | AreaPlot
    | ScatterPlot

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
}

type Axis = {
    Id: string
    /// "cat" | "val" | "date" | "ser"
    Kind: string
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
}

type Legend = {
    /// "r" | "l" | "t" | "b" | "tr"
    Position: string
    Style: TextStyle
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
                  Style = textPropsOf theme labels }
            if spec.ShowValue || spec.ShowPercent || spec.ShowCategory || spec.ShowSeries then Some spec else None)

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
      Fill = shape |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
      Line = shape |> Option.bind (strokeOf theme)
      MarkerSymbol = marker |> Option.bind (valOf "symbol")
      MarkerSize = marker |> Option.bind (valOf "size") |> Option.bind number |> Option.defaultValue 5.0
      MarkerFill = markerShape |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
      MarkerLine = markerShape |> Option.bind (strokeOf theme)
      Smooth = isOn (valOf "smooth" node) false
      PointFills =
        childrenNamed "dPt" node
        |> List.choose (fun point ->
            match valOf "idx" point |> Option.bind number, child "spPr" point |> Option.map (fillOf theme) with
            | Some index, Some fill when fill <> FillAuto -> Some (int index, fill)
            | _ -> None)
        |> Map.ofList
      Labels = labelsOf theme node }

let private plotOf (theme: ThemeColors) (node: XmlDsl.XmlNode) : Plot option =
    let kind =
        match nameOf node with
        | "barChart" -> Some BarPlot
        | "lineChart" -> Some LinePlot
        | "pieChart" -> Some PiePlot
        | "doughnutChart" -> Some DoughnutPlot
        | "areaChart" -> Some AreaPlot
        | "scatterChart" -> Some ScatterPlot
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
          ScatterStyle = valOf "scatterStyle" node |> Option.defaultValue "lineMarker" })

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
          CrossesId = valOf "crossAx" node })

/// The chart of a chart part, or None for a part this module does not draw
/// (no plot of a kind it knows, a 3-D chart, a part that is not a chart).
let parseChart (theme: ThemeColors) (xml: string) : Chart option =
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
                             |> Option.map (fun legend -> { Position = valOf "legendPos" legend |> Option.defaultValue "r"; Style = textPropsOf theme legend })
                           Background = shape |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
                           Border = shape |> Option.bind (strokeOf theme)
                           PlotBackground = child "spPr" area |> Option.map (fillOf theme) |> Option.defaultValue FillAuto
                           TextStyle = textPropsOf theme space
                           RoundedCorners = isOn (valOf "roundedCorners" space) false
                           Date1904 = isOn (valOf "date1904" space) false }
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

type Scene = {
    Width: float
    Height: float
    Items: Primitive list
}

let private fmt (value: float) : string =
    let rounded = Math.Round (value * 100.0) / 100.0
    if Double.IsNaN rounded || Double.IsInfinity rounded then "0" else rounded.ToString(Globalization.CultureInfo.InvariantCulture)

let private xmlText (text: string) : string =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

let private dashArray (dash: string option) (width: float) : string =
    match dash with
    | Some "dash" | Some "lgDash" | Some "sysDash" -> sprintf " stroke-dasharray=\"%s %s\"" (fmt (width * 4.0)) (fmt (width * 3.0))
    | Some "dot" | Some "sysDot" -> sprintf " stroke-dasharray=\"%s %s\"" (fmt width) (fmt (width * 2.0))
    | Some "dashDot" | Some "lgDashDot" | Some "sysDashDot" -> sprintf " stroke-dasharray=\"%s %s %s %s\"" (fmt (width * 4.0)) (fmt (width * 2.0)) (fmt width) (fmt (width * 2.0))
    | _ -> ""

let private paint (fill: string option) (opacity: float) (stroke: string option) (strokeWidth: float) : string =
    (match fill with Some color -> sprintf " fill=\"%s\"" color | None -> " fill=\"none\"")
    + (if opacity < 0.999 && fill.IsSome then sprintf " fill-opacity=\"%s\"" (fmt opacity) else "")
    + (match stroke with Some color when strokeWidth > 0.0 -> sprintf " stroke=\"%s\" stroke-width=\"%s\"" color (fmt strokeWidth) | _ -> "")

let mutable private clipSerial = 0

let rec private primitiveSvg (item: Primitive) : string =
    match item with
    | PRect (x, y, width, height, fill, opacity, stroke, strokeWidth, radius) ->
        sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\"%s%s/>" (fmt x) (fmt y) (fmt (max 0.0 width)) (fmt (max 0.0 height))
            (if radius > 0.0 then sprintf " rx=\"%s\"" (fmt radius) else "") (paint fill opacity stroke strokeWidth)
    | PEllipse (cx, cy, rx, ry, fill, opacity, stroke, strokeWidth) ->
        sprintf "<ellipse cx=\"%s\" cy=\"%s\" rx=\"%s\" ry=\"%s\"%s/>" (fmt cx) (fmt cy) (fmt (max 0.0 rx)) (fmt (max 0.0 ry)) (paint fill opacity stroke strokeWidth)
    | PLine (x1, y1, x2, y2, stroke, strokeWidth, dash) ->
        sprintf "<line x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\" stroke=\"%s\" stroke-width=\"%s\"%s/>" (fmt x1) (fmt y1) (fmt x2) (fmt y2) stroke (fmt strokeWidth) (dashArray dash strokeWidth)
    | PPolyline (points, fill, opacity, stroke, strokeWidth, dash) ->
        let text = points |> List.map (fun (x, y) -> fmt x + "," + fmt y) |> String.concat " "
        sprintf "<%s points=\"%s\"%s stroke-linejoin=\"round\" stroke-linecap=\"round\"%s/>" (if fill.IsSome then "polygon" else "polyline") text (paint fill opacity stroke strokeWidth) (dashArray dash strokeWidth)
    | PPath (d, fill, opacity, stroke, strokeWidth) ->
        sprintf "<path d=\"%s\"%s stroke-linejoin=\"round\"/>" d (paint fill opacity stroke strokeWidth)
    | PText (x, y, text, size, color, anchor, bold, italic, rotate) ->
        sprintf "<text x=\"%s\" y=\"%s\" font-size=\"%s\" fill=\"%s\" text-anchor=\"%s\" dominant-baseline=\"central\"%s%s%s>%s</text>"
            (fmt x) (fmt y) (fmt size) color anchor (if bold then " font-weight=\"bold\"" else "") (if italic then " font-style=\"italic\"" else "")
            (if rotate <> 0.0 then sprintf " transform=\"rotate(%s %s %s)\"" (fmt rotate) (fmt x) (fmt y) else "") (xmlText text)
    | PTextRuns (x, y, anchor, runs) ->
        let spans =
            runs
            |> List.map (fun (text, size, color, bold, italic) ->
                sprintf "<tspan font-size=\"%s\" fill=\"%s\"%s%s>%s</tspan>" (fmt size) color (if bold then " font-weight=\"bold\"" else "") (if italic then " font-style=\"italic\"" else "") (xmlText text))
            |> String.concat ""
        sprintf "<text x=\"%s\" y=\"%s\" text-anchor=\"%s\" dominant-baseline=\"central\" xml:space=\"preserve\" style=\"white-space:pre\">%s</text>" (fmt x) (fmt y) anchor spans
    | PImage (x, y, width, height, href) ->
        sprintf "<image x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" preserveAspectRatio=\"none\" href=\"%s\"/>" (fmt x) (fmt y) (fmt width) (fmt height) (xmlText href)
    | PGroup (x, y, width, height, items) ->
        clipSerial <- clipSerial + 1
        let id = sprintf "rbt-chart-clip-%d" clipSerial
        sprintf "<clipPath id=\"%s\"><rect x=\"0\" y=\"0\" width=\"%s\" height=\"%s\"/></clipPath><g transform=\"translate(%s %s)\" clip-path=\"url(#%s)\">%s</g>"
            id (fmt width) (fmt height) (fmt x) (fmt y) id (items |> List.map primitiveSvg |> String.concat "")

/// A scene as an `<svg>` element of its size (the font is Calibri, as Excel's charts).
let sceneSvg (scene: Scene) : string =
    sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%s\" height=\"%s\" viewBox=\"0 0 %s %s\" font-family=\"Calibri, Carlito, Arial, sans-serif\" style=\"display:block;overflow:hidden\">%s</svg>"
        (fmt scene.Width) (fmt scene.Height) (fmt scene.Width) (fmt scene.Height) (scene.Items |> List.map primitiveSvg |> String.concat "")

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
        else Some (blocks |> List.collect (fun block -> block.Value |> Array.toList |> List.collect Array.toList)))

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

/// The width of a text in pixels, near enough for a layout (Calibri averages about half its size per character).
let private textWidth (size: float) (text: string) : float =
    let mutable units = 0.0
    for ch in text do
        units <- units + (if ch = 'i' || ch = 'l' || ch = '.' || ch = ',' || ch = ' ' || ch = '\'' || ch = '|' || ch = ':' then 0.27
                          elif ch = 'm' || ch = 'w' || ch = 'M' || ch = 'W' || ch = '%' then 0.82
                          elif Char.IsUpper ch then 0.62
                          elif Char.IsDigit ch then 0.51
                          else 0.5)
    // (A little wide: where Calibri is not installed the browser's font is wider.)
    units * size * 1.08

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
    Look: SeriesLook
    FormatCode: string option
    /// On the secondary value axis.
    Secondary: bool
}

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
    let primaryValueAxis =
        // (The value axis of the first plot; on a scatter chart the one at the side.)
        let ofPlot (plot: Plot) = plot.AxisIds |> List.choose axisById |> List.filter (fun axis -> axis.Kind = "val")
        match chart.Plots with
        | first :: _ ->
            (match ofPlot first with
             | [ only ] -> Some only
             | several when first.Kind = ScatterPlot -> several |> List.tryFind (fun axis -> axis.Position = "l" || axis.Position = "r") |> Option.orElse (List.tryLast several)
             | several -> List.tryHead several)
        | [] -> None
    let mutable ordinal = 0
    let series : SeriesData list =
        chart.Plots
        |> List.collect (fun plot ->
            let plotValueAxis = plot.AxisIds |> List.choose axisById |> List.filter (fun axis -> axis.Kind = "val" && (plot.Kind <> ScatterPlot || axis.Position = "l" || axis.Position = "r")) |> List.tryHead
            let secondary =
                match plotValueAxis, primaryValueAxis with
                | Some mine, Some first -> mine.Id <> first.Id
                | _ -> false
            plot.Series
            |> List.map (fun source ->
                let at = ordinal
                ordinal <- ordinal + 1
                let auto = autoColor theme at
                let lineOnly = plot.Kind = LinePlot || plot.Kind = ScatterPlot
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
                            | FillAuto -> auto, 1.0, true
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
                  Look = look
                  FormatCode = formatCodeOf data source.Values
                  Secondary = secondary }))
    let firstPlot = chart.Plots |> List.tryHead
    let isPie = match firstPlot with Some plot -> plot.Kind = PiePlot || plot.Kind = DoughnutPlot | None -> false
    let isScatter = match firstPlot with Some plot -> plot.Kind = ScatterPlot | None -> false
    let categories : string [] =
        let stated = series |> List.tryPick (fun item -> match textsOf data item.Source.Categories with [||] -> None | texts -> Some texts)
        let count = series |> List.fold (fun most item -> max most item.Values.Length) 0
        match stated with
        | Some texts -> Array.init (max count texts.Length) (fun index -> if index < texts.Length then texts.[index] else "")
        | None -> Array.init count (fun index -> string (index + 1))
    let pointColor (item: SeriesData) (index: int) : string * float =
        match Map.tryFind index item.Source.PointFills with
        | Some (FillColor (color, opacity)) -> color, opacity
        | _ -> if item.Plot.VaryColors then autoColor theme index, 1.0 else item.Look.Color, item.Look.Opacity

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
        let lines = text.Split('\n')
        lines |> Array.iteri (fun index line ->
            add (PText (width / 2.0, top + size * 0.7 + float index * size * 1.25, line, size, textColor chart.TitleStyle, "middle",
                        (chart.TitleStyle.Bold |> Option.defaultValue false), (chart.TitleStyle.Italic |> Option.defaultValue false), 0.0)))
        top <- top + float lines.Length * size * 1.25 + 8.0
     | None -> ())

    // ---- legend ----
    let legendEntries : (string * string * bool) list =
        // (text, colour, a line rather than a box)
        if isPie then
            match series with
            | first :: _ -> categories |> Array.toList |> List.mapi (fun index text -> text, fst (pointColor first index), false)
            | [] -> []
        else series |> List.map (fun item -> item.Label, (match item.Look.LineColor with Some color when not item.Look.Filled -> color | _ -> item.Look.Color), (item.Plot.Kind = LinePlot || item.Plot.Kind = ScatterPlot))
    let mutable left = 8.0
    let mutable right = width - 8.0
    let mutable bottom = height - 8.0
    (match chart.Legend with
     | Some legend when not legendEntries.IsEmpty ->
        let size = textSize legend.Style 9.0
        let color = textColor legend.Style
        let swatch = size * 0.75
        let entryWidth (text: string) = swatch + 5.0 + textWidth size text + 14.0
        let drawEntry (x: float) (y: float) (text: string, entryColor: string, line: bool) =
            if line then add (PLine (x, y, x + swatch * 1.6, y, entryColor, 2.25, None))
            else add (PRect (x, y - swatch / 2.0, swatch, swatch, Some entryColor, 1.0, None, 0.0, 0.0))
            add (PText (x + (if line then swatch * 1.6 else swatch) + 5.0, y, text, size, color, "start", false, false, 0.0))
        match legend.Position with
        | "b" | "t" ->
            // Rows of entries, centred.
            let rowHeight = size * 1.5
            let rows = ResizeArray<ResizeArray<string * string * bool>>()
            let mutable current = ResizeArray<string * string * bool>()
            let mutable used = 0.0
            for entry in legendEntries do
                let (text, _, line) = entry
                let entryW = entryWidth text + (if line then swatch * 0.6 else 0.0)
                if used + entryW > width - 16.0 && current.Count > 0 then
                    rows.Add current
                    current <- ResizeArray<string * string * bool>()
                    used <- 0.0
                current.Add entry
                used <- used + entryW
            if current.Count > 0 then rows.Add current
            let rows = rows |> Seq.truncate 4 |> Seq.toList
            let blockHeight = float rows.Length * rowHeight
            let startY = if legend.Position = "b" then bottom - blockHeight else top
            rows |> List.iteri (fun rowIndex row ->
                let total = row |> Seq.sumBy (fun (text, _, line) -> entryWidth text + (if line then swatch * 0.6 else 0.0))
                let mutable x = max 8.0 ((width - total) / 2.0 + 7.0)
                for entry in row do
                    let (text, _, line) = entry
                    drawEntry x (startY + rowHeight * (float rowIndex + 0.5)) entry
                    x <- x + entryWidth text + (if line then swatch * 0.6 else 0.0))
            if legend.Position = "b" then bottom <- bottom - blockHeight - 4.0 else top <- top + blockHeight + 4.0
        | position ->
            let rowHeight = size * 1.6
            let widest = legendEntries |> List.fold (fun most (text, _, line) -> max most (entryWidth text + (if line then swatch * 0.6 else 0.0))) 0.0
            let blockWidth = min (width * 0.4) widest
            let shown = legendEntries |> List.truncate (max 1 (int ((bottom - top) / rowHeight)))
            let blockHeight = float shown.Length * rowHeight
            let x = if position = "l" then left else right - blockWidth + 4.0
            let startY = if position = "tr" then top else top + max 0.0 ((bottom - top - blockHeight) / 2.0)
            shown |> List.iteri (fun index entry -> drawEntry x (startY + rowHeight * (float index + 0.5)) entry)
            if position = "l" then left <- left + blockWidth + 6.0 else right <- right - blockWidth - 2.0
     | _ -> ())

    let labelStyleOf (item: SeriesData) = item.Source.Labels |> Option.orElse item.Plot.Labels
    let labelText (item: SeriesData) (spec: LabelSpec) (index: int) (value: float) (share: float option) : string =
        [ if spec.ShowSeries then yield item.Label
          if spec.ShowCategory && index < categories.Length then yield categories.[index]
          if spec.ShowValue then
              yield (match spec.FormatCode |> Option.orElse item.FormatCode with
                     | Some code -> data.Format code value
                     | None -> generalNumber value)
          if spec.ShowPercent then
              match share with
              | Some part -> yield generalNumber (Math.Round (part * 100.0)) + "%"
              | None -> () ]
        |> String.concat "; "

    if isPie then
        // ---- pie, doughnut ----
        let rings = series |> List.filter (fun item -> item.Plot.Kind = PiePlot || item.Plot.Kind = DoughnutPlot)
        let rings = (match firstPlot with Some plot when plot.Kind = PiePlot -> List.truncate 1 rings | _ -> rings)
        let cx, cy = (left + right) / 2.0, (top + bottom) / 2.0
        let radius = max 4.0 (min (right - left) (bottom - top) / 2.0 - 4.0)
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
                        (match labelStyleOf item with
                         | Some spec when sweep > 0.12 ->
                            let size = textSize spec.Style 9.0
                            let middle = angle + sweep / 2.0
                            let at = (if inner > 0.0 then (inner + outer) / 2.0 elif spec.Position = Some "outEnd" then outer + size else outer * 0.68)
                            let x, y = point at middle
                            add (PText (x, y, labelText item spec index value (Some (value / total)), size, (spec.Style.Color |> Option.defaultValue (if spec.Position = Some "outEnd" then "#404040" else "#FFFFFF")), "middle", (spec.Style.Bold |> Option.defaultValue false), false, 0.0))
                         | _ -> ())
                        angle <- angle + sweep))
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
        let extent (secondary: bool) : (float * float) option =
            let mutable lo = Double.PositiveInfinity
            let mutable hi = Double.NegativeInfinity
            let see (value: float) = (if value < lo then lo <- value); (if value > hi then hi <- value)
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
        let scaleFor (axis: Axis option) (secondary: bool) : float * float * float =
            let percent = chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            match extent secondary with
            | None -> 0.0, 1.0, 0.2
            | Some (lo, hi) ->
                if percent then (if lo < 0.0 then -1.0 else 0.0), 1.0, 0.1
                else valueScale lo hi (axis |> Option.bind (fun a -> a.Min)) (axis |> Option.bind (fun a -> a.Max)) (axis |> Option.bind (fun a -> a.MajorUnit))
        let vMin, vMax, vStep = scaleFor primaryValueAxis false
        let sMin, sMax, sStep = scaleFor secondaryAxis true
        let xMin, xMax, xStep =
            if not isScatter then 0.0, 1.0, 1.0
            else
                let values = series |> List.collect (fun item -> item.XValues |> Array.choose id |> Array.toList)
                match values with
                | [] -> 0.0, float (max 1 count), 1.0
                | _ -> valueScale (List.min values) (List.max values) (xAxis |> Option.bind (fun a -> a.Min)) (xAxis |> Option.bind (fun a -> a.Max)) (xAxis |> Option.bind (fun a -> a.MajorUnit))
        let ticks (lo: float) (hi: float) (step: float) : float list =
            [ let mutable value = lo
              let mutable guard = 0
              while value <= hi + step * 1e-6 && guard < 60 do
                  yield (if abs value < step * 1e-9 then 0.0 else value)
                  value <- value + step
                  guard <- guard + 1 ]
        let axisFormat (axis: Axis option) (secondary: bool) : float -> string =
            let percent = chart.Plots |> List.exists (fun plot -> plot.Grouping = "percentStacked" && series |> List.exists (fun item -> obj.ReferenceEquals(item.Plot, plot) && item.Secondary = secondary))
            let linked = series |> List.tryPick (fun item -> if item.Secondary = secondary then item.FormatCode else None)
            let code =
                match axis with
                | Some axis when not axis.SourceLinked -> axis.FormatCode |> Option.filter (fun code -> code <> "General")
                | Some axis -> linked |> Option.orElse (axis.FormatCode |> Option.filter (fun code -> code <> "General"))
                | None -> linked
            match code with
            | _ when percent -> (fun value -> generalNumber (Math.Round (value * 100.0)) + "%")
            | Some code -> (fun value -> try data.Format code value with _ -> generalNumber value)
            | None -> generalNumber
        let valueLabels = ticks vMin vMax vStep |> List.map (axisFormat primaryValueAxis false)
        let secondaryLabels = if secondaryAxis.IsSome then ticks sMin sMax sStep |> List.map (axisFormat secondaryAxis true) else []
        let shown (axis: Axis option) = match axis with Some axis -> not axis.Deleted && axis.LabelsShown | None -> false
        let valueSize = textSize (primaryValueAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let categorySize = textSize (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle) 9.0
        let valueColor = textColor (primaryValueAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle)
        let categoryColor = textColor (categoryAxis |> Option.map (fun axis -> axis.LabelStyle) |> Option.defaultValue noTextStyle)
        let widest (labels: string list) (size: float) = labels |> List.fold (fun most text -> max most (textWidth size text)) 0.0
        // Axis titles.
        let axisTitle (axis: Axis option) = axis |> Option.bind (fun axis -> if axis.Deleted then None else axis.Title |> Option.map (fun text -> text, textSize axis.TitleStyle 10.0, textColor axis.TitleStyle))
        let sideAxis, baseAxis = (if horizontal then categoryAxis, primaryValueAxis elif isScatter then primaryValueAxis, xAxis else primaryValueAxis, categoryAxis)
        (match axisTitle sideAxis with
         | Some (text, size, color) ->
            add (PText (left + size * 0.6, (top + bottom) / 2.0, text, size, color, "middle", false, false, -90.0))
            left <- left + size * 1.5
         | None -> ())
        (match axisTitle baseAxis with
         | Some (text, size, color) ->
            add (PText ((left + right) / 2.0, bottom - size * 0.6, text, size, color, "middle", false, false, 0.0))
            bottom <- bottom - size * 1.5
         | None -> ())
        // The room the tick labels take.
        let categoryLabels = categories |> Array.toList
        let xLabels = if isScatter then ticks xMin xMax xStep |> List.map (axisFormat xAxis false) else []
        let sideLabels, baseLabels = (if horizontal then categoryLabels, valueLabels elif isScatter then valueLabels, xLabels else valueLabels, categoryLabels)
        let sideSize, baseSize = (if horizontal then categorySize, valueSize else valueSize, categorySize)
        let sideShown, baseShown = (if horizontal then shown categoryAxis, shown primaryValueAxis elif isScatter then shown primaryValueAxis, shown xAxis else shown primaryValueAxis, shown categoryAxis)
        let sideRoom = if sideShown then min (width * 0.4) (widest sideLabels sideSize) + 8.0 else 4.0
        let baseRoom = if baseShown then baseSize * 1.25 + 6.0 else 4.0
        let secondaryRoom = if secondaryAxis.IsSome && shown secondaryAxis then widest secondaryLabels valueSize + 8.0 else 0.0
        let plotLeft = left + sideRoom
        let plotRight = max (plotLeft + 10.0) (right - secondaryRoom - (if baseShown && not horizontal then 0.0 else 6.0) - 4.0)
        let plotTop = top + (if title.IsNone then 6.0 else 0.0)
        let plotBottom = max (plotTop + 10.0) (bottom - baseRoom)
        let plotWidth, plotHeight = plotRight - plotLeft, plotBottom - plotTop
        (match chart.PlotBackground with
         | FillColor (color, opacity) -> add (PRect (plotLeft, plotTop, plotWidth, plotHeight, Some color, opacity, None, 0.0, 0.0))
         | _ -> ())
        // Where a value is on the value axis (along the plot's height, or its width for bars that run to the right).
        let valueReversed = primaryValueAxis |> Option.map (fun axis -> axis.Reversed) |> Option.defaultValue false
        let along (lo: float) (hi: float) (value: float) : float =
            let share = if hi > lo then (value - lo) / (hi - lo) else 0.0
            let share = if valueReversed then 1.0 - share else share
            if horizontal then plotLeft + share * plotWidth else plotBottom - share * plotHeight
        let valueAt (secondary: bool) (value: float) = if secondary then along sMin sMax value else along vMin vMax value
        // Where a category's band begins, and its width (bars that run to the right list their categories upwards).
        let categoryReversed = categoryAxis |> Option.map (fun axis -> axis.Reversed) |> Option.defaultValue false
        let band = (if horizontal then plotHeight else plotWidth) / float count
        let bandStart (index: int) : float =
            if horizontal then (if categoryReversed then plotTop + float index * band else plotBottom - float (index + 1) * band)
            else (if categoryReversed then plotRight - float (index + 1) * band else plotLeft + float index * band)
        // ---- grid lines ----
        let gridOf (axis: Axis option) = axis |> Option.bind (fun axis -> strokeColor axis.Gridlines)
        (match gridOf primaryValueAxis with
         | Some (color, strokeWidth, dash) ->
            for value in ticks vMin vMax vStep do
                let at = along vMin vMax value
                if horizontal then add (PLine (at, plotTop, at, plotBottom, color, strokeWidth, dash))
                else add (PLine (plotLeft, at, plotRight, at, color, strokeWidth, dash))
         | None -> ())
        (match (if isScatter then gridOf xAxis else None) with
         | Some (color, strokeWidth, dash) ->
            for value in ticks xMin xMax xStep do
                let at = plotLeft + (value - xMin) / (xMax - xMin) * plotWidth
                add (PLine (at, plotTop, at, plotBottom, color, strokeWidth, dash))
         | None -> ())
        // ---- the series, plot by plot (areas and bars under lines) ----
        let pendingLabels = ResizeArray<Primitive>()
        let labelAt (item: SeriesData) (index: int) (value: float) (x: float) (y: float) (anchor: string) (inside: bool) =
            match labelStyleOf item with
            | Some spec ->
                let size = textSize spec.Style 9.0
                pendingLabels.Add (PText (x, y, labelText item spec index value None, size, (spec.Style.Color |> Option.defaultValue (if inside then "#FFFFFF" else "#404040")), anchor, (spec.Style.Bold |> Option.defaultValue false), false, 0.0))
            | None -> ()
        let order (plot: Plot) = match plot.Kind with AreaPlot -> 0 | BarPlot -> 1 | LinePlot -> 2 | ScatterPlot -> 3 | _ -> 4
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
                                else 0.0, value
                            let lo, hi = (if item.Secondary then sMin, sMax else vMin, vMax)
                            let clamp (v: float) = max lo (min hi v)
                            let a, b = valueAt item.Secondary (clamp from), valueAt item.Secondary (clamp upto)
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
            | LinePlot | AreaPlot ->
                let sums = Array.zeroCreate<float> count
                for item in members do
                    let points = ResizeArray<float * float * int * float>()
                    let floors = ResizeArray<float * float>()
                    let runs = ResizeArray<(float * float) list>()
                    let mutable run : (float * float) list = []
                    for index in 0 .. count - 1 do
                        match (if index < item.Values.Length then item.Values.[index] else None) with
                        | Some raw ->
                            let value = scaled index raw
                            let basis = if stacked then sums.[index] else 0.0
                            let total = basis + value
                            if stacked then sums.[index] <- total
                            let lo, hi = (if item.Secondary then sMin, sMax else vMin, vMax)
                            let x = bandStart index + band / 2.0
                            let y = valueAt item.Secondary (max lo (min hi total))
                            points.Add ((x, y, index, raw))
                            floors.Add ((x, valueAt item.Secondary (max lo (min hi basis))))
                            run <- (x, y) :: run
                        | None ->
                            if not run.IsEmpty then runs.Add (List.rev run)
                            run <- []
                    if not run.IsEmpty then runs.Add (List.rev run)
                    if plot.Kind = AreaPlot && points.Count > 0 then
                        let upper = points |> Seq.map (fun (x, y, _, _) -> x, y) |> Seq.toList
                        let lower = floors |> Seq.toList |> List.rev
                        add (PPolyline (upper @ lower, (if item.Look.Filled then Some item.Look.Color else None), item.Look.Opacity, item.Look.LineColor, item.Look.LineWidth, None))
                    elif plot.Kind = LinePlot then
                        (match item.Look.LineColor with
                         | Some color -> for line in runs do (if line.Length > 1 then add (PPolyline (line, None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash)))
                         | None -> ())
                        // Markers: the chart's switch, unless the series says none.
                        let symbol = item.Source.MarkerSymbol |> Option.defaultValue (if plot.Markers then "circle" else "none")
                        if symbol <> "none" then
                            let size = max 3.0 (item.Source.MarkerSize * 1.1)
                            let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                            let stroke = (match strokeColor item.Source.MarkerLine with Some (color, _, _) -> Some color | None -> Some item.Look.Color)
                            for (x, y, _, _) in points do
                                match symbol with
                                | "square" -> add (PRect (x - size / 2.0, y - size / 2.0, size, size, fill, 1.0, stroke, 0.75, 0.0))
                                | "diamond" -> add (PPolyline ([ x, y - size / 1.6; x + size / 1.6, y; x, y + size / 1.6; x - size / 1.6, y ], (fill |> Option.orElse (Some "#FFFFFF")), 1.0, stroke, 0.75, None))
                                | "triangle" -> add (PPolyline ([ x, y - size / 1.6; x + size / 1.6, y + size / 2.0; x - size / 1.6, y + size / 2.0 ], (fill |> Option.orElse (Some "#FFFFFF")), 1.0, stroke, 0.75, None))
                                | _ -> add (PEllipse (x, y, size / 2.0, size / 2.0, fill, 1.0, stroke, 0.75))
                    for (x, y, index, raw) in points do
                        let position = labelStyleOf item |> Option.bind (fun spec -> spec.Position) |> Option.defaultValue (if plot.Kind = AreaPlot then "ctr" else "r")
                        (match position with
                         | "t" -> labelAt item index raw x (y - 9.0) "middle" false
                         | "b" -> labelAt item index raw x (y + 9.0) "middle" false
                         | "l" -> labelAt item index raw (x - 6.0) y "end" false
                         | "ctr" -> labelAt item index raw x y "middle" false
                         | _ -> labelAt item index raw (x + 6.0) y "start" false)
            | ScatterPlot ->
                for item in members do
                    let points =
                        [ for index in 0 .. item.Values.Length - 1 do
                            let x = (if index < item.XValues.Length then item.XValues.[index] else Some (float (index + 1)))
                            match x, item.Values.[index] with
                            | Some x, Some y when xMax > xMin ->
                                let lo, hi = (if item.Secondary then sMin, sMax else vMin, vMax)
                                yield plotLeft + (max xMin (min xMax x) - xMin) / (xMax - xMin) * plotWidth, valueAt item.Secondary (max lo (min hi y)), index, y
                            | _ -> () ]
                    let lined = plot.ScatterStyle.StartsWith "line" || plot.ScatterStyle.StartsWith "smooth"
                    (match item.Look.LineColor with
                     | Some color when lined && points.Length > 1 && (item.Source.Line.IsSome || plot.ScatterStyle <> "lineMarker" || item.Source.MarkerSymbol = Some "none") ->
                        add (PPolyline (points |> List.map (fun (x, y, _, _) -> x, y), None, 1.0, Some color, item.Look.LineWidth, item.Look.LineDash))
                     | _ -> ())
                    let symbol = item.Source.MarkerSymbol |> Option.defaultValue (if plot.ScatterStyle.EndsWith "Marker" || plot.ScatterStyle = "marker" then "circle" else "none")
                    if symbol <> "none" then
                        let size = max 4.0 (item.Source.MarkerSize * 1.1)
                        let fill = (match item.Source.MarkerFill with FillColor (color, _) -> Some color | FillNone -> None | FillAuto -> Some item.Look.Color)
                        for (x, y, _, _) in points do add (PEllipse (x, y, size / 2.0, size / 2.0, fill, 1.0, Some item.Look.Color, 0.75))
                    for (x, y, index, raw) in points do labelAt item index raw (x + 6.0) y "start" false
            | _ -> ()
        // ---- axes: lines and labels ----
        let axisLine (axis: Axis option) = axis |> Option.bind (fun axis -> if axis.Deleted then None else strokeColor axis.Line)
        let zeroAt = along vMin vMax (max vMin (min vMax 0.0))
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
            List.zip (ticks vMin vMax vStep) valueLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then
                    let at = along vMin vMax value
                    if horizontal then add (PText (at, plotBottom + baseSize * 0.85, text, valueSize, valueColor, "middle", false, false, 0.0))
                    else add (PText (plotLeft - 6.0, at, text, valueSize, valueColor, "end", false, false, 0.0)))
        if secondaryAxis.IsSome && shown secondaryAxis then
            let every = thinned secondaryLabels (plotHeight / float (max 1 (secondaryLabels.Length - 1))) false valueSize
            List.zip (ticks sMin sMax sStep) secondaryLabels |> List.iteri (fun index (value, text) ->
                if index % every = 0 then add (PText (plotRight + 6.0, along sMin sMax value, text, valueSize, valueColor, "start", false, false, 0.0)))
        if isScatter then
            if shown xAxis then
                let every = thinned xLabels (plotWidth / float (max 1 (xLabels.Length - 1))) true valueSize
                List.zip (ticks xMin xMax xStep) xLabels |> List.iteri (fun index (value, text) ->
                    if index % every = 0 then add (PText (plotLeft + (value - xMin) / (xMax - xMin) * plotWidth, plotBottom + baseSize * 0.85, text, valueSize, valueColor, "middle", false, false, 0.0)))
        elif shown categoryAxis then
            // (As many labels as fit: every n-th one when they are crowded.)
            let widestLabel = widest categoryLabels categorySize
            if horizontal then
                let every = max 1 (int (Math.Ceiling ((categorySize * 1.15) / band)))
                categories |> Array.iteri (fun index text ->
                    if index % every = 0 then
                        let room = sideRoom - 8.0
                        let text = if textWidth categorySize text > room && text.Length > 4 then text.Substring(0, max 1 (int (float text.Length * room / textWidth categorySize text) - 1)) + "…" else text
                        add (PText (plotLeft - 6.0, bandStart index + band / 2.0, text, categorySize, categoryColor, "end", false, false, 0.0)))
            else
                let every = max 1 (int (Math.Ceiling ((widestLabel + 6.0) / band)))
                categories |> Array.iteri (fun index text ->
                    if index % every = 0 then add (PText (bandStart index + band / 2.0, plotBottom + baseSize * 0.85, text, categorySize, categoryColor, "middle", false, false, 0.0)))
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

let private shapeItems (parts: DrawingParts) (shape: XmlDsl.XmlNode) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    let props = child "spPr" shape
    let style = child "style" shape
    let styleColor (name: string) = style |> Option.bind (child name) |> Option.bind (fun reference -> kids reference |> List.tryPick (colorOfNode parts.Theme))
    let isConnector = nameOf shape = "cxnSp"
    let geometry = props |> Option.bind (child "prstGeom") |> Option.bind (attr "prst") |> Option.defaultValue (if isConnector then "line" else "rect")
    let fill =
        match props |> Option.map (fillOf parts.Theme) |> Option.defaultValue FillAuto with
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
            if fillColor.IsNone && strokeColor.IsNone then [] else [ PRect (x, y, width, height, fillColor, opacity, strokeColor, strokeWidth, 0.0) ]
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
                                Some (child "t" run |> Option.map innerText |> Option.defaultValue "", style)
                            | _ -> None)
                    let endStyle = child "endParaRPr" paragraph |> Option.map (runStyle parts.Theme) |> Option.defaultValue noTextStyle
                    align, runs, over endStyle defaults)
            let sizeOf (style: TextStyle) = px (style.Size |> Option.defaultValue 11.0)
            let wraps = (bodyProps |> Option.bind (attr "wrap")) <> Some "none"
            // (Lines are broken where Excel breaks them: by the width the text has in Calibri.)
            let measured (style: TextStyle) (text: string) = textWidth (sizeOf style) text / 1.08
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
    outline @ text

/// What an element of a drawing draws in a rectangle (pixels): a shape, a
/// connector, a picture, a frame with a chart, a group of these. `charts`
/// counts the charts that were drawn, `unknown` what could not be.
let rec private elementItems (parts: DrawingParts) (data: ChartData) (charts: int ref) (unknown: int ref) (element: XmlDsl.XmlNode) (x: float) (y: float) (width: float) (height: float) : Primitive list =
    match nameOf element with
    | "sp" | "cxnSp" -> shapeItems parts element x y width height
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
