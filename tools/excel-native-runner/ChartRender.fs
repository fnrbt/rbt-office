/// Minimal Excel chart → PaintItem renderer (prototype). Parses a workbook's
/// first chart (xl/charts/chart1.xml) and its drawing anchor, and lays out a
/// clustered column chart the way Excel does, emitting the vector PaintItems
/// (VecPath/VecGroup/DrawText) added to the shared paint model.
module ChartRender

open System
open System.IO
open System.IO.Compression
open System.Xml.Linq
open RedBat.Spreadsheet
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Layout
open RedBat.Spreadsheet.Paint

// ---- parsing -------------------------------------------------------------

type ChartSpec =
    { Kind: string                 // "col" (barDir)
      Title: string option
      Categories: string list
      Values: float list
      Gridlines: bool
      GapWidth: int
      SeriesColor: string          // "#RRGGBB" (= accent1)
      Palette: string list }       // theme accent1..6, for pie slices etc.

let private localName (e: XElement) = e.Name.LocalName
let private descendants (name: string) (e: XElement) =
    e.Descendants() |> Seq.filter (fun d -> d.Name.LocalName = name)
let private firstDesc name e = descendants name e |> Seq.tryHead

let private readEntry (zip: ZipArchive) (path: string) : string option =
    match zip.GetEntry path with
    | null -> None
    | entry ->
        use s = entry.Open()
        use r = new StreamReader(s)
        Some (r.ReadToEnd())

/// Pull the ordered <c:pt><c:v> text values out of a <c:cat>/<c:val> block.
let private points (block: XElement) : string list =
    descendants "pt" block
    |> Seq.sortBy (fun p -> match p.Attribute(XName.Get "idx") with null -> 0 | a -> int a.Value)
    |> Seq.choose (fun p -> firstDesc "v" p |> Option.map (fun v -> v.Value))
    |> List.ofSeq

let private pf (s: string) = match Double.TryParse(s, Globalization.CultureInfo.InvariantCulture) with true, n -> n | _ -> 0.0

let private defaultAccents = [ "#5B9BD5"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#4472C4"; "#70AD47" ]

/// Read the theme accent1..6 colours (the chart series palette).
let private parseTheme (zip: ZipArchive) : string list =
    match readEntry zip "xl/theme/theme1.xml" with
    | None -> defaultAccents
    | Some xml ->
        let root = (XDocument.Parse xml).Root
        [ for i in 1 .. 6 ->
            match firstDesc (sprintf "accent%d" i) root |> Option.bind (firstDesc "srgbClr") with
            | Some c -> "#" + c.Attribute(XName.Get "val").Value
            | None -> List.item (i - 1) defaultAccents ]

let parseChart (xlsxPath: string) (chartIdx: int) : ChartSpec option =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip (sprintf "xl/charts/chart%d.xml" chartIdx) with
    | None -> None
    | Some xml ->
        let doc = XDocument.Parse xml
        let root = doc.Root
        let tryType name k = firstDesc name root |> Option.map (fun p -> p, k)
        let plotKind =
            (firstDesc "barChart" root
             |> Option.map (fun bar ->
                 bar, (firstDesc "barDir" bar |> Option.map (fun e -> e.Attribute(XName.Get "val").Value) |> Option.defaultValue "col")))
            |> Option.orElseWith (fun () -> tryType "lineChart" "line")
            |> Option.orElseWith (fun () -> tryType "pieChart" "pie")
            |> Option.orElseWith (fun () -> tryType "areaChart" "area")
            |> Option.orElseWith (fun () -> tryType "scatterChart" "scatter")
        match plotKind with
        | None -> None
        | Some (plot, kind) ->
            let ser = firstDesc "ser" plot
            let cats, vals =
                if kind = "scatter" then
                    let xs = ser |> Option.bind (firstDesc "xVal") |> Option.map points |> Option.defaultValue []
                    let ys = ser |> Option.bind (firstDesc "yVal") |> Option.map points |> Option.defaultValue []
                    xs, (ys |> List.map pf)
                else
                    let c = ser |> Option.bind (firstDesc "cat") |> Option.map points |> Option.defaultValue []
                    let v = ser |> Option.bind (firstDesc "val") |> Option.map points |> Option.defaultValue [] |> List.map pf
                    c, v
            let title =
                firstDesc "title" root
                |> Option.bind (fun t -> descendants "t" t |> Seq.tryHead)
                |> Option.map (fun t -> t.Value)
            let gridlines = root |> descendants "valAx" |> Seq.exists (fun ax -> firstDesc "majorGridlines" ax |> Option.isSome)
            let gap = firstDesc "gapWidth" plot |> Option.map (fun e -> int (e.Attribute(XName.Get "val").Value)) |> Option.defaultValue 150
            let accents = parseTheme zip
            Some { Kind = kind; Title = title; Categories = cats; Values = vals
                   Gridlines = gridlines; GapWidth = gap
                   SeriesColor = List.head accents; Palette = accents }

/// Chart pixel size from the twoCellAnchor (default 64px cols / 20px rows;
/// EMU offsets at 9525 EMU/px). Good enough for the synthetic fixture.
let parseAnchorSize (xlsxPath: string) (chartIdx: int) : float * float =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip "xl/drawings/drawing1.xml" with
    | None -> 480.0, 320.0
    | Some xml ->
        let doc = XDocument.Parse xml
        // Pick the Nth twoCellAnchor (document order). Zoo charts are all the
        // same size, so exact anchor↔chart mapping isn't needed here.
        let anchors = descendants "twoCellAnchor" doc.Root |> Seq.toList
        let anchor = anchors |> List.tryItem (chartIdx - 1) |> Option.defaultValue (List.head anchors)
        let pick tag child =
            firstDesc tag anchor
            |> Option.bind (fun e -> firstDesc child e)
            |> Option.map (fun e -> float e.Value)
            |> Option.defaultValue 0.0
        let colPx c = c * 64.0
        let rowPx r = r * 20.0
        let fromX = colPx (pick "from" "col") + pick "from" "colOff" / 9525.0
        let fromY = rowPx (pick "from" "row") + pick "from" "rowOff" / 9525.0
        let toX = colPx (pick "to" "col") + pick "to" "colOff" / 9525.0
        let toY = rowPx (pick "to" "row") + pick "to" "rowOff" / 9525.0
        (toX - fromX), (toY - fromY)

/// A parsed chart together with its drawing-anchor placement (cell + EMU
/// offset, and pixel size). Used to overlay charts onto the sheet.
type ChartPlacement =
    { Spec: ChartSpec
      FromCol: int; FromColOffPx: float
      FromRow: int; FromRowOffPx: float
      WidthPx: float; HeightPx: float }

/// Parse every chart in the workbook with its anchor. Assumes anchor document
/// order matches chart number (true for charts added in sequence).
let parseAllCharts (xlsxPath: string) : ChartPlacement list =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip "xl/drawings/drawing1.xml" with
    | None -> []
    | Some xml ->
        let anchors = descendants "twoCellAnchor" (XDocument.Parse xml).Root |> Seq.toList
        anchors
        |> List.mapi (fun i anchor ->
            let getf tag child =
                firstDesc tag anchor |> Option.bind (firstDesc child) |> Option.map (fun e -> float e.Value) |> Option.defaultValue 0.0
            let fromCol, fromColOff = getf "from" "col", getf "from" "colOff" / 9525.0
            let fromRow, fromRowOff = getf "from" "row", getf "from" "rowOff" / 9525.0
            let toCol, toColOff = getf "to" "col", getf "to" "colOff" / 9525.0
            let toRow, toRowOff = getf "to" "row", getf "to" "rowOff" / 9525.0
            let wPx = (toCol - fromCol) * 64.0 + toColOff - fromColOff
            let hPx = (toRow - fromRow) * 20.0 + toRowOff - fromRowOff
            parseChart xlsxPath (i + 1)
            |> Option.map (fun spec ->
                { Spec = spec; FromCol = int fromCol; FromColOffPx = fromColOff
                  FromRow = int fromRow; FromRowOffPx = fromRowOff; WidthPx = wPx; HeightPx = hPx }))
        |> List.choose id

// ---- layout --------------------------------------------------------------

let private niceStep (dataMax: float) : float =
    let rough = (max dataMax 1.0) / 5.0
    let mag = 10.0 ** floor (log10 rough)
    let norm = rough / mag
    let nice = if norm < 1.5 then 1.0 elif norm < 3.0 then 2.0 elif norm < 7.0 then 5.0 else 10.0
    nice * mag

let private niceMax (dataMax: float) (step: float) : float =
    let m = ceil (dataMax / step) * step
    if m <= dataMax + 1e-9 then m + step else m

let private text (s: string) (rect: Rect) (size: float) (color: string) (align: TextAlign) (vAlign: VerticalAlign) : PaintItem =
    DrawText { Text = s; Rect = rect; Color = color; FontSize = size; FontName = Some "Calibri"
               Bold = false; Italic = false; Underline = false; UnderlineDouble = false; Strikethrough = false
               Align = align; VerticalAlign = vAlign; LineIndex = 0; LineCount = 1; BaselineShift = 0.0; Rotation = 0 }

let private axisColor = "#595959"
let private gridColor = "#D9D9D9"
// Excel default series palette (accent1..6).
let private palette = [| "#5B9BD5"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#4472C4"; "#70AD47" |]

/// A circle approximated by four cubic Béziers (for markers / pie isn't this).
let private circle (cx: float) (cy: float) (r: float) : VecSeg list =
    let k = r * 0.5522847498
    [ VMove (cx, cy - r)
      VCubic (cx + k, cy - r, cx + r, cy - k, cx + r, cy)
      VCubic (cx + r, cy + k, cx + k, cy + r, cx, cy + r)
      VCubic (cx - k, cy + r, cx - r, cy + k, cx - r, cy)
      VCubic (cx - r, cy - k, cx - k, cy - r, cx, cy - r)
      VClose ]

// Shared plot insets (measured against Excel's column chart).
let private titleSize = 14.0
let private insets (spec: ChartSpec) = (match spec.Title with Some _ -> 49.0 | None -> 10.0), 32.0, 31.0, 28.0  // titleH, leftAxisW, bottomAxisH, rightPad

/// Chart-area background, border, and centred title.
let private frame (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    [ yield FillRect ({ X = 0.0; Y = 0.0; Width = w; Height = h }, "#FFFFFF")
      yield StrokeRect ({ X = 0.5; Y = 0.5; Width = w - 1.0; Height = h - 1.0 }, gridColor, 1.0)
      match spec.Title with
      | Some t -> yield text t { X = 0.0; Y = 15.0; Width = w; Height = titleSize + 4.0 } titleSize "#595959" AlignCenter VAlignTop
      | None -> () ]

let private yAxisLabel (v: float) (y: float) (leftAxisW: float) : PaintItem =
    text (sprintf "%g" v) { X = 0.0; Y = y - 8.0; Width = leftAxisW - 4.0; Height = 14.0 } 10.0 axisColor AlignRight VAlignMiddle

/// Clustered column chart.
let buildColumnChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let n = List.length spec.Values
    let dataMax = if List.isEmpty spec.Values then 1.0 else List.max spec.Values
    let step = niceStep dataMax
    let yMax = niceMax dataMax step
    let nSteps = int (round (yMax / step))
    let titleH, leftAxisW, bottomAxisH, rightPad = insets spec
    let plotX = leftAxisW
    let plotW = w - leftAxisW - rightPad
    let plotH = h - titleH - bottomAxisH
    let baselineY = titleH + plotH
    let yToPx v = baselineY - (v / yMax) * plotH
    [ yield! frame spec w h
      for i in 0 .. nSteps do
          let v = float i * step
          let y = yToPx v
          yield VecPath ([ VMove (plotX, y); VLine (plotX + plotW, y) ], None, Some (gridColor, 1.0))
          yield yAxisLabel v y leftAxisW
      if n > 0 then
          let bandW = plotW / float n
          let barW = bandW / (1.0 + float spec.GapWidth / 100.0)
          for i in 0 .. n - 1 do
              let bandLeft = plotX + float i * bandW
              let barX = bandLeft + (bandW - barW) / 2.0
              let topY = yToPx spec.Values.[i]
              yield VecPath ([ VMove (barX, baselineY); VLine (barX, topY); VLine (barX + barW, topY); VLine (barX + barW, baselineY); VClose ], Some spec.SeriesColor, None)
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = bandLeft; Y = baselineY + 2.0; Width = bandW; Height = bottomAxisH } 10.0 axisColor AlignCenter VAlignTop ]

/// Clustered bar chart (horizontal): value axis along the bottom, categories
/// up the left with category 1 at the bottom.
let buildBarChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let n = List.length spec.Values
    let dataMax = if List.isEmpty spec.Values then 1.0 else List.max spec.Values
    let step = niceStep dataMax
    let vMax = niceMax dataMax step
    let nSteps = int (round (vMax / step))
    let titleH, leftAxisW, bottomAxisH, rightPad = insets spec
    let plotX = leftAxisW
    let plotW = w - leftAxisW - rightPad
    let plotH = h - titleH - bottomAxisH
    let baselineY = titleH + plotH
    let xToPx v = plotX + (v / vMax) * plotW
    [ yield! frame spec w h
      for i in 0 .. nSteps do
          let v = float i * step
          let x = xToPx v
          yield VecPath ([ VMove (x, titleH); VLine (x, baselineY) ], None, Some (gridColor, 1.0))
          yield text (sprintf "%g" v) { X = x - 20.0; Y = baselineY + 2.0; Width = 40.0; Height = 14.0 } 10.0 axisColor AlignCenter VAlignTop
      if n > 0 then
          let bandH = plotH / float n
          let barH = bandH / (1.0 + float spec.GapWidth / 100.0)
          for i in 0 .. n - 1 do
              let bandTop = titleH + plotH - float (i + 1) * bandH   // category 1 at the bottom
              let barY = bandTop + (bandH - barH) / 2.0
              let endX = xToPx spec.Values.[i]
              yield VecPath ([ VMove (plotX, barY); VLine (endX, barY); VLine (endX, barY + barH); VLine (plotX, barY + barH); VClose ], Some spec.SeriesColor, None)
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = 0.0; Y = barY + barH / 2.0 - 7.0; Width = leftAxisW - 4.0; Height = 14.0 } 10.0 axisColor AlignRight VAlignMiddle ]

/// Value-axis chrome (gridlines + y labels) + category x positions, shared by
/// line and area. Returns (items, xOfCategory, yToPx, baselineY, plotX, plotW).
let private valueAxisChrome (edgeMode: bool) (spec: ChartSpec) (w: float) (h: float) =
    let n = List.length spec.Values
    let dataMax = if List.isEmpty spec.Values then 1.0 else List.max spec.Values
    let step = niceStep dataMax
    let yMax = niceMax dataMax step
    let nSteps = int (round (yMax / step))
    let titleH, leftAxisW, bottomAxisH, rightPad = insets spec
    let plotX = leftAxisW
    let plotW = w - leftAxisW - rightPad
    let plotH = h - titleH - bottomAxisH
    let baselineY = titleH + plotH
    let yToPx v = baselineY - (v / yMax) * plotH
    let bandW = if n > 0 then plotW / float n else plotW
    // Line: points at band centres ("between"). Area: edge-to-edge ("on"),
    // first category at plotX, last at plotX+plotW — matches Excel.
    let xOf i =
        if edgeMode then (if n <= 1 then plotX else plotX + float i * plotW / float (n - 1))
        else plotX + (float i + 0.5) * bandW
    let chrome =
        [ yield! frame spec w h
          for i in 0 .. nSteps do
              let v = float i * step
              let y = yToPx v
              yield VecPath ([ VMove (plotX, y); VLine (plotX + plotW, y) ], None, Some (gridColor, 1.0))
              yield yAxisLabel v y leftAxisW
          for i in 0 .. n - 1 do
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = xOf i - bandW / 2.0; Y = baselineY + 2.0; Width = bandW; Height = bandW } 10.0 axisColor AlignCenter VAlignTop ]
    chrome, xOf, yToPx, baselineY

/// Line chart with markers.
let buildLineChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let chrome, xOf, yToPx, _ = valueAxisChrome false spec w h
    let pts = spec.Values |> List.mapi (fun i v -> xOf i, yToPx v)
    [ yield! chrome
      match pts with
      | [] -> ()
      | (x0, y0) :: rest ->
          let segs = VMove (x0, y0) :: (rest |> List.map (fun (x, y) -> VLine (x, y)))
          yield VecPath (segs, None, Some (spec.SeriesColor, 2.0))
          for (x, y) in pts do
              yield VecPath (circle x y 3.5, Some spec.SeriesColor, None) ]

/// Area chart (filled to baseline).
let buildAreaChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let chrome, xOf, yToPx, baselineY = valueAxisChrome true spec w h
    let pts = spec.Values |> List.mapi (fun i v -> xOf i, yToPx v)
    [ yield! chrome
      match pts with
      | [] -> ()
      | (x0, _) :: _ ->
          let xn, _ = List.last pts
          let area = VMove (x0, baselineY) :: (pts |> List.map (fun (x, y) -> VLine (x, y))) @ [ VLine (xn, baselineY); VClose ]
          yield VecPath (area, Some spec.SeriesColor, None) ]

/// Scatter (markers only); both axes numeric.
let buildScatterChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let xsData = spec.Categories |> List.map pf
    let ys = spec.Values
    let xMaxData = if List.isEmpty xsData then 1.0 else List.max xsData
    let yMaxData = if List.isEmpty ys then 1.0 else List.max ys
    let xStep = niceStep xMaxData
    let xMax = niceMax xMaxData xStep
    let yStep = niceStep yMaxData
    let yMax = niceMax yMaxData yStep
    let titleH, leftAxisW, bottomAxisH, rightPad = insets spec
    let plotX = leftAxisW
    let plotW = w - leftAxisW - rightPad
    let plotH = h - titleH - bottomAxisH
    let baselineY = titleH + plotH
    let xToPx v = plotX + (v / xMax) * plotW
    let yToPx v = baselineY - (v / yMax) * plotH
    [ yield! frame spec w h
      for i in 0 .. int (round (yMax / yStep)) do
          let v = float i * yStep
          let y = yToPx v
          yield VecPath ([ VMove (plotX, y); VLine (plotX + plotW, y) ], None, Some (gridColor, 1.0))
          yield yAxisLabel v y leftAxisW
      for i in 0 .. int (round (xMax / xStep)) do
          let v = float i * xStep
          let x = xToPx v
          // Scatter has both horizontal and vertical gridlines.
          yield VecPath ([ VMove (x, titleH); VLine (x, baselineY) ], None, Some (gridColor, 1.0))
          yield text (sprintf "%g" v) { X = x - 20.0; Y = baselineY + 2.0; Width = 40.0; Height = 14.0 } 10.0 axisColor AlignCenter VAlignTop
      for i in 0 .. List.length ys - 1 do
          let x = xToPx (List.item i xsData)
          let y = yToPx ys.[i]
          yield VecPath (circle x y 3.5, Some spec.SeriesColor, None) ]

/// Pie chart (slices coloured from the accent palette).
let buildPieChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let titleH, _, _, _ = insets spec
    let total = List.sum spec.Values
    let pal = if List.isEmpty spec.Palette then palette |> Array.toList else spec.Palette
    let cx = w / 2.0
    // Measured against Excel: centre ~45% down the below-title area, radius
    // ~0.90 of half the smaller plot dimension.
    let cy = titleH + (h - titleH) * 0.45
    let r = (min w (h - titleH)) / 2.0 * 0.90
    let pt a = cx + r * sin a, cy - r * cos a       // a = radians from top, clockwise
    [ yield! frame spec w h
      if total > 0.0 then
          let mutable a0 = 0.0
          for i in 0 .. List.length spec.Values - 1 do
              let v = spec.Values.[i]
              let a1 = a0 + (v / total) * 2.0 * System.Math.PI
              let steps = max 2 (int (ceil ((a1 - a0) / (System.Math.PI / 90.0))))
              let arc = [ for s in 0 .. steps -> let a = a0 + (a1 - a0) * float s / float steps in let px, py = pt a in VLine (px, py) ]
              let segs = VMove (cx, cy) :: arc @ [ VClose ]
              yield VecPath (segs, Some (List.item (i % List.length pal) pal), Some ("#FFFFFF", 1.0))
              a0 <- a1 ]

/// Dispatch on chart kind.
let buildChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    match spec.Kind with
    | "bar" -> buildBarChart spec w h
    | "line" -> buildLineChart spec w h
    | "area" -> buildAreaChart spec w h
    | "scatter" -> buildScatterChart spec w h
    | "pie" -> buildPieChart spec w h
    | _ -> buildColumnChart spec w h
