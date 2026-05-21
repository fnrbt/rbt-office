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
      SeriesColor: string }        // "#RRGGBB"

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

let parseChart (xlsxPath: string) : ChartSpec option =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip "xl/charts/chart1.xml" with
    | None -> None
    | Some xml ->
        let doc = XDocument.Parse xml
        let root = doc.Root
        match firstDesc "barChart" root with
        | None -> None
        | Some bar ->
            let kind = firstDesc "barDir" bar |> Option.map (fun e -> e.Attribute(XName.Get "val").Value) |> Option.defaultValue "col"
            let ser = firstDesc "ser" bar
            let cats = ser |> Option.bind (firstDesc "cat") |> Option.map points |> Option.defaultValue []
            let vals =
                ser |> Option.bind (firstDesc "val") |> Option.map points |> Option.defaultValue []
                |> List.map (fun s -> match Double.TryParse(s, Globalization.CultureInfo.InvariantCulture) with true, n -> n | _ -> 0.0)
            let title =
                firstDesc "title" root
                |> Option.bind (fun t -> descendants "t" t |> Seq.tryHead)
                |> Option.map (fun t -> t.Value)
            let gridlines = root |> descendants "valAx" |> Seq.exists (fun ax -> firstDesc "majorGridlines" ax |> Option.isSome)
            let gap = firstDesc "gapWidth" bar |> Option.map (fun e -> int (e.Attribute(XName.Get "val").Value)) |> Option.defaultValue 150
            Some { Kind = kind; Title = title; Categories = cats; Values = vals
                   Gridlines = gridlines; GapWidth = gap; SeriesColor = "#5B9BD5" }

/// Chart pixel size from the twoCellAnchor (default 64px cols / 20px rows;
/// EMU offsets at 9525 EMU/px). Good enough for the synthetic fixture.
let parseAnchorSize (xlsxPath: string) : float * float =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip "xl/drawings/drawing1.xml" with
    | None -> 480.0, 320.0
    | Some xml ->
        let doc = XDocument.Parse xml
        let pick tag child =
            firstDesc tag doc.Root
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

/// Lay out a clustered column chart into a (w × h) rect, Excel-faithfully.
let buildColumnChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let axisColor = "#595959"
    let gridColor = "#D9D9D9"
    let n = List.length spec.Values
    let dataMax = if List.isEmpty spec.Values then 1.0 else List.max spec.Values
    let step = niceStep dataMax
    let yMax = niceMax dataMax step
    let nSteps = int (round (yMax / step))
    // Title band (Calibri 14, grey), reserved at the top. Insets measured
    // against Excel's plot rect for the synthetic 480x320 chart: plot is
    // x 32..452 (w 420), y 49..289 (h 240) inside the chart area.
    let titleSize = 14.0
    let titleH = match spec.Title with Some _ -> 49.0 | None -> 10.0
    // Plot insets.
    let leftAxisW = 32.0      // room for y labels
    let bottomAxisH = 31.0    // room for x labels
    let rightPad = 28.0
    let plotX = leftAxisW
    let plotY = titleH
    let plotW = w - leftAxisW - rightPad
    let plotH = h - titleH - bottomAxisH
    let baselineY = plotY + plotH
    let yToPx v = baselineY - (v / yMax) * plotH
    [
      // Chart-area background + light border.
      yield FillRect ({ X = 0.0; Y = 0.0; Width = w; Height = h }, "#FFFFFF")
      yield StrokeRect ({ X = 0.5; Y = 0.5; Width = w - 1.0; Height = h - 1.0 }, gridColor, 1.0)
      // Title.
      match spec.Title with
      | Some t -> yield text t { X = 0.0; Y = 15.0; Width = w; Height = titleSize + 4.0 } titleSize "#595959" AlignCenter VAlignTop
      | None -> ()
      // Horizontal gridlines + y-axis labels.
      for i in 0 .. nSteps do
          let v = float i * step
          let y = yToPx v
          yield VecPath ([ VMove (plotX, y); VLine (plotX + plotW, y) ], None, Some (gridColor, 1.0))
          let lbl = if v = floor v then sprintf "%g" v else sprintf "%g" v
          yield text lbl { X = 0.0; Y = y - 8.0; Width = leftAxisW - 4.0; Height = 14.0 } 10.0 axisColor AlignRight VAlignMiddle
      // Bars + category labels.
      if n > 0 then
          let bandW = plotW / float n
          let barW = bandW / (1.0 + float spec.GapWidth / 100.0)
          for i in 0 .. n - 1 do
              let v = spec.Values.[i]
              let bandLeft = plotX + float i * bandW
              let barX = bandLeft + (bandW - barW) / 2.0
              let topY = yToPx v
              yield VecPath ([ VMove (barX, baselineY); VLine (barX, topY); VLine (barX + barW, topY); VLine (barX + barW, baselineY); VClose ],
                             Some spec.SeriesColor, None)
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = bandLeft; Y = baselineY + 2.0; Width = bandW; Height = bottomAxisH } 10.0 axisColor AlignCenter VAlignTop
    ]
