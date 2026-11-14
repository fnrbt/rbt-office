#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open Rbt.Spreadsheet.Model
module Charts = Rbt.Spreadsheet.Charts
module Bridge = Rbt.Spreadsheet.XlsxBridge

// Charts and shapes of a drawing: a chart part read into a model, laid out
// for a rectangle (a scene of rectangles, lines, paths and texts) from the
// cells its series name - or from the file's cache - and written as SVG.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-46s %s" label detail)
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let text (data: byte[]) = Text.Encoding.UTF8.GetString data

/// A workbook's parts, its sheets and what a chart reads of them.
let openBook (file: string) =
    let wb = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, file))) |> Option.get
    let book = Bridge.Import.importWorkbook wb
    let parts = wb.Preserved.Parts |> Map.ofList
    let theme = wb.Preserved.ThemePath |> Option.bind (fun path -> Map.tryFind path parts) |> Option.map (text >> Charts.parseTheme) |> Option.defaultValue Charts.defaultTheme
    wb, book, parts, theme
let dataOf (sheets: (string * SpreadsheetModel) list) : Charts.ChartData =
    { Cells = (fun name r0 c0 r1 c1 ->
        sheets |> List.tryFind (fun (known, _) -> String.Equals(known, name, StringComparison.OrdinalIgnoreCase))
        |> Option.map (fun (_, model) -> Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } model))))
      Display = (fun cell -> Rbt.Spreadsheet.CellText.cellDisplayText None 1000.0 cell.Format cell.Value)
      Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
let noData : Charts.ChartData = { Cells = (fun _ _ _ _ _ -> None); Display = (fun _ -> ""); Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
let rects (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PRect (x, y, w, h, Some fill, _, _, _, _) -> Some (x, y, w, h, fill) | _ -> None)
let texts (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PText (_, _, value, _, _, _, _, _, _) -> Some value | _ -> None)
let paths (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PPath (d, Some fill, _, _, _) -> Some (d, fill) | _ -> None)
let lines (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PPolyline (points, None, _, Some stroke, _, _) -> Some (points, stroke) | _ -> None)

// ---- every chart of the three workbooks is read and laid out -----------------------
for file in [ "real-world/chart-zoo.xlsx"; "real-world/dashboard-sales-analysis.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx" ] do
    let _, book, parts, theme = openBook file
    let data = dataOf (book.Sheets |> Array.toList |> List.map (fun sheet -> sheet.Name, sheet.Model))
    let chartParts = parts |> Map.toList |> List.filter (fun (path, _) -> Text.RegularExpressions.Regex.IsMatch(path, "^xl/charts/chart[0-9]+\\.xml$"))
    ok (sprintf "%s has charts" file) (not chartParts.IsEmpty) "none"
    for path, bytes in chartParts do
        match Charts.parseChart theme (text bytes) with
        | None -> ok (sprintf "%s %s: read" (Path.GetFileName file) path) false "not read"
        | Some chart ->
            let scene = Charts.layoutChart theme data chart 480.0 288.0
            let svg = Charts.sceneSvg scene
            ok (sprintf "%s %s: drawn" (Path.GetFileName file) path)
                (scene.Items.Length >= 1 && svg.StartsWith "<svg" && svg.EndsWith "</svg>" && not (svg.Contains "NaN") && not (svg.Contains "Infinity"))
                (sprintf "%d items" scene.Items.Length)
            ok (sprintf "%s %s: the SVG is XML" (Path.GetFileName file) path) (try (Xml.Linq.XDocument.Parse svg |> ignore; true) with _ -> false) "not XML"
            // Whatever is drawn is inside the rectangle.
            let outside =
                scene.Items |> List.filter (fun item ->
                    match item with
                    | Charts.PRect (x, y, w, h, _, _, _, _, _) -> x < -1.0 || y < -1.0 || x + w > 481.0 || y + h > 289.0
                    | Charts.PLine (x1, y1, x2, y2, _, _, _) -> min x1 x2 < -1.0 || min y1 y2 < -1.0 || max x1 x2 > 481.0 || max y1 y2 > 289.0
                    | _ -> false)
            ok (sprintf "%s %s: inside its rectangle" (Path.GetFileName file) path) outside.IsEmpty (sprintf "%A" (List.truncate 2 outside))

// ---- chart-zoo: what each kind draws -----------------------------------------------
let zooWb, zooBook, zooParts, zooTheme = openBook "real-world/chart-zoo.xlsx"
let zoo = zooBook.Sheets |> Array.find (fun sheet -> sheet.Name = "Zoo")
let zooData = dataOf [ "Zoo", zoo.Model ]
let chartOf (number: int) = Charts.parseChart zooTheme (text zooParts.[sprintf "xl/charts/chart%d.xml" number]) |> Option.get
let accent1 = "#" + zooTheme.["accent1"]

// The column chart: Sales by quarter (10, 25, 15, 30).
let column = chartOf 1
eq "column: one bar plot, upright, clustered" (column.Plots |> List.map (fun plot -> plot.Kind, plot.Horizontal, plot.Grouping, plot.Series.Length)) [ Charts.BarPlot, false, "clustered", 1 ]
eq "column: the title" column.Title (Some "Column")
let columnScene = Charts.layoutChart zooTheme zooData column 480.0 288.0
let bars = rects columnScene |> List.filter (fun (_, _, _, _, fill) -> fill = accent1)
eq "column: four bars in the series' colour (the theme's accent 1)" bars.Length 4
let heights = bars |> List.map (fun (_, _, _, h, _) -> h)
ok "column: the bars are as high as their values (10, 25, 15, 30)" (List.forall2 (fun (h: float) (v: float) -> abs (h / heights.[3] - v / 30.0) < 0.01) heights [ 10.0; 25.0; 15.0; 30.0 ]) (sprintf "%A" heights)
ok "column: they stand on one line, left to right" (bars |> List.map (fun (_, y, _, h, _) -> Math.Round (y + h)) |> List.distinct |> List.length = 1 && (bars |> List.map (fun (x, _, _, _, _) -> x)) = (bars |> List.map (fun (x, _, _, _, _) -> x) |> List.sort)) (sprintf "%A" bars)
ok "column: the title, the categories and the scale 0 .. 35 by 5" ([ "Column"; "Q1"; "Q2"; "Q3"; "Q4"; "0"; "5"; "35" ] |> List.forall (fun label -> List.contains label (texts columnScene))) (sprintf "%A" (texts columnScene))
ok "column: no 40 on the scale" (not (List.contains "40" (texts columnScene))) "a 40"
// Live values: the cell of Q2 changed.
let edited = setCell { Row = 2; Col = 1 } { Value = CellValue.Number 60.0; Format = defaultFormat; Comment = None } zoo.Model
let editedScene = Charts.layoutChart zooTheme (dataOf [ "Zoo", edited ]) column 480.0 288.0
let editedBars = rects editedScene |> List.filter (fun (_, _, _, _, fill) -> fill = accent1) |> List.map (fun (_, _, _, h, _) -> h)
ok "column: a cell changed, its bar is the highest and the scale goes to 70" (editedBars.[1] > editedBars.[3] * 1.9 && List.contains "70" (texts editedScene)) (sprintf "%A %A" editedBars (texts editedScene))
// Without the sheet: the file's cache.
let cachedScene = Charts.layoutChart zooTheme noData column 480.0 288.0
eq "column: without its sheet the cached values are drawn" (rects cachedScene |> List.filter (fun (_, _, _, _, fill) -> fill = accent1) |> List.map (fun (_, _, _, h, _) -> Math.Round h)) (heights |> List.map Math.Round)
ok "column: ... and the cached categories" ([ "Q1"; "Q4" ] |> List.forall (fun label -> List.contains label (texts cachedScene))) (sprintf "%A" (texts cachedScene))

// The bar chart: bars that run to the right, the first category at the bottom.
let bar = chartOf 2
eq "bar: horizontal" (bar.Plots |> List.map (fun plot -> plot.Kind, plot.Horizontal)) [ Charts.BarPlot, true ]
let barScene = Charts.layoutChart zooTheme zooData bar 480.0 288.0
let barRects = rects barScene |> List.filter (fun (_, _, w, h, fill) -> fill <> "#FFFFFF" && w > 0.0 && h < 100.0)
ok "bar: bars of one height that start at one x" (barRects.Length >= 4 && barRects |> List.map (fun (x, _, _, _, _) -> Math.Round x) |> List.distinct |> List.length = 1 && barRects |> List.map (fun (_, _, _, h, _) -> Math.Round (h * 10.0)) |> List.distinct |> List.length = 1) (sprintf "%A" barRects)
ok "bar: the first category is the lowest bar" (let (_, y0, _, _, _), (_, y1, _, _, _) = barRects.[0], barRects.[1] in y0 > y1) (sprintf "%A" barRects)

// The line chart.
let line = chartOf 3
eq "line: a line plot" (line.Plots |> List.map (fun plot -> plot.Kind)) [ Charts.LinePlot ]
let lineScene = Charts.layoutChart zooTheme zooData line 480.0 288.0
ok "line: one polyline per series, a point per category" (let drawn = lines lineScene in drawn.Length = line.Plots.Head.Series.Length && drawn |> List.forall (fun (points, _) -> points.Length = 4)) (sprintf "%A" (lines lineScene |> List.map (fun (points, _) -> points.Length)))

// The pie chart: a wedge per category, each in a colour of its own.
let pie = chartOf 4
let pieScene = Charts.layoutChart zooTheme zooData pie 480.0 288.0
eq "pie: a pie plot" (pie.Plots |> List.map (fun plot -> plot.Kind)) [ Charts.PiePlot ]
ok "pie: four wedges in four colours, from the centre" (let wedges = paths pieScene in wedges.Length = 4 && (wedges |> List.map snd |> List.distinct).Length = 4 && wedges |> List.forall (fun (d, _) -> d.Contains "A" && d.EndsWith "Z")) (sprintf "%A" (paths pieScene |> List.map snd))
ok "pie: a legend, when the chart has one, names the categories" (pie.Legend.IsNone || [ "Q1"; "Q2"; "Q3"; "Q4" ] |> List.forall (fun label -> List.contains label (texts pieScene))) (sprintf "%A" (texts pieScene))
let pieWithLegend = Charts.layoutChart zooTheme zooData { pie with Legend = Some { Position = "r"; Style = { Size = None; Color = None; Bold = None; Italic = None }; Layout = None; Overlay = false; Deleted = Set.empty } } 480.0 288.0
ok "pie: with a legend at the right, its entries are the categories" ([ "Q1"; "Q2"; "Q3"; "Q4" ] |> List.forall (fun label -> List.contains label (texts pieWithLegend))) (sprintf "%A" (texts pieWithLegend))

// The area chart and the scatter chart.
let area = chartOf 5
let areaScene = Charts.layoutChart zooTheme zooData area 480.0 288.0
ok "area: a filled polygon down to the axis" (areaScene.Items |> List.exists (fun item -> match item with Charts.PPolyline (points, Some _, _, _, _, _) -> points.Length = 8 | _ -> false)) "none"
let scatter = chartOf 6
let scatterScene = Charts.layoutChart zooTheme zooData scatter 480.0 288.0
eq "scatter: a scatter plot with x values" (scatter.Plots |> List.map (fun plot -> plot.Kind, plot.Series |> List.forall (fun series -> series.XValues.IsSome))) [ Charts.ScatterPlot, true ]
ok "scatter: a mark or a line per series" (scatterScene.Items |> List.exists (fun item -> match item with Charts.PEllipse _ | Charts.PPolyline _ -> true | _ -> false)) "nothing"

// ---- stacked and 100 % stacked columns, a secondary axis, a doughnut (made here) ------
let chartXml (plots: string) (axes: string) =
    "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><c:chart><c:autoTitleDeleted val=\"1\"/><c:plotArea>"
    + plots + axes + "</c:plotArea><c:legend><c:legendPos val=\"b\"/></c:legend></c:chart></c:chartSpace>"
let literal (values: float list) =
    sprintf "<c:numLit><c:ptCount val=\"%d\"/>%s</c:numLit>" values.Length (values |> List.mapi (fun index value -> sprintf "<c:pt idx=\"%d\"><c:v>%s</c:v></c:pt>" index (string value)) |> String.concat "")
let ser (index: int) (color: string) (values: float list) (extra: string) =
    sprintf "<c:ser><c:idx val=\"%d\"/><c:order val=\"%d\"/><c:tx><c:v>S%d</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill><a:ln w=\"28575\"><a:solidFill><a:srgbClr val=\"%s\"/></a:solidFill></a:ln></c:spPr>%s<c:val>%s</c:val></c:ser>" index index (index + 1) color color extra (literal values)
let axesXml = "<c:catAx><c:axId val=\"1\"/><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx><c:valAx><c:axId val=\"2\"/><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:numFmt formatCode=\"#,##0\" sourceLinked=\"0\"/><c:crossAx val=\"1\"/></c:valAx>"
let barPlot (grouping: string) (series: string) = sprintf "<c:barChart><c:barDir val=\"col\"/><c:grouping val=\"%s\"/>%s<c:gapWidth val=\"100\"/><c:overlap val=\"100\"/><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>" grouping series
let parsed (xml: string) = Charts.parseChart Charts.defaultTheme xml |> Option.get
let stackedScene = Charts.layoutChart Charts.defaultTheme noData (parsed (chartXml (barPlot "stacked" (ser 0 "FF0000" [ 1000.0; 2000.0 ] "" + ser 1 "0000FF" [ 3000.0; 1000.0 ] "")) axesXml)) 400.0 300.0
// (Without the squares of the legend.)
let barsOf (scene: Charts.Scene) (color: string) = rects scene |> List.filter (fun (_, _, w, _, fill) -> fill = color && w > 20.0)
let red, blue = barsOf stackedScene "#FF0000", barsOf stackedScene "#0000FF"
ok "stacked: the second series stands on the first" (red.Length = 2 && blue.Length = 2 && List.forall2 (fun (rx, ry, _, _, _) (bx, by, _, bh, _) -> abs (rx - bx) < 0.01 && abs (by + bh - ry) < 0.01) red blue) (sprintf "%A %A" red blue)
ok "stacked: the scale holds the sums (to 4,500, in the axis' number format)" (List.contains "4,500" (texts stackedScene) && List.contains "1,000" (texts stackedScene)) (sprintf "%A" (texts stackedScene))
ok "stacked: grid lines, and the legend names the series" (stackedScene.Items |> List.exists (fun item -> match item with Charts.PLine _ -> true | _ -> false) && List.contains "S1" (texts stackedScene) && List.contains "S2" (texts stackedScene)) "missing"
let percentScene = Charts.layoutChart Charts.defaultTheme noData (parsed (chartXml (barPlot "percentStacked" (ser 0 "FF0000" [ 1.0; 3.0 ] "" + ser 1 "0000FF" [ 3.0; 1.0 ] "")) axesXml)) 400.0 300.0
let redP, blueP = barsOf percentScene "#FF0000", barsOf percentScene "#0000FF"
ok "100 % stacked: every column is as high as the plot, in the shares of its values" (let (_, _, _, r0, _), (_, _, _, b0, _), (_, _, _, r1, _), (_, _, _, b1, _) = redP.[0], blueP.[0], redP.[1], blueP.[1] in abs (r0 + b0 - (r1 + b1)) < 0.01 && abs (b0 / r0 - 3.0) < 0.01 && abs (r1 / b1 - 3.0) < 0.01) (sprintf "%A %A" redP blueP)
ok "100 % stacked: the scale is 0% .. 100%" (List.contains "0%" (texts percentScene) && List.contains "100%" (texts percentScene) && List.contains "50%" (texts percentScene)) (sprintf "%A" (texts percentScene))
// A combination: columns on the first axis, a line on a second one at the right.
let comboAxes = axesXml + "<c:catAx><c:axId val=\"3\"/><c:delete val=\"1\"/><c:axPos val=\"b\"/><c:crossAx val=\"4\"/></c:catAx><c:valAx><c:axId val=\"4\"/><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:numFmt formatCode=\"0%\" sourceLinked=\"0\"/><c:crossAx val=\"3\"/></c:valAx>"
let combo = parsed (chartXml (barPlot "clustered" (ser 0 "FF0000" [ 100.0; 300.0; 200.0 ] "") + "<c:lineChart><c:grouping val=\"standard\"/>" + ser 1 "00AA00" [ 0.1; 0.5; 0.25 ] "<c:marker><c:symbol val=\"circle\"/><c:size val=\"7\"/></c:marker>" + "<c:marker val=\"1\"/><c:axId val=\"3\"/><c:axId val=\"4\"/></c:lineChart>") comboAxes)
let comboScene = Charts.layoutChart Charts.defaultTheme noData combo 400.0 300.0
eq "combination: a bar plot and a line plot" (combo.Plots |> List.map (fun plot -> plot.Kind)) [ Charts.BarPlot; Charts.LinePlot ]
ok "combination: the line has its own scale, in per cent, at the right" (List.contains "50%" (texts comboScene) && List.contains "300" (texts comboScene) && (comboScene.Items |> List.exists (fun item -> match item with Charts.PText (x, _, "50%", _, _, "start", _, _, _) -> x > 300.0 | _ -> false))) (sprintf "%A" (texts comboScene))
ok "combination: the line is drawn over the columns, with a marker per point" (let drawn = lines comboScene in drawn.Length = 1 && snd drawn.Head = "#00AA00" && (comboScene.Items |> List.filter (fun item -> match item with Charts.PEllipse _ -> true | _ -> false)).Length = 3 && (let lineAt = comboScene.Items |> List.findIndex (fun item -> match item with Charts.PPolyline _ -> true | _ -> false) in let lastBar = comboScene.Items |> List.findIndexBack (fun item -> match item with Charts.PRect (_, _, w, _, Some "#FF0000", _, _, _, _) -> w > 20.0 | _ -> false) in lineAt > lastBar)) "not so"
ok "combination: the line reaches its 50 % near the top (its own scale)" (let points, _ = (lines comboScene).Head in let ys = points |> List.map snd in List.min ys < 80.0) (sprintf "%A" (lines comboScene))
// A doughnut with data labels in per cent.
let doughnut = parsed (chartXml ("<c:doughnutChart><c:varyColors val=\"1\"/>" + ser 0 "FF0000" [ 1.0; 1.0; 2.0 ] "<c:dLbls><c:showVal val=\"0\"/><c:showPercent val=\"1\"/></c:dLbls>" + "<c:firstSliceAng val=\"0\"/><c:holeSize val=\"60\"/></c:doughnutChart>") "")
let doughnutScene = Charts.layoutChart Charts.defaultTheme noData doughnut 300.0 300.0
ok "doughnut: three rings segments with a hole (two arcs each), labelled 25%, 25%, 50%" (let wedges = paths doughnutScene in wedges.Length = 3 && wedges |> List.forall (fun (d, _) -> (d.Split('A').Length - 1) = 2) && [ "25%"; "50%" ] |> List.forall (fun label -> List.contains label (texts doughnutScene))) (sprintf "%A %A" (paths doughnutScene |> List.map fst) (texts doughnutScene))
// What is not drawn says so.
let plotKindsOf xml = Charts.parseChart Charts.defaultTheme xml |> Option.map (fun chart -> chart.Plots |> List.map (fun plot -> plot.Kind))
eq "a radar chart is read" (plotKindsOf (chartXml ("<c:radarChart><c:radarStyle val=\"marker\"/>" + ser 0 "FF0000" [ 1.0 ] "" + "</c:radarChart>") "")) (Some [ Charts.RadarPlot ])
eq "a 3-D bar chart is read as a flat one, beside a known one too" (plotKindsOf (chartXml (barPlot "clustered" (ser 0 "FF0000" [ 1.0 ] "") + "<c:bar3DChart><c:barDir val=\"col\"/>" + ser 1 "00FF00" [ 2.0 ] "" + "</c:bar3DChart>") axesXml)) (Some [ Charts.BarPlot; Charts.BarPlot ])
eq "what is not a chart" (Charts.parseChart Charts.defaultTheme "<x/>", Charts.parseChart Charts.defaultTheme "not xml") (None, None)

// ---- the theme ---------------------------------------------------------------------
let _, _, ecoParts, ecoTheme = openBook "real-world/dashboard-ecommerce-sales.xlsx"
eq "the theme's colours are read (Integral: accent1 1CADE4)" (ecoTheme.["accent1"], ecoTheme.["dk1"], ecoTheme.["lt1"]) ("1CADE4", "000000", "FFFFFF")

// ---- shapes and groups ---------------------------------------------------------------
let noParts : Charts.DrawingParts = { Theme = Charts.defaultTheme; Chart = (fun _ -> None); Picture = (fun _ -> None); Slicer = (fun _ _ _ _ _ -> None) }
let ns = "xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
let shapeXml (geometry: string) (extra: string) (body: string) =
    sprintf "<xdr:sp><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1905000\" cy=\"952500\"/></a:xfrm><a:prstGeom prst=\"%s\"><a:avLst/></a:prstGeom>%s</xdr:spPr>%s</xdr:sp>" geometry extra body
let anchorOf (content: string) = sprintf "<xdr:twoCellAnchor %s><xdr:from/><xdr:to/>%s<xdr:clientData/></xdr:twoCellAnchor>" ns content
let fillXml = "<a:solidFill><a:srgbClr val=\"336699\"/></a:solidFill><a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"112233\"/></a:solidFill></a:ln>"
let bodyXml = "<xdr:txBody><a:bodyPr anchor=\"ctr\"/><a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr sz=\"1400\" b=\"1\"><a:solidFill><a:srgbClr val=\"FFFFFF\"/></a:solidFill></a:rPr><a:t>Total </a:t></a:r><a:r><a:rPr sz=\"1400\"/><a:t>sales</a:t></a:r></a:p></xdr:txBody>"
let rounded = Charts.objectScene noParts noData (anchorOf (shapeXml "roundRect" fillXml bodyXml)) 200.0 100.0 |> Option.get
ok "a rounded rectangle: its fill, its outline (2 px), round corners" (rounded.Items |> List.exists (fun item -> match item with Charts.PRect (0.0, 0.0, 200.0, 100.0, Some "#336699", _, Some "#112233", 2.0, radius) -> radius > 5.0 | _ -> false)) (sprintf "%A" rounded.Items)
ok "... its text is one line in runs: bold white 14 pt, then plain, centred both ways"
    (match rounded.Items |> List.choose (fun item -> match item with Charts.PTextRuns (x, y, anchor, runs) -> Some (x, y, anchor, runs) | _ -> None) with
     | [ (x, y, "middle", [ ("Total ", size, "#FFFFFF", true, false); ("sales", _, "#000000", false, false) ]) ] -> x = 100.0 && abs (y - 50.0) < 1.0 && Math.Round size = 19.0
     | _ -> false) (sprintf "%A" rounded.Items)
let longText = "<xdr:txBody><a:bodyPr/><a:p><a:r><a:rPr sz=\"1100\"/><a:t>This shape holds a text that is much longer than the shape is wide and so it is broken between its words</a:t></a:r></a:p></xdr:txBody>"
let wrappedShape = Charts.objectScene noParts noData (anchorOf (shapeXml "rect" fillXml longText)) 200.0 100.0 |> Option.get
ok "a text longer than its shape is broken between words, line under line"
    (let rows = wrappedShape.Items |> List.choose (fun item -> match item with Charts.PTextRuns (_, y, "start", runs) -> Some (y, runs |> List.map (fun (text, _, _, _, _) -> text) |> String.concat "") | _ -> None)
     rows.Length >= 3 && (rows |> List.map fst) = (rows |> List.map fst |> List.sort) && (rows |> List.map snd |> String.concat "").StartsWith "This shape holds a text" && rows |> List.forall (fun (_, line) -> line.Length < 45))
    (sprintf "%A" wrappedShape.Items)
// What stands for an object of a newer kind (a slicer, a chartEx: `mc:AlternateContent` with a frame in its choice) is not drawn as its fallback text box.
let slicerXml = "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice Requires=\"a14\"><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Region\"/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></xdr:xfrm></xdr:graphicFrame></mc:Choice><mc:Fallback>" + shapeXml "rect" fillXml longText + "</mc:Fallback></mc:AlternateContent>"
eq "a slicer is not drawn as the text box that stands in for it: a placeholder" (Charts.objectScene noParts noData (anchorOf slicerXml) 200.0 100.0) None
ok "an ellipse and a plain rectangle" ((Charts.objectScene noParts noData (anchorOf (shapeXml "ellipse" fillXml "")) 200.0 100.0).Value.Items |> List.exists (fun item -> match item with Charts.PEllipse (100.0, 50.0, 100.0, 50.0, Some "#336699", _, _, _) -> true | _ -> false)
                                       && (Charts.objectScene noParts noData (anchorOf (shapeXml "rect" fillXml "")) 200.0 100.0).Value.Items |> List.exists (fun item -> match item with Charts.PRect (_, _, _, _, Some "#336699", _, _, _, 0.0) -> true | _ -> false)) "not drawn"
eq "a shape without fill, outline or text draws nothing: a placeholder" (Charts.objectScene noParts noData (anchorOf (shapeXml "rect" "<a:noFill/><a:ln><a:noFill/></a:ln>" "")) 200.0 100.0) None
// A group: its children at their place in it, scaled to the rectangle the group has on the sheet.
let groupXml =
    "<xdr:grpSp><xdr:grpSpPr><a:xfrm><a:off x=\"1000000\" y=\"1000000\"/><a:ext cx=\"4000000\" cy=\"2000000\"/><a:chOff x=\"1000000\" y=\"1000000\"/><a:chExt cx=\"4000000\" cy=\"2000000\"/></a:xfrm></xdr:grpSpPr>"
    + "<xdr:sp><xdr:spPr><a:xfrm><a:off x=\"1000000\" y=\"1000000\"/><a:ext cx=\"2000000\" cy=\"2000000\"/></a:xfrm><a:prstGeom prst=\"rect\"/><a:solidFill><a:srgbClr val=\"AA0000\"/></a:solidFill></xdr:spPr></xdr:sp>"
    + "<xdr:graphicFrame><xdr:xfrm><a:off x=\"3000000\" y=\"1000000\"/><a:ext cx=\"2000000\" cy=\"2000000\"/></xdr:xfrm><a:graphic><a:graphicData><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId7\"/></a:graphicData></a:graphic></xdr:graphicFrame></xdr:grpSp>"
let withChart : Charts.DrawingParts = { noParts with Chart = (fun id -> if id = "rId7" then Some column else None) }
let group = Charts.objectScene withChart zooData (anchorOf groupXml) 400.0 200.0 |> Option.get
ok "a group: the shape in its left half, the chart in its right half" (group.Items |> List.exists (fun item -> match item with Charts.PRect (0.0, 0.0, 200.0, 200.0, Some "#AA0000", _, _, _, _) -> true | _ -> false) && group.Items |> List.exists (fun item -> match item with Charts.PGroup (200.0, 0.0, 200.0, 200.0, items) -> items.Length > 8 | _ -> false)) (sprintf "%A" (group.Items |> List.map (fun item -> (sprintf "%A" item).Substring(0, 40))))
ok "a group whose chart is not there still draws its shape, and a box where the chart is" ((Charts.objectScene noParts zooData (anchorOf groupXml) 400.0 200.0).Value.Items |> List.map (fun item -> match item with Charts.PRect (x, _, _, _, Some fill, _, _, _, _) -> sprintf "%s@%d" fill (int x) | Charts.PText _ -> "text" | _ -> "?") = [ "#AA0000@0"; "#F6F8FA@200"; "text" ]) "not so"
// The space a run begins or ends with is text: what is kept of a drawing has it (`<a:t> Sales Analysis Dashboard</a:t>`).
do
    let wb, _, _, _ = openBook "real-world/dashboard-ecommerce-sales.xlsx"
    let kept = wb.Sheets |> List.collect (fun sheet -> sheet.Extras.Objects |> List.map (fun item -> item.Xml)) |> String.concat ""
    ok "a kept shape keeps the space its run begins with" (kept.Contains "<a:t> Sales Analysis Dashboard</a:t>" && kept.Contains "<a:t> sold by Quantity</a:t>" && not (kept.Contains "\uE000")) "the space is gone"
    let titleShape = wb.Sheets |> List.collect (fun sheet -> sheet.Extras.Objects) |> List.find (fun item -> item.Xml.Contains "<a:t> Sales Analysis Dashboard</a:t>")
    let scene = Charts.objectScene noParts noData titleShape.Xml 400.0 50.0 |> Option.get
    ok "... and it is drawn with it" (scene.Items |> List.exists (fun item -> match item with Charts.PTextRuns (_, _, _, runs) -> (runs |> List.map (fun (text, _, _, _, _) -> text) |> String.concat "") = "E-Commerce Sales Analysis Dashboard" | _ -> false)) (sprintf "%A" scene.Items)
// The dashboards group their charts ("Group 57"): every such anchor is drawn with its charts.
for file in [ "real-world/dashboard-ecommerce-sales.xlsx"; "real-world/dashboard-sales-analysis.xlsx" ] do
    let wb, book, parts, theme = openBook file
    let data = dataOf (book.Sheets |> Array.toList |> List.map (fun sheet -> sheet.Name, sheet.Model))
    let mutable groups, drawnCharts = 0, 0
    for sheet in wb.Sheets do
        let target (id: string) = sheet.Extras.DrawingRelationships |> List.tryPick (fun (known, _, target, isExternal) -> if known = id && not isExternal then Some (ZipInterop.normalizePath target) else None)
        let drawing : Charts.DrawingParts =
            { Theme = theme
              Chart = (fun id -> target id |> Option.bind (fun path -> Map.tryFind path parts) |> Option.bind (fun bytes -> Charts.parseChart theme (text bytes)))
              Picture = (fun id -> target id |> Option.map (fun path -> "blob:" + path)); Slicer = (fun _ _ _ _ _ -> None) }
        for item in sheet.Extras.Objects do
            if item.Xml.Contains "grpSp" || item.Kind = "chart" then
                match Charts.objectScene drawing data item.Xml 600.0 400.0 with
                | Some scene ->
                    if item.Xml.Contains "grpSp" then groups <- groups + 1
                    let rec count (items: Charts.Primitive list) = items |> List.sumBy (fun entry -> match entry with Charts.PGroup (_, _, _, _, inner) -> 1 + count inner | _ -> 0)
                    drawnCharts <- drawnCharts + (if item.Kind = "chart" && not (item.Xml.Contains "grpSp") then 1 else count scene.Items)
                | None -> ()
    ok (sprintf "%s: its charts are drawn, also those in groups" (Path.GetFileName file)) (drawnCharts >= 5) (sprintf "%d groups, %d charts" groups drawnCharts)
    printfn "  (%s: %d groups drawn, %d charts)" (Path.GetFileName file) groups drawnCharts

// ---- a pivot chart reads the points its cache has, not the pivot table's grand total --------
// The e-commerce dashboard's charts are of pivot tables: a series names the table's column down
// to its last row (`$A$4:$A$16`, 13 cells) and caches 12 points - the 13th cell is "Grand Total".
// A chart draws as many points as the file cached, from the cells as they are now: the same
// chart whether its source sheet has been converted (shown) or not.
do
    let wb, book, parts, theme = openBook "real-world/dashboard-ecommerce-sales.xlsx"
    let live = dataOf (book.Sheets |> Array.toList |> List.map (fun sheet -> sheet.Name, resolveSheet sheet.Model))
    let rec textsIn (items: Charts.Primitive list) : string list =
        items |> List.collect (fun entry ->
            match entry with
            | Charts.PGroup (_, _, _, _, inner) -> textsIn inner
            | Charts.PText (_, _, value, _, _, _, _, _, _) -> [ value ]
            | _ -> [])
    let mutable charts, withTotal, differing = 0, [], []
    for sheet in wb.Sheets do
        let target (id: string) = sheet.Extras.DrawingRelationships |> List.tryPick (fun (known, _, target, isExternal) -> if known = id && not isExternal then Some (ZipInterop.normalizePath target) else None)
        let drawing : Charts.DrawingParts =
            { Theme = theme
              Chart = (fun id -> target id |> Option.bind (fun path -> Map.tryFind path parts) |> Option.bind (fun bytes -> Charts.parseChart theme (text bytes)))
              Picture = (fun id -> target id |> Option.map (fun path -> "blob:" + path)); Slicer = (fun _ _ _ _ _ -> None) }
        for item in sheet.Extras.Objects do
            if item.Xml.Contains "grpSp" || item.Kind = "chart" then
                match Charts.objectScene drawing live item.Xml 600.0 400.0, Charts.objectScene drawing noData item.Xml 600.0 400.0 with
                | Some fromCells, Some fromCache ->
                    charts <- charts + 1
                    if textsIn fromCells.Items |> List.exists (fun value -> value = "Grand Total") then withTotal <- sheet.Name :: withTotal
                    // (The same categories and as many axis labels: the values are the cells', the points the cache's.)
                    let months (scene: Charts.Scene) = textsIn scene.Items |> List.filter (fun value -> List.contains value [ "Jan"; "Feb"; "Mar"; "Apr"; "May"; "Jun"; "Jul"; "Aug"; "Sep"; "Oct"; "Nov"; "Dec"; "Grand Total" ])
                    if months fromCells <> months fromCache then differing <- (sheet.Name, months fromCells, months fromCache) :: differing
                | _ -> ()
    ok "pivot charts: drawn from the cells and from the cache" (charts >= 5) (sprintf "%d" charts)
    ok "pivot charts: no chart has the grand total among its points" withTotal.IsEmpty (sprintf "on %A" (List.distinct withTotal))
    ok "pivot charts: the categories are the cache's whether the sheets are converted or not" differing.IsEmpty (sprintf "%A" (List.truncate 2 differing))

// ---- a part of the newer kinds is told by its first tag -----------------------------------
// dashboard-sales-analysis.xlsx has a map chart of 8 MB (`cx:chartSpace`): reading all of it to
// find that it is not drawn took a tenth of a second of the workbook's opening. (It is a
// placeholder, read up to its series' name: stage55-chartex.)
do
    let _, _, parts, theme = openBook "real-world/dashboard-sales-analysis.xlsx"
    let chartEx = text parts.["xl/charts/chartEx1.xml"]
    let best =
        [ for _ in 1 .. 3 ->
            let watch = Diagnostics.Stopwatch.StartNew()
            let read = Charts.parseChart theme chartEx
            if not (read |> Option.exists (fun chart -> chart.Plots |> List.exists (fun plot -> match plot.Kind with Charts.UndrawnPlot "regionMap" -> true | _ -> false))) then failwith "a map read as something else than a placeholder"
            watch.Elapsed.TotalMilliseconds ]
        |> List.min
    ok "a chartEx part of 8 MB is not read to say what it is" (chartEx.Length > 8000000 && best < 10.0) (sprintf "%d characters, %.1f ms" chartEx.Length best)
    eq "told by its first tag" (Charts.isChartEx chartEx, Charts.isChartEx (text parts.["xl/charts/chart1.xml"])) (true, false)
    eq "... whatever its prefix"
        [ Charts.isChartEx "<?xml version=\"1.0\"?>\n<cx:chartSpace xmlns:a=\"x\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"><cx:chart/></cx:chartSpace>"
          Charts.isChartEx "<chartSpace xmlns=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"/>"
          Charts.isChartEx "<q:chartSpace xmlns:q='http://schemas.microsoft.com/office/drawing/2014/chartex'/>"
          Charts.isChartEx "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\"/>"
          Charts.isChartEx ""; Charts.isChartEx "no xml" ]
        [ true; true; true; false; false; false ]

// Manual layouts: a legend's and a plot's boxes are the chart's, not worked out.
let boxLayout (target: string) (x: float) (y: float) (w: float) (h: float) =
    sprintf "<c:layout><c:manualLayout>%s<c:xMode val=\"edge\"/><c:yMode val=\"edge\"/><c:x val=\"%g\"/><c:y val=\"%g\"/><c:w val=\"%g\"/><c:h val=\"%g\"/></c:manualLayout></c:layout>" target x y w h
let inPlotArea (layout: string) (xml: string) = xml.Replace("<c:plotArea>", "<c:plotArea>" + layout)
let manualBars = Charts.layoutChart Charts.defaultTheme noData (parsed (inPlotArea (boxLayout "<c:layoutTarget val=\"inner\"/>" 0.5 0.0 0.5 0.5) (chartXml (barPlot "clustered" (ser 0 "FF0000" [ 1.0; 2.0; 3.0; 4.0; 5.0 ] "")) axesXml))) 400.0 300.0
ok "an inner plot layout: every bar inside its box, none lost" (let bars = rects manualBars |> List.filter (fun (_, _, w, _, fill) -> fill = "#FF0000" && w > 12.0) in bars.Length = 5 && bars |> List.forall (fun (x, y, _, h, _) -> x >= 199.0 && y + h <= 151.0)) (sprintf "%A" (rects manualBars))

// The kinds that were not drawn.
let radarScene = Charts.layoutChart Charts.defaultTheme noData (parsed (chartXml ("<c:radarChart><c:radarStyle val=\"marker\"/>" + ser 0 "FF0000" [ 1.0; 2.0; 3.0 ] "" + "</c:radarChart>") "")) 300.0 300.0
ok "a radar chart is drawn: a closed line of the series' colour" ((Charts.sceneSvg radarScene).ToUpperInvariant().Contains "#FF0000") (Charts.sceneSvg radarScene)
let threeD = parsed (chartXml ("<c:bar3DChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>" + ser 0 "FF0000" [ 1.0; 2.0 ] "" + "</c:bar3DChart>") axesXml)
eq "a 3-D bar chart is drawn flat: its bars" (barsOf (Charts.layoutChart Charts.defaultTheme noData threeD 400.0 300.0) "#FF0000" |> List.length) 2

// Shapes.
let shapeParts : Charts.DrawingParts = { Theme = Charts.defaultTheme; Chart = (fun _ -> None); Picture = (fun _ -> None); Slicer = (fun _ _ _ _ _ -> None) }
let shapeOf (attrs: string) (xfrm: string) (geometry: string) (body: string) =
    sprintf "<xdr:twoCellAnchor xmlns:xdr=\"x\" xmlns:a=\"a\"><xdr:sp %s><xdr:spPr><a:xfrm %s><a:off x=\"0\" y=\"0\"/><a:ext cx=\"952500\" cy=\"476250\"/></a:xfrm>%s<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></xdr:spPr>%s</xdr:sp></xdr:twoCellAnchor>" attrs xfrm geometry body
let shapeSvg (data: Charts.ChartData) xml = Charts.objectScene shapeParts data xml 100.0 50.0 |> Option.map Charts.sceneSvg |> Option.defaultValue "(no scene)"
let prst (name: string) = sprintf "<a:prstGeom prst=\"%s\"><a:avLst/></a:prstGeom>" name
let has (svg: string) (part: string) = svg.Contains part
ok "a preset arrow is its outline, not a rectangle" (let svg = shapeSvg noData (shapeOf "" "" (prst "rightArrow") "") in not (has svg "<rect") && (has svg "<polyline" || has svg "<polygon" || has svg "<path")) (shapeSvg noData (shapeOf "" "" (prst "rightArrow") ""))
let custom = "<a:custGeom><a:pathLst><a:path w=\"100\" h=\"100\"><a:moveTo><a:pt x=\"0\" y=\"0\"/></a:moveTo><a:lnTo><a:pt x=\"100\" y=\"50\"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>"
ok "a custom geometry: its path, scaled to the shape" (has (shapeSvg noData (shapeOf "" "" custom "")) "M0 0L100 25Z") (shapeSvg noData (shapeOf "" "" custom ""))
ok "a turned shape turns about its centre" (has (shapeSvg noData (shapeOf "" "rot=\"5400000\"" (prst "rect") "")) "matrix(") (shapeSvg noData (shapeOf "" "rot=\"5400000\"" (prst "rect") ""))
let boxed (paragraph: string) = sprintf "<xdr:txBody><a:bodyPr/><a:p>%s</a:p></xdr:txBody>" paragraph
ok "a flipped shape's outline is mirrored, its text not" (let svg = shapeSvg noData (shapeOf "" "flipH=\"1\"" (prst "rightArrow") (boxed "<a:r><a:t>hi</a:t></a:r>")) in has svg "matrix(-1" && svg.IndexOf ">hi<" > svg.IndexOf "</g>") (shapeSvg noData (shapeOf "" "flipH=\"1\"" (prst "rightArrow") (boxed "<a:r><a:t>hi</a:t></a:r>")))
ok "a line break makes two lines" (let svg = shapeSvg noData (shapeOf "" "" (prst "rect") (boxed "<a:r><a:t>one</a:t></a:r><a:br/><a:r><a:t>two</a:t></a:r>")) in has svg "one" && has svg "two" && not (has svg "onetwo")) ""
let linkData : Charts.ChartData = { noData with Cells = (fun sheet row col _ _ -> if sheet = "KPI" && row = 16 && col = 1 then Some [| [| Some { Value = CellValue.Number 42.0; Format = defaultFormat; Comment = None } |] |] else None); Display = (fun _ -> "live 42") }
let linked = shapeOf "textlink=\"KPI!$B$17\"" "" (prst "rect") (boxed "<a:fld id=\"{1}\" type=\"TxLink\"><a:t>stale</a:t></a:fld>")
ok "a shape linked to a cell shows the cell" (let svg = shapeSvg linkData linked in has svg "live 42" && not (has svg "stale")) (shapeSvg linkData linked)
ok "... and the text of the file when the cell cannot be read" (has (shapeSvg noData linked) "stale") (shapeSvg noData linked)
ok "... which a plain field of no link keeps" (has (shapeSvg linkData (shapeOf "" "" (prst "rect") (boxed "<a:fld id=\"{1}\" type=\"TxLink\"><a:t>stale</a:t></a:fld>"))) "stale") ""

// ---- a chart part that says little (as openpyxl writes them): Excel 2019's defaults ----------
let bareChart =
    "<chartSpace xmlns=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><chart>"
    + "<title><tx><rich><a:bodyPr/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:t>Fleet total</a:t></a:r></a:p></rich></tx></title><plotArea><lineChart><grouping val=\"standard\"/>"
    + "<ser><idx val=\"0\"/><order val=\"0\"/><val><numRef><f>Nowhere!$A$1:$D$1</f><numCache><ptCount val=\"4\"/><pt idx=\"0\"><v>1</v></pt><pt idx=\"1\"><v>3</v></pt><pt idx=\"2\"><v>2</v></pt><pt idx=\"3\"><v>4</v></pt></numCache></numRef></val><smooth val=\"1\"/></ser>"
    + "<axId val=\"10\"/><axId val=\"100\"/></lineChart><catAx><axId val=\"10\"/><delete val=\"0\"/><axPos val=\"b\"/><crossAx val=\"100\"/></catAx>"
    + "<valAx><axId val=\"100\"/><delete val=\"0\"/><axPos val=\"l\"/><majorGridlines/><crossAx val=\"10\"/></valAx></plotArea><legend><legendPos val=\"t\"/></legend></chart></chartSpace>"
let bare = Charts.parseChart zooTheme bareChart |> Option.get
let bareStroke (stroke: Charts.Stroke option) = stroke |> Option.map (fun s -> match s.Paint with Charts.FillColor (color, _) -> color | _ -> "")
eq "no look of its own: a grey frame with rounded corners, the title and legend over the plot"
    (bareStroke bare.Border, bare.RoundedCorners, bare.TitleOverlay, bare.Legend |> Option.map (fun legend -> legend.Overlay)) (Some "#868686", true, true, Some true)
eq "... black grid lines and axis lines" (bare.Axes |> List.map (fun axis -> axis.Kind, bareStroke axis.Line, axis.Gridlines |> Option.map (fun g -> bareStroke (Some g)) |> Option.flatten)) [ "cat", Some "#000000", None; "val", Some "#000000", Some "#000000" ]
let bareScene = Charts.layoutChart zooTheme noData bare 480.0 288.0
let titleText = bareScene.Items |> List.tryPick (fun item -> match item with Charts.PText (_, _, "Fleet total", size, _, _, bold, _, _) -> Some (Math.Round size, bold) | _ -> None)
eq "... its title 10 pt bold" titleText (Some (Math.Round (10.0 * 96.0 / 72.0), true))
ok "... a smoothed line is a curve" (bareScene.Items |> List.exists (fun item -> match item with Charts.PPath (d, None, _, Some _, _) -> d.Contains " C " | _ -> false)) "no curve"

printfn "Stage 52 chart tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
