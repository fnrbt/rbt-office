#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Charts = Rbt.Spreadsheet.Charts
module Bridge = Rbt.Spreadsheet.XlsxBridge

// Charts of the newer kinds (`cx:chartSpace`, chartEx parts): a waterfall and a funnel are read
// - their data from the cells the defined names they name refer to (`_xlchart.v1.0`), else from
// the points the part caches - and drawn: a waterfall's bars from where the last one ended, the
// totals from zero, increases, decreases and totals in the theme's first three accents, lines
// between the bars; a funnel's bars centred, as long as their value is of the largest. A chart of
// another kind (a map, a sunburst, a treemap ...) is a placeholder that says what it is, with its
// title - read without reading the megabytes of a map's geography. The parts go back as they were.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-58s %s" label detail)
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let text (data: byte[]) = Text.Encoding.UTF8.GetString data
let noData : Charts.ChartData =
    { Cells = (fun _ _ _ _ _ -> None); Display = (fun _ -> ""); Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
let rects (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PRect (x, y, w, h, Some fill, _, _, _, _) -> Some (x, y, w, h, fill) | _ -> None)
let rec textsIn (items: Charts.Primitive list) : string list =
    items |> List.collect (fun item ->
        match item with
        | Charts.PText (_, _, value, _, _, _, _, _, _) -> [ value ]
        | Charts.PGroup (_, _, _, _, inner) -> textsIn inner
        | _ -> [])
let texts (scene: Charts.Scene) = textsIn scene.Items
let lines (scene: Charts.Scene) = scene.Items |> List.choose (fun item -> match item with Charts.PLine (x1, y1, x2, y2, color, _, _) -> Some (x1, y1, x2, y2, color) | _ -> None)
let close (a: float) (b: float) = abs (a - b) < 0.01

let space = "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\""
let level (format: string) (points: string list) =
    sprintf "<cx:lvl ptCount=\"%d\"%s>%s</cx:lvl>" points.Length format (points |> List.mapi (fun index point -> sprintf "<cx:pt idx=\"%d\">%s</cx:pt>" index point) |> String.concat "")
let chartEx (layoutId: string) (categories: string list) (values: string list) (series: string) (rest: string) =
    sprintf "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cx:chartSpace %s><cx:chartData><cx:data id=\"0\"><cx:strDim type=\"cat\"><cx:f>Sheet1!$A$2:$A$%d</cx:f>%s</cx:strDim><cx:numDim type=\"val\"><cx:f>Sheet1!$B$2:$B$%d</cx:f>%s</cx:numDim></cx:data></cx:chartData><cx:chart>%s</cx:chart></cx:chartSpace>"
        space (categories.Length + 1) (level "" categories) (values.Length + 1) (level " formatCode=\"General\"" values)
        (rest.Replace("{series}", sprintf "<cx:series layoutId=\"%s\" uniqueId=\"{00000000-0000-0000-0000-000000000001}\"><cx:tx><cx:txData><cx:f>Sheet1!$B$1</cx:f><cx:v>Amount</cx:v></cx:txData></cx:tx><cx:dataLabels pos=\"%s\"><cx:visibility seriesName=\"0\" categoryName=\"0\" value=\"1\"/></cx:dataLabels><cx:dataId val=\"0\"/>%s</cx:series>" layoutId (if layoutId = "funnel" then "ctr" else "outEnd") series))

// ---- a waterfall, from the points the part caches ----------------------------------------------------
let waterfallXml =
    chartEx "waterfall" [ "Start"; "Sales"; "Costs"; "Tax"; "End" ] [ "100"; "50"; "-30"; "-20"; "100" ]
        "<cx:layoutPr><cx:subtotals><cx:idx val=\"0\"/><cx:idx val=\"4\"/></cx:subtotals></cx:layoutPr>"
        ("<cx:title pos=\"t\" align=\"ctr\" overlay=\"0\"><cx:tx><cx:txData><cx:v>Profit bridge</cx:v></cx:txData></cx:tx></cx:title>"
         + "<cx:plotArea><cx:plotAreaRegion>{series}</cx:plotAreaRegion><cx:axis id=\"0\"><cx:catScaling gapWidth=\"0.5\"/><cx:tickLabels/></cx:axis>"
         + "<cx:axis id=\"1\"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis></cx:plotArea><cx:legend pos=\"t\" align=\"ctr\" overlay=\"0\"/>")
let waterfall = Charts.parseChart Charts.defaultTheme waterfallXml
ok "a waterfall part is read" waterfall.IsSome "not read"
match waterfall with
| Some chart ->
    let scene = Charts.layoutChart Charts.defaultTheme noData chart 600.0 400.0
    let shown = texts scene
    ok "waterfall: the title, the legend's three entries, the categories, the values"
        ([ "Profit bridge"; "Increase"; "Decrease"; "Total"; "Start"; "Sales"; "Costs"; "Tax"; "End"; "100"; "50"; "-30"; "-20" ] |> List.forall (fun item -> List.contains item shown)) (sprintf "%A" shown)
    // The bars, left to right (the legend's swatches are small squares above them).
    let bars = rects scene |> List.filter (fun (_, y, w, h, fill) -> fill <> "#FFFFFF" && y > 40.0) |> List.sortBy (fun (x, _, _, _, _) -> x)
    eq "waterfall: five bars: a total, an increase, two decreases, a total" (bars |> List.map (fun (_, _, _, _, fill) -> fill)) [ "#A5A5A5"; "#4472C4"; "#ED7D31"; "#ED7D31"; "#A5A5A5" ]
    match bars with
    | [ (_, startTop, _, startHeight, _); (_, salesTop, _, salesHeight, _); (_, costsTop, _, costsHeight, _); (_, taxTop, _, taxHeight, _); (_, endTop, _, endHeight, _) ] ->
        let bottomOf top height = top + height
        ok "waterfall: the totals stand on zero, as high as each other" (close startTop endTop && close (bottomOf startTop startHeight) (bottomOf endTop endHeight)) (sprintf "%g %g" startTop endTop)
        ok "waterfall: an increase starts where the total ended" (close (bottomOf salesTop salesHeight) startTop) (sprintf "%g / %g" (bottomOf salesTop salesHeight) startTop)
        ok "waterfall: a decrease hangs from where the increase ended" (close costsTop salesTop && close taxTop (bottomOf costsTop costsHeight)) (sprintf "%g %g" costsTop salesTop)
        ok "waterfall: the end total is where the decreases left it" (close (bottomOf taxTop taxHeight) endTop) (sprintf "%g / %g" (bottomOf taxTop taxHeight) endTop)
        ok "waterfall: half as much as a total is half as high" (close (salesHeight * 2.0) startHeight) (sprintf "%g / %g" salesHeight startHeight)
    | _ -> ok "waterfall: five bars" false (sprintf "%A" bars)
    // (Light grey, as the gridlines: Excel 2019's own picture of synthetic-chartex.xlsx.)
    eq "waterfall: a line joins each bar to the next, light grey" (lines scene |> List.filter (fun (x1, y1, x2, y2, color) -> color = "#D9D9D9" && y1 = y2 && x2 - x1 < 100.0) |> List.length) 4
| None -> ()

// ---- a funnel ---------------------------------------------------------------------------------------
let funnelXml =
    chartEx "funnel" [ "Visits"; "Carts"; "Orders" ] [ "500"; "300"; "120" ] ""
        "<cx:plotArea><cx:plotAreaRegion>{series}</cx:plotAreaRegion><cx:axis id=\"0\"><cx:catScaling gapWidth=\"0.06\"/><cx:tickLabels/></cx:axis></cx:plotArea>"
match Charts.parseChart Charts.defaultTheme funnelXml with
| Some chart ->
    let scene = Charts.layoutChart Charts.defaultTheme noData chart 600.0 300.0
    let bars = rects scene |> List.filter (fun (_, _, _, _, fill) -> fill = "#4472C4") |> List.sortBy (fun (_, y, _, _, _) -> y)
    eq "funnel: a bar a category, top to bottom" bars.Length 3
    ok "funnel: the bars are centred on each other" (bars |> List.map (fun (x, _, w, _, _) -> x + w / 2.0) |> List.pairwise |> List.forall (fun (a, b) -> close a b)) (sprintf "%A" bars)
    ok "funnel: as long as their value is of the largest" (match bars with [ (_, _, w0, _, _); (_, _, w1, _, _); (_, _, w2, _, _) ] -> close (w1 / w0) 0.6 && close (w2 / w0) 0.24 | _ -> false) (sprintf "%A" bars)
    ok "funnel: the categories and the values" ([ "Visits"; "Carts"; "Orders"; "500"; "300"; "120" ] |> List.forall (fun item -> List.contains item (texts scene))) (sprintf "%A" (texts scene))
| None -> ok "a funnel part is read" false "not read"

// ---- a kind that is not drawn: a placeholder with its title ----------------------------------------
let treemapXml =
    chartEx "treemap" [ "A"; "B" ] [ "1"; "2" ] ""
        ("<cx:title pos=\"t\" align=\"ctr\" overlay=\"0\"><cx:tx><cx:txData><cx:v>Stock by shelf</cx:v></cx:txData></cx:tx></cx:title>"
         + "<cx:plotArea><cx:plotAreaRegion>{series}</cx:plotAreaRegion></cx:plotArea>")
match Charts.parseChart Charts.defaultTheme treemapXml with
| Some chart ->
    let scene = Charts.layoutChart Charts.defaultTheme noData chart 400.0 300.0
    eq "a treemap: a box that says what it is, with its title" (texts scene) [ "Stock by shelf"; "Treemap chart (not drawn)" ]
| None -> ok "a treemap part is read as a placeholder" false "not read"

// dashboard-sales-analysis.xlsx has a map chart of 8 MB: it is a placeholder named for its series, read in no time.
let openBook (file: string) =
    let wb = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(root, file))) |> Option.get
    let parts = wb.Preserved.Parts |> Map.ofList
    let theme = wb.Preserved.ThemePath |> Option.bind (fun path -> Map.tryFind path parts) |> Option.map (text >> Charts.parseTheme) |> Option.defaultValue Charts.defaultTheme
    wb, parts, theme
do
    let _, parts, theme = openBook "real-world/dashboard-sales-analysis.xlsx"
    let map = text parts.["xl/charts/chartEx1.xml"]
    let best =
        [ for _ in 1 .. 3 ->
            let watch = Diagnostics.Stopwatch.StartNew()
            Charts.parseChart theme map |> ignore
            watch.Elapsed.TotalMilliseconds ]
        |> List.min
    ok "a map of 8 MB is read in no time" (map.Length > 8000000 && best < 10.0) (sprintf "%d characters, %.1f ms" map.Length best)
    match Charts.parseChart theme map with
    | Some chart -> eq "... a placeholder named for its series" (texts (Charts.layoutChart theme noData chart 400.0 300.0)) [ "Sum of Sales"; "Map chart (not drawn)" ]
    | None -> ok "a map part is read as a placeholder" false "not read"

// ---- the e-commerce dashboard's waterfalls: their data through `_xlchart` names, from the cells ----------
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries do
        if not (entry.FullName.EndsWith "/") then
            use stream = entry.Open()
            use buffer = new MemoryStream()
            stream.CopyTo buffer
            yield entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let ecommerceFile = "real-world/dashboard-ecommerce-sales.xlsx"
let ecommerce, ecommerceParts, ecommerceTheme = openBook ecommerceFile
let names (name: string) = ecommerce.NamedRanges |> List.tryPick (fun known -> if known.Name = name then Some known.RefersTo else None)
let book = Bridge.Import.importWorkbook ecommerce
let live : Charts.ChartData =
    let sheets = book.Sheets |> Array.toList |> List.map (fun sheet -> sheet.Name, resolveSheet sheet.Model)
    { Cells = (fun name r0 c0 r1 c1 ->
        sheets |> List.tryFind (fun (known, _) -> String.Equals(known, name, StringComparison.OrdinalIgnoreCase))
        |> Option.map (fun (_, model) -> Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } model))))
      Display = (fun cell -> Rbt.Spreadsheet.CellText.cellDisplayText None 1000.0 cell.Format cell.Value)
      Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
match Charts.parseChartWith names ecommerceTheme (text ecommerceParts.["xl/charts/chartEx1.xml"]) with
| Some chart ->
    let series = chart.Plots.Head.Series.Head
    eq "the waterfall's data is the cells its names refer to"
        (series.Categories |> Option.bind (fun item -> item.Formula), series.Values |> Option.bind (fun item -> item.Formula), chart.Plots.Head.Kind)
        (Some "'Waterfall Chart'!$D$4:$D$7", Some "'Waterfall Chart'!$E$4:$E$7", Charts.WaterfallPlot ([ 3 ], true))
    let scene = Charts.layoutChart ecommerceTheme live chart 600.0 400.0
    ok "... its categories are the cells'" ([ "Technology"; "Office Supplies"; "Furniture"; "Grand Total" ] |> List.forall (fun item -> List.contains item (texts scene))) (sprintf "%A" (texts scene))
    let bars = rects scene |> List.filter (fun (_, _, w, h, _) -> w > 30.0 && h < 380.0) |> List.sortBy (fun (x, _, _, _, _) -> x)
    ok "... four bars, the last (Grand Total) from zero to where the others ended"
        (match bars with
         | [ (_, firstTop, _, firstHeight, _); _; (_, thirdTop, _, _, _); (_, lastTop, _, lastHeight, _) ] -> close lastTop thirdTop && close (lastTop + lastHeight) (firstTop + firstHeight)
         | _ -> false) (sprintf "%A" bars)
    ok "... in the series' own gradient" (bars |> List.forall (fun (_, _, _, _, fill) -> fill.StartsWith "gradient(")) (sprintf "%A" (bars |> List.map (fun (_, _, _, _, fill) -> fill)))
| None -> ok "the e-commerce waterfall is read" false "not read"
// Through the drawing, as the sheet shows it: every chartEx anchor is drawn (the two maps as placeholders).
do
    let mutable drawn = []
    for sheet in ecommerce.Sheets do
        let target (id: string) = sheet.Extras.DrawingRelationships |> List.tryPick (fun (known, _, target, isExternal) -> if known = id && not isExternal then Some (ZipInterop.normalizePath target) else None)
        let drawing : Charts.DrawingParts =
            { Theme = ecommerceTheme
              Chart = (fun id -> target id |> Option.bind (fun path -> Map.tryFind path ecommerceParts) |> Option.bind (fun bytes -> Charts.parseChartWith names ecommerceTheme (text bytes)))
              Picture = (fun _ -> None); Slicer = (fun _ _ _ _ _ -> None) }
        for item in sheet.Extras.Objects do
            if item.Xml.Contains "chartex" then
                match Charts.objectScene drawing live item.Xml 600.0 400.0 with
                | Some scene -> drawn <- (sheet.Name, texts scene |> List.exists (fun value -> value.EndsWith "(not drawn)")) :: drawn
                | None -> drawn <- (sheet.Name + " (none)", false) :: drawn
    eq "every chartEx anchor of the dashboard is drawn" (drawn |> List.sort)
        [ "Dashboard", false; "Dashboard", true; "Map Chart", true; "Waterfall Chart", false ]

// ---- a save keeps the parts as they were, and the names they read through ----------------------------
let exported (sheets: SpreadsheetModel array) =
    let input : Bridge.ExportInput =
        { Sheets = sheets
          Names = book.Sheets |> Array.map (fun sheet -> sheet.Name)
          Visible = book.Sheets |> Array.map (fun sheet -> sheet.Visible)
          ActiveSheetIndex = 0
          NameInfo = Map.empty; DefinedNames = book.Context.Names
          Source = Some book.Source
          Macros = None }
    XlsxWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
let original = partsOf (File.ReadAllBytes(Path.Combine(root, ecommerceFile)))
let chartExParts = original |> Map.toList |> List.map fst |> List.filter (fun path -> path.StartsWith "xl/charts/" && (path.Contains "chartEx" || path.Contains "/colors" || path.Contains "/style"))
ok "the dashboard has chartEx parts, their relationships, styles and colours" (chartExParts.Length >= 12) (sprintf "%A" chartExParts)
let waterfallSheet = book.Sheets |> Array.findIndex (fun sheet -> sheet.Name = "Waterfall Chart")
let typed =
    book.Sheets |> Array.mapi (fun index sheet ->
        if index = waterfallSheet then setCell { Row = 3; Col = 4 } { Value = CellValue.Number 150000.0; Format = defaultFormat; Comment = None } sheet.Model
        else sheet.Model)
for label, saved in [ "as it was opened", exported (book.Sheets |> Array.map (fun sheet -> sheet.Model)); "with a cell of the waterfall's typed", exported typed ] do
    let parts = partsOf saved
    let changed = chartExParts |> List.filter (fun path -> Map.tryFind path parts <> Some original.[path])
    ok (sprintf "saved %s: every chartEx part byte for byte" label) changed.IsEmpty (sprintf "%A" changed)
    let types = text parts.["[Content_Types].xml"]
    ok (sprintf "saved %s: the parts keep their content type" label) ([ 1 .. 4 ] |> List.forall (fun n -> types.Contains (sprintf "/xl/charts/chartEx%d.xml" n))) "missing"
    let again = XlsxReader.loadBytesNative saved |> Option.get
    let namesAgain = again.NamedRanges |> List.filter (fun name -> name.Name.StartsWith "_xlchart") |> List.map (fun name -> name.Name, name.RefersTo) |> List.sort
    let namesBefore = ecommerce.NamedRanges |> List.filter (fun name -> name.Name.StartsWith "_xlchart") |> List.map (fun name -> name.Name, name.RefersTo) |> List.sort
    ok (sprintf "saved %s: the names the charts read through" label) (namesAgain = namesBefore && namesBefore.Length = 14) (sprintf "%A" namesAgain)
    let drawingsWithChartEx = parts |> Map.toList |> List.filter (fun (path, data) -> path.StartsWith "xl/drawings/drawing" && (text data).Contains "chartex")
    ok (sprintf "saved %s: the drawings still hold the charts" label) (drawingsWithChartEx.Length = 3) (sprintf "%A" (drawingsWithChartEx |> List.map fst))

printfn ""
printfn "Stage 55 chartEx tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
