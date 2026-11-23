/// Minimal Excel chart → PaintItem renderer (prototype). Parses a workbook's
/// first chart (xl/charts/chart1.xml) and its drawing anchor, and lays out a
/// clustered column chart the way Excel does, emitting the vector PaintItems
/// (VecPath/VecGroup/DrawText) added to the shared paint model.
module ChartRender

open System
open System.IO
open System.IO.Compression
open System.Xml.Linq
open Rbt.Spreadsheet
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Layout
open Rbt.Spreadsheet.Paint

// ---- parsing -------------------------------------------------------------

type ChartSeries =
    { Name: string option
      Values: float list
      Color: string }              // "#RRGGBB"

type ChartSpec =
    { Kind: string                 // "col" (barDir)
      Title: string option
      TitlePos: (float * float) option  // manualLayout (x,y) fractions, edge mode
      Categories: string list
      Series: ChartSeries list
      Gridlines: bool
      GapWidth: int
      DispUnit: float              // value-axis display unit (1000 = "Thousands")
      Font: string                 // theme minor font (chart text)
      Palette: string list }       // theme accent1..6
    /// First series' values / colour — convenience for single-series paths.
    member s.Values = match s.Series with x :: _ -> x.Values | [] -> []
    member s.SeriesColor = match s.Series with x :: _ -> x.Color | [] -> "#5B9BD5"

let private localName (e: XElement) = e.Name.LocalName
let private descendants (name: string) (e: XElement) =
    e.Descendants() |> Seq.filter (fun d -> d.Name.LocalName = name)
let private firstDesc name e = descendants name e |> Seq.tryHead
let private attr (name: string) (e: XElement) =
    match e.Attribute(XName.Get name) with null -> None | a -> Some a.Value
/// The `r:id` attribute lives in the relationships namespace.
let private relId (e: XElement) =
    match e.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")) with
    | null -> None
    | a -> Some a.Value

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

/// Category labels. A multi-level axis (multiLvlStrRef) has several <lvl>
/// blocks (e.g. months then years); use only the finest (first) level so the
/// labels don't concatenate.
let private catPoints (catEl: XElement) : string list =
    match descendants "lvl" catEl |> Seq.tryHead with
    | Some lvl -> points lvl
    | None -> points catEl

let private pf (s: string) = match Double.TryParse(s, Globalization.CultureInfo.InvariantCulture) with true, n -> n | _ -> 0.0

let private defaultAccents = [ "#5B9BD5"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#4472C4"; "#70AD47" ]

/// Theme minor-font latin typeface (charts use it via `+mn-lt`; cells via the
/// body font). Falls back to Calibri.
let private parseMinorFont (zip: ZipArchive) : string =
    match readEntry zip "xl/theme/theme1.xml" with
    | None -> "Calibri"
    | Some xml ->
        (XDocument.Parse xml).Root |> firstDesc "minorFont" |> Option.bind (firstDesc "latin")
        |> Option.bind (attr "typeface") |> Option.defaultValue "Calibri"

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
            let accents = parseTheme zip
            let valAttr name (e: XElement) = match e.Attribute(XName.Get name) with null -> None | a -> Some a.Value
            let allSer = descendants "ser" plot |> Seq.toList
            // Series colour: explicit spPr solidFill (rgb or theme accent), else
            // the palette by series index.
            let serColor i (ser: XElement) =
                let scheme (v: string) =
                    if v.StartsWith "accent" then
                        match System.Int32.TryParse(v.Substring 6) with
                        | true, n when n >= 1 && n <= List.length accents -> Some (List.item (n - 1) accents)
                        | _ -> None
                    else None
                descendants "solidFill" ser |> Seq.tryHead
                |> Option.bind (fun sf ->
                    match firstDesc "srgbClr" sf with
                    | Some c -> valAttr "val" c |> Option.map (fun v -> "#" + v)
                    | None -> firstDesc "schemeClr" sf |> Option.bind (valAttr "val") |> Option.bind scheme)
                |> Option.defaultValue (List.item (i % List.length accents) accents)
            let categories =
                allSer |> List.tryHead
                |> Option.bind (firstDesc (if kind = "scatter" then "xVal" else "cat"))
                |> Option.map (if kind = "scatter" then points else catPoints) |> Option.defaultValue []
            let series =
                allSer |> List.mapi (fun i ser ->
                    let name = firstDesc "tx" ser |> Option.bind (firstDesc "v") |> Option.map (fun v -> v.Value)
                    let vals =
                        (if kind = "scatter" then firstDesc "yVal" ser else firstDesc "val" ser)
                        |> Option.map points |> Option.defaultValue [] |> List.map pf
                    { Name = name; Values = vals; Color = serColor i ser })
            let titleEl = firstDesc "title" root
            let title =
                titleEl
                |> Option.bind (fun t -> descendants "t" t |> Seq.tryHead)
                |> Option.map (fun t -> t.Value)
            let titlePos =
                titleEl |> Option.bind (firstDesc "manualLayout")
                |> Option.bind (fun ml ->
                    match firstDesc "x" ml |> Option.bind (valAttr "val"), firstDesc "y" ml |> Option.bind (valAttr "val") with
                    | Some xs, Some ys -> Some (float xs, float ys)
                    | _ -> None)
            let gridlines = root |> descendants "valAx" |> Seq.exists (fun ax -> firstDesc "majorGridlines" ax |> Option.isSome)
            let gap = firstDesc "gapWidth" plot |> Option.map (fun e -> int (e.Attribute(XName.Get "val").Value)) |> Option.defaultValue 150
            let dispUnit =
                firstDesc "dispUnits" root |> Option.bind (firstDesc "builtInUnit") |> Option.bind (valAttr "val")
                |> Option.map (function
                    | "hundreds" -> 100.0 | "thousands" -> 1e3 | "tenThousands" -> 1e4 | "hundredThousands" -> 1e5
                    | "millions" -> 1e6 | "tenMillions" -> 1e7 | "hundredMillions" -> 1e8 | "billions" -> 1e9 | "trillions" -> 1e12 | _ -> 1.0)
                |> Option.defaultValue 1.0
            Some { Kind = kind; Title = title; TitlePos = titlePos; Categories = categories; Series = series
                   Gridlines = gridlines; GapWidth = gap; DispUnit = dispUnit; Font = parseMinorFont zip; Palette = accents }

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

/// A parsed chart together with its drawing-anchor placement (from/to cells +
/// EMU offsets). The runner derives the pixel rect from the live cell layout.
type ChartPlacement =
    { Spec: ChartSpec
      FromCol: int; FromColOffPx: float; FromRow: int; FromRowOffPx: float
      ToCol: int; ToColOffPx: float; ToRow: int; ToRowOffPx: float }

/// chart index N from a rels Target like "../charts/chart3.xml". Use the LAST
/// "chart" so the "charts/" directory segment doesn't match first.
let private chartNumFromTarget (t: string) =
    let i = t.LastIndexOf "chart"
    if i < 0 then None
    else
        let ds = t.Substring(i + 5) |> Seq.takeWhile Char.IsDigit |> Seq.toArray |> System.String
        if ds = "" then None else Some (int ds)

/// Parse every chart in the sheet's drawing, mapping each anchor to its chart
/// via the graphicFrame's `c:chart r:id` → drawing rels → chartN.xml (drawing1
/// also holds slicer/textbox anchors, so positional mapping is wrong).
let parseAllCharts (xlsxPath: string) : ChartPlacement list =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    match readEntry zip "xl/drawings/drawing1.xml", readEntry zip "xl/drawings/_rels/drawing1.xml.rels" with
    | Some xml, Some relsXml ->
        let relChart =
            descendants "Relationship" (XDocument.Parse relsXml).Root
            |> Seq.choose (fun r ->
                match attr "Id" r, attr "Target" r |> Option.bind chartNumFromTarget with
                | Some id, Some n -> Some (id, n)
                | _ -> None)
            |> Map.ofSeq
        descendants "twoCellAnchor" (XDocument.Parse xml).Root
        |> Seq.choose (fun a ->
            firstDesc "chart" a |> Option.bind relId |> Option.bind (fun rid -> Map.tryFind rid relChart)
            |> Option.bind (fun n -> parseChart xlsxPath n |> Option.map (fun spec -> a, spec)))
        |> Seq.map (fun (a, spec) ->
            let getf tag child =
                firstDesc tag a |> Option.bind (firstDesc child) |> Option.map (fun e -> float e.Value) |> Option.defaultValue 0.0
            { Spec = spec
              FromCol = int (getf "from" "col"); FromColOffPx = getf "from" "colOff" / 9525.0
              FromRow = int (getf "from" "row"); FromRowOffPx = getf "from" "rowOff" / 9525.0
              ToCol = int (getf "to" "col"); ToColOffPx = getf "to" "colOff" / 9525.0
              ToRow = int (getf "to" "row"); ToRowOffPx = getf "to" "rowOff" / 9525.0 })
        |> List.ofSeq
    | _ -> []

// ---- cell-reference lookup for pivot locations ---------------------------

let private colIndex0 (letters: string) =
    (letters |> Seq.fold (fun acc ch -> acc * 26 + (int (Char.ToUpper ch) - int 'A' + 1)) 0) - 1
/// "B38" -> (row0, col0).
let private parseRef (r: string) =
    let letters = r |> Seq.takeWhile Char.IsLetter |> Seq.toArray |> System.String
    let digits = r |> Seq.skipWhile Char.IsLetter |> Seq.toArray |> System.String
    (int digits - 1), colIndex0 letters

// ---- pivot tables (recompute from the cache) ------------------------------

type PivotPlacement = { Row: int; Col: int; Grid: CellValue[] [] }

/// Parse the workbook's pivot tables and RECOMPUTE each from its cache via the
/// Pivot engine (so imported pivots are live, not just cached cells). Reads the
/// single pivotCache (definition + records), resolves shared items, then for
/// each pivotTableN.xml reads the row/col/data field indices + location and
/// calls Pivot.fromCache. (Convention naming: pivotCacheDefinition1/Records1.)
let parsePivots (xlsxPath: string) : PivotPlacement list =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    let cacheFields =
        match readEntry zip "xl/pivotCache/pivotCacheDefinition1.xml" with
        | None -> [||]
        | Some xml ->
            descendants "cacheField" (XDocument.Parse xml).Root
            |> Seq.map (fun cf ->
                let name = attr "name" cf |> Option.defaultValue ""
                let shared =
                    match firstDesc "sharedItems" cf with
                    | Some si ->
                        si.Elements()
                        |> Seq.choose (fun it ->
                            match it.Name.LocalName, attr "v" it with
                            | "s", Some v -> Some(CellValue.String v)
                            | "n", Some v -> Some(CellValue.Number(pf v))
                            | "b", Some v -> Some(CellValue.Boolean(v = "1"))
                            | "d", Some v -> Some(CellValue.String v)
                            | _ -> None)
                        |> Array.ofSeq
                    | None -> [||]
                name, shared)
            |> Array.ofSeq
    let headers = cacheFields |> Array.map fst
    let sharedOf fi = if fi < cacheFields.Length then snd cacheFields.[fi] else [||]
    let records =
        match readEntry zip "xl/pivotCache/pivotCacheRecords1.xml" with
        | None -> [||]
        | Some xml ->
            descendants "r" (XDocument.Parse xml).Root
            |> Seq.map (fun r ->
                r.Elements()
                |> Seq.mapi (fun fi it ->
                    match it.Name.LocalName, attr "v" it with
                    | "x", Some i -> (let a = sharedOf fi in if int i >= 0 && int i < a.Length then a.[int i] else CellValue.Empty)
                    | "n", Some v -> CellValue.Number(pf v)
                    | "s", Some v -> CellValue.String v
                    | "b", Some v -> CellValue.Boolean(v = "1")
                    | _ -> CellValue.Empty)
                |> Array.ofSeq)
            |> Array.ofSeq
    [ for n in 1 .. 16 do
        match readEntry zip (sprintf "xl/pivotTables/pivotTable%d.xml" n) with
        | None -> ()
        | Some xml ->
            let root = (XDocument.Parse xml).Root
            let loc = firstDesc "location" root |> Option.bind (attr "ref") |> Option.defaultValue "A1"
            let r0, c0 = parseRef ((loc.Split(':')).[0])
            let fieldsOf tag =
                firstDesc tag root
                |> Option.map (fun e -> descendants "field" e |> Seq.choose (attr "x") |> Seq.map int |> Seq.filter (fun i -> i >= 0) |> List.ofSeq)
                |> Option.defaultValue []
            let dataF = firstDesc "dataFields" root |> Option.bind (firstDesc "dataField")
            let dataIdx = dataF |> Option.bind (attr "fld") |> Option.map int |> Option.defaultValue 0
            let agg = dataF |> Option.bind (attr "subtotal") |> Option.bind Rbt.Spreadsheet.Pivot.parseAgg |> Option.defaultValue Rbt.Spreadsheet.Pivot.Sum
            let rowIdxs = fieldsOf "rowFields"
            if not (List.isEmpty rowIdxs) && records.Length > 0 then
                yield { Row = r0; Col = c0; Grid = Rbt.Spreadsheet.Pivot.fromCache headers records rowIdxs (fieldsOf "colFields") dataIdx agg } ]

// ---- slicers --------------------------------------------------------------

type SlicerItem = { Label: string; Selected: bool }
type SlicerSpec =
    { Caption: string
      Items: SlicerItem list
      ColumnCount: int
      Accent: string               // theme accent1 — selected-button fill
      Font: string                 // theme minor font
      FromCol: int; FromColOffPx: float; FromRow: int; FromRowOffPx: float
      ToCol: int; ToColOffPx: float; ToRow: int; ToRowOffPx: float }

/// pivotCache field name -> ordered shared-item labels (for slicer item text).
let private pivotSharedItems (zip: ZipArchive) : Map<string, string list> =
    match readEntry zip "xl/pivotCache/pivotCacheDefinition1.xml" with
    | None -> Map.empty
    | Some xml ->
        descendants "cacheField" (XDocument.Parse xml).Root
        |> Seq.choose (fun cf ->
            attr "name" cf
            |> Option.map (fun nm ->
                let items =
                    firstDesc "sharedItems" cf
                    |> Option.map (fun si -> descendants "s" si |> Seq.choose (attr "v") |> List.ofSeq)
                    |> Option.defaultValue []
                nm, items))
        |> Map.ofSeq

/// Parse slicers on the workbook (slicer1.xml + slicerCaches + drawing anchor).
let parseSlicers (xlsxPath: string) : SlicerSpec list =
    use fs = File.OpenRead xlsxPath
    use zip = new ZipArchive(fs, ZipArchiveMode.Read)
    let accent = parseTheme zip |> List.tryHead |> Option.defaultValue "#5B9BD5"
    let font = parseMinorFont zip
    let shared = pivotSharedItems zip
    // cache name -> (sourceField, [(itemIndex, selected)])
    let caches =
        [ for n in 1 .. 12 do
            match readEntry zip (sprintf "xl/slicerCaches/slicerCache%d.xml" n) with
            | Some xml ->
                let root = (XDocument.Parse xml).Root
                let items =
                    descendants "i" root
                    |> Seq.choose (fun i -> attr "x" i |> Option.map (fun x -> int x, attr "s" i <> Some "0"))
                    |> List.ofSeq
                match attr "name" root with
                | Some c -> yield c, (attr "sourceName" root |> Option.defaultValue "", items)
                | None -> ()
            | None -> () ]
        |> Map.ofList
    // drawing anchor (cell rect) for a slicer by name
    let anchorOf (name: string) =
        match readEntry zip "xl/drawings/drawing1.xml" with
        | None -> None
        | Some xml ->
            descendants "twoCellAnchor" (XDocument.Parse xml).Root
            |> Seq.tryFind (fun a -> descendants "slicer" a |> Seq.exists (fun sl -> attr "name" sl = Some name))
            |> Option.map (fun a ->
                let g tag child = firstDesc tag a |> Option.bind (firstDesc child) |> Option.map (fun e -> float e.Value) |> Option.defaultValue 0.0
                int (g "from" "col"), g "from" "colOff" / 9525.0, int (g "from" "row"), g "from" "rowOff" / 9525.0,
                int (g "to" "col"), g "to" "colOff" / 9525.0, int (g "to" "row"), g "to" "rowOff" / 9525.0)
    match readEntry zip "xl/slicers/slicer1.xml" with
    | None -> []
    | Some xml ->
        [ for sl in descendants "slicer" (XDocument.Parse xml).Root do
            match attr "name" sl, attr "cache" sl with
            | Some name, Some cache ->
                let caption = attr "caption" sl |> Option.defaultValue name
                let sourceName, idxItems = Map.tryFind cache caches |> Option.defaultValue ("", [])
                let labels = Map.tryFind sourceName shared |> Option.defaultValue []
                let items =
                    idxItems |> List.choose (fun (x, sel) ->
                        List.tryItem x labels |> Option.map (fun lbl -> { Label = lbl; Selected = sel }))
                match anchorOf name with
                | Some (fc, fco, fr, fro, tc, tco, tr, tro) ->
                    yield { Caption = caption; Items = items; ColumnCount = (attr "columnCount" sl |> Option.map int |> Option.defaultValue 1)
                            Accent = accent; Font = font
                            FromCol = fc; FromColOffPx = fco; FromRow = fr; FromRowOffPx = fro
                            ToCol = tc; ToColOffPx = tco; ToRow = tr; ToRowOffPx = tro }
                | None -> ()
            | _ -> () ]

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

/// Largest value across every series (drives the value axis).
let private dataMaxOf (spec: ChartSpec) =
    match spec.Series |> List.collect (fun s -> s.Values) with [] -> 1.0 | vs -> List.max vs

/// Font for chart/slicer text — set per render from the workbook theme so we
/// use Excel's actual font (e.g. Garamond) rather than a hardcoded Calibri.
let mutable private chartTextFont = "Calibri"

let private text (s: string) (rect: Rect) (size: float) (color: string) (align: TextAlign) (vAlign: VerticalAlign) : PaintItem =
    DrawText { Text = s; Rect = rect; Color = color; FontSize = size; FontName = Some chartTextFont
               Bold = false; Italic = false; PrintItalicEligible = false; Underline = false; UnderlineDouble = false; Strikethrough = false
               Align = align; VerticalAlign = vAlign; LineIndex = 0; LineCount = 1; BaselineShift = 0.0; Baseline = 0.0; Rotation = 0; Clip = None }

let private axisColor = "#595959"
let private gridColor = "#D9D9D9"
/// Axis-label number format — thousands separators, never scientific notation
/// (Excel shows "3,000,000", not "3e+06"; display-unit scaling is a TODO).
let private fmtNum (v: float) : string =
    let ci = System.Globalization.CultureInfo.InvariantCulture
    if v = 0.0 then "0"
    elif abs v >= 1000.0 then v.ToString("#,0", ci)
    else v.ToString("0.###", ci)
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
// titleH, leftAxisW, bottomAxisH, rightPad. A manual-layout (left-anchored)
// title is small and sits near the top, so it reserves a thinner band than a
// centred title.
let private insets (spec: ChartSpec) =
    let titleH = match spec.Title, spec.TitlePos with Some _, Some _ -> 30.0 | Some _, None -> 49.0 | None, _ -> 10.0
    titleH, 32.0, 31.0, 28.0

/// Chart-area background, border, and title (left-anchored when the chart
/// carries a manualLayout, else centred).
let private frame (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    [ yield FillRect ({ X = 0.0; Y = 0.0; Width = w; Height = h }, "#FFFFFF")
      yield StrokeRect ({ X = 0.5; Y = 0.5; Width = w - 1.0; Height = h - 1.0 }, gridColor, 1.0)
      match spec.Title with
      | Some t ->
          match spec.TitlePos with
          | Some (tx, ty) -> yield text t { X = tx * w + 2.0; Y = ty * h; Width = w - tx * w - 6.0; Height = titleSize + 4.0 } titleSize "#595959" AlignLeft VAlignTop
          | None -> yield text t { X = 0.0; Y = 15.0; Width = w; Height = titleSize + 4.0 } titleSize "#595959" AlignCenter VAlignTop
      | None -> () ]

let private yAxisLabel (dispUnit: float) (v: float) (y: float) (leftAxisW: float) : PaintItem =
    text (fmtNum (v / dispUnit)) { X = 0.0; Y = y - 8.0; Width = leftAxisW - 4.0; Height = 14.0 } 10.0 axisColor AlignRight VAlignMiddle

/// Clustered column chart (one bar per series within each category band).
let buildColumnChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let n = spec.Series |> List.map (fun s -> List.length s.Values) |> (fun ls -> if List.isEmpty ls then 0 else List.max ls)
    let nSer = max 1 (List.length spec.Series)
    let dataMax = dataMaxOf spec
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
          yield yAxisLabel spec.DispUnit v y leftAxisW
      if n > 0 then
          let bandW = plotW / float n
          let groupW = bandW / (1.0 + float spec.GapWidth / 100.0)
          let barW = groupW / float nSer
          for i in 0 .. n - 1 do
              let groupLeft = plotX + float i * bandW + (bandW - groupW) / 2.0
              for si in 0 .. List.length spec.Series - 1 do
                  let s = spec.Series.[si]
                  match List.tryItem i s.Values with
                  | Some v ->
                      let barX = groupLeft + float si * barW
                      let topY = yToPx v
                      yield VecPath ([ VMove (barX, baselineY); VLine (barX, topY); VLine (barX + barW, topY); VLine (barX + barW, baselineY); VClose ], Some s.Color, None)
                  | None -> ()
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = plotX + float i * bandW; Y = baselineY + 2.0; Width = bandW; Height = bottomAxisH } 10.0 axisColor AlignCenter VAlignTop ]

/// Clustered bar chart (horizontal): value axis along the bottom, categories
/// up the left with category 1 at the bottom.
let buildBarChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let n = spec.Series |> List.map (fun s -> List.length s.Values) |> (fun ls -> if List.isEmpty ls then 0 else List.max ls)
    let nSer = max 1 (List.length spec.Series)
    let dataMax = dataMaxOf spec
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
          yield text (fmtNum (v / spec.DispUnit)) { X = x - 20.0; Y = baselineY + 2.0; Width = 40.0; Height = 14.0 } 10.0 axisColor AlignCenter VAlignTop
      if n > 0 then
          let bandH = plotH / float n
          let groupH = bandH / (1.0 + float spec.GapWidth / 100.0)
          let barH = groupH / float nSer
          for i in 0 .. n - 1 do
              let groupTop = titleH + plotH - float (i + 1) * bandH + (bandH - groupH) / 2.0   // category 1 at the bottom
              for si in 0 .. List.length spec.Series - 1 do
                  match List.tryItem i spec.Series.[si].Values with
                  | Some v ->
                      let barY = groupTop + float si * barH
                      let endX = xToPx v
                      yield VecPath ([ VMove (plotX, barY); VLine (endX, barY); VLine (endX, barY + barH); VLine (plotX, barY + barH); VClose ], Some spec.Series.[si].Color, None)
                  | None -> ()
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = 0.0; Y = titleH + plotH - float (i + 1) * bandH + bandH / 2.0 - 7.0; Width = leftAxisW - 4.0; Height = 14.0 } 10.0 axisColor AlignRight VAlignMiddle ]

/// Value-axis chrome (gridlines + y labels) + category x positions, shared by
/// line and area. Returns (items, xOfCategory, yToPx, baselineY, plotX, plotW).
let private valueAxisChrome (edgeMode: bool) (spec: ChartSpec) (w: float) (h: float) =
    let n = spec.Series |> List.map (fun s -> List.length s.Values) |> (fun ls -> if List.isEmpty ls then 0 else List.max ls)
    let dataMax = dataMaxOf spec
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
              yield yAxisLabel spec.DispUnit v y leftAxisW
          for i in 0 .. n - 1 do
              let cat = spec.Categories |> List.tryItem i |> Option.defaultValue ""
              yield text cat { X = xOf i - bandW / 2.0; Y = baselineY + 2.0; Width = bandW; Height = bandW } 10.0 axisColor AlignCenter VAlignTop ]
    chrome, xOf, yToPx, baselineY

/// Line chart — one poly-line per series (markers only for small series, like
/// Excel's marker default).
let buildLineChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let chrome, xOf, yToPx, _ = valueAxisChrome false spec w h
    [ yield! chrome
      for s in spec.Series do
          let pts = s.Values |> List.mapi (fun i v -> xOf i, yToPx v)
          match pts with
          | [] -> ()
          | (x0, y0) :: rest ->
              yield VecPath (VMove (x0, y0) :: (rest |> List.map (fun (x, y) -> VLine (x, y))), None, Some (s.Color, 2.0))
              if List.length pts <= 12 then
                  for (x, y) in pts do
                      yield VecPath (circle x y 3.5, Some s.Color, None) ]

/// Area chart — one filled area per series (overlaid).
let buildAreaChart (spec: ChartSpec) (w: float) (h: float) : PaintItem list =
    let chrome, xOf, yToPx, baselineY = valueAxisChrome true spec w h
    [ yield! chrome
      for s in spec.Series do
          let pts = s.Values |> List.mapi (fun i v -> xOf i, yToPx v)
          match pts with
          | [] -> ()
          | (x0, _) :: _ ->
              let xn, _ = List.last pts
              let area = VMove (x0, baselineY) :: (pts |> List.map (fun (x, y) -> VLine (x, y))) @ [ VLine (xn, baselineY); VClose ]
              yield VecPath (area, Some s.Color, None) ]

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
          yield yAxisLabel spec.DispUnit v y leftAxisW
      for i in 0 .. int (round (xMax / xStep)) do
          let v = float i * xStep
          let x = xToPx v
          // Scatter has both horizontal and vertical gridlines.
          yield VecPath ([ VMove (x, titleH); VLine (x, baselineY) ], None, Some (gridColor, 1.0))
          yield text (fmtNum (v / spec.DispUnit)) { X = x - 20.0; Y = baselineY + 2.0; Width = 40.0; Height = 14.0 } 10.0 axisColor AlignCenter VAlignTop
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
    chartTextFont <- spec.Font
    match spec.Kind with
    | "bar" -> buildBarChart spec w h
    | "line" -> buildLineChart spec w h
    | "area" -> buildAreaChart spec w h
    | "scatter" -> buildScatterChart spec w h
    | "pie" -> buildPieChart spec w h
    | _ -> buildColumnChart spec w h

/// Render a slicer: outer box, caption header, and a grid of item buttons
/// (selected = white, filtered-out = grey).
let buildSlicer (spec: SlicerSpec) (rect: Rect) : PaintItem list =
    chartTextFont <- spec.Font
    let border = "#D9D9D9"
    // Excel's "Light" slicer style tints the accent toward white for selected
    // buttons; approximate with a 45% lighten.
    let selFill =
        let h = spec.Accent.TrimStart('#')
        if h.Length <> 6 then spec.Accent
        else
            let ch i = System.Convert.ToInt32(h.Substring(i, 2), 16)
            let mix v = int (float v + (255.0 - float v) * 0.45)
            sprintf "#%02X%02X%02X" (mix (ch 0)) (mix (ch 2)) (mix (ch 4))
    let headerH = 21.0
    let pad = 2.0
    let cols = max 1 spec.ColumnCount
    let nItems = List.length spec.Items
    let rows = max 1 ((nItems + cols - 1) / cols)
    let bodyTop = rect.Y + headerH
    let cellW = (rect.Width - pad) / float cols
    let cellH = (rect.Height - headerH) / float rows
    [ yield FillRect ({ X = rect.X; Y = rect.Y; Width = rect.Width; Height = rect.Height }, "#FFFFFF")
      yield StrokeRect ({ X = rect.X + 0.5; Y = rect.Y + 0.5; Width = rect.Width - 1.0; Height = rect.Height - 1.0 }, border, 1.0)
      // Caption (bold) + divider under the header.
      yield DrawText { Text = spec.Caption; Rect = { X = rect.X + 5.0; Y = rect.Y; Width = rect.Width - 10.0; Height = headerH }
                       Color = "#595959"; FontSize = 11.0; FontName = Some spec.Font; Bold = true; Italic = false; PrintItalicEligible = false
                       Underline = false; UnderlineDouble = false; Strikethrough = false
                       Align = AlignLeft; VerticalAlign = VAlignMiddle; LineIndex = 0; LineCount = 1; BaselineShift = 0.0; Baseline = 0.0; Rotation = 0; Clip = None }
      yield VecPath ([ VMove (rect.X, bodyTop); VLine (rect.X + rect.Width, bodyTop) ], None, Some (border, 1.0))
      for i in 0 .. nItems - 1 do
          let it = spec.Items.[i]
          let c = i % cols
          let r = i / cols
          let bx = rect.X + pad + float c * cellW
          let by = bodyTop + float r * cellH
          let fill = if it.Selected then selFill else "#F2F2F2"
          let txtColor = if it.Selected then "#000000" else "#A6A6A6"
          yield FillRect ({ X = bx + 1.0; Y = by + 1.0; Width = cellW - 2.0; Height = cellH - 2.0 }, fill)
          yield StrokeRect ({ X = bx + 1.0; Y = by + 1.0; Width = cellW - 2.0; Height = cellH - 2.0 }, border, 1.0)
          yield text it.Label { X = bx + 6.0; Y = by; Width = cellW - 8.0; Height = cellH } 10.0 txtColor AlignLeft VAlignMiddle ]
