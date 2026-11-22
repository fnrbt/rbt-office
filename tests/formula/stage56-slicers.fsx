#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open System
open System.IO
open System.IO.Compression
open Rbt.Spreadsheet.Model
module Charts = Rbt.Spreadsheet.Charts
module Slicers = Rbt.Spreadsheet.Slicers
module Bridge = Rbt.Spreadsheet.XlsxBridge

// Slicers and timelines: read from their parts (the slicer, its cache, the items of the pivot cache's
// field or the values of the table's column; the timeline and its cache's selection and bounds), drawn
// where the drawing anchors them - a slicer as its caption and a button an item, in its columns and
// button height, each button in the look its state has in the slicer's style (a custom one of the
// workbook, or a built-in one in the workbook's theme); a timeline as its caption, its selection in words,
// its level and its strip of periods - and saved as they were, with a cell typed or not.
// synthetic-slicers.xlsx is synthetic (tools/excel-parity/make-slicers-fixture.py); the dashboards are
// real workbooks with pivot slicers.

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-58s %s" label detail)
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let text (data: byte[]) = Text.Encoding.UTF8.GetString data
let rec flat (items: Charts.Primitive list) : Charts.Primitive list =
    items |> List.collect (fun item ->
        match item with
        | Charts.PGroup (x, y, _, _, inner) ->
            flat inner |> List.map (fun primitive ->
                match primitive with
                | Charts.PRect (a, b, w, h, fill, o, stroke, sw, r) -> Charts.PRect (a + x, b + y, w, h, fill, o, stroke, sw, r)
                | Charts.PText (a, b, value, size, color, anchor, bold, italic, turn) -> Charts.PText (a + x, b + y, value, size, color, anchor, bold, italic, turn)
                | other -> other)
        | Charts.PFace (_, inner) -> flat inner
        | other -> [ other ])
let texts (items: Charts.Primitive list) = flat items |> List.choose (fun item -> match item with Charts.PText (_, _, value, _, _, _, _, _, _) -> Some value | _ -> None)
/// The buttons drawn: (label, x, y, fill, text colour, bold), the label at the rectangle's left, 8 px in
/// (as Excel 2019 draws it), centred on its height.
let buttons (items: Charts.Primitive list) =
    let all = flat items
    all |> List.choose (fun item ->
        match item with
        | Charts.PText (x, y, label, _, color, "start", bold, _, _) ->
            all |> List.tryPick (fun other ->
                match other with
                | Charts.PRect (rx, ry, w, h, fill, _, _, _, _) when abs (rx + 8.0 - x) < 0.01 && abs (ry + h / 2.0 - y) < 0.01 && w > 16.0 -> Some (label, rx, ry, fill, color, bold)
                | _ -> None)
        | _ -> None)
let noData : Charts.ChartData =
    { Cells = (fun _ _ _ _ _ -> None); Display = (fun _ -> ""); Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }

let openBook (file: string) =
    let bytes = File.ReadAllBytes(Path.Combine(root, file))
    let wb = XlsxReader.loadBytesNative bytes |> Option.get
    bytes, wb, Bridge.Import.importWorkbook wb
let liveData (book: Bridge.ImportedWorkbook) : Charts.ChartData =
    let sheets = book.Sheets |> Array.toList |> List.map (fun sheet -> sheet.Name, resolveSheet sheet.Model)
    { Cells = (fun name r0 c0 r1 c1 ->
        sheets |> List.tryFind (fun (known, _) -> String.Equals(known, name, StringComparison.OrdinalIgnoreCase))
        |> Option.map (fun (_, model) -> Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } model))))
      Display = (fun cell -> Rbt.Spreadsheet.CellText.cellDisplayText None 1000.0 cell.Format cell.Value)
      Format = (fun code value -> Rbt.Spreadsheet.CellText.formatNumberWithCodeIn false code value) }
let tint (color: string) (amount: float) = "#" + XlsxReader.Xlsx.applyColorTint color amount

// ---- the synthetic workbook: what is read -------------------------------------------------------------
let synthetic = "test-documents-excel/synthetic-slicers.xlsx"
let synBytes, syn, synBook = openBook synthetic
let parts = XlsxReader.SlicerParts.read syn
eq "every slicer and the timeline are read, by name" (parts.Slicers |> Map.toList |> List.map fst, parts.Timelines |> Map.toList |> List.map fst) ([ "Product 1"; "Product 2"; "Region" ], [ "Date" ])
let region = parts.Slicers.["Region"]
eq "Region: its caption, columns, button height (EMU), style (the default)" (region.Caption, region.ColumnCount, region.RowHeight, region.Style) ("Region", 2, 241300.0, None)
let item label selected hasData : XlsxReader.SlicerItem = { Label = label; Selected = selected; HasData = hasData }
eq "Region: its buttons from the pivot cache's field, selected or not, with data or not"
    region.Source (XlsxReader.CachedItems [ item "North" true true; item "South" true true; item "East" false true; item "West" true false ])
eq "Product 1 is the table's: the column's cells, the values its filter lets through"
    parts.Slicers.["Product 1"].Source (XlsxReader.TableColumn ("Data", 1, 12, 2, Some [ "Apples"; "Pears" ]))
let date = parts.Timelines.["Date"]
eq "the timeline: caption, level (months), style, where it is scrolled to, selection, bounds"
    (date.Caption, date.Level, date.Style, date.ScrollPosition, date.Selection, date.Bounds)
    ("Date", 2, Some "Timeline Custom", Some (2012, 1, 1), Some ((2012, 3, 1), (2012, 12, 31)), Some ((2012, 1, 1), (2014, 1, 1)))
eq "the workbook's slicer and timeline styles: the defaults, the custom ones and their elements"
    (syn.SlicerStyles.DefaultSlicer, syn.SlicerStyles.Slicers |> List.map (fun style -> style.Name, style.Elements |> List.map fst),
     syn.SlicerStyles.DefaultTimeline, syn.SlicerStyles.Timelines |> List.map (fun style -> style.Name, style.Elements |> List.map fst))
    (Some "SlicerStyleLight1", [ "Slicer Custom", [ "unselectedItemWithData"; "unselectedItemWithNoData"; "selectedItemWithData"; "selectedItemWithNoData" ] ],
     Some "TimeSlicerStyleLight1", [ "Timeline Custom", [ "selectedTimeBlock" ] ])
// A timeline style's elements point into <x15:dxfs>, a list of their own, not into the slicer styles' <x14:dxfs>.
eq "the custom timeline's selected periods: its <x15:dxfs> format (accent 6), the custom slicer's selected buttons theirs (orange)"
    (syn.SlicerStyles.Timelines |> List.collect (fun style -> style.Elements) |> List.map (fun (_, dxf) -> dxf.Fill),
     syn.SlicerStyles.Slicers |> List.collect (fun style -> style.Elements) |> List.filter (fst >> (=) "selectedItemWithData") |> List.map (fun (_, dxf) -> dxf.Fill))
    ([ Some "70AD47" ], [ Some "FFC000" ])

// ---- ... and drawn, through the drawing's anchors ------------------------------------------------------
let drawer = Slicers.drawer syn
let synData = liveData synBook
let partsFor (wb: XlsxReader.Workbook) (draw: string -> string -> Charts.ChartData -> float -> float -> Charts.Primitive list option) : Charts.DrawingParts =
    { Theme = Charts.defaultTheme; Chart = (fun _ -> None); Picture = (fun _ -> None); Slicer = draw; Camera = (fun _ _ _ -> None) }
let anchorsOf (wb: XlsxReader.Workbook) =
    [ for sheet in wb.Sheets do
        for item in sheet.Extras.Objects do
            yield sheet.Name, item.Xml ]
let sceneNamed (wb: XlsxReader.Workbook) draw data (name: string) width height =
    anchorsOf wb
    |> List.tryFind (fun (_, xml) ->
        let rec named (node: XmlDsl.XmlNode) =
            match node with
            | XmlDsl.Element (tag, attrs, children, _) -> ((tag.EndsWith "slicer" || tag.EndsWith "timeslicer") && Map.tryFind "name" attrs = Some name) || List.exists named children
            | _ -> false
        XmlDsl.tryScan xml (Collections.Generic.HashSet<string>()) |> Option.exists named)
    |> Option.bind (fun (_, xml) -> Charts.objectScene (partsFor wb draw) data xml width height)
let theme = syn.ThemeColors
let accent1 = theme.[4]
match sceneNamed syn drawer synData "Region" 240.0 200.0 with
| Some scene ->
    let shown = buttons scene.Items
    eq "Region: a button an item, in two columns, in the cache's order (no data last)" (shown |> List.map (fun (label, _, _, _, _, _) -> label)) [ "North"; "South"; "East"; "West" ]
    match shown with
    | [ (_, x0, y0, _, _, _); (_, x1, y1, _, _, _); (_, x2, y2, _, _, _); (_, x3, y3, _, _, _) ] ->
        ok "... North and South side by side, East and West under them" (y0 = y1 && x0 < x1 && y2 = y3 && y2 > y0 && x2 = x0 && x3 = x1) (sprintf "%A" shown)
        ok "... as high as the slicer says in whole pixels (241300 EMU: 25.33 px, 25 as Excel draws it) with a gap of 3" (abs (y2 - y0 - 28.0) < 0.01) (sprintf "%g" (y2 - y0))
        // (Excel 2019 at 100%: the buttons 9 px in from the frame at the left and 9.67 at the right, on whole pixels,
        // 3 px apart - synthetic-slicer-looks.xlsx, synthetic-slicer-cut.xlsx, the e-commerce dashboard -, the first row
        // 32 px below the top, the header's line 27 px below it.)
        let widthOf label = flat scene.Items |> List.pick (fun item -> match item with Charts.PRect (x, _, w, _, Some _, _, _, _, _) when x = (if label = 0 then x0 else x1) && w > 16.0 -> Some w | _ -> None)
        eq "... where Excel puts them: 9 px in from the frame at the left, 9.67 at the right, 3 px apart, 32 px from the top" (x0, x1, x1 + widthOf 1, x0 + widthOf 0 + 3.0, y0) (9.0, 121.0, 230.0, 121.0, 32.0)
        ok "... the header's line 27 px down, from 9 px in to the buttons' right"
            (flat scene.Items |> List.exists (fun item -> match item with Charts.PLine (9.0, 27.5, 230.0, 27.5, _, _, _) -> true | _ -> false)) (sprintf "%A" (flat scene.Items))
    | _ -> ok "Region: four buttons" false (sprintf "%A" shown)
    eq "... in SlicerStyleLight1 in the theme: selected with data in accent 1 tint 60%, unselected white, selected without data tint 80% in grey text"
        (shown |> List.map (fun (_, _, _, fill, color, _) -> fill, color))
        [ Some (tint accent1 0.59999389629810485), "#000000"; Some (tint accent1 0.59999389629810485), "#000000"; Some "#FFFFFF", "#000000"; Some (tint accent1 0.79998168889431442), "#828282" ]
    ok "... under its caption" (texts scene.Items |> List.contains "Region") (sprintf "%A" (texts scene.Items))
| None -> ok "the Region slicer is drawn" false "a placeholder"
match sceneNamed syn drawer synData "Product 2" 300.0 120.0 with
| Some scene ->
    let shown = buttons scene.Items
    eq "Product 2: three buttons in a row, in the custom style (orange, bold)"
        (shown |> List.map (fun (label, _, y, fill, color, bold) -> label, fill, color, bold))
        [ "Apples", Some "#FFC000", "#000000", true; "Pears", Some "#FFC000", "#000000", true; "Plums", Some "#FFC000", "#000000", true ]
    ok "... its header in the style's tableStyle: bold dark red" (flat scene.Items |> List.exists (fun item -> match item with Charts.PText (_, _, "Product (pivot)", _, "#C00000", "start", true, _, _) -> true | _ -> false)) (sprintf "%A" scene.Items)
| None -> ok "the Product 2 slicer is drawn" false "a placeholder"
match sceneNamed syn drawer synData "Product 1" 200.0 260.0 with
| Some scene ->
    eq "Product 1 (the table's): the column's values, each once, sorted, Plums not selected"
        (buttons scene.Items |> List.map (fun (label, _, _, fill, _, _) -> label, fill))
        [ "Apples", Some (tint accent1 0.59999389629810485); "Pears", Some (tint accent1 0.59999389629810485); "Plums", Some "#FFFFFF" ]
| None -> ok "the Product 1 slicer is drawn" false "a placeholder"
match sceneNamed syn drawer synData "Date" 600.0 140.0 with
| Some scene ->
    let shown = texts scene.Items
    ok "the timeline: caption, selection in words, level, years and months"
        ([ "Date"; "Mar - Dec 2012"; "MONTHS"; "2012"; "JAN"; "FEB"; "MAR"; "DEC" ] |> List.forall (fun value -> List.contains value shown)) (sprintf "%A" shown)
    let blocks =
        flat scene.Items |> List.choose (fun item -> match item with Charts.PRect (x, y, w, h, fill, _, _, _, _) when h = 17.0 -> Some (x, y, fill) | _ -> None) |> List.sortBy (fun (x, _, _) -> x)
    // (Excel 2019: months 1.6 px a day from 7 px in - Jan 7, Feb 56, Mar 103 ... Dec 543; the custom style has only
    // selectedTimeBlock, so the others have no fill.)
    eq "... a block a month from January 2012 where Excel puts it: March to December in the custom style's green, the others without a fill"
        (blocks |> List.map (fun (_, _, fill) -> fill), blocks |> List.map (fun (x, _, _) -> x))
        ([ None; None ] @ List.replicate 10 (Some ("#" + theme.[9])), [ 7.0; 56.0; 103.0; 152.0; 200.0; 250.0; 298.0; 347.0; 397.0; 445.0; 495.0; 543.0 ])
    let labelAt (value: string) = flat scene.Items |> List.tryPick (fun item -> match item with Charts.PText (x, y, known, _, _, anchor, _, _, _) when known = value -> Some (x, y, anchor) | _ -> None)
    eq "... a month's label at its block's left, above it (as Excel), the year above that"
        (match labelAt "MAR", labelAt "2012", blocks |> List.tryItem 2 with
         | Some (x, y, anchor), Some (_, yearY, _), Some (blockX, blockY, _) -> Some (x - blockX, anchor, y < blockY, yearY < y)
         | _ -> None)
        (Some (1.0, "start", true, true))
| None -> ok "the timeline is drawn" false "a placeholder"
// How Excel 2019 words a timeline's selection at each level (make-timeline-variants.py: synthetic-slicers.xlsx's
// timeline in Excel, its label read off the picture).
for level, first, last, expected in
        [ 2, (2012, 3, 1), (2012, 12, 31), "Mar - Dec 2012"; 2, (2012, 11, 1), (2013, 2, 28), "Nov 2012 - Feb 2013"; 2, (2012, 6, 1), (2012, 6, 30), "Jun 2012"
          1, (2012, 4, 1), (2012, 9, 30), "Q2 - Q3 2012"; 1, (2012, 10, 1), (2013, 3, 31), "Q4 2012 - Q1 2013"; 1, (2012, 4, 1), (2012, 6, 30), "Q2 2012"
          3, (2012, 3, 5), (2012, 3, 20), "Mar 5 - 20 2012"; 3, (2012, 3, 28), (2012, 4, 3), "Mar 28 - Apr 3, 2012"
          3, (2012, 12, 30), (2013, 1, 2), "Dec 30, 2012 - Jan 2, 2013"; 3, (2012, 3, 5), (2012, 3, 5), "Mar 5, 2012"
          0, (2012, 1, 1), (2013, 12, 31), "2012 - 2013"; 0, (2012, 1, 1), (2012, 12, 31), "2012" ] do
    eq (sprintf "a timeline's selection in words: %s" expected) (Slicers.selectionText { date with Level = level; Selection = Some (first, last) }) expected

// ---- slicers Excel made (tools/excel-parity/build-slicer-looks.vbs: synthetic-slicer-looks.xlsx) --------------
let _, looks, looksBook = openBook "test-documents-excel/synthetic-slicer-looks.xlsx"
let looksDrawer = Slicers.drawer looks
let looksData = liveData looksBook
let labels name = match sceneNamed looks looksDrawer looksData name 193.0 254.0 with Some scene -> buttons scene.Items |> List.map (fun (label, _, _, _, _, _) -> label) | None -> []
eq "sorted up: the items with data first, each part sorted (Excel's order)" (labels "RegionUp") [ "East"; "North"; "South"; "Central"; "West" ]
eq "sorted down: likewise, down" (labels "RegionDown") [ "South"; "North"; "East"; "West"; "Central" ]
eq "no data in place: every item in its sorted place" (labels "RegionInPlace") [ "Central"; "East"; "North"; "South"; "West" ]
eq "hide items with no data (x15:slicerCacheHideItemsWithNoData in the cache's extLst): only those with data" (labels "RegionHide") [ "East"; "North"; "South" ]
let headerRed name =
    match sceneNamed looks looksDrawer looksData name 193.0 254.0 with
    | Some scene -> flat scene.Items |> List.exists (fun item -> match item with Charts.PLine (_, _, _, _, "#D65532", _, _) -> true | _ -> false)
    | None -> false
eq "the header's Clear Filter: a red cross while the slicer filters, grey while every item is selected"
    ([ "ProductS"; "RegionUp"; "RegionDown"; "ProductAll" ] |> List.map headerRed) [ true; true; false; false ]

eq "a name the workbook has no slicer of: None (the frame says what it is)" (drawer "slicer" "Nothing" noData 100.0 100.0) None

// ---- where Excel cuts a slicer's text (tools/excel-parity/build-slicer-cut.vbs: synthetic-slicer-cut.xlsx, Calibri,
// and synthetic-slicer-cut-twcen.xlsx, the theme's body font Tw Cen MT; one cache shown by 400 slicers 1 px apart, read
// off Excel 2019's pictures): a button's text where it is no wider than the button less 18 px, else its most characters
// and "..." narrower than that; the caption likewise in the slicer less 72 px; in the theme's body font, by its GDI widths.
for file, face, expected in
        [ "test-documents-excel/synthetic-slicer-cut.xlsx", "Calibri",
          [ "B1_40", ([ ""; ""; "" ], "")
            "B1_41", ([ "."; "."; "." ], "")
            "B1_45", ([ ".."; ".."; ".." ], "")
            "B1_49", ([ "..."; "..."; "..." ], "")
            "B1_56", ([ "0..."; "..."; "..." ], "")
            "B1_70", ([ "000..."; "Ab ..."; "Ce..." ], "")
            "B1_80", ([ "0000..."; "Ab C..."; "Cent..." ], "C")
            "B1_81", ([ "0000..."; "Ab C..."; "Central" ], "C")
            "B2_90", ([ "00000..."; "Ab Cd ..."; "Central" ], "C")
            "B2_91", ([ "000000..."; "Ab Cd ..."; "Central" ], "C")
            "B2_92", ([ "00000000"; "Ab Cd ..."; "Central" ], "C")
            "B2_109", ([ "00000000"; "Ab Cd Efg..."; "Central" ], "C")
            "B2_110", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "C")
            "C1_118", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "0000...")
            "C1_125", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "00000...")
            "C1_126", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "000000...")
            "C1_127", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "00000000") ]
          "test-documents-excel/synthetic-slicer-cut-twcen.xlsx", "Tw Cen MT",
          [ "B1_78", ([ "0000..."; "Ab C..."; "Cent..." ], ".")
            "B1_79", ([ "0000..."; "Ab C..."; "Centr..." ], ".")
            "B1_80", ([ "0000..."; "Ab C..."; "Central" ], "C")
            "C1_127", ([ "00000000"; "Ab Cd Efghij"; "Central" ], "00000...") ] ] do
    let _, wb, book = openBook file
    let draw = Slicers.drawer wb
    let data = liveData book
    for name, (labels, caption) in expected do
        let width = float (int (name.Substring 3)) + 1.0
        match sceneNamed wb draw data name width 126.0 with
        | Some scene ->
            let all = flat scene.Items
            let shown = buttons scene.Items |> List.map (fun (label, _, _, _, _, _) -> label)
            let header = all |> List.tryPick (fun item -> match item with Charts.PText (9.0, 15.0, value, _, _, "start", true, _, _) -> Some value | _ -> None)
            eq (sprintf "%s %s: its buttons and caption as Excel cuts them" (Path.GetFileName file) name) (shown, header) (labels, Some caption)
            let rec faces (items: Charts.Primitive list) =
                items |> List.collect (fun item -> match item with Charts.PFace (known, inner) -> known :: faces inner | Charts.PGroup (_, _, _, _, inner) -> faces inner | _ -> [])
            eq (sprintf "... in %s" face) (faces scene.Items) (if face = "Calibri" then [] else [ face ])
        | None -> ok (sprintf "%s is drawn" name) false "a placeholder"
eq "a slicer's text in its room (GDI's widths: Tw Cen MT 'Central' 44 px, 'Centr...' 42)"
    ([ 44.0; 43.0; 42.0; 41.0 ] |> List.map (fun room -> Slicers.slicerFitted "Tw Cen MT" false room "Central")) [ "Central"; "Centr..."; "Cent..."; "Cent..." ]

// ---- the dashboards' slicers (pivot slicers, a custom style copied from a built-in one) -----------------
for file, expected in [ "real-world/dashboard-ecommerce-sales.xlsx", [ "Region"; "Region 1"; "Segment"; "Ship Mode"; "Ship Mode 1"; "Year" ]
                        "real-world/dashboard-sales-analysis.xlsx", [ "Category 1"; "Financial Year"; "State 1" ] ] do
    let _, wb, book = openBook file
    let read = XlsxReader.SlicerParts.read wb
    eq (sprintf "%s: its slicers are read" (Path.GetFileName file)) (read.Slicers |> Map.toList |> List.map fst) expected
    let draw = Slicers.drawer wb
    let data = liveData book
    let frames = anchorsOf wb |> List.filter (fun (_, xml) -> xml.Contains "sle:slicer")
    let drawn = frames |> List.choose (fun (_, xml) -> Charts.objectScene (partsFor wb draw) data xml 300.0 300.0)
    ok (sprintf "%s: every slicer frame of its drawings is drawn, none a placeholder" (Path.GetFileName file)) (drawn.Length = frames.Length && frames.Length > 0) (sprintf "%d of %d" drawn.Length frames.Length)
do
    let _, wb, book = openBook "real-world/dashboard-ecommerce-sales.xlsx"
    let read = XlsxReader.SlicerParts.read wb
    eq "the Year slicer's buttons: the pivot cache's years in the cache's order, all selected"
        read.Slicers.["Year"].Source (XlsxReader.CachedItems [ for year in 2011 .. 2014 -> item (string year) true true ])
    let regionOne = read.Slicers.["Region 1"]
    eq "Region 1: caption, four columns, its custom style" (regionOne.Caption, regionOne.ColumnCount, regionOne.Style) ("Select Region", 4, Some "SlicerStyleLight2 2")
    match sceneNamed wb (Slicers.drawer wb) (liveData book) "Region 1" 330.0 64.0 with
    | Some scene ->
        let shown = buttons scene.Items
        eq "... four buttons in one row, in the style's selected look (accent 2 tint 60%)"
            (shown |> List.map (fun (label, _, y, fill, _, _) -> label, fill))
            [ for label in [ "Central"; "East"; "South"; "West" ] -> label, Some (tint wb.ThemeColors.[5] 0.59999389629810485) ]
        ok "... side by side" (shown |> List.map (fun (_, _, y, _, _, _) -> y) |> List.distinct |> List.length = 1) (sprintf "%A" shown)
        ok "... under the caption, bold (the style's header)" (flat scene.Items |> List.exists (fun item -> match item with Charts.PText (_, _, "Select Region", _, _, "start", true, _, _) -> true | _ -> false)) (sprintf "%A" (texts scene.Items))
    | None -> ok "Region 1 is drawn" false "a placeholder"
    // (Excel 2019's picture of the Dashboard: Region 1's buttons 61 px wide, 3 apart, "Central" in Tw Cen MT cut.)
    match sceneNamed wb (Slicers.drawer wb) (liveData book) "Region 1" 272.0 64.0 with
    | Some scene ->
        let shown = buttons scene.Items
        eq "... at its size in Excel (271 px): buttons 61 px wide, 3 px apart, \"Central\" cut to \"Centr...\" in Tw Cen MT as Excel shows it"
            (shown |> List.map (fun (label, x, _, _, _, _) -> label, x)) [ "Centr...", 9.0; "East", 73.0; "South", 137.0; "West", 201.0 ]
    | None -> ok "Region 1 is drawn" false "a placeholder"

// ---- saved: every slicer and timeline part, their relationships and content types, the extension list entries ----
let partsOf (bytes: byte[]) : Map<string, byte[]> =
    use archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read)
    [ for entry in archive.Entries do
        if not (entry.FullName.EndsWith "/") then
            use stream = entry.Open()
            use buffer = new MemoryStream()
            stream.CopyTo buffer
            yield entry.FullName, buffer.ToArray() ]
    |> Map.ofList
let scan (xml: string) = XmlDsl.tryScan xml (Collections.Generic.HashSet<string>()) |> Option.get
let rec elements (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    match node with
    | XmlDsl.Element (_, _, children, _) -> node :: (children |> List.collect elements)
    | _ -> []
let localName (node: XmlDsl.XmlNode) = match node with XmlDsl.Element (name, _, _, _) -> (name.Split(':') |> Array.last) | _ -> ""
/// The extension entries of a part's own extension list that hold slicers, timelines or their styles, as written.
let slicerExtensions (xml: string) : string list =
    let rootNode = scan xml
    let kinds = set [ "slicerList"; "timelineRefs"; "slicerCaches"; "timelineCacheRefs"; "slicerStyles"; "timelineStyles"; "dxfs" ]
    match rootNode with
    | XmlDsl.Element (_, _, children, _) ->
        children
        |> List.filter (fun child -> localName child = "extLst")
        |> List.collect (fun extLst -> match extLst with XmlDsl.Element (_, _, exts, _) -> exts | _ -> [])
        |> List.filter (fun ext -> elements ext |> List.exists (fun inner -> kinds.Contains (localName inner)))
        |> List.map (fun ext -> match ext with XmlDsl.Element (_, _, _, span) -> xml.Substring(span.Start, span.End - span.Start) | _ -> "")
    | _ -> []
/// The relationships of a part to slicers, timelines and their caches: (type, target).
let slicerRelationships (relsXml: string) =
    elements (scan relsXml)
    |> List.choose (fun node ->
        match node with
        | XmlDsl.Element (_, attrs, _, _) when localName node = "Relationship" ->
            let kind = attrs.["Type"]
            if kind.EndsWith "/slicer" || kind.EndsWith "/slicerCache" || kind.EndsWith "/timeline" || kind.EndsWith "/timelineCache" then Some (kind, attrs.["Target"]) else None
        | _ -> None)
    |> List.sort
let exported (book: Bridge.ImportedWorkbook) (sheets: SpreadsheetModel array) =
    let input : Bridge.ExportInput =
        { Sheets = sheets
          Names = book.Sheets |> Array.map (fun sheet -> sheet.Name)
          Visible = book.Sheets |> Array.map (fun sheet -> sheet.Visible)
          ActiveSheetIndex = 0
          NameInfo = Map.empty; DefinedNames = book.Context.Names
          Source = Some book.Source
          Macros = None; MacroProject = None }
    XlsxPackageWriter.writeWorkbookPackageBytes (Bridge.Export.toWorkbookOut input)
let isSlicerPart (path: string) = path.StartsWith "xl/slicers/" || path.StartsWith "xl/slicerCaches/" || path.StartsWith "xl/timelines/" || path.StartsWith "xl/timelineCaches/"
for file, typedOn in [ synthetic, [ "Data"; "Report" ]
                       "real-world/dashboard-ecommerce-sales.xlsx", [ "Dashboard" ]
                       "real-world/dashboard-sales-analysis.xlsx", [ "Dashboard" ] ] do
    let bytes, _, book = openBook file
    let original = partsOf bytes
    let watched = original |> Map.toList |> List.map fst |> List.filter isSlicerPart
    let name = Path.GetFileName file
    let relsOf (all: Map<string, byte[]>) =
        all |> Map.toList |> List.filter (fun (path, _) -> path.EndsWith ".rels") |> List.collect (fun (path, data) -> slicerRelationships (text data) |> List.map (fun rel -> path, rel)) |> List.sort
    let extensionsOf (all: Map<string, byte[]>) =
        all |> Map.toList |> List.filter (fun (path, _) -> path = "xl/workbook.xml" || path = "xl/styles.xml" || path.StartsWith "xl/worksheets/sheet")
        |> List.collect (fun (path, data) -> slicerExtensions (text data) |> List.map (fun ext -> path, ext)) |> List.sort
    let typesOf (all: Map<string, byte[]>) =
        elements (scan (text all.["[Content_Types].xml"]))
        |> List.choose (fun node -> match node with XmlDsl.Element (_, attrs, _, _) when localName node = "Override" && isSlicerPart (attrs.["PartName"].TrimStart '/') -> Some (attrs.["PartName"], attrs.["ContentType"]) | _ -> None)
        |> List.sort
    let frames (all: Map<string, byte[]>) =
        all |> Map.toList |> List.filter (fun (path, _) -> path.StartsWith "xl/drawings/drawing") |> List.sumBy (fun (_, data) -> elements (scan (text data)) |> List.filter (fun node -> localName node = "slicer" || localName node = "timeslicer") |> List.length)
    ok (sprintf "%s has slicer parts, relationships, extension entries" name) (watched.Length >= 4 && not (relsOf original).IsEmpty && not (extensionsOf original).IsEmpty) (sprintf "%A" watched)
    let edits =
        ("as opened", book.Sheets |> Array.map (fun sheet -> sheet.Model))
        :: [ for sheetName in typedOn ->
                let index = book.Sheets |> Array.findIndex (fun sheet -> sheet.Name = sheetName)
                sprintf "with a cell typed on %s" sheetName,
                book.Sheets |> Array.mapi (fun at sheet -> if at = index then setCell { Row = 40; Col = 20 } { Value = CellValue.Number 7.0; Format = defaultFormat; Comment = None } sheet.Model else sheet.Model) ]
    for label, sheets in edits do
        let saved = partsOf (exported book sheets)
        let changed = watched |> List.filter (fun path -> Map.tryFind path saved <> Some original.[path])
        ok (sprintf "%s saved %s: every slicer and timeline part byte for byte" name label) changed.IsEmpty (sprintf "%A" changed)
        eq (sprintf "%s saved %s: their relationships" name label) (relsOf saved) (relsOf original)
        eq (sprintf "%s saved %s: their content types" name label) (typesOf saved) (typesOf original)
        eq (sprintf "%s saved %s: the extension entries, as written" name label) (extensionsOf saved) (extensionsOf original)
        eq (sprintf "%s saved %s: the drawings' frames" name label) (frames saved) (frames original)

printfn ""
printfn "Stage 56 slicer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
