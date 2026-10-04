#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let cell value = XlsxWriter.simpleCell value
let formulaCell formula value = XlsxWriter.cellWithFormula formula value

let smallSheet: XlsxWriter.FormattedSheet =
    { Name = "Data"
      Rows =
        [ { Cells = [ cell "A & <x> \"q\" 's'"; cell "42" ] }
          { Cells = [ formulaCell "SUM(A1:B1)" "42" ] } ]
      MergedCells = [ XlsxWriter.createMergedRange 0 0 1 1 ] }

let xml = XlsxWriter.generateFormattedSheetXml smallSheet
ok "xlsx-escapes-inline-text"
   (xml.Contains("A &amp; &lt;x&gt; &quot;q&quot; &apos;s&apos;"))
   xml
ok "xlsx-formula-cell"
   (xml.Contains("<f>SUM(A1:B1)</f><v>42</v>"))
   xml
ok "xlsx-merge-cell"
   (xml.Contains("<mergeCell ref=\"A1:B2\"/>"))
   xml

let minimalXml = XlsxWriter.generateSheetXml [ [ "A & <x> \"q\" 's'"; "42" ] ]
ok "xlsx-minimal-escapes-inline-text"
   (minimalXml.Contains("A &amp; &lt;x&gt; &quot;q&quot; &apos;s&apos;"))
   minimalXml
ok "xlsx-minimal-number-cell"
   (minimalXml.Contains("<c r=\"B1\"><v>42</v></c>"))
   minimalXml

let workbookXml = XlsxWriter.generateWorkbookXml [ "One & <x>"; "Two" ]
ok "xlsx-workbook-escapes-sheet-name"
   (workbookXml.Contains("name=\"One &amp; &lt;x&gt;\"") && workbookXml.IndexOf("One &amp;") < workbookXml.IndexOf("Two"))
   workbookXml

let workbookRels = XlsxWriter.generateWorkbookRels 2
ok "xlsx-workbook-rels"
   (workbookRels.Contains("Id=\"rId2\"") && workbookRels.Contains("Target=\"worksheets/sheet2.xml\""))
   workbookRels

let contentTypes = XlsxWriter.generateContentTypesXlsx 2
ok "xlsx-content-types"
   (contentTypes.Contains("PartName=\"/xl/worksheets/sheet2.xml\""))
   contentTypes

let perfSheet: XlsxWriter.FormattedSheet =
    { Name = "Perf"
      Rows =
        [ for r in 1 .. 2000 ->
            { Cells =
                [ for c in 1 .. 20 ->
                    if c % 3 = 0 then cell (string (r * c))
                    else cell (sprintf "R%d C%d & <text>" r c) ] } ]
      MergedCells = [] }

XlsxWriter.generateFormattedSheetXml perfSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let generated = XlsxWriter.generateFormattedSheetXml perfSheet
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "xlsx-sheet-linear-alloc"
   (generated.Length > 1000000 && allocatedMb < 96.0)
   (sprintf "chars=%d allocated=%.1f MB" generated.Length allocatedMb)

let minimalPerfData =
    [ for r in 1 .. 2000 ->
        [ for c in 1 .. 20 ->
            if c % 3 = 0 then string (r * c)
            else sprintf "R%d C%d & <text>" r c ] ]

XlsxWriter.generateSheetXml minimalPerfData |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let beforeMinimal = GC.GetAllocatedBytesForCurrentThread()
let generatedMinimal = XlsxWriter.generateSheetXml minimalPerfData
let allocatedMinimalMb = float (GC.GetAllocatedBytesForCurrentThread() - beforeMinimal) / 1024.0 / 1024.0
ok "xlsx-minimal-sheet-linear-alloc"
   (generatedMinimal.Length > 1000000 && allocatedMinimalMb < 96.0)
   (sprintf "chars=%d allocated=%.1f MB" generatedMinimal.Length allocatedMinimalMb)

let manySheets = [ for i in 1 .. 5000 -> sprintf "Sheet %d & <x>" i ]
XlsxWriter.generateWorkbookXml manySheets |> ignore
XlsxWriter.generateWorkbookRels manySheets.Length |> ignore
XlsxWriter.generateContentTypesXlsx manySheets.Length |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let beforeMetadata = GC.GetAllocatedBytesForCurrentThread()
let metadataChars =
    (XlsxWriter.generateWorkbookXml manySheets).Length
    + (XlsxWriter.generateWorkbookRels manySheets.Length).Length
    + (XlsxWriter.generateContentTypesXlsx manySheets.Length).Length
let allocatedMetadataMb = float (GC.GetAllocatedBytesForCurrentThread() - beforeMetadata) / 1024.0 / 1024.0
ok "xlsx-metadata-linear-alloc"
   (metadataChars > 1000000 && allocatedMetadataMb < 48.0)
   (sprintf "chars=%d allocated=%.1f MB" metadataChars allocatedMetadataMb)

let loadWorkbook path =
    File.ReadAllBytes path |> XlsxReader.loadBytesNative

let roundTripWorkbook (wb: XlsxReader.Workbook) =
    let bytes = XlsxWriter.packageFromReaderWorkbook wb |> XlsxWriter.writeWorkbookPackageBytes
    XlsxReader.loadBytesNative bytes

let cellMap (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
    sheet.Rows
    |> List.collect (fun row -> row.Cells |> List.map (fun c -> sprintf "%s%d" c.Reference.Column c.Reference.Row, c))
    |> Map.ofList

let resolvedValueText (wb: XlsxReader.Workbook) (cell: XlsxReader.Cell) =
    match XlsxReader.resolveValue wb cell.Value with
    | XlsxReader.Number n -> sprintf "n:%g" n
    | XlsxReader.String s -> "s:" + s
    | XlsxReader.Boolean b -> sprintf "b:%b" b
    | XlsxReader.Error e -> "e:" + e
    | XlsxReader.Formula f -> "f:" + f
    | XlsxReader.Empty -> ""
    | other -> sprintf "%A" other

let styleSig (cell: XlsxReader.Cell) =
    cell.Style
    |> Option.map (fun s ->
        s.NumberFormat, s.FontName, s.FontSize, s.Bold, s.Italic, s.Underline, s.FillColor,
        s.BorderTop |> Option.map (fun b -> b.Style, b.Color),
        s.BorderBottom |> Option.map (fun b -> b.Style, b.Color),
        s.BorderLeft |> Option.map (fun b -> b.Style, b.Color),
        s.BorderRight |> Option.map (fun b -> b.Style, b.Color))
    |> Option.defaultValue (None, None, None, false, false, false, None, None, None, None, None)

let mergeSig (sheet: XlsxReader.Worksheet) =
    sheet.MergedCells
    |> List.map (fun m -> sprintf "%s%d:%s%d" m.StartRef.Column m.StartRef.Row m.EndRef.Column m.EndRef.Row)
    |> List.sort

let columnSig (sheet: XlsxReader.Worksheet) =
    sheet.Columns
    |> List.filter (fun c -> c.Hidden || c.Width.IsSome)
    |> List.collect (fun c -> [ for i in c.Index .. c.MaxIndex -> i, (if c.Hidden then None else c.Width), c.Hidden ])
    |> List.sort

let rowSig (sheet: XlsxReader.Worksheet) =
    sheet.Rows |> List.choose (fun r -> r.Height |> Option.map (fun h -> r.Index, h)) |> List.sort

let tableSig (sheet: XlsxReader.Worksheet) =
    sheet.Tables |> List.map (fun t -> t.Name, t.StartRow, t.StartCol, t.EndRow, t.EndCol, t.Columns, t.HeaderRowCount, t.TotalsRowCount, t.AutoFilter) |> List.sort

let pageSig (sheet: XlsxReader.Worksheet) =
    sheet.PageSetup, sheet.PageMargins, sheet.PrintOptions, sheet.HeaderFooter, sheet.RowBreaks, sheet.ColBreaks, sheet.PrintArea, sheet.PrintTitleRows, sheet.PrintTitleCols

let assertRoundTripFixture path =
    match loadWorkbook path with
    | None -> ok ("load " + Path.GetFileName path) false "reader returned None"
    | Some wb ->
        match roundTripWorkbook wb with
        | None -> ok ("roundtrip-load " + Path.GetFileName path) false "saved reader returned None"
        | Some rt ->
            ok ("rt-sheet-count " + Path.GetFileName path) (wb.Sheets.Length = rt.Sheets.Length) (sprintf "%d <> %d" wb.Sheets.Length rt.Sheets.Length)
            ok ("rt-sheet-names " + Path.GetFileName path) ((wb.Sheets |> List.map (fun s -> s.Name)) = (rt.Sheets |> List.map (fun s -> s.Name))) ""
            for s1, s2 in List.zip wb.Sheets rt.Sheets do
                let before = cellMap wb s1
                let after = cellMap rt s2
                let mismatch =
                    before
                    |> Map.tryPick (fun ref c1 ->
                        after
                        |> Map.tryFind ref
                        |> Option.bind (fun c2 ->
                            if resolvedValueText wb c1 = resolvedValueText rt c2 && styleSig c1 = styleSig c2 then None
                            else Some (sprintf "%s before value=%s style=%A after value=%s style=%A" ref (resolvedValueText wb c1) (styleSig c1) (resolvedValueText rt c2) (styleSig c2))))
                ok ("rt-cells " + s1.Name)
                   (Map.count before = Map.count after
                    && (before |> Map.forall (fun ref c1 -> after |> Map.tryFind ref |> Option.exists (fun c2 -> resolvedValueText wb c1 = resolvedValueText rt c2 && styleSig c1 = styleSig c2))))
                   (sprintf "before=%d after=%d %s" (Map.count before) (Map.count after) (mismatch |> Option.defaultValue ""))
                ok ("rt-merges " + s1.Name) (mergeSig s1 = mergeSig s2) ""
                ok ("rt-columns " + s1.Name)
                   (columnSig s1 = columnSig s2)
                   (sprintf "before=%A after=%A" (columnSig s1 |> List.truncate 8) (columnSig s2 |> List.truncate 8))
                ok ("rt-rows " + s1.Name) (rowSig s1 = rowSig s2) ""
                ok ("rt-freeze " + s1.Name) (s1.FrozenPane = s2.FrozenPane) ""
                ok ("rt-tables " + s1.Name) (tableSig s1 = tableSig s2) ""
                ok ("rt-validations " + s1.Name) (s1.DataValidations = s2.DataValidations) ""
                ok ("rt-page " + s1.Name) (pageSig s1 = pageSig s2) ""

for fixture in [ "dcf-stock-analysis.xlsx"; "dashboard-sales-analysis.xlsx"; "milk-street-office.xlsx" ] do
    assertRoundTripFixture (Path.Combine("tests/real-world", fixture))

let syntheticFormat =
    { XlsxWriter.defaultXlsxCellFormat with
        Bold = true
        FontName = Some "Arial"
        FontSize = Some 12.0
        FillColor = Some "#E2F0D9"
        NumberFormat = Some "0.00"
        BorderBottom = Some { XlsxWriter.Style = "thin"; Color = Some "#112233" }
        HorizontalAlign = Some "center"
        WrapText = true }

let syntheticSheet : XlsxWriter.XlsxSheetOut =
    { Name = "RoundTrip"
      State = XlsxReader.SheetVisible
      Cells =
        [ { Row = 0; Col = 0; Value = XlsxWriter.XlsxText ("Header", []); Format = syntheticFormat; Comment = Some ("Note text", Some "Codex") }
          { Row = 1; Col = 0; Value = XlsxWriter.XlsxNumber 42.5; Format = syntheticFormat; Comment = None }
          { Row = 1; Col = 1; Value = XlsxWriter.XlsxFormula ("A2*2", Some (XlsxWriter.XlsxNumber 85.0)); Format = syntheticFormat; Comment = None } ]
      MergedCells = [ { StartRow = 0; StartCol = 0; EndRow = 0; EndCol = 1 } ]
      ColumnWidths = [ 0, 20.0; 1, 16.0 ]
      RowHeights = [ 0, 24.0 ]
      FrozenRows = 1
      FrozenCols = 1
      ShowGridLines = false
      Tables =
        [ { Name = "Table1"; StartRow = 1; StartCol = 1; EndRow = 2; EndCol = 2; Columns = [ "Header"; "Value" ]; HeaderRowCount = 1; TotalsRowCount = 0; AutoFilter = Some (1, 2); Style = Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false } } ]
      DataValidations =
        [ { Sqref = "A2"; ValidationType = "decimal"; Operator = Some "between"; Formula1 = Some "1"; Formula2 = Some "100"; AllowBlank = true; ShowDropDown = true; ErrorMessage = Some "No"; InputMessage = Some "Pick"; ErrorTitle = Some "Bad"; PromptTitle = Some "Hint"; ErrorStyle = Some "warning"; ShowErrorMessage = true; ShowInputMessage = true } ]
      ConditionalFormats = []
      Hyperlinks = [ { Range = "A1"; Target = "https://example.com"; IsInternal = false } ]
      PageMargins = Some { Top = 0.5; Bottom = 0.5; Left = 0.6; Right = 0.6; Header = 0.2; Footer = 0.2 }
      PageSetup = Some { Orientation = Some "landscape"; FitToPage = true; FitToWidth = 1; FitToHeight = 1; Scale = 90; PaperSize = Some 9; FirstPageNumber = 3; UseFirstPageNumber = true; PageOrder = Some "overThenDown"; BlackAndWhite = false; CellComments = Some "asDisplayed" }
      PrintOptions = Some { GridLines = true; Headings = true; HorizontalCentered = true; VerticalCentered = false }
      HeaderFooter = Some { OddHeader = Some "&CReport"; OddFooter = Some "&P"; EvenHeader = None; EvenFooter = None; FirstHeader = None; FirstFooter = None; DifferentOddEven = false; DifferentFirst = false; ScaleWithDoc = true }
      RowBreaks = [ 3 ]
      ColBreaks = [ 2 ]
      PrintArea = Some "$A$1:$B$2"
      PrintTitleRows = Some "$1:$1"
      PrintTitleCols = Some "$A:$A" }

let synthetic : XlsxWriter.XlsxWorkbookOut =
    { XlsxWriter.Sheets = [ syntheticSheet ]
      ActiveSheetIndex = 0
      DefinedNames = [ { Name = "MY_RANGE"; RefersTo = "'RoundTrip'!$A$1:$B$2"; LocalSheetIndex = None } ]
      VbaProjectSource = Some "Sub Main()\n  Range(\"A1\").Value = 1\nEnd Sub\n"
      Dxfs = [] }

match synthetic |> XlsxWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative with
| None -> ok "synthetic-roundtrip-load" false "reader returned None"
| Some wb ->
    let s = wb.Sheets.[0]
    ok "synthetic-sheet-name" (s.Name = "RoundTrip") s.Name
    ok "synthetic-style" (s.Rows.[0].Cells.[0].Style |> Option.exists (fun st -> st.Bold && st.FontName = Some "Arial" && st.FillColor = Some "E2F0D9" && st.NumberFormat = Some "0.00")) (sprintf "%A" s.Rows.[0].Cells.[0].Style)
    ok "synthetic-merge" (mergeSig s = [ "A1:B1" ]) (sprintf "%A" (mergeSig s))
    ok "synthetic-freeze" (s.FrozenPane = Some { XlsxReader.XSplit = 1; YSplit = 1 }) (sprintf "%A" s.FrozenPane)
    ok "synthetic-table" (s.Tables.Length = 1 && s.Tables.[0].Name = "Table1") (sprintf "%A" s.Tables)
    ok "synthetic-validation" (s.DataValidations.Length = 1 && s.DataValidations.[0].ValidationType = "decimal") (sprintf "%A" s.DataValidations)
    ok "synthetic-comment" (s.Comments.Length = 1 && s.Comments.[0].Text = "Note text") (sprintf "%A" s.Comments)
    ok "synthetic-page" (s.PageSetup |> Option.exists (fun ps -> ps.FitToPage && ps.Orientation = XlsxReader.OrientationLandscape)) (sprintf "%A" s.PageSetup)
    ok "synthetic-defined-name" (wb.NamedRanges |> List.exists (fun n -> n.Name = "MY_RANGE")) (sprintf "%A" wb.NamedRanges)

// ----- writer output Excel accepts --------------------------------------------

let partText (bytes: byte[]) (name: string) =
    use zip = new System.IO.Compression.ZipArchive(new System.IO.MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read)
    match zip.GetEntry name with
    | null -> ""
    | entry ->
        use reader = new System.IO.StreamReader(entry.Open())
        reader.ReadToEnd()

let cfRule (kind: string) : XlsxReader.ConditionalFormatRule =
    { Type = kind; Operator = None; Formulas = []; DxfId = None; ColorScale = []; DataBarColor = None; DataBarCfvo = []
      Text = None; IconSet = None; IconSetCfvo = []; IconSetReverse = false; Rank = None; Percent = false; Bottom = false
      AboveAverage = true; EqualAverage = false; StdDev = None; Priority = None; StopIfTrue = false }

let nastySheet : XlsxWriter.XlsxSheetOut =
    { syntheticSheet with
        Name = "Bob's <Q&A>"
        Cells =
          [ { Row = 0; Col = 0; Value = XlsxWriter.XlsxText ("a\u0001b\r\nc _x0041_ end", []); Format = { syntheticFormat with FontName = Some "A&B \"Font\"" }; Comment = Some ("note\u0002!", Some "Me") }
            { Row = 1; Col = 0; Value = XlsxWriter.XlsxNumber 0.1; Format = { XlsxWriter.defaultXlsxCellFormat with FillColor = Some "#FF0000" }; Comment = None }
            { Row = 2; Col = 0; Value = XlsxWriter.XlsxFormula ("A2*2", Some (XlsxWriter.XlsxNumber 0.2)); Format = { XlsxWriter.defaultXlsxCellFormat with FillColor = Some "#00FF00" }; Comment = None }
            { Row = 3; Col = 0; Value = XlsxWriter.XlsxText ("hidden row", []); Format = XlsxWriter.defaultXlsxCellFormat; Comment = None } ]
        MergedCells = []
        RowHeights = [ 3, 0.0 ]
        ColumnWidths = [ 1, 0.0 ]
        Tables = [ { syntheticSheet.Tables.[0] with Name = "T<1>"; Columns = [ "A&B"; "C" ]; Style = Some { Name = "Style\"X"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false } } ]
        DataValidations =
          [ { Sqref = "A2"; ValidationType = "whole"; Operator = Some "between"; Formula1 = Some "1"; Formula2 = Some "5"; AllowBlank = false; ShowDropDown = true
              ErrorMessage = Some "line1\nline2"; InputMessage = Some "Pick"; ErrorTitle = Some "Bad"; PromptTitle = Some "Hint"; ErrorStyle = Some "warning"; ShowErrorMessage = true; ShowInputMessage = true } ]
        ConditionalFormats =
          [ { Range = "A2:A3"
              Rules =
                [ { cfRule "cellIs" with Operator = Some "greaterThan"; Formulas = [ "0.15" ]; DxfId = Some 0; Priority = Some 1 }
                  { cfRule "containsText" with Operator = Some "containsText"; Text = Some "x<y"; Formulas = [ "NOT(ISERROR(SEARCH(\"x<y\",A2)))" ]; DxfId = Some 7 } ] }
            { Range = "B2:B3"
              Rules =
                [ { cfRule "colorScale" with ColorScale = [ { CfvoType = "min"; CfvoValue = None; Color = "FF0000" }; { CfvoType = "max"; CfvoValue = None; Color = "00FF00" } ] }
                  { cfRule "dataBar" with DataBarColor = Some "638EC6"; DataBarCfvo = [ "min", None; "max", None ] }
                  { cfRule "iconSet" with IconSet = Some "3Arrows"; IconSetCfvo = [ "percent", Some "0"; "percent", Some "33"; "percent", Some "67" ] }
                  { cfRule "top10" with Rank = Some 3; Bottom = true; Percent = true } ] } ]
        PrintArea = Some "$A$1:$B$2" }

let nastyWorkbook : XlsxWriter.XlsxWorkbookOut =
    { Sheets = [ nastySheet ]
      ActiveSheetIndex = 0
      DefinedNames = []
      VbaProjectSource = None
      Dxfs = [ { FontColor = Some "9C0006"; FillColor = Some "FFC7CE"; Bold = Some true; Italic = None } ] }
let nastyBytes = XlsxWriter.writeWorkbookPackageBytes nastyWorkbook

let nastySheetXml = partText nastyBytes "xl/worksheets/sheet1.xml"
let nastyStylesXml = partText nastyBytes "xl/styles.xml"

// Every part must be well-formed XML: control characters and unescaped
// quotes / angle brackets in names used to produce unreadable parts.
let malformedParts =
    use zip = new System.IO.Compression.ZipArchive(new System.IO.MemoryStream(nastyBytes), System.IO.Compression.ZipArchiveMode.Read)
    [ for entry in zip.Entries do
        if entry.FullName.EndsWith ".xml" || entry.FullName.EndsWith ".rels" then
            use reader = new System.IO.StreamReader(entry.Open())
            try System.Xml.Linq.XDocument.Parse(reader.ReadToEnd()) |> ignore
            with ex -> yield entry.FullName + ": " + ex.Message ]
ok "writer-parts-well-formed" (List.isEmpty malformedParts) (String.concat "; " malformedParts)

let indexIn (text: string) (needle: string) = text.IndexOf(needle, System.StringComparison.Ordinal)
// CT_Worksheet order: … mergeCells, conditionalFormatting, dataValidations,
// hyperlinks, page parts, breaks, legacyDrawing, tableParts.
ok "writer-element-order"
   (indexIn nastySheetXml "<conditionalFormatting" > indexIn nastySheetXml "</sheetData>"
    && indexIn nastySheetXml "<conditionalFormatting" < indexIn nastySheetXml "<dataValidations"
    && indexIn nastySheetXml "<dataValidations" < indexIn nastySheetXml "<hyperlinks"
    && indexIn nastySheetXml "<legacyDrawing" > 0
    && indexIn nastySheetXml "<legacyDrawing" < indexIn nastySheetXml "<tableParts")
   nastySheetXml
// Excel reserves fill 0 (none) and fill 1 (gray125); custom fills start at 2.
ok "writer-fill-ids-start-at-two"
   (indexIn nastyStylesXml "<fills count=\"5\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\">" >= 0
    && indexIn nastyStylesXml "fillId=\"1\"" < 0
    && indexIn nastyStylesXml "fillId=\"2\"" > 0 && indexIn nastyStylesXml "fillId=\"3\"" > 0)
   nastyStylesXml
ok "writer-validation-enforced"
   (indexIn nastySheetXml "showErrorMessage=\"1\"" > 0 && indexIn nastySheetXml "showInputMessage=\"1\"" > 0
    && indexIn nastySheetXml "errorStyle=\"warning\"" > 0 && indexIn nastySheetXml "errorTitle=\"Bad\"" > 0
    && indexIn nastySheetXml "promptTitle=\"Hint\"" > 0 && indexIn nastySheetXml "error=\"line1&#10;line2\"" > 0)
   nastySheetXml
ok "writer-hidden-row-and-column"
   (indexIn nastySheetXml "<row r=\"4\" hidden=\"1\">" > 0 && indexIn nastySheetXml "hidden=\"1\"/></cols>" > 0)
   nastySheetXml
ok "writer-numbers-shortest-roundtrip"
   (indexIn nastySheetXml "<v>0.1</v>" > 0 && indexIn nastySheetXml "<v>0.2</v>" > 0)
   nastySheetXml

match XlsxReader.loadBytesNative nastyBytes with
| None -> ok "writer-nasty-roundtrip" false "reader returned None"
| Some wb ->
    let s = wb.Sheets.[0]
    let first = s.Rows.[0].Cells.[0]
    ok "writer-control-chars-roundtrip"
       (XlsxReader.resolveValue wb first.Value = XlsxReader.String "a\u0001b\r\nc _x0041_ end"
        && (s.Comments |> List.map (fun c -> c.Text, c.Author)) = [ "note\u0002!", Some "Me" ])
       (sprintf "%A %A" (XlsxReader.resolveValue wb first.Value) s.Comments)
    ok "writer-names-escaped"
       (s.Name = "Bob's <Q&A>"
        && (first.Style |> Option.bind (fun st -> st.FontName)) = Some "A&B \"Font\""
        && (s.Tables |> List.map (fun t -> t.Name, t.Columns, (t.Style |> Option.map (fun st -> st.Name)))) = [ "T<1>", [ "A&B"; "C" ], Some "Style\"X" ]
        && s.PrintArea = Some "$A$1:$B$2")
       (sprintf "%A %A %A %A" s.Name first.Style s.Tables s.PrintArea)
    ok "writer-fill-colors-roundtrip"
       ((s.Rows |> List.map (fun r -> r.Cells.Head.Style |> Option.bind (fun st -> st.FillColor))) = [ Some "E2F0D9"; Some "FF0000"; Some "00FF00"; None ])
       (sprintf "%A" (s.Rows |> List.map (fun r -> r.Cells.Head.Style |> Option.bind (fun st -> st.FillColor))))
    ok "writer-formula-and-cached-roundtrip"
       (s.Rows.[2].Cells.Head.Formula = Some "A2*2" && s.Rows.[2].Cells.Head.Value = XlsxReader.Number 0.2)
       (sprintf "%A" s.Rows.[2].Cells.Head)
    ok "writer-hidden-roundtrip"
       ((s.Rows |> List.map (fun r -> r.Hidden)) = [ false; false; false; true ]
        && (s.Columns |> List.filter (fun c -> c.Hidden) |> List.map (fun c -> c.Index)) = [ 2 ])
       (sprintf "%A" s.Columns)
    ok "writer-validation-roundtrip"
       (match s.DataValidations with
        | [ dv ] ->
            dv.ShowErrorMessage && dv.ShowInputMessage && dv.ErrorStyle = Some "warning" && dv.ErrorTitle = Some "Bad"
            && dv.PromptTitle = Some "Hint" && dv.Prompt = Some "Pick" && dv.ErrorMessage = Some "line1\nline2" && not dv.AllowBlank
        | _ -> false)
       (sprintf "%A" s.DataValidations)
    ok "writer-conditional-formats-roundtrip"
       (match s.ConditionalFormats with
        | [ a; b ] ->
            a.Range = "A2:A3"
            && (a.Rules |> List.map (fun r -> r.Type, r.Operator, r.Formulas, r.DxfId, r.Text))
               = [ "cellIs", Some "greaterThan", [ "0.15" ], Some 0, None
                   "containsText", Some "containsText", [ "NOT(ISERROR(SEARCH(\"x<y\",A2)))" ], None, Some "x<y" ]
            && (b.Rules |> List.map (fun r -> r.Type)) = [ "colorScale"; "dataBar"; "iconSet"; "top10" ]
            && (b.Rules.[0].ColorScale |> List.map (fun stop -> stop.CfvoType, stop.Color)) = [ "min", "FF0000"; "max", "00FF00" ]
            && b.Rules.[1].DataBarColor = Some "638EC6" && b.Rules.[1].DataBarCfvo.Length = 2
            && b.Rules.[2].IconSet = Some "3Arrows" && b.Rules.[2].IconSetCfvo.Length = 3
            && b.Rules.[3].Rank = Some 3 && b.Rules.[3].Bottom && b.Rules.[3].Percent
        | _ -> false)
       (sprintf "%A" s.ConditionalFormats)
    ok "writer-dxf-roundtrip"
       (wb.Dxfs.Length = 1 && wb.Dxfs.[0].FontColor = Some "9C0006" && wb.Dxfs.[0].FillColor = Some "FFC7CE" && wb.Dxfs.[0].Bold = Some true)
       (sprintf "%A" wb.Dxfs)

// Reader -> writer -> reader keeps formulas (with their cached values),
// conditional formatting and sheet visibility.
for cfFile in [ "real-world/saopaulo-condos.xlsx"; "test-documents-excel/synthetic-condformat.xlsx" ] do
  let cfPath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", cfFile)
  match XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes cfPath) with
  | None -> ok ("cf-load " + cfFile) false "reader returned None"
  | Some wb1 ->
    match XlsxWriter.packageFromReaderWorkbook wb1 |> XlsxWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative with
    | None -> ok ("cf-reload " + cfFile) false "reader returned None"
    | Some wb2 ->
        let cfSig (wb: XlsxReader.Workbook) =
            wb.Sheets |> List.map (fun s -> s.ConditionalFormats |> List.map (fun cf -> cf.Range, cf.Rules))
        ok ("cf-rules-survive-save " + cfFile)
           (cfSig wb1 = cfSig wb2 && (cfSig wb1 |> List.exists (fun sheet -> not (List.isEmpty sheet))))
           (sprintf "%A vs %A" (cfSig wb1) (cfSig wb2))
        ok ("cf-dxfs-survive-save " + cfFile) (wb1.Dxfs = wb2.Dxfs) (sprintf "%A vs %A" wb1.Dxfs wb2.Dxfs)

printfn ""
printfn "OfficeFormat XLSX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
