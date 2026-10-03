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
      Visible = true
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
        [ { Sqref = "A2"; ValidationType = "decimal"; Operator = Some "between"; Formula1 = Some "1"; Formula2 = Some "100"; AllowBlank = true; ShowDropDown = true; ErrorMessage = Some "No"; InputMessage = Some "Pick" } ]
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
      VbaProjectSource = Some "Sub Main()\n  Range(\"A1\").Value = 1\nEnd Sub\n" }

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

printfn ""
printfn "OfficeFormat XLSX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
