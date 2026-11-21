#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.IO

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let cell value = XlsxPackageWriter.simpleCell value
let formulaCell formula value = XlsxPackageWriter.cellWithFormula formula value

let smallSheet: XlsxWriter.FormattedSheet =
    { Name = "Data"
      Rows =
        [ { Cells = [ cell "A & <x> \"q\" 's'"; cell "42" ] }
          { Cells = [ formulaCell "SUM(A1:B1)" "42" ] } ]
      MergedCells = [ XlsxPackageWriter.createMergedRange 0 0 1 1 ] }

let xml = XlsxPackageWriter.generateFormattedSheetXml smallSheet
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

XlsxPackageWriter.generateFormattedSheetXml perfSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let generated = XlsxPackageWriter.generateFormattedSheetXml perfSheet
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
    let bytes = XlsxWriter.packageFromReaderWorkbook wb |> XlsxPackageWriter.writeWorkbookPackageBytes
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
      CellSource =
        XlsxWriter.sheetCells <|
        [ { Row = 0; Col = 0; Value = XlsxWriter.XlsxText ("Header", []); Format = syntheticFormat; Comment = Some ("Note text", Some "Codex") }
          { Row = 1; Col = 0; Value = XlsxWriter.XlsxNumber 42.5; Format = syntheticFormat; Comment = None }
          { Row = 1; Col = 1; Value = XlsxWriter.XlsxFormula ("A2*2", Some (XlsxWriter.XlsxNumber 85.0)); Format = syntheticFormat; Comment = None } ]
      MergedCells = [ { StartRow = 0; StartCol = 0; EndRow = 0; EndCol = 1 } ]
      ColumnWidths = [ 0, 20.0; 1, 16.0 ]
      RowHeights = [ 0, 24.0 ]
      FrozenRows = 1
      FrozenCols = 1
      ShowGridLines = false
      View = { XlsxWriter.defaultSheetViewOut with ShowZeros = false; ShowHeaders = false; ShowFormulas = true; ZoomScale = 85 }
      Tables =
        [ { Name = "Table1"; StartRow = 1; StartCol = 1; EndRow = 2; EndCol = 2; Columns = [ "Header"; "Value" ]; HeaderRowCount = 1; TotalsRowCount = 0; Totals = []; AutoFilter = Some (1, 2); Style = Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false }; SourcePart = None } ]
      DataValidations =
        [ { Sqref = "A2"; ValidationType = "decimal"; Operator = Some "between"; Formula1 = Some "1"; Formula2 = Some "100"; AllowBlank = true; ShowDropDown = true; ErrorMessage = Some "No"; InputMessage = Some "Pick"; ErrorTitle = Some "Bad"; PromptTitle = Some "Hint"; ErrorStyle = Some "warning"; ShowErrorMessage = true; ShowInputMessage = true } ]
      ConditionalFormats = []
      Hyperlinks = [ { Range = "A1"; Target = "https://example.com"; IsInternal = false; Tooltip = None; Display = None } ]
      PageMargins = Some { Top = 0.5; Bottom = 0.5; Left = 0.6; Right = 0.6; Header = 0.2; Footer = 0.2 }
      PageSetup = Some { Orientation = Some "landscape"; FitToPage = true; FitToWidth = 1; FitToHeight = 1; Scale = 90; PaperSize = Some 9; FirstPageNumber = 3; UseFirstPageNumber = true; PageOrder = Some "overThenDown"; BlackAndWhite = false; CellComments = Some "asDisplayed"; Kept = [] }
      PrintOptions = Some { GridLines = true; Headings = true; HorizontalCentered = true; VerticalCentered = false }
      HeaderFooter = Some { OddHeader = Some "&CReport"; OddFooter = Some "&P"; EvenHeader = None; EvenFooter = None; FirstHeader = None; FirstFooter = None; DifferentOddEven = false; DifferentFirst = false; ScaleWithDoc = true }
      RowBreaks = [ 3 ]
      ColBreaks = [ 2 ]
      PrintArea = Some "$A$1:$B$2"
      PrintTitleRows = Some "$1:$1"
      PrintTitleCols = Some "$A:$A"
      Extras = XlsxReader.emptySheetExtras }

let synthetic : XlsxWriter.XlsxWorkbookOut =
    { XlsxWriter.Sheets = [ syntheticSheet ]
      ActiveSheetIndex = 0
      DefinedNames = [ { Name = "MY_RANGE"; RefersTo = "'RoundTrip'!$A$1:$B$2"; LocalSheetIndex = None; Comment = None; Hidden = false } ]
      VbaProjectSource = Some "Sub Main()\n  Range(\"A1\").Value = 1\nEnd Sub\n"
      VbaProjectBytes = None
      Dxfs = []
      Preserved = XlsxReader.emptyWorkbookPreservation }

match synthetic |> XlsxPackageWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative with
| None -> ok "synthetic-roundtrip-load" false "reader returned None"
| Some wb ->
    let s = wb.Sheets.[0]
    ok "synthetic-sheet-name" (s.Name = "RoundTrip") s.Name
    ok "synthetic-style" (s.Rows.[0].Cells.[0].Style |> Option.exists (fun st -> st.Bold && st.FontName = Some "Arial" && st.FillColor = Some "E2F0D9" && st.NumberFormat = Some "0.00")) (sprintf "%A" s.Rows.[0].Cells.[0].Style)
    ok "synthetic-merge" (mergeSig s = [ "A1:B1" ]) (sprintf "%A" (mergeSig s))
    ok "synthetic-freeze" (s.FrozenPane = Some { XlsxReader.XSplit = 1; YSplit = 1 }) (sprintf "%A" s.FrozenPane)
    ok "synthetic-view" ((s.ShowZeros, s.ShowGridLines, s.View) = (false, false, { XlsxReader.ShowRowColHeaders = false; ShowFormulas = true; ZoomScale = 85; TabColor = None; SummaryBelow = true; SummaryRight = true; ShowOutlineSymbols = true })) (sprintf "%A" (s.ShowZeros, s.View))
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

// Formula results are inline strings, not shared strings. Without xml:space,
// Excel discards a whitespace-only cached result when opening the workbook.
let whitespaceFormulaCell row formula text : XlsxWriter.XlsxCellOut =
    { Row = row; Col = 4; Value = XlsxWriter.XlsxFormula (formula, Some (XlsxWriter.XlsxText (text, [])))
      Format = XlsxWriter.defaultXlsxCellFormat; Comment = None }
let whitespaceSheet =
    { syntheticSheet with
        CellSource = XlsxWriter.sheetCells
            [ whitespaceFormulaCell 22 "CONCATENATE(C13,\" \",D13)" " "
              whitespaceFormulaCell 24 "\" pre & post \"" " pre & post " ]
        Tables = []
        MergedCells = []
        Extras = XlsxReader.emptySheetExtras }
let whitespaceBytes =
    XlsxPackageWriter.writeWorkbookPackageBytes { synthetic with Sheets = [ whitespaceSheet ]; VbaProjectSource = None }
let whitespaceXml = partText whitespaceBytes "xl/worksheets/sheet1.xml"
let whitespaceNamespace = System.Xml.Linq.XNamespace.Get "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
let parsedWhitespace =
    (System.Xml.Linq.XDocument.Parse whitespaceXml).Descendants(whitespaceNamespace + "v")
    |> Seq.map (fun value -> value.Value)
    |> Seq.toList
ok "formula cached text survives whitespace-discarding XML readers"
   (parsedWhitespace = [ " "; " pre & post " ])
   (sprintf "%A" parsedWhitespace)
match XlsxReader.loadBytesNative whitespaceBytes with
| None -> ok "formula-cached-whitespace-reload" false "reader returned None"
| Some wb ->
    let cells = cellMap wb wb.Sheets.Head
    ok "formula-cached-whitespace-reload"
       ([ "E23", " "; "E25", " pre & post " ]
        |> List.forall (fun (address, expected) ->
            Map.tryFind address cells |> Option.exists (fun cell -> cell.Value = XlsxReader.String expected)))
       (sprintf "%A" (cells |> Map.toList |> List.map (fun (address, cell) -> address, cell.Value)))

let builtins =
    { synthetic with
        Sheets =
            [ { syntheticSheet with
                    Tables = []
                    Extras = { syntheticSheet.Extras with AutoFilter = Some { Ref = "A1:B2"; Columns = [] } } } ]
        DefinedNames =
            [ { Name = "_xlnm.Criteria"; RefersTo = "'RoundTrip'!$A$1:$B$2"; LocalSheetIndex = Some 0; Comment = Some "Imported criterion"; Hidden = true }
              { Name = "_xlnm._FilterDatabase"; RefersTo = "'RoundTrip'!$A$1:$B$2"; LocalSheetIndex = Some 0; Comment = Some "Imported filter"; Hidden = true } ] }
let builtinBytes = XlsxPackageWriter.writeWorkbookPackageBytes builtins
match XlsxReader.loadBytesNative builtinBytes with
| None -> ok "built-in-names-import" false "reader returned None"
| Some wb ->
    ok "built-in-names-import"
        (wb.NamedRanges |> List.exists (fun name ->
            name.Name = "_xlnm.Criteria" && name.Scope = Some "0" && name.Comment = Some "Imported criterion" && name.Hidden))
        (sprintf "%A" wb.NamedRanges)
    ok "filter-database-single"
        (wb.NamedRanges |> List.filter (fun name -> name.Name = "_xlnm._FilterDatabase") |> List.length = 1)
        (sprintf "%A" wb.NamedRanges)

let cfRule (kind: string) : XlsxReader.ConditionalFormatRule =
    { Type = kind; Operator = None; Formulas = []; DxfId = None; ColorScale = []; DataBarColor = None; DataBarCfvo = []
      Text = None; IconSet = None; IconSetCfvo = []; IconSetReverse = false; TimePeriod = None; Rank = None; Percent = false; Bottom = false
      AboveAverage = true; EqualAverage = false; StdDev = None; Priority = None; StopIfTrue = false; Ext = None; SourceXml = None }

let nastySheet : XlsxWriter.XlsxSheetOut =
    { syntheticSheet with
        Name = "Bob's <Q&A>"
        CellSource =
          XlsxWriter.sheetCells <|
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
                  { cfRule "containsText" with Operator = Some "containsText"; Text = Some "x<y"; Formulas = [ "NOT(ISERROR(SEARCH(\"x<y\",A2)))" ]; DxfId = Some 7 } ]
              GroupAttributes = [] }
            { Range = "B2:B3"
              Rules =
                [ { cfRule "colorScale" with ColorScale = [ { CfvoType = "min"; CfvoValue = None; Color = "FF0000" }; { CfvoType = "max"; CfvoValue = None; Color = "00FF00" } ] }
                  { cfRule "dataBar" with DataBarColor = Some "638EC6"; DataBarCfvo = [ "min", None; "max", None ] }
                  { cfRule "iconSet" with IconSet = Some "3Arrows"; IconSetCfvo = [ "percent", Some "0"; "percent", Some "33"; "percent", Some "67" ] }
                  { cfRule "top10" with Rank = Some 3; Bottom = true; Percent = true } ]
              GroupAttributes = [] } ]
        PrintArea = Some "$A$1:$B$2" }

let nastyWorkbook : XlsxWriter.XlsxWorkbookOut =
    { Sheets = [ nastySheet ]
      ActiveSheetIndex = 0
      DefinedNames = []
      VbaProjectSource = None
      VbaProjectBytes = None
      Dxfs = [ { FontColor = Some "9C0006"; FillColor = Some "FFC7CE"; Bold = Some true; Italic = None; Underline = None; Strike = None; BorderTop = None; BorderBottom = None; BorderLeft = None; BorderRight = None; SourceDxfIndex = None } ]
      Preserved = XlsxReader.emptyWorkbookPreservation }
let nastyBytes = XlsxPackageWriter.writeWorkbookPackageBytes nastyWorkbook

let nastySheetXml = partText nastyBytes "xl/worksheets/sheet1.xml"
let nastyStylesXml = partText nastyBytes "xl/styles.xml"


let rawX14Id = "{11111111-2222-3333-4444-555555555555}"
let opaqueX14 =
    { cfRule "futureRule" with
        Priority = Some 5
        Ext = Some { XlsxReader.emptyCfRuleExt with InExtList = true }
        SourceXml =
            Some (sprintf "<x14:cfRule type=\"futureRule\" id=\"%s\" priority=\"100\" future=\"kept\"><xm:f>1=1</xm:f><x14:future value=\"extra\"/></x14:cfRule>" rawX14Id) }
let advancedBar: XlsxReader.CfDataBarExt =
    { MinLength = 5; MaxLength = 95; ShowValue = false; Border = true; Gradient = false
      Direction = "rightToLeft"; NegativeBarColorSameAsPositive = false
      NegativeBarBorderColorSameAsPositive = false; AxisPosition = "middle"
      Cfvo = [ "autoMin", None; "autoMax", None ]; BorderColor = Some "123456"
      NegativeFillColor = Some "CC0000"; NegativeBorderColor = Some "EE0000"; AxisColor = Some "ABCDEF" }
let withOpaque =
    { nastyWorkbook with
        Sheets =
            [ { nastySheet with
                    ConditionalFormats =
                        [ { Range = "E2"; GroupAttributes = [ "pivot", "1" ]
                            Rules = [ { cfRule "cellIs" with Operator = Some "greaterThan"; Formulas = [ "3" ]; Priority = Some 4 } ] }
                          { Range = "B2:B5"; GroupAttributes = []
                            Rules =
                                [ { cfRule "dataBar" with
                                        Priority = Some 2
                                        DataBarColor = Some "4472C4"
                                        DataBarCfvo = [ "min", None; "max", None ]
                                        Ext = Some { XlsxReader.emptyCfRuleExt with DataBar = Some advancedBar } } ] }
                          { Range = "E3"; GroupAttributes = []; Rules = [ opaqueX14 ] } ] } ] }
let opaqueBytes = XlsxPackageWriter.writeWorkbookPackageBytes withOpaque
let opaqueXml = partText opaqueBytes "xl/worksheets/sheet1.xml"
ok "writer-pivot-rule-group"
    (opaqueXml.Contains("<conditionalFormatting sqref=\"E2\" pivot=\"1\">"))
    opaqueXml
ok "writer-opaque-x14-payload"
    (opaqueXml.Contains(rawX14Id) && opaqueXml.Contains("priority=\"5\" future=\"kept\"")
     && opaqueXml.Contains("<x14:future value=\"extra\"/>") && opaqueXml.Contains("<xm:sqref>E3</xm:sqref>"))
    opaqueXml
match XlsxReader.loadBytesNative opaqueBytes with
| None -> ok "mixed-x14-reader" false "reader returned None"
| Some wb ->
    let imported = wb.Sheets.[0].ConditionalFormats |> List.collect (fun block -> block.Rules)
    let bar = imported |> List.tryFind (fun rule -> rule.Type = "dataBar")
    let unknown = imported |> List.tryFind (fun rule -> rule.Type = "futureRule")
    ok "mixed-x14-bar-options"
        (bar |> Option.bind (fun rule -> rule.Ext) |> Option.bind (fun ext -> ext.DataBar)
         |> Option.exists (fun ext -> ext.MinLength = 5 && ext.MaxLength = 95 && ext.Direction = "rightToLeft"
                                      && ext.BorderColor = Some "123456" && ext.Cfvo = advancedBar.Cfvo))
        (sprintf "%A" bar)
    ok "mixed-x14-unknown-opaque"
        (unknown |> Option.exists (fun rule -> rule.SourceXml |> Option.exists (fun xml -> xml.Contains("future=\"kept\""))))
        (sprintf "%A" unknown)
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

// Editing a supported differential property must not erase unsupported source
// properties or alter another rule that shares the original differential style.
do
    let ns = System.Xml.Linq.XNamespace.Get "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    use stream = new MemoryStream()
    stream.Write(nastyBytes, 0, nastyBytes.Length)
    do
        use zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Update, true)
        let entry = zip.GetEntry "xl/styles.xml"
        let styles =
            use input = entry.Open()
            System.Xml.Linq.XDocument.Load input
        let dxf = styles.Root.Element(ns + "dxfs").Elements(ns + "dxf") |> Seq.head
        let font = dxf.Element(ns + "font")
        font.Add(System.Xml.Linq.XElement(ns + "sz", System.Xml.Linq.XAttribute("val", "16")))
        font.Add(System.Xml.Linq.XElement(ns + "name", System.Xml.Linq.XAttribute("val", "Arial")))
        font.AddAfterSelf(System.Xml.Linq.XElement(ns + "numFmt", System.Xml.Linq.XAttribute("numFmtId", "165"), System.Xml.Linq.XAttribute("formatCode", "0.000")))
        dxf.Add(System.Xml.Linq.XElement(ns + "border", System.Xml.Linq.XElement(ns + "left", System.Xml.Linq.XAttribute("style", "medium"))))
        entry.Delete()
        use output = zip.CreateEntry("xl/styles.xml").Open()
        styles.Save output
    let original = XlsxReader.loadBytesNative (stream.ToArray()) |> Option.get |> XlsxWriter.packageFromReaderWorkbook
    // Source-derived DXFs require the same original-part contract as the grid exporter.
    let preserved = original.Preserved
    let sourcePaths = preserved.SourceParts |> List.map fst |> Set.ofList
    let cloneIndex = original.Dxfs.Length
    let group = original.Sheets.Head.ConditionalFormats.Head
    let changedGroup =
        { group with Range = "D2:D3"
                     Rules = [ { group.Rules.Head with DxfId = Some cloneIndex; Formulas = [ "99" ]; SourceXml = None } ] }
    let changed =
        { original with
            Preserved =
                { preserved with InPlace = true
                                 Parts = (preserved.Parts |> List.filter (fun (path, _) -> not (Set.contains path sourcePaths))) @ preserved.SourceParts }
            Dxfs = original.Dxfs @ [ { original.Dxfs.Head with FillColor = Some "AADD99"; SourceDxfIndex = Some 0 } ]
            Sheets = original.Sheets |> List.mapi (fun index sheet ->
                if index = 0 then { sheet with ConditionalFormats = sheet.ConditionalFormats @ [ changedGroup ] } else sheet) }
    let saved = XlsxPackageWriter.writeWorkbookPackageBytes changed
    let dxfs = (System.Xml.Linq.XDocument.Parse(partText saved "xl/styles.xml")).Root.Element(ns + "dxfs").Elements(ns + "dxf") |> Seq.toArray
    let attribute name (element: System.Xml.Linq.XElement) = element.Attribute(System.Xml.Linq.XName.Get name).Value
    let fill index = dxfs.[index].Descendants(ns + "bgColor") |> Seq.head |> attribute "rgb"
    ok "editing a shared DXF clones only its requested fill"
        (fill 0 = "FFFFC7CE" && fill cloneIndex = "FFAADD99")
        (sprintf "%s / %s" (fill 0) (fill cloneIndex))
    let clone = dxfs.[cloneIndex]
    ok "a differential-style clone preserves source font metrics, number format and border"
        (attribute "val" (clone.Descendants(ns + "sz") |> Seq.head) = "16"
         && attribute "val" (clone.Descendants(ns + "name") |> Seq.head) = "Arial"
         && attribute "formatCode" (clone.Element(ns + "numFmt")) = "0.000"
         && attribute "style" (clone.Descendants(ns + "left") |> Seq.head) = "medium")
        (string clone)

// Reader -> writer -> reader keeps formulas (with their cached values),
// conditional formatting and sheet visibility.
for cfFile in [ "real-world/saopaulo-condos.xlsx"; "test-documents-excel/synthetic-condformat.xlsx" ] do
  let cfPath = System.IO.Path.Combine(__SOURCE_DIRECTORY__, "..", cfFile)
  match XlsxReader.loadBytesNative (System.IO.File.ReadAllBytes cfPath) with
  | None -> ok ("cf-load " + cfFile) false "reader returned None"
  | Some wb1 ->
    match XlsxWriter.packageFromReaderWorkbook wb1 |> XlsxPackageWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative with
    | None -> ok ("cf-reload " + cfFile) false "reader returned None"
    | Some wb2 ->
        let cfSig (wb: XlsxReader.Workbook) =
            wb.Sheets |> List.map (fun s -> s.ConditionalFormats |> List.map (fun cf -> cf.Range, cf.Rules))
        ok ("cf-rules-survive-save " + cfFile)
           (cfSig wb1 = cfSig wb2 && (cfSig wb1 |> List.exists (fun sheet -> not (List.isEmpty sheet))))
           (sprintf "%A vs %A" (cfSig wb1) (cfSig wb2))
        ok ("cf-dxfs-survive-save " + cfFile) (wb1.Dxfs = wb2.Dxfs) (sprintf "%A vs %A" wb1.Dxfs wb2.Dxfs)

// ---- pictures, charts, threaded comments, named styles and the theme ----------
let drawingSummary (wb: XlsxReader.Workbook) =
    wb.Sheets
    |> List.map (fun sheet ->
        sheet.Name,
        sheet.Extras.Images |> List.map (fun image -> image.Anchor, image.Name, image.Description, image.ContentType, image.Data),
        sheet.Extras.Objects,
        sheet.Extras.ThreadedComments,
        sheet.Extras.DrawingRelationships),
    wb.Preserved.Parts, wb.Preserved.ThemePath, wb.Preserved.NamedStyles
// Real workbooks with pictures and charts: everything the reader finds is
// found again, identical, in what the writer saves (chart parts byte for byte).
let drawingFiles =
    [ "real-world/chart-zoo.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx"; "real-world/dashboard-sales-analysis.xlsx"
      "test-documents-excel/pandas-chartsheet.xlsx"; "test-documents-excel/phpsp-27template.xlsx"; "test-documents-excel/phpsp-32chartreadwrite.xlsx"
      "test-documents-excel/poi-47504.xlsx"; "test-documents-excel/poi-47668.xlsx"; "test-documents-excel/poi-47813.xlsx"
      "test-documents-excel/poi-49156.xlsx"; "test-documents-excel/synthetic-chart-column.xlsx" ]
let mutable drawingPictures = 0
let mutable drawingCharts = 0
for relative in drawingFiles do
    let path = Path.Combine(__SOURCE_DIRECTORY__, "..", relative)
    if File.Exists path then
        match loadWorkbook path with
        | None -> ok ("drawings-load " + relative) false "reader returned None"
        | Some wb1 ->
            match roundTripWorkbook wb1 with
            | None -> ok ("drawings-reload " + relative) false "reader returned None"
            | Some wb2 ->
                drawingPictures <- drawingPictures + (wb1.Sheets |> List.sumBy (fun sheet -> sheet.Extras.Images.Length))
                drawingCharts <- drawingCharts + (wb1.Sheets |> List.sumBy (fun sheet -> sheet.Extras.Objects |> List.filter (fun item -> item.Kind = "chart") |> List.length))
                ok ("drawings-survive-save " + relative) (drawingSummary wb1 = drawingSummary wb2) "pictures, objects, kept parts, theme or named styles differ after a save"
ok "drawings-corpus-has-pictures-and-charts" (drawingPictures >= 10 && drawingCharts >= 25) (sprintf "%d pictures, %d charts" drawingPictures drawingCharts)

// A picture, a chart, a threaded comment with a reply, a named style and a theme, synthetic.
let extrasPackage =
    let ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    let rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    use stream = new MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new StreamWriter(archive.CreateEntry(name).Open(), System.Text.UTF8Encoding(false))
            writer.Write text
        let addBytes (name: string) (bytes: byte[]) =
            use entry = archive.CreateEntry(name).Open()
            entry.Write(bytes, 0, bytes.Length)
        add "[Content_Types].xml" "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"png\" ContentType=\"image/png\"/><Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/><Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>"
        add "xl/workbook.xml" (sprintf "<workbook xmlns=\"%s\" xmlns:r=\"%s\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>" ns rel)
        add "xl/_rels/workbook.xml.rels" (sprintf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"%s/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"%s/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"%s/theme\" Target=\"theme/theme1.xml\"/><Relationship Id=\"rId4\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"persons/person.xml\"/></Relationships>" rel rel rel)
        add "xl/theme/theme1.xml" "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Kept theme\"><a:themeElements><a:clrScheme name=\"Kept\"><a:accent1><a:srgbClr val=\"112233\"/></a:accent1></a:clrScheme></a:themeElements></a:theme>"
        add "xl/styles.xml" (sprintf "<styleSheet xmlns=\"%s\"><numFmts count=\"1\"><numFmt numFmtId=\"170\" formatCode=\"0.000 &quot;kg&quot;\"/></numFmts><fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"14\"/><color rgb=\"FF9C0006\"/><name val=\"Georgia\"/></font></fonts><fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill><fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFC7CE\"/></patternFill></fill></fills><borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/><xf numFmtId=\"170\" fontId=\"1\" fillId=\"2\" borderId=\"0\" applyFont=\"1\" applyFill=\"1\"><alignment horizontal=\"center\"/></xf></cellStyleXfs><cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"2\"><cellStyle name=\"Weighty\" xfId=\"1\"/><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>" ns)
        add "xl/worksheets/sheet1.xml" (sprintf "<worksheet xmlns=\"%s\" xmlns:r=\"%s\"><sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c></row></sheetData><drawing r:id=\"rId1\"/></worksheet>" ns rel)
        add "xl/worksheets/_rels/sheet1.xml.rels" (sprintf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"%s/drawing\" Target=\"../drawings/drawing1.xml\"/><Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"../threadedComments/threadedComment1.xml\"/></Relationships>" rel)
        add "xl/drawings/drawing1.xml" (sprintf "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"%s\" xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>1</xdr:col><xdr:colOff>9525</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>6</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"2\" name=\"Logo\" descr=\"the logo\"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed=\"rId1\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1219200\" cy=\"762000\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor><xdr:oneCellAnchor><xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"3048000\" cy=\"1905000\"/><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"3\" name=\"Chart 1\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart r:id=\"rId2\"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>" rel)
        add "xl/drawings/_rels/drawing1.xml.rels" (sprintf "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"%s/image\" Target=\"../media/logo.png\"/><Relationship Id=\"rId2\" Type=\"%s/chart\" Target=\"../charts/chart1.xml\"/></Relationships>" rel rel)
        addBytes "xl/media/logo.png" [| 137uy; 80uy; 78uy; 71uy; 1uy; 2uy; 3uy; 250uy |]
        add "xl/charts/chart1.xml" "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea/></c:chart></c:chartSpace>"
        add "xl/persons/person.xml" "<personList xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\"><person displayName=\"Ann Author\" id=\"{11111111-1111-1111-1111-111111111111}\" userId=\"ann@example.com\" providerId=\"AD\"/><person displayName=\"Bob Reply\" id=\"{22222222-2222-2222-2222-222222222222}\" userId=\"bob\" providerId=\"None\"/></personList>"
        add "xl/threadedComments/threadedComment1.xml" "<ThreadedComments xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments\"><threadedComment ref=\"B2\" dT=\"2024-01-02T03:04:05.00\" personId=\"{11111111-1111-1111-1111-111111111111}\" id=\"{AAAAAAAA-0000-0000-0000-000000000001}\"><text>Is this right?</text></threadedComment><threadedComment ref=\"B2\" dT=\"2024-01-03T04:05:06.00\" personId=\"{22222222-2222-2222-2222-222222222222}\" id=\"{AAAAAAAA-0000-0000-0000-000000000002}\" parentId=\"{AAAAAAAA-0000-0000-0000-000000000001}\"><text>Yes &amp; checked</text></threadedComment></ThreadedComments>"
    stream.ToArray()
match XlsxReader.loadBytesNative extrasPackage with
| None -> ok "extras-load" false "reader returned None"
| Some wb1 ->
    let sheet = wb1.Sheets.Head
    ok "picture-read-with-anchor-and-data"
       ((sheet.Extras.Images |> List.map (fun image -> image.Anchor.Kind, image.Anchor.FromCol, image.Anchor.FromColOffset, image.Anchor.FromRow, image.Anchor.ToCol, image.Anchor.ToRow, image.Anchor.EditAs, image.Anchor.Width, image.Name, image.Description, image.ContentType, image.Data.Length)) =
          [ "twoCell", 1, 9525.0, 2, Some 3, Some 6, Some "oneCell", Some 1219200.0, Some "Logo", Some "the logo", "image/png", 8 ])
       (sprintf "%A" sheet.Extras.Images)
    ok "chart-kept-as-object"
       ((sheet.Extras.Objects |> List.map (fun item -> item.Kind, item.Name, item.Anchor.Kind, item.Anchor.FromCol, item.Anchor.Width, item.Xml.Contains "<c:chart r:id=\"rId2\"/>")) = [ "chart", Some "Chart 1", "oneCell", 5, Some 3048000.0, true ])
       (sprintf "%A" sheet.Extras.Objects)
    ok "threaded-comment-with-author-and-reply"
       ((sheet.Extras.ThreadedComments |> List.map (fun thread -> thread.CellRef.Column, thread.CellRef.Row, thread.Author, thread.Date, thread.Text, thread.Replies |> List.map (fun reply -> reply.Author, reply.Text))) =
          [ "B", 2, Some "Ann Author", Some "2024-01-02T03:04:05.00", "Is this right?", [ Some "Bob Reply", "Yes & checked" ] ])
       (sprintf "%A" sheet.Extras.ThreadedComments)
    ok "named-styles-read"
       ((wb1.Preserved.NamedStyles |> List.map (fun style -> style.Name, style.NumFmtCode, style.FontXml.IsSome, style.FillXml |> Option.map (fun xml -> xml.Contains "FFFFC7CE"), style.XfChildrenXml)) =
          [ "Normal", None, true, Some false, ""; "Weighty", Some "0.000 \"kg\"", true, Some true, "<alignment horizontal=\"center\"/>" ])
       (sprintf "%A" wb1.Preserved.NamedStyles)
    ok "theme-and-chart-parts-kept"
       // (... and the sheet's drawing part itself, with its relationships and picture: it goes back as it is while the drawing is untouched.)
       ((wb1.Preserved.Parts |> List.map fst |> List.filter (fun path -> not (path.StartsWith "xl/drawings/" || path.StartsWith "xl/media/"))) = [ "xl/charts/chart1.xml"; "xl/theme/theme1.xml" ]
        && (wb1.Preserved.Parts |> List.map fst |> List.filter (fun path -> path.StartsWith "xl/drawings/" && not (path.Contains "_rels"))).Length = 1
        && wb1.Preserved.ThemePath = Some "xl/theme/theme1.xml")
       (sprintf "%A" (wb1.Preserved.Parts |> List.map fst))
    let savedBytes = XlsxWriter.packageFromReaderWorkbook wb1 |> XlsxPackageWriter.writeWorkbookPackageBytes
    match XlsxReader.loadBytesNative savedBytes with
    | None -> ok "extras-reload" false "reader returned None"
    | Some wb2 ->
        ok "extras-survive-save" (drawingSummary wb1 = drawingSummary wb2) (sprintf "%A\n%A" (drawingSummary wb1) (drawingSummary wb2))
        ok "persons-survive-save" (wb1.Preserved.Persons = wb2.Preserved.Persons && wb2.Preserved.Persons.Length = 2) (sprintf "%A" wb2.Preserved.Persons)
    // The saved package itself: the named style with its font / fill / number format, the theme, the chart part.
    use savedArchive = new System.IO.Compression.ZipArchive(new MemoryStream(savedBytes), System.IO.Compression.ZipArchiveMode.Read)
    let part (name: string) =
        match savedArchive.GetEntry name with
        | null -> ""
        | entry -> (use reader = new StreamReader(entry.Open()) in reader.ReadToEnd())
    let stylesPart = part "xl/styles.xml"
    ok "named-style-written-with-its-formatting"
       (stylesPart.Contains "<cellStyle name=\"Weighty\" xfId=\"1\"/>" && stylesPart.Contains "<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/>"
        && stylesPart.Contains "<name val=\"Georgia\"/>" && stylesPart.Contains "FFFFC7CE" && stylesPart.Contains "formatCode=\"0.000 &quot;kg&quot;\""
        && stylesPart.Contains "<alignment horizontal=\"center\"/>")
       stylesPart
    ok "theme-written-back-unchanged" ((part "xl/theme/theme1.xml").Contains "name=\"Kept theme\"" && System.Text.RegularExpressions.Regex.IsMatch(part "xl/_rels/workbook.xml.rels", "/theme\" Target=\"(/xl/)?theme/theme1.xml\"")) (part "xl/_rels/workbook.xml.rels")
    ok "chart-part-written-back-unchanged" ((part "xl/charts/chart1.xml") = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><c:plotArea/></c:chart></c:chartSpace>") (part "xl/charts/chart1.xml")
    ok "threaded-comments-written"
       ((part "xl/threadedComments/threadedComment1.xml").Contains "parentId=\"{AAAAAAAA-0000-0000-0000-000000000001}\"" && (part "xl/persons/person.xml").Contains "displayName=\"Bob Reply\"" && (part "[Content_Types].xml").Contains "threadedcomments+xml")
       (part "xl/threadedComments/threadedComment1.xml")
    ok "drawing-written-before-other-sheet-parts" ((part "xl/worksheets/sheet1.xml").Contains "<drawing r:id=\"rIdDrawing1\"/></worksheet>") (part "xl/worksheets/sheet1.xml")

// A workbook made from scratch has none of this, and still no dangling parts.
ok "new-workbook-has-no-drawing-parts"
   (use archive = new System.IO.Compression.ZipArchive(new MemoryStream(XlsxPackageWriter.writeWorkbookPackageBytes synthetic), System.IO.Compression.ZipArchiveMode.Read)
    archive.Entries |> Seq.forall (fun entry -> not (entry.FullName.Contains "drawings/sheetDrawing") && not (entry.FullName.Contains "theme") && not (entry.FullName.Contains "threadedComment")))
   "unexpected parts"

// ---- <col> ranges and styles, fullCalcOnLoad, named cell styles, z-order ----
let realWorkbook (relative: string) = XlsxReader.loadBytesNative (File.ReadAllBytes(Path.Combine(__SOURCE_DIRECTORY__, "..", relative))) |> Option.get
for file in [ "real-world/lyric-multifamily.xlsx"; "real-world/saopaulo-condos.xlsx"; "real-world/milk-street-office.xlsx"; "test-documents-excel/poi-49156.xlsx" ] do
    let before = realWorkbook file
    let bytes = XlsxPackageWriter.writeWorkbookPackageBytes (XlsxWriter.packageFromReaderWorkbook before)
    let after = XlsxReader.loadBytesNative bytes |> Option.get
    // Column definitions come back as the file had them: the same ranges, widths, hidden flags and default cell styles.
    let columnsOf (workbook: XlsxReader.Workbook) =
        workbook.Sheets
        |> List.map (fun sheet ->
            sheet.Columns
            |> List.map (fun column ->
                column.Index, column.MaxIndex, column.Width |> Option.map (fun w -> Math.Round(w, 2)), column.Hidden,
                column.Style |> Option.bind (fun index -> if index > 0 && index < workbook.Styles.Length then Some workbook.Styles.[index] else None)))
    let merged (columns: (int * int * float option * bool * XlsxReader.CellStyle option) list) =
        // The file may itself split a run of equal columns; compare with equal neighbours joined.
        columns
        |> List.fold (fun acc (lo, hi, width, hidden, style) ->
            match acc with
            | (plo, phi, pwidth, phidden, pstyle) :: rest when phi + 1 = lo && pwidth = width && phidden = hidden && pstyle = style -> (plo, hi, width, hidden, style) :: rest
            | _ -> (lo, hi, width, hidden, style) :: acc) []
        |> List.rev
    ok (file + " col ranges and styles") (List.map merged (columnsOf after) = List.map merged (columnsOf before))
       (sprintf "%A\n   became\n%A" (columnsOf before |> List.map (List.truncate 6)) (columnsOf after |> List.map (List.truncate 6)))
    ok (file + " col ranges are not split") (List.sum (List.map List.length (columnsOf after)) <= List.sum (List.map List.length (columnsOf before)))
       (sprintf "%d <col> became %d" (List.sum (List.map List.length (columnsOf before))) (List.sum (List.map List.length (columnsOf after))))
    // Every cell is still based on the named cell style it had.
    let namedOf (workbook: XlsxReader.Workbook) =
        let names = workbook.Preserved.NamedStyles |> List.map (fun style -> style.Name) |> List.toArray
        [ for sheet in workbook.Sheets do
            for row in sheet.Rows do
                for cell in row.Cells do
                    let name =
                        match cell.StyleIndex with
                        | Some index when index >= 0 && index < workbook.Preserved.CellXfNamedStyles.Length ->
                            let named = workbook.Preserved.CellXfNamedStyles.[index]
                            if named >= 0 && named < names.Length then names.[named] else "?"
                        | _ -> if names.Length > 0 then names.[0] else "?"
                    yield sheet.Name, cell.Reference, name ]
    let namedAfter = namedOf after |> List.map (fun (sheet, reference, name) -> (sheet, reference), name) |> Map.ofList
    let lost = namedOf before |> List.filter (fun (sheet, reference, name) -> Map.tryFind (sheet, reference) namedAfter <> Some name)
    ok (file + " named style of every cell") lost.IsEmpty (sprintf "%d cells, first: %A -> %A" lost.Length (List.tryHead lost) (lost |> List.tryHead |> Option.bind (fun (sheet, reference, _) -> Map.tryFind (sheet, reference) namedAfter)))
    ok (file + " every cell format knows its named style") (before.Styles.Length = before.Preserved.CellXfNamedStyles.Length) (sprintf "%d styles, %d links" before.Styles.Length before.Preserved.CellXfNamedStyles.Length)
    // An Excel-calculated workbook has every result cached: no forced recalculation.
    ok (file + " no fullCalcOnLoad when results are cached") (not ((partText bytes "xl/workbook.xml").Contains "fullCalcOnLoad")) "calcPr written"
let namedStylesInUse =
    let workbook = realWorkbook "real-world/milk-street-office.xlsx"
    workbook.Preserved.CellXfNamedStyles |> Array.filter (fun named -> named > 0) |> Array.distinct |> Array.length
ok "the named-style check is not vacuous" (namedStylesInUse >= 1) (sprintf "%d named styles other than the first are used" namedStylesInUse)
// Before this was kept, every cell format was written with xfId="0".
ok "a cell format is written with the xfId of its named style"
   (let styles = partText (XlsxPackageWriter.writeWorkbookPackageBytes (XlsxWriter.packageFromReaderWorkbook (realWorkbook "real-world/milk-street-office.xlsx"))) "xl/styles.xml"
    let cellXfs = styles.Substring(styles.IndexOf "<cellXfs")
    System.Text.RegularExpressions.Regex.IsMatch(cellXfs.Substring(0, cellXfs.IndexOf "</cellXfs>"), "xfId=\"[1-9]"))
   "every xf has xfId 0"

// fullCalcOnLoad: only when a formula goes out without a result.
let calcWorkbook (cached: XlsxWriter.XlsxCellValue option) =
    let sheet = (XlsxWriter.packageFromReaderWorkbook (realWorkbook "test-documents-excel/poi-49156.xlsx")).Sheets.Head
    let formula : XlsxWriter.XlsxCellOut = { Row = 60; Col = 0; Value = XlsxWriter.XlsxFormula ("1+1", cached); Format = XlsxWriter.defaultXlsxCellFormat; Comment = None }
    { (XlsxWriter.packageFromReaderWorkbook (realWorkbook "test-documents-excel/poi-49156.xlsx")) with Sheets = [ { sheet with CellSource = XlsxWriter.sheetCells (sheet.Cells @ [ formula ]) } ] }
let uncachedXml = partText (XlsxPackageWriter.writeWorkbookPackageBytes (calcWorkbook None)) "xl/workbook.xml"
ok "fullCalcOnLoad for a formula without a cached value" (uncachedXml.Contains "<calcPr fullCalcOnLoad=\"1\"/></workbook>") uncachedXml
ok "no fullCalcOnLoad when the formula has its value" (not ((partText (XlsxPackageWriter.writeWorkbookPackageBytes (calcWorkbook (Some (XlsxWriter.XlsxNumber 2.0)))) "xl/workbook.xml").Contains "calcPr")) "calcPr written"
ok "the uncached workbook still opens" ((XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (calcWorkbook None))).IsSome) "unreadable"

// z-order: pictures and kept objects stay interleaved as the drawing part had them.
let zOrderOf (workbook: XlsxReader.Workbook) =
    workbook.Sheets
    |> List.map (fun sheet ->
        (sheet.Extras.Images |> List.map (fun image -> image.Anchor.Order, "picture", image.Name))
        @ (sheet.Extras.Objects |> List.map (fun item -> item.Anchor.Order, item.Kind, item.Name))
        |> List.sortBy (fun (order, _, _) -> order)
        |> List.map (fun (_, kind, name) -> kind, name))
for file in [ "test-documents-excel/phpsp-32chartreadwrite.xlsx"; "real-world/dashboard-ecommerce-sales.xlsx"; "test-documents-excel/phpsp-27template.xlsx" ] do
    let before = realWorkbook file
    let after = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes (XlsxWriter.packageFromReaderWorkbook before)) |> Option.get
    ok (file + " z-order of pictures and objects") (zOrderOf after = zOrderOf before) (sprintf "%A became %A" (zOrderOf before) (zOrderOf after))
// A chart behind a picture stays behind it: swap the orders and write.
let swapped =
    let package = XlsxWriter.packageFromReaderWorkbook (realWorkbook "test-documents-excel/phpsp-32chartreadwrite.xlsx")
    { package with
        Sheets =
            package.Sheets
            |> List.map (fun sheet ->
                { sheet with
                    Extras =
                        { sheet.Extras with
                            Images = sheet.Extras.Images |> List.map (fun image -> { image with Anchor = { image.Anchor with Order = image.Anchor.Order + 100 } }) } }) }
let swappedBack = XlsxReader.loadBytesNative (XlsxPackageWriter.writeWorkbookPackageBytes swapped) |> Option.get
ok "objects are written in front of or behind pictures by their order"
   (zOrderOf swappedBack |> List.forall (fun sheet -> match List.tryFindIndex (fun (kind, _) -> kind = "picture") sheet with | Some first -> sheet |> List.skip first |> List.forall (fun (kind, _) -> kind = "picture") | None -> true)
    && zOrderOf swappedBack |> List.exists (fun sheet -> sheet |> List.exists (fun (kind, _) -> kind <> "picture") && sheet |> List.exists (fun (kind, _) -> kind = "picture")))
   (sprintf "%A" (zOrderOf swappedBack))

// ---- functions newer than Excel 2007 are written under their prefixes (Excel refuses a bare one it knows) ----
for formula, want in
    [ "SUM(A1:A3)", "SUM(A1:A3)"
      "XLOOKUP(2,A1:A4,B1:B4)", "_xlfn.XLOOKUP(2,A1:A4,B1:B4)"
      "_xlfn.XLOOKUP(2,A1:A4,B1:B4)", "_xlfn.XLOOKUP(2,A1:A4,B1:B4)"
      "SORT(FILTER(A2:A7,A2:A7>3))", "_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A2:A7,A2:A7>3))"
      "maxifs(B1:B6,A1:A6,\"X\")", "_xlfn.maxifs(B1:B6,A1:A6,\"X\")"
      "IF(A1=\"XLOOKUP(\",1,2)", "IF(A1=\"XLOOKUP(\",1,2)"
      "'IFS(x'!A1+STDEV.S(A1:A3)", "'IFS(x'!A1+_xlfn.STDEV.S(A1:A3)"
      "LET(x,5,y,x*2,x+y)", "_xlfn.LET(_xlpm.x,5,_xlpm.y,_xlpm.x*2,_xlpm.x+_xlpm.y)"
      "MAP(A2:A5,LAMBDA(x,x*x))", "_xlfn.MAP(A2:A5,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*_xlpm.x))"
      "LET(f,LAMBDA(x,x*x),f(7))", "_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlpm.x*_xlpm.x),_xlpm.f(7))"
      "Table1[IFS]+CONCAT(A1,B1)", "Table1[IFS]+_xlfn.CONCAT(A1,B1)" ] do
    let got = XlsxWriter.fileFormula formula
    ok ("file formula: " + formula) (got = want) got

// ... and read back as Excel shows them; a prefixed name the table does not know keeps its prefix.
for file, model in
    [ "_xlfn.XLOOKUP(2,A1:A4,B1:B4)", "XLOOKUP(2,A1:A4,B1:B4)"
      "_xlfn._xlws.SORT(_xlfn._xlws.FILTER(A2:A7,A2:A7>3))", "SORT(FILTER(A2:A7,A2:A7>3))"
      "_xlfn.LET(_xlpm.x,5,_xlpm.x+1)", "LET(x,5,x+1)"
      "_xlfn.SOMEDAYFUNCTION(1)", "_xlfn.SOMEDAYFUNCTION(1)"
      "IF(A1=\"_xlfn.XLOOKUP(\",1,2)", "IF(A1=\"_xlfn.XLOOKUP(\",1,2)" ] do
    let got = XlsxWriter.modelFormula file
    ok ("model formula: " + file) (got = model && XlsxWriter.fileFormula got = file) got

let modernSheet = { syntheticSheet with Name = "Modern"; Tables = []; DataValidations = []; Hyperlinks = []; PrintArea = None; PrintTitleRows = None; PrintTitleCols = None
                                        CellSource = XlsxWriter.sheetCells [ { Row = 0; Col = 0; Value = XlsxWriter.XlsxFormula ("XLOOKUP(1,B1:B2,C1:C2)", None); Format = XlsxWriter.defaultXlsxCellFormat; Comment = None } ] }
let modernBytes = { synthetic with Sheets = [ modernSheet ]; DefinedNames = [ { Name = "Pick"; RefersTo = "TEXTJOIN(\",\",TRUE,Modern!$B$1:$B$2)"; LocalSheetIndex = None; Comment = None; Hidden = false } ]; VbaProjectSource = None } |> XlsxPackageWriter.writeWorkbookPackageBytes
let modernPart (name: string) =
    use archive = new System.IO.Compression.ZipArchive(new MemoryStream(modernBytes))
    use reader = new StreamReader(archive.GetEntry(name).Open())
    reader.ReadToEnd()
ok "written: a cell's XLOOKUP is _xlfn.XLOOKUP" ((modernPart "xl/worksheets/sheet1.xml").Contains "<f>_xlfn.XLOOKUP(1,B1:B2,C1:C2)</f>") (modernPart "xl/worksheets/sheet1.xml")
ok "written: a name's TEXTJOIN is _xlfn.TEXTJOIN" ((modernPart "xl/workbook.xml").Contains ">_xlfn.TEXTJOIN(") (modernPart "xl/workbook.xml")

printfn ""
printfn "OfficeFormat XLSX writer tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
