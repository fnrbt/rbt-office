#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"

open System
open System.Collections.Generic
open System.IO
open System.Text

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-30s  %s" label detail)

let sharedStringsXml entryCount runsPerEntry =
    let sb = StringBuilder()
    sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">") |> ignore
    for i in 1 .. entryCount do
        sb.Append("<si>") |> ignore
        if i % 2 = 0 then
            sb.Append("<t>plain ") |> ignore
            sb.Append(i) |> ignore
            sb.Append(" &amp; text</t>") |> ignore
        else
            for r in 1 .. runsPerEntry do
                sb.Append("<r>") |> ignore
                if r % 5 = 0 then
                    sb.Append("<rPr><b/><sz val=\"11\"/><rFont val=\"Calibri\"/></rPr>") |> ignore
                sb.Append("<t>") |> ignore
                sb.Append(char (int 'a' + (r % 26))) |> ignore
                sb.Append("</t></r>") |> ignore
        sb.Append("</si>") |> ignore
    sb.Append("</sst>") |> ignore
    sb.ToString()

let small = XlsxReader.Xlsx.sharedStrings (sharedStringsXml 2 6)
ok "shared-string-count" (small.Length = 2) (sprintf "%d" small.Length)
ok "shared-string-rich-text"
   (small.[0].Text = "bcdefg" && small.[0].Runs.Length = 6 && (small.[0].Runs |> List.exists (fun (_, props) -> props.Bold)))
   (sprintf "%A" small.[0])
ok "shared-string-plain-text"
   (small.[1].Text = "plain 2 & text" && small.[1].Runs.IsEmpty)
   (sprintf "%A" small.[1])

let preservedAndStyled =
    XlsxReader.Xlsx.sharedStrings
        """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t xml:space="preserve">  padded  </t></si><si><r><rPr><i/><u val="double"/><strike/><color rgb="FF112233"/><rFont val="Arial"/><sz val="12.5"/><vertAlign val="superscript"/><scheme val="minor"/></rPr><t>x</t></r></si></sst>"""
// <u val="none"/> says a run is not underlined; <u/> and <u val="single"/> say it is.
let underlinedRuns =
    XlsxReader.Xlsx.sharedStrings
        """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><u val="none"/></rPr><t>a</t></r><r><rPr><u/></rPr><t>b</t></r><r><rPr><u val="single"/></rPr><t>c</t></r><r><rPr><b val="0"/></rPr><t>d</t></r></si></sst>"""
ok "shared-string-underline-none"
   (underlinedRuns.[0].Runs |> List.map (fun (text, props) -> text, props.Underline, props.UnderlineDouble) = [ "a", false, false; "b", true, false; "c", true, false; "d", false, false ])
   (sprintf "%A" underlinedRuns.[0])
// A comment's text in runs: the runs are read with it.
let commentRuns =
    XlsxReader.CellComments.parseComments
        """<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>Ann</author></authors><commentList><comment ref="B2" authorId="0"><text><r><rPr><b/><sz val="9"/><rFont val="Tahoma"/></rPr><t>Ann:</t></r><r><rPr><sz val="9"/><color rgb="FFFF0000"/><rFont val="Tahoma"/></rPr><t xml:space="preserve"> red</t></r></text></comment><comment ref="C3" authorId="0"><text><t>plain</t></text></comment></commentList></comments>"""
let commentRunsRead = commentRuns |> List.map (fun c -> c.CellRef.Column, c.Text, c.Author, c.Runs |> List.map (fun (text, props) -> text, props.Bold, props.Color, props.FontName, props.FontSize))
ok "comment-runs"
   (commentRunsRead = [ "B", "Ann: red", Some "Ann", [ "Ann:", true, None, Some "Tahoma", Some 9.0; " red", false, Some "FF0000", Some "Tahoma", Some 9.0 ]; "C", "plain", Some "Ann", [] ])
   (sprintf "%A" commentRuns)
ok "shared-string-preserves-space"
   (preservedAndStyled.[0].Text = "  padded  ")
   (sprintf "%A" preservedAndStyled.[0])
ok "shared-string-rich-props"
   (match preservedAndStyled.[1].Runs with
    | [ "x", props ] ->
        props.Italic
        && props.Underline
        && props.UnderlineDouble
        && props.Strikethrough
        && props.Color = Some "112233"
        && props.FontName = Some "Arial"
        && props.FontSize = Some 12.5
        && props.VertAlign = XlsxReader.VertAlignSuperscript
        && props.Scheme = "minor"
    | _ -> false)
   (sprintf "%A" preservedAndStyled.[1])

let perfXml = sharedStringsXml 1000 24
XlsxReader.Xlsx.sharedStrings perfXml |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let before = GC.GetAllocatedBytesForCurrentThread()
let parsed = XlsxReader.Xlsx.sharedStrings perfXml
let allocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - before) / 1024.0 / 1024.0
ok "shared-string-linear-alloc"
   (parsed.Length = 1000 && allocatedMb < 16.0)
   (sprintf "entries=%d allocated=%.1f MB" parsed.Length allocatedMb)

let columnName oneBased =
    let mutable n = oneBased
    let mutable name = ""
    while n > 0 do
        let r = (n - 1) % 26
        name <- string (char (int 'A' + r)) + name
        n <- (n - 1) / 26
    name

let mkCell col row value : XlsxReader.Cell =
    {
        Reference = { Column = col; Row = row }
        Value = value
        Formula = None
        ArrayRef = None
        ArrayDynamic = false
        StyleIndex = None
        Style = None
        InlineRuns = []
    }

let mkRow index cells : XlsxReader.Row =
    {
        Index = index
        Cells = cells
        Height = None
        RawHeight = None
        CustomHeight = None
        Hidden = false
        OutlineLevel = 0
        Collapsed = false
        Style = None
        ThickEdges = 0
    }

let mkSheetWithColumns columns rows : XlsxReader.Worksheet =
    {
        Name = "Sheet1"
        State = XlsxReader.SheetVisible
        RowSource = XlsxReader.SheetRows.Of rows
        MergedCells = []
        Columns = columns
        Format = XlsxReader.Xlsx.emptySheetFormat
        Comments = []
        ConditionalFormats = []
        DataValidations = []
        Charts = []
        PivotTables = []
        Tables = []
        FrozenPane = None
        ShowZeros = true
        ShowGridLines = true
        View = XlsxReader.defaultSheetView
        Hyperlinks = []
        PageSetup = None
        PageMargins = None
        PrintOptions = XlsxReader.defaultPrintOptions
        HeaderFooter = XlsxReader.defaultHeaderFooter
        RowBreaks = []
        ColBreaks = []
        PrintArea = None
        PrintTitleRows = None
        PrintTitleCols = None
        Extras = XlsxReader.emptySheetExtras
    }

let mkSheet rows = mkSheetWithColumns [] rows

let mkWorkbook sheet : XlsxReader.Workbook =
    {
        Sheets = [ sheet ]
        SharedStrings = [| "shared" |]
        SharedStringRuns = [| [] |]
        Styles = [||]
        Dxfs = [||]
        Relationships = []
        NamedRanges = []
        ThemeAccents = [||]
        ThemeColors = [||]
        TableStyles = []
        SlicerStyles = XlsxReader.SlicerStyleDefinitions.none
        DefaultFont = None, None
        Date1904 = false
        Preserved = XlsxReader.emptyWorkbookPreservation
    }

let isString expected (value: obj) =
    match value with
    | :? string as actual -> actual = expected
    | _ -> false

let isFloat expected (value: obj) =
    match value with
    | :? float as actual -> actual = expected
    | _ -> false

let cellObjMap (value: obj) =
    match value with
    | :? IDictionary<string, obj> as fields -> Some fields
    | _ -> None

let objBool name (fields: IDictionary<string, obj>) =
    match fields.TryGetValue name with
    | true, (:? bool as value) -> value
    | _ -> false

let objString name (fields: IDictionary<string, obj>) =
    match fields.TryGetValue name with
    | true, (:? string as value) -> value
    | _ -> ""

let findCellObj ref (cells: obj array) =
    cells
    |> Array.tryPick (fun cellObj ->
        match cellObjMap cellObj with
        | Some fields when objString "ref" fields = ref -> Some fields
        | _ -> None)

let sparseSheet =
    mkSheet [
        mkRow 1 [ mkCell "B" 1 (XlsxReader.String "b1") ]
        mkRow 3 [ mkCell "A" 3 (XlsxReader.Number 7.0) ]
    ]
let sparseArray = XlsxReader.sheetToArray (mkWorkbook sparseSheet) sparseSheet
ok "sheet-to-array-sparse-shape"
   (sparseArray.Length = 3
    && sparseArray.[0].Length = 2
    && isString "" sparseArray.[0].[0]
    && isString "b1" sparseArray.[0].[1]
    && isString "" sparseArray.[1].[0]
    && isFloat 7.0 sparseArray.[2].[0])
   (sprintf "shape=%dx%d" sparseArray.Length (if sparseArray.Length = 0 then 0 else sparseArray.[0].Length))

let metadataSheet =
    { mkSheetWithColumns [
        { Index = 2; MaxIndex = 3; Width = Some 12.5; BestFit = false; Hidden = false; OutlineLevel = 0; Collapsed = false; Style = None }
        { Index = 5; MaxIndex = 5; Width = Some 9.0; BestFit = false; Hidden = true; OutlineLevel = 0; Collapsed = false; Style = None }
      ] [
        { mkRow 2 [] with Height = Some 22.0 }
      ] with
        Format = { XlsxReader.Xlsx.emptySheetFormat with DefaultColWidth = Some 10.0; DefaultRowHeight = Some 14.0 } }
let metadataWidths = Main.Excel.columnWidths 6 metadataSheet
let metadataHeights = Main.Excel.rowHeights 4 metadataSheet
ok "sheet-metadata-bulk-values"
   (metadataWidths.Length = 6
    && metadataWidths.[0] = 10.0
    && metadataWidths.[1] = 12.5
    && metadataWidths.[2] = 12.5
    && metadataWidths.[4] = 0.0
    && metadataHeights.Length = 4
    && metadataHeights.[0] = 14.0
    && metadataHeights.[1] = 22.0
    && metadataHeights.[3] = 14.0)
   (sprintf "widths=%A heights=%A" metadataWidths metadataHeights)

let heightRows = [ for r in 1 .. 5000 -> { mkRow r [] with Height = Some (float r) } ]
let heightSheet = mkSheet heightRows
Main.Excel.rowHeights heightRows.Length heightSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let heightsBefore = GC.GetAllocatedBytesForCurrentThread()
let heightsSw = System.Diagnostics.Stopwatch.StartNew()
let bulkHeights = Main.Excel.rowHeights heightRows.Length heightSheet
heightsSw.Stop()
let heightsAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - heightsBefore) / 1024.0 / 1024.0
ok "sheet-row-heights-linear"
   (bulkHeights.Length = heightRows.Length
    && bulkHeights.[0] = 1.0
    && bulkHeights.[bulkHeights.Length - 1] = float heightRows.Length
    && heightsSw.Elapsed.TotalMilliseconds < 100.0
    && heightsAllocatedMb < 1.0)
   (sprintf "elapsed=%.1fms allocated=%.2f MB" heightsSw.Elapsed.TotalMilliseconds heightsAllocatedMb)

let styledRows = 700
let styledCols = 20
let styledTargetRef = sprintf "%s%d" (columnName styledCols) styledRows
let styledRowsData =
    [ for r in 1 .. styledRows ->
        mkRow r [
            for c in 1 .. styledCols ->
                mkCell (columnName c) r (XlsxReader.String "x")
        ] ]
let styledSheet =
    { mkSheet styledRowsData with
        Comments =
            [ for r in 1 .. styledRows ->
                { CellRef = { Column = columnName styledCols; Row = r }
                  Author = None
                  Text = "note"
                  Runs = [] } ]
        DataValidations =
            [ for i in 1 .. 16 ->
                { Range =
                    if i = 16 then sprintf "$A$1:$%s$%d" (columnName styledCols) styledRows
                    else sprintf "$%s$1:$%s$%d" (columnName (styledCols + i)) (columnName (styledCols + i)) styledRows
                  ValidationType = "list"
                  Operator = None
                  Formula1 = Some "\"A,B\""
                  Formula2 = None
                  AllowBlank = true
                  ShowDropDown = true
                  ErrorMessage = None
                  ErrorTitle = None
                  ErrorStyle = None
                  ShowErrorMessage = false
                  Prompt = None
                  PromptTitle = None
                  ShowInputMessage = false } ] }
let styledWorkbook = mkWorkbook styledSheet
Main.Excel.cellsWithStyle styledWorkbook styledSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let styledBefore = GC.GetAllocatedBytesForCurrentThread()
let styledSw = System.Diagnostics.Stopwatch.StartNew()
let styledCells = Main.Excel.cellsWithStyle styledWorkbook styledSheet
styledSw.Stop()
let styledAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - styledBefore) / 1024.0 / 1024.0
let styledFirst = findCellObj "A1" styledCells
let styledLast = findCellObj styledTargetRef styledCells
ok "cells-with-style-indexed-flags"
   (match styledFirst, styledLast with
    | Some first, Some last ->
        not (objBool "hasComment" first)
        && objBool "hasListDropdown" first
        && objBool "hasComment" last
        && objBool "hasListDropdown" last
    | _ -> false)
   (sprintf "first=%A last=%A" styledFirst styledLast)
ok "cells-with-style-indexed-alloc"
   (styledCells.Length = styledRows * styledCols
    && styledSw.Elapsed.TotalMilliseconds < 1000.0
    && styledAllocatedMb < 80.0)
   (sprintf "cells=%d elapsed=%.1fms allocated=%.1f MB"
            styledCells.Length
            styledSw.Elapsed.TotalMilliseconds
            styledAllocatedMb)

let sheetArrayRows = 3000
let sheetArrayCols = 50
let denseRows =
    [ for r in 1 .. sheetArrayRows ->
        mkRow r [
            for c in 1 .. sheetArrayCols do
                let value =
                    if c = 1 then XlsxReader.SharedString 0
                    elif c = sheetArrayCols then XlsxReader.String (sprintf "R%dC%d" r c)
                    elif c % 17 = 0 then XlsxReader.Number (float (r + c))
                    else XlsxReader.String "x"
                mkCell (columnName c) r value
        ] ]
let denseSheet = mkSheet denseRows
let denseWorkbook = mkWorkbook denseSheet
XlsxReader.sheetToArray denseWorkbook denseSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let sheetBefore = GC.GetAllocatedBytesForCurrentThread()
let sheetSw = System.Diagnostics.Stopwatch.StartNew()
let denseArray = XlsxReader.sheetToArray denseWorkbook denseSheet
sheetSw.Stop()
let sheetAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - sheetBefore) / 1024.0 / 1024.0
ok "sheet-to-array-linear-alloc"
   (denseArray.Length = sheetArrayRows
    && denseArray.[0].Length = sheetArrayCols
    && isString "shared" denseArray.[0].[0]
    && isString (sprintf "R%dC%d" sheetArrayRows sheetArrayCols) denseArray.[sheetArrayRows - 1].[sheetArrayCols - 1]
    && sheetSw.Elapsed.TotalMilliseconds < 1000.0
    && sheetAllocatedMb < 10.0)
   (sprintf "shape=%dx%d elapsed=%.1fms allocated=%.1f MB"
            denseArray.Length
            (if denseArray.Length = 0 then 0 else denseArray.[0].Length)
            sheetSw.Elapsed.TotalMilliseconds
            sheetAllocatedMb)

let targetRef = sprintf "%s%d" (columnName sheetArrayCols) sheetArrayRows
Main.Excel.cellAt targetRef denseSheet |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let lookupBefore = GC.GetAllocatedBytesForCurrentThread()
let lookupSw = System.Diagnostics.Stopwatch.StartNew()
let lookupValue = Main.Excel.cellAt targetRef denseSheet
lookupSw.Stop()
let lookupAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - lookupBefore) / 1024.0 / 1024.0
ok "sheet-cell-at-targeted-alloc"
   (match lookupValue with
    | Some (XlsxReader.String value) ->
        value = sprintf "R%dC%d" sheetArrayRows sheetArrayCols
        && lookupSw.Elapsed.TotalMilliseconds < 20.0
        && lookupAllocatedMb < 0.5
    | _ -> false)
   (sprintf "ref=%s elapsed=%.1fms allocated=%.2f MB value=%A"
            targetRef
            lookupSw.Elapsed.TotalMilliseconds
            lookupAllocatedMb
            lookupValue)

let dashboardPath = Path.Combine(__SOURCE_DIRECTORY__, "..", "real-world", "dashboard-sales-analysis.xlsx")
let dashboardBytes = File.ReadAllBytes dashboardPath
XlsxReader.loadBytesNative dashboardBytes |> ignore
GC.Collect()
GC.WaitForPendingFinalizers()
GC.Collect()
let dashboardBefore = GC.GetAllocatedBytesForCurrentThread()
let dashboardSw = System.Diagnostics.Stopwatch.StartNew()
let dashboardWorkbook = XlsxReader.loadBytesNative dashboardBytes
dashboardSw.Stop()
let dashboardAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - dashboardBefore) / 1024.0 / 1024.0
let dashboardData =
    dashboardWorkbook
    |> Option.bind (fun wb -> wb.Sheets |> List.tryFind (fun sheet -> sheet.Name = "Data"))
let dashboardCellCount =
    dashboardData
    |> Option.map (fun sheet -> sheet.Rows |> List.sumBy (fun row -> row.Cells.Length))
    |> Option.defaultValue 0
ok "native-large-worksheet-load"
   (match dashboardWorkbook, dashboardData with
    | Some wb, Some data ->
        wb.Sheets.Length = 8
        && data.Rows.Length = 76931
        && dashboardCellCount = 846241
        && data.Columns.Length = 12
        && data.Tables.Length = 1
        && (match data.PageSetup with Some setup -> setup.Orientation = XlsxReader.OrientationPortrait | None -> false)
        && dashboardAllocatedMb < 1600.0
    | _ -> false)
   (sprintf "allocated=%.1f MB elapsed=%.1fms cells=%d"
            dashboardAllocatedMb
            dashboardSw.Elapsed.TotalMilliseconds
            dashboardCellCount)

let objInt name (fields: IDictionary<string, obj>) =
    match fields.TryGetValue name with
    | true, (:? int as value) -> value
    | true, (:? float as value) -> int value
    | _ -> 0

let objArray<'T> name (fields: IDictionary<string, obj>) =
    match fields.TryGetValue name with
    | true, (:? ('T[]) as values) -> Some values
    | _ -> None

ok "sheet-dimensions-dashboard"
   (match dashboardData with
    | Some data ->
        match Main.Excel.dimensions data |> cellObjMap with
        | Some dims -> objInt "rowCount" dims = 76931 && objInt "colCount" dims = 11
        | None -> false
    | None -> false)
   (sprintf "data=%A" (dashboardData |> Option.map (fun data -> data.Rows.Length)))

match dashboardWorkbook, dashboardData with
| Some wb, Some data ->
    Main.Excel.cellsWithStyle wb data |> ignore
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    let dashboardStyledBefore = GC.GetAllocatedBytesForCurrentThread()
    let dashboardStyledSw = System.Diagnostics.Stopwatch.StartNew()
    let dashboardStyledCells = Main.Excel.cellsWithStyle wb data
    dashboardStyledSw.Stop()
    let dashboardStyledAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - dashboardStyledBefore) / 1024.0 / 1024.0
    ok "cells-with-style-dashboard-alloc"
       (dashboardStyledCells.Length = dashboardCellCount
        && dashboardStyledAllocatedMb < 1600.0)
       (sprintf "cells=%d allocated=%.1f MB elapsed=%.1fms"
                dashboardStyledCells.Length
                dashboardStyledAllocatedMb
                dashboardStyledSw.Elapsed.TotalMilliseconds)
| _ ->
    ok "cells-with-style-dashboard-alloc" false "dashboard workbook missing"

match dashboardWorkbook, dashboardData with
| Some wb, Some data ->
    Main.Excel.cellsCompact wb data |> ignore
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    let dashboardCompactBefore = GC.GetAllocatedBytesForCurrentThread()
    let dashboardCompactSw = System.Diagnostics.Stopwatch.StartNew()
    let dashboardCompactCells = Main.Excel.cellsCompact wb data
    dashboardCompactSw.Stop()
    let dashboardCompactAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - dashboardCompactBefore) / 1024.0 / 1024.0
    let firstCompact =
        dashboardCompactCells
        |> Array.tryHead
        |> Option.bind (fun value ->
            match value with
            | :? (obj[]) as slots -> Some slots
            | _ -> None)
    ok "cells-compact-dashboard-alloc"
       (dashboardCompactCells.Length = dashboardCellCount
        && dashboardCompactAllocatedMb < 400.0
        && (match firstCompact with
            | Some slots when slots.Length = 8 ->
                match slots.[0], slots.[1], slots.[4] with
                | (:? int as row), (:? int as col), (:? int as styleIndex) ->
                    row = 0 && col = 0 && styleIndex >= 0
                | _ -> false
            | _ -> false))
       (sprintf "cells=%d allocated=%.1f MB elapsed=%.1fms first=%A"
                dashboardCompactCells.Length
                dashboardCompactAllocatedMb
                dashboardCompactSw.Elapsed.TotalMilliseconds
                firstCompact)

    let firstStyle = Main.Excel.styleAt 0 wb |> cellObjMap
    ok "style-at-dashboard"
       (match firstStyle with
        | Some style -> objString "fontName" style <> ""
        | None -> false)
       (sprintf "style=%A" firstStyle)

    Main.Excel.cellsColumnar wb data |> ignore
    GC.Collect()
    GC.WaitForPendingFinalizers()
    GC.Collect()
    let dashboardColumnarBefore = GC.GetAllocatedBytesForCurrentThread()
    let dashboardColumnarSw = System.Diagnostics.Stopwatch.StartNew()
    let dashboardColumnarCells = Main.Excel.cellsColumnar wb data
    dashboardColumnarSw.Stop()
    let dashboardColumnarAllocatedMb = float (GC.GetAllocatedBytesForCurrentThread() - dashboardColumnarBefore) / 1024.0 / 1024.0
    let columnarFields = dashboardColumnarCells |> cellObjMap
    ok "cells-columnar-dashboard-alloc"
       (match columnarFields with
        | Some fields ->
            match objArray<int> "rows" fields, objArray<int> "cols" fields, objArray<int> "styleIndexes" fields, objArray<obj> "values" fields with
            | Some rows, Some cols, Some styleIndexes, Some values ->
                rows.Length = dashboardCellCount
                && cols.Length = dashboardCellCount
                && styleIndexes.Length = dashboardCellCount
                && values.Length = dashboardCellCount
                && rows.[0] = 0
                && cols.[0] = 0
                && styleIndexes.[0] >= 0
                && dashboardColumnarAllocatedMb < 260.0
            | _ -> false
        | None -> false)
       (sprintf "allocated=%.1f MB elapsed=%.1fms fields=%A"
                dashboardColumnarAllocatedMb
                dashboardColumnarSw.Elapsed.TotalMilliseconds
                (columnarFields |> Option.map (fun fields -> fields.Count)))
| _ ->
    ok "cells-compact-dashboard-alloc" false "dashboard workbook missing"
    ok "style-at-dashboard" false "dashboard workbook missing"
    ok "cells-columnar-dashboard-alloc" false "dashboard workbook missing"

// ----- theme, phonetic text and relationship targets ------------------------

let themeXml =
    """<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="x"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="x"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Georgia"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>"""
let themeColors = XlsxReader.Xlsx.parseThemeColors themeXml
ok "theme-colors-prefixed"
   (themeColors.[0] = "FEFEFE" && themeColors.[1] = "111111" && themeColors.[4] = "4F81BD" && themeColors.[9] = "F79646")
   (sprintf "%A" themeColors)
ok "theme-accents-prefixed"
   ((XlsxReader.Xlsx.parseThemeAccents themeXml).[0] = "4F81BD")
   (sprintf "%A" (XlsxReader.Xlsx.parseThemeAccents themeXml))
ok "theme-fonts-prefixed"
   (XlsxReader.Xlsx.parseThemeFonts themeXml = (Some "Cambria", Some "Georgia"))
   (sprintf "%A" (XlsxReader.Xlsx.parseThemeFonts themeXml))

let phonetic =
    XlsxReader.Xlsx.sharedStrings
        """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="1"/></si></sst>"""
ok "shared-string-skips-phonetic" (phonetic.[0].Text = "漢字") phonetic.[0].Text

let xlsxWithSheetTarget (target: string) =
    use stream = new MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        let add (name: string) (text: string) =
            use writer = new StreamWriter(archive.CreateEntry(name).Open(), UTF8Encoding(false))
            writer.Write text
        add "[Content_Types].xml" """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"""
        add "xl/workbook.xml" """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"""
        add "xl/_rels/workbook.xml.rels" (sprintf """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="%s"/></Relationships>""" target)
        add "xl/worksheets/sheet1.xml" """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>42</v></c></row></sheetData></worksheet>"""
    stream.ToArray()

for label, target in [ "relative", "worksheets/sheet1.xml"; "absolute", "/xl/worksheets/sheet1.xml"; "dotted", "../xl/worksheets/sheet1.xml" ] do
    let sheets =
        XlsxReader.loadBytesNative (xlsxWithSheetTarget target)
        |> Option.map (fun wb -> wb.Sheets |> List.map (fun sheet -> sheet.Name, sheet.Rows.Length))
    ok (sprintf "sheet-target-%s" label) (sheets = Some [ "Data", 1 ]) (sprintf "%A" sheets)

ok "resolve-target"
   (ZipInterop.resolveTarget "word" "media/a.png" = "word/media/a.png"
    && ZipInterop.resolveTarget "word" "/word/media/a.png" = "word/media/a.png"
    && ZipInterop.resolveTarget "word" "../customXml/item1.xml" = "customXml/item1.xml"
    && ZipInterop.resolveTarget "" "word/document.xml" = "word/document.xml")
   (ZipInterop.resolveTarget "word" "/word/media/a.png")

// ----- synthetic packages ----------------------------------------------------

let ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
let relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
let pkgRelNs = "http://schemas.openxmlformats.org/package/2006/relationships"

let zipOf (parts: (string * string) list) =
    use stream = new MemoryStream()
    do
        use archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Create, true)
        for name, text in parts do
            use writer = new StreamWriter(archive.CreateEntry(name).Open(), UTF8Encoding(false))
            writer.Write text
    stream.ToArray()

let workbookXml (extra: string) (sheets: (string * string) list) =
    let sheetXml =
        sheets
        |> List.mapi (fun i (name, attrs) -> sprintf "<sheet name=\"%s\" sheetId=\"%d\" r:id=\"rId%d\"%s/>" name (i + 1) (i + 1) attrs)
        |> String.concat ""
    sprintf "<workbook xmlns=\"%s\" xmlns:r=\"%s\">%s<sheets>%s</sheets></workbook>" ssNs relNs extra sheetXml

let workbookRels (count: int) (extra: string) =
    let rels =
        [ for i in 1 .. count ->
            sprintf "<Relationship Id=\"rId%d\" Type=\"%s/worksheet\" Target=\"worksheets/sheet%d.xml\"/>" i relNs i ]
        |> String.concat ""
    sprintf "<Relationships xmlns=\"%s\">%s%s</Relationships>" pkgRelNs rels extra

let worksheetXml (body: string) = sprintf "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">%s</worksheet>" ssNs relNs body

/// Single-sheet package. `pad` pushes the sheet XML past the 1,000,000-char
/// threshold so the streaming parser takes it instead of the DOM parser;
/// every test that touches cell parsing runs through both.
let onePackage (pad: bool) (wbExtra: string) (sheetBody: string) (extraParts: (string * string) list) =
    let padding = if pad then "<!--" + String('x', 1000100) + "-->" else ""
    zipOf ([ "xl/workbook.xml", workbookXml wbExtra [ "Data", "" ]
             "xl/_rels/workbook.xml.rels", workbookRels 1 ""
             "xl/worksheets/sheet1.xml", worksheetXml (padding + sheetBody) ] @ extraParts)

let loadOne pad wbExtra sheetBody extraParts =
    match XlsxReader.loadBytesNative (onePackage pad wbExtra sheetBody extraParts) with
    | Some wb -> wb
    | None -> failwith "package did not load"

let cellOf (wb: XlsxReader.Workbook) (ref: string) =
    wb.Sheets.Head.Rows
    |> List.collect (fun r -> r.Cells)
    |> List.tryFind (fun c -> sprintf "%s%d" c.Reference.Column c.Reference.Row = ref)

let bothPaths = [ "dom", false; "stream", true ]

// ----- formula text is kept next to the cached value -------------------------

let formulaBody =
    "<sheetData>"
    + "<row r=\"1\"><c r=\"A1\"><v>2</v></c><c r=\"B1\"><f>A1*2</f><v>4</v></c>"
    + "<c r=\"C1\"><f t=\"shared\" ref=\"C1:C3\" si=\"0\">A1*2+$A$1+Sheet2!B$1+SUM(A:A)+LOG10(A1)+LEN(\"A1\")+'A1'!A1+SUM(1:1)</f><v>10</v></c>"
    + "<c r=\"D1\"><f t=\"array\" ref=\"D1:D2\">A1:A2*2</f><v>4</v></c>"
    + "<c r=\"E1\"><f>1+1</f></c>"
    + "<c r=\"F1\" t=\"str\"><f>\"a\"&amp;\"b\"</f><v>ab</v></c>"
    + "<c r=\"G1\" t=\"e\"><f>1/0</f><v>#DIV/0!</v></c>"
    + "<c r=\"H1\"><f t=\"shared\" ref=\"H1:I1\" si=\"1\">A1+$B1+C$1+A:A</f><v>1</v></c><c r=\"I1\"><f t=\"shared\" si=\"1\"/><v>2</v></c></row>"
    + "<row r=\"2\"><c r=\"A2\"><v>3</v></c><c r=\"C2\"><f t=\"shared\" si=\"0\"/><v>12</v></c></row>"
    + "<row r=\"3\"><c r=\"C3\"><f t=\"shared\" si=\"0\"/><v>14</v></c></row>"
    + "</sheetData>"
for label, pad in bothPaths do
    let wb = loadOne pad "" formulaBody []
    let formula ref = cellOf wb ref |> Option.bind (fun c -> c.Formula)
    let value ref = cellOf wb ref |> Option.map (fun c -> c.Value)
    ok (sprintf "formula-keeps-cached-%s" label)
       (formula "B1" = Some "A1*2" && value "B1" = Some (XlsxReader.Number 4.0))
       (sprintf "%A" (cellOf wb "B1"))
    ok (sprintf "formula-shared-expanded-%s" label)
       (formula "C2" = Some "A2*2+$A$1+Sheet2!B$1+SUM(A:A)+LOG10(A2)+LEN(\"A1\")+'A1'!A2+SUM(2:2)"
        && formula "C3" = Some "A3*2+$A$1+Sheet2!B$1+SUM(A:A)+LOG10(A3)+LEN(\"A1\")+'A1'!A3+SUM(3:3)"
        && value "C2" = Some (XlsxReader.Number 12.0))
       (sprintf "%A / %A" (formula "C2") (formula "C3"))
    ok (sprintf "formula-shared-across-columns-%s" label)
       (formula "I1" = Some "B1+$B1+D$1+B:B")
       (sprintf "%A" (formula "I1"))
    ok (sprintf "formula-array-%s" label)
       (formula "D1" = Some "A1:A2*2" && (cellOf wb "D1" |> Option.bind (fun c -> c.ArrayRef)) = Some "D1:D2")
       (sprintf "%A" (cellOf wb "D1"))
    ok (sprintf "formula-uncached-%s" label)
       (formula "E1" = Some "1+1" && value "E1" = Some (XlsxReader.Formula "1+1"))
       (sprintf "%A" (cellOf wb "E1"))
    ok (sprintf "formula-string-and-error-cached-%s" label)
       (value "F1" = Some (XlsxReader.String "ab") && formula "F1" = Some "\"a\"&\"b\""
        && value "G1" = Some (XlsxReader.Error "DIV/0!") && formula "G1" = Some "1/0")
       (sprintf "%A / %A" (cellOf wb "F1") (cellOf wb "G1"))

// ----- rows and cells without `r` --------------------------------------------

let noRefBody =
    "<sheetData><row><c><v>1</v></c><c><v>2</v></c></row>"
    + "<row r=\"5\"><c r=\"C5\"><v>3</v></c><c><v>4</v></c></row>"
    + "<row><c t=\"inlineStr\"><is><t>x</t></is></c></row></sheetData>"
for label, pad in bothPaths do
    let wb = loadOne pad "" noRefBody []
    let cells =
        wb.Sheets.Head.Rows
        |> List.collect (fun r -> r.Cells |> List.map (fun c -> sprintf "%s%d" c.Reference.Column c.Reference.Row, c.Value))
    ok (sprintf "positions-inferred-%s" label)
       (cells = [ "A1", XlsxReader.Number 1.0; "B1", XlsxReader.Number 2.0
                  "C5", XlsxReader.Number 3.0; "D5", XlsxReader.Number 4.0
                  "A6", XlsxReader.String "x" ]
        && (wb.Sheets.Head.Rows |> List.map (fun r -> r.Index)) = [ 1; 5; 6 ])
       (sprintf "%A" cells)

// ----- 1904 date system -------------------------------------------------------

let dateStyles =
    sprintf "<styleSheet xmlns=\"%s\"><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"14\" applyNumberFormat=\"1\"/><xf numFmtId=\"4\" applyNumberFormat=\"1\"/></cellXfs></styleSheet>" ssNs
// A font's <b val="0"/> and <u val="none"/> say what the font is not.
let flagStyles =
    sprintf "<styleSheet xmlns=\"%s\"><fonts count=\"3\"><font/><font><b val=\"0\"/><i val=\"false\"/><strike val=\"0\"/><u val=\"none\"/></font><font><b/><i val=\"1\"/><strike/><u/></font></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"0\" fontId=\"1\" applyFont=\"1\"/><xf numFmtId=\"0\" fontId=\"2\" applyFont=\"1\"/></cellXfs></styleSheet>" ssNs
for label, pad in bothPaths do
    let wb = loadOne pad "<workbookPr/>" "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>1</v></c><c r=\"B1\" s=\"2\"><v>2</v></c></row></sheetData>" [ "xl/styles.xml", flagStyles ]
    let flags index = wb.Styles.[index] |> fun style -> style.Bold, style.Italic, style.Underline, style.Strikethrough
    ok (sprintf "font-flags-off-%s" label) (flags 1 = (false, false, false, false) && flags 2 = (true, true, true, true)) (sprintf "%A %A" (flags 1) (flags 2))
let dateBody = "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>42000</v></c><c r=\"B1\"><v>42000</v></c><c r=\"C1\" s=\"2\"><v>42000</v></c><c r=\"D1\" s=\"1\"><f>A1+1</f><v>42001</v></c></row></sheetData>"
for label, pad in bothPaths do
    let wb1904 = loadOne pad "<workbookPr date1904=\"1\"/>" dateBody [ "xl/styles.xml", dateStyles ]
    let wb1900 = loadOne pad "<workbookPr/>" dateBody [ "xl/styles.xml", dateStyles ]
    let v wb ref = cellOf wb ref |> Option.map (fun c -> c.Value)
    // (The serials of a 1904 workbook are the file's - date-formatted or not: `Date1904` says what they mean.)
    ok (sprintf "date1904-serials-kept-%s" label)
       (wb1904.Date1904
        && v wb1904 "A1" = Some (XlsxReader.Number 42000.0)
        && v wb1904 "D1" = Some (XlsxReader.Number 42001.0)
        && v wb1904 "B1" = Some (XlsxReader.Number 42000.0)
        && v wb1904 "C1" = Some (XlsxReader.Number 42000.0))
       (sprintf "%A %A %A %A" (v wb1904 "A1") (v wb1904 "B1") (v wb1904 "C1") (v wb1904 "D1"))
    ok (sprintf "date1900-untouched-%s" label)
       (not wb1900.Date1904 && v wb1900 "A1" = Some (XlsxReader.Number 42000.0))
       (sprintf "%A" (v wb1900 "A1"))

// ----- sheet state, hidden rows / columns ------------------------------------

let stateWorkbook =
    zipOf [ "xl/workbook.xml", workbookXml "" [ "Shown", ""; "Hidden", " state=\"hidden\""; "Deep", " state=\"veryHidden\"" ]
            "xl/_rels/workbook.xml.rels", workbookRels 3 ""
            "xl/worksheets/sheet1.xml", worksheetXml "<cols><col min=\"2\" max=\"3\" width=\"9\" hidden=\"true\"/></cols><sheetData><row r=\"1\" hidden=\"1\"><c r=\"A1\"><v>1</v></c></row><row r=\"2\"><c r=\"A2\"><v>2</v></c></row></sheetData>"
            "xl/worksheets/sheet2.xml", worksheetXml "<sheetData/>"
            "xl/worksheets/sheet3.xml", worksheetXml "<sheetData/>" ]
match XlsxReader.loadBytesNative stateWorkbook with
| Some wb ->
    ok "sheet-state"
       ((wb.Sheets |> List.map (fun s -> s.Name, s.State)) = [ "Shown", XlsxReader.SheetVisible; "Hidden", XlsxReader.SheetHidden; "Deep", XlsxReader.SheetVeryHidden ])
       (sprintf "%A" (wb.Sheets |> List.map (fun s -> s.Name, s.State)))
    ok "sheet-visible-names"
       (Main.Excel.visibleSheetNames wb = [| "Shown" |] && Main.Excel.sheetStates wb = [| "visible"; "hidden"; "veryHidden" |])
       (sprintf "%A %A" (Main.Excel.visibleSheetNames wb) (Main.Excel.sheetStates wb))
    let shown = wb.Sheets.Head
    ok "hidden-rows-and-columns"
       ((shown.Rows |> List.map (fun r -> r.Hidden)) = [ true; false ]
        && (shown.Columns |> List.map (fun c -> c.Index, c.MaxIndex, c.Hidden)) = [ 2, 3, true ])
       (sprintf "%A %A" shown.Rows shown.Columns)
| None -> ok "sheet-state" false "package did not load"

// ----- namespace prefixes and package relationships --------------------------

let prefixedWorkbook =
    zipOf [ "_rels/.rels", sprintf "<pr:Relationships xmlns:pr=\"%s\"><pr:Relationship Id=\"rId1\" Type=\"%s/officeDocument\" Target=\"/book/main.xml\"/></pr:Relationships>" pkgRelNs relNs
            "book/main.xml", sprintf "<x:workbook xmlns:x=\"%s\" xmlns:rel=\"%s\"><x:sheets><x:sheet name=\"Data\" sheetId=\"1\" rel:id=\"rId7\"/></x:sheets></x:workbook>" ssNs relNs
            "book/_rels/main.xml.rels", sprintf "<pr:Relationships xmlns:pr=\"%s\"><pr:Relationship Id=\"rId7\" Type=\"%s/worksheet\" Target=\"sheets/one.xml\"/><pr:Relationship Id=\"rId8\" Type=\"%s/sharedStrings\" Target=\"strings.xml\"/></pr:Relationships>" pkgRelNs relNs relNs
            "book/strings.xml", sprintf "<sst xmlns=\"%s\"><si><t>hello</t></si></sst>" ssNs
            "book/sheets/one.xml", sprintf "<x:worksheet xmlns:x=\"%s\" xmlns:rel=\"%s\"><x:sheetData><x:row r=\"1\"><x:c r=\"A1\" t=\"s\"><x:v>0</x:v></x:c></x:row></x:sheetData><x:hyperlinks><x:hyperlink ref=\"A1\" rel:id=\"rId1\"/></x:hyperlinks></x:worksheet>" ssNs relNs
            "book/sheets/_rels/one.xml.rels", sprintf "<Relationships xmlns=\"%s\"><Relationship Id=\"rId1\" Type=\"%s/hyperlink\" Target=\"https://example.com/\" TargetMode=\"External\"/></Relationships>" pkgRelNs relNs ]
match XlsxReader.loadBytesNative prefixedWorkbook with
| Some wb ->
    let sheet = wb.Sheets |> List.tryHead
    ok "prefixed-workbook-loads"
       ((sheet |> Option.map (fun s -> s.Name)) = Some "Data"
        && (sheet |> Option.bind (fun s -> s.Rows |> List.tryHead) |> Option.map (fun r -> XlsxReader.resolveValue wb r.Cells.Head.Value)) = Some (XlsxReader.String "hello"))
       (sprintf "%A" wb.Sheets)
    ok "prefixed-hyperlink-resolved"
       ((sheet |> Option.map (fun s -> s.Hyperlinks |> List.map (fun h -> h.Target))) = Some [ "https://example.com/" ])
       (sprintf "%A" (sheet |> Option.map (fun s -> s.Hyperlinks)))
| None -> ok "prefixed-workbook-loads" false "package did not load"

// ----- invariant-culture parsing ---------------------------------------------

let cultureBody = "<sheetData><row r=\"1\" ht=\"12.5\" customHeight=\"1\"><c r=\"A1\"><v>1.5</v></c><c r=\"B1\"><v>-2</v></c><c r=\"C1\" t=\"s\"><v>oops</v></c><c r=\"D1\" t=\"s\"><v>0.0</v></c></row></sheetData>"
let savedCulture = Globalization.CultureInfo.CurrentCulture
for cultureName in [ "de-DE"; "sv-SE" ] do
    for label, pad in bothPaths do
        Globalization.CultureInfo.CurrentCulture <- Globalization.CultureInfo cultureName
        let result =
            try Result.Ok (loadOne pad "" cultureBody [ "xl/sharedStrings.xml", sprintf "<sst xmlns=\"%s\"><si><t>zero</t></si></sst>" ssNs ])
            with ex -> Result.Error ex.Message
        Globalization.CultureInfo.CurrentCulture <- savedCulture
        match result with
        | Result.Ok wb ->
            let v ref = cellOf wb ref |> Option.map (fun c -> c.Value)
            ok (sprintf "culture-%s-%s" cultureName label)
               (v "A1" = Some (XlsxReader.Number 1.5)
                && v "B1" = Some (XlsxReader.Number -2.0)
                && v "C1" = Some XlsxReader.Empty
                && v "D1" = Some (XlsxReader.SharedString 0)
                && wb.Sheets.Head.Rows.Head.Height = Some 12.5)
               (sprintf "%A %A %A %A %A" (v "A1") (v "B1") (v "C1") (v "D1") wb.Sheets.Head.Rows.Head.Height)
        | Result.Error message -> ok (sprintf "culture-%s-%s" cultureName label) false message

// ----- comments: authors and rich runs ---------------------------------------

let commentsPart =
    sprintf "<comments xmlns=\"%s\"><authors><author>Ada</author><author>Grace</author></authors><commentList><comment ref=\"A1\" authorId=\"1\"><text><r><rPr><b/></rPr><t>Hel</t></r><r><t xml:space=\"preserve\">lo </t></r><r><t>there_x000D_</t></r></text></comment><comment ref=\"B2\" authorId=\"0\"><text><t>plain</t></text></comment></commentList></comments>" ssNs
let commentRels =
    sprintf "<Relationships xmlns=\"%s\"><Relationship Id=\"rId1\" Type=\"%s/comments\" Target=\"../notes/c.xml\"/></Relationships>" pkgRelNs relNs
for label, pad in bothPaths do
    let wb = loadOne pad "" "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>" [ "xl/notes/c.xml", commentsPart; "xl/worksheets/_rels/sheet1.xml.rels", commentRels ]
    let comments = wb.Sheets.Head.Comments |> List.map (fun c -> c.CellRef.Column, c.Author, c.Text)
    ok (sprintf "comments-author-and-runs-%s" label)
       (comments = [ "A", Some "Grace", "Hello there\r"; "B", Some "Ada", "plain" ])
       (sprintf "%A" comments)

// ----- data validation attributes, _xHHHH_ escapes ---------------------------

let dvBody =
    "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"inlineStr\"><is><t>x_x0009_y</t></is></c><c r=\"C1\" t=\"str\"><f>\"q\"</f><v>a_x000A_b</v></c></row></sheetData>"
    + "<dataValidations count=\"2\"><dataValidation type=\"whole\" errorStyle=\"warning\" operator=\"between\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" errorTitle=\"Bad\" error=\"1 to 5\" promptTitle=\"Pick\" prompt=\"A number\" sqref=\"A1\"><formula1>1</formula1><formula2>5</formula2></dataValidation>"
    + "<dataValidation type=\"list\" showDropDown=\"1\" sqref=\"B1\"><formula1>\"a,b\"</formula1></dataValidation></dataValidations>"
let dvWb = loadOne false "" dvBody [ "xl/sharedStrings.xml", sprintf "<sst xmlns=\"%s\"><si><t>line1_x000D_\nline2 _x005F_x0041_ _x0041_</t></si></sst>" ssNs ]
ok "validation-attributes"
   (match dvWb.Sheets.Head.DataValidations with
    | [ a; b ] ->
        a.ShowErrorMessage && a.ShowInputMessage && a.AllowBlank && a.ShowDropDown
        && a.ErrorTitle = Some "Bad" && a.ErrorMessage = Some "1 to 5" && a.ErrorStyle = Some "warning"
        && a.PromptTitle = Some "Pick" && a.Prompt = Some "A number"
        && not b.ShowErrorMessage && not b.ShowInputMessage && not b.AllowBlank && not b.ShowDropDown
    | _ -> false)
   (sprintf "%A" dvWb.Sheets.Head.DataValidations)
ok "xstring-escapes-decoded"
   (dvWb.SharedStrings.[0] = "line1\r\nline2 _x0041_ A"
    && (cellOf dvWb "B1" |> Option.map (fun c -> c.Value)) = Some (XlsxReader.String "x\ty")
    && (cellOf dvWb "C1" |> Option.map (fun c -> c.Value)) = Some (XlsxReader.String "a\nb"))
   (sprintf "%A %A %A" dvWb.SharedStrings.[0] (cellOf dvWb "B1") (cellOf dvWb "C1"))

// A table's filter without the columns whose button is hidden (templates that use a table for its look).
let tableXml (filter: string) =
    "<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"1\" name=\"T\" displayName=\"T\" ref=\"B5:D9\">" + filter
    + "<tableColumns count=\"3\"><tableColumn id=\"1\" name=\"a\"/><tableColumn id=\"2\" name=\"b\"/><tableColumn id=\"3\" name=\"c\"/></tableColumns></table>"
let filterOf filter = XlsxReader.Xlsx.parseTable (tableXml filter) |> Option.bind (fun t -> t.AutoFilter)
let buttons = [ filterOf "<autoFilter ref=\"B5:D9\"/>"
                filterOf "<autoFilter ref=\"B5:D9\"><filterColumn colId=\"0\" hiddenButton=\"1\"/></autoFilter>"
                filterOf "<autoFilter ref=\"B5:D9\"><filterColumn colId=\"0\" hiddenButton=\"1\"/><filterColumn colId=\"1\" hiddenButton=\"1\"/><filterColumn colId=\"2\" hiddenButton=\"1\"/></autoFilter>"
                filterOf "" ]
ok "table filter buttons: all, hidden on the first column, all hidden (no column), no filter"
   (buttons = [ Some (2, 4); Some (3, 4); Some (2, 1); None ]) (sprintf "%A" buttons)

// A run in a colour of the theme ("OPEN vs. CLOSED" titles): kept as "theme:index[:tint]" and
// resolved with the workbook's theme.
let themedRuns =
    XlsxReader.Xlsx.sharedStrings
        """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/><color theme="4"/></rPr><t>OPEN</t></r><r><rPr><color theme="1" tint="0.5"/></rPr><t> vs.</t></r></si></sst>"""
let runColors = themedRuns.[0].Runs |> List.map (fun (_, props) -> props.Color)
let theme = [| "FFFFFF"; "000000"; "E7E6E6"; "44546A"; "4472C4" |]
ok "run colours of the theme: kept, then resolved"
   (runColors = [ Some "theme:4"; Some "theme:1:0.5" ]
    && (runColors |> List.map (Option.bind (XlsxReader.Xlsx.themeRunColor theme))) = [ Some "4472C4"; Some "808080" ])
   (sprintf "%A %A" runColors (runColors |> List.map (Option.bind (XlsxReader.Xlsx.themeRunColor theme))))

printfn ""
printfn "OfficeFormat XLSX reader tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
