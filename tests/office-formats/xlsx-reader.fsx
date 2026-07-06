#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Office.Formats/bin/Debug/netstandard2.0/OfficeFormat.dll"

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
    }

let mkSheetWithColumns columns rows : XlsxReader.Worksheet =
    {
        Name = "Sheet1"
        Rows = rows
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
        { Index = 2; MaxIndex = 3; Width = Some 12.5; Hidden = false; Style = None }
        { Index = 5; MaxIndex = 5; Width = Some 9.0; Hidden = true; Style = None }
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
                  Text = "note" } ]
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
                  ErrorMessage = None } ] }
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

printfn ""
printfn "OfficeFormat XLSX reader tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
