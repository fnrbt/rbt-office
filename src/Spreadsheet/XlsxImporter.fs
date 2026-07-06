module RedBat.Spreadsheet.XlsxImporter

open Fable.Core
open Fable.Core.JsInterop
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

/// <summary>
/// XLSX import adapter for office-format library integration.
/// Converts office-format workbook data into RedBat.Grid SpreadsheetModel.
/// </summary>

    // Type definitions matching office-format API structure
    type ExcelBorder = {
        style: string option  // "thin", "medium", "thick", "dashed", "dotted", "double"
        color: string option
    }

    type ExcelCell = {
        value: obj option
        formula: string option
        bold: bool option
        italic: bool option
        underline: bool option
        underlineDouble: bool option
        strikethrough: bool option
        fontName: string option
        fontColor: string option
        fillColor: string option
        fontSize: float option
        numberFormat: string option
        horizontalAlign: string option  // "left", "center", "right", "justify"
        verticalAlign: string option    // "top", "middle", "bottom"
        wrapText: bool option
        borderTop: ExcelBorder option
        borderBottom: ExcelBorder option
        borderLeft: ExcelBorder option
        borderRight: ExcelBorder option
        borderDiagonal: ExcelBorder option
        diagonalUp: bool option
        diagonalDown: bool option
        hasComment: bool option
        hasListDropdown: bool option
    }

    type ExcelRichRun = {
        text: string
        bold: bool option
        italic: bool option
        underline: bool option
        underlineDouble: bool option
        strikethrough: bool option
        color: string option
        fontName: string option
        fontSize: float option
        vertAlign: string option
    }

    type ExcelWorkbook = obj
    type ExcelSheet = obj

    /// <summary>
    /// Import interface to office-format library (loaded at runtime).
    /// Expected API structure from office-format:
    ///   api().excel.load(buffer) -> workbook
    ///   api().excel.sheetNames(workbook) -> string[]
    ///   api().excel.sheet(name, workbook) -> sheet
    ///   api().excel.toArray(workbook, sheet) -> any[][]
    ///   api().excel.cellWithStyleAt(ref, sheet) -> CellWithStyle
    ///   api().excel.columnWidth(col, sheet) -> number
    ///   api().excel.rowHeight(row, sheet) -> number
    ///   api().excel.mergedCells(sheet) -> string[] (e.g., ["A1:B2", "C3:D5"])
    ///   api().excel.commentAt(ref, sheet) -> {text: string, author?: string} | null
    /// </summary>
    // Path expressed relative to the F# source dir, but chosen so Fable's
    // source→output rewrite produces "../office-formats/src/Main.js" in the
    // emitted JS — i.e., a sibling lookup inside dist/. No redbat reference.
    [<ImportAll("../../dist/office-formats/src/Main.js")>]
let private officeFormatApi: obj = jsNative

[<Emit("Array.from($0)")>]
let private arrayFrom (items: obj): obj[] = jsNative

[<Emit("typeof $0 === 'function'")>]
let private isFunction (value: obj): bool = jsNative

[<Emit("$0 == null")>]
let private isNullOrUndefined (value: obj): bool = jsNative

[<Emit("$0 == null ? null : $0[$1]")>]
let private getProp (value: obj) (name: string): obj = jsNative

let private getIndex (value: obj) (index: int): obj =
    getProp value (string index)

let private optStringProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (string value)

let private optBoolProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<bool> value)

let private optFloatProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<float> value)

let private optIntProp source name =
    let value = getProp source name
    if isNullOrUndefined value then None else Some (unbox<int> value)

let private hexColor (value: string) =
    let trimmed = value.TrimStart('#')
    let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
    if cleaned.Length = 6 then Some ("#" + cleaned) else None

    /// <summary>
    /// Convert office-format cell value to RedBat.Grid CellValue.
    /// </summary>
let private convertCellValue (excelCell: ExcelCell): CellValue =
        match excelCell.formula with
        | Some formula when not (System.String.IsNullOrWhiteSpace(formula)) ->
            // Formula cell: store formula text and empty cached value
            // (RedBat.Grid will re-evaluate the formula)
            CellValue.Formula (formula, CellValue.Empty)
        | _ ->
            match excelCell.value with
            | None -> CellValue.Empty
            | Some v ->
                match v with
                | :? float as n -> CellValue.Number n
                | :? int as i -> CellValue.Number (float i)
                | :? bool as b -> CellValue.Boolean b
                | :? string as s -> CellValue.String s
                | _ -> CellValue.String (string v)

    /// <summary>
    /// Convert ExcelBorder to BorderStyle.
    /// </summary>
let private convertBorder (border: ExcelBorder option): BorderStyle option =
        border |> Option.bind (fun b ->
            match b.style with
            | Some style when not (System.String.IsNullOrWhiteSpace(style)) ->
                Some { Style = style; Color = b.color |> Option.bind hexColor }
            | _ -> None
        )

    /// <summary>
    /// Convert office-format cell style to RedBat.Spreadsheet CellFormat.
    /// </summary>
let private convertCellFormat (excelCell: ExcelCell): CellFormat =
        {
            Bold = excelCell.bold |> Option.defaultValue false
            Italic = excelCell.italic |> Option.defaultValue false
            Underline = excelCell.underline |> Option.defaultValue false
            UnderlineDouble = excelCell.underlineDouble |> Option.defaultValue false
            FontName = excelCell.fontName
            TextColor = excelCell.fontColor |> Option.bind hexColor
            BgColor = excelCell.fillColor |> Option.bind hexColor
            FontSize = excelCell.fontSize |> Option.map int  // Convert float to int
            HorizontalAlign = excelCell.horizontalAlign
            VerticalAlign = excelCell.verticalAlign
            WrapText = excelCell.wrapText |> Option.defaultValue false
            Indent = 0
            ShrinkToFit = false
            TextRotation = 0
            NumberFormat = excelCell.numberFormat
            BorderTop = convertBorder excelCell.borderTop
            BorderBottom = convertBorder excelCell.borderBottom
            BorderLeft = convertBorder excelCell.borderLeft
            BorderRight = convertBorder excelCell.borderRight
            BorderDiagonal = convertBorder excelCell.borderDiagonal
            DiagonalUp = excelCell.diagonalUp |> Option.defaultValue false
            DiagonalDown = excelCell.diagonalDown |> Option.defaultValue false
            FilterButton = false
            Strikethrough = excelCell.strikethrough |> Option.defaultValue false
            HasComment = excelCell.hasComment |> Option.defaultValue false
            HasListDropdown = excelCell.hasListDropdown |> Option.defaultValue false
            DataBar = None
            IconSet = None
        }

let private equalsIgnoreCase (left: string) (right: string) =
        if left.Length <> right.Length then
            false
        else
            let mutable index = 0
            let mutable equal = true
            while equal && index < left.Length do
                let l = left.[index]
                let r = right.[index]
                if l <> r && System.Char.ToUpperInvariant l <> System.Char.ToUpperInvariant r then
                    equal <- false
                index <- index + 1
            equal

let private resizeArrayToList (items: ResizeArray<'T>) =
        let mutable result = []
        for i = items.Count - 1 downto 0 do
            result <- items.[i] :: result
        result

let private convertRichRun (runObj: obj): string * RunOverride =
        let vertAlign =
            match optStringProp runObj "vertAlign" with
            | Some value when equalsIgnoreCase value "superscript" -> Some VertAlignSuperscript
            | Some value when equalsIgnoreCase value "subscript" -> Some VertAlignSubscript
            | _ -> None
        let text = optStringProp runObj "text" |> Option.defaultValue ""
        text, {
            Bold = optBoolProp runObj "bold"
            Italic = optBoolProp runObj "italic"
            Underline = optBoolProp runObj "underline"
            UnderlineDouble = optBoolProp runObj "underlineDouble"
            Strikethrough = optBoolProp runObj "strikethrough"
            Color = optStringProp runObj "color" |> Option.bind hexColor
            FontName = optStringProp runObj "fontName"
            FontSize = optFloatProp runObj "fontSize"
            VertAlign = vertAlign
        }

    /// <summary>
    /// Convert Excel column letter (A, B, ..., Z, AA, AB, ...) to zero-based index.
    /// </summary>
let private columnLetterToIndex (letter: string): int =
        let mutable result = 0
        for c in letter do
            let c = System.Char.ToUpperInvariant c
            result <- result * 26 + (int c - int 'A' + 1)
        result - 1

let private parseCellRefSpan (text: string) startIndex endExclusive =
        let mutable index = startIndex
        let mutable col = 0
        let mutable row = 0
        let mutable hasCol = false
        let mutable hasRow = false
        let mutable inRow = false
        let mutable valid = startIndex < endExclusive
        while valid && index < endExclusive do
            let c = text.[index]
            if System.Char.IsDigit c then
                inRow <- true
                hasRow <- true
                let digit = int c - int '0'
                if row > (System.Int32.MaxValue - digit) / 10 then
                    valid <- false
                else
                    row <- row * 10 + digit
            elif System.Char.IsLetter c && not inRow then
                hasCol <- true
                let upper = System.Char.ToUpperInvariant c
                let digit = int upper - int 'A' + 1
                if col > (System.Int32.MaxValue - digit) / 26 then
                    valid <- false
                else
                    col <- col * 26 + digit
            else
                valid <- false
            index <- index + 1
        if valid && hasCol && hasRow && row > 0 && col > 0 then
            Some (row - 1, col - 1)
        else
            None

let private parseMergedCellRange (refRange: string) =
        let colon = refRange.IndexOf(':')
        if colon >= 0 && colon = refRange.LastIndexOf(':') then
            match parseCellRefSpan refRange 0 colon, parseCellRefSpan refRange (colon + 1) refRange.Length with
            | Some (startRow, startCol), Some (endRow, endCol) ->
                Some {
                    StartRow = startRow
                    StartCol = startCol
                    EndRow = endRow
                    EndCol = endCol
                }
            | _ -> None
        else
            None

    /// <summary>
    /// Import a sheet from office-format workbook into SpreadsheetModel.
    /// </summary>
    /// <param name="workbook">office-format workbook object</param>
    /// <param name="sheetName">Name of sheet to import (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importSheet (workbook: ExcelWorkbook) (sheetName: string option): SpreadsheetModel =
        try
            // Get API interface
            let api = officeFormatApi?``default``()
            let excel = api?excel

            // Get sheet (use provided name or first sheet)
            let actualSheetName =
                match sheetName with
                | Some name -> name
                | None ->
                    let sheetNames =
                        excel?sheetNames(workbook)
                        |> arrayFrom
                        |> Array.map string
                    if sheetNames.Length > 0 then sheetNames.[0]
                    else failwith "Workbook has no sheets"

            let sheet = excel?sheet(actualSheetName, workbook)

            let mutable rowCount = 0
            let mutable colCount = 0
            let setDimensionsFromSheet () =
                if isFunction excel?dimensions then
                    let dims = excel?dimensions(sheet)
                    rowCount <- optIntProp dims "rowCount" |> Option.defaultValue 0
                    colCount <- optIntProp dims "colCount" |> Option.defaultValue 0

            // Build cell map with styles
            let mutable cellMap = Map.empty<CellAddress, CellData>
            let mutable richRunsMap = Map.empty<CellAddress, (string * RunOverride) list>
            let mutable columnWidths = Map.empty<int, float>
            let mutable rowHeights = Map.empty<int, float>
            let mutable formulaCount = 0

            let noteImportedValue value =
                match value with
                | CellValue.Formula _ -> formulaCount <- formulaCount + 1
                | _ -> ()

            let extractBorder (borderObj: obj) =
                if isNullOrUndefined borderObj then None
                else
                    Some {
                        style = optStringProp borderObj "style"
                        color = optStringProp borderObj "color"
                    }

            let excelCellFromObj rawValue cellObj formulaOverride hasCommentOverride hasListDropdownOverride =
                {
                    value = if isNullOrUndefined rawValue || string rawValue = "" then None else Some rawValue
                    formula =
                        match formulaOverride with
                        | Some value -> Some value
                        | None -> optStringProp cellObj "formula"
                    bold = optBoolProp cellObj "bold"
                    italic = optBoolProp cellObj "italic"
                    underline = optBoolProp cellObj "underline"
                    underlineDouble = optBoolProp cellObj "underlineDouble"
                    strikethrough = optBoolProp cellObj "strikethrough"
                    fontName = optStringProp cellObj "fontName"
                    fontColor = optStringProp cellObj "fontColor"
                    fillColor = optStringProp cellObj "fillColor"
                    fontSize = optFloatProp cellObj "fontSize"
                    numberFormat = optStringProp cellObj "numberFormat"
                    horizontalAlign = optStringProp cellObj "horizontalAlign"
                    verticalAlign = optStringProp cellObj "verticalAlign"
                    wrapText = optBoolProp cellObj "wrapText"
                    borderTop = extractBorder (getProp cellObj "borderTop")
                    borderBottom = extractBorder (getProp cellObj "borderBottom")
                    borderLeft = extractBorder (getProp cellObj "borderLeft")
                    borderRight = extractBorder (getProp cellObj "borderRight")
                    borderDiagonal = extractBorder (getProp cellObj "borderDiagonal")
                    diagonalUp = optBoolProp cellObj "diagonalUp"
                    diagonalDown = optBoolProp cellObj "diagonalDown"
                    hasComment =
                        match hasCommentOverride with
                        | Some value -> Some value
                        | None -> optBoolProp cellObj "hasComment"
                    hasListDropdown =
                        match hasListDropdownOverride with
                        | Some value -> Some value
                        | None -> optBoolProp cellObj "hasListDropdown"
                }

            let importCell row col excelRef rawValue cellObj directRichRuns knownHasComment =
                if not (isNullOrUndefined cellObj) || not (isNullOrUndefined rawValue) then
                    let excelCell: ExcelCell =
                        excelCellFromObj rawValue cellObj None None None

                    let cellValue = convertCellValue excelCell
                    let cellFormat = convertCellFormat excelCell

                    // Read comment if present. Older office-format builds expose comments as a collection only.
                    let commentObj =
                        match knownHasComment with
                        | Some false -> null
                        | _ ->
                            if isFunction excel?commentAt then excel?commentAt(excelRef, sheet)
                            else null
                    let comment =
                        if isNullOrUndefined commentObj then None
                        else
                            let text = optStringProp commentObj "text" |> Option.defaultValue ""
                            let author = optStringProp commentObj "author"
                            if System.String.IsNullOrWhiteSpace(text) then None
                            else Some { Text = text; Author = author }

                    let cellFormat =
                        if comment.IsSome then { cellFormat with HasComment = true }
                        else cellFormat

                    // Keep style-only cells. Excel renders fills, borders,
                    // comment markers, and validation dropdowns even when
                    // the cell has no value.
                    if cellValue <> CellValue.Empty || cellFormat <> defaultFormat || comment.IsSome then
                        noteImportedValue cellValue
                        let addr = { Row = row; Col = col }
                        cellMap <- Map.add addr { Value = cellValue; Format = cellFormat; Comment = comment } cellMap

                        let richRunObjs =
                            match directRichRuns with
                            | Some runs -> runs
                            | None ->
                                if isFunction excel?richRunsAt then excel?richRunsAt(excelRef, workbook, sheet)
                                else null
                        if not (isNullOrUndefined richRunObjs) then
                            let richRunArray = arrayFrom richRunObjs
                            let runs = ResizeArray<string * RunOverride>(richRunArray.Length)
                            for runObj in richRunArray do
                                runs.Add(convertRichRun runObj)
                            if runs.Count > 0 then
                                let runs = resizeArrayToList runs
                                richRunsMap <- Map.add addr runs richRunsMap

            // Import actual cells with styles in one pass when the
            // office-format API supports it. The coordinate fallback is
            // kept for older browser bundles.
            let styleCache = System.Collections.Generic.Dictionary<int, CellFormat>()
            let styleFormat styleIndex =
                if styleIndex < 0 then
                    defaultFormat
                else
                    match styleCache.TryGetValue styleIndex with
                    | true, cached -> cached
                    | false, _ ->
                        let styleObj = excel?styleAt(styleIndex, workbook)
                        let excelCell = excelCellFromObj null styleObj None None None
                        let format = convertCellFormat excelCell
                        styleCache.[styleIndex] <- format
                        format

            let importCompactCell row col rawValue formulaObj styleIndex hasComment hasListDropdown richRunObjs =
                if row >= 0 && col >= 0 then
                    if row + 1 > rowCount then rowCount <- row + 1
                    if col + 1 > colCount then colCount <- col + 1
                    let formula =
                        if isNullOrUndefined formulaObj then None else Some (string formulaObj)
                    let valueCell = excelCellFromObj rawValue null formula None None
                    let cellValue = convertCellValue valueCell
                    let baseFormat = styleFormat styleIndex
                    let cellFormat =
                        { baseFormat with
                            HasComment = baseFormat.HasComment || hasComment
                            HasListDropdown = baseFormat.HasListDropdown || hasListDropdown }
                    let comment =
                        if hasComment && isFunction excel?commentAt then
                            let excelRef = sprintf "%s%d" (columnToLetter col) (row + 1)
                            let commentObj = excel?commentAt(excelRef, sheet)
                            if isNullOrUndefined commentObj then None
                            else
                                let text = optStringProp commentObj "text" |> Option.defaultValue ""
                                let author = optStringProp commentObj "author"
                                if System.String.IsNullOrWhiteSpace(text) then None
                                else Some { Text = text; Author = author }
                        else
                            None
                    let cellFormat =
                        if comment.IsSome then { cellFormat with HasComment = true }
                        else cellFormat
                    if cellValue <> CellValue.Empty || cellFormat <> defaultFormat || comment.IsSome then
                        noteImportedValue cellValue
                        let addr = { Row = row; Col = col }
                        cellMap <- Map.add addr { Value = cellValue; Format = cellFormat; Comment = comment } cellMap

                        if not (isNullOrUndefined richRunObjs) then
                            let richRunArray = arrayFrom richRunObjs
                            let runs = ResizeArray<string * RunOverride>(richRunArray.Length)
                            for runObj in richRunArray do
                                runs.Add(convertRichRun runObj)
                            if runs.Count > 0 then
                                richRunsMap <- Map.add addr (resizeArrayToList runs) richRunsMap

            if isFunction excel?cellsColumnar && isFunction excel?styleAt then
                setDimensionsFromSheet ()
                let columnarCells = excel?cellsColumnar(workbook, sheet)
                let rows: int[] = unbox (getProp columnarCells "rows")
                let cols: int[] = unbox (getProp columnarCells "cols")
                let values: obj[] = unbox (getProp columnarCells "values")
                let formulas: obj[] = unbox (getProp columnarCells "formulas")
                let styleIndexes: int[] = unbox (getProp columnarCells "styleIndexes")
                let hasComments: bool[] = unbox (getProp columnarCells "hasComments")
                let hasListDropdowns: bool[] = unbox (getProp columnarCells "hasListDropdowns")
                let richRuns: obj[] = unbox (getProp columnarCells "richRuns")
                let count =
                    [| rows.Length; cols.Length; values.Length; formulas.Length; styleIndexes.Length
                       hasComments.Length; hasListDropdowns.Length; richRuns.Length |]
                    |> Array.min
                for index in 0 .. count - 1 do
                    importCompactCell
                        rows.[index]
                        cols.[index]
                        values.[index]
                        formulas.[index]
                        styleIndexes.[index]
                        hasComments.[index]
                        hasListDropdowns.[index]
                        richRuns.[index]
            elif isFunction excel?cellsCompact && isFunction excel?styleAt then
                setDimensionsFromSheet ()
                let compactCells = excel?cellsCompact(workbook, sheet) |> arrayFrom
                for compactCell in compactCells do
                    let row = getIndex compactCell 0 |> unbox<int>
                    let col = getIndex compactCell 1 |> unbox<int>
                    importCompactCell
                        row
                        col
                        (getIndex compactCell 2)
                        (getIndex compactCell 3)
                        (getIndex compactCell 4 |> unbox<int>)
                        (getIndex compactCell 5 |> unbox<bool>)
                        (getIndex compactCell 6 |> unbox<bool>)
                        (getIndex compactCell 7)
            elif isFunction excel?cellsWithStyle then
                setDimensionsFromSheet ()
                let cellObjs = excel?cellsWithStyle(workbook, sheet) |> arrayFrom
                for cellObj in cellObjs do
                    let excelRef = optStringProp cellObj "ref" |> Option.defaultValue ""
                    match parseCellRef excelRef with
                    | Some (row, col) ->
                        if row + 1 > rowCount then rowCount <- row + 1
                        if col + 1 > colCount then colCount <- col + 1
                        let hasComment = optBoolProp cellObj "hasComment" |> Option.defaultValue false
                        importCell row col excelRef (getProp cellObj "value") cellObj (Some (getProp cellObj "richRuns")) (Some hasComment)
                    | None -> ()
            else
                // Older office-format bundles do not expose streaming cell
                // APIs. Import raw values from toArray without calling
                // cellWithStyleAt per coordinate; that lookup scans worksheet
                // rows and makes this legacy path quadratic.
                let dataArray: obj[][] = excel?toArray(workbook, sheet)
                rowCount <- dataArray.Length
                colCount <-
                    if rowCount > 0 then
                        dataArray |> Array.map (fun row -> row.Length) |> Array.max
                    else
                        0
                for row in 0 .. rowCount - 1 do
                    for col in 0 .. colCount - 1 do
                        let colLetter = columnToLetter col
                        let excelRef = sprintf "%s%d" colLetter (row + 1)
                        let rawValue =
                            if row < dataArray.Length && col < dataArray.[row].Length then dataArray.[row].[col]
                            else null
                        importCell row col excelRef rawValue null None None

            let addColumnWidth col width =
                if width > 0.0 then
                    // Convert Excel width units to pixels (approximate: 1 Excel unit ≈ 7 pixels)
                    let pixelWidth = width * 7.0
                    columnWidths <- Map.add col pixelWidth columnWidths

            let addRowHeight row height =
                if height > 0.0 then
                    // Convert points to pixels (1 point = 1.333 pixels at 96 DPI)
                    let pixelHeight = height * 1.333
                    rowHeights <- Map.add row pixelHeight rowHeights

            // Import column widths (office-format provides in Excel units)
            if isFunction excel?columnWidths then
                let widths = excel?columnWidths(colCount, sheet) |> arrayFrom
                let last = min colCount widths.Length
                if last > 0 then
                    for col in 0 .. last - 1 do
                        addColumnWidth col (unbox<float> widths.[col])
            else
                for col in 0 .. colCount - 1 do
                    let width: float =
                        if isFunction excel?columnWidth then excel?columnWidth(col, sheet)
                        else 0.0
                    addColumnWidth col width

            // Import row heights (office-format provides in points)
            if isFunction excel?rowHeights then
                let heights = excel?rowHeights(rowCount, sheet) |> arrayFrom
                let last = min rowCount heights.Length
                if last > 0 then
                    for row in 0 .. last - 1 do
                        addRowHeight row (unbox<float> heights.[row])
            else
                for row in 0 .. rowCount - 1 do
                    let height: float =
                        if isFunction excel?rowHeight then excel?rowHeight(row, sheet)
                        else 0.0
                    addRowHeight row height

            // Import merged cells (office-format provides as array of ref strings like "A1:B2")
            let mergedCells: string[] = excel?mergedCells(sheet)
            let mergedCellList =
                if isNull mergedCells then []
                else
                    let ranges = ResizeArray<MergedCell>()
                    for refRange in mergedCells do
                        try
                            match parseMergedCellRange refRange with
                            | Some range -> ranges.Add range
                            | None -> ()
                        with _ -> ()
                    resizeArrayToList ranges

            // Create model with imported data
            let importedModel = {
                Cells = cellMap
                RowCount = max rowCount 100  // Minimum 100 rows
                ColCount = max colCount 26   // Minimum 26 columns (A-Z)
                Selection = SingleCell (0, 0)
                EditingCell = None
                EditBuffer = ""
                Clipboard = None
                History = { Past = []; Future = []; MaxSize = 50 }
                SortState = None
                Filters = Map.empty
                FilterVisibleRows = None
                FilterHiddenRows = [||]
                RowHeights = rowHeights
                ColumnWidths = columnWidths
                SearchQuery = None
                SearchCaseSensitive = None
                SearchResults = [||]
                SearchIndex = 0
                ConditionalFormats = Map.empty
                RichRuns = richRunsMap
                MergedCells = mergedCellList
                FrozenRows = 0
                FrozenCols = 0
                PageSetup = defaultPageSetup
                Tables = []
                CfBoldCells = Set.empty
                ShowGridLines = true
                Spill = Map.empty
                Validations = Map.empty
                Macros = ""
            }
            if formulaCount = 0 then
                importedModel
            else
                recalculateAllFormulas None None importedModel

        with ex ->
            printfn "Error importing XLSX: %s" ex.Message
            // Return empty model on error
            {
                Cells = Map.empty
                RowCount = 100
                ColCount = 26
                Selection = SingleCell (0, 0)
                EditingCell = None
                EditBuffer = ""
                Clipboard = None
                History = { Past = []; Future = []; MaxSize = 50 }
                SortState = None
                Filters = Map.empty
                FilterVisibleRows = None
                FilterHiddenRows = [||]
                RowHeights = Map.empty
                ColumnWidths = Map.empty
                SearchQuery = None
                SearchCaseSensitive = None
                SearchResults = [||]
                SearchIndex = 0
                ConditionalFormats = Map.empty
                RichRuns = Map.empty
                MergedCells = []
                FrozenRows = 0
                FrozenCols = 0
                PageSetup = defaultPageSetup
                Tables = []
                CfBoldCells = Set.empty
                ShowGridLines = true
                Spill = Map.empty
                Validations = Map.empty
                Macros = ""
            }

    /// <summary>
    /// Import XLSX file from ArrayBuffer.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <param name="sheetName">Optional sheet name (or first sheet if not specified)</param>
    /// <returns>SpreadsheetModel populated with sheet data</returns>
let importFromBuffer (buffer: JS.ArrayBuffer) (sheetName: string option): Async<SpreadsheetModel> =
        async {
            try
                let api = officeFormatApi?``default``()
                let excel = api?excel

                // Load only the requested sheet when the office-format bundle
                // supports it. Full workbook loading still exists for callers
                // that need every sheet, but single-sheet import should not
                // parse unrelated large worksheets.
                let loadPromise =
                    if isFunction excel?loadSelected then excel?loadSelected(buffer, sheetName)
                    else excel?load(buffer)
                let! workbook = loadPromise |> Async.AwaitPromise

                // Import sheet
                return importSheet workbook sheetName

            with ex ->
                printfn "Error loading XLSX from buffer: %s" ex.Message
                return {
                    Cells = Map.empty
                    RowCount = 100
                    ColCount = 26
                    Selection = SingleCell (0, 0)
                    EditingCell = None
                    EditBuffer = ""
                    Clipboard = None
                    History = { Past = []; Future = []; MaxSize = 50 }
                    SortState = None
                    Filters = Map.empty
                    FilterVisibleRows = None
                    FilterHiddenRows = [||]
                    RowHeights = Map.empty
                    ColumnWidths = Map.empty
                    SearchQuery = None
                    SearchCaseSensitive = None
                    SearchResults = [||]
                    SearchIndex = 0
                    ConditionalFormats = Map.empty
                    RichRuns = Map.empty
                    MergedCells = []
                    FrozenRows = 0
                    FrozenCols = 0
                    PageSetup = defaultPageSetup
                    Tables = []
                    CfBoldCells = Set.empty
                    ShowGridLines = true
                    Spill = Map.empty
                    Validations = Map.empty
                    Macros = ""
                }
        }


    /// <summary>
    /// Import every sheet in the workbook in source order, paired with
    /// its sheet name. Used by the multi-sheet load path so the chrome
    /// can render a tab strip and the user can swap between sheets
    /// without re-parsing the XLSX each time.
    /// </summary>
let importAllSheets (buffer: JS.ArrayBuffer): Async<(string * SpreadsheetModel) array> =
        async {
            try
                let api = officeFormatApi?``default``()
                let excel = api?excel
                let! workbook = excel?load(buffer) |> Async.AwaitPromise
                let names: string[] = excel?sheetNames(workbook)
                return
                    names
                    |> Array.map (fun n -> n, importSheet workbook (Some n))
            with ex ->
                printfn "Error loading workbook: %s" ex.Message
                return [||]
        }

    /// <summary>
    /// Get list of sheet names from workbook buffer.
    /// Useful for showing sheet selector before import.
    /// </summary>
    /// <param name="buffer">ArrayBuffer containing XLSX file data</param>
    /// <returns>Array of sheet names</returns>
let getSheetNames (buffer: JS.ArrayBuffer): Async<string[]> =
        async {
            try
                let api = officeFormatApi?``default``()
                let excel = api?excel

                let! workbook = excel?load(buffer) |> Async.AwaitPromise
                let sheetNames: string[] = excel?sheetNames(workbook)
                return sheetNames

            with ex ->
                printfn "Error getting sheet names: %s" ex.Message
                return [||]
        }
