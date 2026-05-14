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
                Some { Style = style; Color = b.color }
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
            UnderlineDouble = false
            FontName = None
            TextColor = excelCell.fontColor
            BgColor = excelCell.fillColor
            FontSize = excelCell.fontSize |> Option.map int  // Convert float to int
            HorizontalAlign = excelCell.horizontalAlign
            VerticalAlign = excelCell.verticalAlign
            WrapText = excelCell.wrapText |> Option.defaultValue false
            NumberFormat = excelCell.numberFormat
            BorderTop = convertBorder excelCell.borderTop
            BorderBottom = convertBorder excelCell.borderBottom
            BorderLeft = convertBorder excelCell.borderLeft
            BorderRight = convertBorder excelCell.borderRight
            BorderDiagonal = None
            DiagonalUp = false
            DiagonalDown = false
            FilterButton = false
            Strikethrough = false
            HasComment = false
            HasListDropdown = false
            DataBar = None
            IconSet = None
        }

    /// <summary>
    /// Convert Excel column letter (A, B, ..., Z, AA, AB, ...) to zero-based index.
    /// </summary>
let private columnLetterToIndex (letter: string): int =
        let mutable result = 0
        for c in letter.ToUpper() do
            result <- result * 26 + (int c - int 'A' + 1)
        result - 1

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

            // Get basic data array to determine dimensions
            let dataArray: obj[][] = excel?toArray(workbook, sheet)
            let rowCount = dataArray.Length
            let colCount =
                if rowCount > 0 then
                    dataArray |> Array.map (fun row -> row.Length) |> Array.max
                else
                    0

            // Build cell map with styles
            let mutable cellMap = Map.empty<CellAddress, CellData>
            let mutable columnWidths = Map.empty<int, float>
            let mutable rowHeights = Map.empty<int, float>

            // Import cells with styles
            for row in 0 .. rowCount - 1 do
                for col in 0 .. colCount - 1 do
                    // Convert to Excel reference (A1, B1, etc.)
                    let colLetter = columnToLetter col
                    let excelRef = sprintf "%s%d" colLetter (row + 1)

                    let rawValue =
                        if row < dataArray.Length && col < dataArray.[row].Length then dataArray.[row].[col]
                        else null

                    // Get cell with style
                    let cellObj = excel?cellWithStyleAt(excelRef, sheet)
                    if not (isNull cellObj) || not (isNull rawValue) then
                        // Helper to extract border from cell object
                        let extractBorder (borderObj: obj) =
                            if isNull borderObj then None
                            else
                                Some {
                                    style = if borderObj?style = null then None else Some (string borderObj?style)
                                    color = if borderObj?color = null then None else Some (string borderObj?color)
                                }

                        let excelCell: ExcelCell =
                            {
                                value = if isNull rawValue || string rawValue = "" then None else Some rawValue
                                formula = if isNull cellObj || cellObj?formula = null then None else Some (string cellObj?formula)
                                bold = if isNull cellObj || cellObj?bold = null then None else Some (unbox cellObj?bold)
                                italic = if isNull cellObj || cellObj?italic = null then None else Some (unbox cellObj?italic)
                                underline = if isNull cellObj || cellObj?underline = null then None else Some (unbox cellObj?underline)
                                fontColor = if isNull cellObj || cellObj?fontColor = null then None else Some (string cellObj?fontColor)
                                fillColor = if isNull cellObj || cellObj?fillColor = null then None else Some (string cellObj?fillColor)
                                fontSize = if isNull cellObj || cellObj?fontSize = null then None else Some (unbox cellObj?fontSize)
                                numberFormat = if isNull cellObj || cellObj?numberFormat = null then None else Some (string cellObj?numberFormat)
                                horizontalAlign = if isNull cellObj || cellObj?horizontalAlign = null then None else Some (string cellObj?horizontalAlign)
                                verticalAlign = if isNull cellObj || cellObj?verticalAlign = null then None else Some (string cellObj?verticalAlign)
                                wrapText = if isNull cellObj || cellObj?wrapText = null then None else Some (unbox cellObj?wrapText)
                                borderTop = if isNull cellObj then None else extractBorder cellObj?borderTop
                                borderBottom = if isNull cellObj then None else extractBorder cellObj?borderBottom
                                borderLeft = if isNull cellObj then None else extractBorder cellObj?borderLeft
                                borderRight = if isNull cellObj then None else extractBorder cellObj?borderRight
                            }

                        let cellValue = convertCellValue excelCell
                        let cellFormat = convertCellFormat excelCell

                        // Read comment if present. Older office-format builds expose comments as a collection only.
                        let commentObj =
                            if isFunction excel?commentAt then excel?commentAt(excelRef, sheet)
                            else null
                        let comment =
                            if isNull commentObj then None
                            else
                                let text = if commentObj?text = null then "" else string commentObj?text
                                let author = if commentObj?author = null then None else Some (string commentObj?author)
                                if System.String.IsNullOrWhiteSpace(text) then None
                                else Some { Text = text; Author = author }

                        // Only store non-empty cells
                        if cellValue <> CellValue.Empty then
                            let addr = { Row = row; Col = col }
                            cellMap <- Map.add addr { Value = cellValue; Format = cellFormat; Comment = comment } cellMap

            // Import column widths (office-format provides in Excel units)
            for col in 0 .. colCount - 1 do
                let width: float =
                    if isFunction excel?columnWidth then excel?columnWidth(col, sheet)
                    else 0.0
                if width > 0.0 then
                    // Convert Excel width units to pixels (approximate: 1 Excel unit ≈ 7 pixels)
                    let pixelWidth = width * 7.0
                    columnWidths <- Map.add col pixelWidth columnWidths

            // Import row heights (office-format provides in points)
            for row in 0 .. rowCount - 1 do
                let height: float =
                    if isFunction excel?rowHeight then excel?rowHeight(row, sheet)
                    else 0.0
                if height > 0.0 then
                    // Convert points to pixels (1 point = 1.333 pixels at 96 DPI)
                    let pixelHeight = height * 1.333
                    rowHeights <- Map.add row pixelHeight rowHeights

            // Import merged cells (office-format provides as array of ref strings like "A1:B2")
            let mergedCells: string[] = excel?mergedCells(sheet)
            let mergedCellList =
                if isNull mergedCells then []
                else
                    mergedCells
                    |> Array.choose (fun refRange ->
                        try
                            // Parse range like "A1:B2"
                            let parts = refRange.Split(':')
                            if parts.Length = 2 then
                                match Model.parseCellRef parts.[0], Model.parseCellRef parts.[1] with
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
                        with _ -> None
                    )
                    |> Array.toList

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
                RowHeights = rowHeights
                ColumnWidths = columnWidths
                SearchQuery = None
                SearchResults = []
                SearchIndex = 0
                ConditionalFormats = Map.empty
                RichRuns = Map.empty
                MergedCells = mergedCellList
                FrozenRows = 0
                FrozenCols = 0
            }
            recalculateAllFormulas None importedModel

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
                RowHeights = Map.empty
                ColumnWidths = Map.empty
                SearchQuery = None
                SearchResults = []
                SearchIndex = 0
                ConditionalFormats = Map.empty
                RichRuns = Map.empty
                MergedCells = []
                FrozenRows = 0
                FrozenCols = 0
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

                // Load workbook from buffer
                let! workbook = excel?load(buffer) |> Async.AwaitPromise

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
                    RowHeights = Map.empty
                    ColumnWidths = Map.empty
                    SearchQuery = None
                    SearchResults = []
                    SearchIndex = 0
                    ConditionalFormats = Map.empty
                    RichRuns = Map.empty
                    MergedCells = []
                    FrozenRows = 0
                    FrozenCols = 0
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
