module XlsxWriter

open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop

/// Convert column index to Excel letter (0 -> A, 25 -> Z, 26 -> AA)
let columnToLetter (col: int) : string =
    let rec loop n acc =
        if n < 0 then acc
        else loop ((n / 26) - 1) (char (65 + (n % 26)) :: acc)
    loop col [] |> Array.ofList |> System.String

/// Convert row/col to cell reference (0, 0 -> A1)
let cellRef (row: int) (col: int) : string =
    sprintf "%s%d" (columnToLetter col) (row + 1)

/// Try to parse a value as a number
let tryParseNumber (value: string) : float option =
    match System.Double.TryParse(value) with
    | (true, num) -> Some num
    | _ -> None

/// Generate cell element
let generateCell (rowIdx: int) (colIdx: int) (value: string) : string =
    let ref = cellRef rowIdx colIdx

    match tryParseNumber value with
    | Some num ->
        // Numeric cell
        element "c" [("r", ref)] [
            elementWithText "v" [] (string num)
        ]
    | None ->
        // Inline string cell
        element "c" [("r", ref); ("t", "inlineStr")] [
            element "is" [] [
                elementWithText "t" [] value
            ]
        ]

/// Generate sheet XML from 2D array
let generateSheetXml (data: string list list) : string =
    let rowElements =
        data
        |> List.mapi (fun rowIdx cells ->
            let cellElements =
                cells
                |> List.mapi (fun colIdx value ->
                    generateCell rowIdx colIdx value
                )
                |> String.concat ""

            element "row" [("r", string (rowIdx + 1))] [cellElements]
        )
        |> String.concat ""

    let sheetData = element "sheetData" [] [rowElements]

    let root =
        element "worksheet"
            [("xmlns", Namespaces.spreadsheet); ("xmlns:r", Namespaces.relationships)]
            [sheetData]

    document root

/// Generate workbook.xml
let generateWorkbookXml (sheetNames: string list) : string =
    let sheets =
        sheetNames
        |> List.mapi (fun i name ->
            selfClosing "sheet" [("name", name); ("sheetId", string (i + 1)); ("r:id", sprintf "rId%d" (i + 1))]
        )
        |> String.concat ""

    let sheetsElement = element "sheets" [] [sheets]
    let root =
        element "workbook"
            [("xmlns", Namespaces.spreadsheet); ("xmlns:r", Namespaces.relationships)]
            [sheetsElement]

    document root

/// Generate workbook relationships
let generateWorkbookRels (sheetCount: int) : string =
    let rels =
        [1..sheetCount]
        |> List.map (fun i ->
            selfClosing "Relationship" [("Id", sprintf "rId%d" i); ("Type", RelationshipTypes.worksheet); ("Target", sprintf "worksheets/sheet%d.xml" i)]
        )
        |> String.concat ""

    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rels]
    document root

/// Generate [Content_Types].xml for XLSX
let generateContentTypesXlsx (sheetCount: int) : string =
    let worksheetOverrides =
        [1..sheetCount]
        |> List.map (fun i ->
            selfClosing "Override" [("PartName", sprintf "/xl/worksheets/sheet%d.xml" i); ("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")]
        )
        |> String.concat ""

    let children = [
        selfClosing "Default" [("Extension", "rels"); ("ContentType", "application/vnd.openxmlformats-package.relationships+xml")]
        selfClosing "Default" [("Extension", "xml"); ("ContentType", "application/xml")]
        selfClosing "Override" [("PartName", "/xl/workbook.xml"); ("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")]
        worksheetOverrides
    ]

    let root = element "Types" [("xmlns", Namespaces.contentTypes)] children
    document root

/// Generate root relationships
let generateRootRelsXlsx () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.officeDocument); ("Target", "xl/workbook.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Type for sheet data
type SheetData = {
    Name: string
    Data: string list list
}

/// Type for formatted cell
type FormattedCell = {
    Value: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    FontSize: float option
    FontFamily: string option
    FontColor: string option
    FillColor: string option
    NumberFormat: string option
    Formula: string option  // Excel formula (without =)
}

/// Type for formatted row
type FormattedRow = {
    Cells: FormattedCell list
}

/// Type for merged cell range
type MergedCellRange = {
    StartRow: int  // 0-indexed
    StartCol: int  // 0-indexed
    EndRow: int    // 0-indexed
    EndCol: int    // 0-indexed
}

/// Type for formatted sheet
type FormattedSheet = {
    Name: string
    Rows: FormattedRow list
    MergedCells: MergedCellRange list
}

/// Helper to create a simple cell from string
let simpleCell (value: string) : FormattedCell = {
    Value = value
    Bold = None
    Italic = None
    Underline = None
    FontSize = None
    FontFamily = None
    FontColor = None
    FillColor = None
    NumberFormat = None
    Formula = None
}

/// Convert hex color to RGB (removes # prefix)
let hexToRgb (color: string) : string =
    if color.StartsWith("#") then
        "FF" + color.Substring(1).ToUpper()  // Add alpha channel
    else
        "FF" + color.ToUpper()

/// Generate cell style XML (inline for Phase 2, will be deduplicated in Phase 3)
let generateCellStyle (cell: FormattedCell) : int option * string =
    let hasStyle =
        cell.Bold.IsSome || cell.Italic.IsSome || cell.Underline.IsSome ||
        cell.FontSize.IsSome || cell.FontFamily.IsSome ||
        cell.FontColor.IsSome || cell.FillColor.IsSome

    if not hasStyle then
        (None, "")
    else
        // For Phase 2, we'll use inline styling (simplified)
        // In Phase 3, we'll implement proper style deduplication
        (None, "")

/// Generate formatted cell element
let generateFormattedCell (rowIdx: int) (colIdx: int) (cell: FormattedCell) : string =
    let ref = cellRef rowIdx colIdx

    match cell.Formula with
    | Some formula ->
        // Cell with formula
        element "c" [("r", ref)] [
            elementWithText "f" [] formula
            elementWithText "v" [] cell.Value  // Cached value
        ]
    | None ->
        match tryParseNumber cell.Value with
        | Some num ->
            // Numeric cell
            element "c" [("r", ref)] [
                elementWithText "v" [] (string num)
            ]
        | None ->
            // Inline string cell
            element "c" [("r", ref); ("t", "inlineStr")] [
                element "is" [] [
                    elementWithText "t" [] cell.Value
                ]
            ]

/// Generate merged cells XML
let generateMergedCells (mergedCells: MergedCellRange list) : string =
    if List.isEmpty mergedCells then
        ""
    else
        let mergeElements =
            mergedCells
            |> List.map (fun range ->
                let startRef = cellRef range.StartRow range.StartCol
                let endRef = cellRef range.EndRow range.EndCol
                let rangeRef = sprintf "%s:%s" startRef endRef
                selfClosing "mergeCell" [("ref", rangeRef)]
            )
            |> String.concat ""

        element "mergeCells" [("count", string (List.length mergedCells))] [mergeElements]

/// Generate formatted sheet XML
let generateFormattedSheetXml (sheet: FormattedSheet) : string =
    let rowElements =
        sheet.Rows
        |> List.mapi (fun rowIdx row ->
            let cellElements =
                row.Cells
                |> List.mapi (fun colIdx cell ->
                    generateFormattedCell rowIdx colIdx cell
                )
                |> String.concat ""

            element "row" [("r", string (rowIdx + 1))] [cellElements]
        )
        |> String.concat ""

    let sheetData = element "sheetData" [] [rowElements]
    let mergedCellsXml = generateMergedCells sheet.MergedCells

    let children = [sheetData; mergedCellsXml] |> List.filter (fun s -> s <> "")

    let root =
        element "worksheet"
            [("xmlns", Namespaces.spreadsheet); ("xmlns:r", Namespaces.relationships)]
            children

    document root

/// Helper to create a cell with formula
let cellWithFormula (formula: string) (cachedValue: string) : FormattedCell = {
    Value = cachedValue
    Bold = None
    Italic = None
    Underline = None
    FontSize = None
    FontFamily = None
    FontColor = None
    FillColor = None
    NumberFormat = None
    Formula = Some formula
}

/// Helper to create a merged cell range
let createMergedRange (startRow: int) (startCol: int) (endRow: int) (endCol: int) : MergedCellRange = {
    StartRow = startRow
    StartCol = startCol
    EndRow = endRow
    EndCol = endCol
}

/// Helper to create a merged range from Excel-style references (e.g., "A1:B2")
let parseMergedRange (range: string) : MergedCellRange option =
    let parts = range.Split(':')
    if parts.Length <> 2 then
        None
    else
        // This is a simplified parser - in production you'd want full Excel reference parsing
        // For now, just return None - users should use createMergedRange directly
        None

/// Create XLSX from formatted sheets
let writeFormattedXlsx (sheets: FormattedSheet list) : Fable.Core.JS.Promise<obj> =
    let sheetNames = sheets |> List.map (fun s -> s.Name)

    let baseFiles = [
        { Path = "[Content_Types].xml"; Content = generateContentTypesXlsx (List.length sheets) }
        { Path = "_rels/.rels"; Content = generateRootRelsXlsx () }
        { Path = "xl/workbook.xml"; Content = generateWorkbookXml sheetNames }
        { Path = "xl/_rels/workbook.xml.rels"; Content = generateWorkbookRels (List.length sheets) }
    ]

    // Add each sheet
    let sheetFiles =
        sheets
        |> List.mapi (fun i sheet ->
            { Path = sprintf "xl/worksheets/sheet%d.xml" (i + 1); Content = generateFormattedSheetXml sheet }
        )

    let files = baseFiles @ sheetFiles

    createPackage files

/// Create minimal XLSX from 2D arrays
let writeMinimalXlsx (sheets: SheetData list) : Fable.Core.JS.Promise<obj> =
    let sheetNames = sheets |> List.map (fun s -> s.Name)

    let baseFiles = [
        { Path = "[Content_Types].xml"; Content = generateContentTypesXlsx (List.length sheets) }
        { Path = "_rels/.rels"; Content = generateRootRelsXlsx () }
        { Path = "xl/workbook.xml"; Content = generateWorkbookXml sheetNames }
        { Path = "xl/_rels/workbook.xml.rels"; Content = generateWorkbookRels (List.length sheets) }
    ]

    // Add each sheet
    let sheetFiles =
        sheets
        |> List.mapi (fun i sheet ->
            { Path = sprintf "xl/worksheets/sheet%d.xml" (i + 1); Content = generateSheetXml sheet.Data }
        )

    let files = baseFiles @ sheetFiles

    createPackage files
