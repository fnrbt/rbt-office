module XlsxWriter

open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop
open System.Text

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

let private appendEscaped (sb: StringBuilder) (text: string) =
    for ch in text do
        match ch with
        | '&' -> sb.Append("&amp;") |> ignore
        | '<' -> sb.Append("&lt;") |> ignore
        | '>' -> sb.Append("&gt;") |> ignore
        | '"' -> sb.Append("&quot;") |> ignore
        | '\'' -> sb.Append("&apos;") |> ignore
        | _ -> sb.Append(ch) |> ignore

let private appendTextElement (sb: StringBuilder) (name: string) (text: string) =
    sb.Append('<').Append(name) |> ignore
    if text = "" then
        sb.Append("/>") |> ignore
    else
        sb.Append('>') |> ignore
        appendEscaped sb text
        sb.Append("</").Append(name).Append('>') |> ignore

let private appendCellRef (sb: StringBuilder) (colLetters: string) (rowNumber: string) =
    sb.Append(colLetters).Append(rowNumber) |> ignore

let private appendSimpleCell (sb: StringBuilder) (rowNumber: string) (colLetters: string) (value: string) =
    sb.Append("<c r=\"") |> ignore
    appendCellRef sb colLetters rowNumber
    match tryParseNumber value with
    | Some num ->
        sb.Append("\"><v>").Append(string num).Append("</v></c>") |> ignore
    | None ->
        sb.Append("\" t=\"inlineStr\"><is>") |> ignore
        appendTextElement sb "t" value
        sb.Append("</is></c>") |> ignore

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
    let columnLetters =
        let maxCols =
            data
            |> List.fold (fun maxCols row -> max maxCols row.Length) 0
        [| for col in 0 .. maxCols - 1 -> columnToLetter col |]

    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<worksheet xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\">") |> ignore
    sb.Append("<sheetData>") |> ignore
    let mutable rowIdx = 0
    for row in data do
        let rowNumber = string (rowIdx + 1)
        sb.Append("<row r=\"").Append(rowNumber).Append("\">") |> ignore
        let mutable colIdx = 0
        for value in row do
            appendSimpleCell sb rowNumber columnLetters.[colIdx] value
            colIdx <- colIdx + 1
        sb.Append("</row>") |> ignore
        rowIdx <- rowIdx + 1
    sb.Append("</sheetData></worksheet>") |> ignore
    sb.ToString()

/// Generate workbook.xml
let generateWorkbookXml (sheetNames: string list) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<workbook xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\"><sheets>") |> ignore
    let mutable index = 0
    for name in sheetNames do
        let sheetId = index + 1
        sb.Append("<sheet name=\"") |> ignore
        appendEscaped sb name
        sb.Append("\" sheetId=\"").Append(sheetId)
            .Append("\" r:id=\"rId").Append(sheetId).Append("\"/>") |> ignore
        index <- index + 1
    sb.Append("</sheets></workbook>") |> ignore
    sb.ToString()

/// Generate workbook relationships
let generateWorkbookRels (sheetCount: int) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<Relationships xmlns=\"").Append(Namespaces.relationshipsPackage).Append("\">") |> ignore
    for i in 1 .. sheetCount do
        sb.Append("<Relationship Id=\"rId").Append(i)
            .Append("\" Type=\"").Append(RelationshipTypes.worksheet)
            .Append("\" Target=\"worksheets/sheet").Append(i).Append(".xml\"/>") |> ignore
    sb.Append("</Relationships>") |> ignore
    sb.ToString()

/// Generate [Content_Types].xml for XLSX
let generateContentTypesXlsx (sheetCount: int) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<Types xmlns=\"").Append(Namespaces.contentTypes).Append("\">") |> ignore
    sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>") |> ignore
    sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>") |> ignore
    sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>") |> ignore
    for i in 1 .. sheetCount do
        sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i)
            .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>") |> ignore
    sb.Append("</Types>") |> ignore
    sb.ToString()

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

let private appendFormattedCell (sb: StringBuilder) (rowNumber: string) (colLetters: string) (cell: FormattedCell) =
    sb.Append("<c r=\"") |> ignore
    appendCellRef sb colLetters rowNumber
    match cell.Formula with
    | Some formula ->
        sb.Append("\">") |> ignore
        appendTextElement sb "f" formula
        appendTextElement sb "v" cell.Value
        sb.Append("</c>") |> ignore
    | None ->
        match tryParseNumber cell.Value with
        | Some num ->
            sb.Append("\"><v>").Append(string num).Append("</v></c>") |> ignore
        | None ->
            sb.Append("\" t=\"inlineStr\"><is>") |> ignore
            appendTextElement sb "t" cell.Value
            sb.Append("</is></c>") |> ignore

let private appendMergeCells (sb: StringBuilder) (mergedCells: MergedCellRange list) =
    if not (List.isEmpty mergedCells) then
        sb.Append("<mergeCells count=\"").Append(List.length mergedCells).Append("\">") |> ignore
        for range in mergedCells do
            sb.Append("<mergeCell ref=\"")
                .Append(cellRef range.StartRow range.StartCol)
                .Append(':')
                .Append(cellRef range.EndRow range.EndCol)
                .Append("\"/>") |> ignore
        sb.Append("</mergeCells>") |> ignore

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
    let columnLetters =
        let maxCols =
            sheet.Rows
            |> List.fold (fun maxCols row -> max maxCols row.Cells.Length) 0
        [| for col in 0 .. maxCols - 1 -> columnToLetter col |]

    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<worksheet xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\">") |> ignore
    sb.Append("<sheetData>") |> ignore
    let mutable rowIdx = 0
    for row in sheet.Rows do
        let rowNumber = string (rowIdx + 1)
        sb.Append("<row r=\"").Append(rowNumber).Append("\">") |> ignore
        let mutable colIdx = 0
        for cell in row.Cells do
            appendFormattedCell sb rowNumber columnLetters.[colIdx] cell
            colIdx <- colIdx + 1
        sb.Append("</row>") |> ignore
        rowIdx <- rowIdx + 1
    sb.Append("</sheetData>") |> ignore
    appendMergeCells sb sheet.MergedCells
    sb.Append("</worksheet>") |> ignore
    sb.ToString()

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
