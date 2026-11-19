module WorkbookBuilder

open Fable.Core.JsInterop
open XlsxWriter
open XlsxPackageWriter

/// Sheet builder for building a single sheet
type SheetBuilder(name: string) =
    let mutable rows: FormattedRow list = []
    let mutable mergedCells: MergedCellRange list = []

    /// Add a row with formatted cells
    member this.AddRow(cells: FormattedCell list) =
        let row = { Cells = cells }
        rows <- row :: rows
        this

    /// Add a simple row from strings
    member this.AddSimpleRow(values: string list) =
        let cells = values |> List.map simpleCell
        this.AddRow(cells)

    /// Add a merged cell range
    member this.AddMergedRange(startRow: int, startCol: int, endRow: int, endCol: int) =
        let range = createMergedRange startRow startCol endRow endCol
        mergedCells <- range :: mergedCells
        this

    /// Get the built sheet
    member this.Build() : FormattedSheet =
        { Name = name; Rows = List.rev rows; MergedCells = List.rev mergedCells }

/// Workbook builder
type WorkbookBuilder() =
    let mutable sheets: FormattedSheet list = []

    /// Add a sheet to the workbook
    member this.AddSheet(name: string, buildAction: SheetBuilder -> unit) =
        let sheetBuilder = SheetBuilder(name)
        buildAction sheetBuilder
        let sheet = sheetBuilder.Build()
        sheets <- sheet :: sheets
        this

    /// Add a simple sheet from 2D array
    member this.AddSimpleSheet(name: string, data: string list list) =
        let rows =
            data
            |> List.map (fun rowData ->
                { Cells = rowData |> List.map simpleCell }
            )
        let sheet = { Name = name; Rows = rows; MergedCells = [] }
        sheets <- sheet :: sheets
        this

    /// Save the workbook to a blob
    member this.Save() =
        writeFormattedXlsx (List.rev sheets)

/// Create a new workbook builder
let create () = WorkbookBuilder()

/// Add sheet (for JavaScript interop)
let addSheet (builder: WorkbookBuilder) (name: string) (buildAction: SheetBuilder -> unit) =
    builder.AddSheet(name, buildAction)

/// Add simple sheet (for JavaScript interop)
let addSimpleSheet (builder: WorkbookBuilder) (name: string) (data: string list list) =
    builder.AddSimpleSheet(name, data)

/// Save workbook (for JavaScript interop)
let saveWorkbook (builder: WorkbookBuilder) =
    builder.Save()

/// Add row to sheet builder (for JavaScript interop)
let addRow (sheetBuilder: SheetBuilder) (cells: FormattedCell list) =
    sheetBuilder.AddRow(cells)

/// Add simple row to sheet builder (for JavaScript interop)
let addSimpleRow (sheetBuilder: SheetBuilder) (values: string list) =
    sheetBuilder.AddSimpleRow(values)

/// Helper to create a formatted cell
let createCell (value: string) (bold: bool option) (italic: bool option) (fontSize: float option) (fontColor: string option) (fillColor: string option) : FormattedCell =
    {
        Value = value
        Bold = bold
        Italic = italic
        Underline = None
        FontSize = fontSize
        FontFamily = Some "Calibri"
        FontColor = fontColor
        FillColor = fillColor
        NumberFormat = None
        Formula = None
    }

/// Helper to create a header cell
let createHeaderCell (value: string) : FormattedCell =
    createCell value (Some true) None (Some 12.0) (Some "#FFFFFF") (Some "#4472C4")

/// Helper to create a numeric cell with format
let createNumericCell (value: string) (format: string) : FormattedCell =
    {
        Value = value
        Bold = None
        Italic = None
        Underline = None
        FontSize = None
        FontFamily = None
        FontColor = None
        FillColor = None
        NumberFormat = Some format
        Formula = None
    }

/// Helper to create a formula cell
let createFormulaCell (formula: string) (cachedValue: string) : FormattedCell =
    cellWithFormula formula cachedValue

/// Add merged range to sheet builder
let addMergedRange (sheetBuilder: SheetBuilder) (startRow: int) (startCol: int) (endRow: int) (endCol: int) =
    sheetBuilder.AddMergedRange(startRow, startCol, endRow, endCol)
