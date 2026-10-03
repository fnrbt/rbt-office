module Rbt.Spreadsheet.Model

open System

/// Cell address (0-indexed)
type CellAddress = {
    Row: int
    Col: int
}

/// Cell value types
type CellValue =
    | Empty
    | Number of float
    | String of string
    | Boolean of bool
    | Error of string
    | Formula of formula: string * cachedValue: CellValue

let unwrapFormulaValue (value: CellValue) : CellValue =
    let mutable current = value
    let mutable doneUnwrapping = false
    while not doneUnwrapping do
        match current with
        | Formula (_, cached) -> current <- cached
        | _ -> doneUnwrapping <- true
    current

/// Border style for cell edges
type BorderStyle = {
    Style: string  // "thin", "medium", "thick", "dashed", "dotted", "double"
    Color: string option
}

/// Cell formatting
type CellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    /// Distinguishes `<u val="double"/>` (two parallel strokes) from the
    /// default single underline. Renderers draw two strokes when true.
    UnderlineDouble: bool
    Strikethrough: bool
    FontName: string option
    TextColor: string option
    BgColor: string option
    FontSize: int option
    HorizontalAlign: string option  // "left", "center", "right", "justify"
    VerticalAlign: string option    // "top", "middle", "bottom"
    WrapText: bool
    /// Excel `<alignment indent="N"/>`. Each unit ≈ 3 character widths and
    /// pads inward from the cell's leading edge: from the left for left /
    /// distributed alignment, from the right for right alignment.
    Indent: int
    /// `<alignment shrinkToFit="1"/>` — when set, the cell's text is
    /// rendered at a reduced font size if its natural size would overflow
    /// the cell width. Excel never increases the size; only shrinks.
    ShrinkToFit: bool
    /// `<alignment textRotation="N"/>` degrees. 0 = horizontal (default),
    /// 1..90 = CCW (text rises left-to-right), 91..180 = CW (encoded as
    /// 90 + |angle|), 255 = stacked.
    TextRotation: int
    NumberFormat: string option
    BorderTop: BorderStyle option
    BorderBottom: BorderStyle option
    BorderLeft: BorderStyle option
    BorderRight: BorderStyle option
    /// Optional `<diagonal>` border style + color. Combined with
    /// DiagonalUp / DiagonalDown to choose which one or two diagonal
    /// strokes Excel paints across the cell rect.
    BorderDiagonal: BorderStyle option
    /// Stroke from bottom-left to top-right of the cell rectangle.
    DiagonalUp: bool
    /// Stroke from top-left to bottom-right of the cell rectangle.
    DiagonalDown: bool
    /// Excel auto-filter dropdown button on the right edge of the cell.
    /// Set for cells in the header row of a table with an <autoFilter> child.
    FilterButton: bool
    /// Small red triangle in the top-right corner of the cell. Excel
    /// renders this whenever the cell has a comment / note attached.
    /// Interactive viewers show the comment text as a tooltip on hover.
    HasComment: bool
    /// Small downward arrow on the right edge of the cell. Excel renders
    /// this for cells inside a `<dataValidation type="list">` range when
    /// the cell is selected; for static snapshot rendering we always show
    /// it so the visual cue carries through to PDF / PNG exports.
    HasListDropdown: bool
    /// Conditional-formatting data bar: (fillFraction 0..1, RRGGBB
    /// color). Renderers draw a horizontal bar from the cell's left
    /// edge spanning fraction × width, behind the text.
    DataBar: (float * string) option
    /// Conditional-formatting icon set: (iconIndex, iconCount). The
    /// renderer draws a small low/mid/high indicator at the cell's left.
    IconSet: (int * int) option
}

/// Default format
let defaultFormat = {
    Bold = false
    Italic = false
    Underline = false
    UnderlineDouble = false
    Strikethrough = false
    FontName = None
    TextColor = None
    BgColor = None
    FontSize = None
    HorizontalAlign = None
    VerticalAlign = None
    WrapText = false
    Indent = 0
    ShrinkToFit = false
    TextRotation = 0
    NumberFormat = None
    BorderTop = None
    BorderBottom = None
    BorderLeft = None
    BorderRight = None
    BorderDiagonal = None
    DiagonalUp = false
    DiagonalDown = false
    FilterButton = false
    HasComment = false
    HasListDropdown = false
    DataBar = None
    IconSet = None
}

/// Cell comment/note
type CellComment = {
    Text: string
    Author: string option
}

/// Vertical glyph alignment within a run / cell. Renderers shift the
/// baseline and scale the glyph size for super/subscript. Maps from
/// Excel's `<vertAlign val="superscript"/>` element.
type VertAlign =
    | VertAlignBaseline
    | VertAlignSuperscript
    | VertAlignSubscript

/// Per-run formatting override for rich-text cells. Each field is optional;
/// when None the run inherits from the cell's base CellFormat. Originates
/// from `<rPr>` children of `<r>` elements inside `<si>` shared strings.
type RunOverride = {
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineDouble: bool option
    Strikethrough: bool option
    Color: string option       // hex with leading '#'
    FontName: string option
    FontSize: float option     // points
    VertAlign: VertAlign option
}

let emptyRunOverride = {
    Bold = None
    Italic = None
    Underline = None
    UnderlineDouble = None
    Strikethrough = None
    Color = None
    FontName = None
    FontSize = None
    VertAlign = None
}

/// Cell data
type CellData = {
    Value: CellValue
    Format: CellFormat
    Comment: CellComment option
}

let private cellDataIsFormula (cell: CellData) =
    match cell.Value with
    | Formula _ -> true
    | _ -> false

let cellsHaveFormulaCells (cells: Map<CellAddress, CellData>) =
    cells |> Map.exists (fun _ cell -> cellDataIsFormula cell)

/// Selection types
type Selection =
    | SingleCell of row: int * col: int
    | Range of startRow: int * startCol: int * endRow: int * endCol: int

/// Clipboard data for copy/paste operations
type ClipboardData = {
    Cells: (int * int * CellData) list  // (relRow, relCol, data)
    IsCut: bool
    SourceStartRow: int
    SourceStartCol: int
}

/// Sort state for a column
type SortState = {
    Column: int
    Ascending: bool
}

/// Column filter (values to show - unchecked values are hidden)
type ColumnFilter = {
    Column: int
    VisibleValues: Set<string>  // String representation of values to show
}

/// Conditional formatting rule
type ConditionalFormatRule =
    | GreaterThan of float
    | LessThan of float
    | Between of float * float
    | Equal of float
    | ContainsText of string

/// Conditional format (rule + style to apply)
type ConditionalFormat = {
    Rule: ConditionalFormatRule
    Style: CellFormat
}

/// Data-validation rule (what a cell is allowed to contain).
type ValidationRule =
    /// Allowed values (a dropdown). Case-insensitive match.
    | ListValues of string list
    /// Numeric range [min, max] (inclusive); `None` = unbounded.
    | WholeNumber of float option * float option
    | DecimalRange of float option * float option
    /// String length range.
    | TextLength of float option * float option
    /// Date (Excel serial) range.
    | DateRange of float option * float option

type DataValidation = {
    Rule: ValidationRule
    /// Shown when an entered value fails the rule.
    ErrorMessage: string option
    /// Shown as a tooltip while the cell is selected.
    InputMessage: string option
}

/// Merged cell range (primary cell is top-left)
type MergedCell = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
}

// ---------------------------------------------------------------------------
// Print / page setup
// ---------------------------------------------------------------------------

type PrintOrientation =
    | PortraitOrientation
    | LandscapeOrientation

/// Walk order for 2D pagination. `DownThenOver` (Excel default) iterates
/// every row-slice within the first column-slice before moving on to the
/// next column-slice. `OverThenDown` does the inverse.
type PageOrder =
    | DownThenOver
    | OverThenDown

/// Print behavior for cell comments. Excel default is `CellCommentsNone`
/// (don't print). `CellCommentsAsDisplayed` shows the red corner marker,
/// `CellCommentsAtEnd` would list comments after the printed pages.
type CellCommentsPrint =
    | CellCommentsNone
    | CellCommentsAsDisplayed
    | CellCommentsAtEnd

/// Page margins, all in *inches* (matching how OOXML encodes them and how
/// Excel exposes them to users; the PDF backend converts to points).
type PageMargins = {
    Top: float
    Bottom: float
    Left: float
    Right: float
    Header: float
    Footer: float
}

/// Boolean print options from `<printOptions>`.
type PrintOptions = {
    /// Print gridlines under cell content.
    GridLines: bool
    /// Print row numbers and column letters as headings.
    Headings: bool
    /// Center the printed range horizontally on the page.
    HorizontalCentered: bool
    /// Center it vertically.
    VerticalCentered: bool
}

/// Header / footer strings carry Excel's `&L &C &R` section markers plus
/// field codes — `&P` = page number, `&N` = total pages, `&A` = sheet
/// name, `&F` = file name, `&D` = date, `&T` = time, `&&` = literal `&`.
/// The PDF backend substitutes these per page.
type HeaderFooter = {
    OddHeader: string option
    OddFooter: string option
    EvenHeader: string option
    EvenFooter: string option
    FirstHeader: string option
    FirstFooter: string option
    /// When true, EvenHeader/EvenFooter apply to even-numbered pages.
    DifferentOddEven: bool
    /// When true, FirstHeader/FirstFooter apply to the first page.
    DifferentFirst: bool
    /// When true (Excel default), H/F text scales with the page's print
    /// scale. When false, H/F always renders at its raw point size.
    ScaleWithDoc: bool
}

/// Inclusive 0-indexed cell range. Used for print area and print titles.
/// Distinct from `MergedCell` (same shape, different meaning) to keep the
/// type system honest and avoid record-literal ambiguity.
type CellRange = {
    From: CellAddress
    To: CellAddress
}

/// Excel table style reference (e.g. `TableStyleMedium7` + show-stripes
/// toggles). Mirrors `Xlsx.TableStyleInfo` so the model carries this
/// without forcing a dependency on the Office.Formats reader.
type TableStyle = {
    Name: string
    ShowFirstColumn: bool
    ShowLastColumn: bool
    ShowRowStripes: bool
    ShowColumnStripes: bool
    /// Pre-resolved header fill / banded-row fill (hex). Set for tables whose
    /// styling is applied at render time from `model.Tables` (e.g. tables made
    /// by `createTable`); left None for imported tables whose runner bakes the
    /// theme-accent styling straight into each cell's format. `applyTableStyles`
    /// only acts when `HeaderFill` is set, so it never double-styles the corpus.
    HeaderFill: string option
    StripeFill: string option
}

/// One table region on a sheet (i.e., `xl/tables/tableN.xml`). Coordinates
/// are 1-indexed to match the XLSX schema and the table-overlay helpers
/// that consume this. `HeaderRowCount` is 0 when the table has no header
/// row (rare); otherwise 1 (effectively always 1 for the corpus we
/// handle). `AutoFilter` is the 1-indexed (startCol, endCol) range that
/// gets dropdown buttons on its header cells. `Name` is the structured-ref
/// identifier; `HasTotals` marks a totals row occupying the last row.
type Table = {
    Name: string
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    HeaderRowCount: int
    Style: TableStyle option
    AutoFilter: (int * int) option
    HasTotals: bool
}

/// Per-sheet page setup. All fields have sensible defaults so a model
/// without explicit setup still has a usable PageSetup record. Page
/// dimensions are in points (1 pt = 1/72 inch).
type PageSetup = {
    PageWidth: float
    PageHeight: float
    Orientation: PrintOrientation
    Margins: PageMargins
    /// When `FitToPage` is true, the printed range is scaled to fit at
    /// most `FitToWidth` × `FitToHeight` pages.
    FitToPage: bool
    FitToWidth: int
    FitToHeight: int
    /// Manual scale percent (10..400). Used when `FitToPage` is false.
    Scale: int
    Options: PrintOptions
    HeaderFooter: HeaderFooter
    /// Manual page breaks: 1-based row / column numbers; the break sits
    /// *after* the given row / column. Empty when there are no manuals.
    ManualRowBreaks: int list
    ManualColBreaks: int list
    /// 0-indexed inclusive range to print; None = print the whole used
    /// range. From `_xlnm.Print_Area`.
    PrintArea: CellRange option
    /// Rows to repeat at the top of every page (inclusive, 0-indexed).
    /// From the row part of `_xlnm.Print_Titles`. None = no repeat.
    RepeatRows: (int * int) option
    /// Columns to repeat on the left of every page (inclusive, 0-indexed).
    /// From the col part of `_xlnm.Print_Titles`.
    RepeatCols: (int * int) option
    /// Page number to display for the first printed page; `&P` substitutes
    /// this value on page 1 and increments from there. Excel-default is
    /// 1; only honored when `UseFirstPageNumber` is true.
    FirstPageNumber: int
    /// `useFirstPageNumber="1"` opts in to FirstPageNumber; when false,
    /// Excel ignores the value and numbers from 1.
    UseFirstPageNumber: bool
    /// Direction of 2D pagination walk. Default DownThenOver.
    PageOrder: PageOrder
    /// When true, every fill renders as white and every text / border /
    /// stroke renders as black. Matches Excel's `blackAndWhite="1"` print
    /// option.
    BlackAndWhite: bool
    /// How cell comments appear in print output. Default `CellCommentsNone`
    /// matches Excel's default of suppressing comments in print.
    CellComments: CellCommentsPrint
}

let defaultPageMargins = {
    Top = 0.75; Bottom = 0.75; Left = 0.7; Right = 0.7
    Header = 0.3; Footer = 0.3
}

let defaultPrintOptions = {
    GridLines = false; Headings = false
    HorizontalCentered = false; VerticalCentered = false
}

let defaultHeaderFooter = {
    OddHeader = None; OddFooter = None
    EvenHeader = None; EvenFooter = None
    FirstHeader = None; FirstFooter = None
    DifferentOddEven = false; DifferentFirst = false
    ScaleWithDoc = true
}

let defaultPageSetup = {
    // US Letter at 72 DPI = 612 × 792 pt.
    PageWidth = 612.0
    PageHeight = 792.0
    Orientation = PortraitOrientation
    Margins = defaultPageMargins
    FitToPage = false
    FitToWidth = 1
    FitToHeight = 1
    Scale = 100
    Options = defaultPrintOptions
    HeaderFooter = defaultHeaderFooter
    ManualRowBreaks = []
    ManualColBreaks = []
    PrintArea = None
    FirstPageNumber = 1
    UseFirstPageNumber = false
    PageOrder = DownThenOver
    BlackAndWhite = false
    CellComments = CellCommentsNone
    RepeatRows = None
    RepeatCols = None
}

/// Standard OOXML paper-size codes → (width, height) in points (portrait).
/// Defaults to Letter when the code is unknown.
let paperSizePoints (code: int) : float * float =
    match code with
    | 1 -> 612.0, 792.0      // Letter   (8.5 × 11 in)
    | 5 -> 612.0, 1008.0     // Legal    (8.5 × 14 in)
    | 8 -> 841.68, 1190.55   // A3       (297 × 420 mm)
    | 9 -> 595.27, 841.89    // A4       (210 × 297 mm)
    | 11 -> 419.53, 595.27   // A5       (148 × 210 mm)
    | 13 -> 728.5, 1031.81   // B5 ISO   (176 × 250 mm)
    | 70 -> 595.27, 841.89   // A4 alt
    | _ -> 612.0, 792.0      // unknown → Letter

/// Spreadsheet model (forward declaration for Command type)
type SpreadsheetModel = {
    /// Sparse cell storage - only non-empty cells
    Cells: Map<CellAddress, CellData>

    /// Current selection
    Selection: Selection

    /// Currently editing cell
    EditingCell: CellAddress option

    /// Edit buffer (what user is typing)
    EditBuffer: string

    /// Grid size
    RowCount: int
    ColCount: int

    /// Clipboard data for copy/paste
    Clipboard: ClipboardData option

    /// History for undo/redo
    History: History

    /// Current sort state
    SortState: SortState option

    /// Active filters by column
    Filters: Map<int, ColumnFilter>

    /// Rows visible under the active AutoFilter projection. None means the
    /// projection was invalidated by a data mutation and reads should fall
    /// back to recomputing from the current cells.
    FilterVisibleRows: Set<int> option

    /// Rows hidden by the active AutoFilter projection.
    FilterHiddenRows: int array

    /// Custom row heights (default 25.0)
    RowHeights: Map<int, float>

    /// Custom column widths (default 100.0)
    ColumnWidths: Map<int, float>

    /// Current search query
    SearchQuery: string option

    /// Case-sensitivity used for the current search results.
    SearchCaseSensitive: bool option

    /// Search results (cell addresses)
    SearchResults: CellAddress array

    /// Current search result index
    SearchIndex: int

    /// Conditional formats by cell address
    ConditionalFormats: Map<CellAddress, ConditionalFormat list>

    /// Merged cell ranges
    MergedCells: MergedCell list

    /// Rich-text runs keyed by cell address. Set for cells whose value comes
    /// from a shared-string entry with multiple `<r>` runs (per-run color,
    /// bold, font). Each entry is the list of (text, override) pairs in
    /// source order; the runs concatenate to the plain cell value.
    RichRuns: Map<CellAddress, (string * RunOverride) list>

    /// Number of rows that stay pinned at the top while the body scrolls
    /// vertically. Set from `<pane ySplit="N" state="frozen"/>` in the
    /// sheet's sheetView. Zero = no freeze.
    FrozenRows: int

    /// Number of columns that stay pinned on the left while the body
    /// scrolls horizontally. Set from `<pane xSplit="N" state="frozen"/>`.
    /// Zero = no freeze.
    FrozenCols: int

    /// Print / page setup for this sheet. Always present (defaults when
    /// the source file has no explicit `<pageSetup>` etc.).
    PageSetup: PageSetup

    /// Table regions defined on the sheet (`xl/tables/tableN.xml`). The
    /// runner translates `Xlsx.TableDefinition` here so renderers and the
    /// layout harness can detect table-style overlays from the model
    /// alone, without needing the original workbook handle.
    Tables: Table list

    /// Cells where the Bold flag came from a conditional-formatting rule
    /// (dxf) rather than the cell's own font. The runner sets these when
    /// it bakes a CF rule's `<font><b/></font>` into the cell's format —
    /// the layout harness reads them so it can report the underlying
    /// font's bold (which is what Excel COM `cell.Font.Bold` returns,
    /// regardless of any CF overlay).
    CfBoldCells: Set<CellAddress>

    /// Whether the sheet shows the light grey screen gridlines. From the
    /// sheetView `showGridLines` attribute (default true). Many polished
    /// workbooks turn this off (`showGridLines="0"`) and rely on explicit
    /// cell borders, so the renderer must suppress the always-on grid.
    ShowGridLines: bool

    /// Dynamic-array spill ownership: each spilled (ghost) cell mapped to its
    /// anchor formula cell. Maintained by the recalc+spill pass so stale
    /// spilled cells are cleared before a re-spill. Empty for non-spilling
    /// workbooks.
    Spill: Map<CellAddress, CellAddress>

    /// Data-validation rules keyed by cell.
    Validations: Map<CellAddress, DataValidation>

    /// VBA/VBScript module source attached to the workbook (the "VBA project").
    /// Empty when there are no macros. Event subs (Worksheet_Change /
    /// Workbook_Open) and user-run macros live here; `Vba.runOnModel` executes
    /// it. Persisted/round-tripped separately from cell data.
    Macros: string
}

/// Command for undo/redo (must come after SpreadsheetModel)
and Command = {
    Undo: SpreadsheetModel -> SpreadsheetModel
    Redo: SpreadsheetModel -> SpreadsheetModel
    Description: string
}

/// History for undo/redo
and History = {
    Past: Command list      // Undo stack
    Future: Command list    // Redo stack
    MaxSize: int           // Limit memory (default 50)
}

/// Empty history
let emptyHistory = {
    Past = []
    Future = []
    MaxSize = 50
}

/// Create empty spreadsheet
let createEmpty rowCount colCount = {
    Cells = Map.empty
    Selection = SingleCell(0, 0)
    EditingCell = None
    EditBuffer = ""
    RowCount = rowCount
    ColCount = colCount
    Clipboard = None
    History = emptyHistory
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
    MergedCells = []
    RichRuns = Map.empty
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

/// Convert column index to letter (0="A", 25="Z", 26="AA")
let rec columnToLetter (col: int): string =
    if col < 0 then ""
    elif col < 26 then
        string (char (int 'A' + col))
    else
        let quotient = (col / 26) - 1
        let remainder = col % 26
        columnToLetter quotient + columnToLetter remainder

/// Convert column letter to index ("A"=0, "Z"=25, "AA"=26)
let letterToColumn (letter: string): int =
    let mutable result = 0
    for c in letter.ToUpper() do
        result <- result * 26 + (int c - int 'A' + 1)
    result - 1

let tryParseNumber (text: string): float option =
    match Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
    | true, value -> Some value
    | false, _ -> None

/// Parse cell reference like "A1" to (row, col)
let parseCellRef (ref: string): (int * int) option =
    try
        let mutable colPart = ""
        let mutable rowPart = ""
        let mutable inRow = false
        let mutable valid = true

        for c in ref do
            if Char.IsDigit(c) then
                inRow <- true
                rowPart <- rowPart + string c
            elif Char.IsLetter(c) && not inRow then
                colPart <- colPart + string c
            else
                valid <- false

        if not valid || String.IsNullOrEmpty(colPart) || String.IsNullOrEmpty(rowPart) then
            None
        else
            let col = letterToColumn colPart
            let row = Int32.Parse(rowPart) - 1  // 1-indexed to 0-indexed
            if row < 0 || col < 0 then None else Some (row, col)
    with
    | _ -> None

/// Format cell reference like (0, 0) to "A1"
let formatCellRef (row: int) (col: int): string =
    sprintf "%s%d" (columnToLetter col) (row + 1)

/// Get cell data at address
let getCell (addr: CellAddress) (model: SpreadsheetModel): CellData option =
    Map.tryFind addr model.Cells

let invalidateFilterProjection (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterVisibleRows.IsNone && model.FilterHiddenRows.Length = 0 then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    else
        { model with FilterVisibleRows = None; FilterHiddenRows = [||] }

/// Set cell data at address
let setCell (addr: CellAddress) (data: CellData) (model: SpreadsheetModel): SpreadsheetModel =
    match Map.tryFind addr model.Cells with
    | Some existing when existing = data -> model
    | _ -> { model with Cells = Map.add addr data model.Cells } |> invalidateFilterProjection

let setCellFormat (addr: CellAddress) (update: CellFormat -> CellFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let existing = getCell addr model |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
    let updatedFormat = update existing.Format
    if updatedFormat = existing.Format then
        model
    else
        setCell addr { existing with Format = updatedFormat } model

let setCellTextColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with TextColor = Some color }) model

let setCellBgColor (addr: CellAddress) color (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with BgColor = Some color }) model

let setCellNumberFormat (addr: CellAddress) numberFormat (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with NumberFormat = Some numberFormat }) model

let setCellHorizontalAlign (addr: CellAddress) align (model: SpreadsheetModel): SpreadsheetModel =
    setCellFormat addr (fun f -> { f with HorizontalAlign = Some align }) model

let private borderMatches style color (border: BorderStyle option) =
    match border with
    | Some existing -> existing.Style = style && existing.Color = Some color
    | None -> false

let private hasRequestedBorder side style color (format: CellFormat) =
    match side with
    | "top" -> borderMatches style color format.BorderTop
    | "bottom" -> borderMatches style color format.BorderBottom
    | "left" -> borderMatches style color format.BorderLeft
    | "right" -> borderMatches style color format.BorderRight
    | "all" ->
        borderMatches style color format.BorderTop
        && borderMatches style color format.BorderBottom
        && borderMatches style color format.BorderLeft
        && borderMatches style color format.BorderRight
    | _ -> true

let private applyBorderToFormat side border (format: CellFormat) =
    match side with
    | "top" -> { format with BorderTop = Some border }
    | "bottom" -> { format with BorderBottom = Some border }
    | "left" -> { format with BorderLeft = Some border }
    | "right" -> { format with BorderRight = Some border }
    | "all" ->
        { format with
            BorderTop = Some border
            BorderBottom = Some border
            BorderLeft = Some border
            BorderRight = Some border }
    | _ -> format

let setCellBorder (addr: CellAddress) (side: string) (style: string) (color: string) (model: SpreadsheetModel): SpreadsheetModel =
    let existing = getCell addr model |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
    if hasRequestedBorder side style color existing.Format then
        model
    else
        let border = { Style = style; Color = Some color }
        setCell addr { existing with Format = applyBorderToFormat side border existing.Format } model

/// Clear cell at address
let clearCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    if Map.containsKey addr model.Cells then
        { model with Cells = Map.remove addr model.Cells } |> invalidateFilterProjection
    else
        model

/// Clear all cells in a selection, preserving identity when nothing changes.
let clearSelectionCells (selection: Selection) (model: SpreadsheetModel): SpreadsheetModel =
    match selection with
    | SingleCell (row, col) ->
        clearCell { Row = row; Col = col } model
    | Range (startRow, startCol, endRow, endCol) ->
        let mutable changed = false
        let cells =
            model.Cells
            |> Map.filter (fun addr _ ->
                let remove =
                    addr.Row >= startRow && addr.Row <= endRow &&
                    addr.Col >= startCol && addr.Col <= endCol
                if remove then changed <- true
                not remove)
        if changed then { model with Cells = cells } |> invalidateFilterProjection else model

/// Round to the nearest integer, halves away from zero (Excel's rule). Written
/// out because Fable ignores MidpointRounding and reads it as a digit count.
let roundHalfAway (x: float) : float =
    let whole = Math.Truncate x
    if abs (x - whole) >= 0.5 then whole + float (sign x) else whole

/// Excel's day zero. UTC-kind so that day arithmetic and calendar fields do
/// not shift across DST changes under Fable, where other DateTimes are local.
let excelEpoch = DateTime(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc)

/// Format a number according to format code
let formatNumber (format: string option) (value: float): string =
    match format with
    | Some "currency" ->
        // Currency format: $1,234.56
        sprintf "$%s" (NumberText.numberToString "N2" value)
    | Some "percentage" | Some "percent" ->
        // Percentage format: 45.67%
        sprintf "%s%%" (NumberText.numberToString "N2" (value * 100.0))
    | Some "date" ->
        try
            let date = excelEpoch.AddDays(value)
            sprintf "%02d/%02d/%04d" date.Month date.Day date.Year
        with
        | _ -> NumberText.numberToString "F2" value
    | Some "decimal2" ->
        // 2 decimal places: 123.46
        NumberText.numberToString "F2" value
    | Some "number" ->
        // General number with comma separators
        NumberText.numberToString "N2" value
    | _ ->
        // Default: just show the number
        string value

/// Display value for cell (handles formulas) - needs to be before getCellValue
let rec displayValue (value: CellValue): string =
    match value with
    | Empty -> ""
    | Number n -> NumberText.numberToString "G" n
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> displayValue cached

/// Formula text as the user edits it. Cells typed in the web editor store the
/// leading "="; imported cells store the bare expression.
let formulaText (text: string) : string =
    if text.StartsWith "=" then text else "=" + text

/// Display value with the legacy named formats applied. Real Excel format
/// codes need CellText.cellDisplayText, which compiles after this module.
let displayValueWithFormat (value: CellValue) (format: CellFormat): string =
    match value with
    | Number n -> formatNumber format.NumberFormat n
    | _ -> displayValue value

/// Get cell value (returns Empty if not found)
let getCellValue (addr: CellAddress) (model: SpreadsheetModel): CellValue =
    match getCell addr model with
    | Some cell -> cell.Value
    | None -> Empty

let private dataRows (model: SpreadsheetModel): Set<int> =
    model.Cells
    |> Map.toSeq
    |> Seq.map (fun (addr, _) -> addr.Row)
    |> Set.ofSeq

/// The block of rows a column-level operation (sort, AutoFilter) acts on, as
/// (first row, last row). Rows with content form unbroken runs separated by
/// empty rows (Excel's "current region"); the block is the run holding most of
/// the column's entries, so a stray title or note elsewhere on the sheet is
/// not mistaken for the data. A table covering that run wins (its totals row
/// excluded). The block's first row is where a header would sit. None for an
/// empty sheet.
let dataBlockForColumn (col: int) (model: SpreadsheetModel) : (int * int) option =
    let mutable rows = Set.empty
    let mutable columnRows = Set.empty
    for KeyValue(addr, cellData) in model.Cells do
        if cellData.Value <> Empty then
            rows <- Set.add addr.Row rows
            if addr.Col = col then columnRows <- Set.add addr.Row columnRows
    if Set.isEmpty rows then None
    else
        let mutable best = None
        let mutable bestCount = -1
        let mutable runStart = -1
        let mutable previous = -2
        let mutable count = 0
        let close () =
            if runStart >= 0 && count > bestCount then
                best <- Some (runStart, previous)
                bestCount <- count
        for row in rows do
            if row <> previous + 1 then
                close ()
                runStart <- row
                count <- 0
            if Set.contains row columnRows then count <- count + 1
            previous <- row
        close ()
        match best with
        | None -> None
        | Some (first, last) ->
            let table =
                model.Tables
                |> List.tryFind (fun t ->
                    col >= t.StartCol - 1 && col <= t.EndCol - 1
                    && t.StartRow - 1 <= last && t.EndRow - 1 >= first)
            match table with
            | Some t -> Some (t.StartRow - 1, t.EndRow - 1 - (if t.HasTotals then 1 else 0))
            | None -> Some (first, last)

/// The block the active AutoFilter applies to: the one around its leftmost
/// filtered column.
let private filterBlock (model: SpreadsheetModel) : (int * int) option =
    if Map.isEmpty model.Filters then None
    else
        let firstFilterCol = model.Filters |> Map.toSeq |> Seq.map fst |> Seq.min
        dataBlockForColumn firstFilterCol model

let private computeVisibleRows (model: SpreadsheetModel): Set<int> =
    let rows = dataRows model
    if Map.isEmpty model.Filters then
        rows
    else
        let inBlock =
            match filterBlock model with
            | Some (first, last) -> fun row -> row >= first && row <= last
            | None -> fun _ -> false
        rows
        |> Set.filter (fun row ->
            // Rows outside the filtered block are never filtered.
            not (inBlock row) ||
            model.Filters
            |> Map.forall (fun col filter ->
                let addr = { Row = row; Col = col }
                let value = getCellValue addr model
                let displayStr = displayValue value
                Set.contains displayStr filter.VisibleValues))

let private hiddenRowsFromVisible (visible: Set<int>) (model: SpreadsheetModel): int array =
    match filterBlock model with
    | None -> [||]
    | Some (headerRow, lastRow) ->
        // The block's first row is its header and always stays.
        dataRows model
        |> Set.filter (fun row -> row > headerRow && row <= lastRow && not (Set.contains row visible))
        |> Set.toArray

let private refreshFilterProjection (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterVisibleRows.IsNone && model.FilterHiddenRows.Length = 0 then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    else
        let visible = computeVisibleRows model
        let hidden = hiddenRowsFromVisible visible model
        { model with FilterVisibleRows = Some visible; FilterHiddenRows = hidden }

/// A row/column insertion or deletion. `at` is the 0-indexed first row/column
/// affected; `count` rows/columns are inserted before it or deleted from it.
type StructureEdit =
    | InsertRows of at: int * count: int
    | DeleteRows of at: int * count: int
    | InsertCols of at: int * count: int
    | DeleteCols of at: int * count: int

/// Formula-text rewrites. They need the formula parser and printer, which
/// compile after this module (Formula.fs), so Formula.fs installs them here.
type FormulaHooks = {
    /// Shift every relative reference by (rows, cols) — copy/paste and fill.
    Shift: int -> int -> string -> string
    /// Re-point references after a structure edit. The predicate says whether
    /// a reference's sheet qualifier (None = unqualified) names the edited sheet.
    Adjust: StructureEdit -> (string option -> bool) -> string -> string
}

let mutable private formulaHooks : FormulaHooks option = None

/// Called once by Formula.fs when it initialises.
let installFormulaHooks (hooks: FormulaHooks) : unit =
    formulaHooks <- Some hooks

let private currentFormulaHooks () : FormulaHooks option =
    match formulaHooks with
    | Some _ -> formulaHooks
    | None ->
#if !FABLE_COMPILER
        // A .NET library initialises a file only when one of its values is
        // first read, so a host that has not evaluated a formula yet would
        // find no hooks. Reading Formula.formulaHooksInstalled runs that
        // file's initialiser, which installs them.
        try
            let formulaModule = typeof<CellAddress>.Assembly.GetType "Rbt.Spreadsheet.Formula"
            if not (isNull formulaModule) then
                let installed = formulaModule.GetProperty "formulaHooksInstalled"
                if not (isNull installed) then
                    installed.GetValue(null, null) |> ignore
        with _ -> ()
#endif
        formulaHooks

/// Shift the relative references of a formula by (deltaRow, deltaCol), the way
/// copying it to another cell does. `$`-locked parts stay put; a reference
/// pushed off the sheet becomes #REF!. Text that does not parse is returned
/// unchanged.
let shiftFormulaReferences (deltaRow: int) (deltaCol: int) (formula: string) : string =
    if deltaRow = 0 && deltaCol = 0 then formula
    else
        match currentFormulaHooks () with
        | Some hooks -> hooks.Shift deltaRow deltaCol formula
        | None -> formula

/// Re-point a formula's references after rows/columns were inserted or
/// deleted on the sheet selected by `targetsEditedSheet`.
let adjustFormulaForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (formula: string) : string =
    match currentFormulaHooks () with
    | Some hooks -> hooks.Adjust edit targetsEditedSheet formula
    | None -> formula

/// Serialize selection to clipboard data
let serializeSelection (selection: Selection) (model: SpreadsheetModel) (isCut: bool): ClipboardData =
    match selection with
    | SingleCell (row, col) ->
        let addr = { Row = row; Col = col }
        let cellData = getCell addr model
        match cellData with
        | Some data ->
            { Cells = [(0, 0, data)]; IsCut = isCut; SourceStartRow = row; SourceStartCol = col }
        | None ->
            { Cells = []; IsCut = isCut; SourceStartRow = row; SourceStartCol = col }

    | Range (startRow, startCol, endRow, endCol) ->
        let cells =
            [startRow .. endRow]
            |> List.collect (fun row ->
                [startCol .. endCol]
                |> List.map (fun col ->
                    let addr = { Row = row; Col = col }
                    let relRow = row - startRow
                    let relCol = col - startCol
                    match getCell addr model with
                    | Some data -> Some (relRow, relCol, data)
                    | None -> None
                )
            )
            |> List.choose id

        { Cells = cells; IsCut = isCut; SourceStartRow = startRow; SourceStartCol = startCol }

/// Paste clipboard data at target address
let pasteClipboard (targetAddr: CellAddress) (clipboardData: ClipboardData) (model: SpreadsheetModel): SpreadsheetModel =
    let deltaRow = targetAddr.Row - clipboardData.SourceStartRow
    let deltaCol = targetAddr.Col - clipboardData.SourceStartCol

    let mutable updatedModel = model

    for (relRow, relCol, cellData) in clipboardData.Cells do
        let pasteRow = targetAddr.Row + relRow
        let pasteCol = targetAddr.Col + relCol

        if pasteRow >= 0 && pasteRow < model.RowCount && pasteCol >= 0 && pasteCol < model.ColCount then
            let pasteAddr = { Row = pasteRow; Col = pasteCol }

            // A copied formula follows its new position; a cut formula is moved
            // as written (Excel keeps a moved formula pointing at the same cells).
            let adjustedValue =
                match cellData.Value with
                | Formula (formula, _) when not clipboardData.IsCut ->
                    Formula (shiftFormulaReferences deltaRow deltaCol formula, Empty)
                | Formula (formula, _) ->
                    Formula (formula, Empty)
                | other -> other

            let adjustedCellData = { cellData with Value = adjustedValue }
            updatedModel <- setCell pasteAddr adjustedCellData updatedModel

    refreshFilterProjection updatedModel

/// Push command to history (clears future stack)
let pushCommand (command: Command) (model: SpreadsheetModel): SpreadsheetModel =
    let newPast = command :: model.History.Past
    // Limit stack size
    let trimmedPast =
        if List.length newPast > model.History.MaxSize then
            List.take model.History.MaxSize newPast
        else
            newPast

    let newHistory = {
        model.History with
            Past = trimmedPast
            Future = []  // Clear future on new command
    }
    { model with History = newHistory }

/// Undo last command
let undo (model: SpreadsheetModel): SpreadsheetModel option =
    match model.History.Past with
    | [] -> None  // Nothing to undo
    | command :: restPast ->
        // Execute undo
        let undoneModel = command.Undo model

        // Move command to future stack
        let newHistory = {
            undoneModel.History with
                Past = restPast
                Future = command :: undoneModel.History.Future
        }

        Some { undoneModel with History = newHistory }

/// Redo last undone command
let redo (model: SpreadsheetModel): SpreadsheetModel option =
    match model.History.Future with
    | [] -> None  // Nothing to redo
    | command :: restFuture ->
        // Execute redo
        let redoneModel = command.Redo model

        // Move command back to past stack
        let newHistory = {
            redoneModel.History with
                Past = command :: redoneModel.History.Past
                Future = restFuture
        }

        Some { redoneModel with History = newHistory }

/// Compare cell values for sorting (numbers < strings < booleans < empty)
let rec compareCellValues (v1: CellValue) (v2: CellValue): int =
    match v1, v2 with
    | Formula (_, cached1), Formula (_, cached2) -> compareCellValues cached1 cached2
    | Formula (_, cached), other -> compareCellValues cached other
    | other, Formula (_, cached) -> compareCellValues other cached
    | Number n1, Number n2 -> compare n1 n2
    | String s1, String s2 -> System.String.Compare(s1, s2, System.StringComparison.OrdinalIgnoreCase)
    | Boolean b1, Boolean b2 -> compare b1 b2
    | Error _, Error _ -> 0
    | Empty, Empty -> 0
    | _ ->
        let rec rank = function
            | Number _ -> 0
            | String _ -> 1
            | Boolean _ -> 2
            | Error _ -> 3
            | Empty -> 4
            | Formula (_, cached) -> rank cached
        compare (rank v1) (rank v2)

/// Sort the data block around a column by that column, in place. The block's
/// rows keep their position on the sheet (nothing is compacted to the top) and
/// rows outside it are untouched. A first row made only of text is taken as
/// the header and stays first. Blanks sort last in either direction, ties keep
/// their order, and a formula that moves keeps pointing at its own row (its
/// relative references shift with it). Cached values of other formulas are
/// stale afterwards — recalculate.
let sortByColumn (col: int) (ascending: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let sorted = { model with SortState = Some { Column = col; Ascending = ascending } }
    match dataBlockForColumn col model with
    | None -> refreshFilterProjection sorted
    | Some (firstRow, lastRow) ->
        let rowCells =
            model.Cells
            |> Map.toList
            |> List.filter (fun (addr, _) -> addr.Row >= firstRow && addr.Row <= lastRow)
        let isTable =
            model.Tables |> List.exists (fun t -> t.StartRow - 1 = firstRow && col >= t.StartCol - 1 && col <= t.EndCol - 1 && t.HeaderRowCount > 0)
        let headerCells = rowCells |> List.filter (fun (addr, cellData) -> addr.Row = firstRow && cellData.Value <> Empty)
        let hasHeader =
            lastRow > firstRow
            && (isTable
                || (not (List.isEmpty headerCells)
                    && headerCells |> List.forall (fun (_, cellData) -> match cellData.Value with String _ -> true | _ -> false)))
        let dataStart = if hasHeader then firstRow + 1 else firstRow
        let isBlank row =
            match getCellValue { Row = row; Col = col } model |> unwrapFormulaValue with
            | Empty | String "" -> true
            | _ -> false
        let compareRows a b =
            match isBlank a, isBlank b with
            | true, true -> 0
            | true, false -> 1
            | false, true -> -1
            | false, false ->
                let va = getCellValue { Row = a; Col = col } model
                let vb = getCellValue { Row = b; Col = col } model
                if ascending then compareCellValues va vb else compareCellValues vb va
        let order = [ dataStart .. lastRow ] |> List.sortWith compareRows
        // order.[i] is the source row that lands on row dataStart + i.
        let target = order |> List.mapi (fun i source -> source, dataStart + i) |> Map.ofList
        let moveMap (items: Map<CellAddress, 'T>) (onMove: int -> 'T -> 'T) : Map<CellAddress, 'T> =
            let moving, staying =
                items |> Map.partition (fun addr _ -> addr.Row >= dataStart && addr.Row <= lastRow)
            moving
            |> Map.fold (fun acc addr value ->
                let row = Map.find addr.Row target
                Map.add { addr with Row = row } (onMove (row - addr.Row) value) acc) staying
        let cells =
            moveMap model.Cells (fun delta cellData ->
                match cellData.Value with
                | Formula (formula, cached) when delta <> 0 ->
                    { cellData with Value = Formula (shiftFormulaReferences delta 0 formula, cached) }
                | _ -> cellData)
        { sorted with
            Cells = cells
            RichRuns = moveMap model.RichRuns (fun _ runs -> runs) }
        |> refreshFilterProjection

/// Get unique values in a column (as strings for display)
let getColumnUniqueValues (col: int) (model: SpreadsheetModel): string list =
    model.Cells
    |> Map.toList
    |> List.filter (fun (addr, _) -> addr.Col = col)
    |> List.map (fun (addr, cellData) -> displayValue cellData.Value)
    |> List.distinct
    |> List.sort

/// Apply filter to a column
let applyFilter (col: int) (visibleValues: Set<string>) (model: SpreadsheetModel): SpreadsheetModel =
    let filter = { Column = col; VisibleValues = visibleValues }
    { model with Filters = Map.add col filter model.Filters }
    |> refreshFilterProjection

/// Remove filter from a column
let removeFilter (col: int) (model: SpreadsheetModel): SpreadsheetModel =
    { model with Filters = Map.remove col model.Filters }
    |> refreshFilterProjection

/// Get visible rows based on all active filters
let getVisibleRows (model: SpreadsheetModel): Set<int> =
    match model.FilterVisibleRows with
    | Some visible when not (Map.isEmpty model.Filters) -> visible
    | _ -> computeVisibleRows model

let private binarySearchInt (values: int array) target =
    let mutable lo = 0
    let mutable hi = values.Length - 1
    let mutable found = -1
    while found < 0 && lo <= hi do
        let mid = lo + (hi - lo) / 2
        let value = values.[mid]
        if value = target then found <- mid
        elif value < target then lo <- mid + 1
        else hi <- mid - 1
    found

let isFilterHiddenRow (row: int) (model: SpreadsheetModel): bool =
    not (Map.isEmpty model.Filters) && binarySearchInt model.FilterHiddenRows row >= 0

let getVisibleRowsHashSet (model: SpreadsheetModel): System.Collections.Generic.HashSet<int> =
    System.Collections.Generic.HashSet<int>(getVisibleRows model |> Set.toArray)

/// Render-time model for active AutoFilters: filtered-out data rows get
/// height 0 so the grid collapses them, exactly like Excel's hidden rows —
/// the surviving rows keep their original numbers and draw contiguously
/// because rowY/visibleRange accumulate getRowHeight. The first row of the
/// data block is treated as the header (AutoFilter range convention) and is
/// never hidden. A strict no-op when no filters are active, so unfiltered
/// sheets (the whole parity corpus) are untouched.
let applyFilterRowHeights (model: SpreadsheetModel): SpreadsheetModel =
    if Map.isEmpty model.Filters then
        if model.FilterVisibleRows.IsNone && model.FilterHiddenRows.Length = 0 then model
        else { model with FilterVisibleRows = None; FilterHiddenRows = [||] }
    else
        match model.FilterVisibleRows with
        | Some _ -> model
        | None -> refreshFilterProjection model

/// Distinct, sorted display strings in a column — the AutoFilter checklist's
/// values. Blanks excluded; uses the same `displayValue` as the filter test.
let columnUniqueValues (col: int) (model: SpreadsheetModel) : string list =
    model.Cells
    |> Map.toList
    |> List.choose (fun (addr, _) ->
        if addr.Col = col then
            match displayValue (getCellValue addr model) with
            | "" -> None
            | s -> Some s
        else None)
    |> List.distinct
    |> List.sort

let applyFilterValues (col: int) (visibleValues: string array) (model: SpreadsheetModel): SpreadsheetModel =
    applyFilter col (Set.ofArray visibleValues) model

let columnUniqueValueCount (col: int) (model: SpreadsheetModel): int =
    columnUniqueValues col model |> List.length

let visibleRowCount (model: SpreadsheetModel): int =
    getVisibleRows model |> Set.count

let cellCount (model: SpreadsheetModel): int =
    model.Cells.Count

/// Create / register an Excel table over a 0-indexed region (header row r0,
/// data rows r0+1..r1). Blank header cells are filled with "Column N";
/// the table is registered in `model.Tables` (1-indexed) with a
/// TableStyleMedium2 style + an autofilter across every column. When
/// `withTotals`, a totals row is appended below the data: a "Total" label in
/// the first column and `=SUBTOTAL(109,<col>)` in the rest, and the table
/// range grows to include it. Returns the updated model. Structured refs over
/// the result work once it's bridged into the workbook (Workbook.makeTable…).
let createTable (name: string) (r0: int) (c0: int) (r1: int) (c1: int)
                (withTotals: bool) (model: SpreadsheetModel): SpreadsheetModel =
    // Fill any blank header cell with a default column name.
    let mutable m = model
    for c in c0 .. c1 do
        let addr = { Row = r0; Col = c }
        match getCellValue addr m with
        | Empty | String "" ->
            m <- setCell addr { Value = String (sprintf "Column %d" (c - c0 + 1)); Format = defaultFormat; Comment = None } m
        | _ -> ()
    // Optional totals row of SUBTOTAL(109,…) formulas, just below the data.
    let totalsRow = r1 + 1
    if withTotals then
        for c in c0 .. c1 do
            let addr = { Row = totalsRow; Col = c }
            if c = c0 then
                m <- setCell addr { Value = String "Total"; Format = defaultFormat; Comment = None } m
            else
                let col = columnToLetter c
                let formula = sprintf "SUBTOTAL(109,%s%d:%s%d)" col (r0 + 2) col (r1 + 1)
                m <- setCell addr { Value = Formula (formula, Number 0.0); Comment = None; Format = defaultFormat } m
    let endRow0 = if withTotals then totalsRow else r1
    let table : Table =
        { Name = name
          StartRow = r0 + 1
          StartCol = c0 + 1
          EndRow = endRow0 + 1
          EndCol = c1 + 1
          HeaderRowCount = 1
          Style = Some { Name = "TableStyleMedium2"; ShowFirstColumn = false; ShowLastColumn = false; ShowRowStripes = true; ShowColumnStripes = false
                         HeaderFill = Some "#4472C4"; StripeFill = Some "#D9E1F2" }
          AutoFilter = Some (c0 + 1, c1 + 1)
          HasTotals = withTotals }
    { m with Tables = m.Tables @ [ table ] }

/// The table whose region contains the 0-indexed cell, if any. Table
/// coordinates are 1-indexed, so they're shifted for the comparison.
let tableAt (row: int) (col: int) (model: SpreadsheetModel): Table option =
    model.Tables
    |> List.tryFind (fun t ->
        row >= t.StartRow - 1 && row <= t.EndRow - 1 && col >= t.StartCol - 1 && col <= t.EndCol - 1)

/// Distinct, sorted display values in one column of a table — data rows only
/// (header and totals excluded). The checklist for a table's filter button.
let tableColumnUniqueValues (table: Table) (col: int) (model: SpreadsheetModel): string list =
    let dataStart = (table.StartRow - 1) + table.HeaderRowCount
    let dataEnd = (table.EndRow - 1) - (if table.HasTotals then 1 else 0)
    [ for r in dataStart .. dataEnd -> displayValue (getCellValue { Row = r; Col = col } model) ]
    |> List.filter (fun s -> s <> "")
    |> List.distinct
    |> List.sort

/// Render-time: bake table styling (header accent fill + white bold text +
/// filter buttons, and banded data rows) into cell formats from `model.Tables`.
/// Acts ONLY on tables whose `Style.HeaderFill` is set — i.e. tables created in
/// the app (`createTable`); imported tables leave it None because the importer
/// already bakes theme-accent styling into each cell, so this never
/// double-styles or disturbs the parity corpus. Uses orElse/|| merges so a
/// cell's own explicit colors still win. Empty cells in the region are created
/// styled so the band/header spans the full width.
let applyTableStyles (model: SpreadsheetModel): SpreadsheetModel =
    let styleTable (m: SpreadsheetModel) (t: Table) : SpreadsheetModel =
        match t.Style with
        | Some s when s.HeaderFill.IsSome ->
            let r0, c0, r1, c1 = t.StartRow - 1, t.StartCol - 1, t.EndRow - 1, t.EndCol - 1
            let mutable acc = m
            for r in r0 .. r1 do
                let isHeader = r < r0 + t.HeaderRowCount
                let dataIdx = r - r0 - t.HeaderRowCount
                for c in c0 .. c1 do
                    let addr = { Row = r; Col = c }
                    let baseData = getCell addr acc |> Option.defaultValue { Value = Empty; Format = defaultFormat; Comment = None }
                    let fmt = baseData.Format
                    let hasFilter =
                        match t.AutoFilter with Some (fc1, fc2) -> (c + 1) >= fc1 && (c + 1) <= fc2 | None -> false
                    let newFmt =
                        if isHeader then
                            { fmt with
                                BgColor = fmt.BgColor |> Option.orElse s.HeaderFill
                                TextColor = fmt.TextColor |> Option.orElse (Some "#FFFFFF")
                                Bold = true
                                FilterButton = fmt.FilterButton || hasFilter }
                        elif s.ShowRowStripes && dataIdx >= 0 && dataIdx % 2 = 0 then
                            { fmt with BgColor = fmt.BgColor |> Option.orElse s.StripeFill }
                        else fmt
                    if newFmt <> fmt then acc <- setCell addr { baseData with Format = newFmt } acc
            acc
        | _ -> m
    model.Tables |> List.fold styleTable model

// --- Row / column insertion and deletion ------------------------------------
// One implementation serves all four operations: everything a sheet stores by
// position (cells, merges, sizes, conditional formats, validations, tables,
// the spill map, filters) is moved the same way, and formulas are re-pointed.

let private structureEditParts (edit: StructureEdit) =
    match edit with
    | InsertRows (at, count) -> true, true, at, max 0 count
    | DeleteRows (at, count) -> true, false, at, max 0 count
    | InsertCols (at, count) -> false, true, at, max 0 count
    | DeleteCols (at, count) -> false, false, at, max 0 count

/// Where an index on the edited axis ends up; None when it was deleted.
let private moveEditIndex isInsert at count index =
    if isInsert then Some (if index >= at then index + count else index)
    elif index >= at + count then Some (index - count)
    elif index >= at then None
    else Some index

/// Where an inclusive span on the edited axis ends up: an insertion inside it
/// grows it, a deletion overlapping it shrinks it, None when nothing is left.
let private moveEditSpan isInsert at count lo hi =
    if isInsert then
        Some ((if lo >= at then lo + count else lo), (if hi >= at then hi + count else hi))
    elif lo >= at && hi < at + count then
        None
    else
        let newLo = if lo >= at + count then lo - count elif lo >= at then at else lo
        let newHi = if hi >= at + count then hi - count elif hi >= at then at - 1 else hi
        Some (newLo, newHi)

/// Re-point the formulas of a sheet after `edit` happened on the sheet that
/// `targetsEditedSheet` selects. Used for the *other* sheets of a workbook;
/// returns the same model when no formula changed.
let rewriteFormulasForEdit (edit: StructureEdit) (targetsEditedSheet: string option -> bool) (model: SpreadsheetModel) : SpreadsheetModel =
    let mutable cells = model.Cells
    let mutable changed = false
    for KeyValue(addr, cellData) in model.Cells do
        match cellData.Value with
        | Formula (formula, cached) ->
            let adjusted = adjustFormulaForEdit edit targetsEditedSheet formula
            if adjusted <> formula then
                cells <- Map.add addr { cellData with Value = Formula (adjusted, cached) } cells
                changed <- true
        | _ -> ()
    if changed then { model with Cells = cells } |> invalidateFilterProjection else model

/// Insert or delete rows/columns on a sheet. `targetsEditedSheet` tells which
/// sheet qualifiers in this sheet's own formulas refer back to it (always the
/// unqualified ones; a workbook also passes the sheet's name).
let applyStructureEditWith (targetsEditedSheet: string option -> bool) (edit: StructureEdit) (model: SpreadsheetModel) : SpreadsheetModel =
    let onRows, isInsert, at, count = structureEditParts edit
    if count = 0 || at < 0 then model else
    let moveIndex = moveEditIndex isInsert at count
    let moveSpan = moveEditSpan isInsert at count
    let moveAddr (addr: CellAddress) =
        if onRows then moveIndex addr.Row |> Option.map (fun row -> { addr with Row = row })
        else moveIndex addr.Col |> Option.map (fun col -> { addr with Col = col })
    let moveAddressMap (items: Map<CellAddress, 'T>) : Map<CellAddress, 'T> =
        items
        |> Map.fold (fun acc addr value ->
            match moveAddr addr with
            | Some moved -> Map.add moved value acc
            | None -> acc) Map.empty
    let moveIndexMap (items: Map<int, 'T>) : Map<int, 'T> =
        items
        |> Map.fold (fun acc index value ->
            match moveIndex index with
            | Some moved -> Map.add moved value acc
            | None -> acc) Map.empty

    let cells =
        model.Cells
        |> Map.fold (fun acc addr cellData ->
            match moveAddr addr with
            | None -> acc
            | Some moved ->
                let cellData =
                    match cellData.Value with
                    | Formula (formula, cached) ->
                        let adjusted = adjustFormulaForEdit edit targetsEditedSheet formula
                        if adjusted = formula then cellData
                        else { cellData with Value = Formula (adjusted, cached) }
                    | _ -> cellData
                Map.add moved cellData acc) Map.empty

    // A deleted selection falls back onto the row/column that took its place.
    let moveSelected index =
        match moveIndex index with
        | Some moved -> moved
        | None -> max 0 (at - 1)
    let selection =
        match model.Selection with
        | SingleCell (row, col) ->
            if onRows then SingleCell (moveSelected row, col) else SingleCell (row, moveSelected col)
        | Range (sr, sc, er, ec) ->
            if onRows then Range (moveSelected sr, sc, moveSelected er, ec)
            else Range (sr, moveSelected sc, er, moveSelected ec)

    let merges =
        model.MergedCells
        |> List.choose (fun merge ->
            if onRows then
                moveSpan merge.StartRow merge.EndRow
                |> Option.map (fun (lo, hi) -> { merge with StartRow = lo; EndRow = hi })
            else
                moveSpan merge.StartCol merge.EndCol
                |> Option.map (fun (lo, hi) -> { merge with StartCol = lo; EndCol = hi }))
        // A merge squeezed down to one cell is no merge at all.
        |> List.filter (fun merge -> merge.StartRow <> merge.EndRow || merge.StartCol <> merge.EndCol)

    // Table coordinates are 1-indexed; spans are computed 0-indexed.
    let tables =
        model.Tables
        |> List.choose (fun table ->
            if onRows then
                moveSpan (table.StartRow - 1) (table.EndRow - 1)
                |> Option.map (fun (lo, hi) -> { table with StartRow = lo + 1; EndRow = hi + 1 })
            else
                moveSpan (table.StartCol - 1) (table.EndCol - 1)
                |> Option.map (fun (lo, hi) ->
                    let autoFilter =
                        table.AutoFilter
                        |> Option.bind (fun (first, last) -> moveSpan (first - 1) (last - 1))
                        |> Option.map (fun (first, last) -> first + 1, last + 1)
                    { table with StartCol = lo + 1; EndCol = hi + 1; AutoFilter = autoFilter }))

    let spill =
        model.Spill
        |> Map.fold (fun acc ghost anchor ->
            match moveAddr ghost, moveAddr anchor with
            | Some movedGhost, Some movedAnchor -> Map.add movedGhost movedAnchor acc
            | _ -> acc) Map.empty

    let delta = if isInsert then count else -count
    { model with
        Cells = cells
        Selection = selection
        RowCount = if onRows then max 1 (model.RowCount + delta) else model.RowCount
        ColCount = if onRows then model.ColCount else max 1 (model.ColCount + delta)
        RowHeights = if onRows then moveIndexMap model.RowHeights else model.RowHeights
        ColumnWidths = if onRows then model.ColumnWidths else moveIndexMap model.ColumnWidths
        ConditionalFormats = moveAddressMap model.ConditionalFormats
        RichRuns = moveAddressMap model.RichRuns
        MergedCells = merges
        Tables = tables
        CfBoldCells = model.CfBoldCells |> Seq.choose moveAddr |> Set.ofSeq
        Spill = spill
        Validations = moveAddressMap model.Validations
        Filters =
            if onRows then model.Filters
            else
                model.Filters
                |> Map.fold (fun acc col filter ->
                    match moveIndex col with
                    | Some moved -> Map.add moved { filter with Column = moved } acc
                    | None -> acc) Map.empty
        SortState =
            if onRows then model.SortState
            else
                model.SortState
                |> Option.bind (fun sort -> moveIndex sort.Column |> Option.map (fun col -> { sort with Column = col }))
        SearchResults = model.SearchResults |> Array.choose moveAddr
        SearchIndex = 0
    }
    |> refreshFilterProjection

/// Insert or delete rows/columns on a stand-alone sheet. Unqualified
/// references in its formulas follow the edit; a workbook host should call
/// `Workbook.applyStructureEdit`, which also fixes the other sheets, defined
/// names and structured-reference tables.
let applyStructureEdit (edit: StructureEdit) (model: SpreadsheetModel) : SpreadsheetModel =
    applyStructureEditWith Option.isNone edit model

/// Insert a row at the specified index (shifts rows down)
let insertRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (InsertRows (atRow, 1)) model

/// Delete a row at the specified index (shifts rows up)
let deleteRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (DeleteRows (atRow, 1)) model

/// Insert a column at the specified index (shifts columns right)
let insertColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (InsertCols (atCol, 1)) model

/// Delete a column at the specified index (shifts columns left)
let deleteColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    applyStructureEdit (DeleteCols (atCol, 1)) model

/// Get row height (with default)
let getRowHeight (row: int) (model: SpreadsheetModel): float =
    if isFilterHiddenRow row model then
        0.0
    else
        Map.tryFind row model.RowHeights
        |> Option.defaultValue 25.0

/// Set row height
let setRowHeight (row: int) (height: float) (model: SpreadsheetModel): SpreadsheetModel =
    if height = 25.0 then
        if Map.containsKey row model.RowHeights then { model with RowHeights = Map.remove row model.RowHeights } else model
    else
        match Map.tryFind row model.RowHeights with
        | Some existing when existing = height -> model
        | _ -> { model with RowHeights = Map.add row height model.RowHeights }

/// Get column width (with default)
let getColumnWidth (col: int) (model: SpreadsheetModel): float =
    Map.tryFind col model.ColumnWidths
    |> Option.defaultValue 100.0

/// Set column width
let setColumnWidth (col: int) (width: float) (model: SpreadsheetModel): SpreadsheetModel =
    if width = 100.0 then
        if Map.containsKey col model.ColumnWidths then { model with ColumnWidths = Map.remove col model.ColumnWidths } else model
    else
        match Map.tryFind col model.ColumnWidths with
        | Some existing when existing = width -> model
        | _ -> { model with ColumnWidths = Map.add col width model.ColumnWidths }

let usedBounds (model: SpreadsheetModel) : (int * int) option =
    let mutable any = false
    let mutable maxRow = 0
    let mutable maxCol = 0
    for KeyValue(addr, _) in model.Cells do
        any <- true
        if addr.Row > maxRow then maxRow <- addr.Row
        if addr.Col > maxCol then maxCol <- addr.Col
    for merge in model.MergedCells do
        any <- true
        if merge.EndRow > maxRow then maxRow <- merge.EndRow
        if merge.EndCol > maxCol then maxCol <- merge.EndCol
    if any then Some (maxRow, maxCol) else None

let lastUsedColumnInRow (row: int) (model: SpreadsheetModel) : int option =
    let mutable found = false
    let mutable lastCol = 0
    for KeyValue(addr, _) in model.Cells do
        if addr.Row = row then
            found <- true
            if addr.Col > lastCol then lastCol <- addr.Col
    for merge in model.MergedCells do
        if row >= merge.StartRow && row <= merge.EndRow then
            found <- true
            if merge.EndCol > lastCol then lastCol <- merge.EndCol
    if found then Some lastCol else None

/// Search cells for a query string
let private containsSearchText (query: string) (caseSensitive: bool) (text: string) =
#if FABLE_COMPILER
    if caseSensitive then text.Contains(query)
    else text.ToLowerInvariant().Contains(query.ToLowerInvariant())
#else
    let comparison = if caseSensitive then StringComparison.Ordinal else StringComparison.OrdinalIgnoreCase
    text.IndexOf(query, comparison) >= 0
#endif

let searchCells (query: string) (caseSensitive: bool) (model: SpreadsheetModel): CellAddress array =
    model.Cells
    |> Seq.choose (fun (KeyValue(addr, cellData)) ->
        let cellText =
            match cellData.Value with
            | Formula (formula, _) -> formula  // Search in formula text
            | other -> displayValue other

        if containsSearchText query caseSensitive cellText then Some addr else None)
    |> Seq.sortBy (fun addr -> (addr.Row, addr.Col))  // Sort by row then column
    |> Seq.toArray

/// Initialize search with a query
let startSearch (query: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let results = searchCells query caseSensitive model
    { model with
        SearchQuery = Some query
        SearchCaseSensitive = Some caseSensitive
        SearchResults = results
        SearchIndex = 0
    }

let nextSearchResultValue (model: SpreadsheetModel): SpreadsheetModel voption =
    if model.SearchResults.Length = 0 then
        ValueNone
    else
        let newIndex = (model.SearchIndex + 1) % model.SearchResults.Length
        let nextAddr = model.SearchResults.[newIndex]
        let updatedModel =
            { model with
                SearchIndex = newIndex
                Selection = SingleCell (nextAddr.Row, nextAddr.Col)
            }
        ValueSome updatedModel

/// Move to next search result
let nextSearchResult (model: SpreadsheetModel): SpreadsheetModel option =
    match nextSearchResultValue model with
    | ValueSome next -> Some next
    | ValueNone -> None

/// Replace text at current selection
let replaceAtSelection (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel option =
    match model.Selection with
    | SingleCell (row, col) ->
        let addr = { Row = row; Col = col }
        match getCell addr model with
        | Some cellData ->
            let currentText = displayValue cellData.Value
            let textToCheck = if caseSensitive then currentText else currentText.ToLower()
            let findToCheck = if caseSensitive then findText else findText.ToLower()

            if textToCheck.Contains(findToCheck) then
                // Perform replacement
                let newText =
                    if caseSensitive then
                        currentText.Replace(findText, replaceText)
                    else
                        // Case-insensitive replacement is more complex - use simple approach
                        let regex = System.Text.RegularExpressions.Regex(findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        regex.Replace(currentText, replaceText)

                let newValue = String newText
                let newCellData = { cellData with Value = newValue }
                Some (setCell addr newCellData model)
            else
                None
        | None -> None
    | _ -> None

/// Replace all occurrences
let rec replaceAll (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let updated, _ = replaceAllWithCount findText replaceText caseSensitive model
    updated

and replaceAllWithCount (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int =
    let updated, count, _, _ = replaceAllWithChanges findText replaceText caseSensitive model
    updated, count

and replaceAllWithChanges (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel * int * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let oldCells = ResizeArray<CellAddress * CellData option>()
    let newCells = ResizeArray<CellAddress * CellData option>()
    if String.IsNullOrEmpty findText then
        model, 0, oldCells, newCells
    else
        let replacement = replaceText.Replace("$", "$$")
        let regex =
            if caseSensitive then None
            else Some (System.Text.RegularExpressions.Regex(System.Text.RegularExpressions.Regex.Escape findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
        let replaceIn (text: string) =
            if caseSensitive then text.Replace(findText, replaceText)
            else regex.Value.Replace(text, replacement)

        let mutable cells = model.Cells
        let mutable count = 0
        for KeyValue(addr, cellData) in model.Cells do
            let currentText = displayValue cellData.Value
            if containsSearchText findText caseSensitive currentText then
                let newText = replaceIn currentText
                if newText <> currentText then
                    let nextCell = { cellData with Value = String newText }
                    oldCells.Add(addr, Some cellData)
                    newCells.Add(addr, Some nextCell)
                    cells <- Map.add addr nextCell cells
                    count <- count + 1
        if count = 0 then model, 0, oldCells, newCells
        else ({ model with Cells = cells } |> invalidateFilterProjection), count, oldCells, newCells

/// Evaluate if a cell value matches a conditional format rule
let evaluateRule (rule: ConditionalFormatRule) (value: CellValue): bool =
    match rule, value with
    | GreaterThan threshold, Number n -> n > threshold
    | LessThan threshold, Number n -> n < threshold
    | Between (min, max), Number n -> n >= min && n <= max
    | Equal target, Number n -> n = target
    | ContainsText text, String s -> s.Contains(text)
    | ContainsText text, Formula (formula, _) -> formula.Contains(text)
    | _ -> false

/// Get conditional format to apply to a cell (returns first matching rule)
let getConditionalFormat (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel): CellFormat option =
    match Map.tryFind addr model.ConditionalFormats with
    | Some formats ->
        formats
        |> List.tryFind (fun cf -> evaluateRule cf.Rule value)
        |> Option.map (fun cf -> cf.Style)
    | None -> None

/// Add a conditional format rule to a cell
let addConditionalFormat (addr: CellAddress) (format: ConditionalFormat) (model: SpreadsheetModel): SpreadsheetModel =
    let existingFormats = Map.tryFind addr model.ConditionalFormats |> Option.defaultValue []
    let newFormats = format :: existingFormats
    { model with ConditionalFormats = Map.add addr newFormats model.ConditionalFormats }

/// Remove all conditional formats from a cell
let removeConditionalFormats (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    { model with ConditionalFormats = Map.remove addr model.ConditionalFormats }

// --- Conditional-format VISUALS: data bars / color scales / icon sets -------
// These compute over a range and bake the result into each cell's CellFormat
// (DataBar / IconSet / BgColor), which the renderers already paint. Applying
// them is a range operation (from the CF dialog or programmatically).

let rec private unwrapVal (v: CellValue) : CellValue =
    match v with Formula (_, cached) -> unwrapVal cached | x -> x

let private cellNumber (model: SpreadsheetModel) (addr: CellAddress) : float option =
    match getCellValue addr model |> unwrapVal with
    | Number n -> Some n
    | Boolean b -> Some(if b then 1.0 else 0.0)
    | _ -> None

let private rangeNumbers sr sc er ec (model: SpreadsheetModel) : (CellAddress * float) list =
    [ for r in sr .. er do
        for c in sc .. ec do
            match cellNumber model { Row = r; Col = c } with
            | Some n -> yield { Row = r; Col = c }, n
            | None -> () ]

let private updateFormat (addr: CellAddress) (f: CellFormat -> CellFormat) (model: SpreadsheetModel) : SpreadsheetModel =
    match getCell addr model with
    | Some cd -> setCell addr { cd with Format = f cd.Format } model
    | None -> setCell addr { Value = Empty; Format = f defaultFormat; Comment = None } model

let private hexToRgb (s: string) =
    let h = (if s.StartsWith "#" then s.Substring 1 else s)
    if h.Length = 6 then
        try (Convert.ToInt32(h.Substring(0, 2), 16), Convert.ToInt32(h.Substring(2, 2), 16), Convert.ToInt32(h.Substring(4, 2), 16))
        with _ -> (255, 255, 255)
    else (255, 255, 255)

let private lerpHex (a: string) (b: string) (t: float) : string =
    let (ar, ag, ab), (br, bg, bb) = hexToRgb a, hexToRgb b
    let m x y = int (float x + (float y - float x) * t)
    sprintf "#%02X%02X%02X" (m ar br) (m ag bg) (m ab bb)

/// Data bars over a range, baselined at 0 when all values are ≥ 0 (Excel's
/// automatic behaviour), else min→max.
let applyDataBars (sr: int) (sc: int) (er: int) (ec: int) (color: string) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells ->
        let vals = cells |> List.map snd
        let lo = min 0.0 (List.min vals)
        let hi = max 0.0 (List.max vals)
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            updateFormat addr (fun f -> { f with DataBar = Some(max 0.0 (min 1.0 ((v - lo) / span)), color) }) m) model

/// 2- or 3-color scale (pass `mid = None` for 2-color) → per-cell BgColor.
let applyColorScale (sr: int) (sc: int) (er: int) (ec: int) (loColor: string) (mid: string option) (hiColor: string) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells ->
        let vals = cells |> List.map snd
        let lo, hi = List.min vals, List.max vals
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            let t = (v - lo) / span
            let color =
                match mid with
                | None -> lerpHex loColor hiColor t
                | Some mc -> if t <= 0.5 then lerpHex loColor mc (t * 2.0) else lerpHex mc hiColor ((t - 0.5) * 2.0)
            updateFormat addr (fun f -> { f with BgColor = Some color }) m) model

/// Icon set over a range: `count` buckets by value (Excel's percent
/// thresholds), each cell tagged with its bucket index for the renderer.
let applyIconSet (sr: int) (sc: int) (er: int) (ec: int) (count: int) (model: SpreadsheetModel) : SpreadsheetModel =
    match rangeNumbers sr sc er ec model with
    | [] -> model
    | cells when count > 1 ->
        let vals = cells |> List.map snd
        let lo, hi = List.min vals, List.max vals
        let span = if hi - lo < 1e-12 then 1.0 else hi - lo
        cells
        |> List.fold (fun m (addr, v) ->
            let bucket = min (count - 1) (int ((v - lo) / span * float count))
            updateFormat addr (fun f -> { f with IconSet = Some(bucket, count) }) m) model
    | _ -> model

// --- Data validation --------------------------------------------------------

let private displayText (v: CellValue) : string =
    match unwrapVal v with
    | String s -> s
    | Number n -> if n = floor n && abs n < 1e15 then string (int64 n) else string n
    | Boolean b -> if b then "TRUE" else "FALSE"
    | _ -> ""

/// Check a value against a validation rule. Blank is always allowed (Excel's
/// default). Returns Ok or a human-readable reason.
let validateValue (rule: ValidationRule) (value: CellValue) : Result<unit, string> =
    let v = unwrapVal value
    let isBlank = match v with Empty -> true | String "" -> true | _ -> false
    let within lo hi x =
        (match lo with Some l -> x >= l - 1e-9 | None -> true) && (match hi with Some h -> x <= h + 1e-9 | None -> true)
    if isBlank then Result.Ok()
    else
        match rule with
        | ListValues vs ->
            let t = (displayText v).Trim().ToLowerInvariant()
            if vs |> List.exists (fun a -> a.Trim().ToLowerInvariant() = t) then Result.Ok()
            else Result.Error(sprintf "Value must be one of: %s" (String.concat ", " vs))
        | WholeNumber (lo, hi) ->
            (match v with
             | Number n when n = floor n -> if within lo hi n then Result.Ok() else Result.Error "Enter a whole number within the allowed range."
             | Number _ -> Result.Error "Enter a whole number (no decimals)."
             | _ -> Result.Error "Enter a whole number.")
        | DecimalRange (lo, hi) ->
            (match v with
             | Number n -> if within lo hi n then Result.Ok() else Result.Error "Enter a number within the allowed range."
             | _ -> Result.Error "Enter a number.")
        | TextLength (lo, hi) ->
            let len = match v with String s -> float s.Length | _ -> float (displayText v).Length
            if within lo hi len then Result.Ok() else Result.Error "Text length is outside the allowed range."
        | DateRange (lo, hi) ->
            (match v with
             | Number n -> if within lo hi n then Result.Ok() else Result.Error "Enter a date within the allowed range."
             | _ -> Result.Error "Enter a date.")

let getValidation (addr: CellAddress) (model: SpreadsheetModel) : DataValidation option =
    Map.tryFind addr model.Validations

/// Attach a validation rule to a cell. A list rule also flags the cell so the
/// renderer paints the dropdown arrow.
let setValidation (addr: CellAddress) (dv: DataValidation) (model: SpreadsheetModel) : SpreadsheetModel =
    let m = { model with Validations = Map.add addr dv model.Validations }
    match dv.Rule with
    | ListValues _ -> updateFormat addr (fun f -> { f with HasListDropdown = true }) m
    | _ -> m

/// Validate a prospective edit against the cell's rule (if any). Ok when valid
/// or unconstrained; Error carries the rule's message (or a default).
let validateEdit (addr: CellAddress) (value: CellValue) (model: SpreadsheetModel) : Result<unit, string> =
    match getValidation addr model with
    | None -> Result.Ok()
    | Some dv ->
        match validateValue dv.Rule value with
        | Result.Ok() -> Result.Ok()
        | Result.Error def -> Result.Error(dv.ErrorMessage |> Option.defaultValue def)

/// Parse CSV string to list of rows
let parseCSV (csvText: string): string list list =
    let rows = ResizeArray<string list>()
    let mutable row = ResizeArray<string>()
    let cell = System.Text.StringBuilder()
    let mutable inQuotes = false
    let mutable i = 0

    let finishCell () =
        row.Add(cell.ToString())
        cell.Clear() |> ignore

    let finishRow () =
        finishCell()
        rows.Add(row |> Seq.toList)
        row <- ResizeArray<string>()

    while i < csvText.Length do
        let c = csvText.[i]
        if inQuotes then
            if c = '"' then
                if i + 1 < csvText.Length && csvText.[i + 1] = '"' then
                    cell.Append('"') |> ignore
                    i <- i + 1
                else
                    inQuotes <- false
            else
                cell.Append(c) |> ignore
        else
            match c with
            | '"' when cell.Length = 0 ->
                inQuotes <- true
            | ',' ->
                finishCell()
            | '\r' ->
                if i + 1 < csvText.Length && csvText.[i + 1] = '\n' then
                    i <- i + 1
                finishRow()
            | '\n' ->
                finishRow()
            | _ ->
                cell.Append(c) |> ignore
        i <- i + 1

    if cell.Length > 0 || row.Count > 0 then
        finishRow()

    rows |> Seq.toList

/// Import CSV data into spreadsheet
let rec importCSV (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    let imported, _, _ = importCSVWithChanges csvText startRow startCol model
    imported

and importCSVWithChanges (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel * ResizeArray<CellAddress * CellData option> * ResizeArray<CellAddress * CellData option> =
    let rows = parseCSV csvText
    let mutable updatedModel = model
    let oldCells = ResizeArray<CellAddress * CellData option>()
    let newCells = ResizeArray<CellAddress * CellData option>()

    let record addr next =
        let old = getCell addr updatedModel
        if old <> next then
            oldCells.Add(addr, old)
            newCells.Add(addr, next)

    rows |> List.iteri (fun rowIdx rowData ->
        rowData |> List.iteri (fun colIdx cellText ->
            let addr = { Row = startRow + rowIdx; Col = startCol + colIdx }

            // Parse cell value
            let value =
                if String.IsNullOrEmpty(cellText) then
                    Empty
                elif cellText.StartsWith("=") then
                    Formula (cellText, Empty)
                elif cellText.ToLower() = "true" || cellText.ToLower() = "false" then
                    Boolean (cellText.ToLower() = "true")
                else
                    // Try to parse as number, otherwise string
                    match tryParseNumber cellText with
                    | Some n -> Number n
                    | None -> String cellText

            if value = Empty then
                record addr None
                updatedModel <- clearCell addr updatedModel
            else
                let cellData = {
                    Value = value
                    Format = defaultFormat
                    Comment = None
                }
                record addr (Some cellData)
                updatedModel <- setCell addr cellData updatedModel
        )
    )

    refreshFilterProjection updatedModel, oldCells, newCells

/// Export spreadsheet to CSV string
let exportCSV (model: SpreadsheetModel): string =
    // Find the range of used cells
    let allAddrs = model.Cells |> Map.toList |> List.map fst

    if List.isEmpty allAddrs then
        ""
    else
        let maxRow = allAddrs |> List.map (fun addr -> addr.Row) |> List.max
        let maxCol = allAddrs |> List.map (fun addr -> addr.Col) |> List.max

        // Build CSV row by row
        let rows =
            [0 .. maxRow]
            |> List.map (fun row ->
                [0 .. maxCol]
                |> List.map (fun col ->
                    let addr = { Row = row; Col = col }
                    match getCell addr model with
                    | Some cellData ->
                        let valueText = displayValueWithFormat cellData.Value cellData.Format
                        // Escape commas and quotes
                        if valueText.Contains(",") || valueText.Contains("\"") || valueText.Contains("\n") || valueText.Contains("\r") then
                            sprintf "\"%s\"" (valueText.Replace("\"", "\"\""))
                        else
                            valueText
                    | None -> ""
                )
                |> String.concat ","
            )

        String.concat "\n" rows

/// Check if a cell address is within a merged range
let isInMergedRange (addr: CellAddress) (merge: MergedCell): bool =
    addr.Row >= merge.StartRow && addr.Row <= merge.EndRow &&
    addr.Col >= merge.StartCol && addr.Col <= merge.EndCol

/// Find merged cell containing this address (if any)
let findMergedCell (addr: CellAddress) (model: SpreadsheetModel): MergedCell option =
    model.MergedCells
    |> List.tryFind (fun merge -> isInMergedRange addr merge)

/// Get primary cell address (top-left) of a merge
let getPrimaryCell (merge: MergedCell): CellAddress =
    { Row = merge.StartRow; Col = merge.StartCol }

/// Check if a cell is the primary cell of a merge
let isPrimaryCell (addr: CellAddress) (merge: MergedCell): bool =
    addr.Row = merge.StartRow && addr.Col = merge.StartCol

/// Add a merged cell range
let addMergedCell (startRow: int) (startCol: int) (endRow: int) (endCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    if startRow > endRow || startCol > endCol then
        model  // Invalid range
    else
        let merge = {
            StartRow = startRow
            StartCol = startCol
            EndRow = endRow
            EndCol = endCol
        }
        { model with MergedCells = merge :: model.MergedCells }

/// Remove merged cell containing this address
let removeMergedCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    let newMerges =
        model.MergedCells
        |> List.filter (fun merge -> not (isInMergedRange addr merge))
    { model with MergedCells = newMerges }

let private mergeIntersectsRange startRow endRowExclusive startCol endColExclusive (merge: MergedCell) =
    merge.EndRow >= startRow &&
    merge.StartRow < endRowExclusive &&
    merge.EndCol >= startCol &&
    merge.StartCol < endColExclusive

let mergedCellsIntersectingRange startRow endRowExclusive startCol endColExclusive (model: SpreadsheetModel) =
    model.MergedCells
    |> List.filter (mergeIntersectsRange startRow endRowExclusive startCol endColExclusive)
