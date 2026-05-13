module RedBat.Spreadsheet.Model

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
    FontName: string option
    TextColor: string option
    BgColor: string option
    FontSize: int option
    HorizontalAlign: string option  // "left", "center", "right", "justify"
    VerticalAlign: string option    // "top", "middle", "bottom"
    WrapText: bool
    NumberFormat: string option
    BorderTop: BorderStyle option
    BorderBottom: BorderStyle option
    BorderLeft: BorderStyle option
    BorderRight: BorderStyle option
    /// Excel auto-filter dropdown button on the right edge of the cell.
    /// Set for cells in the header row of a table with an <autoFilter> child.
    FilterButton: bool
}

/// Default format
let defaultFormat = {
    Bold = false
    Italic = false
    Underline = false
    FontName = None
    TextColor = None
    BgColor = None
    FontSize = None
    HorizontalAlign = None
    VerticalAlign = None
    WrapText = false
    NumberFormat = None
    BorderTop = None
    BorderBottom = None
    BorderLeft = None
    BorderRight = None
    FilterButton = false
}

/// Cell comment/note
type CellComment = {
    Text: string
    Author: string option
}

/// Per-run formatting override for rich-text cells. Each field is optional;
/// when None the run inherits from the cell's base CellFormat. Originates
/// from `<rPr>` children of `<r>` elements inside `<si>` shared strings.
type RunOverride = {
    Bold: bool option
    Italic: bool option
    Underline: bool option
    Color: string option       // hex with leading '#'
    FontName: string option
    FontSize: float option     // points
}

let emptyRunOverride = {
    Bold = None
    Italic = None
    Underline = None
    Color = None
    FontName = None
    FontSize = None
}

/// Cell data
type CellData = {
    Value: CellValue
    Format: CellFormat
    Comment: CellComment option
}

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

/// Merged cell range (primary cell is top-left)
type MergedCell = {
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
}

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

    /// Custom row heights (default 25.0)
    RowHeights: Map<int, float>

    /// Custom column widths (default 100.0)
    ColumnWidths: Map<int, float>

    /// Current search query
    SearchQuery: string option

    /// Search results (cell addresses)
    SearchResults: CellAddress list

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
    RowHeights = Map.empty
    ColumnWidths = Map.empty
    SearchQuery = None
    SearchResults = []
    SearchIndex = 0
    ConditionalFormats = Map.empty
    MergedCells = []
    RichRuns = Map.empty
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

/// Set cell data at address
let setCell (addr: CellAddress) (data: CellData) (model: SpreadsheetModel): SpreadsheetModel =
    { model with Cells = Map.add addr data model.Cells }

/// Clear cell at address
let clearCell (addr: CellAddress) (model: SpreadsheetModel): SpreadsheetModel =
    { model with Cells = Map.remove addr model.Cells }

/// Format a number according to format code
let formatNumber (format: string option) (value: float): string =
    match format with
    | Some "currency" ->
        // Currency format: $1,234.56
        sprintf "$%s" (value.ToString("N2"))
    | Some "percentage" ->
        // Percentage format: 45.67%
        sprintf "%s%%" ((value * 100.0).ToString("N2"))
    | Some "date" ->
        // Excel date format (days since 1900-01-01)
        try
            let baseDate = System.DateTime(1900, 1, 1)
            let date = baseDate.AddDays(value - 1.0)  // Excel dates are 1-indexed
            date.ToString("MM/dd/yyyy")
        with
        | _ -> value.ToString("F2")
    | Some "decimal2" ->
        // 2 decimal places: 123.46
        value.ToString("F2")
    | Some "number" ->
        // General number with comma separators
        value.ToString("N2")
    | _ ->
        // Default: just show the number
        string value

/// Display value for cell (handles formulas) - needs to be before getCellValue
let rec displayValue (value: CellValue): string =
    match value with
    | Empty -> ""
    | Number n -> n.ToString("G", Globalization.CultureInfo.InvariantCulture)
    | String s -> s
    | Boolean b -> if b then "TRUE" else "FALSE"
    | Error e -> sprintf "#%s" e
    | Formula (_, cached) -> displayValue cached

/// Display value with formatting applied
let displayValueWithFormat (value: CellValue) (format: CellFormat): string =
    match value with
    | Number n -> formatNumber format.NumberFormat n
    | _ -> displayValue value

/// Get cell value (returns Empty if not found)
let getCellValue (addr: CellAddress) (model: SpreadsheetModel): CellValue =
    match getCell addr model with
    | Some cell -> cell.Value
    | None -> Empty

/// Adjust formula references when pasting (relative reference adjustment)
let rec adjustFormulaReferences (formula: string) (deltaRow: int) (deltaCol: int): string =
    try
        // Simple adjustment: parse the formula and adjust cell references
        // For now, we'll use a simple string replacement approach
        // TODO: Implement proper AST-based formula adjustment for absolute refs ($A$1)
        let mutable adjusted = formula

        // Find all cell references in the formula (simple regex-like approach)
        let mutable i = 0
        let mutable result = System.Text.StringBuilder()

        while i < formula.Length do
            let c = formula.[i]
            if Char.IsLetter(c) && (i = 0 || not (Char.IsLetterOrDigit(formula.[i-1]))) then
                // Potential cell reference
                let mutable colPart = string c
                let mutable j = i + 1

                // Collect column letters
                while j < formula.Length && Char.IsLetter(formula.[j]) do
                    colPart <- colPart + string formula.[j]
                    j <- j + 1

                // Collect row digits
                let mutable rowPart = ""
                while j < formula.Length && Char.IsDigit(formula.[j]) do
                    rowPart <- rowPart + string formula.[j]
                    j <- j + 1

                // Check if this is a valid cell reference
                if not (String.IsNullOrEmpty(rowPart)) then
                    // Valid cell reference - adjust it
                    let col = letterToColumn colPart
                    let row = Int32.Parse(rowPart) - 1

                    let newRow = row + deltaRow
                    let newCol = col + deltaCol

                    if newRow >= 0 && newCol >= 0 then
                        let newRef = formatCellRef newRow newCol
                        result.Append(newRef) |> ignore
                    else
                        // Reference out of bounds, keep original
                        result.Append(colPart + rowPart) |> ignore

                    i <- j
                else
                    // Not a cell reference, just a word
                    result.Append(c) |> ignore
                    i <- i + 1
            else
                result.Append(c) |> ignore
                i <- i + 1

        result.ToString()
    with
    | _ -> formula  // On error, return original formula

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

            // Adjust formula references if this is a formula
            let adjustedValue =
                match cellData.Value with
                | Formula (formula, _) ->
                    let adjustedFormula = adjustFormulaReferences formula deltaRow deltaCol
                    // Will be re-evaluated when set
                    Formula (adjustedFormula, Empty)
                | other -> other

            let adjustedCellData = { cellData with Value = adjustedValue }
            updatedModel <- setCell pasteAddr adjustedCellData updatedModel

    updatedModel

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

/// Sort spreadsheet by column
let sortByColumn (col: int) (ascending: bool) (model: SpreadsheetModel): SpreadsheetModel =
    // Collect all rows that have data in any column
    let allRows =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, _) -> addr.Row)
        |> List.distinct
        |> List.sort

    if List.isEmpty allRows then
        // No data to sort
        { model with SortState = Some { Column = col; Ascending = ascending } }
    else
        // Get the value for the sort column for each row
        let rowsWithValues =
            allRows
            |> List.map (fun row ->
                let addr = { Row = row; Col = col }
                let value = getCellValue addr model
                (row, value)
            )

        // Sort rows by their values in the sort column
        let sortedRows =
            if ascending then
                rowsWithValues |> List.sortWith (fun (_, v1) (_, v2) -> compareCellValues v1 v2)
            else
                rowsWithValues |> List.sortWith (fun (_, v1) (_, v2) -> compareCellValues v2 v1)
            |> List.map fst

        // Create mapping from old row index to new row index
        let rowMapping =
            sortedRows
            |> List.mapi (fun newIdx oldRow -> (oldRow, newIdx))
            |> Map.ofList

        // Rebuild cells with new row indices
        let newCells =
            model.Cells
            |> Map.toList
            |> List.map (fun (addr, cellData) ->
                match Map.tryFind addr.Row rowMapping with
                | Some newRow ->
                    let newAddr = { addr with Row = newRow }
                    (newAddr, cellData)
                | None ->
                    // Row not in sorted set, keep as is
                    (addr, cellData)
            )
            |> Map.ofList

        { model with
            Cells = newCells
            SortState = Some { Column = col; Ascending = ascending }
        }

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

/// Remove filter from a column
let removeFilter (col: int) (model: SpreadsheetModel): SpreadsheetModel =
    { model with Filters = Map.remove col model.Filters }

/// Get visible rows based on all active filters
let getVisibleRows (model: SpreadsheetModel): Set<int> =
    if Map.isEmpty model.Filters then
        // No filters, all rows visible
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, _) -> addr.Row)
        |> Set.ofList
    else
        // Get all rows
        let allRows =
            model.Cells
            |> Map.toList
            |> List.map (fun (addr, _) -> addr.Row)
            |> Set.ofList

        // Filter rows based on each column filter (AND logic)
        allRows
        |> Set.filter (fun row ->
            // Row is visible if it passes ALL filters
            model.Filters
            |> Map.forall (fun col filter ->
                let addr = { Row = row; Col = col }
                let value = getCellValue addr model
                let displayStr = displayValue value
                // Row passes this filter if value is in visible set
                Set.contains displayStr filter.VisibleValues
            )
        )

/// Insert a row at the specified index (shifts rows down)
let insertRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    // Shift all rows >= atRow down by 1
    let newCells =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, cellData) ->
            if addr.Row >= atRow then
                let newAddr = { addr with Row = addr.Row + 1 }
                (newAddr, cellData)
            else
                (addr, cellData)
        )
        |> Map.ofList

    // Update selection if affected
    let newSelection =
        match model.Selection with
        | SingleCell (row, col) ->
            if row >= atRow then SingleCell (row + 1, col)
            else model.Selection
        | Range (sr, sc, er, ec) ->
            let newSr = if sr >= atRow then sr + 1 else sr
            let newEr = if er >= atRow then er + 1 else er
            Range (newSr, sc, newEr, ec)

    { model with
        Cells = newCells
        Selection = newSelection
        RowCount = model.RowCount + 1
    }

/// Delete a row at the specified index (shifts rows up)
let deleteRow (atRow: int) (model: SpreadsheetModel): SpreadsheetModel =
    // Remove cells in the deleted row and shift rows down
    let newCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cellData) ->
            if addr.Row = atRow then
                None  // Remove this row
            elif addr.Row > atRow then
                let newAddr = { addr with Row = addr.Row - 1 }
                Some (newAddr, cellData)
            else
                Some (addr, cellData)
        )
        |> Map.ofList

    // Update selection if affected
    let newSelection =
        match model.Selection with
        | SingleCell (row, col) ->
            if row = atRow then SingleCell (max 0 (atRow - 1), col)
            elif row > atRow then SingleCell (row - 1, col)
            else model.Selection
        | Range (sr, sc, er, ec) ->
            let newSr = if sr > atRow then sr - 1 elif sr = atRow then max 0 (atRow - 1) else sr
            let newEr = if er > atRow then er - 1 elif er = atRow then max 0 (atRow - 1) else er
            Range (newSr, sc, newEr, ec)

    { model with
        Cells = newCells
        Selection = newSelection
        RowCount = max 1 (model.RowCount - 1)
    }

/// Insert a column at the specified index (shifts columns right)
let insertColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    // Shift all columns >= atCol right by 1
    let newCells =
        model.Cells
        |> Map.toList
        |> List.map (fun (addr, cellData) ->
            if addr.Col >= atCol then
                let newAddr = { addr with Col = addr.Col + 1 }
                (newAddr, cellData)
            else
                (addr, cellData)
        )
        |> Map.ofList

    // Update selection if affected
    let newSelection =
        match model.Selection with
        | SingleCell (row, col) ->
            if col >= atCol then SingleCell (row, col + 1)
            else model.Selection
        | Range (sr, sc, er, ec) ->
            let newSc = if sc >= atCol then sc + 1 else sc
            let newEc = if ec >= atCol then ec + 1 else ec
            Range (sr, newSc, er, newEc)

    { model with
        Cells = newCells
        Selection = newSelection
        ColCount = model.ColCount + 1
    }

/// Delete a column at the specified index (shifts columns left)
let deleteColumn (atCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    // Remove cells in the deleted column and shift columns left
    let newCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cellData) ->
            if addr.Col = atCol then
                None  // Remove this column
            elif addr.Col > atCol then
                let newAddr = { addr with Col = addr.Col - 1 }
                Some (newAddr, cellData)
            else
                Some (addr, cellData)
        )
        |> Map.ofList

    // Update selection if affected
    let newSelection =
        match model.Selection with
        | SingleCell (row, col) ->
            if col = atCol then SingleCell (row, max 0 (atCol - 1))
            elif col > atCol then SingleCell (row, col - 1)
            else model.Selection
        | Range (sr, sc, er, ec) ->
            let newSc = if sc > atCol then sc - 1 elif sc = atCol then max 0 (atCol - 1) else sc
            let newEc = if ec > atCol then ec - 1 elif ec = atCol then max 0 (atCol - 1) else ec
            Range (sr, newSc, er, newEc)

    { model with
        Cells = newCells
        Selection = newSelection
        ColCount = max 1 (model.ColCount - 1)
    }

/// Get row height (with default)
let getRowHeight (row: int) (model: SpreadsheetModel): float =
    Map.tryFind row model.RowHeights
    |> Option.defaultValue 25.0

/// Set row height
let setRowHeight (row: int) (height: float) (model: SpreadsheetModel): SpreadsheetModel =
    { model with RowHeights = Map.add row height model.RowHeights }

/// Get column width (with default)
let getColumnWidth (col: int) (model: SpreadsheetModel): float =
    Map.tryFind col model.ColumnWidths
    |> Option.defaultValue 100.0

/// Set column width
let setColumnWidth (col: int) (width: float) (model: SpreadsheetModel): SpreadsheetModel =
    { model with ColumnWidths = Map.add col width model.ColumnWidths }

/// Search cells for a query string
let searchCells (query: string) (caseSensitive: bool) (model: SpreadsheetModel): CellAddress list =
    let queryLower = if caseSensitive then query else query.ToLower()

    model.Cells
    |> Map.toList
    |> List.filter (fun (addr, cellData) ->
        let cellText =
            match cellData.Value with
            | Formula (formula, _) -> formula  // Search in formula text
            | other -> displayValue other

        let textToSearch = if caseSensitive then cellText else cellText.ToLower()
        textToSearch.Contains(queryLower)
    )
    |> List.map fst
    |> List.sortBy (fun addr -> (addr.Row, addr.Col))  // Sort by row then column

/// Initialize search with a query
let startSearch (query: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let results = searchCells query caseSensitive model
    { model with
        SearchQuery = Some query
        SearchResults = results
        SearchIndex = 0
    }

/// Move to next search result
let nextSearchResult (model: SpreadsheetModel): SpreadsheetModel option =
    if List.isEmpty model.SearchResults then
        None
    else
        let newIndex = (model.SearchIndex + 1) % List.length model.SearchResults
        let nextAddr = List.item newIndex model.SearchResults
        let updatedModel =
            { model with
                SearchIndex = newIndex
                Selection = SingleCell (nextAddr.Row, nextAddr.Col)
            }
        Some updatedModel

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
let replaceAll (findText: string) (replaceText: string) (caseSensitive: bool) (model: SpreadsheetModel): SpreadsheetModel =
    let matches = searchCells findText caseSensitive model

    let mutable updatedModel = model
    for addr in matches do
        match getCell addr updatedModel with
        | Some cellData ->
            let currentText = displayValue cellData.Value
            let newText =
                if caseSensitive then
                    currentText.Replace(findText, replaceText)
                else
                    let regex = System.Text.RegularExpressions.Regex(findText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    regex.Replace(currentText, replaceText)

            let newValue = String newText
            let newCellData = { cellData with Value = newValue }
            updatedModel <- setCell addr newCellData updatedModel
        | None -> ()

    updatedModel

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
let importCSV (csvText: string) (startRow: int) (startCol: int) (model: SpreadsheetModel): SpreadsheetModel =
    let rows = parseCSV csvText
    let mutable updatedModel = model

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
                updatedModel <- clearCell addr updatedModel
            else
                let cellData = {
                    Value = value
                    Format = defaultFormat
                    Comment = None
                }
                updatedModel <- setCell addr cellData updatedModel
        )
    )

    updatedModel

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
