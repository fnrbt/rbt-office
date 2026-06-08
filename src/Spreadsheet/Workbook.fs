module RedBat.Spreadsheet.Workbook

open System.Collections.Generic
open RedBat.Spreadsheet.Model

/// Sheet ID (unique identifier)
type SheetId = string

/// A cell addressed absolutely within a workbook — sheet + cell. The unit
/// the cross-sheet dependency graph is keyed on.
type SheetCell = {
    Sheet: SheetId
    Cell: CellAddress
}

/// Sheet in a workbook
type Sheet = {
    Id: SheetId
    Name: string
    Model: SpreadsheetModel
    Visible: bool
}

/// Table definition (structured reference). Coordinates are 0-indexed model
/// rows/cols (the form `getTableColumn`/`expandTableSel` read directly).
/// `HasTotals` marks a totals row occupying `EndRow` (Excel's `totalsRowCount`).
type Table = {
    Name: string
    SheetId: SheetId
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    Columns: string list  // Column names (header order, left → right)
    HasHeaders: bool
    HasTotals: bool
}

/// Workbook containing multiple sheets
type Workbook = {
    Sheets: Map<SheetId, Sheet>
    SheetOrder: SheetId list  // Tab order
    ActiveSheet: SheetId
    Tables: Map<string, Table>  // Named tables
    /// Defined names → their refers-to formula text (e.g. "Sheet1!$A$1:$B$10").
    /// Keyed by (scope, UPPER-cased name): scope is the owning SheetId for
    /// sheet-local names, None for workbook-global names. A sheet-local name
    /// shadows a global one of the same name when evaluating on that sheet.
    Names: Map<SheetId option * string, string>
}

/// Create empty workbook with one sheet
let createWorkbook (): Workbook =
    let sheetId = "sheet1"
    let sheet = {
        Id = sheetId
        Name = "Sheet1"
        Model = createEmpty 1000 26
        Visible = true
    }
    {
        Sheets = Map.ofList [(sheetId, sheet)]
        SheetOrder = [sheetId]
        ActiveSheet = sheetId
        Tables = Map.empty
        Names = Map.empty
    }

/// Look up a defined name's refers-to text, scoped to the current sheet: a
/// sheet-local name shadows a workbook-global one of the same name.
/// Case-insensitive.
let getName (currentSheet: SheetId option) (name: string) (workbook: Workbook): string option =
    let key = name.ToUpperInvariant()
    let local = currentSheet |> Option.bind (fun s -> Map.tryFind (Some s, key) workbook.Names)
    match local with
    | Some _ -> local
    | None -> Map.tryFind (None, key) workbook.Names

/// Build the Names map from raw (scope, name, refersTo) triples. Keys are
/// upper-cased for case-insensitive lookup; later entries win on collision.
let namesOfTriples (triples: (SheetId option * string * string) seq): Map<SheetId option * string, string> =
    triples
    |> Seq.fold (fun acc (scope, n, refersTo) -> Map.add (scope, n.ToUpperInvariant()) refersTo acc) Map.empty

/// Convenience for workbook-global names only (scope = None).
let namesOfPairs (pairs: (string * string) seq): Map<SheetId option * string, string> =
    pairs |> Seq.map (fun (n, refersTo) -> None, n, refersTo) |> namesOfTriples

/// Get active sheet
let getActiveSheet (workbook: Workbook): Sheet option =
    Map.tryFind workbook.ActiveSheet workbook.Sheets

/// Add a new sheet
let addSheet (name: string) (workbook: Workbook): Workbook =
    let sheetId = sprintf "sheet%d" (Map.count workbook.Sheets + 1)
    let sheet = {
        Id = sheetId
        Name = name
        Model = createEmpty 1000 26
        Visible = true
    }
    {
        workbook with
            Sheets = Map.add sheetId sheet workbook.Sheets
            SheetOrder = workbook.SheetOrder @ [sheetId]
    }

/// Delete a sheet
let deleteSheet (sheetId: SheetId) (workbook: Workbook): Workbook option =
    if Map.count workbook.Sheets <= 1 then
        // Can't delete last sheet
        None
    else
        let newSheets = Map.remove sheetId workbook.Sheets
        let newOrder = List.filter ((<>) sheetId) workbook.SheetOrder
        let newActive =
            if workbook.ActiveSheet = sheetId then
                List.head newOrder
            else
                workbook.ActiveSheet
        Some {
            workbook with
                Sheets = newSheets
                SheetOrder = newOrder
                ActiveSheet = newActive
        }

/// Rename a sheet
let renameSheet (sheetId: SheetId) (newName: string) (workbook: Workbook): Workbook =
    match Map.tryFind sheetId workbook.Sheets with
    | Some sheet ->
        let updatedSheet = { sheet with Name = newName }
        { workbook with Sheets = Map.add sheetId updatedSheet workbook.Sheets }
    | None -> workbook

/// Switch active sheet
let setActiveSheet (sheetId: SheetId) (workbook: Workbook): Workbook =
    if Map.containsKey sheetId workbook.Sheets then
        { workbook with ActiveSheet = sheetId }
    else
        workbook

/// Reorder sheets
let reorderSheets (newOrder: SheetId list) (workbook: Workbook): Workbook =
    // Validate all sheet IDs exist
    let valid = List.forall (fun id -> Map.containsKey id workbook.Sheets) newOrder
    if valid && List.length newOrder = Map.count workbook.Sheets then
        { workbook with SheetOrder = newOrder }
    else
        workbook

/// Parse cross-sheet reference: Sheet2!A1 or 'Sheet Name'!B2
let parseSheetReference (ref: string): (string option * string) =
    let parts = ref.Split('!')
    if parts.Length = 2 then
        let sheetPart = parts.[0]
        let cellPart = parts.[1]

        // Remove quotes if present
        let sheetName =
            if sheetPart.StartsWith("'") && sheetPart.EndsWith("'") then
                sheetPart.Substring(1, sheetPart.Length - 2)
            else
                sheetPart

        (Some sheetName, cellPart)
    else
        (None, ref)

/// Resolve sheet name to sheet ID
let resolveSheetName (name: string) (workbook: Workbook): SheetId option =
    workbook.Sheets
    |> Map.tryPick (fun id sheet ->
        if sheet.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase) then
            Some id
        else
            None
    )

/// Define a table (structured reference)
let defineTable (table: Table) (workbook: Workbook): Workbook =
    { workbook with Tables = Map.add table.Name table workbook.Tables }

/// Delete a table
let deleteTable (name: string) (workbook: Workbook): Workbook =
    { workbook with Tables = Map.remove name workbook.Tables }

/// Get table by name (case-insensitive — Excel table names are).
let getTable (name: string) (workbook: Workbook): Table option =
    match Map.tryFind name workbook.Tables with
    | Some t -> Some t
    | None ->
        workbook.Tables
        |> Map.tryPick (fun k t ->
            if System.String.Equals(k, name, System.StringComparison.OrdinalIgnoreCase) then Some t else None)

/// Index of a column within a table by name (case-insensitive). None if absent.
let tableColumnIndex (table: Table) (columnName: string): int option =
    table.Columns
    |> List.tryFindIndex (fun c -> System.String.Equals(c, columnName, System.StringComparison.OrdinalIgnoreCase))

/// The table on `sheet` whose body region contains `cell`, if any — used to
/// resolve unqualified structured refs (`[@Col]`) from the evaluating cell.
let findTableContaining (sheet: SheetId option) (cell: CellAddress) (workbook: Workbook): Table option =
    workbook.Tables
    |> Map.tryPick (fun _ t ->
        let sheetOk = match sheet with Some s -> s = t.SheetId | None -> true
        if sheetOk && cell.Row >= t.StartRow && cell.Row <= t.EndRow
           && cell.Col >= t.StartCol && cell.Col <= t.EndCol then Some t
        else None)

/// Resolve table column reference: Table1[Amount]
let parseTableReference (ref: string): (string * string) option =
    if ref.Contains("[") && ref.Contains("]") then
        let tablePart = ref.Substring(0, ref.IndexOf('['))
        let columnPart = ref.Substring(ref.IndexOf('[') + 1)
        let columnName = columnPart.Replace("]", "")
        Some (tablePart, columnName)
    else
        None

/// Get table column range
let getTableColumn (tableName: string) (columnName: string) (workbook: Workbook): (SheetId * CellAddress * CellAddress) option =
    match Map.tryFind tableName workbook.Tables with
    | Some table ->
        // Find column index
        match List.tryFindIndex ((=) columnName) table.Columns with
        | Some colIndex ->
            let actualCol = table.StartCol + colIndex
            let startAddr = { Row = table.StartRow + (if table.HasHeaders then 1 else 0); Col = actualCol }
            let endAddr = { Row = table.EndRow; Col = actualCol }
            Some (table.SheetId, startAddr, endAddr)
        | None -> None
    | None -> None

/// Update sheet model
let updateSheetModel (sheetId: SheetId) (model: SpreadsheetModel) (workbook: Workbook): Workbook =
    match Map.tryFind sheetId workbook.Sheets with
    | Some sheet ->
        let updatedSheet = { sheet with Model = model }
        { workbook with Sheets = Map.add sheetId updatedSheet workbook.Sheets }
    | None -> workbook

/// Get all sheet names
let getSheetNames (workbook: Workbook): (SheetId * string) list =
    workbook.SheetOrder
    |> List.choose (fun id ->
        Map.tryFind id workbook.Sheets
        |> Option.map (fun sheet -> (id, sheet.Name))
    )

/// Get all table names
let getTableNames (workbook: Workbook): string list =
    workbook.Tables
    |> Map.toList
    |> List.map fst

/// Workbook statistics
type WorkbookStats = {
    SheetCount: int
    TableCount: int
    TotalCells: int
    TotalFormulas: int
}

let getWorkbookStats (workbook: Workbook): WorkbookStats =
    let totalCells =
        workbook.Sheets
        |> Map.toList
        |> List.sumBy (fun (_, sheet) -> Map.count sheet.Model.Cells)

    let totalFormulas =
        workbook.Sheets
        |> Map.toList
        |> List.sumBy (fun (_, sheet) ->
            sheet.Model.Cells
            |> Map.toList
            |> List.filter (fun (_, cell) ->
                match cell.Value with
                | CellValue.Formula _ -> true
                | _ -> false
            )
            |> List.length
        )

    {
        SheetCount = Map.count workbook.Sheets
        TableCount = Map.count workbook.Tables
        TotalCells = totalCells
        TotalFormulas = totalFormulas
    }
