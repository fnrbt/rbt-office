module Rbt.Spreadsheet.Workbook

open System.Collections.Generic
open Rbt.Spreadsheet.Model

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

/// What a defined name carries besides what it refers to.
type DefinedNameInfo = {
    /// The name as it was written (the key of `Names` is upper-cased).
    Name: string
    Comment: string option
    /// A name the file hides from the name manager (`hidden="1"`).
    Hidden: bool
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
    /// The spelling, comment and hidden flag of defined names, by the key
    /// of `Names`. A name without an entry is spelled as its key.
    NameInfo: Map<SheetId option * string, DefinedNameInfo>
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
        NameInfo = Map.empty
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

/// What is wrong with a defined name, or None for one Excel accepts: it
/// starts with a letter, `_` or `\`; holds letters, digits, `_`, `.`, `\`
/// and `?` only; is at most 255 characters; and cannot be read as a cell
/// reference (A1, XFD1048576, R1C1, R, C) or as TRUE / FALSE.
let definedNameProblem (name: string) : string option =
    let name = if isNull name then "" else name
    let upper = name.ToUpperInvariant()
    let isDigits (text: string) = text |> Seq.forall System.Char.IsDigit
    let isR1C1 =
        // R, C, R12, C3, R1C1, RC, R2C ...
        upper.Length > 0 && upper.[0] = 'R'
        && (let rest = upper.Substring 1
            let c = rest.IndexOf 'C'
            if c < 0 then isDigits rest
            else isDigits (rest.Substring(0, c)) && isDigits (rest.Substring(c + 1)))
        || (upper.Length > 0 && upper.[0] = 'C' && isDigits (upper.Substring 1))
    let isA1 =
        let letters = upper |> Seq.takeWhile (fun c -> c >= 'A' && c <= 'Z') |> Seq.length
        letters >= 1 && letters <= 3 && letters < upper.Length && isDigits (upper.Substring letters)
        && (letterToColumn (upper.Substring(0, letters)) <= 16383)
        && (match System.Int32.TryParse (upper.Substring letters) with
            | true, row -> row >= 1 && row <= 1048576
            | _ -> false)
    if name = "" then Some "a name cannot be empty"
    elif name.Length > 255 then Some "a name can have at most 255 characters"
    elif not (System.Char.IsLetter name.[0] || name.[0] = '_' || name.[0] = '\\') then
        Some (sprintf "the name \"%s\" must start with a letter, an underscore or a backslash" name)
    elif name |> Seq.exists (fun c -> not (System.Char.IsLetterOrDigit c || c = '_' || c = '.' || c = '\\' || c = '?')) then
        Some (sprintf "the name \"%s\" may hold letters, digits, underscores, periods, backslashes and question marks only" name)
    elif isA1 || isR1C1 then Some (sprintf "the name \"%s\" reads as a cell reference" name)
    elif upper = "TRUE" || upper = "FALSE" then Some (sprintf "the name \"%s\" is a logical value" name)
    else None

/// A defined name with what it refers to and what else is known of it.
type DefinedName = {
    Scope: SheetId option
    Name: string
    /// The refers-to formula, without a leading `=`.
    RefersTo: string
    Comment: string option
    Hidden: bool
}

let private refersToBody (text: string) : string =
    let text = if isNull text then "" else text.Trim()
    if text.StartsWith "=" then text.Substring 1 else text

/// The defined names of the workbook: global ones first, then by name.
let definedNames (workbook: Workbook) : DefinedName list =
    workbook.Names
    |> Map.toList
    |> List.map (fun ((scope, key), refersTo) ->
        match Map.tryFind (scope, key) workbook.NameInfo with
        | Some info -> { Scope = scope; Name = info.Name; RefersTo = refersToBody refersTo; Comment = info.Comment; Hidden = info.Hidden }
        | None -> { Scope = scope; Name = key; RefersTo = refersToBody refersTo; Comment = None; Hidden = false })

/// The defined name `name` of a scope (not the one a sheet would see: see getName).
let tryFindDefinedName (scope: SheetId option) (name: string) (workbook: Workbook) : DefinedName option =
    let key = scope, name.ToUpperInvariant()
    Map.tryFind key workbook.Names
    |> Option.map (fun refersTo ->
        match Map.tryFind key workbook.NameInfo with
        | Some info -> { Scope = scope; Name = info.Name; RefersTo = refersToBody refersTo; Comment = info.Comment; Hidden = info.Hidden }
        | None -> { Scope = scope; Name = snd key; RefersTo = refersToBody refersTo; Comment = None; Hidden = false })

/// Define a name or replace the one of the same scope (names differ by
/// more than case). The refers-to text is kept without a leading `=`. A
/// replaced name keeps its hidden flag; its spelling and comment are the
/// new ones.
let setDefinedName (scope: SheetId option) (name: string) (refersTo: string) (comment: string option) (workbook: Workbook) : Workbook =
    let key = scope, name.ToUpperInvariant()
    let hidden = match Map.tryFind key workbook.NameInfo with Some info -> info.Hidden | None -> false
    { workbook with
        Names = Map.add key (refersToBody refersTo) workbook.Names
        NameInfo = Map.add key { Name = name; Comment = comment; Hidden = hidden } workbook.NameInfo }

/// Remove the name of a scope. None when there is no such name.
let deleteDefinedName (scope: SheetId option) (name: string) (workbook: Workbook) : Workbook option =
    let key = scope, name.ToUpperInvariant()
    if Map.containsKey key workbook.Names then
        Some { workbook with Names = Map.remove key workbook.Names; NameInfo = Map.remove key workbook.NameInfo }
    else None

/// Put one name back as it was (`before`: None = it did not exist), for undo / redo.
let restoreDefinedName (scope: SheetId option) (name: string) (before: DefinedName option) (workbook: Workbook) : Workbook =
    let key = scope, name.ToUpperInvariant()
    match before with
    | Some known ->
        { workbook with
            Names = Map.add key known.RefersTo workbook.Names
            NameInfo = Map.add key { Name = known.Name; Comment = known.Comment; Hidden = known.Hidden } workbook.NameInfo }
    | None -> { workbook with Names = Map.remove key workbook.Names; NameInfo = Map.remove key workbook.NameInfo }

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

/// Build a structured-reference `Table` from a 1-indexed XLSX table region
/// (rows and cols 1-based, A=1 — the form `Xlsx.TableDefinition`/`Model.Table`
/// use) by shifting to the 0-indexed model coordinates this module stores.
/// `columns` are the declared `<tableColumn>` names; when empty they're read
/// left→right from the header row of `model`.
let makeTableFromRegion (sheetId: SheetId) (name: string)
                        (startRow1: int) (startCol1: int) (endRow1: int) (endCol1: int)
                        (headerRowCount: int) (totalsRowCount: int)
                        (columns: string list) (model: SpreadsheetModel): Table =
    let sr = startRow1 - 1
    let sc = startCol1 - 1
    let er = endRow1 - 1
    let ec = endCol1 - 1
    let cols =
        if not (List.isEmpty columns) then columns
        else [ for c in sc .. ec -> displayValue (getCellValue { Row = sr; Col = c } model) ]
    { Name = name
      SheetId = sheetId
      StartRow = sr
      StartCol = sc
      EndRow = er
      EndCol = ec
      Columns = cols
      HasHeaders = headerRowCount > 0
      HasTotals = totalsRowCount > 0 }

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

// --- Row / column insertion and deletion across the workbook -----------------

/// What `applyStructureEdit` changed outside the edited sheet's own model,
/// plus that model as it was — enough for `undoStructureEdit` to put
/// everything back.
type StructureEditUndo = {
    Sheet: SheetId
    Before: SpreadsheetModel
    /// Formula cells rewritten on other sheets: (sheet, cell, before, after).
    ForeignCells: (SheetId * CellAddress * CellData * CellData) list
    /// Structured-reference tables moved or removed: (key, before, after).
    Tables: (string * Table * Table option) list
    /// Defined names whose refers-to text changed: (key, before, after).
    Names: ((SheetId option * string) * string * string) list
}

let private moveTableForEdit (edit: StructureEdit) (table: Table) : Table option =
    let onRows, isInsert, at, count =
        match edit with
        | InsertRows (at, count) -> true, true, at, count
        | DeleteRows (at, count) -> true, false, at, count
        | InsertCols (at, count) -> false, true, at, count
        | DeleteCols (at, count) -> false, false, at, count
    let deleted index = not isInsert && index >= at && index < at + count
    let moveSpan lo hi =
        if isInsert then
            Some ((if lo >= at then lo + count else lo), (if hi >= at then hi + count else hi))
        elif lo >= at && hi < at + count then None
        else
            let newLo = if lo >= at + count then lo - count elif lo >= at then at else lo
            let newHi = if hi >= at + count then hi - count elif hi >= at then at - 1 else hi
            Some (newLo, newHi)
    if onRows then
        // Without its header row the column names are gone: drop the table.
        if table.HasHeaders && deleted table.StartRow then None
        else
            moveSpan table.StartRow table.EndRow
            |> Option.map (fun (lo, hi) ->
                { table with
                    StartRow = lo
                    EndRow = hi
                    HasTotals = table.HasTotals && not (deleted table.EndRow) })
    else
        moveSpan table.StartCol table.EndCol
        |> Option.map (fun (lo, hi) ->
            let columns =
                if isInsert then
                    if at > table.StartCol && at <= table.EndCol then
                        let before = List.truncate (at - table.StartCol) table.Columns
                        let after = table.Columns |> List.skip (min (at - table.StartCol) table.Columns.Length)
                        let taken = System.Collections.Generic.HashSet<string>(table.Columns |> List.map (fun c -> c.ToUpperInvariant()))
                        let mutable next = 1
                        let fresh =
                            [ for _ in 1 .. count do
                                while taken.Contains(sprintf "COLUMN%d" next) do next <- next + 1
                                taken.Add(sprintf "COLUMN%d" next) |> ignore
                                yield sprintf "Column%d" next ]
                        before @ fresh @ after
                    else table.Columns
                else
                    table.Columns
                    |> List.mapi (fun i name -> table.StartCol + i, name)
                    |> List.filter (fun (col, _) -> not (deleted col))
                    |> List.map snd
            { table with StartCol = lo; EndCol = hi; Columns = columns })

/// Insert or delete rows/columns on one sheet of a workbook, with Excel's
/// semantics everywhere the edit is visible: the sheet's own content moves,
/// formulas on *every* sheet that point at it are re-pointed, and so are
/// defined names and structured-reference tables. Cached formula values are
/// left for the caller to recalculate.
let applyStructureEdit (sheetId: SheetId) (edit: StructureEdit) (workbook: Workbook) : Workbook * StructureEditUndo =
    match Map.tryFind sheetId workbook.Sheets with
    | None ->
        workbook, { Sheet = sheetId; Before = createEmpty 1 1; ForeignCells = []; Tables = []; Names = [] }
    | Some sheet ->
        let namesThisSheet (qualifier: string) =
            qualifier.Equals(sheet.Name, System.StringComparison.OrdinalIgnoreCase)
        let fromEditedSheet (qualifier: string option) =
            match qualifier with
            | None -> true
            | Some name -> namesThisSheet name
        let fromOtherSheet (qualifier: string option) =
            match qualifier with
            | None -> false
            | Some name -> namesThisSheet name

        let foreign = ResizeArray<SheetId * CellAddress * CellData * CellData>()
        let sheets =
            workbook.Sheets
            |> Map.map (fun id other ->
                if id = sheetId then
                    { other with Model = applyStructureEditWith fromEditedSheet edit other.Model }
                else
                    let rewritten = rewriteFormulasForEdit edit fromOtherSheet other.Model
                    if not (obj.ReferenceEquals(rewritten, other.Model)) then
                        for KeyValue(addr, after) in rewritten.Cells do
                            match Map.tryFind addr other.Model.Cells with
                            | Some before when before <> after -> foreign.Add((id, addr, before, after))
                            | _ -> ()
                    { other with Model = rewritten })

        let tableChanges = ResizeArray<string * Table * Table option>()
        let tables =
            workbook.Tables
            |> Map.fold (fun acc key table ->
                if table.SheetId <> sheetId then acc
                else
                    let moved = moveTableForEdit edit table
                    if moved <> Some table then tableChanges.Add((key, table, moved))
                    match moved with
                    | Some updated -> Map.add key updated acc
                    | None -> Map.remove key acc) workbook.Tables

        let nameChanges = ResizeArray<(SheetId option * string) * string * string>()
        let names =
            workbook.Names
            |> Map.map (fun key refersTo ->
                let targets = if fst key = Some sheetId then fromEditedSheet else fromOtherSheet
                let adjusted = adjustFormulaForEdit edit targets refersTo
                if adjusted <> refersTo then nameChanges.Add((key, refersTo, adjusted))
                adjusted)

        { workbook with Sheets = sheets; Tables = tables; Names = names },
        { Sheet = sheetId
          Before = sheet.Model
          ForeignCells = List.ofSeq foreign
          Tables = List.ofSeq tableChanges
          Names = List.ofSeq nameChanges }

/// Reverse `applyStructureEdit`. The edited sheet goes back to its earlier
/// state wholesale (cells, merges, sizes, formats, validations, tables, spill
/// map — its undo history is linear); on other sheets only the formulas the
/// edit rewrote are restored, and only if nobody has changed them since.
let undoStructureEdit (undo: StructureEditUndo) (workbook: Workbook) : Workbook =
    let sheets =
        workbook.Sheets
        |> Map.map (fun id sheet ->
            if id = undo.Sheet then
                let current = sheet.Model
                { sheet with
                    Model =
                        { undo.Before with
                            History = current.History
                            Clipboard = current.Clipboard
                            EditingCell = current.EditingCell
                            EditBuffer = current.EditBuffer
                            Macros = current.Macros } }
            else
                let mutable model = sheet.Model
                for (owner, addr, before, after) in undo.ForeignCells do
                    if owner = id then
                        match Map.tryFind addr model.Cells with
                        | Some current ->
                            let sameFormula =
                                match current.Value, after.Value with
                                | Formula (a, _), Formula (b, _) -> a = b
                                | _ -> false
                            if sameFormula then model <- setCell addr before model
                        | None -> ()
                if obj.ReferenceEquals(model, sheet.Model) then sheet else { sheet with Model = model })
    let tables =
        undo.Tables
        |> List.fold (fun acc (key, before, after) ->
            if Map.tryFind key acc = after then Map.add key before acc else acc) workbook.Tables
    let names =
        undo.Names
        |> List.fold (fun acc (key, before, after) ->
            if Map.tryFind key acc = Some after then Map.add key before acc else acc) workbook.Names
    { workbook with Sheets = sheets; Tables = tables; Names = names }

// --- Cut and paste: references follow the moved cells -------------------------

/// What `repointMovedCells` rewrote, for `undoRepointMovedCells`.
type CellMoveUndo = {
    /// Formula cells rewritten: (sheet, cell, before, after).
    Cells: (SheetId * CellAddress * CellData * CellData) list
    /// Defined names whose refers-to text changed: (key, before, after).
    Names: ((SheetId option * string) * string * string) list
}

/// After the block `(top, left, bottom, right)` of `sourceSheet` was cut and
/// pasted `(deltaRow, deltaCol)` away on `targetSheet`, re-point every formula
/// and defined name in the workbook that referred to the moved cells — Excel
/// moves references with the cells. Call it once the block sits at its new
/// place. The formulas inside the block keep pointing at the cells they
/// named (qualified with the old sheet's name when the block changed sheet),
/// except that references to cells of the block itself travel with it.
let repointMovedCells (sourceSheet: SheetId) (top: int, left: int, bottom: int, right: int)
                      (targetSheet: SheetId) (deltaRow: int, deltaCol: int)
                      (workbook: Workbook) : Workbook * CellMoveUndo =
    let nameOf id = Map.tryFind id workbook.Sheets |> Option.map (fun sheet -> sheet.Name)
    match nameOf sourceSheet, nameOf targetSheet with
    | Some sourceName, Some targetName when deltaRow <> 0 || deltaCol <> 0 || sourceSheet <> targetSheet ->
        let sameSheet = sourceSheet = targetSheet
        let names (sheetName: string) (qualifier: string) =
            qualifier.Equals(sheetName, System.StringComparison.OrdinalIgnoreCase)
        /// `owner` is the sheet an unqualified reference of the formula means.
        let moveFor (host: SheetId option) (owner: SheetId option) (carried: bool) : CellMove =
            { Top = top; Left = left; Bottom = bottom; Right = right
              DeltaRow = deltaRow; DeltaCol = deltaCol
              OnSourceSheet = (fun q -> match q with None -> owner = Some sourceSheet | Some n -> names sourceName n)
              OnTargetSheet = (fun q -> match q with None -> owner = Some targetSheet | Some n -> names targetName n)
              Requalify = (fun q -> if sameSheet then q elif host = Some targetSheet then None else Some targetName)
              QualifyStationary = if carried && not sameSheet then Some sourceName else None }
        let changedCells = ResizeArray<SheetId * CellAddress * CellData * CellData>()
        let sheets =
            workbook.Sheets
            |> Map.map (fun id sheet ->
                let mutable cells = sheet.Model.Cells
                for KeyValue(addr, cellData) in sheet.Model.Cells do
                    match cellData.Value with
                    | Formula (formula, cached) ->
                        let carried =
                            id = targetSheet
                            && addr.Row - deltaRow >= top && addr.Row - deltaRow <= bottom
                            && addr.Col - deltaCol >= left && addr.Col - deltaCol <= right
                        let move = moveFor (Some id) (Some (if carried then sourceSheet else id)) carried
                        let rewritten = moveFormulaReferences move formula
                        if rewritten <> formula then
                            let after = { cellData with Value = Formula (rewritten, cached) }
                            cells <- Map.add addr after cells
                            changedCells.Add((id, addr, cellData, after))
                    | _ -> ()
                if obj.ReferenceEquals(cells, sheet.Model.Cells) then sheet
                else { sheet with Model = { sheet.Model with Cells = cells } |> invalidateFilterProjection })
        let changedNames = ResizeArray<(SheetId option * string) * string * string>()
        let definedNames =
            workbook.Names
            |> Map.map (fun key refersTo ->
                let rewritten = moveFormulaReferences (moveFor None (fst key) false) refersTo
                if rewritten <> refersTo then changedNames.Add((key, refersTo, rewritten))
                rewritten)
        { workbook with Sheets = sheets; Names = definedNames },
        { Cells = List.ofSeq changedCells; Names = List.ofSeq changedNames }
    | _ -> workbook, { Cells = []; Names = [] }

/// Reverse `repointMovedCells`: every formula it rewrote goes back, unless it
/// was edited since.
let undoRepointMovedCells (undo: CellMoveUndo) (workbook: Workbook) : Workbook =
    let sheets =
        workbook.Sheets
        |> Map.map (fun id sheet ->
            let mutable model = sheet.Model
            for (owner, addr, before, after) in undo.Cells do
                if owner = id then
                    match Map.tryFind addr model.Cells with
                    | Some current ->
                        (match current.Value, after.Value with
                         | Formula (a, _), Formula (b, _) when a = b -> model <- setCell addr before model
                         | _ -> ())
                    | None -> ()
            if obj.ReferenceEquals(model, sheet.Model) then sheet else { sheet with Model = model })
    let definedNames =
        undo.Names
        |> List.fold (fun acc (key, before, after) ->
            if Map.tryFind key acc = Some after then Map.add key before acc else acc) workbook.Names
    { workbook with Sheets = sheets; Names = definedNames }

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
