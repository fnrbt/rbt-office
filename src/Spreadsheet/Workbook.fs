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

/// Iterative calculation of circular references (see `Workbook.Iteration`).
type IterativeCalculation = {
    /// Excel's default: 100.
    MaxIterations: int
    /// Excel's default: 0.001.
    MaxChange: float
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
    /// Excel's "enable iterative calculation" (`<calcPr iterate="1">`): the
    /// formulas of a circular reference are calculated again and again from
    /// the values they have, until none changes by `MaxChange` or more, or
    /// `MaxIterations` times. None (the default): they show the cycle error.
    Iteration: IterativeCalculation option
    /// What the workbook last read of the workbooks its formulas refer to (`[1]Sheet!A1`,
    /// `'[1]My sheet'!A1`): the values the file keeps of them (`xl/externalLinks`), by
    /// "[n]SHEET NAME" upper-cased, n the external reference's number (1 the first).
    /// A recalculation reads these, as Excel's does without the other workbook open.
    External: Map<string, SpreadsheetModel>
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
        Iteration = None
        External = Map.empty
    }

/// The workbook is in the 1904 date system: its sheets say so
/// (`SpreadsheetModel.Date1904`; every sheet of a workbook says the same).
let date1904 (workbook: Workbook) : bool =
    workbook.Sheets |> Map.exists (fun _ sheet -> sheet.Model.Date1904)

/// The workbook in the 1904 date system (true) or the 1900 one: every sheet
/// is told. The serials of the cells stay as they are.
let withDate1904 (flag: bool) (workbook: Workbook) : Workbook =
    if workbook.Sheets |> Map.forall (fun _ sheet -> sheet.Model.Date1904 = flag) then workbook
    else { workbook with Sheets = workbook.Sheets |> Map.map (fun _ sheet -> if sheet.Model.Date1904 = flag then sheet else { sheet with Model = { sheet.Model with Date1904 = flag } }) }

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
        // (A sheet added to a workbook is in the workbook's date system.)
        Model = { createEmpty 1000 26 with Date1904 = date1904 workbook }
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

let private findSheetNamed (name: string) (workbook: Workbook): SheetId option =
    workbook.Sheets
    |> Map.tryPick (fun id sheet ->
        if sheet.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase) then
            Some id
        else
            None
    )

#if FABLE_COMPILER
// Asked for every sheet-qualified reference of every formula when the
// dependency graph is built: what a name (as written) resolved to is kept
// for as long as the questions are about the same sheets.
// (The sheets are held weakly: a workbook that is let go of is not kept by this.)
[<Fable.Core.Emit("new WeakRef($0)")>]
let weakRefOf (target: obj) : obj = failwith "JS only"
[<Fable.Core.Emit("($0 !== null && $0.deref() === $1)")>]
let weakRefIs (reference: obj) (target: obj) : bool = failwith "JS only"
let mutable private sheetNamesOf : obj = null
let private sheetNamesFound = System.Collections.Generic.Dictionary<string, SheetId option>()
let mutable private sheetNamesHeld = false

/// `work` with the names of `workbook`'s sheets held: the questions asked
/// while it runs are all about sheets of these names (a recalculation puts
/// a new model into a sheet with every result, and names none anew), so
/// what a name resolved to is kept throughout.
let withSheetNamesHeld (workbook: Workbook) (work: unit -> 'T) : 'T =
    if sheetNamesHeld then work ()
    else
        sheetNamesOf <- null
        sheetNamesFound.Clear()
        sheetNamesHeld <- true
        try work ()
        finally
            sheetNamesHeld <- false
            sheetNamesFound.Clear()

/// Resolve sheet name to sheet ID
let resolveSheetName (name: string) (workbook: Workbook): SheetId option =
    if not sheetNamesHeld && not (weakRefIs sheetNamesOf (box workbook.Sheets)) then
        sheetNamesOf <- weakRefOf (box workbook.Sheets)
        sheetNamesFound.Clear()
    match Portable.Dict.tryFind name sheetNamesFound with
    | ValueSome found -> found
    | ValueNone ->
        let found = findSheetNamed name workbook
        sheetNamesFound.[name] <- found
        found
#else
// (The same on .NET, per thread: workbooks are calculated on any thread.)
type private SheetNameMemo() =
    [<System.ThreadStatic; DefaultValue>]
    static val mutable private sheets : System.WeakReference
    [<System.ThreadStatic; DefaultValue>]
    static val mutable private found : System.Collections.Generic.Dictionary<string, SheetId option>
    [<System.ThreadStatic; DefaultValue>]
    static val mutable private held : bool
    static member Held (work: unit -> 'T) : 'T =
        if SheetNameMemo.held then work ()
        else
            if isNull SheetNameMemo.found then SheetNameMemo.found <- System.Collections.Generic.Dictionary<string, SheetId option>()
            SheetNameMemo.sheets <- null
            SheetNameMemo.found.Clear()
            SheetNameMemo.held <- true
            try work ()
            finally
                SheetNameMemo.held <- false
                SheetNameMemo.found.Clear()
    static member Resolve (name: string, workbook: Workbook) : SheetId option =
        if isNull SheetNameMemo.found then SheetNameMemo.found <- System.Collections.Generic.Dictionary<string, SheetId option>()
        if not SheetNameMemo.held && (isNull SheetNameMemo.sheets || not (obj.ReferenceEquals(SheetNameMemo.sheets.Target, workbook.Sheets))) then
            SheetNameMemo.sheets <- System.WeakReference(box workbook.Sheets)
            SheetNameMemo.found.Clear()
        match Portable.Dict.tryFind name SheetNameMemo.found with
        | ValueSome known -> known
        | ValueNone ->
            let resolved = findSheetNamed name workbook
            SheetNameMemo.found.[name] <- resolved
            resolved

/// Resolve sheet name to sheet ID
let resolveSheetName (name: string) (workbook: Workbook): SheetId option = SheetNameMemo.Resolve (name, workbook)

/// `work` with the names of `workbook`'s sheets held: the questions asked
/// while it runs are all about sheets of these names (a recalculation puts
/// a new model into a sheet with every result, and names none anew), so
/// what a name resolved to is kept throughout.
let withSheetNamesHeld (workbook: Workbook) (work: unit -> 'T) : 'T = (ignore workbook; SheetNameMemo.Held work)
#endif

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

/// Replace a sheet model after a model-level command, keeping its table
/// registry in step. Structural commands use applyStructureEdit instead:
/// that operation also moves column identities and rewrites references.
let updateSheetModelWithTables (sheetId: SheetId) (model: SpreadsheetModel) (workbook: Workbook): Workbook =
    match Map.tryFind sheetId workbook.Sheets with
    | None -> workbook
    | Some sheet when obj.ReferenceEquals(sheet.Model, model) -> workbook
    | Some sheet ->
        let updated = updateSheetModel sheetId model workbook
        if obj.ReferenceEquals(sheet.Model.Tables, model.Tables) then updated
        else
            let mutable tables = workbook.Tables |> Map.filter (fun _ table -> table.SheetId <> sheetId)
            for table in model.Tables do
                if table.Name <> "" then
                    let columns =
                        match getTable table.Name workbook with
                        | Some previous when previous.SheetId = sheetId
                                             && List.length previous.Columns = table.EndCol - table.StartCol + 1 ->
                            previous.Columns
                        | _ -> []
                    let entry =
                        makeTableFromRegion sheetId table.Name table.StartRow table.StartCol table.EndRow table.EndCol
                            table.HeaderRowCount (if table.HasTotals then 1 else 0) columns model
                    tables <- Map.add table.Name entry tables
            { updated with Tables = tables }

// --- Row / column insertion and deletion across the workbook -----------------

/// What `applyStructureEdit` changed outside the edited sheet's own model,
/// plus that model as it was — enough for `undoStructureEdit` to put
/// everything back.
type StructureEditUndo = {
    Sheet: SheetId
    Before: SpreadsheetModel
    /// Formula cells rewritten on other sheets: (sheet, cell, before, after).
    ForeignCells: (SheetId * CellAddress * CellData * CellData) list
    /// Formula-bearing rules on other sheets: (sheet, rule id, before, after).
    ForeignFormats: (SheetId * int * ConditionalFormatRule * ConditionalFormatRule) list
    ForeignValidations: (SheetId * int * ValidationRule * ValidationRule) list
    /// Authored chart references on other sheets, before and after this edit.
    ForeignCharts: (SheetId * Map<string, string> * Map<string, string>) list
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
        workbook, { Sheet = sheetId; Before = createEmpty 1 1; ForeignCells = []; ForeignFormats = []; ForeignValidations = []; ForeignCharts = []; Tables = []; Names = [] }
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
        let foreignFormats = ResizeArray<SheetId * int * ConditionalFormatRule * ConditionalFormatRule>()
        let foreignValidations = ResizeArray<SheetId * int * ValidationRule * ValidationRule>()
        let foreignCharts = ResizeArray<SheetId * Map<string, string> * Map<string, string>>()
        let sheets =
            workbook.Sheets
            |> Map.map (fun id other ->
                if id = sheetId then
                    let edited = applyStructureEditWith fromEditedSheet edit other.Model
                    let rewrite formula = adjustFormulaForEdit edit fromEditedSheet formula
                    { other with Model = mapAuthoredChartFormulas rewrite edited }
                else
                    let rewritten = mapModelFormulas (adjustFormulaForEdit edit fromOtherSheet) other.Model
                    if rewritten.AuthoredCharts <> other.Model.AuthoredCharts then
                        foreignCharts.Add((id, other.Model.AuthoredCharts, rewritten.AuthoredCharts))
                    if not (obj.ReferenceEquals(rewritten, other.Model)) then
                        if not (obj.ReferenceEquals(rewritten.Cells, other.Model.Cells)) then
                            for KeyValue(addr, after) in rewritten.Cells do
                                match Map.tryFind addr other.Model.Cells with
                                | Some before when before <> after -> foreign.Add((id, addr, before, after))
                                | _ -> ()
                        List.iter2 (fun (before: ConditionalFormatRange) (after: ConditionalFormatRange) ->
                            if before.Format.Rule <> after.Format.Rule then
                                foreignFormats.Add((id, before.Id, before.Format.Rule, after.Format.Rule)))
                            other.Model.ConditionalFormats rewritten.ConditionalFormats
                        List.iter2 (fun (before: ValidationRange) (after: ValidationRange) ->
                            if before.Validation.Rule <> after.Validation.Rule then
                                foreignValidations.Add((id, before.Id, before.Validation.Rule, after.Validation.Rule)))
                            other.Model.Validations rewritten.Validations
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
          ForeignFormats = List.ofSeq foreignFormats
          ForeignValidations = List.ofSeq foreignValidations
          ForeignCharts = List.ofSeq foreignCharts
          Tables = List.ofSeq tableChanges
          Names = List.ofSeq nameChanges }

// Restore only formula fields still holding this edit's output. A later change
// to a rule's range, style, options, or another operand must survive undo.
let private undoFormulaText before after current =
    if before <> after && current = after then before else current

let private undoFormulaOption before after current =
    match before, after, current with
    | Some before, Some after, Some current -> Some (undoFormulaText before after current)
    | _ -> current

let private undoThreshold (before: CfThreshold) (after: CfThreshold) (current: CfThreshold) =
    if before.Kind = CfvoFormula && after.Kind = CfvoFormula && current.Kind = CfvoFormula then
        { current with Value = undoFormulaOption before.Value after.Value current.Value }
    else current

let private undoThresholds before after current =
    if List.length before = List.length after && List.length after = List.length current then
        List.map3 undoThreshold before after current
    else current

let private undoFormatRule before after current =
    let restore = undoFormulaText
    match before, after, current with
    | Expression b, Expression a, Expression c -> Expression (restore b a c)
    | CellIs (_, b1, b2), CellIs (_, a1, a2), CellIs (op, c1, c2) ->
        CellIs (op, restore b1 a1 c1, undoFormulaOption b2 a2 c2)
    | ColorScale b, ColorScale a, ColorScale c when List.length b = List.length a && List.length a = List.length c ->
        ColorScale (List.map3 (fun (bt, _) (at, _) (ct, color) -> undoThreshold bt at ct, color) b a c)
    | DataBar (bl, bh, _), DataBar (al, ah, _), DataBar (cl, ch, color) ->
        DataBar (undoThreshold bl al cl, undoThreshold bh ah ch, color)
    | IconSet (_, b, _), IconSet (_, a, _), IconSet (family, c, reverse) ->
        IconSet (family, undoThresholds b a c, reverse)
    | DataBarEx (bl, bh, _, _), DataBarEx (al, ah, _, _), DataBarEx (cl, ch, color, options) ->
        DataBarEx (undoThreshold bl al cl, undoThreshold bh ah ch, color, options)
    | IconSetEx (_, b, _, _), IconSetEx (_, a, _, _), IconSetEx (family, c, reverse, options) ->
        IconSetEx (family, undoThresholds b a c, reverse, options)
    | _ -> current

let private undoValidationRule before after current =
    let restore = undoFormulaText
    match before, after, current with
    | CustomFormula b, CustomFormula a, CustomFormula c -> CustomFormula (restore b a c)
    | ListSource b, ListSource a, ListSource c -> ListSource (restore b a c)
    | Comparison (_, _, b1, b2), Comparison (_, _, a1, a2), Comparison (kind, op, c1, c2) ->
        Comparison (kind, op, restore b1 a1 c1, undoFormulaOption b2 a2 c2)
    | _ -> current

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
                for (owner, before, after) in undo.ForeignCharts do
                    if owner = id && model.AuthoredCharts = after then
                        model <- { model with AuthoredCharts = before }
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
                let formatChanges =
                    undo.ForeignFormats
                    |> List.fold (fun changes (owner, ruleId, before, after) ->
                        if owner = id then Map.add ruleId (before, after) changes else changes) Map.empty
                if not formatChanges.IsEmpty then
                    let mutable changed = false
                    let updated =
                        model.ConditionalFormats
                        |> List.map (fun entry ->
                            match Map.tryFind entry.Id formatChanges with
                            | None -> entry
                            | Some (before, after) ->
                                let rule = undoFormatRule before after entry.Format.Rule
                                if rule = entry.Format.Rule then entry
                                else
                                    changed <- true
                                    { entry with Format = { entry.Format with Rule = rule } })
                    if changed then model <- { model with ConditionalFormats = updated }
                let validationChanges =
                    undo.ForeignValidations
                    |> List.fold (fun changes (owner, ruleId, before, after) ->
                        if owner = id then Map.add ruleId (before, after) changes else changes) Map.empty
                if not validationChanges.IsEmpty then
                    let mutable changed = false
                    let updated =
                        model.Validations
                        |> List.map (fun entry ->
                            match Map.tryFind entry.Id validationChanges with
                            | None -> entry
                            | Some (before, after) ->
                                let rule = undoValidationRule before after entry.Validation.Rule
                                if rule = entry.Validation.Rule then entry
                                else
                                    changed <- true
                                    { entry with Validation = { entry.Validation with Rule = rule } })
                    if changed then model <- { model with Validations = updated }
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

// --- Sheets: names, renaming, deleting, moving, copying -----------------------

/// What is wrong with a sheet name, or None: Excel's rules (1 to 31
/// characters, none of : \ / ? * [ ], no apostrophe at either end, not the
/// reserved "History").
let sheetNameProblem (name: string) : string option =
    if isNull name || name.Trim() = "" then Some "a sheet name cannot be empty"
    elif name.Length > 31 then Some "a sheet name cannot be longer than 31 characters"
    elif name.IndexOfAny [| ':'; '\\'; '/'; '?'; '*'; '['; ']' |] >= 0 then Some "a sheet name cannot contain any of : \\ / ? * [ ]"
    elif name.StartsWith "'" || name.EndsWith "'" then Some "a sheet name cannot begin or end with an apostrophe"
    elif name.Equals("History", System.StringComparison.OrdinalIgnoreCase) then Some "\"History\" is a reserved sheet name"
    else None

/// Whether a sheet other than `except` has the name (whatever the case).
let sheetNameTaken (except: SheetId option) (name: string) (workbook: Workbook) : bool =
    workbook.Sheets
    |> Map.exists (fun id sheet -> Some id <> except && sheet.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase))

/// The first of "Sheet1", "Sheet2", ... no sheet has.
let nextSheetName (workbook: Workbook) : string =
    let mutable number = 1
    while sheetNameTaken None (sprintf "Sheet%d" number) workbook do number <- number + 1
    sprintf "Sheet%d" number

/// The name of a copy of the sheet `name`: "Name (2)", "Name (3)", ... (the
/// name cut short where it would pass 31 characters).
let copySheetName (name: string) (workbook: Workbook) : string =
    let named (number: int) =
        let suffix = sprintf " (%d)" number
        let stem = if name.Length + suffix.Length > 31 then name.Substring(0, 31 - suffix.Length) else name
        // (A name may not end with an apostrophe; the suffix never does.)
        stem + suffix
    let mutable number = 2
    while sheetNameTaken None (named number) workbook do number <- number + 1
    named number

/// The position of a sheet in the tab order, or -1.
let sheetIndex (sheetId: SheetId) (workbook: Workbook) : int =
    workbook.SheetOrder |> List.tryFindIndex ((=) sheetId) |> Option.defaultValue -1

/// A sheet id no sheet of the workbook has and that is not in `used` (ids
/// that were given out before: a deleted sheet's id comes back with an undo).
let freshSheetId (used: SheetId -> bool) (workbook: Workbook) : SheetId =
    let mutable number = Map.count workbook.Sheets + 1
    while Map.containsKey (sprintf "sheet%d" number) workbook.Sheets || used (sprintf "sheet%d" number) do number <- number + 1
    sprintf "sheet%d" number

/// A formula's quoted sheet qualifier doubles apostrophes. Keep the cheap
/// prefilter (most formulas are unrelated) without skipping an undo/redo
/// rename or a deletion of a sheet whose name contains an apostrophe.
let private mentionsSheetName (name: string) (escapedName: string) (formula: string) =
    formula.IndexOf(name, System.StringComparison.OrdinalIgnoreCase) >= 0
    || (not (isNull escapedName) && formula.IndexOf(escapedName, System.StringComparison.OrdinalIgnoreCase) >= 0)

/// Rename a sheet, with everything that names it following: the formulas of
/// every sheet (cells, conditional formats, validations), the defined names
/// and the hyperlinks into the sheet. The new name is written quoted where a
/// reference needs it. (The name is not checked: see sheetNameProblem.)
let renameSheetEverywhere (sheetId: SheetId) (newName: string) (workbook: Workbook) : Workbook =
    match Map.tryFind sheetId workbook.Sheets with
    | None -> workbook
    | Some sheet when sheet.Name = newName -> workbook
    | Some sheet ->
        let oldName = sheet.Name
        let escapedName = if oldName.IndexOf('\'') >= 0 then oldName.Replace("'", "''") else null
        let change = RenameSheetRefs (oldName, newName)
        // (Most formulas do not mention the sheet: those are not parsed.)
        let rewrite (formula: string) =
            if mentionsSheetName oldName escapedName formula then rewriteFormulaForSheets change formula
            else formula
        let rewriteLink (link: SheetHyperlink) =
            if not link.IsInternal then link
            else
                let rewritten = rewrite ("=" + link.Target)
                let target = if rewritten.StartsWith "=" then rewritten.Substring 1 else rewritten
                if target = link.Target then link else { link with Target = target }
        let sheets =
            workbook.Sheets
            |> Map.map (fun id other ->
                let model = mapModelFormulas rewrite other.Model
                let links = model.Hyperlinks |> List.map rewriteLink
                let model = if links = model.Hyperlinks then model else { model with Hyperlinks = links }
                if id = sheetId then { other with Name = newName; Model = model }
                elif obj.ReferenceEquals(model, other.Model) then other
                else { other with Model = model })
        let names = workbook.Names |> Map.map (fun _ refersTo -> rewrite refersTo)
        { workbook with Sheets = sheets; Names = (if names = workbook.Names then workbook.Names else names) }

/// What `removeSheet` took out of a workbook and changed in it, for `restoreSheet`.
type SheetRemoval = {
    Sheet: Sheet
    /// Where the sheet was in the tab order.
    Index: int
    /// The sheet that was active.
    ActiveBefore: SheetId
    /// Formula cells of other sheets that referred to it: (sheet, cell, before, after).
    Cells: (SheetId * CellAddress * CellData * CellData) list
    /// Conditional formats and validations of other sheets whose formulas referred to it: (sheet, before, after).
    Formats: (SheetId * ConditionalFormatRange list * ConditionalFormatRange list) list
    Validations: (SheetId * ValidationRange list * ValidationRange list) list
    /// Links of other sheets that led into it: (sheet, before, after).
    Hyperlinks: (SheetId * SheetHyperlink list * SheetHyperlink list) list
    Charts: (SheetId * Map<string, string> * Map<string, string>) list
    /// Defined names that referred to it: (key, before, after).
    Names: ((SheetId option * string) * string * string) list
    /// The names of its own scope, gone with it: (key, refers-to, spelling and comment).
    ScopedNames: ((SheetId option * string) * string * DefinedNameInfo option) list
    /// Its structured-reference tables, gone with it.
    Tables: (string * Table) list
}

/// Why a sheet cannot be deleted (or hidden), or None: a workbook keeps at
/// least one visible sheet.
let lastVisibleSheetProblem (sheetId: SheetId) (workbook: Workbook) : string option =
    match Map.tryFind sheetId workbook.Sheets with
    | None -> Some "no such sheet"
    | Some sheet ->
        let otherVisible = workbook.Sheets |> Map.exists (fun id other -> id <> sheetId && other.Visible)
        if sheet.Visible && not otherVisible then Some "a workbook must keep at least one visible sheet"
        elif Map.count workbook.Sheets <= 1 then Some "a workbook must keep at least one sheet"
        else None

/// The visible sheet nearest to position `index` of the tab order (the one
/// after it first), leaving out `except`.
let private nearestVisibleSheet (index: int) (except: SheetId) (workbook: Workbook) : SheetId option =
    let order = List.toArray workbook.SheetOrder
    let visible (id: SheetId) = id <> except && (match Map.tryFind id workbook.Sheets with Some sheet -> sheet.Visible | None -> false)
    let after = order |> Array.mapi (fun i id -> i, id) |> Array.tryFind (fun (i, id) -> i >= index && visible id)
    let before = order |> Array.mapi (fun i id -> i, id) |> Array.tryFindBack (fun (i, id) -> i < index && visible id)
    match after, before with
    | Some (_, id), _ -> Some id
    | None, Some (_, id) -> Some id
    | None, None -> order |> Array.tryFind (fun id -> id <> except)

/// Delete a sheet as Excel does: references to it in the formulas of the
/// other sheets (cells, conditional formats, validations) and in defined
/// names become #REF!, the names of its own scope and its tables go with it,
/// and when it was the active sheet the nearest visible one becomes active.
/// None when the sheet is not there or is the only one. (Whether it may be
/// deleted - the last visible sheet may not - is lastVisibleSheetProblem.)
let removeSheet (sheetId: SheetId) (workbook: Workbook) : (Workbook * SheetRemoval) option =
    match Map.tryFind sheetId workbook.Sheets with
    | None -> None
    | Some _ when Map.count workbook.Sheets <= 1 -> None
    | Some sheet ->
        let tabNames = workbook.SheetOrder |> List.choose (fun id -> Map.tryFind id workbook.Sheets |> Option.map (fun known -> known.Name))
        let change = DeleteSheetRefs (sheet.Name, tabNames)
        let escapedName = if sheet.Name.IndexOf('\'') >= 0 then sheet.Name.Replace("'", "''") else null
        let rewrite (formula: string) =
            if mentionsSheetName sheet.Name escapedName formula then rewriteFormulaForSheets change formula
            else formula
        let cells = ResizeArray<SheetId * CellAddress * CellData * CellData>()
        let formats = ResizeArray<SheetId * ConditionalFormatRange list * ConditionalFormatRange list>()
        let validations = ResizeArray<SheetId * ValidationRange list * ValidationRange list>()
        let hyperlinks = ResizeArray<SheetId * SheetHyperlink list * SheetHyperlink list>()
        let charts = ResizeArray<SheetId * Map<string, string> * Map<string, string>>()
        // A link into the deleted sheet leads nowhere: its target says #REF!, as in Excel.
        let rewriteLink (link: SheetHyperlink) =
            if not link.IsInternal then link
            else
                let rewritten = rewrite ("=" + link.Target)
                let target = if rewritten.StartsWith "=" then rewritten.Substring 1 else rewritten
                if target = link.Target then link else { link with Target = target }
        let sheets =
            workbook.Sheets
            |> Map.remove sheetId
            |> Map.map (fun id other ->
                let model = mapModelFormulas rewrite other.Model
                let model =
                    if model.Hyperlinks.IsEmpty then model
                    else
                        let links = model.Hyperlinks |> List.map rewriteLink
                        if links = model.Hyperlinks then model
                        else
                            hyperlinks.Add((id, model.Hyperlinks, links))
                            { model with Hyperlinks = links }
                if obj.ReferenceEquals(model, other.Model) then other
                else
                    if not (obj.ReferenceEquals(model.Cells, other.Model.Cells)) then
                        for KeyValue(addr, after) in model.Cells do
                            match Map.tryFind addr other.Model.Cells with
                            | Some before when not (obj.ReferenceEquals(before, after)) && before <> after -> cells.Add((id, addr, before, after))
                            | _ -> ()
                    if not (obj.ReferenceEquals(model.ConditionalFormats, other.Model.ConditionalFormats)) then
                        formats.Add((id, other.Model.ConditionalFormats, model.ConditionalFormats))
                    if not (obj.ReferenceEquals(model.Validations, other.Model.Validations)) then
                        validations.Add((id, other.Model.Validations, model.Validations))
                    if model.AuthoredCharts <> other.Model.AuthoredCharts then
                        charts.Add((id, other.Model.AuthoredCharts, model.AuthoredCharts))
                    { other with Model = model })
        let scoped = workbook.Names |> Map.toList |> List.filter (fun ((scope, _), _) -> scope = Some sheetId)
        let nameChanges = ResizeArray<(SheetId option * string) * string * string>()
        let names =
            workbook.Names
            |> Map.filter (fun (scope, _) _ -> scope <> Some sheetId)
            |> Map.map (fun key refersTo ->
                let rewritten = rewrite refersTo
                if rewritten <> refersTo then nameChanges.Add((key, refersTo, rewritten))
                rewritten)
        let tables = workbook.Tables |> Map.toList |> List.filter (fun (_, table) -> table.SheetId = sheetId)
        let index = sheetIndex sheetId workbook
        let active =
            if workbook.ActiveSheet <> sheetId then workbook.ActiveSheet
            else nearestVisibleSheet index sheetId workbook |> Option.defaultValue workbook.ActiveSheet
        Some ({ workbook with
                  Sheets = sheets
                  SheetOrder = workbook.SheetOrder |> List.filter ((<>) sheetId)
                  ActiveSheet = active
                  Names = names
                  NameInfo = workbook.NameInfo |> Map.filter (fun (scope, _) _ -> scope <> Some sheetId)
                  Tables = workbook.Tables |> Map.filter (fun _ table -> table.SheetId <> sheetId) },
              { Sheet = sheet
                Index = index
                ActiveBefore = workbook.ActiveSheet
                Cells = List.ofSeq cells
                Formats = List.ofSeq formats
                Validations = List.ofSeq validations
                Hyperlinks = List.ofSeq hyperlinks
                Charts = List.ofSeq charts
                Names = List.ofSeq nameChanges
                ScopedNames = scoped |> List.map (fun (key, refersTo) -> key, refersTo, Map.tryFind key workbook.NameInfo)
                Tables = tables })

/// Put a removed sheet back where it was, with its names and tables, and
/// give the formulas that `removeSheet` turned into #REF! their references
/// back - those nobody changed since. The sheet that was active is active again.
let restoreSheet (removal: SheetRemoval) (workbook: Workbook) : Workbook =
    let sheetId = removal.Sheet.Id
    if Map.containsKey sheetId workbook.Sheets then workbook
    else
        let sheets =
            workbook.Sheets
            |> Map.map (fun id sheet ->
                let mutable model = sheet.Model
                for (owner, addr, before, after) in removal.Cells do
                    if owner = id then
                        match Map.tryFind addr model.Cells with
                        | Some current ->
                            let sameFormula =
                                match current.Value, after.Value with
                                | Formula (a, _), Formula (b, _) -> a = b
                                | _ -> false
                            if sameFormula then model <- setCell addr before model
                        | None -> ()
                for (owner, before, after) in removal.Formats do
                    if owner = id && model.ConditionalFormats = after then model <- { model with ConditionalFormats = before }
                for (owner, before, after) in removal.Validations do
                    if owner = id && model.Validations = after then model <- { model with Validations = before }
                for (owner, before, after) in removal.Hyperlinks do
                    if owner = id && model.Hyperlinks = after then model <- { model with Hyperlinks = before }
                for (owner, before, after) in removal.Charts do
                    if owner = id && model.AuthoredCharts = after then model <- { model with AuthoredCharts = before }
                if obj.ReferenceEquals(model, sheet.Model) then sheet else { sheet with Model = model })
            |> Map.add sheetId removal.Sheet
        let order =
            let current = workbook.SheetOrder
            let at = max 0 (min (List.length current) removal.Index)
            List.take at current @ [ sheetId ] @ List.skip at current
        let names =
            removal.Names
            |> List.fold (fun acc (key, before, after) -> if Map.tryFind key acc = Some after then Map.add key before acc else acc) workbook.Names
        let names = removal.ScopedNames |> List.fold (fun acc (key, refersTo, _) -> Map.add key refersTo acc) names
        let nameInfo =
            removal.ScopedNames
            |> List.fold (fun acc (key, _, info) -> match info with Some info -> Map.add key info acc | None -> acc) workbook.NameInfo
        let tables =
            removal.Tables |> List.fold (fun acc (key, table) -> if Map.containsKey key acc then acc else Map.add key table acc) workbook.Tables
        { workbook with
            Sheets = sheets
            SheetOrder = order
            ActiveSheet = (if Map.containsKey removal.ActiveBefore sheets then removal.ActiveBefore else workbook.ActiveSheet)
            Names = names
            NameInfo = nameInfo
            Tables = tables }

/// Put a new sheet at position `index` of the tab order (clamped) and make it active.
let insertSheet (sheet: Sheet) (index: int) (workbook: Workbook) : Workbook =
    if Map.containsKey sheet.Id workbook.Sheets then workbook
    else
        let at = max 0 (min (List.length workbook.SheetOrder) index)
        { workbook with
            Sheets = Map.add sheet.Id sheet workbook.Sheets
            SheetOrder = List.take at workbook.SheetOrder @ [ sheet.Id ] @ List.skip at workbook.SheetOrder
            ActiveSheet = sheet.Id }

/// Move a sheet to position `toIndex` of the tab order (the position it has
/// afterwards, clamped). Nothing refers to a sheet by its position but 3-D
/// references, which name their first and last sheet.
let moveSheet (sheetId: SheetId) (toIndex: int) (workbook: Workbook) : Workbook =
    if not (Map.containsKey sheetId workbook.Sheets) then workbook
    else
        let rest = workbook.SheetOrder |> List.filter ((<>) sheetId)
        let at = max 0 (min (List.length rest) toIndex)
        let order = List.take at rest @ [ sheetId ] @ List.skip at rest
        if order = workbook.SheetOrder then workbook else { workbook with SheetOrder = order }

/// Show or hide a sheet's tab. `veryHidden` (with `visible` false) is
/// Excel's "very hidden". Hiding the active sheet makes the nearest visible
/// one active. (Whether it may be hidden: lastVisibleSheetProblem.)
let setSheetVisibility (sheetId: SheetId) (visible: bool) (veryHidden: bool) (workbook: Workbook) : Workbook =
    match Map.tryFind sheetId workbook.Sheets with
    | None -> workbook
    | Some sheet ->
        let veryHidden = veryHidden && not visible
        if sheet.Visible = visible && sheet.Model.VeryHidden = veryHidden then workbook
        else
            let model = if sheet.Model.VeryHidden = veryHidden then sheet.Model else { sheet.Model with VeryHidden = veryHidden }
            let changed = { workbook with Sheets = Map.add sheetId { sheet with Visible = visible; Model = model } workbook.Sheets }
            if visible || workbook.ActiveSheet <> sheetId then changed
            else
                match nearestVisibleSheet (sheetIndex sheetId workbook) sheetId changed with
                | Some active -> { changed with ActiveSheet = active }
                | None -> changed

/// "visible", "hidden" or "veryHidden".
let sheetVisibility (sheet: Sheet) : string =
    if sheet.Visible then "visible" elif sheet.Model.VeryHidden then "veryHidden" else "hidden"

/// A name for a copy of a table that no table has: "Name_2", "Name_3", ...
let private copyTableName (name: string) (taken: string -> bool) : string =
    let mutable number = 2
    while taken (sprintf "%s_%d" name number) do number <- number + 1
    sprintf "%s_%d" name number

/// Copy a sheet as Excel does: the copy `newId`, named `newName`, is put at
/// position `index` and made active. Its formulas that named the source
/// sheet name the copy; the names of the source's scope are copied to the
/// copy's; its tables are copied under new names ("Sales_2"), which the
/// copy's structured references follow. Pictures are copied, and so are the
/// drawing objects the editor only keeps of a file - charts, shapes: written
/// to XLSX, a copied chart is a chart part of its own whose series read the
/// copy where the source's read the source. Pivot tables stay with the
/// source: the copy has their cells as values, not the pivot tables.
let duplicateSheet (sourceId: SheetId) (newId: SheetId) (newName: string) (index: int) (workbook: Workbook) : Workbook =
    match Map.tryFind sourceId workbook.Sheets with
    | None -> workbook
    | Some _ when Map.containsKey newId workbook.Sheets -> workbook
    | Some source ->
        let sourceModel = resolveSheet source.Model
        let mutable taken = workbook.Tables |> Map.toList |> List.map (fun (name, _) -> name.ToUpperInvariant()) |> Set.ofList
        let tableNames =
            sourceModel.Tables
            |> List.filter (fun table -> table.Name <> "")
            |> List.map (fun table ->
                let copyName = copyTableName table.Name (fun name -> Set.contains (name.ToUpperInvariant()) taken)
                taken <- Set.add (copyName.ToUpperInvariant()) taken
                table.Name, copyName)
        let rewrites = RenameSheetRefs (source.Name, newName) :: (tableNames |> List.map RenameTableRefs)
        let rewrite (formula: string) = rewrites |> List.fold (fun text change -> rewriteFormulaForSheets change text) formula
        let renamedTable (name: string) =
            tableNames |> List.tryFind (fun (was, _) -> was = name) |> Option.map snd |> Option.defaultValue name
        let model =
            { mapModelFormulas rewrite sourceModel with
                Tables = sourceModel.Tables |> List.map (fun table -> { table with Name = renamedTable table.Name })
                // (The kept drawing objects - charts, shapes - come along: an export copies
                // the chart parts behind them, and the parts of the pivot tables, each
                // where `KeptRanges` has it. The slicers and timelines stay with the
                // source: the frame of one is not copied, a sheet that shows a slicer
                // it does not list is a damaged file to Excel.)
                Drawings =
                    sourceModel.Drawings
                    |> List.filter (fun drawing ->
                        match drawing.Content with
                        | KeptObject (_, xml) ->
                            not (xml.Contains "drawing/2010/slicer" || xml.Contains "drawing/2012/timeslicer" || xml.Contains ":slicer " || xml.Contains ":timeslicer ")
                        | _ -> true)
                KeptRanges = sourceModel.KeptRanges |> List.filter (fun kept -> kept.Key.StartsWith "xl/pivotTables/")
                SourceSheet = None
                KeptFrom = (match sourceModel.SourceSheet with Some origin -> Some origin | None -> sourceModel.KeptFrom)
                Deferred = None
                History = emptyHistory
                Clipboard = None
                EditingCell = None
                EditBuffer = "" }
        let tables =
            workbook.Tables
            |> Map.fold (fun acc _ table ->
                if table.SheetId <> sourceId then acc
                else
                    let name = renamedTable table.Name
                    if name = table.Name then acc else Map.add name { table with Name = name; SheetId = newId } acc) workbook.Tables
        let scoped = workbook.Names |> Map.toList |> List.filter (fun ((scope, _), _) -> scope = Some sourceId)
        let names = scoped |> List.fold (fun acc ((_, key), refersTo) -> Map.add (Some newId, key) (rewrite refersTo) acc) workbook.Names
        let nameInfo =
            scoped
            |> List.fold (fun acc ((_, key), _) ->
                match Map.tryFind (Some sourceId, key) workbook.NameInfo with
                | Some info -> Map.add (Some newId, key) info acc
                | None -> acc) workbook.NameInfo
        { insertSheet { Id = newId; Name = newName; Model = model; Visible = true } index workbook with
            Tables = tables
            Names = names
            NameInfo = nameInfo }

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
