module Rbt.Spreadsheet.Recalc

open System.Collections.Generic
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

// ---------------------------------------------------------------------------
// Incremental recalculation of a workbook after any cell edit
// ---------------------------------------------------------------------------
// The engine holds the dependency graph of a calculated workbook: a node per
// formula cell and per cell a formula reads on its own, and the ranges
// formulas read, indexed by row / column so that the ranges holding a cell
// are found without looking at every formula. It is built once per workbook
// and kept up to date by the edits it is shown.
//
// `recalculate` is given the workbook as an edit left it - one cell typed, a
// block pasted or deleted, an undo, a replace-all: any change to the cells of
// any sheets. It finds the cells whose value or formula changed (told, or by
// walking the old and the new cell maps together), rewires the formulas among
// them, collects everything that reads them directly or through other
// formulas, and evaluates those formulas in dependency order - a formula only
// when something it reads did change value. The result is what a full
// recalculation of the edited workbook gives (provided the workbook was
// calculated before the edit).
//
// It declines (None: the caller runs the full recalculation) when the edit
// is not one of cells alone (names, tables, sheets, filters, structure), when
// the formulas to evaluate are part of a cycle, and when the dynamic-array
// pass is needed: a formula that spills is entered or reached, or a cell is
// emptied / a spilled cell overwritten in a workbook that has spills.
// Workbooks whose formulas choose their precedents while they are evaluated
// (OFFSET, INDIRECT), or read through 3-D or spill-range references, are
// always left to the full recalculation. Other volatile functions (NOW,
// RAND ...) are evaluated again at every edit, with what reads them.

/// A range some formulas read, with those formulas.
type RangeGroup(sheet: SheetState, r0: int, c0: int, r1: int, c1: int, key: string) =
    member _.Sheet = sheet
    member _.R0 = r0
    member _.C0 = c0
    member _.R1 = r1
    member _.C1 = c1
    member _.Key = key
    member val Owners = ResizeArray<Node>() with get
    member val Mark = 0 with get, set

/// A formula cell, or a cell a formula reads on its own.
and Node(sheet: SheetState, row: int, col: int) =
    member _.Sheet = sheet
    member _.Row = row
    member _.Col = col
    /// The formula's text and parsed form (null / None: not a formula, or one that does not parse).
    member val Formula : string = null with get, set
    member val Expr : Expr option = None with get, set
    member val IsFormula = false with get, set
    /// The single cells and the ranges the formula reads.
    member val Precedents : Node array = [||] with get, set
    member val Ranges : RangeGroup array = [||] with get, set
    /// The formulas that read this cell on its own.
    member val Dependents = ResizeArray<Node>() with get
    member val Volatile = false with get, set
    member val Spills = false with get, set
    /// The formula's precedents are not the graph's to know (see `Engine.Unsupported`).
    member val Untracked = false with get, set
    // State of the pass in hand.
    member val Pass = 0 with get, set
    member val Pending = 0 with get, set
    member val NeedsEval = false with get, set
    member val Outs : ResizeArray<Node> = null with get, set
    member val EdgeMark = 0 with get, set
    /// The value the cell had before the edit, when the edit changed the cell itself (BaselinePass = the pass).
    member val BaselinePass = 0 with get, set
    member val Baseline : CellValue = CellValue.Empty with get, set

and SheetState(id: SheetId) =
    member _.Id = id
    /// The sheet as last calculated.
    member val Model : SpreadsheetModel = Unchecked.defaultof<SpreadsheetModel> with get, set
    member val Name = "" with get, set
    member val Nodes = Dictionary<float, Node>() with get
    member val Groups = Dictionary<string, RangeGroup>() with get
    /// Ranges by each row (ranges at most as tall as wide) / each column they cover; the rest in Wide.
    member val ByRow = Dictionary<int, ResizeArray<RangeGroup>>() with get
    member val ByCol = Dictionary<int, ResizeArray<RangeGroup>>() with get
    member val Wide = ResizeArray<RangeGroup>() with get

/// What a pass did, for hosts and tests.
type RecalcStats = {
    /// Cells whose value or formula the edit changed.
    ChangedCells: int
    /// Formulas that read them, directly or through other formulas.
    ReachedFormulas: int
    /// Of those, the formulas that were evaluated (something they read changed value).
    EvaluatedFormulas: int
}

type Engine = {
    Sheets: Dictionary<SheetId, SheetState>
    mutable Workbook: Workbook
    /// Formulas whose precedents the graph cannot know (OFFSET / INDIRECT, 3-D and
    /// spill-range references): while there is one, the workbook is left to the
    /// full recalculation.
    mutable UntrackedFormulas: int
    /// Formulas that spill.
    Spilling: ResizeArray<Node>
    Volatile: ResizeArray<Node>
    mutable Pass: int
    mutable EdgeMark: int
    mutable Formulas: int
    mutable LastStats: RecalcStats
    /// Why the last edit was left to the full recalculation.
    mutable LastDecline: string
}

let private keyOf (row: int) (col: int) : float = float row * 16384.0 + float col

/// The largest span (rows or columns, whichever is smaller) a range is indexed by.
let private indexedSpan = 64

let private isLoaded (model: SpreadsheetModel) =
    match model.Deferred with
    | Some deferred -> deferred.Loaded.IsSome
    | None -> true

/// The cells of a sheet without converting a sheet that is not loaded.
let private loadedCells (model: SpreadsheetModel) : Map<CellAddress, CellData> =
    match model.Deferred with
    | Some deferred ->
        match deferred.Loaded with
        | Some loaded -> loaded.Cells
        | None -> Map.empty
    | None -> model.Cells

/// What a function does to the formula that calls it (1: volatile, 2: its
/// precedents are chosen while it is evaluated), by its name as written -
/// asked of Formula once per name.
let private functionTraits = Dictionary<string, int>()
let private traitsOf (name: string) =
    match functionTraits.TryGetValue name with
    | true, known -> known
    | false, _ ->
        let call = Function (name, [])
        let traits = (if containsVolatile call then 1 else 0) ||| (if containsDynamicReference call then 2 else 0)
        if functionTraits.Count > 4096 then functionTraits.Clear()
        functionTraits.[name] <- traits
        traits

/// One walk over a formula: 1 volatile, 2 dynamic references, 4 a spill-range
/// reference, 8 a reference that names a sheet (a 3-D one is looked for then).
let private classify (expr: Expr) : int =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable traits = 0
    while stack.Count > 0 do
        let last = stack.Count - 1
        let current = stack.[last]
        stack.RemoveAt last
        match current with
        | SpillRef r ->
            traits <- traits ||| 4
            if r.Sheet.IsSome then traits <- traits ||| 8
        | CellRef r -> if r.Sheet.IsSome then traits <- traits ||| 8
        | RangeRef (a, _) -> if a.Sheet.IsSome then traits <- traits ||| 8
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | Function (name, args) ->
            traits <- traits ||| traitsOf name
            for arg in args do
                stack.Add arg
        | _ -> ()
    traits

let private sheetState (engine: Engine) (sheetId: SheetId) : SheetState =
    match engine.Sheets.TryGetValue sheetId with
    | true, state -> state
    | false, _ ->
        let state = SheetState sheetId
        engine.Sheets.[sheetId] <- state
        state

let private nodeAt (sheet: SheetState) (row: int) (col: int) : Node =
    let key = keyOf row col
    match sheet.Nodes.TryGetValue key with
    | true, node -> node
    | false, _ ->
        let node = Node(sheet, row, col)
        sheet.Nodes.[key] <- node
        node

let private addToBucket (index: Dictionary<int, ResizeArray<RangeGroup>>) (key: int) (group: RangeGroup) =
    match index.TryGetValue key with
    | true, groups -> groups.Add group
    | false, _ ->
        let groups = ResizeArray<RangeGroup>()
        groups.Add group
        index.[key] <- groups

let private removeFromBucket (index: Dictionary<int, ResizeArray<RangeGroup>>) (key: int) (group: RangeGroup) =
    match index.TryGetValue key with
    | true, groups ->
        groups.Remove group |> ignore
        if groups.Count = 0 then index.Remove key |> ignore
    | false, _ -> ()

let private groupAt (sheet: SheetState) (r0: int) (c0: int) (r1: int) (c1: int) : RangeGroup =
    let key = string r0 + ":" + string c0 + ":" + string r1 + ":" + string c1
    match sheet.Groups.TryGetValue key with
    | true, group -> group
    | false, _ ->
        let group = RangeGroup(sheet, r0, c0, r1, c1, key)
        sheet.Groups.[key] <- group
        let rows = r1 - r0 + 1
        let cols = c1 - c0 + 1
        if rows <= cols && rows <= indexedSpan then
            for row in r0 .. r1 do addToBucket sheet.ByRow row group
        elif cols <= indexedSpan then
            for col in c0 .. c1 do addToBucket sheet.ByCol col group
        elif rows <= indexedSpan then
            for row in r0 .. r1 do addToBucket sheet.ByRow row group
        else
            sheet.Wide.Add group
        group

let private dropGroup (sheet: SheetState) (group: RangeGroup) =
    sheet.Groups.Remove group.Key |> ignore
    let rows = group.R1 - group.R0 + 1
    let cols = group.C1 - group.C0 + 1
    if rows <= cols && rows <= indexedSpan then
        for row in group.R0 .. group.R1 do removeFromBucket sheet.ByRow row group
    elif cols <= indexedSpan then
        for col in group.C0 .. group.C1 do removeFromBucket sheet.ByCol col group
    elif rows <= indexedSpan then
        for row in group.R0 .. group.R1 do removeFromBucket sheet.ByRow row group
    else
        sheet.Wide.Remove group |> ignore

/// Take a formula out of the graph: what it read no longer leads to it.
let private unwire (engine: Engine) (node: Node) =
    for precedent in node.Precedents do
        precedent.Dependents.Remove node |> ignore
        if not precedent.IsFormula && precedent.Dependents.Count = 0 then
            precedent.Sheet.Nodes.Remove (keyOf precedent.Row precedent.Col) |> ignore
    for group in node.Ranges do
        group.Owners.Remove node |> ignore
        // A range no formula reads any more leaves the index.
        if group.Owners.Count = 0 then dropGroup group.Sheet group
    if node.Volatile then engine.Volatile.Remove node |> ignore
    if node.Spills then engine.Spilling.Remove node |> ignore
    if node.Untracked then engine.UntrackedFormulas <- engine.UntrackedFormulas - 1
    if node.IsFormula then engine.Formulas <- engine.Formulas - 1
    node.Precedents <- [||]
    node.Ranges <- [||]
    node.Formula <- null
    node.Expr <- None
    node.IsFormula <- false
    node.Volatile <- false
    node.Spills <- false
    node.Untracked <- false
    if node.Dependents.Count = 0 then
        node.Sheet.Nodes.Remove (keyOf node.Row node.Col) |> ignore

/// Put a formula into the graph.
let private wire (engine: Engine) (workbook: Workbook) (sheet: SheetState) (row: int) (col: int) (formula: string) : Node =
    let node = nodeAt sheet row col
    node.Formula <- formula
    node.IsFormula <- true
    engine.Formulas <- engine.Formulas + 1
    match parseFormula formula with
    | Ok expr ->
        node.Expr <- Some expr
        let precedents = ResizeArray<Node>()
        let ranges = ResizeArray<RangeGroup>()
        // (A cell or a range the formula names twice is one precedent.)
        engine.EdgeMark <- engine.EdgeMark + 1
        let mark = engine.EdgeMark
        iterReferenceAreas (Some workbook) sheet.Id (Some { Row = row; Col = col }) expr (fun sheetId r0 c0 r1 c1 ->
            let target = sheetState engine sheetId
            if r0 = r1 && c0 = c1 then
                let precedent = nodeAt target r0 c0
                if precedent.EdgeMark <> mark then
                    precedent.EdgeMark <- mark
                    precedents.Add precedent
                    precedent.Dependents.Add node
            else
                let group = groupAt target r0 c0 r1 c1
                if group.Mark <> mark then
                    group.Mark <- mark
                    ranges.Add group
                    group.Owners.Add node)
        node.Precedents <- precedents.ToArray()
        node.Ranges <- ranges.ToArray()
        let traits = classify expr
        if traits &&& 1 <> 0 then
            node.Volatile <- true
            engine.Volatile.Add node
        if exprSpillsStatically expr then
            node.Spills <- true
            engine.Spilling.Add node
        if traits &&& 2 <> 0 || traits &&& 4 <> 0 || (traits &&& 8 <> 0 && containsThreeDReference expr) then
            node.Untracked <- true
            engine.UntrackedFormulas <- engine.UntrackedFormulas + 1
    | Result.Error _ ->
        node.Expr <- None
    node

let private noStats = { ChangedCells = 0; ReachedFormulas = 0; EvaluatedFormulas = 0 }

/// Build the dependency graph of a workbook. The formulas' cached values are
/// taken as they are: the workbook is the calculated state edits start from.
let build (workbook: Workbook) : Engine =
    let engine = {
        Sheets = Dictionary<SheetId, SheetState>()
        Workbook = workbook
        UntrackedFormulas = 0
        Spilling = ResizeArray<Node>()
        Volatile = ResizeArray<Node>()
        Pass = 0
        EdgeMark = 0
        Formulas = 0
        LastStats = noStats
        LastDecline = ""
    }
    for KeyValue (sheetId, sheet) in workbook.Sheets do
        let state = sheetState engine sheetId
        state.Model <- sheet.Model
        state.Name <- sheet.Name
    for KeyValue (sheetId, sheet) in workbook.Sheets do
        let state = sheetState engine sheetId
        match sheet.Model.Deferred with
        | Some deferred when deferred.Loaded.IsNone ->
            // A sheet nothing has asked for yet: its formulas are read off the
            // file. It is converted when one of them has to be evaluated.
            for addr, formula in deferred.Formulas () do
                wire engine workbook state addr.Row addr.Col formula |> ignore
        | _ ->
            for KeyValue (addr, cell) in loadedCells sheet.Model do
                match cell.Value with
                | CellValue.Formula (formula, _) -> wire engine workbook state addr.Row addr.Col formula |> ignore
                | _ -> ()
    engine

/// Number of formula cells in the graph.
let formulaCount (engine: Engine) = engine.Formulas

/// True while the workbook holds a formula whose precedents the graph cannot
/// know (OFFSET / INDIRECT, a 3-D or a spill-range reference): every edit is
/// left to the full recalculation.
let isUnsupported (engine: Engine) = engine.UntrackedFormulas > 0

let lastStats (engine: Engine) = engine.LastStats

/// Why `recalculate` last declined an edit.
let lastDecline (engine: Engine) = engine.LastDecline

/// The formulas that read a cell: on its own, or in a range.
let private addReaders (engine: Engine) (sheet: SheetState) (row: int) (col: int) (explicit: ResizeArray<Node>) (into: ResizeArray<Node>) =
    engine.EdgeMark <- engine.EdgeMark + 1
    let mark = engine.EdgeMark
    let add (reader: Node) =
        if reader.EdgeMark <> mark then
            reader.EdgeMark <- mark
            into.Add reader
    if not (isNull explicit) then
        for reader in explicit do add reader
    let inGroups (groups: ResizeArray<RangeGroup>) =
        for group in groups do
            if row >= group.R0 && row <= group.R1 && col >= group.C0 && col <= group.C1 then
                for owner in group.Owners do add owner
    match sheet.ByRow.TryGetValue row with
    | true, groups -> inGroups groups
    | false, _ -> ()
    match sheet.ByCol.TryGetValue col with
    | true, groups -> inGroups groups
    | false, _ -> ()
    if sheet.Wide.Count > 0 then inGroups sheet.Wide

let private effectiveValue (cell: CellData voption) : CellValue =
    match cell with
    | ValueSome data -> unwrapFormulaValue data.Value
    | ValueNone -> CellValue.Empty

let private formulaOf (cell: CellData voption) : string =
    match cell with
    | ValueSome data ->
        match data.Value with
        | CellValue.Formula (formula, _) -> formula
        | _ -> null
    | ValueNone -> null

/// A cell the edit changed: what it held and what it holds.
[<Struct>]
type private CellChange = {
    State: SheetState
    Row: int
    Col: int
    Before: CellData voption
    After: CellData voption
}

/// The cells of a sheet that differ between two cell maps, by walking both in order.
let private diffCells (state: SheetState) (before: Map<CellAddress, CellData>) (after: Map<CellAddress, CellData>) (changes: ResizeArray<CellChange>) =
    use left = (before :> seq<KeyValuePair<CellAddress, CellData>>).GetEnumerator()
    use right = (after :> seq<KeyValuePair<CellAddress, CellData>>).GetEnumerator()
    let mutable hasLeft = left.MoveNext()
    let mutable hasRight = right.MoveNext()
    while hasLeft || hasRight do
        let order =
            if not hasLeft then 1
            elif not hasRight then -1
            else
                let a = left.Current.Key
                let b = right.Current.Key
                if a.Row <> b.Row then (if a.Row < b.Row then -1 else 1)
                elif a.Col <> b.Col then (if a.Col < b.Col then -1 else 1)
                else 0
        if order = 0 then
            let was = left.Current.Value
            let now = right.Current.Value
            if not (obj.ReferenceEquals(was, now)) && not (obj.ReferenceEquals(was.Value, now.Value)) && was.Value <> now.Value then
                changes.Add { State = state; Row = left.Current.Key.Row; Col = left.Current.Key.Col; Before = ValueSome was; After = ValueSome now }
            hasLeft <- left.MoveNext()
            hasRight <- right.MoveNext()
        elif order < 0 then
            let was = left.Current.Value
            if was.Value <> CellValue.Empty then
                changes.Add { State = state; Row = left.Current.Key.Row; Col = left.Current.Key.Col; Before = ValueSome was; After = ValueNone }
            hasLeft <- left.MoveNext()
        else
            let now = right.Current.Value
            if now.Value <> CellValue.Empty then
                changes.Add { State = state; Row = right.Current.Key.Row; Col = right.Current.Key.Col; Before = ValueNone; After = ValueSome now }
            hasRight <- right.MoveNext()

let private sameFilterState (before: SpreadsheetModel) (after: SpreadsheetModel) =
    (obj.ReferenceEquals(before.Filters, after.Filters) || (Map.isEmpty before.Filters && Map.isEmpty after.Filters))
    && (obj.ReferenceEquals(before.FilterHiddenRows, after.FilterHiddenRows) || (before.FilterHiddenRows.Length = 0 && after.FilterHiddenRows.Length = 0))

/// Whether the edit left everything a formula can read, apart from cells, as it was.
let private sameSurroundings (engine: Engine) (edited: Workbook) =
    let known = engine.Workbook
    (obj.ReferenceEquals(known.Tables, edited.Tables) || known.Tables = edited.Tables)
    && (obj.ReferenceEquals(known.Names, edited.Names) || known.Names = edited.Names)
    && (obj.ReferenceEquals(known.SheetOrder, edited.SheetOrder) || known.SheetOrder = edited.SheetOrder)
    && known.Sheets.Count = edited.Sheets.Count
    && (edited.Sheets |> Map.forall (fun sheetId sheet ->
            match engine.Sheets.TryGetValue sheetId with
            | true, state when not (isNull (box state.Model)) ->
                let before = state.Model
                let after = sheet.Model
                obj.ReferenceEquals(before, after)
                || (state.Name = sheet.Name
                    && before.StructureVersion = after.StructureVersion
                    && (obj.ReferenceEquals(before.Spill, after.Spill) || (Map.isEmpty before.Spill && Map.isEmpty after.Spill))
                    && (obj.ReferenceEquals(before.Tables, after.Tables) || before.Tables = after.Tables)
                    && sameFilterState before after)
            | _ -> false))

/// The cells an edit changed: those named by `hint`, or - without one - the
/// cells of every sheet whose cell map is another one, compared cell by cell.
let private collectChanges (engine: Engine) (edited: Workbook) (hint: (SheetId * CellAddress) list option) : ResizeArray<CellChange> =
    let changes = ResizeArray<CellChange>()
    match hint with
    | Some cells ->
        for sheetId, addr in cells do
            match engine.Sheets.TryGetValue sheetId, Map.tryFind sheetId edited.Sheets with
            | (true, state), Some sheet ->
                let before = match Map.tryFind addr (loadedCells state.Model) with Some cell -> ValueSome cell | None -> ValueNone
                let after = match Map.tryFind addr (loadedCells sheet.Model) with Some cell -> ValueSome cell | None -> ValueNone
                let differs =
                    match before, after with
                    | ValueSome was, ValueSome now -> not (obj.ReferenceEquals(was, now)) && was.Value <> now.Value
                    | ValueNone, ValueNone -> false
                    | ValueSome was, ValueNone -> was.Value <> CellValue.Empty
                    | ValueNone, ValueSome now -> now.Value <> CellValue.Empty
                if differs then changes.Add { State = state; Row = addr.Row; Col = addr.Col; Before = before; After = after }
            | _ -> ()
    | None ->
        for KeyValue (sheetId, sheet) in edited.Sheets do
            let state = engine.Sheets.[sheetId]
            let before = state.Model
            let after = sheet.Model
            if not (obj.ReferenceEquals(before, after)) then
                // A sheet that was a stand-in and is loaded now holds what its file held.
                let beforeCells = if isLoaded before then loadedCells before else loadedCells after
                let afterCells = loadedCells after
                if not (obj.ReferenceEquals(beforeCells, afterCells)) then
                    diffCells state beforeCells afterCells changes
    changes

/// Make the graph follow the formulas among the changed cells. `onFormula`
/// is given each changed cell that holds a formula now, `onValue` each one
/// that does not.
let private rewire (engine: Engine) (edited: Workbook) (changes: ResizeArray<CellChange>) (onFormula: Node -> CellChange -> unit) (onValue: CellChange -> unit) =
    for change in changes do
        let state = change.State
        let formulaBefore = formulaOf change.Before
        let formulaAfter = formulaOf change.After
        let key = keyOf change.Row change.Col
        if not (isNull formulaBefore) && (isNull formulaAfter || formulaAfter <> formulaBefore) then
            match state.Nodes.TryGetValue key with
            | true, node when node.IsFormula -> unwire engine node
            | _ -> ()
        if not (isNull formulaAfter) then
            let node =
                match state.Nodes.TryGetValue key with
                | true, node when node.IsFormula && node.Formula = formulaAfter -> node
                | true, node when node.IsFormula ->
                    unwire engine node
                    wire engine edited state change.Row change.Col formulaAfter
                | _ -> wire engine edited state change.Row change.Col formulaAfter
            onFormula node change
        else
            onValue change

let private holdWorkbook (engine: Engine) (workbook: Workbook) =
    for KeyValue (sheetId, sheet) in workbook.Sheets do
        let state = sheetState engine sheetId
        state.Model <- sheet.Model
        state.Name <- sheet.Name
    engine.Workbook <- workbook

/// Take a workbook that was recalculated elsewhere (the full recalculation,
/// after `recalculate` declined an edit) as the engine's calculated state:
/// the graph follows the formulas that differ from the state it held. False
/// when more than cells differ - the engine is then to be built again.
let adopt (engine: Engine) (calculated: Workbook) : bool =
    if not (sameSurroundings engine calculated) then false
    else
        let changes = collectChanges engine calculated None
        rewire engine calculated changes (fun _ _ -> ()) ignore
        holdWorkbook engine calculated
        true

/// Recalculate the workbook after an edit of cells. `edited` is the engine's
/// workbook with the edit applied and nothing recalculated. `hint` names the
/// cells the edit touched when the caller knows them (a typed entry); without
/// it the sheets whose cells are another map are compared cell by cell.
/// Some: the recalculated workbook - the engine now holds it. None: the edit
/// needs the full recalculation; afterwards the engine is shown its result
/// (`adopt`) or built again.
let recalculate (engine: Engine) (edited: Workbook) (hint: (SheetId * CellAddress) list option) : Workbook option =
    let decline (reason: string) : Workbook option =
        engine.LastDecline <- reason
        None
    if engine.UntrackedFormulas > 0 then decline "a formula chooses its precedents while it is evaluated (OFFSET, INDIRECT) or reads a 3-D or spill-range reference"
    elif not (sameSurroundings engine edited) then decline "more than cells changed (sheets, names, tables, filters, rows or columns)"
    else
        let changes = collectChanges engine edited hint
        // What the dynamic-array pass has to see: a spilled cell overwritten,
        // and - while a spill is blocked (#SPILL!) - a cell emptied, which may free it.
        let hasSpills =
            engine.Spilling.Count > 0
            || (edited.Sheets |> Map.exists (fun _ sheet -> not (Map.isEmpty sheet.Model.Spill)))
        let mutable needsFull = false
        if hasSpills then
            // (A blocked spill runs right and down from its formula: a cell emptied there may free it.)
            let blocked =
                engine.Spilling |> Seq.filter (fun node ->
                    match Map.tryFind { Row = node.Row; Col = node.Col } (loadedCells node.Sheet.Model) with
                    | Some cell -> unwrapFormulaValue cell.Value = CellValue.Error Err.spill
                    | None -> false)
                |> Seq.toArray
            let mayFree (change: CellChange) =
                blocked |> Array.exists (fun node -> obj.ReferenceEquals(node.Sheet, change.State) && change.Row >= node.Row && change.Col >= node.Col)
            for change in changes do
                if not needsFull then
                    let spill = change.State.Model.Spill
                    // (A spilling formula changed or removed: its spilled cells go with it.)
                    let spilling =
                        match change.State.Nodes.TryGetValue (keyOf change.Row change.Col) with
                        | true, node -> node.Spills
                        | false, _ -> false
                    if spilling then needsFull <- true
                    elif blocked.Length > 0 && effectiveValue change.After = CellValue.Empty && isNull (formulaOf change.After) && mayFree change then needsFull <- true
                    elif not (Map.isEmpty spill) && Map.containsKey { Row = change.Row; Col = change.Col } spill then needsFull <- true
        if needsFull then decline "the edit touches a spill (a spilling formula, a spilled cell, or a cell that may free a blocked spill)"
        else
            engine.Pass <- engine.Pass + 1
            let pass = engine.Pass
            let reached = ResizeArray<Node>()
            let pending = Stack<Node>()
            let reach (node: Node) =
                if node.Pass <> pass then
                    node.Pass <- pass
                    node.Pending <- 0
                    node.NeedsEval <- false
                    node.Outs <- null
                    reached.Add node
                    pending.Push node
            let readers = ResizeArray<Node>()
            rewire engine edited changes
                (fun node change ->
                    reach node
                    node.NeedsEval <- true
                    node.BaselinePass <- pass
                    node.Baseline <- effectiveValue change.Before)
                (fun change ->
                    // What reads the cell sees another value.
                    let explicit =
                        match change.State.Nodes.TryGetValue (keyOf change.Row change.Col) with
                        | true, node -> node.Dependents
                        | false, _ -> null
                    readers.Clear()
                    addReaders engine change.State change.Row change.Col explicit readers
                    for reader in readers do
                        reach reader
                        reader.NeedsEval <- true)
            if engine.UntrackedFormulas > 0 then decline "a formula whose precedents are not known beforehand was entered"
            else
                for node in engine.Volatile do
                    reach node
                    node.NeedsEval <- true
                // Everything that reads what was reached, with the edges among them.
                while pending.Count > 0 do
                    let node = pending.Pop()
                    let outs = ResizeArray<Node>()
                    addReaders engine node.Sheet node.Row node.Col node.Dependents outs
                    node.Outs <- outs
                    for reader in outs do
                        reach reader
                        reader.Pending <- reader.Pending + 1
                let mutable spills = false
                for node in reached do
                    if node.Spills then spills <- true
                if spills then decline "the edit reaches a spilling formula"
                else
                    // Dependency order: a formula is ready when everything reached that it reads is done.
                    let ready = Stack<Node>()
                    for node in reached do
                        if node.Pending = 0 then ready.Push node
                    let order = ResizeArray<Node>(reached.Count)
                    while ready.Count > 0 do
                        let node = ready.Pop()
                        order.Add node
                        for reader in node.Outs do
                            reader.Pending <- reader.Pending - 1
                            if reader.Pending = 0 then ready.Push reader
                    // (Fewer in order than reached: a cycle among the formulas to evaluate.)
                    if order.Count < reached.Count then decline "a circular reference among the formulas to evaluate"
                    else
                        let mutable workbook = edited
                        let models = Dictionary<SheetId, SpreadsheetModel>()
                        let touched = HashSet<SheetId>()
                        let modelOf (state: SheetState) =
                            match models.TryGetValue state.Id with
                            | true, model -> model
                            | false, _ ->
                                // (A formula is evaluated in its sheet, not in the sheet's stand-in.)
                                let model = resolveSheet (Map.find state.Id workbook.Sheets).Model
                                models.[state.Id] <- model
                                model
                        let mutable evaluated = 0
                        for node in order do
                            if node.NeedsEval && node.IsFormula then
                                evaluated <- evaluated + 1
                                let state = node.Sheet
                                let model = modelOf state
                                let addr = { Row = node.Row; Col = node.Col }
                                match Map.tryFind addr model.Cells with
                                | Some cell ->
                                    let result =
                                        match node.Expr with
                                        | Some expr -> eval expr { Model = model; Workbook = Some workbook; Sheet = Some state.Id; CurrentCell = Some addr }
                                        | None -> CellValue.Error Err.syntax
                                    let previous = match cell.Value with CellValue.Formula (_, cached) -> cached | other -> other
                                    if previous <> result then
                                        let updated = { model with Cells = Map.add addr { cell with Value = CellValue.Formula (node.Formula, result) } model.Cells }
                                        models.[state.Id] <- updated
                                        touched.Add state.Id |> ignore
                                        workbook <- updateSheetModel state.Id updated workbook
                                    // What reads it was calculated with the value it had before the edit.
                                    let before = if node.BaselinePass = pass then node.Baseline else previous
                                    if before <> result then
                                        for reader in node.Outs do
                                            reader.NeedsEval <- true
                                | None -> ()
                        // As the full recalculation leaves a sheet whose formulas gave other values.
                        for sheetId in touched do
                            let model = models.[sheetId]
                            let projected = invalidateFilterProjection model
                            if not (obj.ReferenceEquals(model, projected)) then
                                workbook <- updateSheetModel sheetId projected workbook
                        holdWorkbook engine workbook
                        engine.LastStats <- { ChangedCells = changes.Count; ReachedFormulas = reached.Count; EvaluatedFormulas = evaluated }
                        Some workbook
