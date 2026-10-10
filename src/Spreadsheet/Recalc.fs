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
// An edit can be more than one of cells: `recalculate` compares what the
// engine holds with the edited workbook and has the formulas that can tell
// wired again and calculated - those that use a defined name that changed,
// structured references when a table changed, SUBTOTAL / AGGREGATE when a
// filter or hidden rows changed, what reads a sheet that was removed, renamed
// or whose rows or columns moved, 3-D references when the sheets or their
// order changed (see `Surroundings`). Rows or columns inserted or deleted are
// known as such (`Model.StructureMove`): a formula that only followed its
// cells is wired again and keeps its value (`followsMove`).
//
// A circular reference the edit makes or reaches is found among the formulas
// the pass has to calculate (their strongly connected components): its
// formulas show the cycle error, or - in a workbook with iterative
// calculation - are calculated in sweeps, in reading order. Volatile
// functions (NOW, RAND ...) are evaluated again at every edit, with what
// reads them. A 3-D reference reads its range on every sheet of its span.
// Formulas that choose their precedents while they are evaluated (OFFSET,
// INDIRECT, the range and intersection operators) are evaluated at every
// edit too, and once more whenever a value changed after they were: see
// `recalculate`. Dynamic arrays are part of the graph: see there.
//
// It declines (None: the caller runs the full recalculation) when the
// workbook's setting of iterative calculation changed, and when dynamic
// arrays or OFFSET / INDIRECT formulas do not settle (a circular reference
// through them).
//
// The full recalculation of a workbook (`recalculateAll`, what
// Formula.recalculateWorkbook runs) is a pass of the same graph over every
// formula: one dependency order for the whole workbook, so one run reaches
// the values however they travel between sheets, and the formulas on a
// circular reference - also one that crosses sheets - show the cycle error.

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
    /// The batch of formulas taken out of the graph that last swept the owners (`unwireMany`).
    member val SweepMark = 0 with get, set

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
    /// The formula chooses precedents while it is evaluated (OFFSET, INDIRECT,
    /// a range or intersection operator): the graph knows only the references it
    /// spells out.
    member val Dynamic = false with get, set
    /// The formula gives another value every time (RAND, NOW ...).
    member val Random = false with get, set
    /// The formula reads a 3-D reference (`Jan:Mar!B1`), itself or through a name.
    member val ThreeD = false with get, set
    /// What the formula is made of (`classify`): 1 a function that gives another value every time,
    /// 2 references chosen while it is evaluated, 4 a 3-D reference, 8 a function that leaves out
    /// hidden rows (SUBTOTAL, AGGREGATE), 16 a defined name, 32 a structured reference, 64 a
    /// function of the workbook's sheets (SHEET, SHEETS), 128 a reference to a sheet there is none of.
    member val Traits = 0 with get, set
    /// Which wiring of a formula this is (`Engine.WireCount` when it was wired).
    member val WireStamp = 0 with get, set
    /// The batch that takes the formula out of the graph / that last swept what reads this cell (`unwireMany`).
    member val UnwireMark = 0 with get, set
    member val SweepMark = 0 with get, set
    /// Found on a circular reference by the last pass that reached it.
    member val OnCycle = false with get, set
    /// The round of a pass that last found it on a circular reference.
    member val CircularRound = 0 with get, set
    /// With iterative calculation: the circular reference it is on did not settle when it was
    /// last calculated (or has not been calculated here yet) - it is calculated at every edit.
    member val Unsettled = true with get, set
    // Strongly connected components of a round (the full pass marks cycles).
    member val SccPass = 0 with get, set
    member val SccIndex = 0 with get, set
    member val SccLow = 0 with get, set
    member val SccOn = false with get, set
    /// The round in which the formula was seen to read itself.
    member val SelfPass = 0 with get, set
    /// The first round of the full pass that found the formula on a circular reference.
    member val CyclicPass = 0 with get, set
    // State of the pass in hand.
    member val Pass = 0 with get, set
    member val Pending = 0 with get, set
    member val NeedsEval = false with get, set
    member val Outs : ResizeArray<Node> = null with get, set
    /// The cell's dynamic array has to be laid out again (its spilled cells written, cleared or blocked).
    member val NeedsSpill = false with get, set
    /// The node's place in the order of the pass in hand.
    member val Order = -1 with get, set
    /// Rows and columns of the array the formula last gave here (-1: not known).
    member val ExtentRows = -1 with get, set
    member val ExtentCols = -1 with get, set
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
    /// The cells each dynamic array has spilled into, by the key of its formula's
    /// cell: the sheet's spill map (`GhostsFor`) the other way round.
    member val Ghosts = Dictionary<float, ResizeArray<CellAddress>>() with get
    member val GhostsFor : obj = null with get, set

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
    /// Formulas that spill.
    Spilling: ResizeArray<Node>
    Volatile: ResizeArray<Node>
    /// Formulas that choose precedents while they are evaluated.
    Dynamic: ResizeArray<Node>
    mutable Pass: int
    mutable EdgeMark: int
    mutable Formulas: int
    mutable LastStats: RecalcStats
    /// Why the last edit was left to the full recalculation.
    mutable LastDecline: string
    /// Formulas the last full pass found on a circular reference.
    mutable LastCyclic: int
    /// What the defined names bring to a formula that uses them (see `classify`),
    /// by sheet and name: the names are those the graph was built with.
    NameTraits: Dictionary<string, int>
    /// Formulas that read a 3-D reference: what they read depends on the order of the sheets.
    mutable ThreeD: int
    /// Formulas found on a circular reference (`Node.OnCycle`; entries whose flag is down are stale).
    /// With iterative calculation they are calculated again at every edit, as in Excel.
    mutable WireCount: int
    mutable Cyclic: ResizeArray<Node>
    /// Every circular reference of the graph is in `Cyclic` (a pass over all of it has looked).
    mutable CyclesKnown: bool
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

/// What a function does to the formula that calls it (1: it gives another
/// value every time - RAND, NOW; 2: its precedents are chosen while it is
/// evaluated - OFFSET, INDIRECT, the range and intersection operators), by
/// its name as written - asked of Formula once per name.
let private functionTraits = Dictionary<string, int>()
let private traitsOf (name: string) =
    match functionTraits.TryGetValue name with
    | true, known -> known
    | false, _ ->
        let call = Function (name, [])
        let traits = if containsDynamicReference call then 2 elif containsVolatile call then 1 else 0
        let upper = name.ToUpperInvariant()
        let traits =
            traits
            ||| (if upper = "SUBTOTAL" || upper = "AGGREGATE" then 8 else 0)
            ||| (if upper = "SHEET" || upper = "SHEETS" then 64 else 0)
        if functionTraits.Count > 4096 then functionTraits.Clear()
        functionTraits.[name] <- traits
        traits

/// One walk over a formula and the defined names it uses: see `Node.Traits`.
let rec private classifyWith (engine: Engine) (workbook: Workbook) (sheetId: SheetId) (seen: Set<string>) (expr: Expr) : int =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable traits = 0
    while stack.Count > 0 do
        let last = stack.Count - 1
        let current = stack.[last]
        stack.RemoveAt last
        match current with
        | CellRef r | SpillRef r | RangeRef (r, _) ->
            match r.Sheet with
            | Some qualifier when qualifier.IndexOf ':' >= 0 -> traits <- traits ||| 4
            | Some qualifier -> if (resolveSheetName qualifier workbook).IsNone then traits <- traits ||| 128
            | None -> ()
        | TableRef _ -> traits <- traits ||| 32
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | Function (name, args) ->
            traits <- traits ||| traitsOf name
            for arg in args do
                stack.Add arg
        | NameRef name ->
            traits <- traits ||| 16
            let upper = name.ToUpperInvariant()
            let key = sheetId + "|" + upper
            match engine.NameTraits.TryGetValue key with
            | true, known -> traits <- traits ||| known
            | false, _ ->
                if not (Set.contains upper seen) then
                    let ofName =
                        match resolveNameExpr name (Some sheetId) workbook with
                        | Some resolved -> classifyWith engine workbook sheetId (Set.add upper seen) resolved
                        | None -> 0
                    // (A name met inside its own definition is not settled yet.)
                    if Set.isEmpty seen then engine.NameTraits.[key] <- ofName
                    traits <- traits ||| ofName
        | _ -> ()
    traits

let private classify (engine: Engine) (workbook: Workbook) (sheetId: SheetId) (expr: Expr) : int =
    classifyWith engine workbook sheetId Set.empty expr

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

let private noNodes : Node array = [||]
let private noGroups : RangeGroup array = [||]

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
    if node.Dynamic then engine.Dynamic.Remove node |> ignore
    if node.ThreeD then engine.ThreeD <- engine.ThreeD - 1
    if node.IsFormula then engine.Formulas <- engine.Formulas - 1
    node.Precedents <- [||]
    node.Ranges <- [||]
    node.Formula <- null
    node.Expr <- None
    node.IsFormula <- false
    node.Volatile <- false
    node.Spills <- false
    node.Dynamic <- false
    node.Random <- false
    node.ThreeD <- false
    node.Traits <- 0
    node.OnCycle <- false
    node.Unsettled <- true
    if node.Dependents.Count = 0 then
        node.Sheet.Nodes.Remove (keyOf node.Row node.Col) |> ignore

/// Take many formulas out of the graph at once. A cell that thousands of
/// formulas read (the cell of a defined name, a total) holds them in a list:
/// taking them out one by one searches that list once for each, so each list
/// is gone through once instead, for the whole batch.
let private unwireMany (engine: Engine) (nodes: ResizeArray<Node>) =
    if nodes.Count <= 4 then
        for node in nodes do
            if node.IsFormula then unwire engine node
    else
        engine.EdgeMark <- engine.EdgeMark + 1
        let batch = engine.EdgeMark
        for node in nodes do
            if node.IsFormula then node.UnwireMark <- batch
        for node in nodes do
            if node.IsFormula && node.UnwireMark = batch then
                for precedent in node.Precedents do
                    if precedent.SweepMark <> batch then
                        precedent.SweepMark <- batch
                        precedent.Dependents.RemoveAll (fun reader -> reader.UnwireMark = batch) |> ignore
                        if not precedent.IsFormula && precedent.Dependents.Count = 0 then
                            precedent.Sheet.Nodes.Remove (keyOf precedent.Row precedent.Col) |> ignore
                for group in node.Ranges do
                    if group.SweepMark <> batch then
                        group.SweepMark <- batch
                        group.Owners.RemoveAll (fun owner -> owner.UnwireMark = batch) |> ignore
                        if group.Owners.Count = 0 then dropGroup group.Sheet group
                node.Precedents <- noNodes
                node.Ranges <- noGroups
        // (What is left of each: as `unwire` leaves a formula that reads nothing.)
        for node in nodes do
            if node.IsFormula && node.UnwireMark = batch then
                node.UnwireMark <- 0
                unwire engine node

/// Put a formula into the graph.
let private wire (engine: Engine) (workbook: Workbook) (sheet: SheetState) (row: int) (col: int) (formula: string) : Node =
    let node = nodeAt sheet row col
    node.Formula <- formula
    node.IsFormula <- true
    engine.WireCount <- engine.WireCount + 1
    node.WireStamp <- engine.WireCount
    engine.Formulas <- engine.Formulas + 1
    match parseFormula formula with
    | Ok expr ->
        node.Expr <- Some expr
        let precedents = ResizeArray<Node>()
        let ranges = ResizeArray<RangeGroup>()
        // (A cell or a range the formula names twice is one precedent.)
        engine.EdgeMark <- engine.EdgeMark + 1
        let mark = engine.EdgeMark
        // (Most references are to the formula's own sheet, or to the sheet of the reference before.)
        let lastTarget = ref sheet
        iterReferenceAreas (Some workbook) sheet.Id (Some { Row = row; Col = col }) expr (fun sheetId r0 c0 r1 c1 ->
            if not (obj.ReferenceEquals(sheetId, lastTarget.Value.Id)) && sheetId <> lastTarget.Value.Id then lastTarget.Value <- sheetState engine sheetId
            let target = lastTarget.Value
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
        node.Precedents <- (if precedents.Count = 0 then noNodes else precedents.ToArray())
        node.Ranges <- (if ranges.Count = 0 then noGroups else ranges.ToArray())
        let traits = classify engine workbook sheet.Id expr
        node.Traits <- traits
        if traits &&& 3 <> 0 then
            node.Volatile <- true
            engine.Volatile.Add node
        if traits &&& 4 <> 0 then
            node.ThreeD <- true
            engine.ThreeD <- engine.ThreeD + 1
        node.Random <- traits &&& 1 <> 0
        if traits &&& 2 <> 0 then
            node.Dynamic <- true
            engine.Dynamic.Add node
        // (OFFSET / INDIRECT may give a range: the benefit of the doubt, as in
        // the layout of the arrays. One cell is laid out as no array.)
        if (if node.Dynamic then exprSpills expr else exprSpillsStatically expr) then
            node.Spills <- true
            engine.Spilling.Add node
        // (A spill-range reference `A1#` is a reference to its formula's cell: what
        // reads it is evaluated whenever the array is laid out differently.)
    | Result.Error _ ->
        node.Expr <- None
    node

let private noStats = { ChangedCells = 0; ReachedFormulas = 0; EvaluatedFormulas = 0 }

let private buildGraph (unloadedToo: bool) (workbook: Workbook) : Engine =
    let engine = {
        Sheets = Dictionary<SheetId, SheetState>()
        Workbook = workbook
        Spilling = ResizeArray<Node>()
        Volatile = ResizeArray<Node>()
        Dynamic = ResizeArray<Node>()
        Pass = 0
        EdgeMark = 0
        Formulas = 0
        LastStats = noStats
        LastDecline = ""
        LastCyclic = 0
        NameTraits = Dictionary<string, int>()
        ThreeD = 0
        WireCount = 0
        Cyclic = ResizeArray<Node>()
        CyclesKnown = false
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
            if unloadedToo then
                for addr, formula in deferred.Formulas () do
                    wire engine workbook state addr.Row addr.Col formula |> ignore
        | _ ->
            // (The sheet's formula index: its values are not walked to find its formulas.)
            let cells = loadedCells sheet.Model
            for addr in formulaIndexOf cells do
                match Map.tryFind addr cells with
                | Some { Value = CellValue.Formula (formula, _) } -> wire engine workbook state addr.Row addr.Col formula |> ignore
                | _ -> ()
    engine

/// Build the dependency graph of a workbook. The formulas' cached values are
/// taken as they are: the workbook is the calculated state edits start from.
let build (workbook: Workbook) : Engine = buildGraph true workbook

/// Number of formula cells in the graph.
let formulaCount (engine: Engine) = engine.Formulas

/// Whether the workbook holds formulas the graph cannot follow, so that every
/// edit is left to the full recalculation. No formula is of that kind any
/// more (OFFSET / INDIRECT and 3-D references are part of the graph): false.
let isUnsupported (engine: Engine) = ignore engine; false

/// Number of formulas the last full pass (`recalculateAll`) found on a circular
/// reference: they show the cycle error, or - in a workbook with iterative
/// calculation - were calculated until they settled.
let cyclicCount (engine: Engine) = engine.LastCyclic

/// Number of formulas that read a 3-D reference (`Jan:Mar!B1`): what they read
/// changes with the order of the sheets, without any cell or formula changing.
let threeDCount (engine: Engine) = engine.ThreeD

/// Number of formulas that choose their precedents while they are evaluated.
let dynamicCount (engine: Engine) = engine.Dynamic.Count

let lastStats (engine: Engine) = engine.LastStats

/// Why `recalculate` last declined an edit.
let lastDecline (engine: Engine) = engine.LastDecline

/// The formulas that read a cell - on its own, or in a range - added to `into`
/// once each under the edge mark `mark`.
let private addReadersMarked (mark: int) (sheet: SheetState) (row: int) (col: int) (explicit: ResizeArray<Node>) (into: ResizeArray<Node>) =
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

/// The formulas that read a cell: on its own, or in a range.
let private addReaders (engine: Engine) (sheet: SheetState) (row: int) (col: int) (explicit: ResizeArray<Node>) (into: ResizeArray<Node>) =
    engine.EdgeMark <- engine.EdgeMark + 1
    addReadersMarked engine.EdgeMark sheet row col explicit into

/// The spill map of a sheet without converting a sheet that is not loaded.
let private loadedSpill (model: SpreadsheetModel) : Map<CellAddress, CellAddress> =
    match model.Deferred with
    | Some deferred ->
        match deferred.Loaded with
        | Some loaded -> loaded.Spill
        | None -> Map.empty
    | None -> model.Spill

/// The spilled cells of a sheet by the formula they belong to, for the spill map given.
let private ghostsOf (state: SheetState) (spill: Map<CellAddress, CellAddress>) : Dictionary<float, ResizeArray<CellAddress>> =
    if not (obj.ReferenceEquals(state.GhostsFor, spill)) then
        state.Ghosts.Clear()
        for KeyValue (cell, anchor) in spill do
            let key = float anchor.Row * 16384.0 + float anchor.Col
            match state.Ghosts.TryGetValue key with
            | true, cells -> cells.Add cell
            | false, _ ->
                let cells = ResizeArray<CellAddress>()
                cells.Add cell
                state.Ghosts.[key] <- cells
        state.GhostsFor <- box spill
    state.Ghosts

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

/// The cells of a sheet that differ between two cell maps, by walking both in
/// order. A spilled cell (one in `spill`) that differs in more than its value -
/// it was given a format - goes to `restyled`: its formula lays it out again.
let private diffCells (state: SheetState) (before: Map<CellAddress, CellData>) (after: Map<CellAddress, CellData>) (spill: Map<CellAddress, CellAddress>)
                      (changes: ResizeArray<CellChange>) (restyled: ResizeArray<struct (SheetState * CellAddress)>) =
    use left = (before :> seq<KeyValuePair<CellAddress, CellData>>).GetEnumerator()
    use right = (after :> seq<KeyValuePair<CellAddress, CellData>>).GetEnumerator()
    let hasSpill = not (Map.isEmpty spill)
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
            if not (obj.ReferenceEquals(was, now)) then
                if not (obj.ReferenceEquals(was.Value, now.Value)) && was.Value <> now.Value then
                    changes.Add { State = state; Row = left.Current.Key.Row; Col = left.Current.Key.Col; Before = ValueSome was; After = ValueSome now }
                elif hasSpill && Map.containsKey left.Current.Key spill then
                    restyled.Add (struct (state, left.Current.Key))
            hasLeft <- left.MoveNext()
            hasRight <- right.MoveNext()
        elif order < 0 then
            let was = left.Current.Value
            if was.Value <> CellValue.Empty then
                changes.Add { State = state; Row = left.Current.Key.Row; Col = left.Current.Key.Col; Before = ValueSome was; After = ValueNone }
            elif hasSpill && Map.containsKey left.Current.Key spill then
                restyled.Add (struct (state, left.Current.Key))
            hasLeft <- left.MoveNext()
        else
            let now = right.Current.Value
            if now.Value <> CellValue.Empty then
                changes.Add { State = state; Row = right.Current.Key.Row; Col = right.Current.Key.Col; Before = ValueNone; After = ValueSome now }
            elif hasSpill && Map.containsKey right.Current.Key spill then
                restyled.Add (struct (state, right.Current.Key))
            hasRight <- right.MoveNext()

/// The cells whose entry differs between two spill maps: each with the formula
/// cells it belonged / belongs to.
let private diffSpill (before: Map<CellAddress, CellAddress>) (after: Map<CellAddress, CellAddress>) (onAnchor: CellAddress -> unit) =
    use left = (before :> seq<KeyValuePair<CellAddress, CellAddress>>).GetEnumerator()
    use right = (after :> seq<KeyValuePair<CellAddress, CellAddress>>).GetEnumerator()
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
            if left.Current.Value <> right.Current.Value then
                onAnchor left.Current.Value
                onAnchor right.Current.Value
            hasLeft <- left.MoveNext()
            hasRight <- right.MoveNext()
        elif order < 0 then
            onAnchor left.Current.Value
            hasLeft <- left.MoveNext()
        else
            onAnchor right.Current.Value
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
    && known.Iteration = edited.Iteration
    && known.Sheets.Count = edited.Sheets.Count
    && (edited.Sheets |> Map.forall (fun sheetId sheet ->
            match engine.Sheets.TryGetValue sheetId with
            | true, state when not (isNull (box state.Model)) ->
                let before = state.Model
                let after = sheet.Model
                obj.ReferenceEquals(before, after)
                || (state.Name = sheet.Name
                    && before.StructureVersion = after.StructureVersion
                    && (obj.ReferenceEquals(before.Tables, after.Tables) || before.Tables = after.Tables)
                    && sameFilterState before after)
            | _ -> false))

/// The cells an edit changed: those named by `hint`, or - without one - the
/// cells of every sheet whose cell map is another one, compared cell by cell.
/// `restyled` takes the spilled cells that changed in their format alone.
///
/// `structural` names the sheets whose rows or columns moved: of those only
/// the cells that hold or held a formula are compared (through the formula
/// index - the values are not walked); the caller has everything that reads
/// the sheet calculated again.
let private collectChanges (engine: Engine) (edited: Workbook) (hint: (SheetId * CellAddress) list option)
                           (restyled: ResizeArray<struct (SheetState * CellAddress)>) (structural: HashSet<SheetId>) : ResizeArray<CellChange> =
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
                else
                    let same =
                        match before, after with
                        | ValueSome was, ValueSome now -> obj.ReferenceEquals(was, now)
                        | ValueNone, ValueNone -> true
                        | _ -> false
                    if not same && Map.containsKey addr (loadedSpill sheet.Model) then restyled.Add (struct (state, addr))
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
                if not (isLoaded before) && isLoaded after then
                    // The sheet was converted since (its formulas were known from the file):
                    // a formula that was rewritten on the way - a sheet it names was renamed,
                    // rows it reads moved - is another formula than the graph holds.
                    let seen = HashSet<float>()
                    for node in List.ofSeq state.Nodes.Values do
                        if node.IsFormula then
                            seen.Add (keyOf node.Row node.Col) |> ignore
                            let now = match Map.tryFind { Row = node.Row; Col = node.Col } afterCells with Some cell -> ValueSome cell | None -> ValueNone
                            let text = formulaOf now
                            if isNull text || text <> node.Formula then
                                let was : CellData voption =
                                    match now with
                                    | ValueSome cell -> ValueSome { cell with Value = CellValue.Formula (node.Formula, effectiveValue now) }
                                    | ValueNone -> ValueSome { Value = CellValue.Formula (node.Formula, CellValue.Empty); Format = defaultFormat; Comment = None }
                                changes.Add { State = state; Row = node.Row; Col = node.Col; Before = was; After = now }
                    for addr in formulaIndexOf afterCells do
                        if seen.Add (keyOf addr.Row addr.Col) then
                            changes.Add { State = state; Row = addr.Row; Col = addr.Col; Before = ValueNone; After = (match Map.tryFind addr afterCells with Some cell -> ValueSome cell | None -> ValueNone) }
                elif obj.ReferenceEquals(beforeCells, afterCells) then ()
                elif structural.Contains sheetId then
                    let seen = HashSet<float>()
                    let consider (addr: CellAddress) =
                        if seen.Add (keyOf addr.Row addr.Col) then
                            let was = match Map.tryFind addr beforeCells with Some cell -> ValueSome cell | None -> ValueNone
                            let now = match Map.tryFind addr afterCells with Some cell -> ValueSome cell | None -> ValueNone
                            let differs =
                                match was, now with
                                | ValueSome a, ValueSome b -> not (obj.ReferenceEquals(a, b)) && a.Value <> b.Value
                                | ValueNone, ValueNone -> false
                                | ValueSome a, ValueNone -> a.Value <> CellValue.Empty
                                | ValueNone, ValueSome b -> b.Value <> CellValue.Empty
                            if differs then changes.Add { State = state; Row = addr.Row; Col = addr.Col; Before = was; After = now }
                    for node in List.ofSeq state.Nodes.Values do
                        if node.IsFormula then consider { Row = node.Row; Col = node.Col }
                    for addr in formulaIndexOf afterCells do consider addr
                else
                    diffCells state beforeCells afterCells (loadedSpill after) changes restyled
    changes

/// Make the graph follow the formulas among the changed cells. `onFormula`
/// is given each changed cell that holds a formula now, `onValue` each one
/// that does not.
let private rewire (engine: Engine) (edited: Workbook) (changes: ResizeArray<CellChange>) (onFormula: Node -> CellChange -> unit) (onValue: CellChange -> unit) =
    // The formulas that are gone or are other formulas now leave the graph, together.
    let going = ResizeArray<Node>()
    for change in changes do
        let formulaAfter = formulaOf change.After
        match change.State.Nodes.TryGetValue (keyOf change.Row change.Col) with
        | true, node when node.IsFormula && (isNull formulaAfter || node.Formula <> formulaAfter) -> going.Add node
        | _ -> ()
    unwireMany engine going
    for change in changes do
        let state = change.State
        let formulaAfter = formulaOf change.After
        if not (isNull formulaAfter) then
            let node =
                match state.Nodes.TryGetValue (keyOf change.Row change.Col) with
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

/// Whether a formula as an insertion / deletion of rows or columns left it
/// (`after`) reads the very cells it read before (`before`): every reference
/// of it either stands where it stood, wholly before the edit or on another
/// sheet, or moved with its cells, both corners by the edit's count. Such a
/// formula gives the value it gave: it is wired again and not evaluated.
/// False - evaluate - for anything else: a range the edit runs through or
/// took cells from, a name that changed, structured and spill references, arrays, and the
/// functions that answer with a position or with the rows that are shown.
/// `onMovedSheet` says whether a reference's sheet (None: the formula's own)
/// is the sheet that was edited.
/// `nameStands` says of a defined name whether it refers to the cells it referred to.
let private followsMove (move: StructureMove) (onMovedSheet: string option -> bool) (nameStands: string -> bool) (before: Expr) (after: Expr) : bool =
    let sameOnOtherAxis (a: RefAst) (b: RefAst) =
        if move.OnRows then a.Col = b.Col else a.Row = b.Row
    let along (r: RefAst) = if move.OnRows then r.Row else r.Col
    let follows (a: RefAst) (b: RefAst) (a2: RefAst) (b2: RefAst) =
        a.Sheet = a2.Sheet && b.Sheet = b2.Sheet
        && (if not (onMovedSheet a.Sheet) then a.Row = a2.Row && a.Col = a2.Col && b.Row = b2.Row && b.Col = b2.Col
            else
                sameOnOtherAxis a a2 && sameOnOtherAxis b b2
                && (let lo, hi = min (along a) (along b), max (along a) (along b)
                    if hi < move.At then along a2 = along a && along b2 = along b
                    elif move.Insert then lo >= move.At && along a2 = along a + move.Count && along b2 = along b + move.Count
                    else lo >= move.At + move.Count && along a2 = along a - move.Count && along b2 = along b - move.Count))
    let left = Stack<Expr>()
    let right = Stack<Expr>()
    left.Push before
    right.Push after
    let mutable same = true
    while same && left.Count > 0 do
        match left.Pop(), right.Pop() with
        | Number a, Number b -> same <- a = b
        | Str a, Str b -> same <- a = b
        | Bool a, Bool b -> same <- a = b
        | ErrorLit a, ErrorLit b -> same <- a = b
        | MissingArg, MissingArg -> ()
        | NameRef a, NameRef b -> same <- a = b && nameStands a
        | CellRef a, CellRef b -> same <- not (a.Sheet |> Option.exists (fun sheet -> sheet.IndexOf ':' >= 0)) && follows a a b b
        | RangeRef (a, b), RangeRef (a2, b2) -> same <- not (a.Sheet |> Option.exists (fun sheet -> sheet.IndexOf ':' >= 0)) && follows a b a2 b2
        | UnaryOp (opA, a), UnaryOp (opB, b) ->
            same <- opA = opB
            left.Push a
            right.Push b
        | BinOp (a1, opA, a2), BinOp (b1, opB, b2) ->
            same <- opA = opB
            left.Push a1
            right.Push b1
            left.Push a2
            right.Push b2
        | Function (nameA, argsA), Function (nameB, argsB) ->
            if nameA <> nameB || List.length argsA <> List.length argsB then same <- false
            else
                let upper = nameA.ToUpperInvariant()
                if upper = "ROW" || upper = "COLUMN" || upper = "CELL" || upper = "SHEET" || upper = "SHEETS" || upper = "FORMULATEXT" || upper = "ISFORMULA" || traitsOf nameA <> 0 then same <- false
                else
                    for arg in argsA do left.Push arg
                    for arg in argsB do right.Push arg
        | _ -> same <- false
    same

/// What an edit changed besides cells, as far as a recalculation has to know.
type private Surroundings = {
    Any: bool
    /// Why the edit is left to the full recalculation (null: it is not).
    Problem: string
    Added: ResizeArray<SheetId>
    Removed: ResizeArray<SheetState>
    Renamed: ResizeArray<SheetState>
    /// Sheets were added, removed, renamed or put in another order.
    SheetsChanged: bool
    NamesChanged: bool
    /// The names (upper case, whatever their scope) that were defined, changed or deleted,
    /// and the names that are defined through one of them.
    ChangedNames: HashSet<string>
    /// Names that only followed rows or columns that moved (`followsMove`): they refer to
    /// the cells they referred to, which stand elsewhere now.
    FollowerNames: HashSet<string>
    TablesChanged: bool
    /// Rows may be hidden or shown: a filter, row heights, rows that moved.
    VisibilityChanged: bool
    /// Sheets whose rows or columns moved.
    Structural: HashSet<SheetId>
    /// Of those, the sheets where that was one insertion or deletion, with it.
    Moves: Dictionary<SheetId, StructureMove>
}

let private surroundingsOf (engine: Engine) (edited: Workbook) : Surroundings =
    let known = engine.Workbook
    let added = ResizeArray<SheetId>()
    let removed = ResizeArray<SheetState>()
    let renamed = ResizeArray<SheetState>()
    let structural = HashSet<SheetId>()
    let moves = Dictionary<SheetId, StructureMove>()
    let mutable problem : string = null
    let mutable tables = not (obj.ReferenceEquals(known.Tables, edited.Tables) || known.Tables = edited.Tables)
    let mutable visibility = false
    for KeyValue (sheetId, sheet) in edited.Sheets do
        match engine.Sheets.TryGetValue sheetId with
        | true, state when not (isNull (box state.Model)) ->
            if state.Name <> sheet.Name then renamed.Add state
            let before = state.Model
            let after = sheet.Model
            if not (obj.ReferenceEquals(before, after)) then
                if not (obj.ReferenceEquals(before.Tables, after.Tables) || before.Tables = after.Tables) then tables <- true
                if not (sameFilterState before after) then visibility <- true
                if not (obj.ReferenceEquals(before.RowHeights, after.RowHeights)) && before.RowHeights <> after.RowHeights then visibility <- true
                if before.StructureVersion <> after.StructureVersion then
                    if not (isLoaded after) || not (isLoaded before) then problem <- "rows or columns changed on a sheet that is not converted"
                    structural.Add sheetId |> ignore
                    visibility <- true
                    // (One edit on from the sheet as the engine holds it: which cells moved is known.)
                    match after.StructureMove with
                    | Some move when move.Version = after.StructureVersion && after.StructureVersion = before.StructureVersion + 1 -> moves.[sheetId] <- move
                    | _ -> ()
        | _ -> added.Add sheetId
    for KeyValue (sheetId, state) in engine.Sheets do
        if not (Map.containsKey sheetId edited.Sheets) then removed.Add state
    let names = not (obj.ReferenceEquals(known.Names, edited.Names) || known.Names = edited.Names)
    let changedNames = HashSet<string>()
    let followerNames = HashSet<string>()
    /// The one sheet whose rows or columns moved, with its name and the move, when that is all that moved.
    let theMove =
        if moves.Count = 1 && structural.Count = 1 then
            let pair = Seq.head moves
            Some ((Map.find pair.Key edited.Sheets).Name, pair.Value)
        else None
    let parsed (refersTo: string) = parseFormula (if refersTo.StartsWith "=" then refersTo else "=" + refersTo)
    // A name that stands as it was written while rows or columns moved under it (a whole
    // column of the sheet, a range the edit runs through) refers to other cells now.
    match theMove with
    | Some (sheetName, move) ->
        for KeyValue ((scope, name), refersTo) in edited.Names do
            match Map.tryFind (scope, name) known.Names with
            | Some was when was = refersTo ->
                let stands =
                    match parsed refersTo with
                    | Ok expr ->
                        let rec plain (e: Expr) =
                            match e with
                            | CellRef r -> r.Sheet.IsSome
                            | RangeRef (a, _) -> a.Sheet.IsSome
                            | Number _ | Str _ | Bool _ | ErrorLit _ -> true
                            | UnaryOp (_, inner) -> plain inner
                            | BinOp (left, _, right) -> plain left && plain right
                            | _ -> false
                        // (Anything but plain references and constants - a function, another name,
                        // a reference without a sheet - may read what moved.)
                        plain expr
                        && followsMove move (fun qualifier -> match qualifier with Some q -> q.Equals(sheetName, System.StringComparison.OrdinalIgnoreCase) | None -> false) (fun _ -> false) expr expr
                    | Result.Error _ -> true
                if not stands then changedNames.Add (name.ToUpperInvariant()) |> ignore
            | _ -> ()
    | None -> ()
    let names = names || changedNames.Count > 0
    if names then
        for KeyValue ((scope, name), refersTo) in known.Names do
            match Map.tryFind (scope, name) edited.Names with
            | Some now when now = refersTo -> ()
            | Some now ->
                let follows =
                    match theMove, parsed refersTo, parsed now with
                    | Some (sheetName, move), Ok was, Ok is ->
                        // (A name inside a name: the name is taken as changed.)
                        followsMove move (fun qualifier -> match qualifier with Some q -> q.Equals(sheetName, System.StringComparison.OrdinalIgnoreCase) | None -> false) (fun _ -> false) was is
                        // (A reference without a sheet in a name is relative to whoever uses it: not followed here.)
                        && (let rec qualified (expr: Expr) =
                                match expr with
                                | CellRef r -> r.Sheet.IsSome
                                | RangeRef (a, _) -> a.Sheet.IsSome
                                | UnaryOp (_, inner) -> qualified inner
                                | BinOp (left, _, right) -> qualified left && qualified right
                                | Function (_, args) -> args |> List.forall qualified
                                | _ -> true
                            qualified was)
                    | _ -> false
                if follows then followerNames.Add (name.ToUpperInvariant()) |> ignore
                else changedNames.Add (name.ToUpperInvariant()) |> ignore
            | None -> changedNames.Add (name.ToUpperInvariant()) |> ignore
        for KeyValue ((scope, name), _) in edited.Names do
            if not (Map.containsKey (scope, name) known.Names) then changedNames.Add (name.ToUpperInvariant()) |> ignore
        // A name defined through a changed one reads something else too.
        let mentions (refersTo: string) =
            match parseFormula (if refersTo.StartsWith "=" then refersTo else "=" + refersTo) with
            | Ok expr ->
                let stack = ResizeArray<Expr>()
                stack.Add expr
                let mutable found = false
                while not found && stack.Count > 0 do
                    let current = stack.[stack.Count - 1]
                    stack.RemoveAt (stack.Count - 1)
                    match current with
                    | NameRef name -> if changedNames.Contains (name.ToUpperInvariant()) then found <- true
                    | UnaryOp (_, inner) -> stack.Add inner
                    | BinOp (left, _, right) ->
                        stack.Add left
                        stack.Add right
                    | Function (_, args) -> for arg in args do stack.Add arg
                    | _ -> ()
                found
            | Result.Error _ -> true
        let mutable growing = true
        while growing do
            growing <- false
            for KeyValue ((_, name), refersTo) in edited.Names do
                let upper = name.ToUpperInvariant()
                if not (changedNames.Contains upper) && mentions refersTo then
                    changedNames.Add upper |> ignore
                    growing <- true
        // (A name that is both - by scope - counts as changed.)
        followerNames.ExceptWith changedNames
        // A name defined through a follower follows too.
        let mentionsFollower (refersTo: string) =
            match parsed refersTo with
            | Ok expr ->
                let stack = ResizeArray<Expr>()
                stack.Add expr
                let mutable found = false
                while not found && stack.Count > 0 do
                    let current = stack.[stack.Count - 1]
                    stack.RemoveAt (stack.Count - 1)
                    match current with
                    | NameRef name -> if followerNames.Contains (name.ToUpperInvariant()) then found <- true
                    | UnaryOp (_, inner) -> stack.Add inner
                    | BinOp (left, _, right) ->
                        stack.Add left
                        stack.Add right
                    | Function (_, args) -> for arg in args do stack.Add arg
                    | _ -> ()
                found
            | Result.Error _ -> false
        let mutable following = followerNames.Count > 0
        while following do
            following <- false
            for KeyValue ((_, name), refersTo) in edited.Names do
                let upper = name.ToUpperInvariant()
                if not (changedNames.Contains upper) && not (followerNames.Contains upper) && mentionsFollower refersTo then
                    followerNames.Add upper |> ignore
                    following <- true
    let order = not (obj.ReferenceEquals(known.SheetOrder, edited.SheetOrder) || known.SheetOrder = edited.SheetOrder)
    let sheets = added.Count > 0 || removed.Count > 0 || renamed.Count > 0 || order
    { Any = sheets || names || tables || visibility || structural.Count > 0
      Problem = problem
      Added = added; Removed = removed; Renamed = renamed
      SheetsChanged = sheets; NamesChanged = names; ChangedNames = changedNames; FollowerNames = followerNames; TablesChanged = tables; VisibilityChanged = visibility
      Structural = structural; Moves = moves }

let private blankModel = createEmpty 1 1

/// Take a workbook that was recalculated elsewhere (the full recalculation,
/// after `recalculate` declined an edit) as the engine's calculated state:
/// the graph follows the formulas that differ from the state it held. False
/// when more than cells differ - the engine is then to be built again.
let adopt (engine: Engine) (calculated: Workbook) : bool =
    if not (sameSurroundings engine calculated) then false
    else
        let changes = collectChanges engine calculated None (ResizeArray()) (HashSet<SheetId>())
        rewire engine calculated changes (fun _ _ -> ()) ignore
        // (Formulas that came in may close a circular reference nothing here has looked for.)
        if changes.Count > 0 then engine.CyclesKnown <- false
        // (The arrays were laid out elsewhere: their sizes are not known here.)
        for node in engine.Spilling do
            node.ExtentRows <- -1
            node.ExtentCols <- -1
        holdWorkbook engine calculated
        true

/// Recalculate the workbook after an edit of cells. `edited` is the engine's
/// workbook with the edit applied and nothing recalculated. `hint` names the
/// cells the edit touched when the caller knows them (a typed entry); without
/// it the sheets whose cells are another map are compared cell by cell.
/// Some: the recalculated workbook - the engine now holds it. None: the edit
/// needs the full recalculation; afterwards the engine is shown its result
/// (`adopt`) or built again.
///
/// Dynamic arrays. A formula that spills owns the cells it spilled into (the
/// sheet's spill map) as its outputs: what reads one of them is ordered after
/// the formula. The formula is laid out again - its array evaluated, its
/// cells written, cleared, or the formula blocked (#SPILL!) - when something
/// it reads changed, when it was entered, changed or removed, when one of its
/// cells was typed over, cleared or restyled, and when a cell in the way of a
/// blocked formula changed. Laying it out can change which cells it owns: what
/// reads a cell it newly covers, a formula it took a cell from (the formula
/// further up / left wins a cell two arrays want, as in the full
/// recalculation) and a blocked formula that wants a cell it gave up are
/// calculated in the same pass when they come later in its order, else in a
/// further round of the pass.
///
/// Precedents chosen while a formula is evaluated (OFFSET, INDIRECT). Such a
/// formula is evaluated in every pass, at the place its spelled-out references
/// give it. A cell it turns out to read may be calculated after it: whenever
/// a value changes after the first of these formulas was evaluated, they are
/// all evaluated once more in a further round (what reads them follows when
/// their value changed), until a round changes nothing after them.
///
/// `full`: the pass is the full recalculation - every formula is evaluated,
/// whatever the engine held, and formulas on a circular reference are given
/// the cycle error instead of the pass being declined (the strongly connected
/// components of what is left when the order cannot go on: a formula that
/// reads itself, or two or more that read one another, also across sheets).
/// A formula marked so stays marked for the rest of the pass. In a workbook
/// with iterative calculation (`Workbook.Iteration`, Excel's setting) the
/// formulas of each such component are instead calculated again and again,
/// from the values they have, in the order they were met - until a sweep
/// changes no value by the workbook's largest change or more (that sweep is
/// dropped, so a workbook that had settled is left exactly as it is), or the
/// workbook's number of iterations is reached.
let rec private run (mode: int) (engine: Engine) (edited: Workbook) (hint: (SheetId * CellAddress) list option) : Workbook option =
    // 0: an edit (incremental); 1: the full recalculation; 2: no calculation - the circular references of the graph are looked for.
    let full = mode = 1
    let scan = mode = 2
    let decline (reason: string) : Workbook option =
        engine.LastDecline <- reason
        // (The spilled cells by formula may have been moved on by the rounds that ran.)
        for KeyValue (_, state) in engine.Sheets do state.GhostsFor <- null
        None
    // (The circular references of a graph that was built, not calculated, are not known:
    // they are taken as settled - the workbook came calculated - until an edit reaches
    // one, or `findCycles` has looked for them.)
    let surroundings =
        if mode = 0 then surroundingsOf engine edited
        else
            { Any = false; Problem = null; Added = ResizeArray(); Removed = ResizeArray(); Renamed = ResizeArray()
              SheetsChanged = false; NamesChanged = false; ChangedNames = HashSet<string>(); FollowerNames = HashSet<string>(); TablesChanged = false; VisibilityChanged = false
              Structural = HashSet<SheetId>(); Moves = Dictionary<SheetId, StructureMove>() }
    if mode = 0 && engine.Workbook.Iteration <> edited.Iteration then decline "the setting of iterative calculation changed"
    elif not (isNull surroundings.Problem) then decline surroundings.Problem
    else
        // A sheet that is new: nothing was there.
        for sheetId in surroundings.Added do
            let state = sheetState engine sheetId
            state.Model <- blankModel
            state.Name <- (Map.find sheetId edited.Sheets).Name
        let restyled = ResizeArray<struct (SheetState * CellAddress)>()
        let changes =
            if mode <> 0 then ResizeArray<CellChange>()
            else collectChanges engine edited (if surroundings.Any then None else hint) restyled surroundings.Structural
        let mutable workbook = edited
        let models = Dictionary<SheetId, SpreadsheetModel>()
        /// The cells of a sheet as the pass first took them: the pass changes values, never which cells hold formulas.
        let cellsAtStart = Dictionary<SheetId, Map<CellAddress, CellData>>()
        let touched = HashSet<SheetId>()
        let modelOf (state: SheetState) =
            match models.TryGetValue state.Id with
            | true, model -> model
            | false, _ ->
                // (A formula is evaluated in its sheet, not in the sheet's stand-in.)
                let model = resolveSheet (Map.find state.Id workbook.Sheets).Model
                models.[state.Id] <- model
                if not (cellsAtStart.ContainsKey state.Id) then cellsAtStart.[state.Id] <- model.Cells
                model
        /// The sheet's spill map as the pass has it now (a sheet that is not loaded has none to speak of).
        let spillOf (state: SheetState) : Map<CellAddress, CellAddress> =
            match models.TryGetValue state.Id with
            | true, model -> model.Spill
            | false, _ ->
                match Map.tryFind state.Id workbook.Sheets with
                | Some sheet -> loadedSpill sheet.Model
                | None -> Map.empty
        let cellsOf (state: SheetState) : Map<CellAddress, CellData> =
            match models.TryGetValue state.Id with
            | true, model -> model.Cells
            | false, _ ->
                match Map.tryFind state.Id workbook.Sheets with
                | Some sheet -> loadedCells sheet.Model
                | None -> Map.empty

        let mutable pass = 0
        let mutable reached = ResizeArray<Node>()
        let pending = Stack<Node>()
        let reach (node: Node) =
            if node.Pass <> pass then
                node.Pass <- pass
                node.Pending <- 0
                node.NeedsEval <- false
                node.NeedsSpill <- false
                node.Order <- -1
                node.Outs <- null
                reached.Add node
                pending.Push node
        let newRound () =
            engine.Pass <- engine.Pass + 1
            pass <- engine.Pass
            reached <- ResizeArray<Node>()
        newRound ()
        let firstPass = pass

        // ---- what changed besides cells ------------------------------------------------
        // Names, tables, the sheets themselves, a filter, rows and columns: the formulas
        // that can tell are wired again (what they read may be something else now) and
        // calculated; the others are not looked at.
        /// The formulas that read a sheet at or beyond a row / column edit of it, with their wiring.
        let moveReaders = ResizeArray<struct (Node * int)>()
        if mode = 0 && surroundings.Any then
            let rewireKeys = ResizeArray<struct (SheetState * int * int)>()
            /// Formulas to wire again that give the value they gave (a name they use followed its cells).
            let quietKeys = ResizeArray<struct (SheetState * int * int)>()
            let seedKeys = ResizeArray<struct (SheetState * int * int)>()
            engine.EdgeMark <- engine.EdgeMark + 3
            let rewireMark = engine.EdgeMark - 2
            let quietMark = engine.EdgeMark - 1
            let seedMark = engine.EdgeMark
            let toRewire (node: Node) =
                if node.IsFormula && node.EdgeMark <> rewireMark then
                    node.EdgeMark <- rewireMark
                    rewireKeys.Add (struct (node.Sheet, node.Row, node.Col))
            let toRewireQuietly (node: Node) =
                if node.IsFormula && node.EdgeMark <> rewireMark && node.EdgeMark <> quietMark then
                    node.EdgeMark <- quietMark
                    quietKeys.Add (struct (node.Sheet, node.Row, node.Col))
            let toSeed (node: Node) =
                if node.IsFormula && node.EdgeMark <> rewireMark && node.EdgeMark <> seedMark then
                    node.EdgeMark <- seedMark
                    seedKeys.Add (struct (node.Sheet, node.Row, node.Col))
            /// Every formula that reads a cell or a range of the sheet.
            let readersOf (state: SheetState) (onReader: Node -> unit) =
                for node in state.Nodes.Values do
                    for reader in node.Dependents do onReader reader
                for group in state.Groups.Values do
                    for owner in group.Owners do onReader owner
            let rewireTraits =
                (if surroundings.SheetsChanged then 4 ||| 64 ||| 128 else 0)
                ||| (if surroundings.TablesChanged then 32 else 0)
            let seedTraits = if surroundings.VisibilityChanged then 8 else 0
            /// The formula uses one of the names that changed (itself: a name used through another is in the set).
            let usesNameOf (names: HashSet<string>) (node: Node) =
                match node.Expr with
                | None -> true
                | Some expr ->
                    let stack = ResizeArray<Expr>()
                    stack.Add expr
                    let mutable found = false
                    while not found && stack.Count > 0 do
                        let current = stack.[stack.Count - 1]
                        stack.RemoveAt (stack.Count - 1)
                        match current with
                        | NameRef name -> if names.Contains (name.ToUpperInvariant()) then found <- true
                        | UnaryOp (_, inner) -> stack.Add inner
                        | BinOp (left, _, right) ->
                            stack.Add left
                            stack.Add right
                        | Function (_, args) -> for arg in args do stack.Add arg
                        | _ -> ()
                    found
            if rewireTraits <> 0 || surroundings.NamesChanged then
                for state in engine.Sheets.Values do
                    for node in state.Nodes.Values do
                        if node.IsFormula then
                            if node.Traits &&& rewireTraits <> 0 then toRewire node
                            elif surroundings.NamesChanged && node.Traits &&& 16 <> 0 then
                                if surroundings.ChangedNames.Count > 0 && usesNameOf surroundings.ChangedNames node then toRewire node
                                elif surroundings.FollowerNames.Count > 0 && usesNameOf surroundings.FollowerNames node then toRewireQuietly node
            // (A reference to a sheet that is gone, or that has another name now, reads something else.)
            for state in surroundings.Removed do readersOf state toRewire
            for state in surroundings.Renamed do readersOf state toRewire
            if seedTraits <> 0 then
                for state in engine.Sheets.Values do
                    for node in state.Nodes.Values do
                        if node.IsFormula && node.Traits &&& seedTraits <> 0 then toSeed node
            // One insertion or deletion: what reads the sheet at or beyond the edit, as it is wired now.
            for KeyValue (sheetId, move) in surroundings.Moves do
                match engine.Sheets.TryGetValue sheetId with
                | true, state ->
                    let note (reader: Node) = if reader.IsFormula then moveReaders.Add (struct (reader, reader.WireStamp))
                    for node in state.Nodes.Values do
                        if (if move.OnRows then node.Row else node.Col) >= move.At then
                            for reader in node.Dependents do note reader
                    for group in state.Groups.Values do
                        if (if move.OnRows then group.R1 else group.C1) >= move.At then
                            for owner in group.Owners do note owner
                | false, _ -> ()
            // (Rows or columns moved: the cells of the sheet are not compared one by one -
            // whatever reads the sheet is calculated. When it was one insertion or
            // deletion, less is: see below, after the formulas that follow the rows.)
            for sheetId in surroundings.Structural do
                if not (surroundings.Moves.ContainsKey sheetId) then
                    match engine.Sheets.TryGetValue sheetId with
                    | true, state -> readersOf state toSeed
                    | false, _ -> ()
            if surroundings.NamesChanged || surroundings.SheetsChanged then engine.NameTraits.Clear()
            // They leave the graph together (see `unwireMany`) and are put back one by one.
            let going = ResizeArray<Node>()
            for state in surroundings.Removed do
                for node in state.Nodes.Values do
                    if node.IsFormula then going.Add node
            let again = ResizeArray<struct (SheetState * int * int * string * bool)>()
            let wireAgain (evaluate: bool) (state: SheetState) (row: int) (col: int) =
                match Map.tryFind state.Id edited.Sheets with
                | Some sheet ->
                    match state.Nodes.TryGetValue (keyOf row col) with
                    | true, node when node.IsFormula ->
                        // (The formula as the edit left it; of a sheet that is not converted, as it was.)
                        let text =
                            if isLoaded sheet.Model then
                                formulaOf (match Map.tryFind { Row = row; Col = col } (loadedCells sheet.Model) with Some cell -> ValueSome cell | None -> ValueNone)
                            else node.Formula
                        // (Whether its array has to be laid out again is not known of a formula that is only wired again.)
                        let evaluate = evaluate || node.Spills || node.Traits &&& (1 ||| 2 ||| 4 ||| 8 ||| 32 ||| 64 ||| 128) <> 0
                        going.Add node
                        if not (isNull text) then again.Add (struct (state, row, col, text, evaluate))
                    | _ -> ()
                | None -> ()
            for struct (state, row, col) in rewireKeys do wireAgain true state row col
            for struct (state, row, col) in quietKeys do wireAgain false state row col
            unwireMany engine going
            for struct (state, row, col, text, evaluate) in again do
                let fresh = wire engine edited state row col text
                // (A formula that gives the value it gave is not part of the pass, unless something it reads changes.)
                if evaluate || fresh.Spills then
                    reach fresh
                    fresh.NeedsEval <- true
            for state in surroundings.Removed do engine.Sheets.Remove state.Id |> ignore
            // A new sheet that is not converted brings its formulas as its file has them.
            for sheetId in surroundings.Added do
                let sheet = Map.find sheetId edited.Sheets
                match sheet.Model.Deferred with
                | Some deferred when deferred.Loaded.IsNone ->
                    let state = sheetState engine sheetId
                    for addr, formula in deferred.Formulas () do
                        wire engine edited state addr.Row addr.Col formula |> ignore
                | _ -> ()
            for struct (state, row, col) in seedKeys do
                match state.Nodes.TryGetValue (keyOf row col) with
                | true, node when node.IsFormula ->
                    reach node
                    node.NeedsEval <- true
                | _ -> ()

        let readers = ResizeArray<Node>()
        /// The one insertion / deletion of rows or columns of the edit, with its sheet (see `followsMove`).
        let theMove =
            if surroundings.Moves.Count = 1 && surroundings.Structural.Count = 1 then
                let pair = Seq.head surroundings.Moves
                match engine.Sheets.TryGetValue pair.Key, Map.tryFind pair.Key edited.Sheets with
                | (true, state), Some sheet when not (isNull (box state.Model)) -> Some (state, sheet.Name, pair.Value, loadedCells state.Model)
                | _ -> None
            else None
        /// What a changed cell held before the edit: the cell that stood there, or - on a sheet whose
        /// rows or columns moved - the cell that moved there.
        let heldBefore (change: CellChange) : CellData voption =
            match theMove with
            | Some (moved, _, move, cellsBefore) when obj.ReferenceEquals(change.State, moved) ->
                let along = if move.OnRows then change.Row else change.Col
                let from =
                    if along < move.At then Some along
                    elif move.Insert then (if along >= move.At + move.Count then Some (along - move.Count) else None)
                    else Some (along + move.Count)
                match from with
                | Some index ->
                    (match Map.tryFind (if move.OnRows then { Row = index; Col = change.Col } else { Row = change.Row; Col = index }) cellsBefore with
                     | Some cell -> ValueSome cell
                     | None -> ValueNone)
                | None -> ValueNone
            | _ -> change.Before
        /// A changed formula that only followed the cells it reads (they moved) and gives the value it gave.
        let followed (node: Node) (change: CellChange) (was: CellData voption) : bool =
            match theMove, node.Expr with
            | Some (moved, movedName, move, _), Some expr when not node.Spills && node.Traits &&& ~~~16 = 0 ->
                let onMoved = obj.ReferenceEquals(change.State, moved)
                let text = formulaOf was
                if isNull text || effectiveValue was <> effectiveValue change.After then false
                else
                    match parseFormula text with
                    | Ok before ->
                        followsMove move (fun qualifier ->
                            match qualifier with
                            | None -> onMoved
                            | Some q -> q.Equals(movedName, System.StringComparison.OrdinalIgnoreCase))
                            (fun name -> not (surroundings.ChangedNames.Contains (name.ToUpperInvariant()))) before expr
                    | Result.Error _ -> false
            | _ -> false
        rewire engine edited changes
            (fun node change ->
                let was = heldBefore change
                if not (followed node change was) then
                    reach node
                    node.NeedsEval <- true
                    // (What reads it was calculated with the value it had before the edit.)
                    node.BaselinePass <- pass
                    node.Baseline <- effectiveValue was)
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
        // One insertion or deletion of rows / columns: every formula that reads a cell
        // that moved was rewritten to follow it (it was wired again above, and gives the
        // value it gave - see `followed` - or is calculated). What is left to calculate
        // reads the sheet at a place that stayed while the cells under it moved - a whole
        // column, a range in another sheet's name - : of what read the sheet at or
        // beyond the edit, the formulas that were not wired again.
        for struct (reader, stamp) in moveReaders do
            if reader.IsFormula && reader.WireStamp = stamp then
                reach reader
                reader.NeedsEval <- true
        if true then
            // ---- the dynamic arrays the edit touches ----------------------------------
            /// Formulas that show #SPILL! (or may), by sheet: a cell changed or given up in their way lays them out again.
            let blockedBySheet = Dictionary<SheetId, ResizeArray<Node>>()
            let noteBlocked (node: Node) =
                match blockedBySheet.TryGetValue node.Sheet.Id with
                | true, nodes -> if not (nodes.Contains node) then nodes.Add node
                | false, _ ->
                    let nodes = ResizeArray<Node>()
                    nodes.Add node
                    blockedBySheet.[node.Sheet.Id] <- nodes
            let spillError = CellValue.Error Err.spill
            for node in engine.Spilling do
                match Map.tryFind { Row = node.Row; Col = node.Col } (cellsOf node.Sheet) with
                | Some cell when unwrapFormulaValue cell.Value = spillError -> noteBlocked node
                | _ -> ()
            /// The blocked formulas whose array reaches a cell (as far as its size is known).
            let blockedOver (state: SheetState) (row: int) (col: int) (onNode: Node -> unit) =
                match blockedBySheet.TryGetValue state.Id with
                | true, nodes ->
                    for node in nodes do
                        if node.Row <= row && node.Col <= col && not (node.Row = row && node.Col = col)
                           && (node.ExtentRows < 0 || (row < node.Row + node.ExtentRows && col < node.Col + node.ExtentCols)) then
                            onNode node
                | false, _ -> ()
            let seedSpill (node: Node) =
                reach node
                node.NeedsSpill <- true
            let seedAnchor (state: SheetState) (anchor: CellAddress) = seedSpill (nodeAt state anchor.Row anchor.Col)
            // Rows or columns of a sheet moved: every array of the sheet is laid out again.
            for sheetId in surroundings.Structural do
                match engine.Sheets.TryGetValue sheetId, Map.tryFind sheetId edited.Sheets with
                | (true, state), Some sheet ->
                    if not (isNull (box state.Model)) then
                        for KeyValue (_, anchor) in loadedSpill state.Model do seedAnchor state anchor
                    for KeyValue (_, anchor) in loadedSpill sheet.Model do seedAnchor state anchor
                    for node in List.ofSeq engine.Spilling do
                        if obj.ReferenceEquals(node.Sheet, state) then seedSpill node
                | _ -> ()
            let anySpill =
                engine.Spilling.Count > 0 || (edited.Sheets |> Map.exists (fun _ sheet -> not (Map.isEmpty (loadedSpill sheet.Model))))
                || (engine.Sheets.Values |> Seq.exists (fun state -> not (isNull (box state.Model)) && not (Map.isEmpty (loadedSpill state.Model))))
            if anySpill then
                // Spilled cells that are another formula's now, or nobody's.
                for KeyValue (sheetId, sheet) in edited.Sheets do
                    match engine.Sheets.TryGetValue sheetId with
                    | true, state when not (isNull (box state.Model)) && not (obj.ReferenceEquals(state.Model, sheet.Model)) ->
                        let held = loadedSpill state.Model
                        let now = loadedSpill sheet.Model
                        if not (obj.ReferenceEquals(held, now)) && not (Map.isEmpty held && Map.isEmpty now) then
                            diffSpill held now (seedAnchor state)
                    | _ -> ()
                for change in changes do
                    let state = change.State
                    let addr = { Row = change.Row; Col = change.Col }
                    let now = spillOf state
                    if not (Map.isEmpty now) then
                        // The formula of a changed spilled cell; a changed cell that has spilled cells of its own.
                        match Map.tryFind addr now with
                        | Some anchor -> seedAnchor state anchor
                        | None -> ()
                        if (ghostsOf state now).ContainsKey (keyOf change.Row change.Col) then seedAnchor state addr
                    let held = loadedSpill state.Model
                    if not (Map.isEmpty held) then
                        match Map.tryFind addr held with
                        | Some anchor -> seedAnchor state anchor
                        | None -> ()
                    blockedOver state change.Row change.Col seedSpill
                for struct (state, addr) in restyled do
                    match Map.tryFind addr (spillOf state) with
                    | Some anchor -> seedAnchor state anchor
                    | None -> ()

            for node in engine.Volatile do
                reach node
                node.NeedsEval <- true
            if full || scan then
                // Every formula, and every array there is or was.
                for KeyValue (_, state) in engine.Sheets do
                    for node in List.ofSeq state.Nodes.Values do
                        if node.IsFormula then
                            reach node
                            node.NeedsEval <- full
                    if full && not (isNull (box state.Model)) then
                        for KeyValue (_, anchor) in loadedSpill state.Model do seedAnchor state anchor

            /// Formulas found on a circular reference show the cycle error, or - with
            /// iterative calculation - are calculated in sweeps.
            let iteration = edited.Iteration
            let isCyclic (node: Node) = iteration.IsNone && node.CyclicPass >= firstPass
            if mode = 0 && iteration.IsSome then
                // Excel calculates every circular reference at every calculation: one
                // that does not settle moves on by the workbook's number of iterations.
                engine.EdgeMark <- engine.EdgeMark + 1
                let mark = engine.EdgeMark
                let live = ResizeArray<Node>()
                for node in engine.Cyclic do
                    if node.OnCycle && node.IsFormula && node.EdgeMark <> mark then
                        node.EdgeMark <- mark
                        live.Add node
                        // (One that had settled gives what it has when nothing it reads changed:
                        // the sweep over it would be dropped. It is not calculated for nothing.)
                        if node.Unsettled then
                            reach node
                            node.NeedsEval <- true
                engine.Cyclic <- live
            /// The place of a sheet among the tabs: the formulas of a circular reference are calculated in reading order.
            let sheetPlaces = Dictionary<SheetId, int>()
            let sheetPlace (sheetId: SheetId) =
                if sheetPlaces.Count = 0 then edited.SheetOrder |> List.iteri (fun place id -> sheetPlaces.[id] <- place)
                match sheetPlaces.TryGetValue sheetId with
                | true, place -> place
                | false, _ -> System.Int32.MaxValue
            let mutable cyclicCount = 0
            // Precedents chosen while evaluating: see the doc comment.
            let hasDynamic = engine.Dynamic.Count > 0
            /// Every such formula is part of the round in hand / one of them was evaluated in it /
            /// a value changed that one of them may have read before.
            let mutable dynamicSeeded = true
            let mutable dynamicSeen = false
            let mutable dynamicStale = false
            let valueChanged () =
                if hasDynamic && (dynamicSeen || not dynamicSeeded) then dynamicStale <- true

            let mutable reachedTotal = 0
            let mutable evaluated = 0
            let mutable failure : string = null
            let mutable rounds = 0
            let mutable nextEval = ResizeArray<Node>()
            let mutable nextSpill = ResizeArray<Node>()
            let mutable running = true
            while running do
                rounds <- rounds + 1
                // Everything that reads what was reached - a cell itself or, of a
                // formula that spilled, one of its spilled cells - with the edges among them.
                while pending.Count > 0 do
                    let node = pending.Pop()
                    let outs = ResizeArray<Node>()
                    engine.EdgeMark <- engine.EdgeMark + 1
                    let mark = engine.EdgeMark
                    addReadersMarked mark node.Sheet node.Row node.Col node.Dependents outs
                    let spill = spillOf node.Sheet
                    if not (Map.isEmpty spill) then
                        match (ghostsOf node.Sheet spill).TryGetValue (keyOf node.Row node.Col) with
                        | true, ghosts ->
                            for ghost in ghosts do
                                let explicit =
                                    match node.Sheet.Nodes.TryGetValue (keyOf ghost.Row ghost.Col) with
                                    | true, ghostNode -> ghostNode.Dependents
                                    | false, _ -> null
                                addReadersMarked mark node.Sheet ghost.Row ghost.Col explicit outs
                        | false, _ -> ()
                    node.Outs <- outs
                    for reader in outs do
                        reach reader
                        reader.Pending <- reader.Pending + 1
                reachedTotal <- reachedTotal + reached.Count
                // Dependency order: a formula is ready when everything reached that it reads is done.
                let ready = Stack<Node>()
                for node in reached do
                    if node.Pending = 0 then ready.Push node
                let order = ResizeArray<Node>(reached.Count)
                let drain () =
                    while ready.Count > 0 do
                        let node = ready.Pop()
                        if node.Order < 0 then
                            node.Order <- order.Count
                            order.Add node
                            for reader in node.Outs do
                                reader.Pending <- reader.Pending - 1
                                if reader.Pending = 0 && reader.Order < 0 then ready.Push reader
                drain ()
                /// The circular references of the round, as stretches of the order: first index -> last.
                let blocks = Dictionary<int, int>()
                if order.Count < reached.Count then
                    // What is left reads a cycle or is on one. On one: the strongly
                    // connected components of more than one formula, and formulas that
                    // read themselves (Tarjan's algorithm, without recursion). It gives
                    // the components readers first: taken from the last to the first they
                    // are the rest of the order.
                    let components = ResizeArray<ResizeArray<Node>>()
                    let circular = ResizeArray<bool>()
                    let stacked = Stack<Node>()
                    let work = Stack<struct (Node * int)>()
                    let mutable counter = 0
                    for root in reached do
                        if root.Order < 0 && root.SccPass <> pass then
                            work.Push (struct (root, 0))
                            while work.Count > 0 do
                                let struct (node, start) = work.Pop()
                                if start = 0 then
                                    node.SccPass <- pass
                                    node.SccIndex <- counter
                                    node.SccLow <- counter
                                    counter <- counter + 1
                                    stacked.Push node
                                    node.SccOn <- true
                                else
                                    let child = node.Outs.[start - 1]
                                    if child.SccLow < node.SccLow then node.SccLow <- child.SccLow
                                let mutable next = start
                                let mutable descended = false
                                while not descended && next < node.Outs.Count do
                                    let out = node.Outs.[next]
                                    next <- next + 1
                                    if out.Order < 0 && out.Pass = pass then
                                        if obj.ReferenceEquals(out, node) then node.SelfPass <- pass
                                        if out.SccPass <> pass then
                                            work.Push (struct (node, next))
                                            work.Push (struct (out, 0))
                                            descended <- true
                                        elif out.SccOn && out.SccIndex < node.SccLow then node.SccLow <- out.SccIndex
                                if not descended && node.SccLow = node.SccIndex then
                                    let members = ResizeArray<Node>()
                                    let mutable popping = true
                                    while popping do
                                        let memberNode = stacked.Pop()
                                        memberNode.SccOn <- false
                                        members.Add memberNode
                                        if obj.ReferenceEquals(memberNode, node) then popping <- false
                                    // (In the order they were met: a formula before most of what reads it.)
                                    members.Reverse()
                                    components.Add members
                                    circular.Add (members.Count > 1 || node.SelfPass = pass)
                    for at = components.Count - 1 downto 0 do
                        let first = order.Count
                        // (A circular reference in the same order whatever led the pass to it:
                        // sheet by sheet, row by row. Its sweeps then give the same numbers.)
                        let members =
                            if circular.[at] && components.[at].Count > 1 then
                                components.[at] |> Seq.sortBy (fun node -> sheetPlace node.Sheet.Id, node.Row, node.Col) |> Seq.toArray
                            else components.[at].ToArray()
                        for node in members do
                            node.Order <- order.Count
                            order.Add node
                            if circular.[at] then
                                node.CircularRound <- pass
                                // (Every formula of a circular reference shows it, whichever the edit reached first.)
                                node.NeedsEval <- true
                                if not node.OnCycle then
                                    node.OnCycle <- true
                                    engine.Cyclic.Add node
                                if node.CyclicPass < firstPass then
                                    node.CyclicPass <- pass
                                    cyclicCount <- cyclicCount + 1
                        if circular.[at] then blocks.[first] <- order.Count - 1
                // (Reached and on no circular reference: one it was on is broken.)
                for node in reached do
                    if node.OnCycle && node.CircularRound <> pass then node.OnCycle <- false
                // (Fewer in order than reached: a cycle among the formulas to evaluate.)
                if order.Count < reached.Count then
                    failure <- "the formulas to evaluate could not be put in order"
                    running <- false
                elif scan then running <- false
                else
                    nextEval <- ResizeArray<Node>()
                    nextSpill <- ResizeArray<Node>()
                    let mutable index = 0
                    /// A formula has to be evaluated (again): later in this round when it is still to come, else in the next.
                    let evalLater (reader: Node) =
                        if reader.Pass = pass && reader.Order > index then reader.NeedsEval <- true
                        else nextEval.Add reader
                    let spillLater (node: Node) =
                        if node.Pass = pass && node.Order > index then node.NeedsSpill <- true
                        elif not (node.Pass = pass && node.Order = index) then nextSpill.Add node
                    let cellReaders = ResizeArray<Node>()
                    let readersChanged (state: SheetState) (cell: CellAddress) =
                        let explicit =
                            match state.Nodes.TryGetValue (keyOf cell.Row cell.Col) with
                            | true, cellNode -> cellNode.Dependents
                            | false, _ -> null
                        cellReaders.Clear()
                        addReaders engine state cell.Row cell.Col explicit cellReaders
                        for reader in cellReaders do evalLater reader
                    /// Lay out the dynamic array of a cell: evaluate it, write the cells it
                    /// spills into, clear the ones it no longer covers. True when it is blocked.
                    let layOut (node: Node) : bool =
                        let state = node.Sheet
                        let model = modelOf state
                        let addr = { Row = node.Row; Col = node.Col }
                        let key = keyOf node.Row node.Col
                        let ghosts = ghostsOf state model.Spill
                        let old = match ghosts.TryGetValue key with | true, cells -> cells.ToArray() | false, _ -> [||]
                        let mutable cells = model.Cells
                        let mutable spill = model.Spill
                        let mutable blocked = false
                        let mutable layoutChanged = false
                        let covered = HashSet<float>()
                        let current = ResizeArray<CellAddress>()
                        let anchorCell =
                            match Map.tryFind addr cells with
                            | Some cell when node.IsFormula && node.Spills && node.Expr.IsSome && not (isCyclic node) ->
                                (match cell.Value with CellValue.Formula _ -> Some cell | _ -> None)
                            | _ -> None
                        match anchorCell with
                        | Some cell ->
                            let arr = evalArray node.Expr.Value { Model = model; Workbook = Some workbook; Sheet = Some state.Id; CurrentCell = Some addr }
                            let nr = arr.Length
                            let nc = if nr > 0 then arr.[0].Length else 0
                            node.ExtentRows <- nr
                            node.ExtentCols <- nc
                            if nr * nc > 1 then
                                // In the way: a cell of the user's, or a cell of an array further up / left.
                                let mutable r = 0
                                while not blocked && r < nr do
                                    let mutable c = 0
                                    while not blocked && c < nc do
                                        if not (r = 0 && c = 0) then
                                            let target = { Row = addr.Row + r; Col = addr.Col + c }
                                            match Map.tryFind target spill with
                                            | Some owner when owner = addr -> ()
                                            | Some owner ->
                                                let ownerLive =
                                                    match state.Nodes.TryGetValue (keyOf owner.Row owner.Col) with
                                                    | true, ownerNode -> ownerNode.IsFormula && ownerNode.Spills
                                                    | false, _ -> false
                                                if ownerLive && (owner.Row < addr.Row || (owner.Row = addr.Row && owner.Col < addr.Col)) then blocked <- true
                                            | None ->
                                                match Map.tryFind target cells with
                                                | Some existing when existing.Value <> CellValue.Empty -> blocked <- true
                                                | _ -> ()
                                        c <- c + 1
                                    r <- r + 1
                                if not blocked then
                                    for r in 0 .. nr - 1 do
                                        for c in 0 .. nc - 1 do
                                            if not (r = 0 && c = 0) then
                                                let target = { Row = addr.Row + r; Col = addr.Col + c }
                                                let next = { Value = arr.[r].[c]; Format = cell.Format; Comment = None }
                                                let existing = Map.tryFind target cells
                                                match existing with
                                                | Some known when known = next -> ()
                                                | _ ->
                                                    cells <- Map.add target next cells
                                                    let before = match existing with Some known -> unwrapFormulaValue known.Value | None -> CellValue.Empty
                                                    if before <> unwrapFormulaValue next.Value then
                                                        layoutChanged <- true
                                                        readersChanged state target
                                                match Map.tryFind target spill with
                                                | Some owner when owner = addr -> ()
                                                | Some owner ->
                                                    // Taken from the array of a formula further down / right: that one is blocked now.
                                                    match ghosts.TryGetValue (keyOf owner.Row owner.Col) with
                                                    | true, theirs -> theirs.Remove target |> ignore
                                                    | false, _ -> ()
                                                    spill <- Map.add target addr spill
                                                    layoutChanged <- true
                                                    spillLater (nodeAt state owner.Row owner.Col)
                                                | None ->
                                                    spill <- Map.add target addr spill
                                                    layoutChanged <- true
                                                covered.Add (keyOf target.Row target.Col) |> ignore
                                                current.Add target
                        | None -> ()
                        // The cells it spilled into and no longer covers.
                        for ghost in old do
                            if not (covered.Contains (keyOf ghost.Row ghost.Col)) then
                                spill <- Map.remove ghost spill
                                layoutChanged <- true
                                match Map.tryFind ghost cells with
                                | Some known ->
                                    match known.Value with
                                    | CellValue.Formula _ -> ()
                                    | value ->
                                        cells <- Map.remove ghost cells
                                        if value <> CellValue.Empty then readersChanged state ghost
                                | None -> ()
                                // A blocked formula that wants the cell can have it.
                                blockedOver state ghost.Row ghost.Col (fun waiting -> if not (obj.ReferenceEquals(waiting, node)) then spillLater waiting)
                        if current.Count > 0 then ghosts.[key] <- current else ghosts.Remove key |> ignore
                        if not (obj.ReferenceEquals(cells, model.Cells)) || not (obj.ReferenceEquals(spill, model.Spill)) then
                            if not (obj.ReferenceEquals(cells, model.Cells)) then valueChanged ()
                            let updated = { model with Cells = cells; Spill = spill }
                            models.[state.Id] <- updated
                            touched.Add state.Id |> ignore
                            workbook <- updateSheetModel state.Id updated workbook
                        state.GhostsFor <- box spill
                        if blocked then noteBlocked node
                        // What reads the array as a whole (`A1#`) sees another one: other cells, or other values in them.
                        if layoutChanged then
                            for reader in node.Dependents do evalLater reader
                        blocked
                    while index < order.Count do
                        // A circular reference of a workbook with iterative calculation.
                        (match iteration with
                         | Some setting when blocks.ContainsKey index ->
                            let last = blocks.[index]
                            let cachedOf (node: Node) =
                                match Map.tryFind { Row = node.Row; Col = node.Col } (modelOf node.Sheet).Cells with
                                | Some cell -> (match cell.Value with CellValue.Formula (_, cached) -> cached | other -> other)
                                | None -> CellValue.Empty
                            let before = [| for at in index .. last -> cachedOf order.[at] |]
                            // (A formula that showed the cycle error - iteration was off - starts from nothing, as in Excel.)
                            for at in index .. last do
                                let node = order.[at]
                                if node.IsFormula && before.[at - index] = CellValue.Error Err.cycle then
                                    let model = modelOf node.Sheet
                                    let addr = { Row = node.Row; Col = node.Col }
                                    match Map.tryFind addr model.Cells with
                                    | Some cell ->
                                        let updated = { model with Cells = Map.add addr { cell with Value = CellValue.Formula (node.Formula, CellValue.Empty) } model.Cells }
                                        models.[node.Sheet.Id] <- updated
                                        touched.Add node.Sheet.Id |> ignore
                                        workbook <- updateSheetModel node.Sheet.Id updated workbook
                                    | None -> ()
                            let mutable sweeps = 0
                            let mutable going = true
                            while going && sweeps < max 1 setting.MaxIterations do
                                sweeps <- sweeps + 1
                                let heldWorkbook = workbook
                                let heldModels = [| for KeyValue (sheetId, held) in models -> sheetId, held |]
                                let mutable change = 0.0
                                for at in index .. last do
                                    let node = order.[at]
                                    if node.IsFormula then
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
                                                change <-
                                                    match previous, result with
                                                    | CellValue.Number a, CellValue.Number b -> max change (abs (a - b))
                                                    | _ -> System.Double.PositiveInfinity
                                                let updated = { model with Cells = Map.add addr { cell with Value = CellValue.Formula (node.Formula, result) } model.Cells }
                                                models.[state.Id] <- updated
                                                touched.Add state.Id |> ignore
                                                workbook <- updateSheetModel state.Id updated workbook
                                        | None -> ()
                                if change < setting.MaxChange || change = 0.0 then
                                    // Settled: the sweep's last small steps are not kept.
                                    going <- false
                                    if change > 0.0 then
                                        workbook <- heldWorkbook
                                        models.Clear()
                                        for sheetId, held in heldModels do models.[sheetId] <- held
                            for at in index .. last do
                                let node = order.[at]
                                node.NeedsEval <- false
                                node.Unsettled <- going
                                if node.Dynamic then dynamicSeen <- true
                                // (What reads a formula the edit changed was calculated with the value it had before.)
                                let was = if node.BaselinePass = pass && pass = firstPass then node.Baseline else before.[at - index]
                                if cachedOf node <> was then
                                    valueChanged ()
                                    for reader in node.Outs do
                                        if reader.Order > last then reader.NeedsEval <- true
                         | _ -> ())
                        let node = order.[index]
                        let lays = node.NeedsSpill || (node.NeedsEval && node.IsFormula && node.Spills)
                        let blocked = if lays then layOut node else false
                        if (node.NeedsEval || lays) && node.IsFormula then
                            evaluated <- evaluated + 1
                            let state = node.Sheet
                            let model = modelOf state
                            let addr = { Row = node.Row; Col = node.Col }
                            match Map.tryFind addr model.Cells with
                            | Some cell ->
                                let result =
                                    if isCyclic node then CellValue.Error Err.cycle
                                    elif blocked then spillError
                                    else
                                        match node.Expr with
                                        | Some expr -> eval expr { Model = model; Workbook = Some workbook; Sheet = Some state.Id; CurrentCell = Some addr }
                                        | None -> CellValue.Error Err.syntax
                                let previous = match cell.Value with CellValue.Formula (_, cached) -> cached | other -> other
                                if previous <> result then
                                    valueChanged ()
                                    let updated = { model with Cells = Map.add addr { cell with Value = CellValue.Formula (node.Formula, result) } model.Cells }
                                    models.[state.Id] <- updated
                                    touched.Add state.Id |> ignore
                                    workbook <- updateSheetModel state.Id updated workbook
                                // What reads it was calculated with the value it had before the edit.
                                let before = if node.BaselinePass = pass && pass = firstPass then node.Baseline else previous
                                if before <> result then
                                    for reader in node.Outs do
                                        reader.NeedsEval <- true
                            | None -> ()
                            if node.Dynamic then dynamicSeen <- true
                        index <- index + 1
                    if nextEval.Count = 0 && nextSpill.Count = 0 && not dynamicStale then running <- false
                    elif rounds >= 200 then
                        // The full pass has no one to hand the workbook to: it stops with what it has.
                        // (This is the one bound left in it. The layout of a dynamic array can depend
                        // on the layout of another - one takes a cell, another is blocked and gives
                        // up cells a third wants - and OFFSET / INDIRECT can read what was calculated
                        // after them; both are settled by rounds over the formulas concerned, which
                        // end unless the dependence is circular.)
                        if not full then
                            failure <-
                                if nextEval.Count = 0 && nextSpill.Count = 0 then "formulas that choose their precedents (OFFSET, INDIRECT) do not settle (a circular reference through them)"
                                else "the dynamic arrays do not settle (a circular reference through spilled cells)"
                        running <- false
                    else
                        newRound ()
                        dynamicSeeded <- dynamicStale
                        dynamicSeen <- false
                        if dynamicStale then
                            dynamicStale <- false
                            for node in engine.Dynamic do
                                if not node.Random then
                                    reach node
                                    node.NeedsEval <- true
                        for node in nextEval do
                            reach node
                            node.NeedsEval <- true
                        for node in nextSpill do
                            reach node
                            node.NeedsSpill <- true
            if not (isNull failure) then decline failure
            elif scan then
                engine.CyclesKnown <- true
                Some edited
            else
                // As the full recalculation leaves a sheet whose formulas gave other values.
                for sheetId in touched do
                    let model = models.[sheetId]
                    // (The formula index of the sheet goes on to its recalculated cells.)
                    match cellsAtStart.TryGetValue sheetId with
                    | true, before -> inheritFormulaIndex before model.Cells
                    | false, _ -> ()
                    let projected = invalidateFilterProjection model
                    if not (obj.ReferenceEquals(model, projected)) then
                        workbook <- updateSheetModel sheetId projected workbook
                holdWorkbook engine workbook
                engine.LastStats <- { ChangedCells = changes.Count; ReachedFormulas = reachedTotal; EvaluatedFormulas = evaluated }
                if full then
                    engine.LastCyclic <- cyclicCount
                    engine.CyclesKnown <- true
                Some workbook
        else None

/// Look for the circular references of the graph (no formula is evaluated).
/// A pass over the whole workbook knows them; an engine that was built for a
/// calculated workbook (`build`) does not, and takes the circular references
/// of that workbook as settled until an edit reaches one. After this, in a
/// workbook with iterative calculation, every one of them is calculated at
/// the next edit - as the full recalculation, and Excel, do at every
/// calculation - and from then on those that did not settle.
let findCycles (engine: Engine) : unit =
    if not engine.CyclesKnown then run 2 engine engine.Workbook None |> ignore

/// See `run`.
let recalculate (engine: Engine) (edited: Workbook) (hint: (SheetId * CellAddress) list option) : Workbook option =
    run 0 engine edited hint

/// The full recalculation of a workbook: every formula of every sheet (of the
/// sheets that are converted; the stand-in of a sheet nothing asked for has no
/// cells to calculate), in one dependency order across the sheets - see `run`.
/// Gives the recalculated workbook and the engine that now holds it, for the
/// edits that follow.
let recalculateAll (workbook: Workbook) : Workbook * Engine =
    let engine = buildGraph false workbook
    // (Nothing to calculate: no formula, and no cell left behind by an array.)
    let hasWork = engine.Formulas > 0 || (workbook.Sheets |> Map.exists (fun _ sheet -> not (Map.isEmpty (loadedSpill sheet.Model))))
    if not hasWork then workbook, engine
    else
        match run 1 engine workbook None with
        | Some recalculated -> recalculated, engine
        | None -> workbook, engine

/// Formula.recalculateWorkbook runs `recalculateAll` once this is read (it
/// reads it itself on .NET; in the browser the module is loaded with the grid).
let workbookPassInstalled : bool =
    installWorkbookPass { Run = fun workbook -> fst (recalculateAll workbook) }
    true
