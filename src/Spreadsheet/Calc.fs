module Rbt.Spreadsheet.Calc

open System.Collections.Generic
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Formula

/// Incremental calculation engine with cross-sheet dependency tracking.
///
/// The dependency graph is keyed on `SheetCell` (sheet + cell), so an edge
/// can cross a sheet boundary: `Sheet2!B1 = Sheet1!A1 + 1` makes
/// `Sheet1!A1` a dependency of `Sheet2!B1`, and editing `Sheet1!A1`
/// dirties — and recalculates — `Sheet2!B1`. The engine therefore operates
/// on a whole `Workbook` rather than a single sheet model.

/// Calculation metadata for one formula cell.
type CalcMetadata = {
    /// Cells this cell directly references (may live on other sheets).
    Dependencies: Set<SheetCell>
    /// Large ranges this cell depends on, tracked as rectangles instead of
    /// one edge per cell (so `SUM(A1:A100000)` is one rect, not 100k edges).
    /// A cell edit inside any of these dirties this formula; for scheduling,
    /// formula cells that fall inside a rect become ordering dependencies.
    DepRects: DepRect list
    /// Cells that reference this cell (may live on other sheets).
    Dependents: Set<SheetCell>
    /// Is this cell marked for recalculation?
    IsDirty: bool
    /// Formula AST (parsed once, cached).
    ParsedFormula: Expr option
}

/// What in a workbook the incremental engine cannot calculate by itself:
/// formulas that spill (dynamic arrays, legacy array formulas), which need the
/// spill pass, and formulas whose precedents the dependency graph does not
/// know (3-D references, references to a spill range `A1#`).
type PipelineNeeds = {
    /// A formula spills, or a sheet holds spilled cells.
    Spills: bool
    /// A formula reads through a 3-D reference or a spill-range reference, or
    /// a spilling formula has not been calculated yet.
    UntrackedPrecedents: bool
    /// A formula has no value yet (a workbook that was never calculated).
    Uncalculated: bool
}

/// Calculation engine state.
type CalcEngine = {
    /// Metadata for each formula cell, across every sheet.
    Metadata: Map<SheetCell, CalcMetadata>
    /// Formula cells containing volatile functions, forced dirty every pass.
    VolatileCells: Set<SheetCell>
    /// Pre-computed calculation order (topological sort).
    CalcChain: SheetCell list
    /// Cells marked dirty (need recalculation).
    DirtySet: Set<SheetCell>
    /// What `trySingleCellEdit` found the workbook to need, remembered across
    /// incremental edits (which cannot change it); None until first asked.
    Needs: PipelineNeeds option
}

/// Result of a guarded incremental single-cell recalculation attempt.
type IncrementalRecalcResult =
    | IncrementalApplied of Workbook * CalcEngine
    | IncrementalFallback of string

let emptyMetadata = {
    Dependencies = Set.empty
    DepRects = []
    Dependents = Set.empty
    IsDirty = false
    ParsedFormula = None
}

let emptyEngine = {
    Metadata = Map.empty
    VolatileCells = Set.empty
    CalcChain = []
    DirtySet = Set.empty
    Needs = None
}

let private exprUsesSpillRef (expr: Expr) =
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        stack.RemoveAt last
        match expr with
        | SpillRef _ -> found <- true
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | Function (_, args) ->
            for arg in args do
                stack.Add arg
        | _ -> ()
    found

let private metadataIsEmpty (metadata: CalcMetadata) =
    Set.isEmpty metadata.Dependencies
    && List.isEmpty metadata.DepRects
    && Set.isEmpty metadata.Dependents
    && not metadata.IsDirty
    && Option.isNone metadata.ParsedFormula

let private putMetadata (cell: SheetCell) (metadata: CalcMetadata) (metadataByCell: Map<SheetCell, CalcMetadata>) =
    if metadataIsEmpty metadata then
        Map.remove cell metadataByCell
    else
        Map.add cell metadata metadataByCell

let private resizeArrayToList (items: ResizeArray<'T>) =
    let mutable result = []
    for i = items.Count - 1 downto 0 do
        result <- items.[i] :: result
    result

let private formulaCellsWithRectContaining (cell: SheetCell) (engine: CalcEngine) =
    let matches = ResizeArray<SheetCell>()
    for KeyValue(formulaCell, metadata) in engine.Metadata do
        if not (List.isEmpty metadata.DepRects) then
            let mutable found = false
            let mutable remaining = metadata.DepRects
            while not found && not remaining.IsEmpty do
                match remaining with
                | rect :: tail ->
                    if rectContainsCell rect cell then
                        found <- true
                    remaining <- tail
                | [] -> ()
            if found then
                matches.Add formulaCell
    matches

/// Look up a sheet's model within the workbook.
let private sheetModel (sheetId: SheetId) (workbook: Workbook) : SpreadsheetModel option =
    Map.tryFind sheetId workbook.Sheets |> Option.map (fun s -> s.Model)

/// Extract the cross-sheet references of a formula AST. Delegates to the
/// shared implementation in Formula.fs so the two stay in lock-step as the
/// AST grows. `homeSheet` qualifies unqualified refs; the workbook resolves
/// sheet names and `NameRef`s.
let extractReferences (workbook: Workbook option) (homeSheet: SheetId)
                      (expr: Expr): Set<SheetCell> =
    referencesQualified workbook homeSheet expr

let private clearFormulaDependencies (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    match Map.tryFind cell engine.Metadata with
    | Some metadata ->
        let mutable updatedMetadata = engine.Metadata
        for dep in metadata.Dependencies do
            match Map.tryFind dep updatedMetadata with
            | Some depMeta ->
                let depMeta = { depMeta with Dependents = Set.remove cell depMeta.Dependents }
                updatedMetadata <-
                    putMetadata dep depMeta updatedMetadata
            | None -> ()
        let clearedMetadata = {
            metadata with
                Dependencies = Set.empty
                DepRects = []
                IsDirty = false
                ParsedFormula = None
        }
        { engine with
            Metadata = putMetadata cell clearedMetadata updatedMetadata
            VolatileCells = Set.remove cell engine.VolatileCells
            DirtySet = Set.remove cell engine.DirtySet }
    | None -> engine

/// Build (or rebuild) the dependency edges for one formula cell.
let buildDependencies (cell: SheetCell) (formula: string)
                      (workbook: Workbook) (engine: CalcEngine): CalcEngine =
    match parseFormula formula with
    | Ok expr ->
        // Large ranges become rectangle deps (rects); small ranges + single
        // cells stay explicit edges (deps).
        let deps, rects = referencesRangedAt (Some workbook) cell.Sheet (Some cell.Cell) expr
        let engine = clearFormulaDependencies cell engine
        let metadata =
            engine.Metadata
            |> Map.tryFind cell
            |> Option.defaultValue emptyMetadata

        let mutable updatedMetadata = engine.Metadata
        let newMetadata = {
            metadata with
                Dependencies = deps
                DepRects = rects
                ParsedFormula = Some expr
        }
        updatedMetadata <- Map.add cell newMetadata updatedMetadata

        // Register this cell as a dependent of everything it now references.
        for dep in deps do
            let depMeta =
                updatedMetadata
                |> Map.tryFind dep
                |> Option.defaultValue emptyMetadata
            updatedMetadata <-
                Map.add dep { depMeta with Dependents = Set.add cell depMeta.Dependents } updatedMetadata

        let volatileCells =
            if containsVolatile expr then Set.add cell engine.VolatileCells
            else Set.remove cell engine.VolatileCells
        { engine with
            Metadata = updatedMetadata
            VolatileCells = volatileCells }
    | Result.Error _ ->
        clearFormulaDependencies cell engine

/// Topological sort (DFS). Returns calculation order — dependencies first
/// — or None when the graph has a cycle. `iterDeps` visits a node's ordering
/// dependencies (explicit cell edges plus any rect-implied formula edges).
let topologicalSort (nodes: SheetCell seq) (iterDeps: SheetCell -> (SheetCell -> unit) -> unit): SheetCell list option =
    let visitState = Dictionary<SheetCell, byte>()
    let result = ResizeArray<SheetCell>()
    let mutable hasCycle = false

    let depsFor node =
        let deps = ResizeArray<SheetCell>()
        iterDeps node deps.Add
        deps

    for node in nodes do
        if not hasCycle then
            let mutable state = 0uy
            if not (visitState.TryGetValue(node, &state)) then
                let stack = Stack<struct (SheetCell * ResizeArray<SheetCell> * int)>()
                visitState.[node] <- 1uy
                stack.Push(struct (node, depsFor node, 0))
                while stack.Count > 0 && not hasCycle do
                    let struct (current, deps, index) = stack.Pop()
                    if index >= deps.Count then
                        visitState.[current] <- 2uy
                        result.Add current
                    else
                        stack.Push(struct (current, deps, index + 1))
                        let dep = deps.[index]
                        let mutable depState = 0uy
                        if visitState.TryGetValue(dep, &depState) then
                            if depState = 1uy then
                                hasCycle <- true
                        else
                            visitState.[dep] <- 1uy
                            stack.Push(struct (dep, depsFor dep, 0))

    if hasCycle then None else Some (resizeArrayToList result)

let private dirtySubsetSortMaxCells = 1024

let private dirtyCellsInDependencyOrder (engine: CalcEngine) =
    let dirtySet = engine.DirtySet
    let iterDirtyDeps (sc: SheetCell) (emit: SheetCell -> unit) =
        match Map.tryFind sc engine.Metadata with
        | Some meta ->
            for dep in meta.Dependencies do
                if Set.contains dep dirtySet then
                    emit dep
            if not (List.isEmpty meta.DepRects) then
                for rect in meta.DepRects do
                    for dirtyCell in dirtySet do
                        if rectContainsCell rect dirtyCell then
                            emit dirtyCell
        | None -> ()
    match topologicalSort dirtySet iterDirtyDeps with
    | Some cells -> cells
    | None -> engine.CalcChain

let private recalculateVisitOrder (engine: CalcEngine) =
    if Set.count engine.DirtySet <= dirtySubsetSortMaxCells then
        dirtyCellsInDependencyOrder engine
    else
        engine.CalcChain

/// Rebuild the calculation chain — a topological sort over every formula
/// cell in the workbook.
let rebuildCalcChain (workbook: Workbook) (engine: CalcEngine): CalcEngine =
    let formulaCells = ResizeArray<SheetCell>()
    let addIndexed (indexBySheet: Dictionary<SheetId, Dictionary<int, ResizeArray<SheetCell>>>) sheetId key cell =
        let index =
            match indexBySheet.TryGetValue sheetId with
            | true, existing -> existing
            | false, _ ->
                let created = Dictionary<int, ResizeArray<SheetCell>>()
                indexBySheet.[sheetId] <- created
                created
        match index.TryGetValue key with
        | true, cells -> cells.Add cell
        | false, _ ->
            let cells = ResizeArray<SheetCell>()
            cells.Add cell
            index.[key] <- cells
    let addFormulaCell sheetId addr =
        let formulaCell = { Sheet = sheetId; Cell = addr }
        formulaCells.Add formulaCell
    workbook.Sheets
    |> Map.iter (fun sheetId sheet ->
        sheet.Model.Cells
        |> Map.iter (fun addr cell ->
            match cell.Value with
            | CellValue.Formula _ -> addFormulaCell sheetId addr
            | _ -> ()))
    let mutable formulaRowsBySheet = ValueNone
    let mutable formulaColsBySheet = ValueNone
    let formulaIndexBy keyOf =
        let indexBySheet = Dictionary<SheetId, Dictionary<int, ResizeArray<SheetCell>>>()
        for formulaCell in formulaCells do
            addIndexed indexBySheet formulaCell.Sheet (keyOf formulaCell.Cell) formulaCell
        indexBySheet
    let formulaRowsIndex () =
        match formulaRowsBySheet with
        | ValueSome index -> index
        | ValueNone ->
            let index = formulaIndexBy (fun addr -> addr.Row)
            formulaRowsBySheet <- ValueSome index
            index
    let formulaColsIndex () =
        match formulaColsBySheet with
        | ValueSome index -> index
        | ValueNone ->
            let index = formulaIndexBy (fun addr -> addr.Col)
            formulaColsBySheet <- ValueSome index
            index
    // A formula's ordering deps = its explicit edges plus any *formula* cell
    // that falls inside one of its dependency rectangles (so a large-range
    // aggregate still computes after formula cells it covers — without
    // enumerating the rect's data cells). Formula cells are few, so this scan
    // is cheap even when the rect spans 100k data cells. Rects are sheet-local,
    // so only formula cells from the rect's sheet need to be tested.
    let iterDeps (sc: SheetCell) (emit: SheetCell -> unit) =
        match Map.tryFind sc engine.Metadata with
        | Some meta ->
            for dep in meta.Dependencies do
                emit dep
            if not (List.isEmpty meta.DepRects) then
                for rect in meta.DepRects do
                    let rowCount = rect.R1 - rect.R0 + 1
                    let colCount = rect.C1 - rect.C0 + 1
                    if rowCount <= colCount then
                        let formulaRowsBySheet = formulaRowsIndex ()
                        match formulaRowsBySheet.TryGetValue rect.Sheet with
                        | true, rowIndex ->
                            for row in rect.R0 .. rect.R1 do
                                match rowIndex.TryGetValue row with
                                | true, candidates ->
                                    for formulaCell in candidates do
                                        if formulaCell.Cell.Col >= rect.C0 && formulaCell.Cell.Col <= rect.C1 then
                                            emit formulaCell
                                | false, _ -> ()
                        | false, _ -> ()
                    else
                        let formulaColsBySheet = formulaColsIndex ()
                        match formulaColsBySheet.TryGetValue rect.Sheet with
                        | true, colIndex ->
                            for col in rect.C0 .. rect.C1 do
                                match colIndex.TryGetValue col with
                                | true, candidates ->
                                    for formulaCell in candidates do
                                        if formulaCell.Cell.Row >= rect.R0 && formulaCell.Cell.Row <= rect.R1 then
                                            emit formulaCell
                                | false, _ -> ()
                        | false, _ -> ()
        | None -> ()
    match topologicalSort formulaCells iterDeps with
    | Some chain -> { engine with CalcChain = chain }
    | None ->
        // Circular reference — clear the chain to signal the error state.
        { engine with CalcChain = [] }

/// Build dependency metadata and a calculation chain for every formula in a
/// workbook. Cached formula results are left untouched; only the engine graph
/// is constructed.
let buildEngine (workbook: Workbook): CalcEngine =
    let mutable engine = emptyEngine
    workbook.Sheets
    |> Map.iter (fun sheetId sheet ->
        sheet.Model.Cells
        |> Map.iter (fun addr cell ->
            match cell.Value with
            | CellValue.Formula (formula, _) ->
                engine <- buildDependencies { Sheet = sheetId; Cell = addr } formula workbook engine
            | _ -> ()))
    rebuildCalcChain workbook engine

let private cellDataNeedsFullPipeline (cellData: CellData) =
    match cellData.Value with
    | CellValue.Formula (formula, _) ->
        match parseFormula formula with
        | Ok expr -> exprSpillsStatically expr || exprUsesSpillRef expr || containsThreeDReference expr
        | Result.Error _ -> false
    | _ -> false

let private markDirtyQueued (stack: Stack<SheetCell>) (queued: HashSet<SheetCell>) (engine: CalcEngine): CalcEngine =
    let dirtyCells = ResizeArray<SheetCell>()

    while stack.Count > 0 do
        let cell = stack.Pop()
        match Map.tryFind cell engine.Metadata with
        | Some metadata when not (Set.contains cell engine.DirtySet) ->
            dirtyCells.Add cell
            for dependent in metadata.Dependents do
                if queued.Add dependent then
                    stack.Push dependent
        | _ -> ()

    let dirtySet =
        if dirtyCells.Count = 0 then
            engine.DirtySet
        elif Set.isEmpty engine.DirtySet then
            Set.ofSeq dirtyCells
        else
            Set.union engine.DirtySet (Set.ofSeq dirtyCells)

    { engine with
        DirtySet = dirtySet }

let private markDirtyFrom (cells: SheetCell seq) (engine: CalcEngine): CalcEngine =
    let stack = Stack<SheetCell>()
    let queued = HashSet<SheetCell>()
    for cell in cells do
        if queued.Add cell then
            stack.Push cell
    markDirtyQueued stack queued engine

/// Mark a cell and everything transitively downstream of it dirty.
let markDirty (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    let stack = Stack<SheetCell>()
    let queued = HashSet<SheetCell>()
    queued.Add cell |> ignore
    stack.Push cell
    markDirtyQueued stack queued engine

/// Clear the dirty flag for a cell.
let clearDirty (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    { engine with DirtySet = Set.remove cell engine.DirtySet }

/// Mark a cell dirty AND dirty any formula whose dependency *rectangle*
/// contains it. Per-cell reverse edges (`Dependents`) only exist for small
/// ranges; large ranges are tracked as rectangles with no reverse edges, so a
/// data-cell edit inside one is propagated by this containment scan instead.
/// Used on every edit. The scan skips formulas without rects, so it is cheap
/// unless the workbook actually has large-range aggregates.
let markDirtyForCell (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    let dirtySeeds = ResizeArray<SheetCell>()
    dirtySeeds.Add cell
    for formulaCell in formulaCellsWithRectContaining cell engine do
        dirtySeeds.Add formulaCell
    markDirtyFrom dirtySeeds engine

/// Recalculate a single cell from its cached formula AST, evaluating
/// against its own sheet's model with the full workbook in context.
let recalculateCell (cell: SheetCell) (workbook: Workbook) (engine: CalcEngine): CellValue option =
    match Map.tryFind cell engine.Metadata with
    | Some metadata ->
        match metadata.ParsedFormula, sheetModel cell.Sheet workbook with
        | Some expr, Some model ->
            let ctx : EvalContext =
                { Model = model
                  Workbook = Some workbook
                  Sheet = Some cell.Sheet
                  CurrentCell = Some cell.Cell }
            Some (eval expr ctx)
        | _ -> None
    | None -> None

let private canSkipUnchangedCellUpdate (cell: SheetCell) (targetCellData: CellData)
                                       (workbook: Workbook) (engine: CalcEngine) =
    match sheetModel cell.Sheet workbook with
    | Some model ->
        match getCell cell.Cell model with
        | Some existing when existing = targetCellData ->
            match targetCellData.Value with
            | CellValue.Formula _ ->
                match Map.tryFind cell engine.Metadata with
                | Some metadata ->
                    match metadata.ParsedFormula with
                    | Some _ -> true
                    | None -> false
                | None -> false
            | _ -> true
        | _ -> false
    | None -> false

/// One pass over the dirty cells, in dependency order.
let private recalculateDirty (workbook: Workbook) (engine: CalcEngine): Workbook =
    let mutable wb = workbook
    for cell in recalculateVisitOrder engine do
        match Map.tryFind cell engine.Metadata with
        | Some _ when Set.contains cell engine.DirtySet ->
            match recalculateCell cell wb engine with
            | Some newValue ->
                match sheetModel cell.Sheet wb with
                | Some model ->
                    match getCell cell.Cell model with
                    | Some cellData ->
                        match cellData.Value with
                        | CellValue.Formula (formula, _) ->
                            let model' =
                                setCell cell.Cell
                                    { cellData with Value = CellValue.Formula (formula, newValue) }
                                    model
                            wb <- updateSheetModel cell.Sheet model' wb
                        | _ -> ()
                    | None -> ()
                | None -> ()
            | None -> ()
        | _ -> ()
    wb

/// Smart recalculation: recompute only dirty cells, in dependency order,
/// across every sheet. Volatile cells (RAND/NOW/OFFSET/…) are forced dirty
/// every pass so their values — and anything downstream — stay live.
let smartRecalculate (workbook: Workbook) (engine: CalcEngine): Workbook * CalcEngine =
    let engine =
        if Set.isEmpty engine.VolatileCells then engine
        else markDirtyFrom engine.VolatileCells engine
    if Set.isEmpty engine.DirtySet then
        (workbook, engine)
    else
        let mutable wb = recalculateDirty workbook engine
        // OFFSET / INDIRECT choose their precedents while evaluating, so they
        // have no edges to order them after the cells they read and may have
        // run too early. Run them (and their dependents) once more now that
        // every other dirty cell is up to date.
        let dynamicCells =
            engine.VolatileCells
            |> Set.filter (fun cell ->
                match Map.tryFind cell engine.Metadata with
                | Some metadata ->
                    (match metadata.ParsedFormula with
                     | Some expr -> containsDynamicReference expr
                     | None -> false)
                | None -> false)
        if not (Set.isEmpty dynamicCells) then
            let second = markDirtyFrom dynamicCells { engine with DirtySet = Set.empty }
            wb <- recalculateDirty wb second
        // Anything still dirty after the chain pass is a non-formula source
        // cell (the *edited* cell itself, which has nothing to recompute) or
        // — when the graph has a cycle — an unresolvable formula. Either way
        // clear it: leaving a stale dirty flag would make the next edit's
        // markDirty short-circuit and skip propagating to dependents.
        (wb, { engine with DirtySet = Set.empty })

/// Update a cell's value, rewire its dependency edges, mark dependents
/// dirty, and rebuild the calc chain. Returns the updated workbook + engine.
let updateCell (cell: SheetCell) (value: CellValue) (cellData: CellData)
               (workbook: Workbook) (engine: CalcEngine): Workbook * CalcEngine =
    let targetCellData = { cellData with Value = value }
    if canSkipUnchangedCellUpdate cell targetCellData workbook engine then
        workbook, engine
    else
        let newWorkbook =
            match sheetModel cell.Sheet workbook with
            | Some model ->
                let model' = setCell cell.Cell targetCellData model
                updateSheetModel cell.Sheet model' workbook
            | None -> workbook

        let wasFormula =
            match Map.tryFind cell engine.Metadata with
            | Some metadata -> Option.isSome metadata.ParsedFormula
            | None -> false

        let newEngine, graphChanged =
            match value with
            | CellValue.Formula (formula, _) ->
                buildDependencies cell formula newWorkbook engine, true
            | _ ->
                if wasFormula then
                    // Not a formula anymore: drop its outgoing edges but keep the
                    // incoming ones — other cells may still reference this cell.
                    clearFormulaDependencies cell engine, true
                else
                    engine, false

        let dirtyEngine = markDirtyForCell cell newEngine
        let finalEngine =
            if graphChanged then rebuildCalcChain newWorkbook dirtyEngine
            else dirtyEngine
        (newWorkbook, finalEngine)

let private pipelineNeeds (workbook: Workbook) : PipelineNeeds =
    let mutable spills = false
    let mutable untracked = false
    let mutable uncalculated = false
    for KeyValue(_, sheet) in workbook.Sheets do
        if not (Map.isEmpty sheet.Model.Spill) then spills <- true
        if not untracked then
            for KeyValue(_, cell) in sheet.Model.Cells do
                if not untracked then
                    match cell.Value with
                    | CellValue.Formula (formula, cached) ->
                        if cached = CellValue.Empty then uncalculated <- true
                        match parseFormula formula with
                        | Ok expr ->
                            if exprSpillsStatically expr then
                                spills <- true
                                // A spilling formula that was never calculated has no
                                // spill range yet: the full pipeline has to run first.
                                (match cell.Value with
                                 | CellValue.Formula (_, CellValue.Empty) -> untracked <- true
                                 | _ -> ())
                            if exprUsesSpillRef expr || containsThreeDReference expr then
                                spills <- true
                                untracked <- true
                        | Result.Error _ -> ()
                    | _ -> ()
    { Spills = spills; UntrackedPrecedents = untracked; Uncalculated = uncalculated }

/// Safely apply a plain single-cell edit through the incremental engine.
/// An edit is rejected - so that the caller runs the full recalculation -
/// when it produces a circular calculation chain, hits an engine exception,
/// or needs the dynamic-array / spill pipeline: the entered formula spills,
/// the workbook has formulas whose precedents the graph does not track, or
/// the edit touches a spill range or reaches a spilling formula through its
/// dependents. A workbook that merely contains a spilling or array formula
/// somewhere else is edited incrementally.
let trySingleCellEdit (cell: SheetCell) (cellData: CellData)
                      (workbook: Workbook) (engine: CalcEngine option): IncrementalRecalcResult =
    try
        // Scanning every formula is done once per engine: incremental edits
        // cannot add a spilling formula (that is a fallback) or change a spill range.
        let needs =
            match engine with
            | Some { Needs = Some known } -> known
            | _ -> pipelineNeeds workbook
        let isGhost (target: SheetCell) =
            match sheetModel target.Sheet workbook with
            | Some model -> Map.containsKey target.Cell model.Spill
            | None -> false
        let clearsCell = (match cellData.Value with CellValue.Empty -> true | _ -> false)
        if cellDataNeedsFullPipeline cellData then
            IncrementalFallback "spill or array formula entered"
        elif needs.UntrackedPrecedents then
            IncrementalFallback "3-D or spill-range references, or an uncalculated array formula, present"
        elif needs.Spills && (isGhost cell || clearsCell) then
            // Typing over a spilled cell blocks the spill; emptying a cell may unblock one.
            IncrementalFallback "edit inside a spill range"
        else
            let engine =
                match engine with
                | Some existing -> existing
                | None -> buildEngine workbook
            let updatedWorkbook, updatedEngine =
                updateCell cell cellData.Value cellData workbook engine
            let hasCircularRef =
                let mutable formulaCells = 0
                updatedWorkbook.Sheets
                |> Map.iter (fun _ sheet ->
                    sheet.Model.Cells
                    |> Map.iter (fun _ cell ->
                        match cell.Value with
                        | CellValue.Formula _ -> formulaCells <- formulaCells + 1
                        | _ -> ()))
                formulaCells > 0 && List.isEmpty updatedEngine.CalcChain
            // The cells this edit makes stale (and the volatile ones, which are
            // recalculated on every pass): none of them may be a spilling formula.
            let reachesSpill () =
                let spillsAt (target: SheetCell) =
                    isGhost target
                    || (match Map.tryFind target updatedEngine.Metadata with
                        | Some metadata ->
                            (match metadata.ParsedFormula with
                             | Some expr -> exprSpillsStatically expr
                             | None -> false)
                        | None -> false)
                Set.exists spillsAt updatedEngine.DirtySet || Set.exists spillsAt updatedEngine.VolatileCells
            // Nothing reads the edited cell: there is nothing to recalculate,
            // whatever else the workbook holds (a circular chain included).
            let nothingDependsOnIt =
                Set.forall (fun dirty -> dirty = cell) updatedEngine.DirtySet && Set.isEmpty updatedEngine.VolatileCells
            if hasCircularRef && nothingDependsOnIt && not needs.Uncalculated then
                IncrementalApplied (updatedWorkbook, { updatedEngine with DirtySet = Set.empty; Needs = Some needs })
            elif hasCircularRef then
                IncrementalFallback "circular reference"
            elif needs.Spills && reachesSpill () then
                IncrementalFallback "edit reaches a spill or array formula"
            else
                let recalculatedWorkbook, recalculatedEngine =
                    smartRecalculate updatedWorkbook updatedEngine
                IncrementalApplied (recalculatedWorkbook, { recalculatedEngine with Needs = Some needs })
    with ex ->
        IncrementalFallback ex.Message

/// Statistics about the calculation engine.
type CalcStats = {
    TotalCells: int
    FormulaCells: int
    DirtyCells: int
    Dependencies: int
    CalcChainLength: int
    HasCircularRef: bool
}

let getStats (workbook: Workbook) (engine: CalcEngine): CalcStats =
    let mutable totalCells = 0
    let mutable formulaCells = 0
    workbook.Sheets
    |> Map.iter (fun _ sheet ->
        sheet.Model.Cells
        |> Map.iter (fun _ cell ->
            totalCells <- totalCells + 1
            match cell.Value with
            | CellValue.Formula _ -> formulaCells <- formulaCells + 1
            | _ -> ()))

    let mutable totalDeps = 0
    engine.Metadata
    |> Map.iter (fun _ meta ->
        totalDeps <- totalDeps + Set.count meta.Dependencies)

    {
        TotalCells = totalCells
        FormulaCells = formulaCells
        DirtyCells = Set.count engine.DirtySet
        Dependencies = totalDeps
        CalcChainLength = List.length engine.CalcChain
        HasCircularRef = formulaCells > 0 && List.isEmpty engine.CalcChain
    }
