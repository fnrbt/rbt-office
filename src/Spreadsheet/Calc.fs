module RedBat.Spreadsheet.Calc

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Workbook
open RedBat.Spreadsheet.Formula

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
    /// Cells that reference this cell (may live on other sheets).
    Dependents: Set<SheetCell>
    /// Is this cell marked for recalculation?
    IsDirty: bool
    /// Formula AST (parsed once, cached).
    ParsedFormula: Expr option
}

/// Calculation engine state.
type CalcEngine = {
    /// Metadata for each formula cell, across every sheet.
    Metadata: Map<SheetCell, CalcMetadata>
    /// Pre-computed calculation order (topological sort).
    CalcChain: SheetCell list
    /// Cells marked dirty (need recalculation).
    DirtySet: Set<SheetCell>
}

let emptyMetadata = {
    Dependencies = Set.empty
    Dependents = Set.empty
    IsDirty = false
    ParsedFormula = None
}

let emptyEngine = {
    Metadata = Map.empty
    CalcChain = []
    DirtySet = Set.empty
}

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

/// Build (or rebuild) the dependency edges for one formula cell.
let buildDependencies (cell: SheetCell) (formula: string)
                      (workbook: Workbook) (engine: CalcEngine): CalcEngine =
    match parseFormula formula with
    | Ok expr ->
        let deps = extractReferences (Some workbook) cell.Sheet expr
        let metadata =
            engine.Metadata
            |> Map.tryFind cell
            |> Option.defaultValue emptyMetadata

        // Drop this cell from the dependents list of every cell it used to
        // reference, before wiring up the new edges.
        let mutable updatedMetadata = engine.Metadata
        for oldDep in metadata.Dependencies do
            match Map.tryFind oldDep updatedMetadata with
            | Some oldDepMeta ->
                updatedMetadata <-
                    Map.add oldDep { oldDepMeta with Dependents = Set.remove cell oldDepMeta.Dependents } updatedMetadata
            | None -> ()

        let newMetadata = {
            metadata with
                Dependencies = deps
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

        { engine with Metadata = updatedMetadata }
    | Result.Error _ ->
        engine

/// Topological sort (DFS). Returns calculation order — dependencies first
/// — or None when the graph has a cycle.
let topologicalSort (nodes: SheetCell list) (metadata: Map<SheetCell, CalcMetadata>): SheetCell list option =
    let mutable visited = Set.empty
    let mutable recStack = Set.empty
    let mutable result = []
    let mutable hasCycle = false

    let rec visit (node: SheetCell) =
        if recStack.Contains node then
            hasCycle <- true
        elif not (visited.Contains node) then
            visited <- Set.add node visited
            recStack <- Set.add node recStack
            match Map.tryFind node metadata with
            | Some meta -> for dep in meta.Dependencies do visit dep
            | None -> ()
            recStack <- Set.remove node recStack
            result <- node :: result

    for node in nodes do
        if not hasCycle then visit node

    if hasCycle then None else Some (List.rev result)

/// Rebuild the calculation chain — a topological sort over every formula
/// cell in the workbook.
let rebuildCalcChain (workbook: Workbook) (engine: CalcEngine): CalcEngine =
    let formulaCells =
        [ for KeyValue(sheetId, sheet) in workbook.Sheets do
            for KeyValue(addr, cell) in sheet.Model.Cells do
                match cell.Value with
                | CellValue.Formula _ -> yield { Sheet = sheetId; Cell = addr }
                | _ -> () ]
    match topologicalSort formulaCells engine.Metadata with
    | Some chain -> { engine with CalcChain = chain }
    | None ->
        // Circular reference — clear the chain to signal the error state.
        { engine with CalcChain = [] }

/// Mark a cell and everything transitively downstream of it dirty.
let rec markDirty (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    let metadata =
        engine.Metadata
        |> Map.tryFind cell
        |> Option.defaultValue emptyMetadata

    if metadata.IsDirty then
        engine  // already dirty — its dependents were marked too
    else
        let mutable result =
            { engine with
                Metadata = Map.add cell { metadata with IsDirty = true } engine.Metadata
                DirtySet = Set.add cell engine.DirtySet }
        for dependent in metadata.Dependents do
            result <- markDirty dependent result
        result

/// Clear the dirty flag for a cell.
let clearDirty (cell: SheetCell) (engine: CalcEngine): CalcEngine =
    match Map.tryFind cell engine.Metadata with
    | Some metadata ->
        { engine with
            Metadata = Map.add cell { metadata with IsDirty = false } engine.Metadata
            DirtySet = Set.remove cell engine.DirtySet }
    | None -> engine

/// Recalculate a single cell from its cached formula AST, evaluating
/// against its own sheet's model with the full workbook in context.
let recalculateCell (cell: SheetCell) (workbook: Workbook) (engine: CalcEngine): CellValue option =
    match Map.tryFind cell engine.Metadata with
    | Some metadata ->
        match metadata.ParsedFormula, sheetModel cell.Sheet workbook with
        | Some expr, Some model ->
            let ctx : EvalContext =
                { Model = model; Workbook = Some workbook; Sheet = Some cell.Sheet }
            Some (eval expr ctx)
        | _ -> None
    | None -> None

/// Smart recalculation: recompute only dirty cells, in dependency order,
/// across every sheet. Volatile cells (RAND/NOW/OFFSET/…) are forced dirty
/// every pass so their values — and anything downstream — stay live.
let smartRecalculate (workbook: Workbook) (engine: CalcEngine): Workbook * CalcEngine =
    let engine =
        engine.Metadata
        |> Map.fold (fun eng cell meta ->
            match meta.ParsedFormula with
            | Some expr when containsVolatile expr -> markDirty cell eng
            | _ -> eng) engine
    if Set.isEmpty engine.DirtySet then
        (workbook, engine)
    else
        let mutable wb = workbook
        let mutable eng = engine
        for cell in engine.CalcChain do
            match Map.tryFind cell eng.Metadata with
            | Some metadata when metadata.IsDirty ->
                match recalculateCell cell wb eng with
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
                    eng <- clearDirty cell eng
                | None -> ()
            | _ -> ()
        // Anything still dirty after the chain pass is a non-formula source
        // cell (the *edited* cell itself, which has nothing to recompute) or
        // — when the graph has a cycle — an unresolvable formula. Either way
        // clear it: leaving a stale dirty flag would make the next edit's
        // markDirty short-circuit and skip propagating to dependents.
        let eng = eng.DirtySet |> Set.fold (fun e c -> clearDirty c e) eng
        (wb, eng)

/// Update a cell's value, rewire its dependency edges, mark dependents
/// dirty, and rebuild the calc chain. Returns the updated workbook + engine.
let updateCell (cell: SheetCell) (value: CellValue) (cellData: CellData)
               (workbook: Workbook) (engine: CalcEngine): Workbook * CalcEngine =
    let newWorkbook =
        match sheetModel cell.Sheet workbook with
        | Some model ->
            let model' = setCell cell.Cell { cellData with Value = value } model
            updateSheetModel cell.Sheet model' workbook
        | None -> workbook

    let newEngine =
        match value with
        | CellValue.Formula (formula, _) ->
            buildDependencies cell formula newWorkbook engine
        | _ ->
            // Not a formula anymore: drop its outgoing edges but keep the
            // incoming ones — other cells may still reference this cell.
            match Map.tryFind cell engine.Metadata with
            | Some metadata ->
                let mutable updatedMetadata = engine.Metadata
                for dep in metadata.Dependencies do
                    match Map.tryFind dep updatedMetadata with
                    | Some depMeta ->
                        updatedMetadata <-
                            Map.add dep { depMeta with Dependents = Set.remove cell depMeta.Dependents } updatedMetadata
                    | None -> ()
                let newMetadata = { metadata with Dependencies = Set.empty; ParsedFormula = None }
                { engine with Metadata = Map.add cell newMetadata updatedMetadata }
            | None -> engine

    let dirtyEngine = markDirty cell newEngine
    let finalEngine = rebuildCalcChain newWorkbook dirtyEngine
    (newWorkbook, finalEngine)

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
    for KeyValue(_, sheet) in workbook.Sheets do
        for KeyValue(_, cell) in sheet.Model.Cells do
            totalCells <- totalCells + 1
            match cell.Value with
            | CellValue.Formula _ -> formulaCells <- formulaCells + 1
            | _ -> ()

    let totalDeps =
        engine.Metadata
        |> Map.toList
        |> List.sumBy (fun (_, meta) -> Set.count meta.Dependencies)

    {
        TotalCells = totalCells
        FormulaCells = formulaCells
        DirtyCells = Set.count engine.DirtySet
        Dependencies = totalDeps
        CalcChainLength = List.length engine.CalcChain
        HasCircularRef = formulaCells > 0 && List.isEmpty engine.CalcChain
    }
