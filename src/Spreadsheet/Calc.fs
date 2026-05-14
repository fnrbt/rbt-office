module RedBat.Spreadsheet.Calc

open System.Collections.Generic
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

/// Calculation engine with dependency tracking and smart recalculation

/// Dependency graph edge: cell A depends on cell B
type Dependency = {
    From: CellAddress  // The cell that has the formula
    To: CellAddress    // The cell it references
}

/// Calculation metadata for a cell
type CalcMetadata = {
    /// Cells this cell directly references
    Dependencies: Set<CellAddress>
    /// Cells that reference this cell
    Dependents: Set<CellAddress>
    /// Is this cell marked for recalculation?
    IsDirty: bool
    /// Formula AST (parsed once, cached)
    ParsedFormula: Expr option
}

/// Calculation engine state
type CalcEngine = {
    /// Metadata for each cell
    Metadata: Map<CellAddress, CalcMetadata>
    /// Pre-computed calculation order (topological sort)
    CalcChain: CellAddress list
    /// Cells marked dirty (need recalculation)
    DirtySet: Set<CellAddress>
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

/// Extract cell references from formula AST. Delegates to the shared
/// implementation in Formula.fs so the two stay in lock-step as the AST
/// grows.
let extractReferences (expr: Expr): Set<CellAddress> =
    referencesInExpr expr

/// Build dependency graph for a cell with a formula
let buildDependencies (addr: CellAddress) (formula: string) (engine: CalcEngine): CalcEngine =
    match parseFormula formula with
    | Ok expr ->
        let deps = extractReferences expr
        let metadata =
            engine.Metadata
            |> Map.tryFind addr
            |> Option.defaultValue emptyMetadata

        let mutable updatedMetadata = engine.Metadata
        for oldDep in metadata.Dependencies do
            match Map.tryFind oldDep updatedMetadata with
            | Some oldDepMeta ->
                updatedMetadata <-
                    Map.add oldDep { oldDepMeta with Dependents = Set.remove addr oldDepMeta.Dependents } updatedMetadata
            | None -> ()

        let newMetadata = {
            metadata with
                Dependencies = deps
                ParsedFormula = Some expr
        }

        // Update dependents for referenced cells
        updatedMetadata <- Map.add addr newMetadata updatedMetadata

        for dep in deps do
            let depMeta =
                updatedMetadata
                |> Map.tryFind dep
                |> Option.defaultValue emptyMetadata

            let newDepMeta = {
                depMeta with
                    Dependents = Set.add addr depMeta.Dependents
            }
            updatedMetadata <- Map.add dep newDepMeta updatedMetadata

        { engine with Metadata = updatedMetadata }
    | Result.Error _ ->
        engine

/// Topological sort using DFS
/// Returns calculation order (dependencies first) or None if there's a cycle
let topologicalSort (addresses: CellAddress list) (metadata: Map<CellAddress, CalcMetadata>): CellAddress list option =
    let mutable visited = Set.empty
    let mutable recStack = Set.empty
    let mutable result = []
    let mutable hasCycle = false

    let rec visit (addr: CellAddress) =
        if recStack.Contains(addr) then
            hasCycle <- true
        elif not (visited.Contains(addr)) then
            visited <- Set.add addr visited
            recStack <- Set.add addr recStack

            match Map.tryFind addr metadata with
            | Some meta ->
                for dep in meta.Dependencies do
                    visit dep
            | None -> ()

            recStack <- Set.remove addr recStack
            result <- addr :: result

    for addr in addresses do
        if not hasCycle then
            visit addr

    // Reverse to get dependencies-first order
    if hasCycle then None else Some (List.rev result)

/// Rebuild calculation chain (topological sort of all formula cells)
let rebuildCalcChain (model: SpreadsheetModel) (engine: CalcEngine): CalcEngine =
    // Get all cells with formulas
    let formulaCells =
        model.Cells
        |> Map.toList
        |> List.choose (fun (addr, cell) ->
            match cell.Value with
            | CellValue.Formula _ -> Some addr
            | _ -> None
        )

    match topologicalSort formulaCells engine.Metadata with
    | Some chain -> { engine with CalcChain = chain }
    | None ->
        // Circular reference detected - clear calc chain to indicate error
        { engine with CalcChain = [] }

/// Mark a cell and all its dependents as dirty
let rec markDirty (addr: CellAddress) (engine: CalcEngine): CalcEngine =
    let metadata =
        engine.Metadata
        |> Map.tryFind addr
        |> Option.defaultValue emptyMetadata

    if metadata.IsDirty then
        // Already dirty, no need to propagate
        engine
    else
        // Mark this cell dirty
        let newMetadata = { metadata with IsDirty = true }
        let updatedMetadata = Map.add addr newMetadata engine.Metadata
        let updatedDirtySet = Set.add addr engine.DirtySet

        // Recursively mark all dependents
        let mutable result = {
            engine with
                Metadata = updatedMetadata
                DirtySet = updatedDirtySet
        }

        for dependent in metadata.Dependents do
            result <- markDirty dependent result

        result

/// Clear dirty flag for a cell
let clearDirty (addr: CellAddress) (engine: CalcEngine): CalcEngine =
    match Map.tryFind addr engine.Metadata with
    | Some metadata ->
        let newMetadata = { metadata with IsDirty = false }
        { engine with
            Metadata = Map.add addr newMetadata engine.Metadata
            DirtySet = Set.remove addr engine.DirtySet
        }
    | None -> engine

/// Recalculate a single cell using cached formula
let recalculateCell (addr: CellAddress) (model: SpreadsheetModel) (engine: CalcEngine): CellValue option =
    match Map.tryFind addr engine.Metadata with
    | Some metadata ->
        match metadata.ParsedFormula with
        | Some expr ->
            let result = eval expr model None
            Some result
        | None -> None
    | None -> None

/// Smart recalculation: only recalculate dirty cells in correct order
let smartRecalculate (model: SpreadsheetModel) (engine: CalcEngine): SpreadsheetModel * CalcEngine =
    if Set.isEmpty engine.DirtySet then
        // Nothing to recalculate
        (model, engine)
    else
        // Recalculate dirty cells in chain order
        let mutable updatedModel = model
        let mutable updatedEngine = engine

        for addr in engine.CalcChain do
            match Map.tryFind addr engine.Metadata with
            | Some metadata when metadata.IsDirty ->
                // Recalculate this cell
                match recalculateCell addr updatedModel updatedEngine with
                | Some newValue ->
                    // Update cell with new calculated value
                    match getCell addr updatedModel with
                    | Some cellData ->
                        match cellData.Value with
                        | CellValue.Formula (formula, _) ->
                            let newCellData = {
                                cellData with
                                    Value = CellValue.Formula (formula, newValue)
                            }
                            updatedModel <- setCell addr newCellData updatedModel
                        | _ -> ()
                    | None -> ()

                    // Clear dirty flag
                    updatedEngine <- clearDirty addr updatedEngine
                | None -> ()
            | _ -> ()

        (updatedModel, updatedEngine)

/// Update cell value and mark dependents dirty
let updateCell (addr: CellAddress) (value: CellValue) (cellData: CellData) (model: SpreadsheetModel) (engine: CalcEngine): SpreadsheetModel * CalcEngine =
    // Update the cell
    let newCellData = { cellData with Value = value }
    let newModel = setCell addr newCellData model

    // If it's a formula, rebuild dependencies
    let newEngine =
        match value with
        | CellValue.Formula (formula, _) ->
            buildDependencies addr formula engine
        | _ ->
            // Clear dependencies but keep dependents (other cells may reference this)
            match Map.tryFind addr engine.Metadata with
            | Some metadata ->
                // Remove this cell from dependents of cells it referenced
                let mutable updatedMetadata = engine.Metadata
                for dep in metadata.Dependencies do
                    match Map.tryFind dep updatedMetadata with
                    | Some depMeta ->
                        let newDepMeta = {
                            depMeta with
                                Dependents = Set.remove addr depMeta.Dependents
                        }
                        updatedMetadata <- Map.add dep newDepMeta updatedMetadata
                    | None -> ()

                // Keep dependents, clear dependencies and formula
                let newMetadata = {
                    metadata with
                        Dependencies = Set.empty
                        ParsedFormula = None
                }
                { engine with Metadata = Map.add addr newMetadata updatedMetadata }
            | None -> engine

    // Mark dependents dirty
    let dirtyEngine = markDirty addr newEngine

    // Rebuild calc chain
    let finalEngine = rebuildCalcChain newModel dirtyEngine

    (newModel, finalEngine)

/// Statistics about calculation engine
type CalcStats = {
    TotalCells: int
    FormulaCells: int
    DirtyCells: int
    Dependencies: int
    CalcChainLength: int
    HasCircularRef: bool
}

let getStats (model: SpreadsheetModel) (engine: CalcEngine): CalcStats =
    let formulaCells =
        model.Cells
        |> Map.toList
        |> List.filter (fun (_, cell) ->
            match cell.Value with
            | CellValue.Formula _ -> true
            | _ -> false
        )
        |> List.length

    let totalDeps =
        engine.Metadata
        |> Map.toList
        |> List.sumBy (fun (_, meta) -> Set.count meta.Dependencies)

    {
        TotalCells = Map.count model.Cells
        FormulaCells = formulaCells
        DirtyCells = Set.count engine.DirtySet
        Dependencies = totalDeps
        CalcChainLength = List.length engine.CalcChain
        HasCircularRef = formulaCells > 0 && List.isEmpty engine.CalcChain
    }
