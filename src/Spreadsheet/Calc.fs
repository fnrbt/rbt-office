module RedBat.Spreadsheet.Calc

open System.Collections.Generic
open System.Runtime.CompilerServices
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

/// Calculation engine state.
type CalcEngine = {
    /// Metadata for each formula cell, across every sheet.
    Metadata: Map<SheetCell, CalcMetadata>
    /// Formula cells that have at least one large-range rectangle dependency.
    RectDependentCells: Set<SheetCell>
    /// Coarse bounds of all large-range rectangles by sheet. Bounds may be
    /// conservative after formula removal; that only loses the fast skip.
    RectDependencyBounds: Map<SheetId, struct (int * int * int * int)>
    /// Formula cells containing volatile functions, forced dirty every pass.
    VolatileCells: Set<SheetCell>
    /// Pre-computed calculation order (topological sort).
    CalcChain: SheetCell list
    /// Cells marked dirty (need recalculation).
    DirtySet: Set<SheetCell>
}

[<Struct>]
type private RectDependencyIndexEntry = {
    FormulaCell: SheetCell
    Rect: DepRect
}

type private RectDependencySheetIndex = {
    RowBuckets: Dictionary<int, ResizeArray<RectDependencyIndexEntry>>
    ColBuckets: Dictionary<int, ResizeArray<RectDependencyIndexEntry>>
    Overflow: ResizeArray<RectDependencyIndexEntry>
}

type private RectDependencyIndex = {
    Sheets: Dictionary<SheetId, RectDependencySheetIndex>
}

let private rectDependencyIndexBucketSize = 4096
let private rectDependencyIndexMaxBucketsPerRect = 512
let private rectDependencyIndexCache = ConditionalWeakTable<Map<SheetCell, CalcMetadata>, RectDependencyIndex>()

let emptyMetadata = {
    Dependencies = Set.empty
    DepRects = []
    Dependents = Set.empty
    IsDirty = false
    ParsedFormula = None
}

let emptyEngine = {
    Metadata = Map.empty
    RectDependentCells = Set.empty
    RectDependencyBounds = Map.empty
    VolatileCells = Set.empty
    CalcChain = []
    DirtySet = Set.empty
}

let private dependencyParseCache = Dictionary<string, Result<Expr, string>>()
let private dependencyParseCacheOrder = Queue<string>()
let private dependencyParseCacheMaxEntries = 8192
let private dependencyParseCacheMaxTotalFormulaChars = 5000000
let mutable private dependencyParseCacheTotalFormulaChars = 0

let private dependencyParseCacheAddCapped (amount: int) (total: int) =
    if amount <= 0 then total
    elif total > dependencyParseCacheMaxTotalFormulaChars
         || amount > dependencyParseCacheMaxTotalFormulaChars - total then
        dependencyParseCacheMaxTotalFormulaChars + 1
    else
        total + amount

let private dependencyParseCacheEntryChars (formula: string) (parsed: Result<Expr, string>) =
    0
    |> dependencyParseCacheAddCapped formula.Length
    |> dependencyParseCacheAddCapped (parsedFormulaResultRetainedTextCharCount dependencyParseCacheMaxTotalFormulaChars parsed)

let private addDependencyParseCacheEntry (formula: string) (parsed: Result<Expr, string>) =
    let entryChars = dependencyParseCacheEntryChars formula parsed
    if entryChars > dependencyParseCacheMaxTotalFormulaChars then
        let mutable previous = Unchecked.defaultof<Result<Expr, string>>
        if dependencyParseCache.TryGetValue(formula, &previous) then
            dependencyParseCache.Remove formula |> ignore
            dependencyParseCacheTotalFormulaChars <- max 0 (dependencyParseCacheTotalFormulaChars - dependencyParseCacheEntryChars formula previous)
        parsed
    else
        let mutable previous = Unchecked.defaultof<Result<Expr, string>>
        if dependencyParseCache.TryGetValue(formula, &previous) then
            dependencyParseCacheTotalFormulaChars <- max 0 (dependencyParseCacheTotalFormulaChars - dependencyParseCacheEntryChars formula previous)
        else
            dependencyParseCacheOrder.Enqueue formula
        dependencyParseCache.[formula] <- parsed
        dependencyParseCacheTotalFormulaChars <- dependencyParseCacheTotalFormulaChars + entryChars
        while (dependencyParseCache.Count > dependencyParseCacheMaxEntries || dependencyParseCacheTotalFormulaChars > dependencyParseCacheMaxTotalFormulaChars)
              && dependencyParseCacheOrder.Count > 0 do
            let oldest = dependencyParseCacheOrder.Dequeue()
            if oldest <> formula then
                let mutable oldestParsed = Unchecked.defaultof<Result<Expr, string>>
                if dependencyParseCache.TryGetValue(oldest, &oldestParsed) then
                    dependencyParseCache.Remove oldest |> ignore
                    dependencyParseCacheTotalFormulaChars <- max 0 (dependencyParseCacheTotalFormulaChars - dependencyParseCacheEntryChars oldest oldestParsed)
        if dependencyParseCache.Count > dependencyParseCacheMaxEntries || dependencyParseCacheTotalFormulaChars > dependencyParseCacheMaxTotalFormulaChars then
            let current = dependencyParseCache.[formula]
            dependencyParseCache.Clear()
            dependencyParseCacheOrder.Clear()
            dependencyParseCacheOrder.Enqueue formula
            dependencyParseCache.[formula] <- current
            dependencyParseCacheTotalFormulaChars <- entryChars
        parsed

let private parseFormulaForDependencies formula =
    let mutable parsed = Unchecked.defaultof<Result<Expr, string>>
    if dependencyParseCache.TryGetValue(formula, &parsed) then
        parsed
    else
        parsed <- parseFormula formula
        addDependencyParseCacheEntry formula parsed

type private DependencyReferenceCacheEntry = {
    Dependencies: Set<SheetCell>
    DepRects: DepRect list
    DependencyCount: int
    RectCount: int
    FormulaChars: int
}

let private dependencyReferenceCache = Dictionary<struct (SheetId * string), DependencyReferenceCacheEntry>()
let private dependencyReferenceCacheOrder = Queue<struct (SheetId * string)>()
let private dependencyReferenceCacheMaxEntries = 8192
let private dependencyReferenceCacheMaxTotalFormulaChars = 5000000
let private dependencyReferenceCacheMaxTotalDependencies = 20000
let private dependencyReferenceCacheMaxTotalRects = 8192
let mutable private dependencyReferenceCacheTotalFormulaChars = 0
let mutable private dependencyReferenceCacheTotalDependencies = 0
let mutable private dependencyReferenceCacheTotalRects = 0

let private dependencyReferenceCacheKeyFormulaLength (struct (sheet, formula): struct (SheetId * string)) =
    sheet.Length + formula.Length

let private dependencyReferenceRectCount (rects: DepRect list) =
    let mutable count = 0
    let mutable remaining = rects
    while count <= dependencyReferenceCacheMaxTotalRects && not remaining.IsEmpty do
        count <- count + 1
        remaining <- remaining.Tail
    count

let private dependencyReferenceCacheEntry key deps rects = {
    Dependencies = deps
    DepRects = rects
    DependencyCount = max 1 (Set.count deps)
    RectCount = max 1 (dependencyReferenceRectCount rects)
    FormulaChars = max 1 (dependencyReferenceCacheKeyFormulaLength key)
}

let private addDependencyReferenceCacheEntry key value =
    let subtractEntry (entry: DependencyReferenceCacheEntry) =
        dependencyReferenceCacheTotalFormulaChars <- dependencyReferenceCacheTotalFormulaChars - entry.FormulaChars
        dependencyReferenceCacheTotalDependencies <- dependencyReferenceCacheTotalDependencies - entry.DependencyCount
        dependencyReferenceCacheTotalRects <- dependencyReferenceCacheTotalRects - entry.RectCount
    if value.FormulaChars > dependencyReferenceCacheMaxTotalFormulaChars
       || value.DependencyCount > dependencyReferenceCacheMaxTotalDependencies
       || value.RectCount > dependencyReferenceCacheMaxTotalRects then
        let mutable oldEntry = Unchecked.defaultof<DependencyReferenceCacheEntry>
        if dependencyReferenceCache.TryGetValue(key, &oldEntry) then
            dependencyReferenceCache.Remove key |> ignore
            subtractEntry oldEntry
        value
    else
    if dependencyReferenceCache.ContainsKey key then
        subtractEntry dependencyReferenceCache.[key]
    else
        dependencyReferenceCacheOrder.Enqueue key
    dependencyReferenceCache.[key] <- value
    dependencyReferenceCacheTotalFormulaChars <- dependencyReferenceCacheTotalFormulaChars + value.FormulaChars
    dependencyReferenceCacheTotalDependencies <- dependencyReferenceCacheTotalDependencies + value.DependencyCount
    dependencyReferenceCacheTotalRects <- dependencyReferenceCacheTotalRects + value.RectCount
    while (dependencyReferenceCache.Count > dependencyReferenceCacheMaxEntries
           || dependencyReferenceCacheTotalFormulaChars > dependencyReferenceCacheMaxTotalFormulaChars
           || dependencyReferenceCacheTotalDependencies > dependencyReferenceCacheMaxTotalDependencies
           || dependencyReferenceCacheTotalRects > dependencyReferenceCacheMaxTotalRects)
          && dependencyReferenceCacheOrder.Count > 0 do
        let oldest = dependencyReferenceCacheOrder.Dequeue()
        let mutable oldEntry = Unchecked.defaultof<DependencyReferenceCacheEntry>
        if oldest <> key && dependencyReferenceCache.TryGetValue(oldest, &oldEntry) then
            dependencyReferenceCache.Remove oldest |> ignore
            subtractEntry oldEntry
    if dependencyReferenceCache.Count > dependencyReferenceCacheMaxEntries
       || dependencyReferenceCacheTotalFormulaChars > dependencyReferenceCacheMaxTotalFormulaChars
       || dependencyReferenceCacheTotalDependencies > dependencyReferenceCacheMaxTotalDependencies
       || dependencyReferenceCacheTotalRects > dependencyReferenceCacheMaxTotalRects then
        let current = dependencyReferenceCache.[key]
        dependencyReferenceCache.Clear()
        dependencyReferenceCacheOrder.Clear()
        dependencyReferenceCacheOrder.Enqueue key
        dependencyReferenceCache.[key] <- current
        dependencyReferenceCacheTotalFormulaChars <- current.FormulaChars
        dependencyReferenceCacheTotalDependencies <- current.DependencyCount
        dependencyReferenceCacheTotalRects <- current.RectCount
    value

let private exprUsesWorkbookResolution expr =
    let refUsesWorkbook (ref: RefAst) = ref.Sheet.IsSome
    let stack = ResizeArray<Expr>()
    stack.Add expr
    let mutable found = false
    while not found && stack.Count > 0 do
        let last = stack.Count - 1
        let expr = stack.[last]
        stack.RemoveAt last
        match expr with
        | Number _ | Str _ | Bool _ | ErrorLit _ | ArrayLit _ -> ()
        | NameRef _
        | TableRef _ -> found <- true
        | CellRef ref
        | SpillRef ref ->
            if refUsesWorkbook ref then
                found <- true
        | RangeRef (left, right) ->
            if refUsesWorkbook left || refUsesWorkbook right then
                found <- true
        | UnaryOp (_, inner) -> stack.Add inner
        | BinOp (left, _, right) ->
            stack.Add left
            stack.Add right
        | Function (_, args) ->
            for arg in args do
                stack.Add arg
    found

let private referencesRangedFor (formula: string) (cell: SheetCell) (workbook: Workbook) (expr: Expr) =
    if exprUsesWorkbookResolution expr then
        referencesRanged (Some workbook) cell.Sheet expr
    else
        let key = struct (cell.Sheet, formula)
        let mutable cached = Unchecked.defaultof<DependencyReferenceCacheEntry>
        if dependencyReferenceCache.TryGetValue(key, &cached) then
            cached.Dependencies, cached.DepRects
        else
            let deps, rects = referencesRanged (Some workbook) cell.Sheet expr
            let cached = addDependencyReferenceCacheEntry key (dependencyReferenceCacheEntry key deps rects)
            cached.Dependencies, cached.DepRects

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

let private depRectsContainCell (rects: DepRect list) (cell: SheetCell) =
    let mutable found = false
    let mutable remaining = rects
    while not found && not remaining.IsEmpty do
        match remaining with
        | rect :: tail ->
            if rectContainsCell rect cell then
                found <- true
            remaining <- tail
        | [] -> ()
    found

let private rectDependencyBucket value =
    if value >= 0 then
        value / rectDependencyIndexBucketSize
    else
        (value - rectDependencyIndexBucketSize + 1) / rectDependencyIndexBucketSize

let private rectDependencySheetIndex sheetId (sheets: Dictionary<SheetId, RectDependencySheetIndex>) =
    let mutable sheetIndex = Unchecked.defaultof<RectDependencySheetIndex>
    if sheets.TryGetValue(sheetId, &sheetIndex) then
        sheetIndex
    else
        let created = {
            RowBuckets = Dictionary<int, ResizeArray<RectDependencyIndexEntry>>()
            ColBuckets = Dictionary<int, ResizeArray<RectDependencyIndexEntry>>()
            Overflow = ResizeArray<RectDependencyIndexEntry>()
        }
        sheets.[sheetId] <- created
        created

let private addRectDependencyBucketEntry bucket entry (buckets: Dictionary<int, ResizeArray<RectDependencyIndexEntry>>) =
    let mutable entries = Unchecked.defaultof<ResizeArray<RectDependencyIndexEntry>>
    if not (buckets.TryGetValue(bucket, &entries)) then
        entries <- ResizeArray<RectDependencyIndexEntry>()
        buckets.[bucket] <- entries
    entries.Add entry

let private addRectDependencyBuckets startBucket endBucket entry buckets =
    for bucket in startBucket .. endBucket do
        addRectDependencyBucketEntry bucket entry buckets

let private buildRectDependencyIndex (engine: CalcEngine) =
    let sheets = Dictionary<SheetId, RectDependencySheetIndex>()
    for formulaCell in engine.RectDependentCells do
        match Map.tryFind formulaCell engine.Metadata with
        | Some metadata ->
            for rect in metadata.DepRects do
                let sheetIndex = rectDependencySheetIndex rect.Sheet sheets
                let entry = { FormulaCell = formulaCell; Rect = rect }
                let rowStartBucket = rectDependencyBucket rect.R0
                let rowEndBucket = rectDependencyBucket rect.R1
                let colStartBucket = rectDependencyBucket rect.C0
                let colEndBucket = rectDependencyBucket rect.C1
                let rowBucketCount = rowEndBucket - rowStartBucket + 1
                let colBucketCount = colEndBucket - colStartBucket + 1
                let rowSpan = rect.R1 - rect.R0
                let colSpan = rect.C1 - rect.C0
                if rowBucketCount <= rectDependencyIndexMaxBucketsPerRect
                   && (rowBucketCount < colBucketCount
                       || (rowBucketCount = colBucketCount && rowSpan >= colSpan)) then
                    addRectDependencyBuckets rowStartBucket rowEndBucket entry sheetIndex.RowBuckets
                elif colBucketCount <= rectDependencyIndexMaxBucketsPerRect then
                    addRectDependencyBuckets colStartBucket colEndBucket entry sheetIndex.ColBuckets
                elif rowBucketCount <= rectDependencyIndexMaxBucketsPerRect then
                    addRectDependencyBuckets rowStartBucket rowEndBucket entry sheetIndex.RowBuckets
                else
                    sheetIndex.Overflow.Add entry
        | None -> ()
    { Sheets = sheets }

let private rememberRectDependencyIndex (engine: CalcEngine) =
    rectDependencyIndexCache.Remove engine.Metadata |> ignore
    if not (Set.isEmpty engine.RectDependentCells) then
        rectDependencyIndexCache.Add(engine.Metadata, buildRectDependencyIndex engine)
    engine

let private rectDependencyIndexFor (engine: CalcEngine) =
    let mutable index = Unchecked.defaultof<RectDependencyIndex>
    if rectDependencyIndexCache.TryGetValue(engine.Metadata, &index) then
        index
    else
        let index = buildRectDependencyIndex engine
        rectDependencyIndexCache.Remove engine.Metadata |> ignore
        rectDependencyIndexCache.Add(engine.Metadata, index)
        index

let private expandRectDependencyBounds (rects: DepRect list) (bounds: Map<SheetId, struct (int * int * int * int)>) =
    let mutable updated = bounds
    for rect in rects do
        let nextBounds =
            match Map.tryFind rect.Sheet updated with
            | Some (struct (minRow, maxRow, minCol, maxCol)) ->
                struct (min minRow rect.R0, max maxRow rect.R1, min minCol rect.C0, max maxCol rect.C1)
            | None ->
                struct (rect.R0, rect.R1, rect.C0, rect.C1)
        updated <- Map.add rect.Sheet nextBounds updated
    updated

let private rectDependencyBoundsMayContain (cell: SheetCell) (engine: CalcEngine) =
    match Map.tryFind cell.Sheet engine.RectDependencyBounds with
    | Some (struct (minRow, maxRow, minCol, maxCol)) ->
        cell.Cell.Row >= minRow && cell.Cell.Row <= maxRow && cell.Cell.Col >= minCol && cell.Cell.Col <= maxCol
    | None -> false

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
            RectDependentCells = Set.remove cell engine.RectDependentCells
            VolatileCells = Set.remove cell engine.VolatileCells
            DirtySet = Set.remove cell engine.DirtySet }
    | None -> engine

/// Build (or rebuild) the dependency edges for one formula cell.
let buildDependencies (cell: SheetCell) (formula: string)
                      (workbook: Workbook) (engine: CalcEngine): CalcEngine =
    match parseFormulaForDependencies formula with
    | Ok expr ->
        // Large ranges become rectangle deps (rects); small ranges + single
        // cells stay explicit edges (deps).
        let deps, rects = referencesRangedFor formula cell workbook expr
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

        let rectDependentCells =
            if List.isEmpty rects then Set.remove cell engine.RectDependentCells
            else Set.add cell engine.RectDependentCells
        let volatileCells =
            if containsVolatile expr then Set.add cell engine.VolatileCells
            else Set.remove cell engine.VolatileCells
        { engine with
            Metadata = updatedMetadata
            RectDependentCells = rectDependentCells
            RectDependencyBounds = expandRectDependencyBounds rects engine.RectDependencyBounds
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
    | Some chain -> { engine with CalcChain = chain } |> rememberRectDependencyIndex
    | None ->
        // Circular reference — clear the chain to signal the error state.
        { engine with CalcChain = [] } |> rememberRectDependencyIndex

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
    if not (Set.isEmpty engine.RectDependentCells) && rectDependencyBoundsMayContain cell engine then
        let index = rectDependencyIndexFor engine
        let mutable sheetIndex = Unchecked.defaultof<RectDependencySheetIndex>
        if index.Sheets.TryGetValue(cell.Sheet, &sheetIndex) then
            let addMatches (entries: ResizeArray<RectDependencyIndexEntry>) =
                for entry in entries do
                    if rectContainsCell entry.Rect cell then
                        dirtySeeds.Add entry.FormulaCell
            let rowBucket = rectDependencyBucket cell.Cell.Row
            let colBucket = rectDependencyBucket cell.Cell.Col
            let mutable rowEntries = Unchecked.defaultof<ResizeArray<RectDependencyIndexEntry>>
            if sheetIndex.RowBuckets.TryGetValue(rowBucket, &rowEntries) then
                addMatches rowEntries
            let mutable colEntries = Unchecked.defaultof<ResizeArray<RectDependencyIndexEntry>>
            if sheetIndex.ColBuckets.TryGetValue(colBucket, &colEntries) then
                addMatches colEntries
            if sheetIndex.Overflow.Count > 0 then
                addMatches sheetIndex.Overflow
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
                  CriteriaCache = Some (CriteriaPredicateCache())
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
                    | Some expr -> not (exprUsesWorkbookResolution expr)
                    | None -> false
                | None -> false
            | _ -> true
        | _ -> false
    | None -> false

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
        let mutable wb = workbook
        let mutable eng = engine
        for cell in recalculateVisitOrder engine do
            match Map.tryFind cell eng.Metadata with
            | Some metadata when Set.contains cell eng.DirtySet ->
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
                | None -> ()
            | _ -> ()
        // Anything still dirty after the chain pass is a non-formula source
        // cell (the *edited* cell itself, which has nothing to recompute) or
        // — when the graph has a cycle — an unresolvable formula. Either way
        // clear it: leaving a stale dirty flag would make the next edit's
        // markDirty short-circuit and skip propagating to dependents.
        (wb, { eng with DirtySet = Set.empty })

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
