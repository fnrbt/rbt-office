#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Workbook
open RedBat.Spreadsheet.Calc

// Exercises the incremental calc engine's cross-sheet dependency graph:
// an edit on Sheet1 must dirty — and recalculate — a formula on Sheet2
// that references it.

let cell v = { Value = v; Format = defaultFormat; Comment = None }

// Sheet1: A1 = 10, A2 = "=A1*2"
let sheet1Model =
    { createEmpty 100 26 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
                  { Row = 1; Col = 0 }, cell (CellValue.Formula ("=A1*2", CellValue.Empty)) ] }

// Sheet2: B1 = "=Sheet1!A1+100"  (cross-sheet reference)
let sheet2Model =
    { createEmpty 100 26 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 1 }, cell (CellValue.Formula ("=Sheet1!A1+100", CellValue.Empty)) ] }

let wb0 =
    { Sheets =
        Map.ofList
            [ "sheet1", { Id = "sheet1"; Name = "Sheet1"; Model = sheet1Model; Visible = true }
              "sheet2", { Id = "sheet2"; Name = "Sheet2"; Model = sheet2Model; Visible = true } ]
      SheetOrder = [ "sheet1"; "sheet2" ]
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names = Map.empty }

let a1 = { Sheet = "sheet1"; Cell = { Row = 0; Col = 0 } }
let a2 = { Sheet = "sheet1"; Cell = { Row = 1; Col = 0 } }
let b1 = { Sheet = "sheet2"; Cell = { Row = 0; Col = 1 } }

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-28s  %s" label detail)

/// Read a cell's (cached) value out of a workbook.
let readWb (sc: SheetCell) (wb: Workbook) : CellValue =
    match Map.tryFind sc.Sheet wb.Sheets with
    | Some s ->
        match getCell sc.Cell s.Model with
        | Some cd -> (match cd.Value with CellValue.Formula (_, v) -> v | v -> v)
        | None -> CellValue.Empty
    | None -> CellValue.Empty

// --- Build the dependency graph -----------------------------------------
let engine0 =
    emptyEngine
    |> buildDependencies a2 "=A1*2" wb0
    |> buildDependencies b1 "=Sheet1!A1+100" wb0
    |> rebuildCalcChain wb0

// The cross-sheet edge exists: Sheet2!B1 depends on Sheet1!A1.
ok "xsheet-dep-edge"
   (engine0.Metadata.[b1].Dependencies |> Set.contains a1)
   (sprintf "B1 deps = %A" engine0.Metadata.[b1].Dependencies)
// Sheet1!A1 has both a same-sheet and a cross-sheet dependent.
ok "xsheet-dependents"
   (engine0.Metadata.[a1].Dependents = Set.ofList [ a2; b1 ])
   (sprintf "A1 dependents = %A" engine0.Metadata.[a1].Dependents)
// A2 depends only on A1 (same sheet).
ok "samesheet-dep-edge"
   (engine0.Metadata.[a2].Dependencies = Set.ofList [ a1 ])
   (sprintf "A2 deps = %A" engine0.Metadata.[a2].Dependencies)
// No cycle → the chain has both formula cells, and the topological sort
// places A1 before everything that depends on it (the chain also includes
// A1 itself as a visited dependency node — harmless, recalculateCell skips
// non-formula cells).
let chainIndex sc = List.findIndex ((=) sc) engine0.CalcChain
ok "chain-built"
   (List.contains a2 engine0.CalcChain
    && List.contains b1 engine0.CalcChain
    && chainIndex a1 < chainIndex a2
    && chainIndex a1 < chainIndex b1)
   (sprintf "chain = %A" engine0.CalcChain)

let sameFormulaSheet1 = { Sheet = "sheet1"; Cell = { Row = 4; Col = 0 } }
let sameFormulaSheet2 = { Sheet = "sheet2"; Cell = { Row = 4; Col = 0 } }
let sameFormulaDep1 = { Sheet = "sheet1"; Cell = { Row = 0; Col = 0 } }
let sameFormulaDep2 = { Sheet = "sheet2"; Cell = { Row = 0; Col = 0 } }
let sameFormulaEngine =
    emptyEngine
    |> buildDependencies sameFormulaSheet1 "=A1+1" wb0
    |> buildDependencies sameFormulaSheet2 "=A1+1" wb0
ok "dependency-cache-home-sheet"
   (engine0.Metadata.[a1].Dependents = Set.ofList [ a2; b1 ]
    && sameFormulaEngine.Metadata.[sameFormulaSheet1].Dependencies = Set.ofList [ sameFormulaDep1 ]
    && sameFormulaEngine.Metadata.[sameFormulaSheet2].Dependencies = Set.ofList [ sameFormulaDep2 ])
   (sprintf "sheet1 deps=%A sheet2 deps=%A"
        sameFormulaEngine.Metadata.[sameFormulaSheet1].Dependencies
        sameFormulaEngine.Metadata.[sameFormulaSheet2].Dependencies)

let manyDependencyCount = 4200
let manyDependencyModel =
    { createEmpty (manyDependencyCount + 10) 4 with
        Cells =
            seq {
                for r in 0 .. manyDependencyCount - 1 do
                    yield { Row = r; Col = 0 }, cell (CellValue.Number (float r))
                    yield { Row = r; Col = 1 }, cell (CellValue.Formula (sprintf "=A%d+%d" (r + 1) (r % 11), CellValue.Empty))
            }
            |> Map.ofSeq }
let manyDependencyWb =
    { Sheets = Map.ofList [ "many", { Id = "many"; Name = "Many"; Model = manyDependencyModel; Visible = true } ]
      SheetOrder = [ "many" ]
      ActiveSheet = "many"
      Tables = Map.empty
      Names = Map.empty }
let buildManyDependencyEngine () =
    let mutable engine = emptyEngine
    for r in 0 .. manyDependencyCount - 1 do
        let sc = { Sheet = "many"; Cell = { Row = r; Col = 1 } }
        engine <- buildDependencies sc (sprintf "=A%d+%d" (r + 1) (r % 11)) manyDependencyWb engine
    engine
let manyDependencyEngineA = buildManyDependencyEngine ()
let manyDependencyEngineB = buildManyDependencyEngine ()
let manyDependencyLast = { Sheet = "many"; Cell = { Row = manyDependencyCount - 1; Col = 1 } }
let manyDependencyLastDep = { Sheet = "many"; Cell = { Row = manyDependencyCount - 1; Col = 0 } }
ok "dependency-cache-many-distinct"
   (manyDependencyEngineA.Metadata.Count = manyDependencyEngineB.Metadata.Count
    && manyDependencyEngineB.Metadata.[manyDependencyLast].Dependencies = Set.ofList [ manyDependencyLastDep ])
   (sprintf "metadata=%d lastDeps=%A"
        manyDependencyEngineB.Metadata.Count
        manyDependencyEngineB.Metadata.[manyDependencyLast].Dependencies)
let largeDependencyFormula index =
    "=" + "\"" + System.String('x', 10000) + sprintf "%04d\"&A1" index
let largeDependencyCacheRetainedBytes =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let mutable engine = emptyEngine
    let formulaCell = { Sheet = "many"; Cell = { Row = 0; Col = 1 } }
    for index in 0 .. 1999 do
        engine <- buildDependencies formulaCell (largeDependencyFormula index) manyDependencyWb engine
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.GetTotalMemory(true) - before
ok "dependency-cache-bounded-source-memory"
   (largeDependencyCacheRetainedBytes < 40000000L)
   (sprintf "large dependency caches retained %.1f MB" (float largeDependencyCacheRetainedBytes / 1024.0 / 1024.0))
let dependencyParseDerivedOversizedFormulaWeakReference () =
    let formula = "=\"" + System.String('D', 3000000) + "\"&Sheet1!A1"
    let formulaCell = { Sheet = "sheet2"; Cell = { Row = 7; Col = 1 } }
    buildDependencies formulaCell formula wb0 emptyEngine |> ignore
    System.WeakReference(formula)
let dependencyParseDerivedOversizedFormulaWeak = dependencyParseDerivedOversizedFormulaWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "dependency-parse-cache-skips-derived-entry"
   (not dependencyParseDerivedOversizedFormulaWeak.IsAlive)
   "dependency parse cache should count retained AST text, not only formula source"
let longExplicitDependencyRefCount = 12000
let longExplicitDependencyFormula index =
    "="
    + String.concat
        "+"
        [ for row in 1 .. longExplicitDependencyRefCount ->
            sprintf "A%d" (row + index) ]
let longExplicitDependencyRetainedBytes, longExplicitDependencyFinalDepCount =
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let before = System.GC.GetTotalMemory(true)
    let formulaCell = { Sheet = "many"; Cell = { Row = 0; Col = 2 } }
    let mutable engine = emptyEngine
    for index in 0 .. 15 do
        engine <- buildDependencies formulaCell (longExplicitDependencyFormula index) manyDependencyWb engine
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    let depCount =
        match Map.tryFind formulaCell engine.Metadata with
        | Some metadata -> Set.count metadata.Dependencies
        | None -> 0
    System.GC.GetTotalMemory(true) - before, depCount
ok "dependency extractor long explicit formula"
   (longExplicitDependencyFinalDepCount = longExplicitDependencyRefCount)
   (sprintf "dependency count %d" longExplicitDependencyFinalDepCount)
ok "dependency cache bounded explicit refs"
   (longExplicitDependencyRetainedBytes < 35000000L)
   (sprintf "explicit dependency caches retained %.1f MB" (float longExplicitDependencyRetainedBytes / 1024.0 / 1024.0))
let dependencyCacheOversizedSheetIdWeakReference () =
    let sheetId = System.String('D', 6 * 1024 * 1024)
    let model =
        { createEmpty 4 2 with
            Cells =
                [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
                  { Row = 0; Col = 1 }, cell (CellValue.Formula ("=A1+1", CellValue.Empty)) ]
                |> Map.ofList }
    let workbook =
        { Sheets =
            Map.ofList [
                sheetId,
                { Id = sheetId
                  Name = "OversizedSheetId"
                  Model = model
                  Visible = true }
            ]
          SheetOrder = [ sheetId ]
          ActiveSheet = sheetId
          Tables = Map.empty
          Names = Map.empty }
    let formulaCell = { Sheet = sheetId; Cell = { Row = 0; Col = 1 } }
    buildDependencies formulaCell "=A1+1" workbook emptyEngine |> ignore
    System.WeakReference(sheetId)
let dependencyCacheOversizedSheetIdWeak = dependencyCacheOversizedSheetIdWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "dependency cache skips oversized sheet id"
   (not dependencyCacheOversizedSheetIdWeak.IsAlive)
   "dependency reference cache should not retain oversized sheet ids in keys"
let volatileFormulaCell = { Sheet = "sheet1"; Cell = { Row = 9; Col = 2 } }
let volatileFormulaEngine =
    emptyEngine
    |> buildDependencies volatileFormulaCell "=RAND()+1" wb0
let nonvolatileFormulaEngine =
    buildDependencies volatileFormulaCell "=A1+1" wb0 volatileFormulaEngine
ok "volatile-cell-index-updates"
   (Set.contains volatileFormulaCell volatileFormulaEngine.VolatileCells
    && not (Set.contains volatileFormulaCell nonvolatileFormulaEngine.VolatileCells))
   (sprintf "volatile=%A nonvolatile=%A"
        volatileFormulaEngine.VolatileCells
        nonvolatileFormulaEngine.VolatileCells)

// --- Edit Sheet1!A1, then recalc ----------------------------------------
let a1Data =
    match getCell a1.Cell sheet1Model with
    | Some cd -> cd
    | None -> failwith "A1 missing"

let (wb1, engine1) = updateCell a1 (CellValue.Number 20.0) a1Data wb0 engine0

// The edit dirtied A1 and propagated across the sheet boundary to B1.
ok "edit-dirties-xsheet"
   (Set.contains a2 engine1.DirtySet && Set.contains b1 engine1.DirtySet)
   (sprintf "dirty = %A" engine1.DirtySet)

let (wb2, engine2) = smartRecalculate wb1 engine1

// Same-sheet dependent recalculated: A2 = A1*2 = 40.
ok "recalc-samesheet" (readWb a2 wb2 = CellValue.Number 40.0)
   (sprintf "A2 = %A" (readWb a2 wb2))
// Cross-sheet dependent recalculated: B1 = Sheet1!A1 + 100 = 120.
ok "recalc-xsheet" (readWb b1 wb2 = CellValue.Number 120.0)
   (sprintf "B1 = %A" (readWb b1 wb2))
// The formula cells were cleared from the dirty set.
ok "dirty-cleared-after-recalc"
   (not (Set.contains a2 engine2.DirtySet) && not (Set.contains b1 engine2.DirtySet))
   (sprintf "dirty = %A" engine2.DirtySet)
let (sameA1Wb, sameA1Engine) = updateCell a1 (CellValue.Number 20.0) a1Data wb2 engine2
ok "same-data-update-identity"
   (obj.ReferenceEquals(sameA1Wb, wb2) && obj.ReferenceEquals(sameA1Engine, engine2))
   (sprintf "dirty = %A" sameA1Engine.DirtySet)

// --- A second edit only recomputes what changed -------------------------
let a1Data2 =
    match getCell a1.Cell (Map.find "sheet1" wb2.Sheets).Model with
    | Some cd -> cd
    | None -> failwith "A1 missing (2)"
let (wb3, engine3) = updateCell a1 (CellValue.Number 1.0) a1Data2 wb2 engine2
let (wb4, _) = smartRecalculate wb3 engine3
ok "recalc-again-samesheet" (readWb a2 wb4 = CellValue.Number 2.0)
   (sprintf "A2 = %A" (readWb a2 wb4))
ok "recalc-again-xsheet" (readWb b1 wb4 = CellValue.Number 101.0)
   (sprintf "B1 = %A" (readWb b1 wb4))

// --- Cross-sheet cycle is detected --------------------------------------
let cycModel1 =
    { createEmpty 100 26 with
        Cells = Map.ofList [ { Row = 5; Col = 5 }, cell (CellValue.Formula ("=Sheet2!F6", CellValue.Empty)) ] }
let cycModel2 =
    { createEmpty 100 26 with
        Cells = Map.ofList [ { Row = 5; Col = 5 }, cell (CellValue.Formula ("=Sheet1!F6", CellValue.Empty)) ] }
let cycWb =
    { wb0 with
        Sheets =
            Map.ofList
                [ "sheet1", { Id = "sheet1"; Name = "Sheet1"; Model = cycModel1; Visible = true }
                  "sheet2", { Id = "sheet2"; Name = "Sheet2"; Model = cycModel2; Visible = true } ] }
let c1 = { Sheet = "sheet1"; Cell = { Row = 5; Col = 5 } }
let c2 = { Sheet = "sheet2"; Cell = { Row = 5; Col = 5 } }
let cycEngine =
    emptyEngine
    |> buildDependencies c1 "=Sheet2!F6" cycWb
    |> buildDependencies c2 "=Sheet1!F6" cycWb
    |> rebuildCalcChain cycWb
ok "xsheet-cycle-detected" (List.isEmpty cycEngine.CalcChain)
   (sprintf "chain = %A" cycEngine.CalcChain)
ok "xsheet-cycle-in-stats" ((getStats cycWb cycEngine).HasCircularRef)
   "expected HasCircularRef = true"

// --- Deep dependency chains sort without recursive stack growth -------------
let longCalcChainCount = 20000
let longCalcChainCells =
    seq {
        for row in 0 .. longCalcChainCount - 1 do
            let formula =
                if row = longCalcChainCount - 1 then
                    "=1"
                else
                    sprintf "=A%d+1" (row + 2)
            yield { Row = row; Col = 0 }, cell (CellValue.Formula (formula, CellValue.Empty))
    }
    |> Map.ofSeq
let longCalcChainModel = { createEmpty (longCalcChainCount + 1) 2 with Cells = longCalcChainCells }
let longCalcChainWb =
    { Sheets = Map.ofList [ "chain", { Id = "chain"; Name = "Chain"; Model = longCalcChainModel; Visible = true } ]
      SheetOrder = [ "chain" ]; ActiveSheet = "chain"; Tables = Map.empty; Names = Map.empty }
let mutable longCalcChainEngine = emptyEngine
for row in 0 .. longCalcChainCount - 1 do
    let formula =
        if row = longCalcChainCount - 1 then
            "=1"
        else
            sprintf "=A%d+1" (row + 2)
    longCalcChainEngine <-
        buildDependencies { Sheet = "chain"; Cell = { Row = row; Col = 0 } } formula longCalcChainWb longCalcChainEngine
let longCalcChainSw = System.Diagnostics.Stopwatch.StartNew()
let longCalcChainSorted = rebuildCalcChain longCalcChainWb longCalcChainEngine
longCalcChainSw.Stop()
let longCalcChainFirst =
    match longCalcChainSorted.CalcChain with
    | first :: _ -> Some first.Cell.Row
    | [] -> None
let rec lastChainRow rows =
    match rows with
    | [] -> None
    | [ last ] -> Some last.Cell.Row
    | _ :: tail -> lastChainRow tail
let longCalcChainLast =
    lastChainRow longCalcChainSorted.CalcChain
ok "calc-chain-deep-sort"
   (longCalcChainSorted.CalcChain.Length = longCalcChainCount
    && longCalcChainFirst = Some (longCalcChainCount - 1)
    && longCalcChainLast = Some 0)
   (sprintf "len=%d first=%A last=%A"
        longCalcChainSorted.CalcChain.Length
        longCalcChainFirst
        longCalcChainLast)
ok "calc-chain-deep-sort-time"
   (longCalcChainSw.Elapsed.TotalMilliseconds < 3000.0)
   (sprintf "elapsed %.1f ms" longCalcChainSw.Elapsed.TotalMilliseconds)
let longCalcChainTail = { Sheet = "chain"; Cell = { Row = longCalcChainCount - 1; Col = 0 } }
let longCalcChainTop = { Sheet = "chain"; Cell = { Row = 0; Col = 0 } }
let longCalcChainDirty = markDirty longCalcChainTail longCalcChainSorted
let longCalcChainRecalcSw = System.Diagnostics.Stopwatch.StartNew()
let longCalcChainRecalcWb, longCalcChainRecalcEngine = smartRecalculate longCalcChainWb longCalcChainDirty
longCalcChainRecalcSw.Stop()
ok "calc-chain-deep-recalc"
   (readWb longCalcChainTop longCalcChainRecalcWb = CellValue.Number (float longCalcChainCount)
    && Set.isEmpty longCalcChainRecalcEngine.DirtySet)
   (sprintf "top=%A dirty=%d"
        (readWb longCalcChainTop longCalcChainRecalcWb)
        (Set.count longCalcChainRecalcEngine.DirtySet))
ok "calc-chain-deep-recalc-time"
   (longCalcChainRecalcSw.Elapsed.TotalMilliseconds < 10000.0)
   (sprintf "elapsed %.1f ms" longCalcChainRecalcSw.Elapsed.TotalMilliseconds)
let longCalcChainCleanSw = System.Diagnostics.Stopwatch.StartNew()
let mutable longCalcChainCleanChecksum = 0
for _ in 1 .. 20 do
    let cleanWorkbook, cleanEngine = smartRecalculate longCalcChainRecalcWb longCalcChainRecalcEngine
    if obj.ReferenceEquals(cleanWorkbook, longCalcChainRecalcWb)
       && obj.ReferenceEquals(cleanEngine, longCalcChainRecalcEngine) then
        longCalcChainCleanChecksum <- longCalcChainCleanChecksum + 1
longCalcChainCleanSw.Stop()
ok "calc-chain-clean-recalc-time"
   (longCalcChainCleanChecksum = 20
    && longCalcChainCleanSw.Elapsed.TotalMilliseconds < 10.0)
   (sprintf "20 clean recalcs took %.1f ms checksum=%d"
        longCalcChainCleanSw.Elapsed.TotalMilliseconds
        longCalcChainCleanChecksum)
let sparseDirtyFormulaCount = 20000
let sparseDirtyCells =
    seq {
        for row in 0 .. sparseDirtyFormulaCount - 1 do
            yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
            yield { Row = row; Col = 1 }, cell (CellValue.Formula (sprintf "=A%d+1" (row + 1), CellValue.Number (float (row + 1))))
    }
    |> Map.ofSeq
let sparseDirtyModel = { createEmpty (sparseDirtyFormulaCount + 1) 4 with Cells = sparseDirtyCells }
let sparseDirtyWb =
    { Sheets = Map.ofList [ "sparse", { Id = "sparse"; Name = "Sparse"; Model = sparseDirtyModel; Visible = true } ]
      SheetOrder = [ "sparse" ]; ActiveSheet = "sparse"; Tables = Map.empty; Names = Map.empty }
let mutable sparseDirtyEngine = emptyEngine
for row in 0 .. sparseDirtyFormulaCount - 1 do
    sparseDirtyEngine <-
        buildDependencies
            { Sheet = "sparse"; Cell = { Row = row; Col = 1 } }
            (sprintf "=A%d+1" (row + 1))
            sparseDirtyWb
            sparseDirtyEngine
let sparseDirtySortedEngine = rebuildCalcChain sparseDirtyWb sparseDirtyEngine
let sparseDirtySource = { Sheet = "sparse"; Cell = { Row = sparseDirtyFormulaCount / 2; Col = 0 } }
let sparseDirtyUpdatedWb, sparseDirtyUpdatedEngine =
    updateCell
        sparseDirtySource
        (CellValue.Number 123456.0)
        (cell (CellValue.Number 123456.0))
        sparseDirtyWb
        sparseDirtySortedEngine
let sparseDirtySw = System.Diagnostics.Stopwatch.StartNew()
let sparseDirtyRecalcWb, sparseDirtyRecalcEngine = smartRecalculate sparseDirtyUpdatedWb sparseDirtyUpdatedEngine
sparseDirtySw.Stop()
let sparseDirtyValue =
    readWb { Sheet = "sparse"; Cell = { Row = sparseDirtyFormulaCount / 2; Col = 1 } } sparseDirtyRecalcWb
ok "calc-chain-sparse-dirty-recalc-time"
   (sparseDirtyValue = CellValue.Number 123457.0
    && Set.isEmpty sparseDirtyRecalcEngine.DirtySet
    && sparseDirtySw.Elapsed.TotalMilliseconds < 60.0)
   (sprintf "single dirty recalc in %d independent formulas took %.1f ms value=%A dirty=%d"
        sparseDirtyFormulaCount
        sparseDirtySw.Elapsed.TotalMilliseconds
        sparseDirtyValue
        (Set.count sparseDirtyRecalcEngine.DirtySet))

let criteriaRecalcFormulaCount = 4000
let criteriaRecalcSourceRows = 32
let criteriaRecalcFormula = "=COUNTIF($A$1:$A$32,\">=16\")"
let criteriaRecalcCells =
    seq {
        for row in 0 .. criteriaRecalcSourceRows - 1 do
            yield { Row = row; Col = 0 }, cell (CellValue.Number (float row))
        for row in 0 .. criteriaRecalcFormulaCount - 1 do
            yield { Row = row; Col = 1 }, cell (CellValue.Formula (criteriaRecalcFormula, CellValue.Empty))
    }
    |> Map.ofSeq
let criteriaRecalcModel =
    { createEmpty (criteriaRecalcFormulaCount + 10) 4 with Cells = criteriaRecalcCells }
let criteriaRecalcWb =
    { Sheets = Map.ofList [ "criteria", { Id = "criteria"; Name = "Criteria"; Model = criteriaRecalcModel; Visible = true } ]
      SheetOrder = [ "criteria" ]; ActiveSheet = "criteria"; Tables = Map.empty; Names = Map.empty }
let mutable criteriaRecalcEngine = emptyEngine
for row in 0 .. criteriaRecalcFormulaCount - 1 do
    criteriaRecalcEngine <-
        buildDependencies
            { Sheet = "criteria"; Cell = { Row = row; Col = 1 } }
            criteriaRecalcFormula
            criteriaRecalcWb
            criteriaRecalcEngine
criteriaRecalcEngine <- rebuildCalcChain criteriaRecalcWb criteriaRecalcEngine
let criteriaRecalcInitialWb, criteriaRecalcInitialEngine =
    smartRecalculate criteriaRecalcWb criteriaRecalcEngine
let criteriaRecalcDirtyWb, criteriaRecalcDirtyEngine =
    updateCell
        { Sheet = "criteria"; Cell = { Row = 0; Col = 0 } }
        (CellValue.Number 99.0)
        (cell (CellValue.Number 99.0))
        criteriaRecalcInitialWb
        criteriaRecalcInitialEngine
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let criteriaRecalcAllocatedBefore = System.GC.GetAllocatedBytesForCurrentThread()
let criteriaRecalcSw = System.Diagnostics.Stopwatch.StartNew()
let criteriaRecalcResultWb, criteriaRecalcResultEngine =
    smartRecalculate criteriaRecalcDirtyWb criteriaRecalcDirtyEngine
criteriaRecalcSw.Stop()
let criteriaRecalcAllocated =
    System.GC.GetAllocatedBytesForCurrentThread() - criteriaRecalcAllocatedBefore
let criteriaRecalcValue =
    readWb { Sheet = "criteria"; Cell = { Row = 0; Col = 1 } } criteriaRecalcResultWb
ok "criteria-cache-shared-incremental-recalc"
   (criteriaRecalcValue = CellValue.Number 17.0
    && Set.isEmpty criteriaRecalcResultEngine.DirtySet
    && criteriaRecalcAllocated < 33L * 1024L * 1024L
    && criteriaRecalcSw.Elapsed.TotalMilliseconds < 500.0)
   (sprintf "criteria recalc took %.1f ms allocated=%d value=%A dirty=%d"
        criteriaRecalcSw.Elapsed.TotalMilliseconds
        criteriaRecalcAllocated
        criteriaRecalcValue
        (Set.count criteriaRecalcResultEngine.DirtySet))

// --- Range-granular dependency: a large range is one rectangle, not N edges -
// A1:A5000 (> rangeDepThreshold) feeding B1=SUM(...) must NOT reify 5000 edges,
// yet an edit to any data cell inside it must still dirty + recalc the SUM.
let bigModel =
    { createEmpty 6000 4 with
        Cells = seq { for r in 0 .. 4999 -> { Row = r; Col = 0 }, cell (CellValue.Number 1.0) } |> Map.ofSeq }
let bigWb0 =
    { Sheets = Map.ofList [ "s1", { Id = "s1"; Name = "Sheet1"; Model = bigModel; Visible = true } ]
      SheetOrder = [ "s1" ]; ActiveSheet = "s1"; Tables = Map.empty; Names = Map.empty }
let bb1 = { Sheet = "s1"; Cell = { Row = 0; Col = 1 } }
let bb1Data = cell (CellValue.Formula ("=SUM(A1:A5000)", CellValue.Empty))
let (bigWb1, bigEng1) = updateCell bb1 (CellValue.Formula ("=SUM(A1:A5000)", CellValue.Empty)) bb1Data bigWb0 emptyEngine
ok "bigrange-no-cell-edges"
   (Set.isEmpty bigEng1.Metadata.[bb1].Dependencies && List.length bigEng1.Metadata.[bb1].DepRects = 1)
   (sprintf "deps=%d rects=%d" (Set.count bigEng1.Metadata.[bb1].Dependencies) (List.length bigEng1.Metadata.[bb1].DepRects))
let (bigWb2, bigEng2) = smartRecalculate bigWb1 bigEng1
ok "bigrange-initial-sum" (readWb bb1 bigWb2 = CellValue.Number 5000.0) (sprintf "B1=%A" (readWb bb1 bigWb2))
let bb1Current =
    match getCell bb1.Cell (Map.find "s1" bigWb2.Sheets).Model with
    | Some cd -> cd
    | None -> failwith "B1 missing"
let (sameBigFormulaWb, sameBigFormulaEng) = updateCell bb1 bb1Current.Value bb1Current bigWb2 bigEng2
ok "same-formula-update-identity"
   (obj.ReferenceEquals(sameBigFormulaWb, bigWb2) && obj.ReferenceEquals(sameBigFormulaEng, bigEng2))
   (sprintf "dirty=%A" sameBigFormulaEng.DirtySet)
let ai = { Sheet = "s1"; Cell = { Row = 2499; Col = 0 } }   // a data cell inside the rectangle
let (bigWb3, bigEng3) = updateCell ai (CellValue.Number 11.0) (cell (CellValue.Number 11.0)) bigWb2 bigEng2
ok "bigrange-edit-dirties-via-rect" (Set.contains bb1 bigEng3.DirtySet) (sprintf "dirty=%A" (Set.count bigEng3.DirtySet))
let (bigWb4, _) = smartRecalculate bigWb3 bigEng3
ok "bigrange-recalc" (readWb bb1 bigWb4 = CellValue.Number 5010.0) (sprintf "B1=%A" (readWb bb1 bigWb4)) // 4999×1 + 11
let outsideBig = { Sheet = "s1"; Cell = { Row = 5500; Col = 3 } }
let (_, bigOutsideEng) = updateCell outsideBig (CellValue.Number 1.0) (cell (CellValue.Number 1.0)) bigWb2 bigEng2
ok "bigrange-outside-edit-clean" (not (Set.contains bb1 bigOutsideEng.DirtySet)) (sprintf "dirty=%A" (Set.count bigOutsideEng.DirtySet))
let wideBigModel =
    { createEmpty 4 3000 with
        Cells = seq { for c in 0 .. 2999 -> { Row = 0; Col = c }, cell (CellValue.Number 1.0) } |> Map.ofSeq }
let wideBigWb0 =
    { Sheets = Map.ofList [ "s1", { Id = "s1"; Name = "Sheet1"; Model = wideBigModel; Visible = true } ]
      SheetOrder = [ "s1" ]; ActiveSheet = "s1"; Tables = Map.empty; Names = Map.empty }
let wideSum = { Sheet = "s1"; Cell = { Row = 1; Col = 0 } }
let wideSumData = cell (CellValue.Formula ("=SUM(A1:DKJ1)", CellValue.Empty))
let (wideBigWb1, wideBigEng1) = updateCell wideSum (CellValue.Formula ("=SUM(A1:DKJ1)", CellValue.Empty)) wideSumData wideBigWb0 emptyEngine
ok "wide-bigrange-no-cell-edges"
   (Set.isEmpty wideBigEng1.Metadata.[wideSum].Dependencies && List.length wideBigEng1.Metadata.[wideSum].DepRects = 1)
   (sprintf "deps=%d rects=%d" (Set.count wideBigEng1.Metadata.[wideSum].Dependencies) (List.length wideBigEng1.Metadata.[wideSum].DepRects))
let (wideBigWb2, wideBigEng2) = smartRecalculate wideBigWb1 wideBigEng1
ok "wide-bigrange-initial-sum" (readWb wideSum wideBigWb2 = CellValue.Number 3000.0) (sprintf "wide sum=%A" (readWb wideSum wideBigWb2))
let crd1 = { Sheet = "s1"; Cell = { Row = 0; Col = 2499 } }
let (wideBigWb3, wideBigEng3) = updateCell crd1 (CellValue.Number 11.0) (cell (CellValue.Number 11.0)) wideBigWb2 wideBigEng2
ok "wide-bigrange-edit-dirties-via-rect" (Set.contains wideSum wideBigEng3.DirtySet) (sprintf "dirty=%A" (Set.count wideBigEng3.DirtySet))
let (wideBigWb4, _) = smartRecalculate wideBigWb3 wideBigEng3
ok "wide-bigrange-recalc" (readWb wideSum wideBigWb4 = CellValue.Number 3010.0) (sprintf "wide sum=%A" (readWb wideSum wideBigWb4))
let outsideWide = { Sheet = "s1"; Cell = { Row = 3; Col = 2999 } }
let (_, wideOutsideEng) = updateCell outsideWide (CellValue.Number 1.0) (cell (CellValue.Number 1.0)) wideBigWb2 wideBigEng2
ok "wide-bigrange-outside-edit-clean" (not (Set.contains wideSum wideOutsideEng.DirtySet)) (sprintf "dirty=%A" (Set.count wideOutsideEng.DirtySet))

let rectIndexFormulaCount = 5000
let rectIndexRangeHeight = 3000
let rectIndexGap = 1
let rectIndexRows = rectIndexFormulaCount * (rectIndexRangeHeight + rectIndexGap) + 10
let rectIndexA1 row = sprintf "A%d" (row + 1)
let rectIndexCells =
    seq {
        for index in 0 .. rectIndexFormulaCount - 1 do
            let startRow = index * (rectIndexRangeHeight + rectIndexGap)
            let endRow = startRow + rectIndexRangeHeight - 1
            yield { Row = startRow; Col = 0 }, cell (CellValue.Number 1.0)
            yield
                { Row = index; Col = 1 },
                cell (CellValue.Formula (sprintf "=SUM(%s:%s)" (rectIndexA1 startRow) (rectIndexA1 endRow), CellValue.Empty))
    }
    |> Map.ofSeq
let rectIndexModel = { createEmpty rectIndexRows 3 with Cells = rectIndexCells }
let rectIndexWb =
    { Sheets = Map.ofList [ "rect", { Id = "rect"; Name = "Rect"; Model = rectIndexModel; Visible = true } ]
      SheetOrder = [ "rect" ]; ActiveSheet = "rect"; Tables = Map.empty; Names = Map.empty }
let mutable rectIndexEngine = emptyEngine
for index in 0 .. rectIndexFormulaCount - 1 do
    let startRow = index * (rectIndexRangeHeight + rectIndexGap)
    let endRow = startRow + rectIndexRangeHeight - 1
    rectIndexEngine <-
        buildDependencies
            { Sheet = "rect"; Cell = { Row = index; Col = 1 } }
            (sprintf "=SUM(%s:%s)" (rectIndexA1 startRow) (rectIndexA1 endRow))
            rectIndexWb
            rectIndexEngine
rectIndexEngine <- rebuildCalcChain rectIndexWb rectIndexEngine
let rectIndexGapRow =
    (rectIndexFormulaCount / 2) * (rectIndexRangeHeight + rectIndexGap) + rectIndexRangeHeight
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
let rectIndexGapBefore = System.GC.GetAllocatedBytesForCurrentThread()
let rectIndexGapSw = System.Diagnostics.Stopwatch.StartNew()
let _, rectIndexGapEngine =
    updateCell
        { Sheet = "rect"; Cell = { Row = rectIndexGapRow; Col = 0 } }
        (CellValue.Number 2.0)
        (cell (CellValue.Number 2.0))
        rectIndexWb
        rectIndexEngine
rectIndexGapSw.Stop()
let rectIndexGapBytes = System.GC.GetAllocatedBytesForCurrentThread() - rectIndexGapBefore
ok "bigrange-rect-index-gap-edit"
   (Set.isEmpty rectIndexGapEngine.DirtySet
    && rectIndexGapBytes < 80000L
    && rectIndexGapSw.Elapsed.TotalMilliseconds < 20.0)
   (sprintf "dirty=%d elapsed=%.1f ms allocated=%d"
        (Set.count rectIndexGapEngine.DirtySet)
        rectIndexGapSw.Elapsed.TotalMilliseconds
        rectIndexGapBytes)

[<System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)>]
let rectIndexMetadataWeakReference () =
    let formulaCount = 200
    let rangeHeight = 3000
    let gap = 1
    let rows = formulaCount * (rangeHeight + gap) + 10
    let a1 row = sprintf "A%d" (row + 1)
    let cells =
        seq {
            for index in 0 .. formulaCount - 1 do
                let startRow = index * (rangeHeight + gap)
                let endRow = startRow + rangeHeight - 1
                yield { Row = startRow; Col = 0 }, cell (CellValue.Number 1.0)
                yield
                    { Row = index; Col = 1 },
                    cell (CellValue.Formula (sprintf "=SUM(%s:%s)" (a1 startRow) (a1 endRow), CellValue.Empty))
        }
        |> Map.ofSeq
    let model = { createEmpty rows 3 with Cells = cells }
    let workbook =
        { Sheets = Map.ofList [ "rect-weak", { Id = "rect-weak"; Name = "RectWeak"; Model = model; Visible = true } ]
          SheetOrder = [ "rect-weak" ]; ActiveSheet = "rect-weak"; Tables = Map.empty; Names = Map.empty }
    let mutable engine = emptyEngine
    for index in 0 .. formulaCount - 1 do
        let startRow = index * (rangeHeight + gap)
        let endRow = startRow + rangeHeight - 1
        engine <-
            buildDependencies
                { Sheet = "rect-weak"; Cell = { Row = index; Col = 1 } }
                (sprintf "=SUM(%s:%s)" (a1 startRow) (a1 endRow))
                workbook
                engine
    engine <- rebuildCalcChain workbook engine
    updateCell
        { Sheet = "rect-weak"; Cell = { Row = rangeHeight; Col = 0 } }
        (CellValue.Number 2.0)
        (cell (CellValue.Number 2.0))
        workbook
        engine
    |> ignore
    System.WeakReference(engine.Metadata)
let rectIndexMetadataWeak = rectIndexMetadataWeakReference ()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
ok "bigrange-rect-index-cache-releases-engine"
   (not rectIndexMetadataWeak.IsAlive)
   "rectangle dependency index cache should not retain discarded calc metadata"

printfn ""
printfn "Stage 5 calc-engine tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
