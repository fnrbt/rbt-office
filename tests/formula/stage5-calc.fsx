#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Calc
open Rbt.Spreadsheet.Formula

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
let ai = { Sheet = "s1"; Cell = { Row = 2499; Col = 0 } }   // a data cell inside the rectangle
let (bigWb3, bigEng3) = updateCell ai (CellValue.Number 11.0) (cell (CellValue.Number 11.0)) bigWb2 bigEng2
ok "bigrange-edit-dirties-via-rect" (Set.contains bb1 bigEng3.DirtySet) (sprintf "dirty=%A" (Set.count bigEng3.DirtySet))
let (bigWb4, _) = smartRecalculate bigWb3 bigEng3
ok "bigrange-recalc" (readWb bb1 bigWb4 = CellValue.Number 5010.0) (sprintf "B1=%A" (readWb bb1 bigWb4)) // 4999×1 + 11

// --- Incremental-vs-full equivalence through the guarded edit API --------
let oneSheetWb (model: SpreadsheetModel) =
    { Sheets = Map.ofList [ "s1", { Id = "s1"; Name = "Sheet1"; Model = model; Visible = true } ]
      SheetOrder = [ "s1" ]; ActiveSheet = "s1"; Tables = Map.empty; Names = Map.empty }

let activeModelOf (wb: Workbook) = (Map.find wb.ActiveSheet wb.Sheets).Model

let fullEdit (addr: CellAddress) (cellData: CellData) (wb: Workbook) =
    let model = activeModelOf wb |> setCell addr cellData
    let recalculated = recalcAndSpill (Some wb) (Some wb.ActiveSheet) model
    updateSheetModel wb.ActiveSheet recalculated wb

let incrementalEdit (addr: CellAddress) (cellData: CellData) (wb: Workbook) (engine: CalcEngine option) =
    let sc = { Sheet = wb.ActiveSheet; Cell = addr }
    match trySingleCellEdit sc cellData wb engine with
    | IncrementalApplied (wb', engine') -> wb', Some engine', false
    | IncrementalFallback _ ->
        fullEdit addr cellData wb, None, true

let recalcRelevantEqual (left: SpreadsheetModel) (right: SpreadsheetModel) =
    left.Cells = right.Cells
    && left.Spill = right.Spill
    && left.RowCount = right.RowCount
    && left.ColCount = right.ColCount

let assertEquivalent label edits initialWb =
    let mutable incWb = initialWb
    let mutable fullWb = initialWb
    let mutable engine = None
    let mutable fellBack = false
    for addr, cd in edits do
        let wb', eng', fallback = incrementalEdit addr cd incWb engine
        incWb <- wb'
        engine <- eng'
        fellBack <- fellBack || fallback
        fullWb <- fullEdit addr cd fullWb
    ok label (recalcRelevantEqual (activeModelOf incWb) (activeModelOf fullWb))
       (sprintf "incremental=%A full=%A fallback=%b" (activeModelOf incWb).Cells (activeModelOf fullWb).Cells fellBack)
    fellBack

let cnum n = cell (CellValue.Number n)
let cform f = cell (CellValue.Formula (f, CellValue.Empty))

let chainInitial =
    { createEmpty 20 5 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cnum 1.0
                  { Row = 0; Col = 1 }, cform "=A1+1"
                  { Row = 0; Col = 2 }, cform "=B1+1" ] }
let chainFallback =
    assertEquivalent "incremental-equiv-chain"
        [ { Row = 0; Col = 0 }, cnum 10.0
          { Row = 0; Col = 0 }, cnum 20.0 ]
        (oneSheetWb chainInitial)
ok "incremental-chain-no-fallback" (not chainFallback) "chain edit should use incremental path"

let volatileInitial =
    { createEmpty 20 5 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cnum 5.0
                  { Row = 0; Col = 1 }, cform "=IF(TODAY()>0,A1+1,0)" ] }
let volatileFallback =
    assertEquivalent "incremental-equiv-volatile"
        [ { Row = 1; Col = 0 }, cnum 1.0
          { Row = 1; Col = 0 }, cnum 2.0 ]
        (oneSheetWb volatileInitial)
ok "incremental-volatile-no-fallback" (not volatileFallback) "volatile edit should use incremental path"

let spillInitial =
    { createEmpty 20 5 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cform "=SEQUENCE(2,1)"
                  { Row = 0; Col = 2 }, cnum 1.0 ] }
let spillFallback =
    assertEquivalent "incremental-equiv-spill-fallback"
        [ { Row = 0; Col = 2 }, cnum 9.0 ]
        (oneSheetWb spillInitial)
ok "incremental-spill-fallback" spillFallback "array formula should force full recalc"

// A workbook that has a calculated spilling formula is still edited
// incrementally, as long as the edit cannot reach the spill: the edited cell
// is not one the formula reads, not inside its spill range, and not emptied.
let spillCalculated =
    // A1 = SEQUENCE(3,1) spilling A1:A3 from C1; D1 = C1*2; E5 = SUM(A1:A3); B1..B3 plain.
    { createEmpty 20 8 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cform "=SEQUENCE(3,1,C1)"
                  { Row = 0; Col = 2 }, cnum 1.0
                  { Row = 0; Col = 3 }, cform "=C1*2"
                  { Row = 4; Col = 4 }, cform "=SUM(A1:A3)"
                  { Row = 0; Col = 6 }, cnum 5.0
                  { Row = 1; Col = 6 }, cform "=G1+1" ] }
    |> recalcAndSpill None None
ok "spill-calculated-precondition" (spillCalculated.Spill.Count = 2 && getCellValue { Row = 2; Col = 0 } spillCalculated = CellValue.Number 3.0)
   (sprintf "spill=%A" spillCalculated.Spill)
let editOutcome (addr: CellAddress) (cd: CellData) =
    match trySingleCellEdit { Sheet = "s1"; Cell = addr } cd (oneSheetWb spillCalculated) None with
    | IncrementalApplied _ -> "incremental"
    | IncrementalFallback reason -> "fallback: " + reason
ok "incremental-unrelated-edit-with-spill-present" (editOutcome { Row = 0; Col = 6 } (cnum 9.0) = "incremental") (editOutcome { Row = 0; Col = 6 } (cnum 9.0))
ok "incremental-new-cell-with-spill-present" (editOutcome { Row = 10; Col = 5 } (cell (CellValue.String "note")) = "incremental") (editOutcome { Row = 10; Col = 5 } (cell (CellValue.String "note")))
ok "incremental-plain-formula-with-spill-present" (editOutcome { Row = 10; Col = 5 } (cform "=G1*3") = "incremental") (editOutcome { Row = 10; Col = 5 } (cform "=G1*3"))
ok "fallback-edit-feeds-spill" ((editOutcome { Row = 0; Col = 2 } (cnum 4.0)).StartsWith "fallback") (editOutcome { Row = 0; Col = 2 } (cnum 4.0))
ok "fallback-edit-inside-spill-range" ((editOutcome { Row = 1; Col = 0 } (cnum 4.0)).StartsWith "fallback") (editOutcome { Row = 1; Col = 0 } (cnum 4.0))
ok "fallback-emptying-a-cell" ((editOutcome { Row = 0; Col = 6 } (cell CellValue.Empty)).StartsWith "fallback") (editOutcome { Row = 0; Col = 6 } (cell CellValue.Empty))
ok "fallback-entering-a-spill-formula" ((editOutcome { Row = 10; Col = 5 } (cform "=SEQUENCE(2)")).StartsWith "fallback") (editOutcome { Row = 10; Col = 5 } (cform "=SEQUENCE(2)"))
// The incremental results equal the full pipeline's, edit after edit.
let spillEquivalent =
    assertEquivalent "incremental-equiv-with-spill-present"
        [ { Row = 0; Col = 6 }, cnum 9.0
          { Row = 10; Col = 5 }, cform "=G2*3"
          { Row = 0; Col = 6 }, cnum 11.0 ]
        (oneSheetWb spillCalculated)
ok "incremental-with-spill-no-fallback" (not spillEquivalent) "edits away from the spill should stay incremental"
// And one that does reach the spill falls back and still agrees.
let spillReached =
    assertEquivalent "incremental-equiv-edit-feeds-spill"
        [ { Row = 0; Col = 6 }, cnum 9.0
          { Row = 0; Col = 2 }, cnum 4.0 ]
        (oneSheetWb spillCalculated)
ok "incremental-feeding-spill-falls-back" spillReached "an edit read by the spill formula must use the full pipeline"

let circularInitial =
    { createEmpty 20 5 with
        Cells =
            Map.ofList
                [ { Row = 0; Col = 0 }, cform "=B1"
                  { Row = 0; Col = 1 }, cform "=A1"
                  { Row = 0; Col = 2 }, cnum 1.0 ] }
let circularFallback =
    assertEquivalent "incremental-equiv-cycle-fallback"
        [ { Row = 0; Col = 2 }, cnum 2.0 ]
        (oneSheetWb circularInitial)
ok "incremental-cycle-fallback" circularFallback "circular chain should force full recalc"

// --- Unary plus on text survives recalculation ---------------------------
// `=+$D$5` where D5 holds text shows the text (unary plus does not coerce),
// through a whole-workbook recalculation and through the incremental engine,
// same sheet and across sheets.
let plusWb =
    let data =
        { createEmpty 100 26 with
            Cells =
                Map.ofList
                    [ { Row = 4; Col = 3 }, cell (CellValue.String "Units:")
                      { Row = 5; Col = 3 }, cell (CellValue.Number 7.0)
                      { Row = 0; Col = 0 }, cell (CellValue.Formula ("=+$D$5", CellValue.String "Units:"))
                      { Row = 1; Col = 0 }, cell (CellValue.Formula ("=+D6", CellValue.Number 7.0))
                      { Row = 2; Col = 0 }, cell (CellValue.Formula ("=-D5", CellValue.Empty)) ] }
    let other =
        { createEmpty 100 26 with
            Cells = Map.ofList [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("+Data!$D$5", CellValue.String "Units:")) ] }
    { Sheets =
        Map.ofList
            [ "sheet1", { Id = "sheet1"; Name = "Data"; Model = data; Visible = true }
              "sheet2", { Id = "sheet2"; Name = "Other"; Model = other; Visible = true } ]
      SheetOrder = [ "sheet1"; "sheet2" ]
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names = Map.empty }
let plusA1 = { Sheet = "sheet1"; Cell = { Row = 0; Col = 0 } }
let plusA2 = { Sheet = "sheet1"; Cell = { Row = 1; Col = 0 } }
let plusA3 = { Sheet = "sheet1"; Cell = { Row = 2; Col = 0 } }
let plusOther = { Sheet = "sheet2"; Cell = { Row = 0; Col = 0 } }
let plusFull = recalculateWorkbook plusWb
ok "unary-plus-text-full-recalc" (readWb plusA1 plusFull = CellValue.String "Units:") (sprintf "A1 = %A" (readWb plusA1 plusFull))
ok "unary-plus-number-full-recalc" (readWb plusA2 plusFull = CellValue.Number 7.0) (sprintf "A2 = %A" (readWb plusA2 plusFull))
ok "unary-minus-text-full-recalc" (readWb plusA3 plusFull = CellValue.Error "VALUE!") (sprintf "A3 = %A" (readWb plusA3 plusFull))
ok "unary-plus-text-xsheet-full-recalc" (readWb plusOther plusFull = CellValue.String "Units:") (sprintf "Other!A1 = %A" (readWb plusOther plusFull))
let plusEngine =
    emptyEngine
    |> buildDependencies plusA1 "=+$D$5" plusWb
    |> buildDependencies plusA2 "=+D6" plusWb
    |> buildDependencies plusA3 "=-D5" plusWb
    |> buildDependencies plusOther "+Data!$D$5" plusWb
    |> rebuildCalcChain plusWb
let plusD5 = { Sheet = "sheet1"; Cell = { Row = 4; Col = 3 } }
let plusEdited, plusEngine1 = updateCell plusD5 (CellValue.String "Suites:") (cell (CellValue.String "Units:")) plusWb plusEngine
let plusIncremental, _ = smartRecalculate plusEdited plusEngine1
ok "unary-plus-text-incremental" (readWb plusA1 plusIncremental = CellValue.String "Suites:") (sprintf "A1 = %A" (readWb plusA1 plusIncremental))
ok "unary-plus-text-xsheet-incremental" (readWb plusOther plusIncremental = CellValue.String "Suites:") (sprintf "Other!A1 = %A" (readWb plusOther plusIncremental))
ok "unary-minus-text-incremental" (readWb plusA3 plusIncremental = CellValue.Error "VALUE!") (sprintf "A3 = %A" (readWb plusA3 plusIncremental))

// In a workbook with a circular chain, an edit nothing reads needs no
// recalculation; one that feeds the cycle still takes the full path.
let cyclicWb =
    oneSheetWb
        { createEmpty 20 8 with
            Cells =
                Map.ofList
                    [ { Row = 0; Col = 0 }, cell (CellValue.Formula ("=B1+C1", CellValue.Number 0.0))
                      { Row = 0; Col = 1 }, cell (CellValue.Formula ("=A1", CellValue.Number 0.0))
                      { Row = 0; Col = 2 }, cnum 1.0 ] }
let cyclicOutcome (addr: CellAddress) (cd: CellData) =
    match trySingleCellEdit { Sheet = "s1"; Cell = addr } cd cyclicWb None with
    | IncrementalApplied (wb, _) -> "incremental", (activeModelOf wb).Cells |> Map.tryFind addr |> Option.map (fun c -> c.Value)
    | IncrementalFallback reason -> "fallback: " + reason, None
ok "cyclic-unread-edit-is-incremental" (cyclicOutcome { Row = 5; Col = 5 } (cnum 7.0) = ("incremental", Some (CellValue.Number 7.0))) (sprintf "%A" (cyclicOutcome { Row = 5; Col = 5 } (cnum 7.0)))
ok "cyclic-edit-feeding-the-cycle-falls-back" ((fst (cyclicOutcome { Row = 0; Col = 2 } (cnum 2.0))).StartsWith "fallback") (sprintf "%A" (cyclicOutcome { Row = 0; Col = 2 } (cnum 2.0)))

printfn ""
printfn "Stage 5 calc-engine tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
