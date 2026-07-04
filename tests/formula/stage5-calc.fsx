#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office-spreadsheet/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

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

printfn ""
printfn "Stage 5 calc-engine tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
