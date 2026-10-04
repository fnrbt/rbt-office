#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// Incremental recalculation (Recalc): after any edit of cells - one cell, a
// block pasted or cleared, formulas entered or overwritten, an undo that puts
// cells back with the values they had - only the formulas that read the
// changed cells are evaluated, in dependency order, and the result is the
// result of a full recalculation of the edited workbook. Checked on built
// workbooks (ranges, cross-sheet references, names, tables, cycles, spills,
// volatile functions) and with long random edit sequences on the real-world
// workbooks.

open System
open System.IO
open Rbt.Spreadsheet.Model
module Formula = Rbt.Spreadsheet.Formula
module Wb = Rbt.Spreadsheet.Workbook
module Recalc = Rbt.Spreadsheet.Recalc
module Bridge = Rbt.Spreadsheet.XlsxBridge

let mutable pass = 0
let mutable fail = 0
let ok label cond detail =
    if cond then pass <- pass + 1
    else (fail <- fail + 1; printfn "FAIL %-40s %s" label detail)
let eq label (got: 'a) (expected: 'a) = ok label (got = expected) (sprintf "expected %A  got %A" expected got)

let number value = { Value = CellValue.Number value; Format = defaultFormat; Comment = None }
let text value = { Value = CellValue.String value; Format = defaultFormat; Comment = None }
let formula body = { Value = CellValue.Formula (body, CellValue.Empty); Format = defaultFormat; Comment = None }
let at row col : CellAddress = { Row = row; Col = col }

let setCells (sheetId: string) (cells: (CellAddress * CellData option) list) (wb: Wb.Workbook) =
    let mutable model = wb.Sheets.[sheetId].Model |> resolveSheet
    for addr, data in cells do
        model <-
            match data with
            | Some data -> setCell addr data model
            | None -> clearCell addr model
    Wb.updateSheetModel sheetId model wb

let valueAt (sheetId: string) row col (wb: Wb.Workbook) =
    match getCell (at row col) (resolveSheet wb.Sheets.[sheetId].Model) with
    | Some cell -> unwrapFormulaValue cell.Value
    | None -> CellValue.Empty

/// The first cell in which two workbooks differ.
let difference (left: Wb.Workbook) (right: Wb.Workbook) =
    left.SheetOrder
    |> List.tryPick (fun sheetId ->
        let a = (resolveSheet left.Sheets.[sheetId].Model).Cells
        let b = (resolveSheet right.Sheets.[sheetId].Model).Cells
        if a = b then None
        else
            let keys = Seq.append (a |> Seq.map (fun kv -> kv.Key)) (b |> Seq.map (fun kv -> kv.Key)) |> Seq.distinct
            keys |> Seq.tryPick (fun key ->
                let x, y = Map.tryFind key a |> Option.map (fun c -> c.Value), Map.tryFind key b |> Option.map (fun c -> c.Value)
                if x <> y then Some (sprintf "%s!R%dC%d: %A vs %A" sheetId key.Row key.Col x y) else None))

/// The full recalculation, run until it gives what it was given: one run
/// visits the sheets in tab order a few times, which leaves a long chain
/// through several sheets (a name on another sheet read by a formula that a
/// third sheet reads ...) short of its values. The incremental engine follows
/// the dependencies, so its result is the settled one.
let fullRecalculation (wb: Wb.Workbook) =
    let mutable result = Formula.recalculateWorkbook wb
    let mutable runs = 1
    let mutable settled = false
    while not settled && runs < 8 do
        let again = Formula.recalculateWorkbook result
        settled <- (difference again result).IsNone
        result <- again
        runs <- runs + 1
    result

/// An edit both ways: incrementally, and by a full recalculation.
type Outcome = Incremental of Recalc.RecalcStats | Declined
let mutable oneRunSteps = 0
let mutable oneRunShort = 0
let mutable engine = Unchecked.defaultof<Recalc.Engine>
let mutable current = Unchecked.defaultof<Wb.Workbook>
let start (wb: Wb.Workbook) =
    current <- fullRecalculation wb
    engine <- Recalc.build current
let edit label (hint: (string * CellAddress) list option) (change: Wb.Workbook -> Wb.Workbook) =
    let edited = change current
    let full = fullRecalculation edited
    // One run of the full recalculation reaches the settled values.
    oneRunSteps <- oneRunSteps + 1
    if (difference (Formula.recalculateWorkbook edited) full).IsSome then oneRunShort <- oneRunShort + 1
    match Recalc.recalculate engine edited hint with
    | Some result ->
        (match difference result full with
         | None -> pass <- pass + 1
         | Some detail -> fail <- fail + 1; printfn "FAIL %s: incremental differs from full at %s" label detail)
        current <- full
        // (The engine goes on from its own result: any difference would show at a later step.)
        Incremental (Recalc.lastStats engine)
    | None ->
        current <- full
        if not (Recalc.adopt engine full) then engine <- Recalc.build full
        Declined

let twoSheets () =
    let wb = Wb.createWorkbook () |> Wb.addSheet "Data"
    let dataId = wb.SheetOrder |> List.last
    let wb = wb |> setCells "sheet1" [ for row in 0 .. 49 -> at row 0, Some (number (float (row + 1))) ]
    let wb =
        wb |> setCells "sheet1"
            [ at 0 1, Some (formula "=A1*2")
              at 1 1, Some (formula "=B1+A2")
              at 2 1, Some (formula "=SUM(A1:A50)")
              at 3 1, Some (formula "=SUM(A:A)")
              at 4 1, Some (formula "=B3+B4")
              at 5 1, Some (formula "=IF(A6>100,\"big\",\"small\")")
              at 6 1, Some (formula "=Data!A1+B1")
              at 7 1, Some (formula "=SUM(A10:C12)") ]
    let wb = wb |> setCells dataId [ at 0 0, Some (number 7.0); at 0 1, Some (formula "=Sheet1!B5*2"); at 1 1, Some (formula "=SUM(Sheet1!A1:A10)") ]
    wb, dataId

// ---- single cells, ranges, cross-sheet -------------------------------------------
let wb0, dataId = twoSheets ()
start wb0
eq "calculated: SUM(A1:A50)" (valueAt "sheet1" 2 1 current) (CellValue.Number 1275.0)
let stats label outcome = match outcome with Incremental s -> s | Declined -> (ok label false "declined"; { ChangedCells = 0; ReachedFormulas = 0; EvaluatedFormulas = 0 } : Recalc.RecalcStats)
let typed = edit "a number typed" (Some [ "sheet1", at 0 0 ]) (setCells "sheet1" [ at 0 0, Some (number 11.0) ]) |> stats "a number typed"
eq "typed A1: the formulas that read it, and theirs" (valueAt "sheet1" 0 1 current, valueAt "sheet1" 2 1 current, valueAt dataId 0 1 current) (CellValue.Number 22.0, CellValue.Number 1285.0, CellValue.Number 5140.0)
ok "typed A1: not every formula was evaluated" (typed.EvaluatedFormulas >= 6 && typed.EvaluatedFormulas < Recalc.formulaCount engine) (sprintf "%A" typed)
let unread = edit "a cell nothing reads" None (setCells "sheet1" [ at 60 5, Some (number 1.0) ]) |> stats "unread"
eq "a cell nothing reads: nothing evaluated" (unread.ChangedCells, unread.EvaluatedFormulas) (1, 0)
let wholeColumn = edit "a cell only a whole-column range reads" None (setCells "sheet1" [ at 500 0, Some (number 5.0) ]) |> stats "whole column"
eq "whole column: SUM(A:A) and the two formulas after it" (wholeColumn.EvaluatedFormulas, valueAt "sheet1" 3 1 current) (3, CellValue.Number 1290.0)
let sameValue = edit "the value it had" None (setCells "sheet1" [ at 5 0, Some (number 6.0) ]) |> stats "same value"
eq "a cell set to the value it had: nothing changed" (sameValue.ChangedCells, sameValue.EvaluatedFormulas) (0, 0)
let unchangedResult = edit "a value that leaves the formula's result" None (setCells "sheet1" [ at 5 0, Some (number 50.0) ]) |> stats "unchanged result"
ok "IF gives the same text: what reads the sums is evaluated, nothing after the IF" (unchangedResult.EvaluatedFormulas >= 3) (sprintf "%A" unchangedResult)
let block = edit "a block pasted" None (setCells "sheet1" [ for row in 9 .. 11 do for col in 0 .. 2 -> at row col, Some (number (float (row * col))) ]) |> stats "block"
ok "a block pasted: its cells, once" (block.ChangedCells >= 6 && valueAt "sheet1" 7 1 current <> CellValue.Empty) (sprintf "%A" block)
let cleared = edit "a block cleared" None (setCells "sheet1" [ for row in 20 .. 39 -> at row 0, None ]) |> stats "cleared"
eq "a block cleared" (cleared.ChangedCells, valueAt "sheet1" 2 1 current = valueAt "sheet1" 2 1 (Formula.recalculateWorkbook current)) (20, true)
edit "a formula entered" (Some [ "sheet1", at 0 3 ]) (setCells "sheet1" [ at 0 3, Some (formula "=B3/2") ]) |> ignore
edit "a formula changed" None (setCells "sheet1" [ at 0 3, Some (formula "=B4/4+Data!B2") ]) |> ignore
edit "what the changed formula no longer reads" None (setCells "sheet1" [ at 1 0, Some (number 77.0) ]) |> ignore
edit "a formula overwritten by a number" None (setCells "sheet1" [ at 0 1, Some (number 3.0) ]) |> ignore
edit "the other sheet edited" None (setCells dataId [ at 0 0, Some (number 70.0) ]) |> ignore
edit "a formula removed" None (setCells "sheet1" [ at 2 1, None ]) |> ignore
// An undo puts cells back as they were, with the values the formulas had then.
let beforeUndo = current
edit "edits to undo" None (setCells "sheet1" [ at 3 0, Some (number 1000.0); at 4 1, Some (formula "=B4*3"); at 6 1, None ]) |> ignore
edit "the undo" None (fun wb ->
    let old = (resolveSheet beforeUndo.Sheets.["sheet1"].Model)
    wb |> setCells "sheet1" [ for addr in [ at 3 0; at 4 1; at 6 1 ] -> addr, getCell addr old ]) |> ignore
ok "after the undo the workbook is what it was" ((difference current beforeUndo).IsNone) (sprintf "%A" (difference current beforeUndo))
// A format alone is no change.
let formatted = edit "a format" None (fun wb -> Wb.updateSheetModel "sheet1" (setCellBgColor (at 0 0) "#ff0000" wb.Sheets.["sheet1"].Model) wb) |> stats "format"
eq "a format changes no value" (formatted.ChangedCells, formatted.EvaluatedFormulas) (0, 0)

// ---- what is left to the full recalculation -------------------------------------
let declined label outcome = ok label (outcome = Declined) "was calculated incrementally"
declined "a formula that reads itself" (edit "cycle" None (setCells "sheet1" [ at 30 3, Some (formula "=D31+1") ]))
eq "the cycle is marked" (valueAt "sheet1" 30 3 current) (CellValue.Error Formula.Err.cycle)
ok "an edit elsewhere is incremental although a cycle exists" (edit "beside a cycle" None (setCells "sheet1" [ at 0 0, Some (number 1.0) ]) <> Declined) ""
declined "a cycle through two cells" (edit "cycle 2" None (setCells "sheet1" [ at 31 3, Some (formula "=D33"); at 32 3, Some (formula "=D32") ]))
edit "the cycles removed" None (setCells "sheet1" [ at 30 3, None; at 31 3, None; at 32 3, None ]) |> ignore
ok "then edits are incremental again" (edit "after the cycles" None (setCells "sheet1" [ at 0 0, Some (number 2.0) ]) <> Declined) ""
declined "a spilling formula entered" (edit "spill" None (setCells "sheet1" [ at 40 5, Some (formula "=SEQUENCE(3)") ]))
eq "the spill is there" (valueAt "sheet1" 42 5 current) (CellValue.Number 3.0)
ok "an edit that does not reach the spill is incremental" (edit "beside a spill" None (setCells "sheet1" [ at 0 0, Some (number 3.0) ]) <> Declined) ""
declined "a spilled cell overwritten" (edit "into the spill" None (setCells "sheet1" [ at 41 5, Some (number 9.0) ]))
eq "the spill is blocked" (valueAt "sheet1" 40 5 current) (CellValue.Error Formula.Err.spill)
declined "the blocking cell cleared" (edit "unblock" None (setCells "sheet1" [ at 41 5, None ]))
eq "the spill is back" (valueAt "sheet1" 42 5 current) (CellValue.Number 3.0)
edit "the spill removed" None (setCells "sheet1" [ at 40 5, None ]) |> ignore
declined "OFFSET: precedents chosen while evaluating" (edit "offset" None (setCells "sheet1" [ at 45 5, Some (formula "=OFFSET(A1,1,0)") ]))
ok "the workbook is unsupported while it holds one" (Recalc.isUnsupported engine) ""
declined "every edit is left to the full recalculation" (edit "with offset" None (setCells "sheet1" [ at 0 0, Some (number 4.0) ]))
edit "the OFFSET removed" None (setCells "sheet1" [ at 45 5, None ]) |> ignore
ok "and supported again without it" (not (Recalc.isUnsupported engine) && edit "after offset" None (setCells "sheet1" [ at 0 0, Some (number 5.0) ]) <> Declined) ""
// More than cells: a defined name.
declined "a defined name added" (edit "name" None (fun wb -> { wb with Names = wb.Names |> Map.add (None, "RATE") "=Sheet1!$A$1" }))
edit "a formula with the name" None (setCells "sheet1" [ at 46 5, Some (formula "=RATE*2") ]) |> ignore
let named = edit "the named cell typed" None (setCells "sheet1" [ at 0 0, Some (number 21.0) ]) |> stats "named"
eq "a name is a precedent" (valueAt "sheet1" 46 5 current) (CellValue.Number 42.0)

// ---- random edit sequences on real workbooks -------------------------------------
let root = Path.Combine(__SOURCE_DIRECTORY__, "..")
let load (relative: string) =
    let bytes = File.ReadAllBytes(Path.Combine(root, relative))
    (Bridge.Import.importWorkbook (XlsxReader.loadBytesNative bytes |> Option.get)).Context

let randomEdits (label: string) (seed: int) (steps: int) (wb: Wb.Workbook) =
    let random = Random seed
    start wb
    let settled = Formula.recalculateWorkbook current
    // (Volatile functions - RAND, NOW - give other values at every run.)
    if (difference settled current).IsSome then
        printfn "  %s: the full recalculation does not settle (%s) - left out" label (difference settled current).Value
        0, 0
    else
        let sheets = current.SheetOrder |> List.filter (fun id -> not (Map.isEmpty (resolveSheet current.Sheets.[id].Model).Cells)) |> List.toArray
        let mutable incremental = 0
        let mutable declinedSteps = 0
        let mutable evaluated = 0
        let history = ResizeArray<string * (CellAddress * CellData option) list>()
        let failuresBefore = fail
        for step in 1 .. steps do
            if fail = failuresBefore then
                let sheetId = sheets.[random.Next sheets.Length]
                let model = resolveSheet current.Sheets.[sheetId].Model
                let cells = model.Cells |> Map.toArray
                let anyCell () = fst cells.[random.Next cells.Length]
                let numbers = cells |> Array.filter (fun (_, cell) -> match cell.Value with CellValue.Number _ -> true | _ -> false)
                let formulas = cells |> Array.filter (fun (_, cell) -> match cell.Value with CellValue.Formula _ -> true | _ -> false)
                let refOf (addr: CellAddress) = sprintf "%s%d" (columnToLetter addr.Col) (addr.Row + 1)
                let kind = random.Next 12
                let name, changes =
                    match kind with
                    | 0 | 1 | 2 when numbers.Length > 0 ->
                        "a number changed", [ fst numbers.[random.Next numbers.Length], Some (number (float (random.Next 2000) / 7.0)) ]
                    | 3 ->
                        let origin = anyCell ()
                        "a block of numbers pasted", [ for r in 0 .. random.Next 12 do for c in 0 .. random.Next 5 -> at (origin.Row + r) (origin.Col + c), Some (number (float (random.Next 500))) ]
                    | 4 ->
                        let origin = anyCell ()
                        "a block cleared", [ for r in 0 .. random.Next 8 do for c in 0 .. random.Next 4 -> at (origin.Row + r) (origin.Col + c), None ]
                    | 5 ->
                        let a, b = anyCell (), anyCell ()
                        "a formula entered", [ at (a.Row + random.Next 3) (a.Col + 1 + random.Next 3), Some (formula (sprintf "=%s+SUM(%s:%s)" (refOf a) (refOf (at (min a.Row b.Row) (min a.Col b.Col))) (refOf (at (min (min a.Row b.Row + 40) (max a.Row b.Row)) (max a.Col b.Col))))) ]
                    | 6 when formulas.Length > 0 ->
                        "a formula overwritten", [ fst formulas.[random.Next formulas.Length], Some (number (float (random.Next 100))) ]
                    | 7 when formulas.Length > 0 ->
                        let addr, cell = formulas.[random.Next formulas.Length]
                        let target = anyCell ()
                        "a formula copied elsewhere", [ at (target.Row + 1) target.Col, Some { cell with Value = (match cell.Value with CellValue.Formula (f, _) -> CellValue.Formula (f, CellValue.Empty) | v -> v) } ]
                    | 8 -> "a text typed", [ anyCell (), Some (text "typed") ]
                    | 9 when history.Count > 0 ->
                        // The cells of an earlier step as they were before it (an undo, however long ago).
                        let at = random.Next history.Count
                        let _, undo = history.[at]
                        "cells put back", undo
                    | _ -> "a number in an empty cell", [ (let a = anyCell () in at (a.Row + random.Next 4) (a.Col + random.Next 4)), Some (number (float (random.Next 50))) ]
                let changes = if kind = 9 && history.Count > 0 then changes else changes |> List.distinctBy fst
                let historySheet = if kind = 9 && history.Count > 0 then fst history.[0] else sheetId
                let targetSheet = if name = "cells put back" then (history |> Seq.find (fun (_, undo) -> obj.ReferenceEquals(undo, changes)) |> fst) else sheetId
                ignore historySheet
                let before = resolveSheet current.Sheets.[targetSheet].Model
                let undo = changes |> List.map (fun (addr, _) -> addr, getCell addr before)
                history.Add (targetSheet, undo)
                let hint = if changes.Length = 1 && random.Next 2 = 0 then Some [ targetSheet, fst changes.Head ] else None
                match edit (sprintf "%s step %d (%s)" label step name) hint (setCells targetSheet changes) with
                | Incremental s ->
                    incremental <- incremental + 1
                    evaluated <- evaluated + s.EvaluatedFormulas
                | Declined -> declinedSteps <- declinedSteps + 1
        printfn "  %s: %d formulas; %d edits incremental (%d formulas evaluated in all), %d left to the full recalculation%s"
            label (Recalc.formulaCount engine) incremental evaluated declinedSteps (if Recalc.isUnsupported engine then " (unsupported: OFFSET / INDIRECT / 3-D / spill-range references)" else "")
        incremental, declinedSteps

let mutable incrementalTotal = 0
for file, seed, steps in [ "real-world/milk-street-office.xlsx", 3, 45
                           "real-world/lyric-multifamily.xlsx", 4, 35
                           "real-world/dcf-stock-analysis.xlsx", 5, 45
                           "real-world/saopaulo-condos.xlsx", 6, 35
                           "real-world/jumeirah-hotel.xlsx", 7, 35
                           "real-world/dashboard-sales-analysis.xlsx", 8, 30 ] do
    let incremental, _ = randomEdits (Path.GetFileNameWithoutExtension file) seed steps (load file)
    incrementalTotal <- incrementalTotal + incremental
// (The synthetic workbook again, so that the sequences cover one the engine supports whatever the files hold.)
let incrementalSynthetic, _ = randomEdits "two sheets" 9 150 (fst (twoSheets ()))
ok "the random sequences were mostly incremental" (incrementalSynthetic > 120 && incrementalTotal > 60) (sprintf "%d synthetic, %d on files" incrementalSynthetic incrementalTotal)

// (Up to six passes over two sheets: a chain that crosses between them more often than that is
// still short after one run - a few steps of the sequences on the schedule workbooks.)
ok "one run of the full recalculation settles nearly every edit of workbooks whose values travel between sheets"
    (oneRunSteps > 300 && oneRunShort * 20 < oneRunSteps) (sprintf "%d of %d steps short" oneRunShort oneRunSteps)
printfn "  one run of the full recalculation was short of the settled values in %d of %d steps" oneRunShort oneRunSteps

// A date function fed something far outside the calendar is #NUM!, on every
// platform, and does not take the recalculation down (.NET's DateTime threw).
do
    let wb = Wb.createWorkbook () |> setCells "sheet1" [ at 0 0, Some (number 1e12); at 0 1, Some (formula "=EOMONTH(A1,1)"); at 0 2, Some (formula "=EDATE(A1,5)"); at 0 3, Some (formula "=YEAR(A1)"); at 1 1, Some (formula "=1+1") ]
    let calculated = Formula.recalculateWorkbook wb
    eq "the recalculation goes on past them" (valueAt "sheet1" 1 1 calculated) (CellValue.Number 2.0)
    ok "and they are errors or held at the calendar's end, not exceptions"
        ([ 1; 2; 3 ] |> List.forall (fun col -> match valueAt "sheet1" 0 col calculated with CellValue.Error _ | CellValue.Number _ -> true | _ -> false)) (sprintf "%A" [ for col in 1 .. 3 -> valueAt "sheet1" 0 col calculated ])

// A paste that makes a formula read itself through another sheet is a circular
// reference: the engine leaves it to the full recalculation and says why.
do
    let wb = Wb.createWorkbook () |> Wb.addSheet "Other"
    let other = wb.SheetOrder |> List.last
    let wb =
        wb |> setCells "sheet1" [ at 0 0, Some (number 3.0); at 1 0, Some (formula "=Other!A1*2") ]
           |> setCells other [ at 0 0, Some (formula "=Sheet1!A1+1") ]
    start wb
    eq "through the other sheet" (valueAt "sheet1" 1 0 current) (CellValue.Number 8.0)
    let edited = current |> setCells "sheet1" [ at 0 0, Some (formula "=A2+1") ]
    eq "the input overwritten by a formula that reads what depends on it: declined, as a circular reference"
        (Recalc.recalculate engine edited None |> Option.isSome, Recalc.lastDecline engine) (false, "a circular reference among the formulas to evaluate")

printfn "stage32-recalc: %d passed, %d failed" pass fail
if fail > 0 then exit 1
