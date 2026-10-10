#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// The full recalculation of a workbook has one dependency order for all its
// sheets: one run reaches the values however they travel between sheets, and
// the formulas on a circular reference - within a sheet, across sheets,
// through a defined name - show the cycle error, as do no others. Formulas
// that choose their precedents while they are evaluated (OFFSET, INDIRECT)
// and 3-D references are part of that order and of the incremental engine.

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
let formula body = { Value = CellValue.Formula (body, CellValue.Empty); Format = defaultFormat; Comment = None }
let at row col : CellAddress = { Row = row; Col = col }
let setCells (sheetId: string) (cells: (CellAddress * CellData option) list) (wb: Wb.Workbook) =
    let mutable model = wb.Sheets.[sheetId].Model |> resolveSheet
    for addr, data in cells do
        model <- (match data with Some data -> setCell addr data model | None -> clearCell addr model)
    Wb.updateSheetModel sheetId model wb
let valueAt (sheetId: string) row col (wb: Wb.Workbook) =
    match getCell (at row col) (resolveSheet wb.Sheets.[sheetId].Model) with
    | Some cell -> unwrapFormulaValue cell.Value
    | None -> CellValue.Empty
let same (left: Wb.Workbook) (right: Wb.Workbook) =
    left.SheetOrder |> List.forall (fun id ->
        let a, b = resolveSheet left.Sheets.[id].Model, resolveSheet right.Sheets.[id].Model
        a.Cells = b.Cells && a.Spill = b.Spill)
let cycle = CellValue.Error Formula.Err.cycle
let n value = CellValue.Number value
let full = Formula.recalculateWorkbook
/// One run gives what a second run gives too.
let settledInOne label (wb: Wb.Workbook) =
    let once = full wb
    ok (label + ": a second run changes nothing") (same once (full once)) ""
    once

// Three sheets, ids sheet1 (Sheet1), sheet2 (Two), sheet3 (Three).
let three () = Wb.createWorkbook () |> Wb.addSheet "Two" |> Wb.addSheet "Three"
let ids = (three ()).SheetOrder
eq "sheet ids" ids [ "sheet1"; "sheet2"; "sheet3" ]

// ---- one run, whatever the order of the sheets ------------------------------------
// A chain that runs against the tab order and crosses the sheets many times:
// the pass that took the sheets one after the other needed a run per crossing.
do
    let links = 30
    let mutable wb = three () |> setCells "sheet3" [ at 0 0, Some (number 1.0) ]
    // Link k is on sheet (3 - k mod 3), reading link k-1 on the sheet after it in tab order.
    let sheetOf k = ids.[2 - (k % 3)]
    let nameOf (id: string) = match id with "sheet1" -> "Sheet1" | "sheet2" -> "Two" | _ -> "Three"
    for k in 1 .. links do
        let from = if k = 1 then "Three!A1" else sprintf "%s!B%d" (nameOf (sheetOf (k - 1))) (k - 1)
        wb <- wb |> setCells (sheetOf k) [ at (k - 1) 1, Some (formula (sprintf "=%s+1" from)) ]
    let once = settledInOne "a chain against the tab order" wb
    eq "the end of the chain after one run" (valueAt (sheetOf links) (links - 1) 1 once) (n (float (links + 1)))
    let old = Formula.recalculateWorkbookSettling wb
    ok "(one run of the pass as it was is short of it)" (valueAt (sheetOf links) (links - 1) 1 old <> n (float (links + 1))) ""
    // An input changed: again one run.
    let edited = once |> setCells "sheet3" [ at 0 0, Some (number 100.0) ]
    eq "after an edit of the input, one run" (valueAt (sheetOf links) (links - 1) 1 (settledInOne "chain edited" edited)) (n (float (links + 100)))
    // ... and through a defined name on another sheet, and a table-less range.
    let named = { once with Names = once.Names |> Map.add (None, "LAST") (sprintf "=%s!$B$%d" (nameOf (sheetOf links)) links) }
    let named = named |> setCells "sheet1" [ at 0 5, Some (formula "=LAST*2"); at 1 5, Some (formula "=SUM(Two!B1:B40)+F1") ] |> setCells "sheet3" [ at 0 0, Some (number 2.0) ]
    let calculated = settledInOne "through a name" named
    eq "a name's target is calculated before what uses the name" (valueAt "sheet1" 0 5 calculated) (n (float (links + 2) * 2.0))

// ---- circular references -------------------------------------------------------------
do
    // Across two sheets.
    let wb = three () |> setCells "sheet1" [ at 0 0, Some (formula "=Two!A1+1"); at 0 1, Some (formula "=A1*2"); at 0 2, Some (formula "=5+5"); at 0 3, Some (formula "=IFERROR(A1,-1)") ]
                      |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!A1+1"); at 1 0, Some (formula "=7") ]
    let calculated = settledInOne "a cycle across two sheets" wb
    eq "both formulas of a cycle across two sheets show the cycle error" (valueAt "sheet1" 0 0 calculated, valueAt "sheet2" 0 0 calculated) (cycle, cycle)
    eq "a formula that reads the cycle has its error; one that catches it, its own value" (valueAt "sheet1" 0 1 calculated, valueAt "sheet1" 0 3 calculated) (cycle, n -1.0)
    eq "formulas beside the cycle are calculated" (valueAt "sheet1" 0 2 calculated, valueAt "sheet2" 1 0 calculated) (n 10.0, n 7.0)
    // An edit that breaks the cycle: the values are there after one run.
    let broken = calculated |> setCells "sheet2" [ at 0 0, Some (number 4.0) ]
    let recovered = settledInOne "the cycle broken" broken
    eq "the cycle broken: values after one run" (valueAt "sheet1" 0 0 recovered, valueAt "sheet1" 0 1 recovered, valueAt "sheet1" 0 3 recovered) (n 5.0, n 10.0, n 5.0)
    // ... and one that makes it again.
    let again = settledInOne "the cycle made again" (recovered |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!B1") ])
    eq "the cycle made again (through B1 now): all three" (valueAt "sheet1" 0 0 again, valueAt "sheet1" 0 1 again, valueAt "sheet2" 0 0 again) (cycle, cycle, cycle)
    let engine = Recalc.build recovered
    let edited = recovered |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!B1") ]
    // The engine calculates an edit that makes a cycle itself, as the full pass does.
    ok "the incremental engine calculates the edit that makes the cycle" (match Recalc.recalculate engine edited None with Some result -> same result (full edited) && valueAt "sheet2" 0 0 result = cycle | None -> false) (Recalc.lastDecline engine)
    let fullResult, fullEngine = Recalc.recalculateAll edited
    ok "the full pass's engine goes on incrementally beside the cycle" (
        match Recalc.recalculate fullEngine (fullResult |> setCells "sheet2" [ at 1 0, Some (number 8.0) ]) None with
        | Some result -> valueAt "sheet2" 1 0 result = n 8.0 && valueAt "sheet1" 0 0 result = cycle
        | None -> false) (Recalc.lastDecline fullEngine)
    ok "... and breaking the cycle is incremental too" (
        match Recalc.recalculate fullEngine (Recalc.recalculate fullEngine (fullResult |> setCells "sheet2" [ at 0 0, Some (number 1.0) ]) None |> Option.defaultValue fullResult) None with
        | Some result -> valueAt "sheet1" 0 0 result = n 2.0 && valueAt "sheet1" 0 1 result = n 4.0
        | None -> false) (Recalc.lastDecline fullEngine)
do
    // Across three sheets, with a tail on each.
    let wb = three () |> setCells "sheet1" [ at 0 0, Some (formula "=Two!A1+1"); at 1 0, Some (formula "=A1+1") ]
                      |> setCells "sheet2" [ at 0 0, Some (formula "=Three!A1+1"); at 1 0, Some (formula "=1+1") ]
                      |> setCells "sheet3" [ at 0 0, Some (formula "=SUM(Sheet1!A1:A1)"); at 1 0, Some (formula "=Two!A2*3") ]
    let calculated = settledInOne "a cycle across three sheets" wb
    eq "a cycle across three sheets (one link a range)" ([ for id in ids -> valueAt id 0 0 calculated ]) [ cycle; cycle; cycle ]
    eq "what is not on it" (valueAt "sheet2" 1 0 calculated, valueAt "sheet3" 1 0 calculated) (n 2.0, n 6.0)
do
    // A formula that reads itself; one that reads itself in a range; a cycle through a defined name.
    let wb = three () |> setCells "sheet1" [ at 0 0, Some (formula "=A1+1"); at 1 1, Some (formula "=SUM(B1:B5)"); at 0 3, Some (formula "=LOOP+1"); at 0 4, Some (formula "=3") ]
    let wb = { wb with Names = wb.Names |> Map.add (None, "LOOP") "=Two!$C$1" } |> setCells "sheet2" [ at 0 2, Some (formula "=Sheet1!D1*2"); at 0 3, Some (formula "=Sheet1!E1") ]
    let calculated = settledInOne "self references and a name" wb
    eq "a formula that reads itself, alone and in a range" (valueAt "sheet1" 0 0 calculated, valueAt "sheet1" 1 1 calculated) (cycle, cycle)
    eq "a cycle through a defined name on another sheet" (valueAt "sheet1" 0 3 calculated, valueAt "sheet2" 0 2 calculated) (cycle, cycle)
    eq "beside them" (valueAt "sheet2" 0 3 calculated) (n 3.0)
    // The name re-pointed: no cycle any more.
    let repointed = { calculated with Names = calculated.Names |> Map.add (None, "LOOP") "=Two!$D$1" }
    eq "the name re-pointed: values after one run" (valueAt "sheet1" 0 3 (settledInOne "name re-pointed" repointed)) (n 4.0)
do
    // Two separate cycles and a formula that reads both; a figure of eight (two cycles sharing a formula).
    let wb = three () |> setCells "sheet1" [ at 0 0, Some (formula "=A2"); at 1 0, Some (formula "=A1"); at 0 1, Some (formula "=Two!B1"); at 0 2, Some (formula "=IFERROR(A1,1)+IFERROR(B1,1)");
                                           at 5 0, Some (formula "=A7+A8"); at 6 0, Some (formula "=A6"); at 7 0, Some (formula "=A6"); at 8 0, Some (formula "=IFERROR(A8,0)+1") ]
                      |> setCells "sheet2" [ at 0 1, Some (formula "=Sheet1!B1") ]
    let calculated = settledInOne "several cycles" wb
    eq "two cycles, a reader of both, a figure of eight, a reader of that" ([ for row, col in [ 0, 0; 1, 0; 0, 1; 0, 2; 5, 0; 6, 0; 7, 0; 8, 0 ] -> valueAt "sheet1" row col calculated ]) [ cycle; cycle; cycle; n 2.0; cycle; cycle; cycle; n 1.0 ]

// ---- OFFSET, INDIRECT ------------------------------------------------------------------
do
    // Precedents chosen while evaluating, reading formulas that come later in every order the references give.
    let wb = three () |> setCells "sheet1" [ for row in 0 .. 5 -> at row 0, Some (number (float (row + 1))) ]
                      |> setCells "sheet1" [ at 0 1, Some (formula "=OFFSET(C1,1,0)"); at 0 2, Some (number 10.0); at 1 2, Some (formula "=C1*2"); at 2 2, Some (formula "=B1+1");
                                             at 0 4, Some (formula "=INDIRECT(\"C\"&F1)"); at 0 5, Some (number 3.0); at 1 4, Some (formula "=SUM(OFFSET(A1,0,0,F1,1))");
                                             at 2 4, Some (formula "=INDIRECT(\"Two!A1\")+1"); at 3 4, Some (formula "=SUM(INDIRECT(\"A1:A\"&F1))"); at 4 4, Some (formula "=INDIRECT(\"TOTAL\")") ]
                      |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!C3*2") ]
    let wb = { wb with Names = wb.Names |> Map.add (None, "TOTAL") "=Sheet1!$E$4" }
    let calculated = settledInOne "OFFSET and INDIRECT" wb
    eq "OFFSET reads a formula; a formula reads the OFFSET" (valueAt "sheet1" 0 1 calculated, valueAt "sheet1" 2 2 calculated) (n 20.0, n 21.0)
    eq "INDIRECT to a cell, a range, another sheet, a name" (valueAt "sheet1" 0 4 calculated, valueAt "sheet1" 3 4 calculated, valueAt "sheet1" 2 4 calculated, valueAt "sheet1" 4 4 calculated) (n 21.0, n 6.0, n 43.0, n 6.0)
    eq "SUM over an OFFSET" (valueAt "sheet1" 1 4 calculated) (n 6.0)
    // Incrementally: every edit, with the values of the full pass.
    let engine = Recalc.build calculated
    ok "the workbook is not left to the full recalculation" (not (Recalc.isUnsupported engine) && Recalc.dynamicCount engine = 6) (string (Recalc.dynamicCount engine))
    let mutable current = calculated
    let step label (change: Wb.Workbook -> Wb.Workbook) =
        let edited = change current
        let expected = full edited
        match Recalc.recalculate engine edited None with
        | Some result ->
            ok (label + ": incremental = full") (same result expected) ""
            current <- result
        | None -> ok (label + ": incremental") false (Recalc.lastDecline engine)
    step "what the OFFSET's target reads" (setCells "sheet1" [ at 0 2, Some (number 11.0) ])
    eq "... followed through the OFFSET to its reader and on to the other sheet" (valueAt "sheet1" 0 1 current, valueAt "sheet1" 2 2 current, valueAt "sheet2" 0 0 current, valueAt "sheet1" 2 4 current) (n 22.0, n 23.0, n 46.0, n 47.0)
    step "the cell INDIRECT builds its reference from" (setCells "sheet1" [ at 0 5, Some (number 2.0) ])
    eq "... another cell, another range" (valueAt "sheet1" 0 4 current, valueAt "sheet1" 3 4 current, valueAt "sheet1" 1 4 current, valueAt "sheet1" 4 4 current) (n 22.0, n 3.0, n 3.0, n 3.0)
    step "a cell only an INDIRECT range reads" (setCells "sheet1" [ at 1 0, Some (number 50.0) ])
    eq "... is read" (valueAt "sheet1" 3 4 current) (n 51.0)
    step "an OFFSET entered" (setCells "sheet1" [ at 8 1, Some (formula "=OFFSET(A1,F1,0)*2") ])
    step "an INDIRECT entered that reads the OFFSET" (setCells "sheet1" [ at 9 1, Some (formula "=INDIRECT(\"B9\")+B9") ])
    eq "... both" (valueAt "sheet1" 8 1 current, valueAt "sheet1" 9 1 current) (n 6.0, n 12.0)
    step "their input" (setCells "sheet1" [ at 2 0, Some (number 30.0) ])
    eq "... both follow" (valueAt "sheet1" 8 1 current, valueAt "sheet1" 9 1 current) (n 60.0, n 120.0)
    step "an OFFSET that spills" (setCells "sheet1" [ at 12 7, Some (formula "=OFFSET(A1,0,0,F1,1)") ])
    eq "... it spilled" (valueAt "sheet1" 12 7 current, valueAt "sheet1" 13 7 current, (resolveSheet current.Sheets.["sheet1"].Model).Spill.Count) (n 1.0, n 50.0, 1)
    step "... and grows" (setCells "sheet1" [ at 0 5, Some (number 4.0) ])
    eq "... to four cells" ([ for row in 12 .. 16 -> valueAt "sheet1" row 7 current ]) [ n 1.0; n 50.0; n 30.0; n 4.0; CellValue.Empty ]
    step "the OFFSET removed" (setCells "sheet1" [ at 0 1, None ])
    step "an edit afterwards" (setCells "sheet1" [ at 0 2, Some (number 1.0) ])

// ---- 3-D references ----------------------------------------------------------------------
do
    let wb = three () |> setCells "sheet1" [ at 0 0, Some (number 1.0); at 1 0, Some (number 2.0); at 0 3, Some (formula "=SUM(Two:Three!A1:A2)"); at 1 3, Some (formula "=SUM(Sheet1:Three!B1)") ; at 2 3, Some (formula "=D1+1") ]
                      |> setCells "sheet2" [ at 0 0, Some (number 10.0); at 1 0, Some (formula "=Three!A2*2"); at 0 1, Some (number 5.0) ]
                      |> setCells "sheet3" [ at 0 0, Some (number 100.0); at 1 0, Some (formula "=Sheet1!A2+1"); at 0 1, Some (formula "=Two!B1+1") ]
    let calculated = settledInOne "3-D references" wb
    eq "a 3-D sum over formulas of two sheets, and its reader" (valueAt "sheet1" 0 3 calculated, valueAt "sheet1" 2 3 calculated) (n 119.0, n 120.0)
    eq "a 3-D reference to one cell on three sheets" (valueAt "sheet1" 1 3 calculated) (n 11.0)
    let engine = Recalc.build calculated
    let mutable current = calculated
    let step label (change: Wb.Workbook -> Wb.Workbook) =
        let edited = change current
        let expected = full edited
        match Recalc.recalculate engine edited None with
        | Some result ->
            ok (label + ": incremental = full") (same result expected) ""
            current <- result
        | None -> ok (label + ": incremental") false (Recalc.lastDecline engine)
    step "a cell of the last sheet of the span" (setCells "sheet3" [ at 0 0, Some (number 200.0) ])
    eq "... is read" (valueAt "sheet1" 0 3 current, valueAt "sheet1" 2 3 current) (n 219.0, n 220.0)
    step "a cell a formula inside the span reads" (setCells "sheet1" [ at 1 0, Some (number 7.0) ])
    eq "... through two formulas on two sheets" (valueAt "sheet1" 0 3 current) (n 234.0)
    step "a cell of the middle sheet" (setCells "sheet2" [ at 0 1, Some (number 6.0) ])
    eq "... both it and what reads it on the next sheet" (valueAt "sheet1" 1 3 current) (n 13.0)
    step "a 3-D reference entered" (setCells "sheet3" [ at 5 5, Some (formula "=AVERAGE(Sheet1:Two!A1)+COUNT(Sheet1:Three!A1:B2)") ])
    eq "... calculated" (valueAt "sheet3" 5 5 current) (n 13.5)
    // The tab order is part of what a 3-D reference reads.
    let moved = { current with SheetOrder = [ "sheet2"; "sheet1"; "sheet3" ] }
    ok "sheets moved: the formulas with 3-D references are calculated, by the engine" (match Recalc.recalculate (Recalc.build current) moved None with Some result -> same result (full moved) | None -> false) ""
    let movedFull = settledInOne "sheets moved" moved
    eq "the span Two:Three holds Sheet1 now, Sheet1:Three no longer Two" (valueAt "sheet1" 0 3 movedFull, valueAt "sheet1" 1 3 movedFull) (n 242.0, n 7.0)
    // A cycle through a 3-D reference.
    let looped = settledInOne "a cycle through a 3-D reference" (current |> setCells "sheet2" [ at 0 1, Some (formula "=Sheet1!D2") ])
    eq "a formula inside the span that reads the 3-D sum, and one between them" (valueAt "sheet1" 1 3 looped, valueAt "sheet2" 0 1 looped, valueAt "sheet3" 0 1 looped) (cycle, cycle, cycle)

// ---- iterative calculation ------------------------------------------------------------
// Excel's setting for workbooks with a deliberate circular reference (a loan whose
// interest is part of what it finances): the formulas of the cycle are calculated
// again and again until they settle, across sheets too; what had settled stays as it is.
do
    let model =
        three () |> setCells "sheet1" [ at 0 0, Some (number 1000.0); at 1 0, Some (number 0.1); at 2 0, Some (formula "=A1+Two!A1"); at 3 0, Some (formula "=A3*2") ]
                 |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!A3*Sheet1!A2"); at 1 0, Some (formula "=A1+1") ]
    let strict = settledInOne "a loan and its interest, no iteration" model
    eq "without the setting: the cycle error" (valueAt "sheet1" 2 0 strict, valueAt "sheet2" 0 0 strict, valueAt "sheet1" 3 0 strict) (cycle, cycle, cycle)
    let iterating = { model with Iteration = Some { MaxIterations = 100; MaxChange = 0.001 } }
    let calculated, engine = Recalc.recalculateAll iterating
    let near (value: CellValue) (expected: float) = match value with CellValue.Number x -> abs (x - expected) < 0.01 | _ -> false
    // total = 1000 + 0.1 * total  ->  1111.11
    ok "with it: the cycle settles at its fixed point" (near (valueAt "sheet1" 2 0 calculated) 1111.111 && near (valueAt "sheet2" 0 0 calculated) 111.111) (sprintf "%A %A" (valueAt "sheet1" 2 0 calculated) (valueAt "sheet2" 0 0 calculated))
    ok "what reads the cycle follows" (near (valueAt "sheet1" 3 0 calculated) 2222.222 && near (valueAt "sheet2" 1 0 calculated) 112.111) ""
    eq "two formulas are on the cycle" (Recalc.cyclicCount engine) 2
    ok "a second run leaves the settled workbook exactly as it is" (same calculated (full calculated)) ""
    ok "switched on in a workbook that showed the cycle error: the formulas start from nothing" (near (valueAt "sheet1" 2 0 (full { strict with Iteration = iterating.Iteration })) 1111.111) ""
    // An input of the cycle edited: the engine settles it again, to the numbers of the full pass.
    let edited = calculated |> setCells "sheet1" [ at 0 0, Some (number 2000.0) ]
    let again, engineAfter = Recalc.recalculateAll edited
    ok "... which settles at the new fixed point" (near (valueAt "sheet1" 2 0 again) 2222.222) (sprintf "%A" (valueAt "sheet1" 2 0 again))
    ok "an edit that reaches the cycle is calculated by the engine, to the same numbers" (match Recalc.recalculate engine edited None with Some result -> same result again | None -> false) (Recalc.lastDecline engine)
    // ... also by an engine that was built for the workbook and never ran a full pass over it.
    ok "... and by an engine built for the calculated workbook" (match Recalc.recalculate (Recalc.build calculated) edited None with Some result -> same result again | None -> false) ""
    // An edit beside the cycle is incremental and leaves the cycle's values alone.
    match Recalc.recalculate engineAfter (again |> setCells "sheet3" [ at 0 0, Some (number 5.0) ]) None with
    | Some result -> ok "an edit beside the cycle is incremental" (valueAt "sheet1" 2 0 result = valueAt "sheet1" 2 0 again && same result (full result)) ""
    | None -> ok "an edit beside the cycle is incremental" false (Recalc.lastDecline engineAfter)
    // A cycle that does not settle stops at the workbook's number of iterations.
    let growing = { (three () |> setCells "sheet1" [ at 0 0, Some (formula "=Two!A1+1") ] |> setCells "sheet2" [ at 0 0, Some (formula "=Sheet1!A1") ]) with Iteration = Some { MaxIterations = 10; MaxChange = 0.001 } }
    eq "a cycle that grows: as many steps as the workbook allows" (valueAt "sheet1" 0 0 (full growing)) (n 10.0)
    eq "... and ten more at the next run, as in Excel" (valueAt "sheet1" 0 0 (full (full growing))) (n 20.0)
    // ... at an edit elsewhere too: Excel calculates a circular reference at every calculation.
    let grown, growingEngine = Recalc.recalculateAll growing
    let beside = grown |> setCells "sheet3" [ at 4 4, Some (number 1.0) ]
    ok "an edit beside a cycle that grows: ten more, as the full pass gives" (match Recalc.recalculate growingEngine beside None with Some result -> valueAt "sheet1" 0 0 result = n 20.0 && same result (full beside) | None -> false) ""
    // An engine that was only built for the workbook takes its circular references as settled (the workbook
    // came calculated) until an edit reaches one - or until it has looked for them.
    ok "... an engine that was only built leaves the cycle as the workbook has it" (match Recalc.recalculate (Recalc.build grown) beside None with Some result -> valueAt "sheet1" 0 0 result = n 10.0 | None -> false) ""
    let looked = Recalc.build grown
    Recalc.findCycles looked
    ok "... and gives ten more once it has looked for the circular references" (match Recalc.recalculate looked beside None with Some result -> valueAt "sheet1" 0 0 result = n 20.0 && same result (full beside) | None -> false) ""
    let reaching = grown |> setCells "sheet2" [ at 5 5, Some (formula "=Sheet1!A1*2") ] |> setCells "sheet1" [ at 0 0, Some (formula "=Two!A1+2") ]
    ok "... an edit that reaches the cycle calculates it, whatever the engine knew" (match Recalc.recalculate (Recalc.build grown) reaching None with Some result -> same result (full reaching) | None -> false) ""

printfn "stage45-order: %d passed, %d failed" pass fail
if fail > 0 then exit 1
