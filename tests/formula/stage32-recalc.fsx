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
        let spillA = (resolveSheet left.Sheets.[sheetId].Model).Spill
        let spillB = (resolveSheet right.Sheets.[sheetId].Model).Spill
        if a = b then
            // The same cells spilled from the same formulas.
            if spillA = spillB then None
            else
                let keys = Seq.append (spillA |> Seq.map (fun kv -> kv.Key)) (spillB |> Seq.map (fun kv -> kv.Key)) |> Seq.distinct
                keys |> Seq.tryPick (fun key ->
                    let x, y = Map.tryFind key spillA, Map.tryFind key spillB
                    if x <> y then Some (sprintf "%s!R%dC%d spilled from %A vs %A" sheetId key.Row key.Col x y) else None)
        else
            let keys = Seq.append (a |> Seq.map (fun kv -> kv.Key)) (b |> Seq.map (fun kv -> kv.Key)) |> Seq.distinct
            keys |> Seq.tryPick (fun key ->
                let x, y = Map.tryFind key a |> Option.map (fun c -> c.Value), Map.tryFind key b |> Option.map (fun c -> c.Value)
                if x <> y then Some (sprintf "%s!R%dC%d: %A vs %A" sheetId key.Row key.Col x y) else None))

/// The reference the two recalculations are checked against: the full pass as
/// it was before the workbook had one dependency order (every sheet on its
/// own, in tab order, a few times), run until it gives what it was given -
/// one run of it leaves a long chain through several sheets short of its
/// values. It shares no ordering code with the engine.
let settledReference (wb: Wb.Workbook) =
    let mutable result = Formula.recalculateWorkbookSettling wb
    let mutable runs = 1
    let mutable settled = false
    while not settled && runs < 12 do
        let again = Formula.recalculateWorkbookSettling result
        settled <- (difference again result).IsNone
        result <- again
        runs <- runs + 1
    result, settled

/// Whether a formula of the workbook shows the cycle error (the reference marks
/// circular references within a sheet only, and not every formula on one).
let hasCycle (wb: Wb.Workbook) =
    wb.Sheets |> Map.exists (fun _ sheet ->
        (resolveSheet sheet.Model).Cells |> Map.exists (fun _ cell -> unwrapFormulaValue cell.Value = CellValue.Error Formula.Err.cycle))

/// The full recalculation: ONE run of the ordered pass. It is compared with
/// the settled reference wherever that settles and nothing is circular.
let mutable fullRuns = 0
let mutable fullCompared = 0
let mutable fullShort = 0
/// Where the last full recalculation differed from the reference, until `countFull` has counted it.
let mutable fullDifference : string option = None
/// The workbook holds arrays whose layout does not settle (a random sequence ends there).
let mutable unsettled = false
let countFull (label: string) =
    match fullDifference with
    | Some detail ->
        fullShort <- fullShort + 1
        if fullShort <= 5 then printfn "FAIL %s: one run of the full recalculation differs from the settled reference at %s" label detail
        fullDifference <- None
    | None -> ()
let fullRecalculation (wb: Wb.Workbook) =
    let result, fullEngine = Recalc.recalculateAll wb
    fullRuns <- fullRuns + 1
    let reference, settled = settledReference wb
    fullDifference <- None
    // (Formulas on a circular reference show the cycle error or - in a workbook with iterative
    // calculation - are iterated: the reference does neither across sheets.)
    if settled && Recalc.cyclicCount fullEngine = 0 && not (hasCycle result) && not (hasCycle reference) then
        fullCompared <- fullCompared + 1
        fullDifference <- difference result reference
    result

/// The formula index of a sheet's cells (`Model.formulaIndexOf`), wherever one is known, names exactly
/// the cells that hold formulas: it follows setCell / clearCell and both recalculations.
let mutable indexChecks = 0
let mutable indexKnown = 0
let checkFormulaIndex (label: string) (wb: Wb.Workbook) =
    for KeyValue (sheetId, sheet) in wb.Sheets do
        match sheet.Model.Deferred with
        | Some deferred when deferred.Loaded.IsNone -> ()
        | _ ->
            let cells = (resolveSheet sheet.Model).Cells
            indexChecks <- indexChecks + 1
            match tryFormulaIndex cells with
            | Some index ->
                indexKnown <- indexKnown + 1
                let walked = cells |> Map.toSeq |> Seq.filter (fun (_, cell) -> match cell.Value with CellValue.Formula _ -> true | _ -> false) |> Seq.map fst |> Set.ofSeq
                if index <> walked then
                    fail <- fail + 1
                    printfn "FAIL %s: the formula index of %s names %d cells, the sheet has %d formulas (%A)" label sheetId index.Count walked.Count ((index - walked) + (walked - index) |> Seq.truncate 3 |> List.ofSeq)
            | None -> ()

/// An edit both ways: incrementally, and by a full recalculation.
type Outcome = Incremental of Recalc.RecalcStats | Declined
/// How often the engine declined, by the reason it gave.
let declines = System.Collections.Generic.Dictionary<string, int>()
let mutable referenceSteps = 0
let mutable referenceShort = 0
let mutable engine = Unchecked.defaultof<Recalc.Engine>
let mutable current = Unchecked.defaultof<Wb.Workbook>
let start (wb: Wb.Workbook) =
    current <- fullRecalculation wb
    countFull "start"
    checkFormulaIndex "start" current
    engine <- Recalc.build current
    // (As the full recalculation, which `edit` compares with: every circular reference is calculated at every edit.)
    Recalc.findCycles engine
let edit label (hint: (string * CellAddress) list option) (change: Wb.Workbook -> Wb.Workbook) =
    let edited = change current
    checkFormulaIndex (label + " (edited)") edited
    let full = fullRecalculation edited
    checkFormulaIndex (label + " (full)") full
    // (How often one run of the reference - the pass as it was - is short of the values.)
    referenceSteps <- referenceSteps + 1
    if (difference (Formula.recalculateWorkbookSettling edited) full).IsSome then referenceShort <- referenceShort + 1
    match Recalc.recalculate engine edited hint with
    | Some result ->
        countFull label
        checkFormulaIndex (label + " (incremental)") result
        (match difference result full with
         | None -> pass <- pass + 1
         | Some detail -> fail <- fail + 1; printfn "FAIL %s: incremental differs from full at %s" label detail)
        current <- full
        // (The engine goes on from its own result: any difference would show at a later step.)
        Incremental (Recalc.lastStats engine)
    | None ->
        current <- full
        let reason = Recalc.lastDecline engine
        // (A circular reference through the cells of dynamic arrays - one array blocks the other
        // and so frees itself - has no settled layout: the two passes stop at different ones.)
        if reason.Contains "do not settle" then (fullDifference <- None; unsettled <- true) else countFull label
        declines.[reason] <- (match declines.TryGetValue reason with | true, n -> n + 1 | _ -> 1)
        if not (Recalc.adopt engine full) then engine <- Recalc.build full
        Recalc.findCycles engine
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

// ---- circular references, and edits that are more than cells --------------------
// (They were left to the full recalculation; `edit` checks each against it.)
let handled label outcome = ok label (outcome <> Declined) (sprintf "left to the full recalculation: %s" (Recalc.lastDecline engine))
handled "a formula that reads itself" (edit "cycle" None (setCells "sheet1" [ at 30 3, Some (formula "=D31+1") ]))
eq "the cycle is marked" (valueAt "sheet1" 30 3 current) (CellValue.Error Formula.Err.cycle)
ok "an edit elsewhere is incremental although a cycle exists" (edit "beside a cycle" None (setCells "sheet1" [ at 0 0, Some (number 1.0) ]) <> Declined) ""
handled "a cycle through two cells" (edit "cycle 2" None (setCells "sheet1" [ at 31 3, Some (formula "=D33"); at 32 3, Some (formula "=D32") ]))
eq "both show it" (valueAt "sheet1" 31 3 current, valueAt "sheet1" 32 3 current) (CellValue.Error Formula.Err.cycle, CellValue.Error Formula.Err.cycle)
handled "what reads a cycle" (edit "reader of a cycle" None (setCells "sheet1" [ at 33 3, Some (formula "=D32*2"); at 34 3, Some (formula "=D34+1") ]))
eq "... shows the error too" (valueAt "sheet1" 33 3 current) (CellValue.Error Formula.Err.cycle)
// A formula that was there before the edit closes the cycle with a formula the edit brings.
handled "an existing formula drawn into a cycle" (edit "cycle 3" (Some [ "sheet1", at 36 3 ]) (setCells "sheet1" [ at 36 3, Some (formula "=D36+1") ]))
handled "... its other half entered" (edit "cycle 3b" (Some [ "sheet1", at 35 3 ]) (setCells "sheet1" [ at 35 3, Some (formula "=D37*2") ]))
eq "both halves show it" (valueAt "sheet1" 35 3 current, valueAt "sheet1" 36 3 current) (CellValue.Error Formula.Err.cycle, CellValue.Error Formula.Err.cycle)
handled "a cycle broken by a value" (edit "cycle 3 broken" None (setCells "sheet1" [ at 35 3, Some (number 4.0) ]))
eq "the other half has its value again" (valueAt "sheet1" 36 3 current) (CellValue.Number 5.0)
handled "the cycles removed" (edit "the cycles removed" None (setCells "sheet1" [ at 30 3, None; at 31 3, None; at 32 3, None; at 33 3, None; at 34 3, None; at 35 3, None; at 36 3, None ]))
ok "then edits are incremental again" (edit "after the cycles" None (setCells "sheet1" [ at 0 0, Some (number 2.0) ]) <> Declined) ""
// (Dynamic arrays, OFFSET / INDIRECT and 3-D references are calculated incrementally: see further down.)
// More than cells: a defined name.
handled "a defined name added" (edit "name" None (fun wb -> { wb with Names = wb.Names |> Map.add (None, "RATE") "=Sheet1!$A$1" }))
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
                let kind = random.Next 14
                // (Rows and columns inserted and deleted among the edits of cells: the formulas of a real workbook follow.)
                let structureEdit =
                    if kind < 12 || cells.Length = 0 then None
                    else
                        let around = anyCell ()
                        Some (match random.Next 5 with
                              | 0 | 1 -> InsertRows (around.Row, 1 + random.Next 2)
                              | 2 -> DeleteRows (around.Row + 1, 1)
                              | 3 -> InsertCols (around.Col, 1)
                              | _ -> DeleteCols (around.Col + 1 + random.Next 3, 1))
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
                let outcome =
                    match structureEdit with
                    | Some structure -> edit (sprintf "%s step %d (%A)" label step structure) None (fun wb -> fst (Wb.applyStructureEdit sheetId structure wb))
                    | None -> edit (sprintf "%s step %d (%s)" label step name) hint (setCells targetSheet changes)
                match outcome with
                | Incremental s ->
                    incremental <- incremental + 1
                    evaluated <- evaluated + s.EvaluatedFormulas
                | Declined -> declinedSteps <- declinedSteps + 1
        printfn "  %s: %d formulas; %d edits incremental (%d formulas evaluated in all), %d left to the full recalculation%s"
            label (Recalc.formulaCount engine) incremental evaluated declinedSteps (if Recalc.isUnsupported engine then " (unsupported: OFFSET / INDIRECT / 3-D / spill-range references)" else "")
        incremental, declinedSteps

let mutable incrementalTotal = 0
for file, seed, steps in List.filter (fun _ -> Environment.GetEnvironmentVariable "SKIP_FILES" <> "1") [
                           "real-world/milk-street-office.xlsx", 3, 45
                           "real-world/lyric-multifamily.xlsx", 4, 35
                           "real-world/dcf-stock-analysis.xlsx", 5, 45
                           "real-world/saopaulo-condos.xlsx", 6, 35
                           "real-world/jumeirah-hotel.xlsx", 7, 35
                           "real-world/dashboard-sales-analysis.xlsx", 8, 30 ] do
    let incremental, _ = randomEdits (Path.GetFileNameWithoutExtension file) seed steps (load file)
    incrementalTotal <- incrementalTotal + incremental
// (The synthetic workbook again, so that the sequences cover one the engine supports whatever the files hold.)
let incrementalSynthetic, _ = randomEdits "two sheets" 9 150 (fst (twoSheets ()))
ok "the random sequences were mostly incremental" (incrementalSynthetic > 120 && (incrementalTotal > 60 || Environment.GetEnvironmentVariable "SKIP_FILES" = "1")) (sprintf "%d synthetic, %d on files" incrementalSynthetic incrementalTotal)

// One run of the full recalculation reaches the values, whatever the edit: it has one dependency
// order for the whole workbook. (One run of the pass as it was - sheet by sheet - was short in some
// steps of the sequences on the schedule workbooks.)
let reportFullRuns (what: string) =
    eq (sprintf "one run of the full recalculation gives the settled values (%s)" what) fullShort 0
    printfn "  %s: %d full recalculations, %d compared with the settled reference, %d differ; one run of the reference was short in %d of %d steps"
        what fullRuns fullCompared fullShort referenceShort referenceSteps
reportFullRuns "directed cases and random edits of real workbooks"
ok "the comparison with the reference took place" (fullCompared > 100 || Environment.GetEnvironmentVariable "SKIP_FILES" = "1") (string fullCompared)

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
    // A circular reference through the other sheet is calculated by the engine: every formula on it shows it.
    ok "the input overwritten by a formula that reads what depends on it: incremental" (edit "a cycle across sheets" None (setCells "sheet1" [ at 0 0, Some (formula "=A2+1") ]) <> Declined) (Recalc.lastDecline engine)
    eq "... all three formulas of the cycle" (valueAt "sheet1" 0 0 current, valueAt "sheet1" 1 0 current, valueAt other 0 0 current) (CellValue.Error Formula.Err.cycle, CellValue.Error Formula.Err.cycle, CellValue.Error Formula.Err.cycle)
    ok "broken again" (edit "the cycle across sheets broken" None (setCells "sheet1" [ at 0 0, Some (number 5.0) ]) <> Declined) (Recalc.lastDecline engine)
    eq "... the values are back" (valueAt "sheet1" 1 0 current) (CellValue.Number 12.0)

// ---- dynamic arrays ---------------------------------------------------------------
// A formula that spills owns the cells it spilled into; entering, changing and
// removing one, a change of what it reads, and typing into its cells are
// calculated incrementally, with the cells and the spill map of the full pass.
let incrementally label outcome = ok label (outcome <> Declined) (sprintf "declined: %s" (Recalc.lastDecline engine))
let spillOf sheetId (wb: Wb.Workbook) = (resolveSheet wb.Sheets.[sheetId].Model).Spill
do
    let wb =
        Wb.createWorkbook ()
        |> setCells "sheet1" [ for row in 0 .. 9 -> at row 0, Some (number (float (10 - row))) ]
        |> setCells "sheet1" [ at 0 11, Some (number 3.0); at 1 11, Some (number 4.0) ]
    start wb
    declines.Clear()
    incrementally "a spilling formula entered" (edit "spill entered" (Some [ "sheet1", at 0 3 ]) (setCells "sheet1" [ at 0 3, Some (formula "=SEQUENCE($L$1)") ]))
    eq "it spilled" ([ for row in 0 .. 3 -> valueAt "sheet1" row 3 current ], spillOf "sheet1" current |> Map.toList)
        ([ CellValue.Number 1.0; CellValue.Number 2.0; CellValue.Number 3.0; CellValue.Empty ], [ at 1 3, at 0 3; at 2 3, at 0 3 ])
    // Readers of the spilled cells: a cell, a range over them, the array as a whole, and a chain off one.
    incrementally "readers entered" (edit "readers" None (setCells "sheet1" [ at 0 12, Some (formula "=D2*10"); at 1 12, Some (formula "=SUM(D1:D9)"); at 2 12, Some (formula "=SUM(D1#)"); at 3 12, Some (formula "=M1+M2"); at 4 12, Some (formula "=COUNT(D:D)") ]))
    eq "they read the array" ([ for row in 0 .. 4 -> valueAt "sheet1" row 12 current ]) [ CellValue.Number 20.0; CellValue.Number 6.0; CellValue.Number 6.0; CellValue.Number 26.0; CellValue.Number 3.0 ]
    let grown = edit "the array grows" (Some [ "sheet1", at 0 11 ]) (setCells "sheet1" [ at 0 11, Some (number 5.0) ])
    incrementally "its size read from a cell: it grows" grown
    eq "... and what reads the cells it newly covers follows" ([ for row in 1 .. 4 -> valueAt "sheet1" row 12 current ], valueAt "sheet1" 4 3 current, (spillOf "sheet1" current).Count)
        ([ CellValue.Number 15.0; CellValue.Number 15.0; CellValue.Number 35.0; CellValue.Number 5.0 ], CellValue.Number 5.0, 4)
    incrementally "it shrinks" (edit "the array shrinks" None (setCells "sheet1" [ at 0 11, Some (number 2.0) ]))
    eq "... its cells are cleared" ([ for row in 0 .. 4 -> valueAt "sheet1" row 3 current ], valueAt "sheet1" 1 12 current, (spillOf "sheet1" current).Count)
        ([ CellValue.Number 1.0; CellValue.Number 2.0; CellValue.Empty; CellValue.Empty; CellValue.Empty ], CellValue.Number 3.0, 1)
    incrementally "it becomes one cell" (edit "the array is one cell" None (setCells "sheet1" [ at 0 11, Some (number 1.0) ]))
    eq "... no spilled cells" (valueAt "sheet1" 1 3 current, spillOf "sheet1" current |> Map.isEmpty, valueAt "sheet1" 0 12 current) (CellValue.Empty, true, CellValue.Number 0.0)
    edit "back to four" None (setCells "sheet1" [ at 0 11, Some (number 4.0) ]) |> ignore
    // Typing into the array blocks it; clearing the cell lets it spill again.
    incrementally "a spilled cell typed over" (edit "into the spill" (Some [ "sheet1", at 2 3 ]) (setCells "sheet1" [ at 2 3, Some (number 99.0) ]))
    eq "the formula is blocked and its other cells are cleared" ([ for row in 0 .. 3 -> valueAt "sheet1" row 3 current ], spillOf "sheet1" current |> Map.isEmpty, valueAt "sheet1" 1 12 current)
        ([ CellValue.Error Formula.Err.spill; CellValue.Empty; CellValue.Number 99.0; CellValue.Empty ], true, CellValue.Error Formula.Err.spill)
    incrementally "an edit elsewhere while it is blocked" (edit "beside the blocked spill" None (setCells "sheet1" [ at 5 0, Some (number 1.0) ]))
    incrementally "the blocking cell cleared" (edit "unblock" (Some [ "sheet1", at 2 3 ]) (setCells "sheet1" [ at 2 3, None ]))
    eq "the array is back" ([ for row in 0 .. 3 -> valueAt "sheet1" row 3 current ], valueAt "sheet1" 1 12 current) ([ for n in 1 .. 4 -> CellValue.Number (float n) ], CellValue.Number 10.0)
    // A spilled cell cleared (Delete on it) is written again; one restyled is the array's again.
    incrementally "a spilled cell cleared" (edit "ghost cleared" None (setCells "sheet1" [ at 1 3, None ]))
    eq "it is back" (valueAt "sheet1" 1 3 current) (CellValue.Number 2.0)
    incrementally "a spilled cell restyled" (edit "ghost restyled" None (fun wb -> Wb.updateSheetModel "sheet1" (setCellBgColor (at 1 3) "#ff0000" wb.Sheets.["sheet1"].Model) wb))
    eq "it has the formula's format, as after a full recalculation" ((getCell (at 1 3) (resolveSheet current.Sheets.["sheet1"].Model)).Value.Format.BgColor) None
    // Arrays over edited ranges.
    incrementally "FILTER, SORT, UNIQUE, array arithmetic, TRANSPOSE entered" (edit "arrays" None (setCells "sheet1"
        [ at 12 3, Some (formula "=FILTER(A1:A10,A1:A10>$L$2)"); at 12 4, Some (formula "=SORT(A1:A6)"); at 12 5, Some (formula "=UNIQUE(A1:A10)")
          at 12 6, Some (formula "=A1:A5*2"); at 24 3, Some (formula "=TRANSPOSE(A1:A4)"); at 26 3, Some (formula "=E13#+1"); at 5 12, Some (formula "=SUM(D13:G23)") ]))
    eq "they are there" (valueAt "sheet1" 12 3 current, valueAt "sheet1" 12 4 current, valueAt "sheet1" 24 6 current, valueAt "sheet1" 26 3 current) (CellValue.Number 10.0, CellValue.Number 1.0, CellValue.Number 7.0, CellValue.Number 2.0)
    incrementally "an input of all of them changed" (edit "input" (Some [ "sheet1", at 0 0 ]) (setCells "sheet1" [ at 0 0, Some (number 0.5) ]))
    incrementally "the filter's threshold changed: the array is shorter" (edit "threshold" None (setCells "sheet1" [ at 1 11, Some (number 7.0) ]))
    incrementally "a text among the inputs" (edit "text" None (setCells "sheet1" [ at 3 0, Some (text "x") ]))
    incrementally "inputs cleared" (edit "cleared" None (setCells "sheet1" [ at 6 0, None; at 7 0, None ]))
    // Two arrays that want the same cells: the one further up / left has them.
    incrementally "an array entered across another" (edit "across" None (setCells "sheet1" [ at 14 2, Some (formula "=SEQUENCE(1,6)") ]))
    eq "the later formula is blocked" (valueAt "sheet1" 14 2 current) (CellValue.Error Formula.Err.spill)
    // An array further up / left takes the cells it wants from one further down / right, which is blocked then.
    incrementally "a wide array entered" (edit "wide" None (setCells "sheet1" [ at 40 3, Some (formula "=SEQUENCE(3,3)") ]))
    incrementally "an array entered above that takes one of its cells" (edit "above" None (setCells "sheet1" [ at 39 4, Some (formula "=SEQUENCE(2)") ]))
    eq "the array below is blocked and its cells are cleared" (valueAt "sheet1" 40 3 current, valueAt "sheet1" 40 4 current, valueAt "sheet1" 41 4 current, valueAt "sheet1" 39 4 current) (CellValue.Error Formula.Err.spill, CellValue.Number 2.0, CellValue.Empty, CellValue.Number 1.0)
    incrementally "the array above removed" (edit "above removed" (Some [ "sheet1", at 39 4 ]) (setCells "sheet1" [ at 39 4, None ]))
    eq "the array below spills again" (valueAt "sheet1" 40 3 current, valueAt "sheet1" 42 5 current) (CellValue.Number 1.0, CellValue.Number 9.0)
    incrementally "a spilling formula changed to a plain one" (edit "to plain" None (setCells "sheet1" [ at 12 4, Some (formula "=A1+1") ]))
    eq "its cells are gone" (valueAt "sheet1" 13 4 current) (CellValue.Empty)
    incrementally "a spilling formula removed" (edit "removed" None (setCells "sheet1" [ at 12 5, None ]))
    // An undo that puts the cells and the spill map back as they were.
    let before = current
    incrementally "arrays deleted in a block" (edit "block deleted" None (setCells "sheet1" [ for row in 0 .. 27 do for col in 2 .. 7 -> at row col, None ]))
    incrementally "the undo: cells and spill map put back" (edit "undo of the block" None (fun wb ->
        let old = resolveSheet before.Sheets.["sheet1"].Model
        Wb.updateSheetModel "sheet1" ({ wb.Sheets.["sheet1"].Model with Cells = old.Cells; Spill = old.Spill } |> invalidateFilterProjection) wb))
    ok "the workbook is what it was" ((difference current before).IsNone) (sprintf "%A" (difference current before))
    eq "none of these edits was left to the full recalculation" (declines |> Seq.map (fun kv -> kv.Key, kv.Value) |> List.ofSeq) []
    // A formula that reads its own spilled cells is a circular reference.
    edit "an array that reads what it spills into" None (setCells "sheet1" [ at 0 8, Some (formula "=SEQUENCE(3)+I3") ]) |> ignore
    ok "... is declined as circular, not as a spill" (declines.Keys |> Seq.forall (fun reason -> reason.Contains "circular")) (sprintf "%A" (List.ofSeq declines.Keys))

// Random sequences with dynamic arrays: sizes read from cells, arrays over
// edited ranges, arrays side by side that block one another, typing into and
// clearing spilled cells, removed formulas, readers of spilled cells, undos.
// With `dynamic`, formulas that choose their precedents while they are
// evaluated (OFFSET, INDIRECT - to cells, ranges over the arrays, the other
// sheet, a name; one that spills) and 3-D references are mixed in.
let randomSpills (dynamic: bool) (seed: int) (steps: int) =
    let random = Random seed
    let wb =
        Wb.createWorkbook () |> Wb.addSheet "Other"
    let other = wb.SheetOrder |> List.last
    let wb =
        wb
        |> setCells "sheet1" [ for row in 0 .. 11 do
                                  yield at row 0, Some (number (float (random.Next 20)))
                                  yield at row 1, Some (if random.Next 3 = 0 then text (string (char (65 + random.Next 4))) else number (float (random.Next 4))) ]
        |> setCells "sheet1" [ at 0 14, Some (number 3.0); at 1 14, Some (number 5.0); at 2 14, Some (number 2.0) ]
        |> setCells "sheet1" [ at 4 14, Some (formula "=SUM(D1:K30)"); at 5 14, Some (formula "=COUNT(E:E)"); at 6 14, Some (formula "=F5*2+G7"); at 7 14, Some (formula "=O5+O6") ]
        |> setCells other [ at 0 0, Some (formula "=SUM(Sheet1!D1:K30)"); at 1 0, Some (formula "=Sheet1!E4&\"\""); at 3 0, Some (formula "=Sheet1!A1:A3*2") ]
    let wb = if dynamic then { wb with Names = wb.Names |> Map.add (None, "SIZE") "=Sheet1!$O$1" } |> setCells other [ at 0 2, Some (number 4.0); at 1 2, Some (formula "=Sheet1!O5+1") ] else wb
    start wb
    let snapshots = ResizeArray<SpreadsheetModel>()
    let failuresBefore = fail
    let mutable incremental = 0
    unsettled <- false
    for step in 1 .. steps do
        if fail = failuresBefore && not unsettled then
            let model = resolveSheet current.Sheets.["sheet1"].Model
            let anchors = model.Spill |> Map.toList |> List.map snd |> List.distinct |> List.toArray
            let ghosts = model.Spill |> Map.toArray |> Array.map fst
            let spot () = at (random.Next 26) (3 + random.Next 8)
            let refOf (addr: CellAddress) = sprintf "%s%d" (columnToLetter addr.Col) (addr.Row + 1)
            let arrayFormula () =
                match random.Next 12 with
                | 0 -> "=SEQUENCE($O$1)"
                | 1 -> "=SEQUENCE($O$3,$O$1)"
                | 2 -> "=FILTER(A1:A12,A1:A12>$O$2)"
                | 3 -> "=SORT(A1:A8)"
                | 4 -> "=UNIQUE(B1:B12)"
                | 5 -> "=A1:A5*2"
                | 6 -> "=TRANSPOSE(A1:A4)"
                | 7 -> "=SEQUENCE(2,3)"
                | 8 when anchors.Length > 0 -> sprintf "=%s#*2" (refOf anchors.[random.Next anchors.Length])
                | 9 when anchors.Length > 0 -> sprintf "=SORT(%s#,1,-1)" (refOf anchors.[random.Next anchors.Length])
                | 10 -> "=FILTER(A1:B12,B1:B12=$O$3)"
                | _ -> "=A1:A3+TRANSPOSE(A1:A3)"
            let dynamicReader () =
                match random.Next 12 with
                | 0 -> "=OFFSET(A1,$O$3,0)"
                | 1 -> "=SUM(OFFSET(A1,0,0,$O$1,2))"
                | 2 -> "=INDIRECT(\"A\"&$O$1)+1"
                | 3 -> "=SUM(INDIRECT(\"D1:F\"&($O$2+3)))"
                | 4 -> "=INDIRECT(\"Other!C2\")"
                | 5 -> "=INDIRECT(\"SIZE\")*2"
                | 6 -> "=SUM(Sheet1:Other!A1:A3)"
                | 7 -> "=COUNT(Sheet1:Other!C1:C2)+MAX(Sheet1:Other!C1)"
                | 8 -> sprintf "=OFFSET(D1,%d,%d)" (random.Next 12) (random.Next 6)
                | 9 -> sprintf "=SUM(OFFSET(%s,0,0,$O$3,$O$1))" (refOf (spot ()))
                | 10 -> "=COUNT(INDIRECT(\"E:E\"))"
                | _ -> "=IFERROR(INDEX(D1:K30,$O$1,$O$3),0)+SUM(D1:INDEX(D1:K30,$O$3,2))"
            let name, change =
                match random.Next (if dynamic then 20 else 16) with
                | 16 | 17 -> "a formula that chooses its precedents, or a 3-D reference, entered", setCells "sheet1" [ at (8 + random.Next 12) (14 + random.Next 3), Some (formula (dynamicReader ())) ]
                | 18 -> "an OFFSET that spills entered", setCells "sheet1" [ spot (), Some (formula (if random.Next 2 = 0 then "=OFFSET(A1,0,0,$O$3,1)" else "=OFFSET(A1,1,0,$O$1,2)")) ]
                | 19 -> "the other sheet edited", setCells other [ at (random.Next 2) 2, Some (number (float (random.Next 9))) ]
                | 0 | 1 -> "an input changed", setCells "sheet1" [ at (random.Next 12) (random.Next 2), Some (number (float (random.Next 20))) ]
                | 2 -> "a size changed", setCells "sheet1" [ at (random.Next 3) 14, Some (number (float (1 + random.Next 5))) ]
                | 3 | 4 | 5 -> "an array formula entered", setCells "sheet1" [ spot (), Some (formula (arrayFormula ())) ]
                | 6 -> "a number typed in the arrays' area", setCells "sheet1" [ spot (), Some (number (float (random.Next 9))) ]
                | 7 when ghosts.Length > 0 -> "a spilled cell typed over", setCells "sheet1" [ ghosts.[random.Next ghosts.Length], Some (number 77.0) ]
                | 8 when ghosts.Length > 0 -> "a spilled cell cleared", setCells "sheet1" [ ghosts.[random.Next ghosts.Length], None ]
                | 9 when anchors.Length > 0 -> "an array formula removed", setCells "sheet1" [ anchors.[random.Next anchors.Length], None ]
                | 10 ->
                    let origin = spot ()
                    "a block cleared", setCells "sheet1" [ for r in 0 .. random.Next 5 do for c in 0 .. random.Next 3 -> at (origin.Row + r) (origin.Col + c), None ]
                | 11 ->
                    let target = spot ()
                    "a reader entered", setCells "sheet1" [ at (8 + random.Next 12) (14 + random.Next 2), Some (formula (if random.Next 2 = 0 then sprintf "=%s+1" (refOf target) else sprintf "=SUM(%s:%s)" (refOf target) (refOf (at (target.Row + random.Next 6) (target.Col + random.Next 3))))) ]
                | 12 when anchors.Length > 0 ->
                    "a reader of a whole array entered", setCells "sheet1" [ at (8 + random.Next 12) 16, Some (formula (sprintf "=SUM(%s#)" (refOf anchors.[random.Next anchors.Length]))) ]
                | 13 when snapshots.Count > 0 ->
                    let old = snapshots.[random.Next snapshots.Count]
                    "cells and spill map put back", (fun wb -> Wb.updateSheetModel "sheet1" ({ wb.Sheets.["sheet1"].Model with Cells = old.Cells; Spill = old.Spill } |> invalidateFilterProjection) wb)
                | 14 -> "a text among the inputs", setCells "sheet1" [ at (random.Next 12) (random.Next 2), Some (text (string (char (65 + random.Next 4)))) ]
                | _ -> "an input cleared", setCells "sheet1" [ at (random.Next 12) (random.Next 2), None ]
            if step % 7 = 0 then snapshots.Add model
            // (A single changed cell is named to the engine half of the time, as a typed entry is.)
            let edited = change current
            let hint =
                let was, now = model.Cells, (resolveSheet edited.Sheets.["sheet1"].Model).Cells
                let keys = Seq.append (was |> Seq.map (fun kv -> kv.Key)) (now |> Seq.map (fun kv -> kv.Key)) |> Seq.distinct |> Seq.filter (fun key -> Map.tryFind key was <> Map.tryFind key now) |> Seq.truncate 2 |> List.ofSeq
                match keys with
                | [ only ] when random.Next 2 = 0 -> Some [ "sheet1", only ]
                | _ -> None
            match edit (sprintf "spills seed %d step %d (%s)" seed step name) hint (fun _ -> edited) with
            | Incremental _ -> incremental <- incremental + 1
            | Declined -> ()
    incremental

declines.Clear()
let mutable spillIncremental = 0
let spillSeeds, spillSteps = (match Environment.GetEnvironmentVariable "SPILL_SEEDS" with null | "" -> 60 | text -> int text), 70
for seed in 1 .. spillSeeds do
    spillIncremental <- spillIncremental + randomSpills false (100 + seed) spillSteps
printfn "  dynamic arrays: %d of %d random edits incremental; left to the full recalculation: %s" spillIncremental (spillSeeds * spillSteps)
    (if declines.Count = 0 then "none" else declines |> Seq.map (fun kv -> sprintf "%d x \"%s\"" kv.Value kv.Key) |> String.concat ", ")
ok "no edit of a dynamic array is declined for being one (only circular references are)"
    (declines.Keys |> Seq.forall (fun reason -> reason.Contains "circular"))
    (sprintf "%A" (List.ofSeq declines.Keys))
ok "the random edits of dynamic arrays were nearly all incremental" (spillIncremental * 10 >= spillSeeds * spillSteps * 9) (sprintf "%d of %d" spillIncremental (spillSeeds * spillSteps))

reportFullRuns "with dynamic arrays"

// OFFSET / INDIRECT / 3-D references among them: such workbooks were left to the full recalculation.
declines.Clear()
let mutable dynamicIncremental = 0
let dynamicSeeds = (match Environment.GetEnvironmentVariable "DYNAMIC_SEEDS" with null | "" -> 40 | text -> int text)
for seed in 1 .. dynamicSeeds do
    dynamicIncremental <- dynamicIncremental + randomSpills true (500 + seed) spillSteps
printfn "  OFFSET / INDIRECT / 3-D with dynamic arrays: %d of %d random edits incremental; left to the full recalculation: %s" dynamicIncremental (dynamicSeeds * spillSteps)
    (if declines.Count = 0 then "none" else declines |> Seq.map (fun kv -> sprintf "%d x \"%s\"" kv.Value kv.Key) |> String.concat ", ")
ok "no edit is declined for OFFSET / INDIRECT / 3-D references (only circular references are)"
    (declines.Keys |> Seq.forall (fun reason -> reason.Contains "circular"))
    (sprintf "%A" (List.ofSeq declines.Keys))
ok "the random edits with OFFSET / INDIRECT / 3-D references were nearly all incremental" (dynamicIncremental * 10 > dynamicSeeds * spillSteps * 9) (sprintf "%d of %d" dynamicIncremental (dynamicSeeds * spillSteps))
reportFullRuns "with OFFSET / INDIRECT / 3-D references"

// ---- edits that are more than cells ---------------------------------------------------
// Defined names, tables, filters and hidden rows, rows and columns inserted and deleted,
// sheets added, removed, renamed, copied and moved: the engine calculates the formulas
// that can tell (those that use a name or a table, SUBTOTAL / AGGREGATE, what reads the
// sheet whose rows moved, 3-D references ...) instead of leaving the workbook to the full
// recalculation. Every step is compared with the full recalculation of the edited workbook.
module Ops = Rbt.Spreadsheet.SheetOps
let moreWorkbook (iterate: bool) =
    let wb = Wb.createWorkbook () |> Wb.addSheet "Data" |> Wb.addSheet "Totals"
    let wb =
        wb |> setCells "sheet1"
                ([ for row in 1 .. 12 do
                     yield at row 0, Some (text (sprintf "item %d" (row % 4)))
                     yield at row 1, Some (number (float (row * 3 % 7 + 1)))
                     yield at row 2, Some (formula (sprintf "=B%d*2" (row + 1))) ]
                 @ [ at 0 0, Some (text "Item"); at 0 1, Some (text "Qty"); at 0 2, Some (text "Double")
                     at 14 1, Some (formula "=SUBTOTAL(109,B2:B13)"); at 14 2, Some (formula "=SUM(C2:C13)"); at 15 1, Some (formula "=AGGREGATE(9,5,B2:B13)")
                     at 16 1, Some (formula "=SUBTOTAL(9,B2:B13)+RATE"); at 17 1, Some (formula "=SUM(Sales[Qty])"); at 18 1, Some (formula "=ROWS(Sales[Qty])+Data!A1")
                     at 19 1, Some (formula "=SUM(Sheet1:Totals!F1)"); at 20 1, Some (formula "=Missing!A1+1"); at 21 1, Some (formula "=INDEX(B:B,5)+ROW()") ])
           |> setCells "sheet2" [ at 0 0, Some (number 5.0); at 1 0, Some (formula "=Sheet1!B15*RATE"); at 2 0, Some (formula "=SUM(Sheet1!C2:C13)+A1"); at 0 5, Some (number 2.0); at 3 0, Some (formula "=LOCALRATE+1") ]
           |> setCells "sheet3" [ at 0 0, Some (formula "=Data!A2+Data!A3"); at 1 0, Some (formula "=Sheet1!B17"); at 0 5, Some (number 3.0); at 2 0, Some (formula "=SUM(Data!A:A)") ]
    let wb = { wb with Names = wb.Names |> Map.add (None, "RATE") "=Data!$A$1" |> Map.add (Some "sheet2", "LOCALRATE") "=Data!$F$1*10" }
    let change, _ = Ops.defineTable "t" "sheet1" { From = at 0 0; To = at 12 2 } (Some "Sales") true None wb
    { change.Book with Iteration = (if iterate then Some { Wb.IterativeCalculation.MaxIterations = 20; MaxChange = 0.001 } else None) }
let handledOrFail label outcome = ok label (outcome <> Declined) (sprintf "left to the full recalculation: %s" (Recalc.lastDecline engine))
do
    start (moreWorkbook false)
    eq "more: calculated" (valueAt "sheet1" 14 1 current, valueAt "sheet1" 17 1 current, valueAt "sheet2" 1 0 current) (CellValue.Number 50.0, CellValue.Number 50.0, CellValue.Number 250.0)
    // A filter: only SUBTOTAL / AGGREGATE and what reads them.
    let filtered = edit "a filter set" None (fun wb -> Wb.updateSheetModel "sheet1" (wb.Sheets.["sheet1"].Model |> setAutoFilter (Some { From = at 0 0; To = at 12 2 }) |> setColumnFilter 0 (FilterValues (Set.ofList [ "item 1" ], false))) wb)
    handledOrFail "a filter set: by the engine" filtered
    (match filtered with
     | Incremental stats -> ok "a filter set: only SUBTOTAL / AGGREGATE and their readers are evaluated" (stats.EvaluatedFormulas >= 3 && stats.EvaluatedFormulas <= 9 && stats.EvaluatedFormulas * 2 < Recalc.formulaCount engine) (sprintf "%A of %d" stats (Recalc.formulaCount engine))
     | Declined -> ())
    ok "a filter set: SUBTOTAL leaves the hidden rows out, SUM does not" (valueAt "sheet1" 14 1 current = CellValue.Number 13.0 && valueAt "sheet1" 14 2 current = CellValue.Number 100.0) (sprintf "%A" (valueAt "sheet1" 14 1 current))
    handledOrFail "the filter cleared" (edit "the filter cleared" None (fun wb -> Wb.updateSheetModel "sheet1" (wb.Sheets.["sheet1"].Model |> setAutoFilter None) wb))
    eq "the filter cleared: the subtotal is back" (valueAt "sheet1" 14 1 current) (CellValue.Number 50.0)
    // A row hidden by hand: AGGREGATE(9,5,...) and SUBTOTAL(109,...) leave it out.
    handledOrFail "a row hidden" (edit "a row hidden" None (fun wb -> Wb.updateSheetModel "sheet1" (setRowHeight 3 0.0 wb.Sheets.["sheet1"].Model) wb))
    // Names.
    let renamed = edit "a name changed" None (Wb.setDefinedName None "RATE" "=Data!$F$1" None)
    handledOrFail "a name changed: by the engine" renamed
    (match renamed with
     | Incremental stats -> ok "a name changed: the formulas that use a name, and their readers" (stats.EvaluatedFormulas >= 2 && stats.EvaluatedFormulas <= 8 && stats.EvaluatedFormulas * 2 < Recalc.formulaCount engine) (sprintf "%A" stats)
     | Declined -> ())
    eq "a name changed: its users follow" (valueAt "sheet2" 1 0 current) (CellValue.Number 94.0)
    handledOrFail "a name deleted" (edit "a name deleted" None (fun wb -> (Wb.deleteDefinedName None "RATE" wb) |> Option.defaultValue wb))
    ok "a name deleted: #NAME?" (match valueAt "sheet2" 1 0 current with CellValue.Error _ -> true | _ -> false) (sprintf "%A" (valueAt "sheet2" 1 0 current))
    handledOrFail "a name defined again" (edit "a name defined" None (Wb.setDefinedName None "RATE" "=Data!$A$1" None))
    handledOrFail "a sheet-scoped name changed" (edit "a sheet-scoped name changed" None (Wb.setDefinedName (Some "sheet2") "LOCALRATE" "=Data!$F$1*100" None))
    eq "a sheet-scoped name changed: its user follows" (valueAt "sheet2" 3 0 current) (CellValue.Number 201.0)
    // Tables.
    handledOrFail "a table resized" (edit "a table resized" None (fun wb -> (Ops.resizeTable "t" "Sales" { From = at 0 0; To = at 8 2 } wb).Book))
    eq "a table resized: the structured references follow" (valueAt "sheet1" 18 1 current) (CellValue.Number 13.0)
    handledOrFail "a table renamed" (edit "a table renamed" None (fun wb -> (Ops.renameTable "t" "Sales" "Orders" wb).Book))
    handledOrFail "a table removed" (edit "a table removed" None (fun wb -> (Ops.deleteTable "t" "Orders" true wb).Book))
    handledOrFail "a table defined" (edit "a table defined" None (fun wb -> (fst (Ops.defineTable "t" "sheet1" { From = at 0 0; To = at 12 2 } (Some "Sales") true None wb)).Book))
    // Rows and columns.
    for label, sheetId, structureEdit in [ "a row inserted", "sheet1", InsertRows (5, 1); "rows deleted", "sheet1", DeleteRows (7, 2); "a column inserted", "sheet1", InsertCols (1, 1)
                                           "a column deleted", "sheet1", DeleteCols (1, 1); "a row inserted on the sheet that is read", "sheet2", InsertRows (0, 1)
                                           "a row deleted that a formula read", "sheet2", DeleteRows (0, 1); "a column inserted before a 3-D reference", "sheet3", InsertCols (2, 2) ] do
        handledOrFail label (edit label None (fun wb -> fst (Wb.applyStructureEdit sheetId structureEdit wb)))
    // Sheets.
    // The undo of a row / column edit is known as the edit taken back: formulas that only follow their cells keep their value unevaluated.
    do
        let before = current
        let inserted, undo = Wb.applyStructureEdit "sheet1" (InsertRows (20, 1)) current
        let forward = edit "a row inserted below the data" None (fun _ -> inserted)
        let back = edit "... and the insert undone" None (Wb.undoStructureEdit undo)
        match forward, back with
        | Incremental there, Incremental home ->
            ok "undo of a row insert: few formulas are evaluated, as for the insert" (home.EvaluatedFormulas <= there.EvaluatedFormulas + 2 && home.EvaluatedFormulas * 2 < Recalc.formulaCount engine) (sprintf "insert %A, undo %A of %d" there home (Recalc.formulaCount engine))
        | _ -> ok "undo of a row insert: by the engine" false (Recalc.lastDecline engine)
        ok "undo of a row insert: the workbook is what it was" ((difference current before).IsNone) (sprintf "%A" (difference current before))
    handledOrFail "a sheet renamed" (edit "a sheet renamed" None (Wb.renameSheetEverywhere "sheet2" "Numbers"))
    handledOrFail "a sheet renamed to a name formulas wait for" (edit "a sheet named Missing" None (fun wb -> Wb.addSheet "Missing" wb |> setCells "sheet4" [ at 0 0, Some (number 41.0) ]))
    eq "the reference to it resolves" (valueAt "sheet1" (if (match valueAt "sheet1" 20 1 current with CellValue.Number _ -> true | _ -> false) then 20 else 19) 1 current |> (fun v -> match v with CellValue.Number _ -> true | _ -> false)) true
    handledOrFail "sheets moved" (edit "sheets moved" None (Wb.moveSheet "sheet3" 0))
    handledOrFail "a sheet copied" (edit "a sheet copied" None (Wb.duplicateSheet "sheet1" "copy1" "Sheet1 (2)" 1))
    handledOrFail "a sheet removed" (edit "a sheet removed" None (fun wb -> match Wb.removeSheet "sheet2" wb with Some (after, _) -> after | None -> wb))
    handledOrFail "another removed" (edit "the copy removed" None (fun wb -> match Wb.removeSheet "copy1" wb with Some (after, _) -> after | None -> wb))
    handledOrFail "an edit of cells afterwards" (edit "cells after all that" None (setCells "sheet1" [ at 2 1, Some (number 40.0) ]))

/// Random sequences of such edits (and of cells) on that workbook.
let randomMore (iterate: bool) (seed: int) (steps: int) =
    let random = Random seed
    start (moreWorkbook iterate)
    let mutable incremental = 0
    let mutable copies = 0
    /// The row / column edit of the step before, to be taken back (an undo right after the edit).
    let mutable lastStructure : Wb.StructureEditUndo option = None
    let failuresBefore = fail
    for step in 1 .. steps do
        if fail = failuresBefore then
            let ids = current.SheetOrder |> List.toArray
            let sheetId = ids.[random.Next ids.Length]
            let model = resolveSheet current.Sheets.[sheetId].Model
            let undoable = lastStructure
            lastStructure <- None
            let structure (edit: StructureEdit) (wb: Wb.Workbook) =
                let edited, undo = Wb.applyStructureEdit sheetId edit wb
                lastStructure <- Some undo
                edited
            let name, change =
                match (if undoable.IsSome && random.Next 2 = 0 then 99 else random.Next 16) with
                | 99 -> "the row / column edit undone", Wb.undoStructureEdit undoable.Value
                | 0 | 1 -> "a number", setCells sheetId [ at (random.Next 14) (random.Next 3), Some (number (float (random.Next 90))) ]
                | 2 ->
                    let bodies = [| "=SUBTOTAL(109,B2:B13)"; "=RATE*3"; "=SUM(Sales[Qty])"; "=Sheet1!B15+1"; "=SUM(Sheet1:Totals!F1)"; "=Data!A1"; "=AGGREGATE(9,5,B2:B9)"; "=E5+1"; "=F6*2"; "=Nowhere!B2"; "=ROW()+COLUMN()"; "=SUM(A:A)"; "=LOCALRATE"; "=INDEX(Sheet1!B:B,6)"; "=SEQUENCE(3)"; "=B2:B4*2"; "=SORT(Sheet1!B2:B5)"; "=OFFSET(Sheet1!A1,2,1)"; "=SUM(H3#)" |]
                    "a formula", setCells sheetId [ at (random.Next 22) (3 + random.Next 4), Some (formula bodies.[random.Next bodies.Length]) ]
                | 3 -> "a name set", Wb.setDefinedName None (if random.Next 2 = 0 then "RATE" else "OTHER") (sprintf "=Data!$%s$%d" (if random.Next 2 = 0 then "A" else "F") (1 + random.Next 3)) None
                | 4 -> "a name deleted", (fun wb -> Wb.deleteDefinedName None (if random.Next 2 = 0 then "RATE" else "OTHER") wb |> Option.defaultValue wb)
                | 5 -> "rows inserted", structure (InsertRows (random.Next 16, 1 + random.Next 2))
                | 6 -> "rows deleted", structure (DeleteRows (1 + random.Next 16, 1 + random.Next 2))
                | 7 -> "columns inserted", structure (InsertCols (random.Next 5, 1))
                | 8 -> "columns deleted", structure (DeleteCols (3 + random.Next 4, 1))
                | 9 ->
                    "a filter", (fun wb ->
                        let m = resolveSheet wb.Sheets.[sheetId].Model
                        let next =
                            if m.AutoFilter.IsSome && random.Next 2 = 0 then setAutoFilter None m
                            else m |> setAutoFilter (Some { From = at 0 0; To = at 12 2 }) |> setColumnFilter 1 (FilterCustom ([ OpGreaterThan, string (random.Next 6) ], true))
                        Wb.updateSheetModel sheetId next wb)
                | 10 -> "a row hidden or shown", (fun wb -> Wb.updateSheetModel sheetId (setRowHeight (1 + random.Next 12) (if random.Next 2 = 0 then 0.0 else 20.0) (resolveSheet wb.Sheets.[sheetId].Model)) wb)
                | 11 -> "a sheet renamed", Wb.renameSheetEverywhere sheetId (sprintf "%s%d" (if random.Next 3 = 0 then "Nowhere" else "Tab") step)
                | 12 -> "a sheet moved", Wb.moveSheet sheetId (random.Next ids.Length)
                | 13 when ids.Length < 6 ->
                    copies <- copies + 1
                    let id = sprintf "extra%d" copies
                    (if random.Next 2 = 0 then "a sheet copied", Wb.duplicateSheet sheetId id (sprintf "Copy %d" copies) (random.Next ids.Length)
                     else "a sheet added", (fun wb -> { wb with Sheets = Map.add id ({ Id = id; Name = sprintf "Added %d" copies; Model = createEmpty 50 10; Visible = true } : Wb.Sheet) wb.Sheets; SheetOrder = wb.SheetOrder @ [ id ] }
                                                     |> setCells id [ at 0 0, Some (number 9.0); at 1 0, Some (formula "=A1+Sheet1!B2") ]))
                | 14 when ids.Length > 2 -> "a sheet removed", (fun wb -> match Wb.removeSheet sheetId wb with Some (after, _) -> after | None -> wb)
                | _ ->
                    "a table changed", (fun wb ->
                        try
                            if wb.Tables.IsEmpty then (fst (Ops.defineTable "t" wb.SheetOrder.Head { From = at 0 0; To = at (6 + random.Next 6) 2 } (Some "Sales") true None wb)).Book
                            elif random.Next 3 = 0 then (Ops.deleteTable "t" (wb.Tables |> Seq.head).Value.Name true wb).Book
                            else
                                let table = (wb.Tables |> Seq.head).Value
                                (Ops.resizeTable "t" table.Name { From = at table.StartRow table.StartCol; To = at (table.StartRow + 3 + random.Next 9) table.EndCol } wb).Book
                        with _ -> wb)
            ignore model
            match edit (sprintf "more%s seed %d step %d (%s)" (if iterate then " (iterating)" else "") seed step name) None change with
            | Incremental _ -> incremental <- incremental + 1
            | Declined -> ()
    incremental
do
    declines.Clear()
    let seeds = (match Environment.GetEnvironmentVariable "MORE_SEEDS" with null | "" -> 40 | value -> int value)
    let steps = 40
    let mutable incremental = 0
    for seed in 1 .. seeds do
        incremental <- incremental + randomMore (seed % 4 = 0) (900 + seed) steps
    printfn "  edits that are more than cells: %d of %d random edits incremental; left to the full recalculation: %s" incremental (seeds * steps)
        (if declines.Count = 0 then "none" else declines |> Seq.map (fun kv -> sprintf "%d x \"%s\"" kv.Value kv.Key) |> String.concat ", ")
    ok "edits that are more than cells are calculated by the engine" (incremental * 100 >= seeds * steps * 97) (sprintf "%d of %d" incremental (seeds * steps))
    reportFullRuns "with edits that are more than cells"

// The formula index: known for most of the cell maps that went by, and never wrong (checked in `edit`).
printfn "  formula index: known for %d of %d cell maps checked" indexKnown indexChecks
ok "the formula index was known for most cell maps" (indexKnown * 2 > indexChecks) (sprintf "%d of %d" indexKnown indexChecks)
// A sheet of values with a few formulas: an opened file's index is there without a walk, and it is exact.
do
    let dashboard = load "real-world/dashboard-sales-analysis.xlsx"
    for KeyValue (sheetId, sheet) in dashboard.Sheets do
        let cells = (resolveSheet sheet.Model).Cells
        let walked = cells |> Map.toSeq |> Seq.filter (fun (_, cell) -> match cell.Value with CellValue.Formula _ -> true | _ -> false) |> Seq.map fst |> Set.ofSeq
        eq (sprintf "an opened sheet has its formula index (%s)" sheetId) (tryFormulaIndex cells) (Some walked)
    let recalculated = Formula.recalculateWorkbook dashboard
    checkFormulaIndex "the opened workbook recalculated" recalculated

// "Does this formula spill" is asked of every formula when the graph is built, in one walk of the
// formula: it answers as the two walks it replaced, on every formula of the real workbooks and on
// formulas of every shape the test looks at.
do
    let shapes =
        [ "=A1"; "=A1:A3"; "=A:A"; "=1:1"; "=SUM(A1:A3)"; "=SUM(A1:A3*2)"; "=A1:A3*2"; "=-A1:A3"; "=LEN(A1:A3)"; "=LEN(A:A)"; "=SEQUENCE(3)"; "=SUM(SEQUENCE(3))"
          "=IF(A1,SEQUENCE(2),1)"; "=OFFSET(A1,0,0,3,1)"; "=SUM(OFFSET(A1,0,0,3,1))"; "=OFFSET(A1,SEQUENCE(2),0)"; "=INDIRECT(\"A1:A3\")"; "=INDEX(A1:C3,0,2)"; "=INDEX(A1:C3,2,2)"
          "=INDEX(A1:C3,B1:B2,1)"; "=INDEX(A1:A9,2)"; "=INDEX(A1:C3,2)"; "=INDEX(SEQUENCE(3),2)"; "=ROW(A1:A5)"; "=ROW(A1)"; "=COLUMN(A1:C1)"; "=ROWS(A1:A5)"; "={1,2;3,4}"; "=SUM({1,2})"
          "=A1#"; "=SUM(A1#)"; "=Sheet1:Data!A1"; "=SUM(Sheet1:Data!A1)"; "=(A1:A3,C1:C3)"; "=SUM((A1:A3,C1:C3))"; "=A1:C3 B2:D4"; "=A1:INDEX(A:A,3)"; "=SUM(A1:INDEX(A:A,3))"
          "=VLOOKUP(A1:A3,C1:D9,2,0)"; "=VLOOKUP(A1,C1:D9,2,0)"; "=IF(A1:A3>1,1,0)"; "=SUMPRODUCT(A1:A3,B1:B3)"; "=SUMPRODUCT((A1:A3>1)*B1:B3)"; "=MYSTERY(A1:A3)"; "=MYSTERY(SORT(A1:A3))"
          "=TEXT(A1:A2,\"0\")&\"x\""; "=T[@Amount]"; "=T[Amount]"; "=SUM(T[Amount])"; "=FILTER(A1:A9,B1:B9>1)"; "=LET(x,A1:A3,x*2)"; "=1+2"; "=NOW()"; "=RATE*2" ]
    let mutable asked = 0
    let mutable differ = 0
    let ask (formula: string) =
        match Formula.parseFormula formula with
        | Ok expr ->
            asked <- asked + 1
            for dynamicReferences in [ false; true ] do
                let one = if dynamicReferences then Formula.exprSpills expr else Formula.exprSpillsStatically expr
                if one <> Formula.exprSpillsReference dynamicReferences expr then
                    differ <- differ + 1
                    if differ <= 5 then printfn "FAIL one walk says %b of %s (OFFSET / INDIRECT counted: %b)" one formula dynamicReferences
        | Result.Error _ -> ()
    for formula in shapes do ask formula
    ok "the shapes parse" (asked >= shapes.Length - 3) (string asked)
    if Environment.GetEnvironmentVariable "SKIP_FILES" <> "1" then
        for file in [ "milk-street-office.xlsx"; "lyric-multifamily.xlsx"; "dcf-stock-analysis.xlsx"; "saopaulo-condos.xlsx"; "jumeirah-hotel.xlsx"; "dashboard-sales-analysis.xlsx"; "chart-zoo.xlsx"; "conditional-formatting-tutorial.xlsx" ] do
            for KeyValue (_, sheet) in (load ("real-world/" + file)).Sheets do
                for KeyValue (_, cell) in (resolveSheet sheet.Model).Cells do
                    match cell.Value with
                    | CellValue.Formula (formula, _) -> ask formula
                    | _ -> ()
    eq (sprintf "spills: one walk answers as two (%d formulas)" asked) differ 0

printfn "stage32-recalc: %d passed, %d failed" pass fail
if fail > 0 then exit 1
