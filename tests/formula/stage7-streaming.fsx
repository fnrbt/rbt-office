#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Streaming/fusion engine (see docs/calc-streaming-fusion-design.md):
//  - RangeView is a lazy struct-enumerator pull over the sparse cell map.
//  - Fold is a one-pass reducer; `both` runs two folds in a SINGLE traversal.
//  - Aggregates (SUM/AVERAGE/MIN/MAX/COUNT/…) stream via foldArgs — no list.
// This stage proves the primitives, that the retrofit preserves semantics on
// the tricky edges (errors, blanks, booleans, empty), and that a 100k-cell
// range aggregates fast without materializing a list.

open System.Diagnostics
open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-30s expected %A got %A" label expected got

// ── 1. Primitives: RangeView + Fold + run + both ──────────────────────
// A1:A6 = 10..60 in column A.
let small =
    { createEmpty 100 26 with
        Cells = [ for r in 0..5 -> { Row = r; Col = 0 }, cell (CellValue.Number (float ((r + 1) * 10))) ] |> Map.ofList }
let view = { Model = small; SR = 0; SC = 0; ER = 5; EC = 0 }

let sumFold = numAggFold 0.0 (+) (fun _ a -> CellValue.Number a)
let numCount = countFold false (fun v -> match v with CellValue.Number _ -> true | _ -> false)

check "run sum over view" (run view sumFold) (CellValue.Number 210.0)

// One-pass: SUM and COUNT from a SINGLE traversal via `both`.
let struct (sumV, cntV) = run view (both sumFold numCount)
check "both: sum"   sumV (CellValue.Number 210.0)
check "both: count" cntV (CellValue.Number 6.0)
let avg = match sumV, cntV with CellValue.Number s, CellValue.Number n -> s / n | _ -> nan
check "both: average (one pass)" avg 35.0

// Prove the producer yields each element exactly once (single pass).
let visits = ref 0
let tap : Fold<CellValue, int, int> = { Init = 0; Step = (fun n _ -> visits.Value <- visits.Value + 1; n + 1); Extract = id }
check "run visits N once" (run view tap) 6
check "visit count == N"  visits.Value 6

// ── 2. Retrofit edge-case parity (via evaluateFormula) ────────────────
let mixed =
    { createEmpty 100 26 with
        Cells =
            [ { Row=0; Col=0 }, cell (CellValue.Number 5.0)
              { Row=1; Col=0 }, cell (CellValue.String "txt")     // skipped by numeric aggs
              { Row=2; Col=0 }, cell (CellValue.Boolean true)     // counts as 1 in SUM, not in COUNT
              { Row=3; Col=0 }, cell (CellValue.Number 15.0)
              { Row=4; Col=0 }, cell (CellValue.String "")        // empty string
              { Row=0; Col=1 }, cell (CellValue.Number 1.0)
              { Row=1; Col=1 }, cell (CellValue.Error "DIV/0!")   // propagates through SUM/COUNT
              { Row=0; Col=2 }, cell (CellValue.String "a")       // all-text column C1:C3
              { Row=1; Col=2 }, cell (CellValue.String "b") ]
            |> Map.ofList }
let ev f = evaluateFormula f mixed None
let n x = CellValue.Number x
let e x = CellValue.Error x

check "sum skips text, counts bool" (ev "=SUM(A1:A5)")   (n 21.0)   // 5 + 1(true) + 15
check "average over numerics+bool"  (ev "=AVERAGE(A1:A5)") (n 7.0)  // 21 / 3
check "min ignores text"            (ev "=MIN(A1:A5)")    (n 1.0)   // true=1
check "max ignores text"            (ev "=MAX(A1:A5)")    (n 15.0)
check "count numbers only"          (ev "=COUNT(A1:A5)")  (n 2.0)   // 5,15 (bool/text excluded)
check "counta non-empty"            (ev "=COUNTA(A1:A5)") (n 5.0)   // 5,txt,true,15,"" — Excel counts "" as non-empty
check "countblank empty+emptystr"   (ev "=COUNTBLANK(A1:A5)") (n 1.0) // "" counts as blank too (Excel quirk)
let sparseBlankModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 9999; Col = 0 }, cell (CellValue.String "")
              { Row = 14999; Col = 0 }, cell CellValue.Empty
              { Row = 19999; Col = 0 }, cell (CellValue.String "x") ]
            |> Map.ofList }
let evSparseBlank f = evaluateFormula f sparseBlankModel None
check "countblank sparse range"     (evSparseBlank "=COUNTBLANK(A1:A20000)") (n 19998.0)
check "countblank sparse + scalar"  (evSparseBlank "=COUNTBLANK(A1:A20000,\"\")") (n 19999.0)
check "sum propagates error"        (ev "=SUM(B1:B2)")    (e "DIV/0!")
check "count propagates error"      (ev "=COUNT(B1:B2)")  (e "DIV/0!")
check "sum empty range = 0"         (ev "=SUM(Z1:Z9)")    (n 0.0)
check "average empty = DIV/0"       (ev "=AVERAGE(Z1:Z9)") (e "DIV/0!")
check "min all-text = 0"            (ev "=MIN(C1:C3)")    (n 0.0)
check "max all-text = 0"            (ev "=MAX(C1:C3)")    (n 0.0)
check "product empty = 0"           (ev "=PRODUCT(Z1:Z9)") (n 0.0)
check "multi-arg sum ranges+scalar" (ev "=SUM(A1:A5,B1:B1,100)") (n 122.0) // 21 + 1 + 100

// ── 3. Large range: 100k cells, streamed (no list materialized) ───────
let bigCells = [ for r in 0..99999 -> { Row = r; Col = 0 }, cell (CellValue.Number (float (r + 1))) ] |> Map.ofList
let big = { createEmpty 100000 4 with Cells = bigCells }
let evBig f = evaluateFormula f big None
let sw = Stopwatch.StartNew()
let bigSum = evBig "=SUM(A1:A100000)"
let bigAvg = evBig "=AVERAGE(A1:A100000)"
let bigCnt = evBig "=COUNT(A1:A100000)"
sw.Stop()
check "100k SUM"     bigSum (n 5000050000.0)
check "100k AVERAGE" bigAvg (n 50000.5)
check "100k COUNT"   bigCnt (n 100000.0)
printfn "  (100k SUM+AVERAGE+COUNT in %d ms)" sw.ElapsedMilliseconds

// ── 4. Large criteria/lookup family, streamed by index (no per-range list) ──
// 50k-row table: A=sorted keys 1..50000, B=values (key*2), C="even"/"odd".
let tbl =
    { createEmpty 50000 4 with
        Cells =
            [ for r in 0..49999 do
                let i = r + 1
                yield { Row=r; Col=0 }, cell (CellValue.Number (float i))
                yield { Row=r; Col=1 }, cell (CellValue.Number (float (i * 2)))
                yield { Row=r; Col=2 }, cell (CellValue.String (if i % 2 = 0 then "even" else "odd")) ]
            |> Map.ofList }
let evT f = evaluateFormula f tbl None
let sw2 = Stopwatch.StartNew()
let r_countif  = evT "=COUNTIF(C1:C50000,\"even\")"
let r_sumif    = evT "=SUMIF(A1:A50000,\">49998\")"
let r_sumifs   = evT "=SUMIFS(B1:B50000,A1:A50000,\">49998\")"
let r_avgifs   = evT "=AVERAGEIFS(B1:B50000,C1:C50000,\"odd\",A1:A50000,\"<=4\")" // odd keys 1,3 → vals 2,6 → avg 4
let r_match    = evT "=MATCH(40000,A1:A50000,0)"
let r_vlookup  = evT "=VLOOKUP(40000,A1:B50000,2,TRUE)"
let r_xlookup  = evT "=XLOOKUP(40000,A1:A50000,B1:B50000)"
let r_xmatch   = evT "=XMATCH(40000,A1:A50000)"
sw2.Stop()
check "50k COUNTIF"   r_countif  (n 25000.0)
check "50k SUMIF"     r_sumif    (n 99999.0)     // 49999 + 50000
check "50k SUMIFS"    r_sumifs   (n 199998.0)    // (49999+50000)*2
check "50k AVERAGEIFS" r_avgifs  (n 4.0)
check "50k MATCH"     r_match    (n 40000.0)
check "50k VLOOKUP"   r_vlookup  (n 80000.0)
check "50k XLOOKUP"   r_xlookup  (n 80000.0)
check "50k XMATCH"    r_xmatch   (n 40000.0)
printfn "  (8 criteria/lookup calls over 50k rows in %d ms)" sw2.ElapsedMilliseconds

let sparseCriteriaModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.String "match")
              { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 0.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 14999; Col = 0 }, cell (CellValue.String "other")
              { Row = 14999; Col = 1 }, cell (CellValue.Number 30.0)
              { Row = 19999; Col = 0 }, cell (CellValue.String "match") ]
            |> Map.ofList }
let evSparseCriteria f = evaluateFormula f sparseCriteriaModel None
check "countif sparse text"  (evSparseCriteria "=COUNTIF(A1:A20000,\"match\")") (n 2.0)
check "countif sparse zero"  (evSparseCriteria "=COUNTIF(A1:A20000,0)") (n 19997.0)
check "countif sparse blank" (evSparseCriteria "=COUNTIF(A1:A20000,\"\")") (n 19996.0)
check "sumif sparse text"    (evSparseCriteria "=SUMIF(A1:A20000,\"match\",B1:B20000)") (n 10.0)
check "sumif sparse zero"    (evSparseCriteria "=SUMIF(A1:A20000,0,B1:B20000)") (n 20.0)
check "sumif sparse blank"   (evSparseCriteria "=SUMIF(A1:A20000,\"\",B1:B20000)") (n 0.0)
check "averageif sparse text" (evSparseCriteria "=AVERAGEIF(A1:A20000,\"match\",B1:B20000)") (n 5.0)
check "countifs sparse driver" (evSparseCriteria "=COUNTIFS(A1:A20000,\"match\",B1:B20000,\">0\")") (n 1.0)
check "sumifs sparse driver" (evSparseCriteria "=SUMIFS(B1:B20000,A1:A20000,\"match\",B1:B20000,\">0\")") (n 10.0)
check "maxifs sparse driver" (evSparseCriteria "=MAXIFS(B1:B20000,A1:A20000,\"match\",B1:B20000,\">0\")") (n 10.0)
check "minifs sparse driver" (evSparseCriteria "=MINIFS(B1:B20000,A1:A20000,\"match\",B1:B20000,\">0\")") (n 10.0)
check "averageifs sparse driver" (evSparseCriteria "=AVERAGEIFS(B1:B20000,A1:A20000,\"match\",B1:B20000,\">0\")") (n 10.0)

let denseSumSparseCriteriaModel =
    { createEmpty 20010 4 with
        Cells =
            [ for r in 0 .. 19999 do
                yield { Row = r; Col = 1 }, cell (CellValue.Number 1.0)
              yield { Row = 0; Col = 0 }, cell (CellValue.String "skip")
              yield { Row = 9999; Col = 0 }, cell (CellValue.String "skip")
              yield { Row = 19999; Col = 0 }, cell (CellValue.String "skip") ]
            |> Map.ofList }
check "countif sparse blank dense unrelated" (evaluateFormula "=COUNTIF(A1:A20000,\"\")" denseSumSparseCriteriaModel None) (n 19997.0)
check "sumif sparse blank dense sum" (evaluateFormula "=SUMIF(A1:A20000,\"\",B1:B20000)" denseSumSparseCriteriaModel None) (n 19997.0)
check "averageif sparse blank dense avg" (evaluateFormula "=AVERAGEIF(A1:A20000,\"\",B1:B20000)" denseSumSparseCriteriaModel None) (n 1.0)
check "countif sparse text dense unrelated" (evaluateFormula "=COUNTIF(A1:A20000,\"skip\")" denseSumSparseCriteriaModel None) (n 3.0)
check "sumif sparse text dense sum" (evaluateFormula "=SUMIF(A1:A20000,\"skip\",B1:B20000)" denseSumSparseCriteriaModel None) (n 3.0)
check "averageif sparse text dense avg" (evaluateFormula "=AVERAGEIF(A1:A20000,\"skip\",B1:B20000)" denseSumSparseCriteriaModel None) (n 1.0)

let sparseAvgSparseCriteriaModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 30.0)
              { Row = 0; Col = 0 }, cell (CellValue.String "skip")
              { Row = 19999; Col = 0 }, cell (CellValue.String "skip") ]
            |> Map.ofList }
check "averageif sparse blank sparse avg" (evaluateFormula "=AVERAGEIF(A1:A20000,\"\",B1:B20000)" sparseAvgSparseCriteriaModel None) (n 20.0)

let denseValueSparseIfsCriteriaModel =
    { createEmpty 20010 4 with
        Cells =
            [ for r in 0 .. 19999 do
                yield { Row = r; Col = 2 }, cell (CellValue.Number 1.0)
              yield { Row = 0; Col = 0 }, cell (CellValue.String "skip")
              yield { Row = 9999; Col = 0 }, cell (CellValue.String "skip")
              yield { Row = 19999; Col = 0 }, cell (CellValue.String "skip")
              yield { Row = 10; Col = 1 }, cell (CellValue.String "skip")
              yield { Row = 10000; Col = 1 }, cell (CellValue.String "skip")
              yield { Row = 19998; Col = 1 }, cell (CellValue.String "skip") ]
            |> Map.ofList }
check "sumifs sparse blank dense values" (evaluateFormula "=SUMIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" denseValueSparseIfsCriteriaModel None) (n 19994.0)
check "countifs sparse blank dense values" (evaluateFormula "=COUNTIFS(A1:A20000,\"\",B1:B20000,\"\")" denseValueSparseIfsCriteriaModel None) (n 19994.0)
check "averageifs sparse blank dense values" (evaluateFormula "=AVERAGEIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" denseValueSparseIfsCriteriaModel None) (n 1.0)
check "countifs sparse text dense values" (evaluateFormula "=COUNTIFS(A1:A20000,\"skip\",C1:C20000,\">0\")" denseValueSparseIfsCriteriaModel None) (n 3.0)
check "sumifs sparse text dense values" (evaluateFormula "=SUMIFS(C1:C20000,A1:A20000,\"skip\",C1:C20000,\">0\")" denseValueSparseIfsCriteriaModel None) (n 3.0)
check "averageifs sparse text dense values" (evaluateFormula "=AVERAGEIFS(C1:C20000,A1:A20000,\"skip\",C1:C20000,\">0\")" denseValueSparseIfsCriteriaModel None) (n 1.0)
check "maxifs sparse text dense values" (evaluateFormula "=MAXIFS(C1:C20000,A1:A20000,\"skip\",C1:C20000,\">0\")" denseValueSparseIfsCriteriaModel None) (n 1.0)
check "minifs sparse text dense values" (evaluateFormula "=MINIFS(C1:C20000,A1:A20000,\"skip\",C1:C20000,\">0\")" denseValueSparseIfsCriteriaModel None) (n 1.0)

let sparseValueSparseIfsCriteriaModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 2 }, cell (CellValue.Number 10.0)
              { Row = 4999; Col = 2 }, cell (CellValue.Number 20.0)
              { Row = 9999; Col = 2 }, cell (CellValue.Number 30.0)
              { Row = 14999; Col = 2 }, cell (CellValue.Number 40.0)
              { Row = 19999; Col = 2 }, cell (CellValue.Number 50.0)
              { Row = 0; Col = 0 }, cell (CellValue.String "skip")
              { Row = 19999; Col = 0 }, cell (CellValue.String "skip")
              { Row = 4999; Col = 1 }, cell (CellValue.String "skip") ]
            |> Map.ofList }
check "sumifs sparse blank sparse values" (evaluateFormula "=SUMIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" sparseValueSparseIfsCriteriaModel None) (n 70.0)
check "averageifs sparse blank sparse values" (evaluateFormula "=AVERAGEIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" sparseValueSparseIfsCriteriaModel None) (n 35.0)
check "maxifs sparse blank sparse values" (evaluateFormula "=MAXIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" sparseValueSparseIfsCriteriaModel None) (n 40.0)
check "minifs sparse blank sparse values" (evaluateFormula "=MINIFS(C1:C20000,A1:A20000,\"\",B1:B20000,\"\")" sparseValueSparseIfsCriteriaModel None) (n 30.0)

let sparseProductModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 3.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 4.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 5.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 6.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 7.0)
              { Row = 0; Col = 2 }, cell (CellValue.Number 11.0)
              { Row = 9999; Col = 2 }, cell (CellValue.Number 13.0)
              { Row = 19999; Col = 2 }, cell (CellValue.Number 17.0) ]
            |> Map.ofList }
let evSparseProduct f = evaluateFormula f sparseProductModel None
check "sumproduct sparse" (evSparseProduct "=SUMPRODUCT(A1:A20000,B1:B20000)") (n 56.0)
check "sumproduct sparse three ranges" (evSparseProduct "=SUMPRODUCT(A1:A20000,B1:B20000,C1:C20000)") (n 820.0)

let sparseLookupModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.String "first")
              { Row = 9999; Col = 0 }, cell (CellValue.String "needle")
              { Row = 19999; Col = 0 }, cell (CellValue.String "last")
              { Row = 0; Col = 1 }, cell (CellValue.Number 1.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 42.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 99.0) ]
            |> Map.ofList }
let evSparseLookup f = evaluateFormula f sparseLookupModel None
check "match sparse exact" (evSparseLookup "=MATCH(\"needle\",A1:A20000,0)") (n 10000.0)
check "xmatch sparse exact" (evSparseLookup "=XMATCH(\"needle\",A1:A20000,0)") (n 10000.0)
check "xlookup sparse exact" (evSparseLookup "=XLOOKUP(\"needle\",A1:A20000,B1:B20000)") (n 42.0)
check "vlookup sparse exact" (evSparseLookup "=VLOOKUP(\"needle\",A1:B20000,2,FALSE)") (n 42.0)
let crowdedSparseLookupModel =
    { sparseLookupModel with
        Cells =
            [ yield! sparseLookupModel.Cells |> Map.toList
              for r in 0 .. 19999 do
                  yield { Row = r; Col = 2 }, cell (CellValue.Number (float r))
                  yield { Row = r; Col = 3 }, cell (CellValue.Number (float (r * 2))) ]
            |> Map.ofList }
let evCrowdedSparseLookup f = evaluateFormula f crowdedSparseLookupModel None
check "match crowded sparse first" (evCrowdedSparseLookup "=MATCH(\"first\",A1:A20000,0)") (n 1.0)
check "match crowded sparse exact" (evCrowdedSparseLookup "=MATCH(\"needle\",A1:A20000,0)") (n 10000.0)
check "vlookup crowded sparse exact" (evCrowdedSparseLookup "=VLOOKUP(\"needle\",A1:B20000,2,FALSE)") (n 42.0)
let sparseHLookupModel =
    { createEmpty 4 2000 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.String "first")
              { Row = 0; Col = 999 }, cell (CellValue.String "needle")
              { Row = 0; Col = 1999 }, cell (CellValue.String "last")
              { Row = 1; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 1; Col = 999 }, cell (CellValue.Number 42.0)
              { Row = 1; Col = 1999 }, cell (CellValue.Number 99.0) ]
            |> Map.ofList }
let evSparseHLookup f = evaluateFormula f sparseHLookupModel None
check "hlookup sparse exact" (evSparseHLookup "=HLOOKUP(\"needle\",A1:BXX2,2,FALSE)") (n 42.0)
let sparseApproxLookupModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 42.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 90.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 100.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 420.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 900.0) ]
            |> Map.ofList }
let evSparseApproxLookup f = evaluateFormula f sparseApproxLookupModel None
check "xmatch sparse next smaller" (evSparseApproxLookup "=XMATCH(50,A1:A20000,-1)") (n 10000.0)
check "xlookup sparse next smaller" (evSparseApproxLookup "=XLOOKUP(50,A1:A20000,B1:B20000,\"x\",-1)") (n 420.0)
check "xlookup sparse next larger" (evSparseApproxLookup "=XLOOKUP(50,A1:A20000,B1:B20000,\"x\",1)") (n 900.0)

let sparseFilterModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 1.0)
              { Row = 0; Col = 1 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 1 }, cell (CellValue.Number 20.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 3.0)
              { Row = 19999; Col = 1 }, cell (CellValue.Number 30.0)
              { Row = 9999; Col = 2 }, cell (CellValue.Boolean true) ]
            |> Map.ofList }
let arrSparseFilter f =
    match parseFormula f with
    | Ok expr -> evalArray expr { Model = sparseFilterModel; Workbook = None; Sheet = None; CriteriaCache = None; CurrentCell = None }
    | Result.Error e -> [| [| CellValue.Error e |] |]
check "filter sparse include rows" (arrSparseFilter "=FILTER(A1:B20000,C1:C20000)") [| [| n 2.0; n 20.0 |] |]
check "filter sparse include empty" (arrSparseFilter "=FILTER(A1:B20000,D1:D20000,\"none\")") [| [| CellValue.String "none" |] |]

printfn ""
printfn "Stage 7 streaming tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
