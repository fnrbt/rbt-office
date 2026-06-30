#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Stats / date / lookup batch: AVERAGEA/MAXA/MINA, MODE, RANK, PERCENTILE,
// QUARTILE, DEVSQ, AVEDEV, GEOMEAN, HARMEAN, VAR.S/.P, STDEV.S/.P; WEEKNUM,
// DATEDIF, NETWORKDAYS, WORKDAY; LOOKUP; NUMBERVALUE; TEXTSPLIT.

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Formula

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let nC r c v = ({ Row=r; Col=c }, cell (CellValue.Number v))
let sC r c v = ({ Row=r; Col=c }, cell (CellValue.String v))
// A1:A8 = 2,4,4,4,5,5,7,9 ; B1:B5 = 10,20,30,40,50 (sorted, for LOOKUP);
// C1:C5 = "a".."e" (result vector) ; D1:D3 = 5,3,8 with E having text for AVERAGEA
let model =
    { createEmpty 100 26 with
        Cells =
            ([ 2.;4.;4.;4.;5.;5.;7.;9. ] |> List.mapi (fun i v -> nC i 0 v))
            @ ([ 10.;20.;30.;40.;50. ] |> List.mapi (fun i v -> nC i 1 v))
            @ ([ "a";"b";"c";"d";"e" ] |> List.mapi (fun i v -> sC i 2 v))
            @ [ nC 0 3 10.0; sC 1 3 "text"; nC 2 3 20.0 ]   // D1:D3 mixed for AVERAGEA
            |> Map.ofList }

let mutable pass = 0
let mutable fail = 0
let ev f = evaluateFormula f model None
let n x = CellValue.Number x
let s x = CellValue.String x
let check label f (expected: CellValue) =
    let got = ev f
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-22s %-30s expected %A got %A" label f expected got
let approx label f (expected: float) =
    match ev f with
    | CellValue.Number g when abs (g - expected) < 1e-9 -> pass <- pass + 1
    | got -> fail <- fail + 1; printfn "FAIL %-22s %-30s expected ~%f got %A" label f expected got
let arrOf f = match parseFormula f with Ok e -> evalArray e (contextOf model) | Result.Error e -> failwithf "parse: %s" e
let checkV label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-22s expected %A got %A" label expected got

// ── stats ──
check  "mode"        "=MODE(A1:A8)"        (n 4.0)
let sparseCountRankModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 0 }, cell (CellValue.String "x")
              { Row = 19999; Col = 0 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
checkV "count sparse blanks" (evaluateFormula "=COUNT(A1:A20000)" sparseCountRankModel None) (n 2.0)
checkV "counta sparse blanks" (evaluateFormula "=COUNTA(A1:A20000)" sparseCountRankModel None) (n 3.0)
checkV "rank sparse blanks" (evaluateFormula "=RANK(10,A1:A20000)" sparseCountRankModel None) (n 2.0)
let sparseLogicalModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Boolean true)
              { Row = 9999; Col = 0 }, cell (CellValue.String "TRUE")
              { Row = 19999; Col = 0 }, cell (CellValue.Number 1.0) ]
            |> Map.ofList }
checkV "and sparse blanks" (evaluateFormula "=AND(A1:A20000)" sparseLogicalModel None) (CellValue.Boolean true)
checkV "or sparse blanks" (evaluateFormula "=OR(A1:A20000)" sparseLogicalModel None) (CellValue.Boolean true)
checkV "xor sparse blanks" (evaluateFormula "=XOR(A1:A20000)" sparseLogicalModel None) (CellValue.Boolean true)
let sparseModeModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 7.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 7.0) ]
            |> Map.ofList }
let sparseMode = evaluateFormula "=MODE(A1:A20000)" sparseModeModel None
checkV "mode sparse blanks" sparseMode (n 7.0)
check  "rank-desc"   "=RANK(7,A1:A8)"      (n 2.0)   // 9>7 → rank 2
check  "rank-asc"    "=RANK(7,A1:A8,1)"    (n 7.0)   // seven values <7
approx "percentile"  "=PERCENTILE(A1:A8,0.25)" 4.0
approx "median-via-pct" "=PERCENTILE(A1:A8,0.5)" 4.5
check  "quartile-0"  "=QUARTILE(A1:A8,0)"  (n 2.0)
check  "quartile-4"  "=QUARTILE(A1:A8,4)"  (n 9.0)
let sparseNumericArrayModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 9999; Col = 0 }, cell (CellValue.Number 8.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
checkV "median sparse blanks" (evaluateFormula "=MEDIAN(A1:A20000)" sparseNumericArrayModel None) (n 8.0)
checkV "large sparse blanks" (evaluateFormula "=LARGE(A1:A20000,1)" sparseNumericArrayModel None) (n 20.0)
checkV "small sparse blanks" (evaluateFormula "=SMALL(A1:A20000,1)" sparseNumericArrayModel None) (n 2.0)
checkV "percentile sparse blanks" (evaluateFormula "=PERCENTILE(A1:A20000,0.5)" sparseNumericArrayModel None) (n 8.0)
checkV "quartile sparse blanks" (evaluateFormula "=QUARTILE(A1:A20000,4)" sparseNumericArrayModel None) (n 20.0)
let sparseTextOnlyModel =
    { createEmpty 20010 4 with
        Cells = [ { Row = 9999; Col = 0 }, cell (CellValue.String "x") ] |> Map.ofList }
checkV "median sparse no numbers" (evaluateFormula "=MEDIAN(A1:A20000)" sparseTextOnlyModel None) (CellValue.Error "NUM!")
checkV "large sparse no numbers" (evaluateFormula "=LARGE(A1:A20000,1)" sparseTextOnlyModel None) (CellValue.Error "NUM!")
checkV "small sparse no numbers" (evaluateFormula "=SMALL(A1:A20000,1)" sparseTextOnlyModel None) (CellValue.Error "NUM!")
checkV "percentile sparse no numbers" (evaluateFormula "=PERCENTILE(A1:A20000,0.5)" sparseTextOnlyModel None) (CellValue.Error "NUM!")
checkV "quartile sparse no numbers" (evaluateFormula "=QUARTILE(A1:A20000,2)" sparseTextOnlyModel None) (CellValue.Error "NUM!")
approx "devsq"       "=DEVSQ(A1:A8)"       32.0
approx "avedev"      "=AVEDEV(2,4,6)"      (4.0/3.0)
let sparseAveDevModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 4.0) ]
            |> Map.ofList }
let sparseAveDev = evaluateFormula "=AVEDEV(A1:A20000)" sparseAveDevModel None
checkV "avedev sparse blanks" sparseAveDev (n 1.0)
let sparseStatsModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 2.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 8.0) ]
            |> Map.ofList }
checkV "devsq sparse blanks" (evaluateFormula "=DEVSQ(A1:A20000)" sparseStatsModel None) (n 18.0)
checkV "geomean sparse blanks" (evaluateFormula "=GEOMEAN(A1:A20000)" sparseStatsModel None) (n 4.0)
checkV "harmean sparse blanks" (evaluateFormula "=HARMEAN(A1:A20000)" sparseStatsModel None) (n 3.2)
checkV "sumsq sparse blanks" (evaluateFormula "=SUMSQ(A1:A20000)" sparseStatsModel None) (n 68.0)
checkV "gcd sparse blanks" (evaluateFormula "=GCD(A1:A20000)" sparseStatsModel None) (n 2.0)
checkV "lcm sparse blanks" (evaluateFormula "=LCM(A1:A20000)" sparseStatsModel None) (n 8.0)
approx "geomean"     "=GEOMEAN(1,2,4)"     2.0
approx "harmean"     "=HARMEAN(1,2,4)"     (3.0/(1.0+0.5+0.25))
check  "averagea"    "=AVERAGEA(D1:D3)"    (n 10.0)  // (10+0+20)/3
check  "maxa"        "=MAXA(D1:D3)"        (n 20.0)
check  "mina"        "=MINA(D1:D3)"        (n 0.0)   // text→0
let sparseAVariantModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 9999; Col = 0 }, cell (CellValue.String "text")
              { Row = 19999; Col = 0 }, cell (CellValue.Number 20.0) ]
            |> Map.ofList }
checkV "averagea sparse blanks" (evaluateFormula "=AVERAGEA(A1:A20000)" sparseAVariantModel None) (n 10.0)
checkV "maxa sparse blanks" (evaluateFormula "=MAXA(A1:A20000)" sparseAVariantModel None) (n 20.0)
checkV "mina sparse blanks" (evaluateFormula "=MINA(A1:A20000)" sparseAVariantModel None) (n 0.0)
approx "var.s"       "=VAR.S(2,4,4,4,5,5,7,9)" 4.571428571428571
approx "var.p"       "=VAR.P(2,4,4,4,5,5,7,9)" 4.0
approx "stdev.p"     "=STDEV.P(2,4,4,4,5,5,7,9)" 2.0

// ── date ──
check  "weeknum"     "=WEEKNUM(DATE(2020,1,1))"  (n 1.0)
check  "datedif-y"   "=DATEDIF(DATE(2010,1,1),DATE(2020,6,15),\"Y\")" (n 10.0)
check  "datedif-m"   "=DATEDIF(DATE(2020,1,15),DATE(2020,4,10),\"M\")" (n 2.0)
check  "datedif-d"   "=DATEDIF(DATE(2020,1,1),DATE(2020,1,11),\"D\")" (n 10.0)
check  "networkdays" "=NETWORKDAYS(DATE(2020,1,1),DATE(2020,1,7))" (n 5.0) // Wed..Tue, excl Sat/Sun
check  "workday"     "=WORKDAY(DATE(2020,1,1),5)" (n (43831.0+7.0)) // +5 workdays over a weekend = +7 cal days
let sparseHolidayModel =
    { createEmpty 20010 4 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 43833.0)
              { Row = 19999; Col = 0 }, cell (CellValue.Number 43837.0) ]
            |> Map.ofList }
checkV "networkdays sparse holidays" (evaluateFormula "=NETWORKDAYS(43831,43837,A1:A20000)" sparseHolidayModel None) (n 3.0)
checkV "workday sparse holidays" (evaluateFormula "=WORKDAY(43831,3,A1:A20000)" sparseHolidayModel None) (n 43838.0)
checkV "networkdays sparse holidays plus zero" (evaluateFormula "=NETWORKDAYS(43831,43837,A1:A20000+0)" sparseHolidayModel None) (n 3.0)
checkV "networkdays sparse holidays plus date" (evaluateFormula "=NETWORKDAYS(43831,43837,A1:A20000+43831)" sparseHolidayModel None) (n 4.0)
checkV "workday sparse holidays plus zero" (evaluateFormula "=WORKDAY(43831,3,A1:A20000+0)" sparseHolidayModel None) (n 43838.0)
checkV "workday sparse holidays plus date" (evaluateFormula "=WORKDAY(43831,3,A1:A20000+43831)" sparseHolidayModel None) (n 43836.0)

// ── lookup ──
check  "lookup"      "=LOOKUP(35,B1:B5,C1:C5)"   (s "c")   // 35→30 (B3)→C3="c"
check  "lookup-exact" "=LOOKUP(40,B1:B5,C1:C5)"  (s "d")
check  "lookup-nov"  "=LOOKUP(25,B1:B5)"         (n 20.0)  // no result vector → lookup vector
let sparseLookupModel =
    { createEmpty 20000 2 with
        Cells =
            [ { Row = 0; Col = 0 }, cell (CellValue.Number 10.0)
              { Row = 0; Col = 1 }, cell (CellValue.String "first")
              { Row = 19999; Col = 0 }, cell (CellValue.Number 20.0)
              { Row = 19999; Col = 1 }, cell (CellValue.String "last") ]
            |> Map.ofList }
checkV "match sparse exact miss" (evaluateFormula "=MATCH(99,A1:A20000,0)" sparseLookupModel None) (CellValue.Error "N/A")
checkV "xlookup sparse exact miss" (evaluateFormula "=XLOOKUP(99,A1:A20000,B1:B20000,\"missing\")" sparseLookupModel None) (s "missing")

// ── numbervalue / textsplit ──
check  "numbervalue" "=NUMBERVALUE(\"1.234,56\",\",\",\".\")" (n 1234.56)
let ts1 = arrOf "=TEXTSPLIT(\"a,b,c\",\",\")"
checkV "textsplit cols" ts1.[0].Length 3
checkV "textsplit val"  ts1.[0].[1] (s "b")
let ts2 = arrOf "=TEXTSPLIT(\"a,b;c,d\",\",\",\";\")"
checkV "textsplit 2d rows" ts2.Length 2
checkV "textsplit 2d v"    ts2.[1].[1] (s "d")
check  "textsplit scalar 2d" "=TEXTSPLIT(\"a,b;c,d\",\",\",\";\")" (s "a")
check  "textsplit scalar no col" "=TEXTSPLIT(\"a,b;c,d\",\"\",\";\")" (s "a,b")

printfn ""
printfn "Stage 11 stats/date/lookup tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
