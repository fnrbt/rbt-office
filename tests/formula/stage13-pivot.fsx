#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Pivot-table engine: group-by + cross-tab + Grand Totals.

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Pivot
open RedBat.Spreadsheet.Formula

let S = CellValue.String
let N = CellValue.Number
let cell v = { Value = v; Format = defaultFormat; Comment = None }
// Region, Product, Sales
let source =
    [| [| S "Region"; S "Product"; S "Sales" |]
       [| S "East"; S "A"; N 100.0 |]
       [| S "East"; S "B"; N 200.0 |]
       [| S "West"; S "A"; N 150.0 |]
       [| S "West"; S "B"; N 250.0 |]
       [| S "East"; S "A"; N 50.0 |] |]

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

// ── group-by Region, Sum of Sales ──
let p1 = computePivot source 0 None 2 Sum
check "sum shape rows"   p1.Length 4                       // header + East + West + Grand
check "sum header0"      p1.[0].[0] (S "Region")
check "sum header1"      p1.[0].[1] (S "Sum of Sales")
check "sum east"         p1.[1] [| S "East"; N 350.0 |]    // 100+200+50
check "sum west"         p1.[2] [| S "West"; N 400.0 |]    // 150+250
check "sum grand"        p1.[3] [| S "Grand Total"; N 750.0 |]

// ── Count / Average ──
let pc = computePivot source 0 None 2 Count
check "count east"       pc.[1].[1] (N 3.0)
check "count grand"      pc.[3].[1] (N 5.0)
let pa = computePivot source 0 None 2 Average
check "avg west"         pa.[2].[1] (N 200.0)
check "avg grand"        pa.[3].[1] (N 150.0)

// ── cross-tab Region × Product, Sum of Sales ──
let p2 = computePivot source 0 (Some 1) 2 Sum
check "xtab header"      p2.[0] [| S "Row Labels"; S "A"; S "B"; S "Grand Total" |]
check "xtab east"        p2.[1] [| S "East"; N 150.0; N 200.0; N 350.0 |]   // A:100+50, B:200
check "xtab west"        p2.[2] [| S "West"; N 150.0; N 250.0; N 400.0 |]
check "xtab grand"       p2.[3] [| S "Grand Total"; N 300.0; N 450.0; N 750.0 |]

// ── Max ──
let pm = computePivot source 0 (Some 1) 2 Max
check "max east A"       pm.[1].[1] (N 100.0)   // max(100,50)

// ── nested rows (Region → Product) with subtotals, Sum of Sales ──
let pn = computePivotN source [0; 1] [ (2, Sum) ]
check "nested rows count" pn.Length 8           // hdr + 2 East + East Total + 2 West + West Total + Grand
check "nested header"     pn.[0] [| S "Region"; S "Product"; S "Sum of Sales" |]
check "nested east-A"     pn.[1] [| S "East"; S "A"; N 150.0 |]
check "nested east-B"     pn.[2] [| S "East"; S "B"; N 200.0 |]
check "nested east-total" pn.[3] [| S "East Total"; CellValue.Empty; N 350.0 |]
check "nested west-total" pn.[6] [| S "West Total"; CellValue.Empty; N 400.0 |]
check "nested grand"      pn.[7] [| S "Grand Total"; CellValue.Empty; N 750.0 |]

// ── nested rows + multiple value fields (Sum + Count) ──
let pn2 = computePivotN source [0; 1] [ (2, Sum); (2, Count) ]
check "multi header"      pn2.[0] [| S "Region"; S "Product"; S "Sum of Sales"; S "Count of Sales" |]
check "multi east-A"      pn2.[1] [| S "East"; S "A"; N 150.0; N 2.0 |]
check "multi grand"       pn2.[7] [| S "Grand Total"; CellValue.Empty; N 750.0; N 5.0 |]

// ── dialog resolution: pivotFromNames (the Insert-Pivot apply core) ──
// Header names → indices, agg parsing, nested vs cross-tab, error messages.
check "parseAgg sum"     (parseAgg "Sum") (Some Sum)
check "parseAgg avg"     (parseAgg "average") (Some Average)
check "parseAgg bad"     (parseAgg "median") None
let okGrid r = match r with Result.Ok (g: CellValue[][]) -> g | Result.Error e -> failwithf "pivot error: %s" e
let g1 = okGrid (pivotFromNames source [ "Region" ] None "Sales" Sum)
check "byNames grand"    g1.[g1.Length - 1] [| S "Grand Total"; N 750.0 |]
let g2 = okGrid (pivotFromNames source [ "Region"; "Product" ] None "Sales" Sum)
check "byNames nested rows" g2.Length 8
let g3 = okGrid (pivotFromNames source [ "Region" ] (Some "Product") "Sales" Sum)
check "byNames crosstab"  g3.[0] [| S "Row Labels"; S "A"; S "B"; S "Grand Total" |]
let errMsg r = match r with Result.Error (m: string) -> m | Result.Ok _ -> ""
check "byNames bad row"   ((errMsg (pivotFromNames source [ "Nope" ] None "Sales" Sum)).StartsWith "Row field not found") true
check "byNames bad value" ((errMsg (pivotFromNames source [ "Region" ] None "Nope" Sum)).StartsWith "Value field not found") true

// ── spilling GROUPBY / PIVOTBY (the "live formula pivot") ──
// matrix entries (data-only parallel ranges)
let dataRows = source.[1..]
let rowM = dataRows |> Array.map (fun r -> [| r.[0] |])   // Region
let colM = dataRows |> Array.map (fun r -> [| r.[1] |])   // Product
let valM = dataRows |> Array.map (fun r -> [| r.[2] |])   // Sales
let gb = groupBy rowM valM Sum
check "groupBy east"   gb.[0] [| S "East"; N 350.0 |]
check "groupBy west"   gb.[1] [| S "West"; N 400.0 |]
check "groupBy grand"  gb.[2] [| S "Grand Total"; N 750.0 |]
let pb = pivotBy rowM colM valM Sum
check "pivotBy header"  pb.[0] [| S "Row Labels"; S "A"; S "B"; S "Grand Total" |]
check "pivotBy east"    pb.[1] [| S "East"; N 150.0; N 200.0; N 350.0 |]

// formula path: GROUPBY over real cell ranges (data only, A1:A5 / C1:C5)
let model =
    { createEmpty 100 26 with
        Cells =
            ([ "East";"East";"West";"West";"East" ] |> List.mapi (fun i v -> { Row=i; Col=0 }, cell (S v)))
            @ ([ 100.;200.;150.;250.;50. ] |> List.mapi (fun i v -> { Row=i; Col=2 }, cell (N v)))
            |> Map.ofList }
let arrAt f = match parseFormula f with Ok e -> evalArray e (contextOf model) | Result.Error e -> failwithf "parse: %s" e
let gf = arrAt "=GROUPBY(A1:A5,C1:C5,\"sum\")"
check "GROUPBY formula east"  gf.[0] [| S "East"; N 350.0 |]
check "GROUPBY formula grand" gf.[gf.Length - 1] [| S "Grand Total"; N 750.0 |]
// composable: the spill is a normal range — feed it to other functions
check "GROUPBY is composable" (evaluateFormula "=SUM(GROUPBY(A1:A5,C1:C5,\"sum\"))" model None) (N 1500.0) // (350+400)+750

// ── dialog writes a LIVE formula (buildFormula) ──
let hdrs = source.[0]   // [Region; Product; Sales]
let fOf r = match r with Result.Ok (s: string) -> s | Result.Error e -> "ERR:" + e
check "formula groupby"  (fOf (buildFormula 0 0 7 2 hdrs [ "Region" ] None "Sales" "sum")) "=GROUPBY(A2:A8,C2:C8,\"sum\")"
check "formula nested"   (fOf (buildFormula 0 0 7 2 hdrs [ "Region"; "Product" ] None "Sales" "sum")) "=GROUPBY(A2:B8,C2:C8,\"sum\")"
check "formula pivotby"  (fOf (buildFormula 0 0 7 2 hdrs [ "Region" ] (Some "Product") "Sales" "sum")) "=PIVOTBY(A2:A8,B2:B8,C2:C8,\"sum\")"
check "formula bad row"  ((fOf (buildFormula 0 0 7 2 hdrs [ "Nope" ] None "Sales" "sum")).StartsWith "ERR:Row field not found") true

// end-to-end: the formula the dialog writes, evaluated + spilled into a model
let srcModel =
    { createEmpty 100 26 with
        Cells =
            ([ S "Region" ] @ ([ "East";"East";"West";"West";"East" ] |> List.map S) |> List.mapi (fun i v -> { Row=i; Col=0 }, cell v))
            @ ([ S "Sales" ] @ ([ 100.;200.;150.;250.;50. ] |> List.map N) |> List.mapi (fun i v -> { Row=i; Col=2 }, cell v))
            |> Map.ofList }
let formula = fOf (buildFormula 0 0 5 2 [| S "Region"; S "x"; S "Sales" |] [ "Region" ] None "Sales" "sum")
let placed = srcModel |> setCell { Row=0; Col=4 } (cell (CellValue.Formula(formula, CellValue.Empty)))
let spilled = placed |> recalculateAllFormulas None None |> spillArrays None None
check "dialog formula" formula "=GROUPBY(A2:A6,C2:C6,\"sum\")"
check "spilled east"  (match getCell { Row=0; Col=4 } spilled with Some c -> (match c.Value with CellValue.Formula(_,v) -> v | v -> v) | None -> CellValue.Empty) (S "East")
check "spilled sum"   (getCellValue { Row=0; Col=5 } spilled) (N 350.0)   // East sum spilled to F1
check "spilled grand" (getCellValue { Row=2; Col=5 } spilled) (N 750.0)   // Grand Total

// ── interactive spill: edit source → auto-refresh + clear stale cells ──
let baseCells =
    ([ "East";"East";"West";"West";"East" ] |> List.mapi (fun i v -> ({ Row=i; Col=0 }, cell (S v))))
    @ ([ 100.;200.;150.;250.;50. ] |> List.mapi (fun i v -> ({ Row=i; Col=2 }, cell (N v))))
    @ [ { Row=0; Col=4 }, cell (CellValue.Formula("=GROUPBY(A1:A5,C1:C5,\"sum\")", CellValue.Empty)) ]
let m0 = { createEmpty 100 26 with Cells = Map.ofList baseCells } |> recalcAndSpill None None
let valAt r c (m: SpreadsheetModel) = match getCellValue { Row=r; Col=c } m with CellValue.Formula(_, v) -> v | v -> v
check "ispill anchor"   (valAt 0 4 m0) (S "East")        // E1 anchor (formula) shows top-left
check "ispill spilled"  (getCellValue { Row=0; Col=5 } m0) (N 350.0)  // F1 = East sum
check "ispill grand"    (getCellValue { Row=2; Col=5 } m0) (N 750.0)  // F3 = Grand Total
check "ispill tracked"  (Map.containsKey { Row=2; Col=5 } m0.Spill) true
// edit a source value → the pivot refreshes (no Refresh needed)
let m1 = m0 |> setCell { Row=0; Col=2 } (cell (N 1100.0)) |> recalcAndSpill None None
check "ispill live F1"   (getCellValue { Row=0; Col=5 } m1) (N 1350.0) // East now 1100+200+50
check "ispill live grand" (getCellValue { Row=2; Col=5 } m1) (N 1750.0)
// shrink the result (all rows → East): the old 3rd row must be cleared
let m2 = m0 |> setCell { Row=2; Col=0 } (cell (S "East")) |> setCell { Row=3; Col=0 } (cell (S "East")) |> recalcAndSpill None None
check "ispill shrink stale-cleared" (getCellValue { Row=2; Col=4 } m2) CellValue.Empty // old E3 gone
check "ispill shrink grand"        (getCellValue { Row=1; Col=5 } m2) (N 750.0)        // Grand now at row 2

// ── nested COLUMN fields (multi-level column header) ──
let nrow = [| [| S "East" |]; [| S "East" |]; [| S "East" |]; [| S "West" |]; [| S "West" |] |]
let ncol = [| [| S "Q1"; S "A" |]; [| S "Q1"; S "B" |]; [| S "Q2"; S "A" |]; [| S "Q1"; S "A" |]; [| S "Q2"; S "B" |] |]
let nval = [| [| N 10. |]; [| N 20. |]; [| N 30. |]; [| N 40. |]; [| N 50. |] |]
let ncx = pivotBy nrow ncol nval Sum
check "ncol header lvl0" ncx.[0] [| S "Row Labels"; S "Q1"; S "Q1"; S "Q2"; S "Q2"; S "Grand Total" |]
check "ncol header lvl1" ncx.[1] [| CellValue.Empty; S "A"; S "B"; S "A"; S "B"; CellValue.Empty |]
check "ncol east row"    ncx.[2] [| S "East"; N 10.; N 20.; N 30.; N 0.; N 60. |]
check "ncol grand row"   ncx.[ncx.Length - 1] [| S "Grand Total"; N 50.; N 20.; N 30.; N 50.; N 150. |]

printfn ""
printfn "Stage 13 pivot tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
