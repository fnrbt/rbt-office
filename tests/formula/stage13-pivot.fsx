#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// Pivot-table engine: group-by + cross-tab + Grand Totals.

open RedBat.Spreadsheet.Model
open RedBat.Spreadsheet.Pivot

let S = CellValue.String
let N = CellValue.Number
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

printfn ""
printfn "Stage 13 pivot tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
