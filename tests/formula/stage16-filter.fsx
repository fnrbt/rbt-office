#r "/home/dave/.nuget/packages/fable.core/4.0.0/lib/netstandard2.0/Fable.Core.dll"
#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// AutoFilter model: unique values for the checklist, apply, and visible rows.

open RedBat.Spreadsheet.Model

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let S = CellValue.String
// Col 0 (rows 0-4): East, West, East, North, West ; Col 1 = sales
let model =
    { createEmpty 20 8 with
        Cells =
            ([ "East";"West";"East";"North";"West" ] |> List.mapi (fun i v -> { Row=i; Col=0 }, cell (S v)))
            @ ([ 10.;20.;30.;40.;50. ] |> List.mapi (fun i v -> { Row=i; Col=1 }, cell (CellValue.Number v)))
            |> Map.ofList }

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

check "unique values"   (columnUniqueValues 0 model) [ "East"; "North"; "West" ]
check "unique col1 count" (List.length (columnUniqueValues 1 model)) 5

// no filter → all rows visible
check "no filter all visible" (getVisibleRows model |> Set.count) 5

// filter col0 to {East} → rows 0 and 2
let fEast = applyFilter 0 (Set.ofList [ "East" ]) model
check "filter East rows"  (getVisibleRows fEast) (Set.ofList [ 0; 2 ])

// filter col0 to {West, North} → rows 1,3,4
let fWN = applyFilter 0 (Set.ofList [ "West"; "North" ]) model
check "filter West+North"  (getVisibleRows fWN) (Set.ofList [ 1; 3; 4 ])

// AND across two columns: East AND sales in {30}
let fAnd = model |> applyFilter 0 (Set.ofList [ "East" ]) |> applyFilter 1 (Set.ofList [ "30" ])
check "filter AND two cols" (getVisibleRows fAnd) (Set.ofList [ 2 ])

// remove filter restores
check "remove filter"     (getVisibleRows (removeFilter 0 fEast) |> Set.count) 5

printfn ""
printfn "Stage 16 filter tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
