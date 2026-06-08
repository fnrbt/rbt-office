#r "/home/dave/Dev/browser/office/src/Spreadsheet/bin/Debug/netstandard2.0/RedBat.Spreadsheet.dll"

// VBA runs natively (cross-platform) AND writes cells back via runOnModel.

open RedBat.Spreadsheet.Model
module V = RedBat.Spreadsheet.Vba

let cell v = { Value = v; Format = defaultFormat; Comment = None }
let model =
    { createEmpty 20 8 with
        Cells = Map.ofList [ { Row=0; Col=0 }, cell (CellValue.Number 10.) ] }

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-26s expected %A got %A" label expected got

let disp m r c = displayValue (getCellValue { Row=r; Col=c } m)

// A macro that reads A1, writes a computed value, a string, and a formula.
let src =
    String.concat "\n"
        [ "Range(\"B1\").Value = Range(\"A1\").Value * 2"
          "Cells(1, 3).Value = \"hello\""
          "Range(\"D1\").Formula = \"=B1+5\"" ]

match V.runOnModel model src with
| Result.Error e -> fail <- fail + 1; printfn "FAIL macro errored: %s" e
| Result.Ok m ->
    check "B1 = A1*2"        (disp m 0 1) "20"
    check "C1 = string"      (disp m 0 2) "hello"
    check "D1 formula compute" (disp m 0 3) "25"   // =B1+5 = 20+5
    check "A1 untouched"     (disp m 0 0) "10"

// Loop macro: fill A3:A5 with 1,2,3.
let src2 =
    String.concat "\n"
        [ "Dim i"
          "For i = 1 To 3"
          "  Cells(i + 2, 1).Value = i"
          "Next i" ]
match V.runOnModel model src2 with
| Result.Error e -> fail <- fail + 1; printfn "FAIL loop macro errored: %s" e
| Result.Ok m ->
    check "loop A3" (disp m 2 0) "1"
    check "loop A4" (disp m 3 0) "2"
    check "loop A5" (disp m 4 0) "3"

// --- events: hasSub + fireEvent ---
let openModel = { model with Macros = "Sub Workbook_Open()\n  Range(\"A1\").Value = 99\nEnd Sub" }
check "hasSub true"  (V.hasSub "Workbook_Open" openModel.Macros) true
check "hasSub false" (V.hasSub "Worksheet_Change" openModel.Macros) false
check "Workbook_Open fires" (disp (V.fireEvent "Workbook_Open" "" openModel) 0 0) "99"
check "no-sub no-op"        (disp (V.fireEvent "Workbook_Open" "" model) 0 0) "10"

// Worksheet_Change receiving the changed cell as Target.
let wcModel = { model with Macros = "Sub Worksheet_Change(ByVal Target As Range)\n  Range(\"Z1\").Value = Target.Address\nEnd Sub" }
check "Worksheet_Change Target" (disp (V.fireEvent "Worksheet_Change" "Range(\"B2\")" wcModel) 0 25) "$B$2"

printfn ""
printfn "Stage 19 VBA tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
