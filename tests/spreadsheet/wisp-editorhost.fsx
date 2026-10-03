#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeUi.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Grid.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.OfficeWisp.dll"
#r "../../../wisp/bin/Debug/net10.0/Wisp.dll"
#r "../../../rbt-renderer/bin/Debug/net10.0/Renderer.dll"
#r "../../../wisp/bin/Debug/net10.0/Platform.dll"

#load "../../spreadsheet-wisp/GridRender.fs"
#load "../../spreadsheet-wisp/EditorHost.fs"

open Rbt.Spreadsheet.Model
open Rbt.SpreadsheetWisp.EditorHost

let fail name detail =
    failwithf "%s failed: %s" name detail

let assertEq name expected actual =
    if expected <> actual then fail name (sprintf "expected %A, got %A" expected actual)

let st = create [| createEmpty 20 10 |] [| "Sheet1" |] 0

applyCellEdit st 0 0 "2" |> ignore
applyCellEdit st 0 1 "=A1*5" |> ignore

match getCellValue { Row = 0; Col = 1 } (activeModel st) with
| Formula ("A1*5", Number 10.0) -> ()
| other -> fail "cell commit with formula" (sprintf "%A" other)

if not (st.LastRecalcWasIncremental || st.CalcEngine.Value.IsSome) then
    fail "incremental recalc" "single-cell formula edit did not use/update the incremental engine"

handleKey st { Key = "ArrowRight"; Shift = false; Ctrl = false; Alt = false } |> ignore
assertEq "arrow navigation" (0, 1) (focusCell (activeModel st))

handleKey st { Key = "Delete"; Shift = false; Ctrl = false; Alt = false } |> ignore
assertEq "delete-selection" Empty (getCellValue { Row = 0; Col = 1 } (activeModel st))

undo st |> ignore
match getCellValue { Row = 0; Col = 1 } (activeModel st) with
| Formula ("A1*5", Number 10.0) -> ()
| other -> fail "undo restores" (sprintf "%A" other)

redo st |> ignore
assertEq "redo delete" Empty (getCellValue { Row = 0; Col = 1 } (activeModel st))

printfn "wisp-editorhost smoke ok"
