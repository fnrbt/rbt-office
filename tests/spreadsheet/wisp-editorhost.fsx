#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeUi.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Grid.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Office.Formats/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.OfficeWisp.dll"
// Resolve packages from the library build: dotnet fsi uses the SDK's Core,
// while the executable output carries Wisp's newer FSharp.Core runtime.
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Wisp.dll"
#r "../../office-wisp/bin/Debug/net10.0/Rbt.Renderer.dll"
#r "../../office-wisp/bin/Debug/net10.0/Platform.dll"

#load "../../spreadsheet-wisp/GridRender.fs"
#load "../../spreadsheet-wisp/EditorHost.fs"

open Rbt.Spreadsheet.Model
open Rbt.SpreadsheetWisp.EditorHost

let fail name detail =
    failwithf "%s failed: %s" name detail

let mutable checks = 0
let assertEq name expected actual =
    checks <- checks + 1
    if expected <> actual then fail name (sprintf "expected %A, got %A" expected actual)

let st = create [| createEmpty 20 10 |] [| "Sheet1" |] 0

applyCellEdit st 0 0 "2" |> ignore
applyCellEdit st 0 1 "=A1*5" |> ignore

match getCellValue { Row = 0; Col = 1 } (activeModel st) with
| Formula ("A1*5", Number 10.0) -> ()
| other -> fail "cell commit with formula" (sprintf "%A" other)


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

let value st r c = getCellValue { Row = r; Col = c } (activeModel st)
let cached st r c = value st r c |> unwrapFormulaValue
let ok = function Result.Ok () -> () | Result.Error message -> failwith message
let key name = { Key = name; Shift = false; Ctrl = false; Alt = false }

let chain = create [| createEmpty 20 10; createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Middle"; "Result" |] 0
applyCellEdit chain 0 0 "4" |> ignore
switchSheet chain 1
applyCellEdit chain 0 0 "=Data!A1*2" |> ignore
switchSheet chain 2
applyCellEdit chain 0 0 "=Middle!A1+1" |> ignore
assertEq "cross-sheet chain" (Number 9.0) (cached chain 0 0)
switchSheet chain 0
handleKey chain (key "Delete") |> ignore
switchSheet chain 2
assertEq "delete precedent across sheets" (Number 1.0) (cached chain 0 0)
undo chain |> ignore
switchSheet chain 2
assertEq "undo precedent deletion" (Number 9.0) (cached chain 0 0)
redo chain |> ignore
switchSheet chain 2
assertEq "redo precedent deletion" (Number 1.0) (cached chain 0 0)

let ranged = create [| createEmpty 20 10 |] [| "Sheet1" |] 0
applyCellEdit ranged 0 0 "3" |> ignore
applyCellEdit ranged 1 0 "7" |> ignore
applyCellEdit ranged 0 1 "=SUM(A1:A2)" |> ignore
applyNoHistory ranged (fun m -> { m with Selection = Range(0,0,1,0) }) |> ignore
handleKey ranged (key "Delete") |> ignore
assertEq "range delete recalculates" (Number 0.0) (cached ranged 0 1)
undo ranged |> ignore
assertEq "range delete undo" (Number 10.0) (cached ranged 0 1)

let moved = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Other" |] 0
applyCellEdit moved 0 0 "8" |> ignore
switchSheet moved 1
applyCellEdit moved 0 1 "=Data!A1*2" |> ignore
switchSheet moved 0
setFocus moved 0 0
cutSelection moved |> ignore
switchSheet moved 1
setFocus moved 0 0
pasteFromClipboard moved |> ignore
assertEq "cut paste dependent" (Number 16.0) (cached moved 0 1)
undo moved |> ignore
assertEq "undo paste destination" Empty (value moved 0 0)
undo moved |> ignore
assertEq "undo cut source" (Number 8.0) (value moved 0 0)
redo moved |> ignore
redo moved |> ignore
assertEq "redo cut paste" (Number 16.0) (cached moved 0 1)

let spills = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Spills" |] 0
applyCellEdit spills 0 0 "2" |> ignore
switchSheet spills 1
applyCellEdit spills 0 0 "=SEQUENCE(Data!A1)" |> ignore
assertEq "spill initially" (Number 2.0) (cached spills 1 0)
switchSheet spills 0
applyCellEdit spills 0 0 "3" |> ignore
switchSheet spills 1
assertEq "spill refresh" (Number 3.0) (cached spills 2 0)

module W = Rbt.Spreadsheet.Workbook
module B = Rbt.Spreadsheet.XlsxBridge
let book = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Local" |] 0
applyCellEdit book 1 0 "12" |> ignore
applyChange book (createTable "Sales" 3 0 5 1 false) |> ignore
let initial = book.Workbook.Value
let localId = initial.SheetOrder.[1]
let info : W.DefinedNameInfo = { Name = "LocalName"; Comment = Some "note"; Hidden = true }
let enriched =
    { initial with
        Names = Map.ofList [ (None, "GLOBAL"), "Data!$A$2"; (Some localId, "LOCALNAME"), "Data!$A$2" ]
        NameInfo = Map.ofList [ (Some localId, "LOCALNAME"), info ]
        Iteration = Some { MaxIterations = 40; MaxChange = 0.01 }
        Tables = initial.Tables |> Map.map (fun _ t -> { t with Columns = [ "Declared"; "Other" ] }) }
replaceWorkbook book enriched
applyChange book (createTable "Extra" 8 0 10 1 false) |> ignore
assertEq "declared table columns retained" [ "Declared"; "Other" ] book.Workbook.Value.Tables.["Sales"].Columns
assertEq "new table registered" true (Map.containsKey "Extra" book.Workbook.Value.Tables)
applyChange book (fun m -> { m with Tables = m.Tables |> List.filter (fun t -> t.Name <> "Extra") }) |> ignore
assertEq "removed table unregistered" false (Map.containsKey "Extra" book.Workbook.Value.Tables)
applyStructureEdit book (InsertRows(0, 1)) |> ignore
assertEq "global name follows structure" "Data!$A$3" book.Workbook.Value.Names.[None, "GLOBAL"]
assertEq "local name follows structure" "Data!$A$3" book.Workbook.Value.Names.[Some localId, "LOCALNAME"]
undo book |> ignore
assertEq "undo restores names" "Data!$A$2" book.Workbook.Value.Names.[None, "GLOBAL"]
redo book |> ignore
switchSheet book 1
renameSheet book "Renamed" |> ok
addSheet book "Added" |> ok
let addedId = book.Workbook.Value.ActiveSheet
deleteSheet book |> ok
undo book |> ignore
assertEq "undo delete restores identity" addedId book.Workbook.Value.ActiveSheet
undo book |> ignore
assertEq "undo add restores active" localId book.Workbook.Value.ActiveSheet
undo book |> ignore
assertEq "undo rename" "Local" (activeSheetName book)
redo book |> ignore
redo book |> ignore
redo book |> ignore
assertEq "redo deletion order" [ "Data"; "Renamed" ] (book.Workbook.Value.SheetOrder |> List.map (fun id -> book.Workbook.Value.Sheets.[id].Name))
let bytes = XlsxWriter.writeWorkbookPackageBytes (exportWorkbook book None)
let imported = bytes |> XlsxReader.loadBytesNative |> Option.get |> B.Import.importWorkbook
assertEq "saved iteration" book.Workbook.Value.Iteration imported.Context.Iteration
assertEq "saved global name" "Data!$A$3" imported.Context.Names.[None, "GLOBAL"]
assertEq "saved scoped name" "Data!$A$3" imported.Context.Names.[Some "sheet2", "LOCALNAME"]
assertEq "saved name metadata" info imported.Context.NameInfo.[Some "sheet2", "LOCALNAME"]
assertEq "saved table" true (Map.containsKey "Sales" imported.Context.Tables)

let reopened = createWithWorkbook imported.Context
switchSheet reopened 0
deleteSheet reopened |> ok
let removed = reopened.Workbook.Value
assertEq "surviving stable id" "sheet2" removed.ActiveSheet
let savedAgain = XlsxWriter.writeWorkbookPackageBytes (exportWorkbook reopened (Some imported.Source))
let reopenedAgain = savedAgain |> XlsxReader.loadBytesNative |> Option.get |> B.Import.importWorkbook
assertEq "local scope follows sheet deletion" true (Map.containsKey (Some "sheet1", "LOCALNAME") reopenedAgain.Context.Names)
assertEq "local metadata follows scope" info reopenedAgain.Context.NameInfo.[Some "sheet1", "LOCALNAME"]
undo reopened |> ignore
assertEq "undo restores sheet identity" [ "sheet1"; "sheet2" ] reopened.Workbook.Value.SheetOrder
redo reopened |> ignore
assertEq "redo preserves survivor identity" "sheet2" reopened.Workbook.Value.ActiveSheet

let references = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Results" |] 0
applyCellEdit references 0 0 "5" |> ignore
switchSheet references 1
applyCellEdit references 0 0 "=Data!A1*3" |> ignore
switchSheet references 0
renameSheet references "New Data" |> ok
switchSheet references 1
assertEq "rename rewrites dependent" (Formula("'New Data'!A1*3", Number 15.0)) (value references 0 0)
let savedReferences = exportWorkbook references None |> XlsxWriter.writeWorkbookPackageBytes |> XlsxReader.loadBytesNative |> Option.get
let savedResults = savedReferences |> B.Import.importWorkbook |> _.Context |> createWithWorkbook
switchSheet savedResults 1
assertEq "saved formula cache after rename" (Number 15.0) (cached savedResults 0 0)
switchSheet references 0
deleteSheet references |> ok
assertEq "deletion invalidates dependent" (Error "REF!") (cached references 0 0)
undo references |> ignore
switchSheet references 1
assertEq "undo deletion restores dependent" (Number 15.0) (cached references 0 0)
redo references |> ignore
assertEq "redo deletion invalidates dependent" (Error "REF!") (cached references 0 0)
addSheet references "Fresh" |> ok
assertEq "deleted identity not reused" false (references.Workbook.Value.ActiveSheet = "sheet1")
assertEq "duplicate sheet name rejected" true (renameSheet references "Results" |> Result.isError)
assertEq "invalid sheet name rejected" true (addSheet references "Bad/Name" |> Result.isError)

let events = create [| createEmpty 20 10; createEmpty 20 10 |] [| "Data"; "Results" |] 0
applyCellEdit events 0 0 "2" |> ignore
applyCellEdit events 0 1 "=A1*5" |> ignore
switchSheet events 1
applyCellEdit events 0 0 "=Data!C1*2" |> ignore
switchSheet events 0
applyChange events (fun m -> { m with Macros = "Sub Worksheet_Change(ByVal Target As Range)\nRange(\"C1\").Value = Range(\"B1\").Value\nEnd Sub" }) |> ignore
applyCellEdit events 0 0 "4" |> ignore
assertEq "event sees recalculated dependent" (Number 20.0) (cached events 0 2)
switchSheet events 1
assertEq "event writes recalculate foreign dependent" (Number 40.0) (cached events 0 0)
undo events |> ignore
assertEq "event and edit undo together" (Number 2.0) (cached events 0 0)
assertEq "event output undone" Empty (value events 0 2)
copySelection events |> ignore
startEditing events 0 0 (Some "99")
let replacement = W.createWorkbook () |> W.withDate1904 true
replaceWorkbook events replacement
assertEq "replacement cancels pending entry" false (commitEdit events)
assertEq "replacement clears clipboard" false (pasteFromClipboard events)
addSheet events "New" |> ok
assertEq "new sheet keeps workbook date system" true (activeModel events).Date1904

printfn "wisp-editorhost: %d assertions passed" checks
