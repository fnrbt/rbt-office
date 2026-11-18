#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// What the showcase macros (tests/macros/showcase: Game of Life, Mandelbrot, mazes, sorting bars ...)
// needed of Excel's object model: pictures painted into cells past the grid's first columns, the
// window's gridlines and zoom, and the members they use.

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
module M = Rbt.Spreadsheet.MacroHost

let mutable pass = 0
let mutable fail = 0
let check label (got: 'a) (expected: 'a) =
    if got = expected then pass <- pass + 1
    else fail <- fail + 1; printfn "FAIL %-50s expected %A got %A" label expected got

let output = ResizeArray<string>()
let hooks = { M.defaultHooks with Output = output.Add; FileName = "Book.xlsm" }
let program (source: string) =
    match M.MacroProgram.Compile [ Rbt.Vb.Api.ModuleSource.standard "Module1" source ] with
    | Result.Ok compiled -> compiled
    | Result.Error failures -> failwithf "compile: %s" (failures |> List.map Rbt.Vb.Api.VbFailure.format |> String.concat "; ")
let run source =
    match (program source).Run(createWorkbook (), hooks, "Main") with
    | M.MacroDone (wb, _) -> wb
    | M.MacroFailed failure -> failwithf "macro failed: %s" (Rbt.Vb.Api.VbFailure.format failure)
let tryRun source =
    match (program source).Run(createWorkbook (), hooks, "Main") with
    | M.MacroDone _ -> "done"
    | M.MacroFailed failure -> sprintf "%d %s" failure.Number failure.Description
let model (wb: Workbook) = wb.Sheets.[wb.ActiveSheet].Model
let valueAt (wb: Workbook) (r, c) = getCellValue { Row = r; Col = c } (model wb) |> unwrapFormulaValue
let fillAt (wb: Workbook) (r, c) = getCell { Row = r; Col = c } (model wb) |> Option.bind (fun cell -> cell.Format.BgColor)

// ---- a picture painted past the grid's columns grows the sheet ----
let painted = run """
Sub Main()
    Columns("A:CZ").ColumnWidth = 1
    Rows("1:120").RowHeight = 6
    Cells(150, 104).Interior.Color = RGB(255, 0, 0)
    Range("AB3").Value = 1
End Sub"""
check "fill past the columns: shown" ((model painted).ColCount >= 104, (model painted).RowCount >= 150) (true, true)
check "the fill" (fillAt painted (149, 103)) (Some "#FF0000")
check "a column width of 1 (12 px)" ((model painted).ColumnWidths |> Map.tryFind 103) (Some 12.0)
check "row heights of 6 pt (8 px)" ((model painted).RowHeights |> Map.tryFind 119) (Some 8.0)

// ---- the window ----
let window = run """
Sub Main()
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.Zoom = 75
    Range("A1").Value = ActiveWindow.Zoom
    Range("B3").Select
    ActiveWindow.FreezePanes = True
    Application.ActiveWindow.ScrollRow = 1
End Sub"""
check "gridlines and headings off" ((model window).ShowGridLines, (model window).ShowHeaders) (false, false)
check "zoom" ((model window).Zoom, valueAt window (0, 0)) (0.75, Number 75.0)
check "panes frozen above and left of the active cell" ((model window).FrozenRows, (model window).FrozenCols) (2, 1)

printfn "Stage 63 macro showcase tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
