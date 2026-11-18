#r "nuget: Fable.Core, 4.3.0"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.OfficeFormat.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Vb.dll"
#r "../../src/Spreadsheet/bin/Debug/netstandard2.0/Rbt.Spreadsheet.dll"

// What the showcase macros (tests/macros/showcase: Game of Life, Mandelbrot, mazes, sorting bars ...)
// needed of Excel's object model: pictures painted into cells past the grid's first columns, the
// window's gridlines and zoom, and the members they use.

open System
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

// ---- borders: one line between two cells (the maze generator reads them back) ----
let borders = run """
Sub Main()
    With Range("B2:D4")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
    End With
    Range("C3").Borders(xlEdgeTop).LineStyle = xlNone
    Range("A1").Value = Range("C2").Borders(xlEdgeBottom).LineStyle
    Range("A2").Value = Range("C4").Borders(xlEdgeTop).LineStyle
    Range("A3").Value = Range("C4").Borders(xlEdgeTop).Weight
    Range("F2:G3").BorderAround xlDashDot, xlMedium
    Range("A4").Value = Range("F2").Borders(xlEdgeLeft).LineStyle & "/" & Range("F2").Borders(xlEdgeRight).LineStyle
    Range("A5").Value = Range("F1").Borders(xlEdgeBottom).Weight
End Sub"""
let side (wb: Workbook) (r, c) (pick: CellFormat -> BorderStyle option) = getCell { Row = r; Col = c } (model wb) |> Option.bind (fun cell -> pick cell.Format) |> Option.map (fun b -> b.Style)
check "Borders on a block: every side, thick" (side borders (1, 1) (fun f -> f.BorderLeft), side borders (2, 2) (fun f -> f.BorderRight)) (Some "thick", Some "thick")
check "C3's top taken away is C2's bottom too" (side borders (2, 2) (fun f -> f.BorderTop), side borders (1, 2) (fun f -> f.BorderBottom), valueAt borders (0, 0)) (None, None, Number -4142.0)
check "a line read from the cell beyond" (valueAt borders (1, 0), valueAt borders (2, 0)) (Number 1.0, Number 4.0)
check "BorderAround: the outline only" (valueAt borders (3, 0), side borders (1, 5) (fun f -> f.BorderRight)) (String "4/-4142", None)
check "the line above F2 is F1's bottom" (valueAt borders (4, 0), side borders (1, 5) (fun f -> f.BorderTop)) (Number -4138.0, Some "mediumDashDot")

// ---- notes, ID, theme colours, center across selection, orientation ----
let misc = run """
Sub Main()
    Range("A1").AddComment "Start"
    Range("A1").Comment.Text "Round 2"
    Range("B1").ID = "ALIVE"
    Range("B2").Value = Range("B1").ID & "/" & Range("A1").Comment.Text & "/" & (Range("C1").Comment Is Nothing)
    Range("C2").Font.ThemeColor = xlThemeColorDark1
    Range("C3").Interior.ThemeColor = xlThemeColorAccent1
    Range("C4").Interior.Color = RGB(100, 100, 100)
    Range("C4").Interior.TintAndShade = 0.5
    Range("A3:G3").HorizontalAlignment = xlCenterAcrossSelection
    Range("D4").Orientation = xlUpward
    Range("D5").Value = Application.Text("March 2026", "mmmm yyyy")
    Range("D6").Value = VBA.RGB(10, 10, 2550)
    ActiveSheet.Protect
    Range("D7").Value = ActiveSheet.ProtectContents
End Sub"""
let formatOf (wb: Workbook) (r, c) = getCell { Row = r; Col = c } (model wb) |> Option.map (fun cell -> cell.Format) |> Option.defaultValue defaultFormat
check "a note, its text changed; ID; no note is Nothing" (valueAt misc (1, 1)) (String "ALIVE/Round 2/True")
check "theme colours (Background 1 is white)" ((formatOf misc (1, 2)).TextColor, (formatOf misc (2, 2)).BgColor) (Some "#FFFFFF", Some "#4472C4")
check "TintAndShade lightens" ((formatOf misc (3, 2)).BgColor) (Some "#B2B2B2")
check "center across selection, upward text" ((formatOf misc (2, 0)).HorizontalAlign, (formatOf misc (3, 3)).TextRotation) (Some "centerContinuous", 90)
check "Application.Text of a date typed as text (written, the text is a date again, as in Excel)" (valueAt misc (4, 3)) (Number 46082.0)
check "VBA.RGB (a component over 255 is 255)" (valueAt misc (5, 3)) (Number 16714250.0)
check "Protect, read back" (valueAt misc (6, 3)) (Boolean true)

// ---- whole sheet / columns: formats on the cells shown and the columns' style, default sizes ----
let whole = run """
Sub Main()
    Cells.Interior.Color = vbWhite
    Columns.ColumnWidth = 2
    Rows.RowHeight = 14
    Columns("B").Font.Color = vbRed
    Range("B500").Value = 7
End Sub"""
check "every cell shown gets the fill" ((formatOf whole (5, 5)).BgColor, (formatOf whole ((model whole).RowCount - 1, 0)).BgColor) (Some "#FFFFFF", Some "#FFFFFF")
check "sizes of every column and row: the defaults" ((model whole).DefaultColumnWidth, Math.Round((model whole).DefaultRowHeight, 2), (model whole).ColumnWidths.Count) (19.0, 18.67, 0)
check "a cell written later in a styled column takes its style" ((formatOf whole (499, 1)).TextColor, (formatOf whole (499, 1)).BgColor) (Some "#FF0000", Some "#FFFFFF")

// ---- conditional formats, Replace, DisplayFormat ----
let rules = run """
Sub Main()
    Range("A1:C1").Value = Array(0, 1, 2)
    Range("A1:C1").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1"
    Range("A1:C1").FormatConditions(1).Interior.Color = RGB(255, 0, 0)
    Range("A3:C3").Value = Array(1, 2, 3)
    Range("A3:C3").FormatConditions.AddColorScale ColorScaleType:=3
    Range("A3:C3").FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(0, 0, 255)
    Range("D1").Value = Range("A1:C1").FormatConditions.Count & "|" & Range("B1").DisplayFormat.Interior.Color & "|" & Range("A1").DisplayFormat.Interior.Color
    Range("A5:C5").Value = Array(10, 20, 0)
    Range("A5:C5").Replace What:="0", Replacement:="", LookAt:=xlPart
End Sub"""
check "a cell-value rule and what B1 shows with it" (valueAt rules (0, 3)) (String "1|255|16777215")
check "two rules on the sheet; the scale's first stop blue" ((model rules).ConditionalFormats.Length, (model rules).ConditionalFormats |> List.tryPick (fun e -> match e.Format.Rule with ColorScale stops -> Some (snd stops.Head) | _ -> None)) (2, Some "#0000FF")
check "Replace in the values (an empty result clears the cell)" (valueAt rules (4, 0), valueAt rules (4, 1), valueAt rules (4, 2)) (Number 1.0, Number 2.0, CellValue.Empty)

// ---- a chart from arrays (Barnsley fern, Archimedean spiral) ----
let charted = run """
Sub Main()
    Dim x(3) As Double, y(3) As Double, i
    For i = 0 To 3: x(i) = i: y(i) = i * i: Next
    Dim c As Chart
    Set c = ActiveSheet.Shapes.AddChart.Chart
    c.ChartType = xlXYScatter
    c.HasLegend = False
    c.SeriesCollection.NewSeries
    c.SeriesCollection.Item(1).XValues = x
    c.SeriesCollection(1).Values = y
    Range("A1").Value = ActiveSheet.Shapes.Count & "|" & c.SeriesCollection.Count
End Sub"""
let chartXml = (model charted).AuthoredCharts |> Map.toList |> List.map snd |> List.tryHead |> Option.defaultValue ""
check "one chart, one series" (valueAt charted (0, 0), (model charted).Drawings.Length) (String "1|1", 1)
check "an XY scatter of literal points" (chartXml.Contains "<c:scatterChart>", chartXml.Contains "<c:yVal><c:numLit>", chartXml.Contains "<c:v>9</c:v>", chartXml.Contains "<c:legend>") (true, true, true, false)

// ---- a sheet by its code name; a project the engine cannot load fails to compile ----
let coded = run """
Sub Main()
    Sheet1.Range("A1").Value = Sheet1.Name
End Sub"""
check "Sheet1 is the first sheet" (valueAt coded (0, 0)) (String coded.Sheets.[coded.ActiveSheet].Name)
let unloadable =
    match M.MacroProgram.Compile [ Rbt.Vb.Api.ModuleSource.standard "Module1" "Sub Show(a)\nEnd Sub\nSub Main()\n    Dim p(5) As String\n    Show p()\nEnd Sub" ] with
    | Result.Ok _ -> "compiled"
    | Result.Error failures -> if (List.head failures).Description.StartsWith "The VBA engine could not" then "a compile error" else (List.head failures).Description
// (Rbt.Vb 0.5.0 passes `p()` as the array; before, this was the engine's exception at load.)
check "an array passed with empty parentheses compiles" unloadable "compiled"

// ---- a class module exported from the VBA editor (VERSION / BEGIN header) ----
let exported = M.macroModule "Counter" "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1  'True\nEND\nAttribute VB_Name = \"Counter\"\nAttribute VB_PredeclaredId = False\nAttribute VB_Exposed = False\nPublic N As Long\n"
check "an exported class: its kind and name, the header skipped" (exported.Kind, exported.Name, exported.Source.Contains "VERSION") (Rbt.Vb.Api.Class, "Counter", false)

// ---- a workbook without formulas is not recalculated after each write; one written later calculates ----
let later = run """
Sub Main()
    Dim i
    For i = 1 To 50: Cells(i, 1).Value = Cells(i, 1).Value + i: Next
    Range("B1").Formula = "=SUM(A1:A50)"
    Range("A1").Value = 101
    Range("B2").Value = Range("B1").Value
    Range("B1").ClearContents
    Range("A2").Value = 7
    Range("B3").Formula = "=A2*2"
    Range("B4").Value = Range("B3").Value
End Sub"""
check "a formula written after values sees them, and what is written after it" (valueAt later (1, 1)) (Number 1375.0)
check "and again after the only formula went" (valueAt later (3, 1)) (Number 14.0)

printfn "Stage 63 macro showcase tests: %d passed, %d failed" pass fail
if fail > 0 then exit 1
