' build-border-fixture.vbs
'
' Builds tests/test-documents-excel/synthetic-diagonal-borders.xlsx covering
' border cases the corpus does not yet exercise: diagonal-down / diagonal-up /
' both diagonals, double and thick box borders, and a bordered merged range.
' Visual parity is judged from an Excel screenshot vs our render.
'
' Usage:
'   cscript //nologo build-border-fixture.vbs <unused-input> <output.xlsx>

Option Explicit

Const xlOpenXMLWorkbook = 51
Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlContinuous = 1
Const xlDouble = -4119
Const xlThin = 2
Const xlMedium = -4138
Const xlThick = 4

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo build-border-fixture.vbs <input> <output.xlsx>"
  WScript.Quit 2
End If
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, wb, sh
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "Borders"

Sub Edge(rng, idx, style, weight)
  rng.Borders(idx).LineStyle = style
  rng.Borders(idx).Weight = weight
End Sub

' Labels in column A.
sh.Cells(2, 1).Value = "diag down"
sh.Cells(4, 1).Value = "diag up"
sh.Cells(6, 1).Value = "both diag"
sh.Cells(2, 4).Value = "thick box"
sh.Cells(4, 4).Value = "double box"
sh.Cells(6, 4).Value = "merged box"

' B2 diagonal down (thin), with text so we see text-under-diagonal.
sh.Cells(2, 2).Value = "x"
Edge sh.Cells(2, 2), xlDiagonalDown, xlContinuous, xlThin

' B4 diagonal up (thin).
sh.Cells(4, 2).Value = "y"
Edge sh.Cells(4, 2), xlDiagonalUp, xlContinuous, xlThin

' B6 both diagonals (medium).
Edge sh.Cells(6, 2), xlDiagonalDown, xlContinuous, xlMedium
Edge sh.Cells(6, 2), xlDiagonalUp, xlContinuous, xlMedium

' E2 thick box (all four edges).
Edge sh.Cells(2, 5), xlEdgeLeft, xlContinuous, xlThick
Edge sh.Cells(2, 5), xlEdgeTop, xlContinuous, xlThick
Edge sh.Cells(2, 5), xlEdgeRight, xlContinuous, xlThick
Edge sh.Cells(2, 5), xlEdgeBottom, xlContinuous, xlThick

' E4 double box.
Edge sh.Cells(4, 5), xlEdgeLeft, xlDouble, xlThick
Edge sh.Cells(4, 5), xlEdgeTop, xlDouble, xlThick
Edge sh.Cells(4, 5), xlEdgeRight, xlDouble, xlThick
Edge sh.Cells(4, 5), xlEdgeBottom, xlDouble, xlThick

' E6:F7 merged box (thick edges around the whole merge, empty content).
Dim mrng
Set mrng = sh.Range("E6:F7")
mrng.Merge
Edge mrng, xlEdgeLeft, xlContinuous, xlThick
Edge mrng, xlEdgeTop, xlContinuous, xlThick
Edge mrng, xlEdgeRight, xlContinuous, xlThick
Edge mrng, xlEdgeBottom, xlContinuous, xlThick

wb.SaveAs outputPath, xlOpenXMLWorkbook
wb.Close False
excel.Quit
Set excel = Nothing
