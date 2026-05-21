' build-rotation-sweep.vbs
'
' Builds tests/test-documents-excel/synthetic-rotation-sweep.xlsx: a grid that
' sweeps cell-text rotation from -90..+90 in 15-degree steps (Excel
' .Orientation; +CCW, -CW) across columns, for three text lengths (rows). Used
' to reverse-engineer Excel's angle-dependent HORIZONTAL anchor for rotated
' text by screenshotting and measuring each cell's text position.
'
' Usage: cscript //nologo build-rotation-sweep.vbs <unused-input> <output.xlsx>

Option Explicit
Const xlOpenXMLWorkbook = 51

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: build-rotation-sweep.vbs <input> <output.xlsx>"
  WScript.Quit 2
End If
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, wb, sh
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "RotSweep"

' 13 orientations: -90..90 step 15.
Dim oris, i
oris = Array(-90,-75,-60,-45,-30,-15,0,15,30,45,60,75,90)

' Three text lengths on three (tall) rows.
Dim texts, rows, t
texts = Array("X", "Hi", "Sample")
rows = Array(1, 3, 5)

' Uniform wide columns so the horizontal anchor is visible; tall rows so
' rotated text fits without auto-resizing column widths.
Dim c
For c = 1 To UBound(oris) + 1
  sh.Columns(c).ColumnWidth = 12
Next

For t = 0 To UBound(texts)
  Dim rr
  rr = rows(t)
  sh.Rows(rr).RowHeight = 96
  For i = 0 To UBound(oris)
    Dim cell
    Set cell = sh.Cells(rr, i + 1)
    cell.Value = texts(t)
    cell.Orientation = oris(i)
  Next
Next

wb.SaveAs outputPath, xlOpenXMLWorkbook
wb.Close False
excel.Quit
Set excel = Nothing
