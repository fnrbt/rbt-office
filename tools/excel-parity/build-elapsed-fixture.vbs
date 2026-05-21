' build-elapsed-fixture.vbs
'
' Builds tests/test-documents-excel/synthetic-elapsed-time.xlsx: a grid of
' time serials (column A) formatted with elapsed-time number formats
' ([h]/[m]/[s] and variants) across columns B.. so the parity harness has a
' real fixture covering Excel's elapsed-duration display. Column headers in
' row 1 carry the format string; column A row 1 is blank.
'
' Usage:
'   cscript //nologo build-elapsed-fixture.vbs <unused-input> <output.xlsx>

Option Explicit

Const xlOpenXMLWorkbook = 51

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo build-elapsed-fixture.vbs <input> <output.xlsx>"
  WScript.Quit 2
End If

outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, workbook, sheet
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)
sheet.Name = "Elapsed"

Dim serials
serials = Array( _
  0.5, _
  0.041666666666667, _
  1.0625, _
  2.5208333333333, _
  0.0006944444444, _
  0.0001157407407, _
  0.052083333333, _
  1.5212615740741, _
  0.00037615740741, _
  0.043831018518519 )

Dim fmts
fmts = Array( _
  "[h]:mm:ss", _
  "[h]:mm", _
  "[hh]:mm:ss", _
  "[h]", _
  "[m]:ss", _
  "[m]", _
  "[mm]:ss", _
  "[s]", _
  "[ss]", _
  "[s].000", _
  "[h]:mm:ss.0", _
  "h:mm:ss" )

Dim i, j, headerCell, valCell
' Header row: each format label as text in row 1.
For j = 0 To UBound(fmts)
  Set headerCell = sheet.Cells(1, j + 2)
  headerCell.NumberFormat = "@"
  headerCell.Value = fmts(j)
Next

For i = 0 To UBound(serials)
  sheet.Cells(i + 2, 1).Value = serials(i)
  For j = 0 To UBound(fmts)
    Set valCell = sheet.Cells(i + 2, j + 2)
    valCell.NumberFormat = fmts(j)
    valCell.Value = serials(i)
  Next
Next

sheet.Columns("A:M").ColumnWidth = 12

workbook.SaveAs outputPath, xlOpenXMLWorkbook
workbook.Close False
excel.Quit
Set excel = Nothing
