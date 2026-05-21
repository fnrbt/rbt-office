' build-numfmt-sections-fixture.vbs
' Builds tests/test-documents-excel/synthetic-numfmt-sections.xlsx: cells with
' conditional ([>=100]...) and colour ([Red]/[Blue]) number-format sections at
' values that exercise each branch. Coverage + gallery for section selection.
' Usage: cscript //nologo build-numfmt-sections-fixture.vbs <input> <out.xlsx>
Option Explicit
Const xlOpenXMLWorkbook = 51
Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then WScript.Quit 2
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
Dim excel, wb, sh
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "NumFmtSections"

' Column A: labels. B: conditional. C: sign+colour. D: conditional+colour.
sh.Cells(1,1).Value = "value"
sh.Cells(1,2).Value = "[>=100]High;[>=50]Mid;Low"
sh.Cells(1,3).Value = "0.0;[Red]-0.0;zero"
sh.Cells(1,4).Value = "[<0][Red]0.0;[Blue]0.0"

Dim vals, i
vals = Array(150, 70, 30, 0, -10)
For i = 0 To UBound(vals)
  Dim r
  r = i + 2
  sh.Cells(r,1).Value = vals(i)
  sh.Cells(r,2).Value = vals(i)
  sh.Cells(r,2).NumberFormat = "[>=100]""High"";[>=50]""Mid"";""Low"""
  sh.Cells(r,3).Value = vals(i)
  sh.Cells(r,3).NumberFormat = "0.0;[Red]-0.0;""zero"""
  sh.Cells(r,4).Value = vals(i)
  sh.Cells(r,4).NumberFormat = "[<0][Red]0.0;[Blue]0.0"
Next
sh.Columns("A:D").ColumnWidth = 16

wb.SaveAs outputPath, xlOpenXMLWorkbook
wb.Close False
excel.Quit
Set excel = Nothing
