' build-merged-text-fixture.vbs
'
' Builds tests/test-documents-excel/synthetic-merged-text.xlsx: merged ranges
' carrying text with various horizontal / vertical alignments + fill + border,
' to check that content centers across the FULL merge rect (not the primary
' cell) and that vertical alignment uses the merge's full height.
'
' Usage: cscript //nologo build-merged-text-fixture.vbs <unused-input> <output.xlsx>

Option Explicit
Const xlOpenXMLWorkbook = 51
Const xlCenter = -4108
Const xlBottom = -4107
Const xlTop = -4160
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12
Const xlContinuous = 1
Const xlThin = 2
Const xlSolid = 1

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: build-merged-text-fixture.vbs <input> <output.xlsx>"
  WScript.Quit 2
End If
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, wb, sh
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "MergedText"

' Title bar: A1:D1 merged, horizontally centered bold text, fill + box.
Dim title
Set title = sh.Range("A1:D1")
title.Merge
sh.Cells(1, 1).Value = "Quarterly Report"
title.HorizontalAlignment = xlCenter
title.VerticalAlignment = xlCenter
sh.Cells(1, 1).Font.Bold = True
title.Interior.Color = RGB(221, 235, 247)
title.Borders(xlEdgeLeft).LineStyle = xlContinuous
title.Borders(xlEdgeTop).LineStyle = xlContinuous
title.Borders(xlEdgeRight).LineStyle = xlContinuous
title.Borders(xlEdgeBottom).LineStyle = xlContinuous

' Vertically + horizontally centered block: A3:B5 merged.
Dim vc
Set vc = sh.Range("A3:B5")
vc.Merge
sh.Cells(3, 1).Value = "Centered"
vc.HorizontalAlignment = xlCenter
vc.VerticalAlignment = xlCenter
vc.Borders(xlEdgeLeft).LineStyle = xlContinuous
vc.Borders(xlEdgeTop).LineStyle = xlContinuous
vc.Borders(xlEdgeRight).LineStyle = xlContinuous
vc.Borders(xlEdgeBottom).LineStyle = xlContinuous

' Bottom-aligned block: D3:D5 merged tall cell, text at bottom.
Dim bb
Set bb = sh.Range("D3:D5")
bb.Merge
sh.Cells(3, 4).Value = "Bottom"
bb.HorizontalAlignment = xlCenter
bb.VerticalAlignment = xlBottom

' Top-aligned block: F3:F5 merged tall cell, text at top.
Dim tb
Set tb = sh.Range("F3:F5")
tb.Merge
sh.Cells(3, 6).Value = "Top"
tb.HorizontalAlignment = xlCenter
tb.VerticalAlignment = xlTop

wb.SaveAs outputPath, xlOpenXMLWorkbook
wb.Close False
excel.Quit
Set excel = Nothing
