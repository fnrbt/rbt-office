' cscript build-pivot-indent.vbs OUT.tsv PNGDIR: builds synthetic-pivot-indent.xlsx in Excel (saved into PNGDIR) - compact pivot
' tables of three row levels with CompactRowIndent 1 (the default), 2 and 3, on a sheet of the default Normal font
' (Calibri 11), and the same in Calibri 16 on a sheet "Big" (the font of the pivot's cells) - for how far Excel 2019
' moves each level of a compact pivot table in. Pictures with shoot-sheets.ps1. OUT gets a line a row label cell:
' sheet, cell, text, its IndentLevel, the pivot's CompactRowIndent.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Do While wb.Worksheets.Count < 3: wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count): Loop
Set data = wb.Worksheets(1): data.Name = "Data"
Set pvt = wb.Worksheets(2): pvt.Name = "Indent"
Set big = wb.Worksheets(3): big.Name = "Big"
data.Range("A1:D1").Value = Array("Region", "Product", "Quarter", "Sales")
r = 2
For Each region In Array("East", "West")
  For Each product In Array("Apples", "Pears")
    For q = 1 To 2
      data.Cells(r, 1).Value = region: data.Cells(r, 2).Value = product: data.Cells(r, 3).Value = "Q" & q: data.Cells(r, 4).Value = r * 10
      r = r + 1
    Next
  Next
Next
Set pc = wb.PivotCaches.Create(1, data.Range("A1:D" & (r - 1)))
k = 0
For Each target In Array(pvt, big)
  For i = 1 To 3
    Set pt = pc.CreatePivotTable(target.Cells(2, 1 + (i - 1) * 3), target.Name & i)
    pt.PivotFields("Region").Orientation = 1
    pt.PivotFields("Product").Orientation = 1
    pt.PivotFields("Quarter").Orientation = 1
    pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", -4157
    pt.CompactRowIndent = i
    If target.Name = "Big" Then pt.TableRange1.Font.Size = 16
    For Each c In pt.RowRange.Cells
      out.WriteLine target.Name & vbTab & c.Address(False, False) & vbTab & c.Text & vbTab & c.IndentLevel & vbTab & pt.CompactRowIndent
    Next
  Next
  target.Columns("A:I").ColumnWidth = 16
Next
pvt.Activate
wb.SaveAs pngDir & "\synthetic-pivot-indent.xlsx", 51
out.Close
wb.Close False: x.Quit
