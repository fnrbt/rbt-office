' cscript build-icon-sets.vbs OUT.tsv PNGDIR: builds synthetic-icon-sets.xlsx in Excel (saved into PNGDIR) - each of
' Excel's 20 icon sets on a row of five cells (1 .. 5: every icon of the set once), then the 3 traffic lights
' reversed, without their values, centred and in a 30-pt row, and a data bar row of each kind - for how Excel 2019
' draws an icon (its picture, size and place in the cell) and a data bar: pictures with shoot-sheets.ps1, compared in
' tools/xlsx-browser `icons-excel-parity`. OUT gets one line a row: row, what it shows.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1): sh.Name = "Icons"
names = Array("3Arrows", "3ArrowsGray", "3Flags", "3TrafficLights1", "3TrafficLights2", "3Signs", "3Symbols", "3Symbols2", _
  "4Arrows", "4ArrowsGray", "4RedToBlack", "4Rating", "4TrafficLights", "5Arrows", "5ArrowsGray", "5Rating", "5Quarters", _
  "3Stars", "3Triangles", "5Boxes")
For k = 1 To 20
  sh.Cells(k, 1).Value = names(k - 1)
  For i = 1 To 5: sh.Cells(k, 1 + i).Value = i: Next
  Set fc = sh.Range(sh.Cells(k, 2), sh.Cells(k, 6)).FormatConditions.AddIconSetCondition
  fc.IconSet = wb.IconSets(k)
  out.WriteLine k & vbTab & names(k - 1)
Next
' 21: reversed; 22: icons only; 23: centred; 24: a 30-pt row
extras = Array("reverse", "noValue", "centred", "tall")
For j = 0 To 3
  r = 21 + j
  sh.Cells(r, 1).Value = "3TrafficLights1 " & extras(j)
  For i = 1 To 5: sh.Cells(r, 1 + i).Value = i: Next
  Set fc = sh.Range(sh.Cells(r, 2), sh.Cells(r, 6)).FormatConditions.AddIconSetCondition
  fc.IconSet = wb.IconSets(4)
  If j = 0 Then fc.ReverseOrder = True
  If j = 1 Then fc.ShowIconOnly = True
  If j = 2 Then sh.Range(sh.Cells(r, 2), sh.Cells(r, 6)).HorizontalAlignment = -4108
  If j = 3 Then sh.Rows(r).RowHeight = 30
  out.WriteLine r & vbTab & sh.Cells(r, 1).Value
Next
' 25 .. 28: data bars - gradient, solid, with negatives, right to left
bars = Array("gradient", "solid", "negative", "rtl")
For j = 0 To 3
  r = 25 + j
  sh.Cells(r, 1).Value = "bar " & bars(j)
  For i = 1 To 5
    v = i * 20
    If j = 2 Then v = (i - 2) * 25
    sh.Cells(r, 1 + i).Value = v
  Next
  Set db = sh.Range(sh.Cells(r, 2), sh.Cells(r, 6)).FormatConditions.AddDatabar
  If j = 1 Then db.BarFillType = 0
  If j = 3 Then db.Direction = -5004
  out.WriteLine r & vbTab & sh.Cells(r, 1).Value
Next
sh.Columns(1).ColumnWidth = 24
wb.SaveAs pngDir & "\synthetic-icon-sets.xlsx", 51
out.Close
wb.Close False: x.Quit
