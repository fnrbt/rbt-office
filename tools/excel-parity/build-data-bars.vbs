' cscript build-data-bars.vbs OUT.tsv PNGDIR: builds synthetic-data-bars.xlsx in Excel (saved into PNGDIR) - gradient data
' bars of six colours (values 1 .. 5 in 64-px columns), right to left, and negative values with an automatic axis, a
' mid-cell axis and none, in Excel 2010's defaults (red negative bars) - for the colour a gradient fades to, where bars
' start and end, and how Excel 2019 draws negative bars and the axis. Pictures with shoot-sheets.ps1, compared in
' tools/xlsx-browser `databars-excel-parity`. OUT gets one line a row: row, what it shows.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1): sh.Name = "Bars"
colors = Array(RGB(99, 142, 198), RGB(255, 0, 0), RGB(99, 190, 123), RGB(255, 182, 40), RGB(0, 0, 128), RGB(214, 0, 127))
names = Array("638EC6", "FF0000", "63BE7B", "FFB628", "000080", "D6007F")
For k = 0 To 5
  r = k + 1
  sh.Cells(r, 1).Value = "gradient " & names(k)
  For i = 1 To 5: sh.Cells(r, 1 + i).Value = i: Next
  Set db = sh.Range(sh.Cells(r, 2), sh.Cells(r, 6)).FormatConditions.AddDatabar
  db.BarColor.Color = colors(k)
  out.WriteLine r & vbTab & sh.Cells(r, 1).Value
Next
' 7: right to left; 8 .. 10: values -2 .. 2 with the axis automatic, at the middle, none; 11: -4 .. 1 automatic;
' 12: solid with a border, negative automatic
labels = Array("rtl", "negative automatic", "negative middle", "negative none", "negative mostly", "solid border negative")
For j = 0 To 5
  r = 7 + j
  sh.Cells(r, 1).Value = labels(j)
  For i = 1 To 5
    v = i
    If j >= 1 Then v = i - 3
    If j = 4 Then v = i - 5
    sh.Cells(r, 1 + i).Value = v
  Next
  Set db = sh.Range(sh.Cells(r, 2), sh.Cells(r, 6)).FormatConditions.AddDatabar
  If j = 0 Then db.Direction = -5004
  ' (A bar made by a script has Excel 2007's look - no axis, negative bars as the positive ones -; these get Excel
  ' 2010's defaults as the ribbon gives them: an automatic axis in black, red negative bars.)
  If j >= 1 Then
    db.AxisPosition = 0
    db.AxisColor.Color = RGB(0, 0, 0)
    db.NegativeBarFormat.ColorType = 0
    db.NegativeBarFormat.Color.Color = RGB(255, 0, 0)
  End If
  If j = 2 Then db.AxisPosition = 1
  If j = 3 Then db.AxisPosition = 2
  If j = 5 Then
    db.BarFillType = 0
    db.BarBorder.Type = 1
  End If
  If j >= 1 Then
    On Error Resume Next
    out.WriteLine r & vbTab & labels(j) & vbTab & "axis " & db.AxisPosition & " negative " & Hex(db.NegativeBarFormat.Color.Color) & " axis colour " & Hex(db.AxisColor.Color)
    On Error GoTo 0
  Else
    out.WriteLine r & vbTab & labels(j)
  End If
Next
sh.Columns(1).ColumnWidth = 24
wb.SaveAs pngDir & "\synthetic-data-bars.xlsx", 51
out.Close
wb.Close False: x.Quit
