' cscript build-sparklines.vbs OUT.tsv PNGDIR: builds synthetic-sparkline-looks.xlsx in Excel (saved into PNGDIR) - one
' sparkline a row, in a cell of 200 x 40 px (column A), its data in C:N of the row - for how Excel 2019 draws a sparkline
' (the room it leaves at the cell's sides, where points and columns stand, markers, the axis, empty cells): pictures
' with shoot-sheets.ps1, compared in tools/xlsx-browser `sparklines-excel-parity`. OUT gets one line a case: row, name.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1): sh.Name = "Sparklines"
sh.Columns(1).ColumnWidth = 27.86   ' 200 px in Calibri 11
sh.Columns(2).ColumnWidth = 14
' name, type (1 line, 2 column, 3 win/loss), values ("" an empty cell), options
cases = Array( _
  Array("LineRamp", 1, "0,1,2,3,4,5,6,7,8,9,10", ""), _
  Array("LineZigzag", 1, "3,7,4,9,6,11,5", ""), _
  Array("LineMarkers", 1, "3,7,4,9,6,11,5", "markers"), _
  Array("LineHighLow", 1, "3,-2,4,9,-6,11,5", "high,low,first,last,negative"), _
  Array("LineAxis", 1, "3,-2,4,9,-6,11,5", "axis"), _
  Array("LineGap", 1, "3,7,,9,,11,5", ""), _
  Array("LineZero", 1, "3,7,,9,,11,5", "zero"), _
  Array("LineSpan", 1, "3,7,,9,,11,5", "span"), _
  Array("LineWeight", 1, "3,7,4,9,6,11,5", "weight"), _
  Array("LineFlat", 1, "5,5,5,5,5", ""), _
  Array("ColRamp", 2, "1,2,3,4,5", ""), _
  Array("ColNeg", 2, "4,-2,5,-1,6,2,-3", ""), _
  Array("ColNegAxis", 2, "4,-2,5,-1,6,2,-3", "axis,negative"), _
  Array("ColMany", 2, "1,2,3,4,5,6,7,8,9,10,11,12", ""), _
  Array("ColCustom", 2, "1,2,3,4,5", "custom"), _
  Array("ColHigh", 2, "3,7,4,9,6,11,5", "high,low,first,last"), _
  Array("WinLoss", 3, "1,-1,1,1,-1,1,-1", ""), _
  Array("WinLossNeg", 3, "4,-2,0,5,-1,6,-3", "negative,axis"), _
  Array("LineRTL", 1, "0,1,2,3,4,5,6,7,8,9,10", "rtl"))
r = 1
For Each c In cases
  sh.Rows(r).RowHeight = 30   ' 40 px
  sh.Cells(r, 2).Value = c(0)
  vals = Split(c(2), ",")
  For i = 0 To UBound(vals)
    If vals(i) <> "" Then sh.Cells(r, 3 + i).Value = CDbl(vals(i))
  Next
  Set g = sh.Cells(r, 1).SparklineGroups.Add(c(1), sh.Name & "!" & sh.Range(sh.Cells(r, 3), sh.Cells(r, 3 + UBound(vals))).Address)
  Set grp = g
  opts = "," & c(3) & ","
  If InStr(opts, ",markers,") > 0 Then grp.Points.Markers.Visible = True
  If InStr(opts, ",high,") > 0 Then grp.Points.Highpoint.Visible = True
  If InStr(opts, ",low,") > 0 Then grp.Points.Lowpoint.Visible = True
  If InStr(opts, ",first,") > 0 Then grp.Points.Firstpoint.Visible = True
  If InStr(opts, ",last,") > 0 Then grp.Points.Lastpoint.Visible = True
  If InStr(opts, ",negative,") > 0 Then grp.Points.Negative.Visible = True
  If InStr(opts, ",axis,") > 0 Then grp.Axes.Horizontal.Axis.Visible = True
  If InStr(opts, ",zero,") > 0 Then grp.DisplayBlanksAs = 1
  If InStr(opts, ",span,") > 0 Then grp.DisplayBlanksAs = 3
  If InStr(opts, ",weight,") > 0 Then grp.LineWeight = 2.25
  If InStr(opts, ",rtl,") > 0 Then grp.Axes.Horizontal.RightToLeftPlotOrder = True
  If InStr(opts, ",custom,") > 0 Then
    grp.Axes.Vertical.MinScaleType = 2: grp.Axes.Vertical.CustomMinScaleValue = 0
    grp.Axes.Vertical.MaxScaleType = 2: grp.Axes.Vertical.CustomMaxScaleValue = 10
  End If
  out.WriteLine r & vbTab & c(0)
  r = r + 1
Next
wb.SaveAs pngDir & "\synthetic-sparkline-looks.xlsx", 51
out.Close
wb.Close False: x.Quit
