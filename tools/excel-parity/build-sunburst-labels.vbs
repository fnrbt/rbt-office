' build-sunburst-labels.vbs OUT.tsv PNGDIR: sunbursts of Excel 2019 (AddChart2 style -1) for where Excel breaks a part's
' label and when it leaves one out, each exported as PNGDIR\<chart>.png (Chart.Export), the workbook saved as
' PNGDIR\synthetic-sunburst-labels.xlsx. Sheet "Widths": three levels (A > B > nine leaves of equal value), the leaves
' "0000" and 0 - 8 "i"s after it (Calibri 9 pt: 24.3 px wide and 2.75 px more a letter, by the design advances), and
' "Osaka", "France", "Europe", "Madrid", "Ab Cd" - in charts 360 x 216, 368 x 222 and 376 x 228 pt (the rings 1 px
' apart). Sheet "Sweeps": "Osaka" leaves of values 1 .. 12 beside a large one (where a part too narrow loses its label).
' OUT: a line a chart - its name, the plot area's inside (points), each point's label text as COM has it.
'   PNGDIR=dir tools/excel-parity/probe-on-vm.sh tools/excel-parity/build-sunburst-labels.vbs out.tsv
Option Explicit
Const xlSunburst = 120
Dim f, x, wb, sh, i, k, labels, shp, c, n, line, sizes, sw, p, j
Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(WScript.Arguments(0), True, True)
Set x = CreateObject("Excel.Application")
x.Visible = True
x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "Widths"
labels = Array("0000", "0000i", "0000ii", "0000iii", "0000iiii", "0000iiiii", "0000iiiiii", "0000iiiiiii", "0000iiiiiiii", "Osaka", "France", "Europe", "Madrid", "Ab Cd")
sh.Range("A1:D1").Value = Array("L1", "L2", "L3", "Value")
For i = 0 To UBound(labels)
  sh.Cells(i + 2, 1).Value = "A"
  sh.Cells(i + 2, 2).Value = "B"
  sh.Cells(i + 2, 3).Value = labels(i)
  sh.Cells(i + 2, 4).Value = 10
Next
n = UBound(labels) + 2
Set sw = wb.Worksheets.Add(, sh)
sw.Name = "Sweeps"
sw.Range("A1:D1").Value = Array("L1", "L2", "L3", "Value")
For i = 0 To 12
  sw.Cells(i + 2, 1).Value = "A"
  sw.Cells(i + 2, 2).Value = "B"
  If i = 0 Then sw.Cells(i + 2, 3).Value = "Big" Else sw.Cells(i + 2, 3).Value = "Osaka" & Chr(64 + i)
  If i = 0 Then sw.Cells(i + 2, 4).Value = 100 Else sw.Cells(i + 2, 4).Value = i
Next
sizes = Array(Array(360, 216), Array(368, 222), Array(376, 228))
On Error Resume Next
For k = 0 To 3
  If k < 3 Then
    sh.Range("A1:D" & n).Select
    Set shp = sh.Shapes.AddChart2(-1, xlSunburst, sh.Range("H1").Left + k * 400, sh.Range("A1").Top, sizes(k)(0), sizes(k)(1))
    shp.Name = "Widths" & k
    shp.Chart.SetSourceData sh.Range("A1:D" & n)
  Else
    sw.Range("A1:D14").Select
    Set shp = sw.Shapes.AddChart2(-1, xlSunburst, sw.Range("H1").Left, sw.Range("A1").Top, 360, 216)
    shp.Name = "Sweeps"
    shp.Chart.SetSourceData sw.Range("A1:D14")
  End If
  Set c = shp.Chart
  ' (Data labels on, as the chart's "+" button turns them on: the category's name.)
  c.SetElement 202
  c.FullSeriesCollection(1).HasDataLabels = True
  c.FullSeriesCollection(1).DataLabels.ShowCategoryName = True
  c.FullSeriesCollection(1).DataLabels.ShowValue = False
  Err.Clear
  line = shp.Name & vbTab & c.PlotArea.InsideLeft & vbTab & c.PlotArea.InsideTop & vbTab & c.PlotArea.InsideWidth & vbTab & c.PlotArea.InsideHeight
  Err.Clear
  Set p = c.FullSeriesCollection(1)
  For j = 1 To p.Points.Count
    Err.Clear
    line = line & vbTab & Replace(p.Points(j).DataLabel.Text, vbLf, "\n")
  Next
  f.WriteLine line
  c.Export WScript.Arguments(1) & "\" & shp.Name & ".png"
  Err.Clear
Next
On Error GoTo 0
wb.SaveAs WScript.Arguments(1) & "\synthetic-sunburst-labels.xlsx", 51
f.Close
wb.Close False
x.Quit
