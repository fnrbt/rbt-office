' legacy-chart-probe.vbs OUT.tsv PNGDIR: a chart made with the old API (ChartObjects.Add, no chart style: a part that
' names no text sizes nor lines, as Charts.Add writes a chart sheet's) on a worksheet - its title's and labels' sizes,
' its gridlines' colour through COM, its picture - saved as PNGDIR\legacy-chart.xlsx to read what Excel wrote.
Option Explicit
Dim out, pngDir, fso, f, x, wb, ws, co, ch, i
out = WScript.Arguments(0)
pngDir = WScript.Arguments(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set x = CreateObject("Excel.Application")
x.Visible = True
x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set ws = wb.Worksheets(1)
ws.Range("A1").Value = "Quarter": ws.Range("B1").Value = "Sales": ws.Range("C1").Value = "Costs"
For i = 0 To 3
  ws.Cells(i + 2, 1).Value = "Q" & (i + 1)
  ws.Cells(i + 2, 2).Value = Array(120, 135, 160, 150)(i)
  ws.Cells(i + 2, 3).Value = Array(80, 95, 90, 110)(i)
Next
Set co = ws.ChartObjects.Add(ws.Range("E2").Left, ws.Range("E2").Top, 360, 216)
co.Name = "Legacy"
Set ch = co.Chart
ch.ChartType = 51
ch.SetSourceData ws.Range("A1:C5")
ch.HasTitle = True
ch.ChartTitle.Text = "Legacy"
On Error Resume Next
f.WriteLine "title" & vbTab & ch.ChartTitle.Font.Size & vbTab & ch.ChartTitle.Font.Bold & vbTab & Hex(ch.ChartTitle.Font.Color)
f.WriteLine "ticks" & vbTab & ch.Axes(2).TickLabels.Font.Size & vbTab & Hex(ch.Axes(2).TickLabels.Font.Color)
f.WriteLine "grid" & vbTab & Hex(ch.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB) & vbTab & ch.Axes(2).MajorGridlines.Format.Line.Weight
f.WriteLine "legend" & vbTab & ch.Legend.Font.Size
f.WriteLine "style" & vbTab & ch.ChartStyle
ch.Export pngDir & "\legacy-chart.png"
On Error GoTo 0
wb.SaveAs pngDir & "\legacy-chart.xlsx", 51
wb.Close False
f.Close
x.Quit
