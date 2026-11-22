' chart-marker-style-probe.vbs WORKBOOK OUT.tsv [PNGDIR]: synthetic-chart-looks.xlsx's "Markers" (eight line series stating
' nothing of their markers) under each chart style 1 .. 48 (Chart.ChartStyle): each series' marker size and style, line
' weight; with PNGDIR a picture of the chart under each style (style-NN.png). Run it in the desktop session
' (probe-on-vm.sh with BOOKFIRST=1, PNGDIR). docs/excel-parity-tracker.md "markers".
Option Explicit
Dim book, out, png, fso, f, xl, wb, ws, co, ch, st, s, p
book = WScript.Arguments(0)
out = WScript.Arguments(1)
png = ""
If WScript.Arguments.Count > 2 Then png = WScript.Arguments(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(book, False, False)
Set ws = wb.Worksheets(1)
On Error Resume Next
Set co = ws.ChartObjects("Markers")
Set ch = co.Chart
For st = 1 To 48
  Err.Clear
  ch.ChartStyle = st
  For s = 1 To 2
    Set p = ch.SeriesCollection(s)
    f.WriteLine st & vbTab & ch.ChartStyle & vbTab & s & vbTab & p.MarkerStyle & vbTab & p.MarkerSize & vbTab & p.Format.Line.Weight & vbTab & Hex(p.MarkerBackgroundColor) & vbTab & Hex(p.MarkerForegroundColor)
  Next
  If png <> "" Then ch.Export png & "\style-" & Right("0" & st, 2) & ".png"
Next
On Error GoTo 0
f.Close
wb.Close False
xl.Quit
