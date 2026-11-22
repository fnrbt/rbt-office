' chart-marker-probe.vbs WORKBOOK OUT.tsv: what Excel 2019 makes of each series' markers and smoothing in every chart of a
' workbook: one line per series - sheet, chart, series, chart type, Smooth, MarkerStyle (xlMarkerStyle: 2 diamond, 1
' square, 3 triangle, -4168 x, 5 star, 8 circle, 9 plus, -4115 dash, -4142 none, -4105 automatic), MarkerSize, the
' marker's fill and outline colours (hex BGR) and whether each is visible, the series' line weight. Run it in the
' desktop session (run-vbs-in-session.ps1; probe-on-vm.sh with BOOKFIRST=1). docs/excel-parity-tracker.md "markers".
Option Explicit
Dim book, out, fso, f, xl, wb, ws, co, s, p, text
book = WScript.Arguments(0)
out = WScript.Arguments(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(book, False, True)
On Error Resume Next
For Each ws In wb.Worksheets
  For Each co In ws.ChartObjects
    For s = 1 To co.Chart.SeriesCollection.Count
      Set p = co.Chart.SeriesCollection(s)
      text = ws.Name & vbTab & co.Name & vbTab & s & vbTab & p.ChartType & vbTab & p.Smooth & vbTab & p.MarkerStyle & vbTab & p.MarkerSize
      text = text & vbTab & Hex(p.MarkerBackgroundColor) & vbTab & Hex(p.MarkerForegroundColor) & vbTab & p.MarkerBackgroundColorIndex & vbTab & p.MarkerForegroundColorIndex
      text = text & vbTab & p.Format.Fill.Visible & vbTab & Hex(p.Format.Fill.ForeColor.RGB) & vbTab & p.Format.Line.Weight
      f.WriteLine text
      Err.Clear
    Next
  Next
Next
f.Close
wb.Close False
xl.Quit
