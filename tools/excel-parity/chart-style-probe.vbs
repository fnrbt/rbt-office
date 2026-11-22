' chart-style-probe.vbs WORKBOOK OUT.tsv: synthetic-chart-lines.xlsx's Plain2 (a line chart stating no lines) and PiePlain
' under each chart style 1 .. 48 (Chart.ChartStyle): the series' line weight and visibility, a pie point's.
Option Explicit
Dim book, out, fso, f, xl, wb, ws, co, ch, st, p
book = WScript.Arguments(0)
out = WScript.Arguments(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Open(book, False, False)
Set ws = wb.Worksheets(1)
On Error Resume Next
For Each co In ws.ChartObjects
  If co.Name = "Plain2" Or co.Name = "PiePlain" Or co.Name = "ScatterPlain" Then
    Set ch = co.Chart
    For st = 1 To 48
      Err.Clear
      ch.ChartStyle = st
      Set p = ch.SeriesCollection(1)
      If co.Name = "PiePlain" Then
        f.WriteLine co.Name & vbTab & st & vbTab & ch.ChartStyle & vbTab & p.Points(1).Format.Line.Visible & vbTab & p.Points(1).Format.Line.Weight & vbTab & Hex(p.Points(1).Format.Line.ForeColor.RGB) & vbTab & p.Points(1).Border.LineStyle & vbTab & p.Points(1).Border.Weight
      Else
        f.WriteLine co.Name & vbTab & st & vbTab & ch.ChartStyle & vbTab & p.Format.Line.Visible & vbTab & p.Format.Line.Weight & vbTab & Hex(p.Format.Line.ForeColor.RGB) & vbTab & p.Border.LineStyle & vbTab & p.Border.Weight
      End If
    Next
  End If
Next
On Error GoTo 0
f.Close
wb.Close False
xl.Quit
