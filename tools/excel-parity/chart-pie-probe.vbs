' chart-pie-probe.vbs OUT.tsv: the size Excel 2019 gives a pie / doughnut made through COM with data labels of each position and size (chart type, size, position, font size, chart area, plot inside rectangle). Run it in the desktop session (run-vbs-in-session.ps1).
Option Explicit
Dim out, fso, f, xl, wb, ws, co, ch, kinds, kind, pos, positions, fs, sizes, dims, d, p
out = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add
Set ws = wb.Worksheets(1)
ws.Range("A1:A4").Value = xl.WorksheetFunction.Transpose(Array("A", "B", "C", "D"))
ws.Range("B1:B4").Value = xl.WorksheetFunction.Transpose(Array(30, 25, 25, 20))
kinds = Array(5, -4120)
positions = Array(0, -4108, 3, 2, 5)   ' none, center, inside end, outside end, best fit
sizes = Array(8, 10, 14)
dims = Array(Array(300, 200), Array(400, 250), Array(250, 300))
On Error Resume Next
For Each kind In kinds
  For Each d In dims
    For Each pos In positions
      For Each fs In sizes
        Set co = ws.ChartObjects.Add(10, 10, d(0), d(1))
        Set ch = co.Chart
        ch.ChartType = kind
        ch.SetSourceData ws.Range("A1:B4")
        ch.HasTitle = False
        ch.HasLegend = False
        Set p = ch.SeriesCollection(1)
        If pos <> 0 Then
          p.HasDataLabels = True
          p.DataLabels.ShowValue = True
          p.DataLabels.Font.Size = fs
          Err.Clear
          If kind = 5 Then p.DataLabels.Position = pos
        End If
        f.WriteLine kind & vbTab & d(0) & vbTab & d(1) & vbTab & pos & vbTab & fs & vbTab & Err.Number & vbTab & ch.ChartArea.Width & vbTab & ch.ChartArea.Height & vbTab & ch.PlotArea.InsideLeft & vbTab & ch.PlotArea.InsideTop & vbTab & ch.PlotArea.InsideWidth & vbTab & ch.PlotArea.InsideHeight
        co.Delete
        If pos = 0 Then Exit For
      Next
    Next
  Next
Next
f.Close
wb.Close False
xl.Quit
