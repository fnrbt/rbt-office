' chart-axis-face-probe.vbs OUT.tsv: the major unit Excel 2019 gives a value axis whose labels are in a given typeface -
' column charts (the axis up the side, heights swept) and bar charts (along the bottom, widths swept) made through COM,
' of one series whose top end is 60, 70, 80, 90 or 100 (so that steps of 10, 20, 50 tell apart 2 .. 10 intervals), the
' tick labels in each face and size. One line each: kind (51 column, 57 bar), face, size, top end, chart width, height,
' the plot's inside width and height, the axis' major unit, its maximum, the plot area's width and height (points). Run it in the desktop session
' (run-vbs-in-session.ps1). docs/excel-parity-tracker.md "a line of the labels' face".
Option Explicit
Dim out, fso, f, xl, wb, ws, co, ch, ax, faces, face, sizes, fs, tops, t, kind, kinds, d, i
out = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.ScreenUpdating = False
Set wb = xl.Workbooks.Add
Set ws = wb.Worksheets(1)
tops = Array(60, 70, 80, 90, 100)
For i = 0 To 4
  ws.Cells(1, i + 1).Value = 1
  ws.Cells(2, i + 1).Value = tops(i) / 2
  ws.Cells(3, i + 1).Value = tops(i) / 1.05 - 0.01
Next
faces = Array("Calibri", "Garamond", "Arial", "Segoe UI", "Times New Roman", "Tw Cen MT", "Georgia", "Verdana", "Trebuchet MS", "Cambria")
sizes = Array(9, 10, 12)
kinds = Array(51, 57)
For Each kind In kinds
  For i = 0 To 4
    Set co = ws.ChartObjects.Add(10, 10, 300, 220)
    Set ch = co.Chart
    ch.ChartType = kind
    ch.SetSourceData ws.Range(ws.Cells(1, i + 1), ws.Cells(3, i + 1))
    ch.HasTitle = False
    ch.HasLegend = False
    Set ax = ch.Axes(2)
    For Each face In faces
      ax.TickLabels.Font.Name = face
      For Each fs In sizes
        ax.TickLabels.Font.Size = fs
        For d = 90 To 420 Step 4
          If kind = 51 Then co.Height = d Else co.Width = d * 2
          ax.MajorUnitIsAuto = True
          ax.MaximumScaleIsAuto = True
          f.WriteLine kind & vbTab & face & vbTab & fs & vbTab & tops(i) & vbTab & co.Width & vbTab & co.Height & vbTab & ch.PlotArea.InsideWidth & vbTab & ch.PlotArea.InsideHeight & vbTab & ax.MajorUnit & vbTab & ax.MaximumScale & vbTab & ch.PlotArea.Width & vbTab & ch.PlotArea.Height
        Next
      Next
    Next
    co.Delete
  Next
Next
f.Close
wb.Close False
xl.Quit
