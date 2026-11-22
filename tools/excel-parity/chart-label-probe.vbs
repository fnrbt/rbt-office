' chart-label-probe.vbs OUT.tsv [PNGDIR]: Excel 2019's automatic category labels on column charts made through COM -
' label kinds (1: 1, 2, 3 ...; 2: "Cat 1"; 3: "Category 1"; 4: "Q1 2000"; 6: mixed lengths) x category counts x
' tick label sizes x chart widths: one line each of kind, count, size, chart width, the plot's inside left and width,
' the label spacing Excel chose (TickLabelSpacing), the category axis' top and height (a line, lines, turned -45 or
' -90 tell apart by it), its orientation, the plot's inside height (points). With PNGDIR a picture whenever the
' spacing or the axis' height changes. Run it in the desktop session (run-vbs-in-session.ps1); 13,000 charts take
' about 15 minutes. docs/excel-parity-tracker.md "category labels".
Option Explicit
Dim mixed, last, cur, out, fso, xl, wb, ws, co, ch, ax, f, i, n, w, fs, kinds, kind, ns, sizes, a, b, c, orient, ext
out = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.ScreenUpdating = False
Set wb = xl.Workbooks.Add
Set ws = wb.Worksheets(1)
For i = 1 To 120
  ws.Cells(i, 1).Value = i
  ws.Cells(i, 2).Value = "Cat " & i
  ws.Cells(i, 3).Value = "Category " & i
  ws.Cells(i, 4).Value = "Q" & ((i - 1) Mod 4 + 1) & " " & (2000 + (i - 1) \ 4)
  ws.Cells(i, 5).Value = 1 + (i Mod 3)
  mixed = Array("A", "Long label " & i, "Mid " & i, "Wordy text label number " & i)
  ws.Cells(i, 6).Value = mixed((i - 1) Mod 4)
Next
kinds = Array(1, 2, 3, 4, 6)
ns = Array(6, 12, 24, 40, 60, 100)
sizes = Array(8, 9, 10, 11)
ext = ""
If WScript.Arguments.Count > 1 Then ext = WScript.Arguments(1)
For Each kind In kinds
  For Each n In ns
    Set co = ws.ChartObjects.Add(10, 10, 300, 220)
    Set ch = co.Chart
    ch.ChartType = 51
    ch.SetSourceData ws.Range(ws.Cells(1, 5), ws.Cells(n, 5))
    ch.SeriesCollection(1).XValues = ws.Range(ws.Cells(1, kind), ws.Cells(n, kind))
    ch.HasTitle = False
    ch.HasLegend = False
    Set ax = ch.Axes(1)
    For Each fs In sizes
      ax.TickLabels.Font.Size = fs
      For w = 90 To 760 Step 6
        co.Width = w
        ax.TickLabelSpacingIsAuto = True
        orient = ax.TickLabels.Orientation
        f.WriteLine kind & vbTab & n & vbTab & fs & vbTab & co.Width & vbTab & ch.PlotArea.InsideLeft & vbTab & ch.PlotArea.InsideWidth & vbTab & ax.TickLabelSpacing & vbTab & ax.Top & vbTab & ax.Height & vbTab & orient & vbTab & ch.PlotArea.InsideHeight
        cur = ax.TickLabelSpacing & "/" & ax.Height
        If ext <> "" And cur <> last Then ch.Export ext & "\k" & kind & "-n" & n & "-w" & w & ".png"
        last = cur
      Next
    Next
    co.Delete
  Next
Next
f.Close
wb.Close False
xl.Quit
