' chartex-treemap-probe.vbs OUT.tsv: where Excel 2019 puts the rectangles of flat treemaps (one level of categories)
' of generated values in charts of several sizes - each point's Left / Top / Width / Height through COM (inches, from
' an origin of Excel's own: only their differences count). One line per point: case, chart width and height (pt),
' the point's number, its value, its rectangle. Data are made here (a Lehmer generator). Run it in the desktop session:
'   tools/excel-parity/probe-on-vm.sh tools/excel-parity/chartex-treemap-probe.vbs out.tsv
Option Explicit
Const xlTreemap = 117
Dim out, fso, f, x, wb, sh, seed, i, j, k, n, c, shp, p, kase, sizes, w, h, col, values, t
out = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set x = CreateObject("Excel.Application")
x.Visible = True
x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1)

Function Rand()
  seed = seed * 16807
  seed = seed - Int(seed / 2147483647) * 2147483647
  Rand = seed / 2147483647
End Function

sizes = Array(Array(360, 216), Array(216, 360), Array(300, 300), Array(500, 160))
On Error Resume Next
For kase = 0 To 39
  seed = 1000 + kase * 7
  n = 3 + Int(Rand() * 12)
  If kase < 4 Then n = 10
  ReDim values(n - 1)
  For i = 0 To n - 1
    Select Case kase Mod 4
      Case 0: values(i) = 10
      Case 1: values(i) = Int(Rand() * 100) + 1
      Case 2: values(i) = Int(Rand() * Rand() * 1000) + 1
      Case 3: values(i) = Int(100 * 0.7 ^ i) + 1
    End Select
    If kase = 1 Then values(i) = 1: If i = 0 Then values(i) = 50
    If kase = 2 Then values(i) = Int(Rand() * 100) + 1
  Next
  col = 1
  sh.Cells.Clear
  For i = 0 To n - 1
    sh.Cells(i + 2, 1).Value = "P" & (i + 1)
    sh.Cells(i + 2, 2).Value = values(i)
  Next
  sh.Cells(1, 1).Value = "Name": sh.Cells(1, 2).Value = "Value"
  For k = 0 To 3
    w = sizes(k)(0): h = sizes(k)(1)
    sh.Range(sh.Cells(1, 1), sh.Cells(n + 1, 2)).Select
    Set shp = sh.Shapes.AddChart2(-1, xlTreemap, 300, 10, w, h)
    Set c = shp.Chart
    c.SetSourceData sh.Range(sh.Cells(1, 1), sh.Cells(n + 1, 2))
    c.HasLegend = False
    c.HasTitle = False
    Set p = c.FullSeriesCollection(1)
    For j = 1 To p.Points.Count
      Err.Clear
      t = p.Points(j).Left * 72 & vbTab & p.Points(j).Top * 72 & vbTab & p.Points(j).Width * 72 & vbTab & p.Points(j).Height * 72
      If Err.Number <> 0 Then t = "-"
      f.WriteLine kase & vbTab & w & vbTab & h & vbTab & j & vbTab & values(j - 1) & vbTab & t
    Next
    shp.Delete
  Next
Next
On Error GoTo 0
f.Close
wb.Close False
x.Quit
