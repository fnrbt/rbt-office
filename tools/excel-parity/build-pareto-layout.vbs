' build-pareto-layout.vbs OUT.tsv PNGDIR: Pareto charts of Excel 2019 (AddChart2 style -1, 360 x 216 pt) whose counts
' are 1/10 to 10000 times six causes' (so the value axis' labels are 1 to 6 figures wide), and a histogram of each
' (no secondary axis), each exported as PNGDIR\<chart>.png (Chart.Export), the workbook saved as
' PNGDIR\synthetic-pareto-layout.xlsx - for how wide Excel lays a Pareto chart's plot out between its two axes.
' OUT: a line a chart - its name, the plot area's inside (InsideLeft, InsideWidth, points), each axis' MaximumScale.
'   PNGDIR=dir tools/excel-parity/probe-on-vm.sh tools/excel-parity/build-pareto-layout.vbs out.tsv
Option Explicit
Const xlHistogram = 118, xlPareto = 122
Dim f, x, wb, sh, pc, pv, scales, k, i, shp, c, name, line, kind
Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(WScript.Arguments(0), True, True)
Set x = CreateObject("Excel.Application")
x.Visible = True
x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1)
sh.Name = "Pareto"
pc = Array("Late", "Damaged", "Wrong item", "Missing", "Billing", "Other")
pv = Array(42, 17, 88, 9, 25, 5)
scales = Array(0.1, 1, 10, 100, 1000, 10000)
For k = 0 To 5
  sh.Cells(1, 1 + 3 * k).Value = "Cause"
  sh.Cells(1, 2 + 3 * k).Value = "Count"
  For i = 0 To 5
    sh.Cells(i + 2, 1 + 3 * k).Value = pc(i)
    sh.Cells(i + 2, 2 + 3 * k).Value = pv(i) * scales(k)
  Next
Next
On Error Resume Next
For k = 0 To 5
  For Each kind In Array(xlPareto, xlHistogram)
    sh.Range(sh.Cells(1, 1 + 3 * k), sh.Cells(7, 2 + 3 * k)).Select
    Set shp = sh.Shapes.AddChart2(-1, kind, sh.Range("T1").Left + (k Mod 2) * 380, sh.Range("A1").Top + (k \ 2) * 230 + (kind = xlHistogram) * -700, 360, 216)
    If kind = xlPareto Then name = "Pareto" & k Else name = "Hist" & k
    shp.Name = name
    Set c = shp.Chart
    c.SetSourceData sh.Range(sh.Cells(1, 1 + 3 * k), sh.Cells(7, 2 + 3 * k))
    line = name & vbTab & c.PlotArea.InsideLeft & vbTab & c.PlotArea.InsideWidth & vbTab & c.PlotArea.InsideTop & vbTab & c.PlotArea.InsideHeight
    Err.Clear
    line = line & vbTab & c.Axes(2, 1).MaximumScale
    Err.Clear
    line = line & vbTab & c.Axes(2, 2).MaximumScale
    f.WriteLine line
    c.Export WScript.Arguments(1) & "\" & name & ".png"
    Err.Clear
  Next
Next
On Error GoTo 0
wb.SaveAs WScript.Arguments(1) & "\synthetic-pareto-layout.xlsx", 51
f.Close
wb.Close False
x.Quit
