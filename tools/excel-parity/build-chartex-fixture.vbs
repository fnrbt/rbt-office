' cscript build-chartex-fixture.vbs IGNORED OUT.xlsx: Excel's own waterfall and funnel charts (chartEx parts)
' with its defaults (AddChart2 style -1), for tests/test-documents-excel/synthetic-chartex.xlsx.
Option Explicit
Const xlOpenXMLWorkbook = 51, xlWaterfall = 119, xlFunnel = 123
Dim fso, o: Set fso = CreateObject("Scripting.FileSystemObject")
o = fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x, wb, sh: Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add(): Set sh = wb.Worksheets(1): sh.Name = "ChartEx"
Dim i, cats, vals
cats = Array("Start", "Sales", "Returns", "Costs", "Tax", "End")
vals = Array(100, 60, -20, -45, -15, 80)
sh.Cells(1, 1).Value = "Step": sh.Cells(1, 2).Value = "Amount"
For i = 0 To 5: sh.Cells(i + 2, 1).Value = cats(i): sh.Cells(i + 2, 2).Value = vals(i): Next
cats = Array("Visits", "Sign-ups", "Trials", "Orders", "Repeat")
vals = Array(1200, 800, 450, 300, 120)
sh.Cells(1, 4).Value = "Stage": sh.Cells(1, 5).Value = "Count"
For i = 0 To 4: sh.Cells(i + 2, 4).Value = cats(i): sh.Cells(i + 2, 5).Value = vals(i): Next
Dim w, f
sh.Range("A1:B7").Select
Set w = sh.Shapes.AddChart2(-1, xlWaterfall, sh.Range("G2").Left, sh.Range("G2").Top, 360, 216)
w.Chart.SetSourceData sh.Range("A1:B7")
w.Chart.HasTitle = True: w.Chart.ChartTitle.Text = "Waterfall"
' The last point is a total (set as total: it starts at the axis).
w.Chart.FullSeriesCollection(1).Points(6).IsTotal = True
w.Chart.FullSeriesCollection(1).Points(1).IsTotal = True
sh.Range("D1:E6").Select
Set f = sh.Shapes.AddChart2(-1, xlFunnel, sh.Range("G18").Left, sh.Range("G18").Top, 360, 216)
f.Chart.SetSourceData sh.Range("D1:E6")
f.Chart.HasTitle = True: f.Chart.ChartTitle.Text = "Funnel"
sh.Range("A1").Select
wb.SaveAs o, xlOpenXMLWorkbook: wb.Close False: x.Quit
