' build-chartex-zoo.vbs OUT.tsv PNGDIR: Excel 2019's own chartEx charts of every kind (histogram, pareto, box and
' whisker, treemap, sunburst, waterfall, funnel), made through COM with Excel's defaults (AddChart2 style -1), saved
' as PNGDIR\synthetic-chartex-zoo.xlsx (tests/test-documents-excel/), each chart exported as PNGDIR\<sheet>-<chart>.png
' and what COM tells of it written to OUT.tsv (points, their labels and rectangles, axes' scales, the plot area).
' Data are made here, deterministically (a Lehmer generator), so the sheet holds them. Run it in the desktop session:
'   PNGDIR=dir tools/excel-parity/probe-on-vm.sh tools/excel-parity/build-chartex-zoo.vbs out.tsv
Option Explicit
Const xlOpenXMLWorkbook = 51
Const xlTreemap = 117, xlHistogram = 118, xlWaterfall = 119, xlSunburst = 120, xlBoxwhisker = 121, xlPareto = 122, xlFunnel = 123
Dim out, pngDir, fso, f, x, wb, sh, seed
out = WScript.Arguments(0)
pngDir = WScript.Arguments(1)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set x = CreateObject("Excel.Application")
x.Visible = True
x.DisplayAlerts = False
Set wb = x.Workbooks.Add()

Function Rand()
  seed = seed * 16807
  seed = seed - Int(seed / 2147483647) * 2147483647
  Rand = seed / 2147483647
End Function

Function NewSheet(name)
  Dim s
  If wb.Worksheets(1).Name = "Sheet1" And wb.Worksheets(1).UsedRange.Count = 1 And IsEmpty(wb.Worksheets(1).Range("A1").Value) Then
    Set s = wb.Worksheets(1)
  Else
    Set s = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
  End If
  s.Name = name
  Set NewSheet = s
End Function

Dim charts(): ReDim charts(0)
Dim nCharts: nCharts = 0
Function AddChart(sheet, kind, rng, name, title, slot)
  Dim shp, c
  sheet.Range(rng).Select
  Set shp = sheet.Shapes.AddChart2(-1, kind, sheet.Range("N1").Left + (slot Mod 2) * 380, sheet.Range("A1").Top + (slot \ 2) * 230, 360, 216)
  shp.Name = name
  Set c = shp.Chart
  c.SetSourceData sheet.Range(rng)
  If title <> "" Then
    c.HasTitle = True
    c.ChartTitle.Text = title
  End If
  Set AddChart = c
End Function

On Error Resume Next
Dim i, j, k, v, n, c
' ---- histograms: the automatic bins of data sets of several sizes and spreads; then bins set
Set sh = NewSheet("Hist")
Dim sizes: sizes = Array(5, 12, 20, 50, 100, 200, 30, 40)
Dim kinds: kinds = Array("u100", "u100", "n50", "u1000", "n50", "u1000", "small", "int")
For k = 0 To 7
  seed = 12345 + k
  sh.Cells(1, k + 1).Value = "D" & (k + 1)
  For i = 1 To sizes(k)
    Select Case kinds(k)
      Case "u100": v = Int(Rand() * 100) + 1
      Case "u1000": v = Round(Rand() * 1000, 1)
      Case "n50": v = Round((Rand() + Rand() + Rand()) * 33.3, 2)
      Case "small": v = Round(Rand(), 3)
      Case "int": v = Int(Rand() * 7) + 1
    End Select
    sh.Cells(i + 1, k + 1).Value = v
  Next
Next
For k = 0 To 7
  Set c = AddChart(sh, xlHistogram, sh.Range(sh.Cells(1, k + 1), sh.Cells(sizes(k) + 1, k + 1)).Address, "Hist" & (k + 1), "", k)
Next
' bins set: a count, a width with overflow / underflow
Set c = AddChart(sh, xlHistogram, "D1:D51", "HistCount", "Bins 5", 8)
c.Axes(1).BinsType = 3: c.Axes(1).BinsCountValue = 5
Set c = AddChart(sh, xlHistogram, "D1:D51", "HistWidth", "Width 150", 9)
c.Axes(1).BinsType = 2: c.Axes(1).BinWidthValue = 150
c.Axes(1).BinsOverflowEnabled = True: c.Axes(1).BinsOverflowValue = 800
c.Axes(1).BinsUnderflowEnabled = True: c.Axes(1).BinsUnderflowValue = 100
' text categories: counted by category
sh.Range("J1").Value = "Fruit"
Dim fruit: fruit = Array("Apple", "Pear", "Apple", "Plum", "Apple", "Pear", "Fig", "Plum", "Apple", "Pear")
For i = 0 To 9: sh.Cells(i + 2, 10).Value = fruit(i): Next
Set c = AddChart(sh, xlHistogram, "J1:J11", "HistText", "", 10)

' ---- pareto
Set sh = NewSheet("Pareto")
Dim pc: pc = Array("Late", "Damaged", "Wrong item", "Missing", "Billing", "Other")
Dim pv: pv = Array(42, 17, 88, 9, 25, 5)
sh.Range("A1").Value = "Cause": sh.Range("B1").Value = "Count"
For i = 0 To 5: sh.Cells(i + 2, 1).Value = pc(i): sh.Cells(i + 2, 2).Value = pv(i): Next
Set c = AddChart(sh, xlPareto, "A1:B7", "Pareto", "Complaints", 0)
seed = 777: sh.Range("D1").Value = "Values"
For i = 1 To 60: sh.Cells(i + 1, 4).Value = Round((Rand() + Rand()) * 50, 1): Next
Set c = AddChart(sh, xlPareto, "D1:D61", "ParetoBins", "", 1)

' ---- box and whisker
Set sh = NewSheet("Box")
sh.Range("A1").Value = "Group": sh.Range("B1").Value = "Score"
seed = 4242
For i = 1 To 45
  sh.Cells(i + 1, 1).Value = Array("North", "South", "East")((i - 1) Mod 3)
  sh.Cells(i + 1, 2).Value = Round(Rand() * 60 + 20 + ((i - 1) Mod 3) * 10, 0)
Next
sh.Range("B44").Value = 150
Set c = AddChart(sh, xlBoxwhisker, "A1:B46", "BoxCat", "", 0)
sh.Range("D1").Value = "A": sh.Range("E1").Value = "B"
Dim ba: ba = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 30)
Dim bb: bb = Array(10, 12, 13, 15, 18, 22, 25, 26)
For i = 0 To 9: sh.Cells(i + 2, 4).Value = ba(i): Next
For i = 0 To 7: sh.Cells(i + 2, 5).Value = bb(i): Next
Set c = AddChart(sh, xlBoxwhisker, "D1:E11", "BoxTwo", "", 1)
Set c = AddChart(sh, xlBoxwhisker, "D1:E11", "BoxInclusive", "Inclusive", 2)
c.FullSeriesCollection(1).QuartileCalculationInclusiveMedian = True
c.FullSeriesCollection(2).QuartileCalculationInclusiveMedian = True

' ---- treemap and sunburst
Set sh = NewSheet("Tree")
Dim t1: t1 = Array("Fruit", "Fruit", "Fruit", "Veg", "Veg", "Dairy", "Dairy", "Dairy", "Dairy", "Bakery")
Dim t2: t2 = Array("Apple", "Pear", "Plum", "Carrot", "Leek", "Milk", "Cheese", "Butter", "Yogurt", "Bread")
Dim t3: t3 = Array(120, 60, 25, 80, 30, 150, 90, 40, 35, 70)
sh.Range("A1").Value = "Aisle": sh.Range("B1").Value = "Item": sh.Range("C1").Value = "Sales"
For i = 0 To 9: sh.Cells(i + 2, 1).Value = t1(i): sh.Cells(i + 2, 2).Value = t2(i): sh.Cells(i + 2, 3).Value = t3(i): Next
Set c = AddChart(sh, xlTreemap, "A1:C11", "Treemap", "", 0)
Set c = AddChart(sh, xlSunburst, "A1:C11", "Sunburst", "", 1)
sh.Range("E1").Value = "Item": sh.Range("F1").Value = "Sales"
Dim f2: f2 = Array("A", "B", "C", "D", "E", "F", "G")
Dim f3: f3 = Array(50, 40, 30, 20, 10, 8, 5)
For i = 0 To 6: sh.Cells(i + 2, 5).Value = f2(i): sh.Cells(i + 2, 6).Value = f3(i): Next
Set c = AddChart(sh, xlTreemap, "E1:F8", "TreemapFlat", "", 2)
Set c = AddChart(sh, xlSunburst, "E1:F8", "SunburstFlat", "", 3)
' three levels
sh.Range("H1").Value = "Region": sh.Range("I1").Value = "Country": sh.Range("J1").Value = "City": sh.Range("K1").Value = "Pop"
Dim r1: r1 = Array("Europe", "Europe", "Europe", "Europe", "Asia", "Asia", "Asia", "America")
Dim r2: r2 = Array("France", "France", "Italy", "Spain", "Japan", "Japan", "India", "USA")
Dim r3: r3 = Array("Paris", "Lyon", "Rome", "Madrid", "Tokyo", "Osaka", "Delhi", "New York")
Dim r4: r4 = Array(11, 2, 4, 6, 37, 19, 30, 19)
For i = 0 To 7: sh.Cells(i + 2, 8).Value = r1(i): sh.Cells(i + 2, 9).Value = r2(i): sh.Cells(i + 2, 10).Value = r3(i): sh.Cells(i + 2, 11).Value = r4(i): Next
Set c = AddChart(sh, xlTreemap, "H1:K9", "Treemap3", "", 4)
Set c = AddChart(sh, xlSunburst, "H1:K9", "Sunburst3", "", 5)

' ---- waterfall and funnel
Set sh = NewSheet("Water")
Dim wc: wc = Array("Start", "Sales", "Returns", "Costs", "Tax", "Loss", "End")
Dim wv: wv = Array(100, 60, -20, -45, -15, -120, -40)
sh.Range("A1").Value = "Step": sh.Range("B1").Value = "Amount"
For i = 0 To 6: sh.Cells(i + 2, 1).Value = wc(i): sh.Cells(i + 2, 2).Value = wv(i): Next
Set c = AddChart(sh, xlWaterfall, "A1:B8", "WaterNeg", "", 0)
c.FullSeriesCollection(1).Points(1).IsTotal = True
c.FullSeriesCollection(1).Points(7).IsTotal = True
Set c = AddChart(sh, xlWaterfall, "A1:B8", "WaterPlain", "", 1)
sh.Range("D1").Value = "Stage": sh.Range("E1").Value = "Count"
Dim fc: fc = Array("Visits", "Sign-ups", "Trials", "Orders", "Repeat", "Advocates")
Dim fv: fv = Array(1200, 800, 450, 300, 120, 30)
For i = 0 To 5: sh.Cells(i + 2, 4).Value = fc(i): sh.Cells(i + 2, 5).Value = fv(i): Next
Set c = AddChart(sh, xlFunnel, "D1:E7", "Funnel6", "", 2)
On Error GoTo 0

wb.Worksheets(1).Activate
wb.Worksheets(1).Range("A1").Select
wb.SaveAs pngDir & "\synthetic-chartex-zoo.xlsx", xlOpenXMLWorkbook

' ---- what COM tells
Dim ws, co, ch, s, p, ax, g, t
On Error Resume Next
For Each ws In wb.Worksheets
  For Each co In ws.ChartObjects
    Set ch = co.Chart
    f.WriteLine ws.Name & vbTab & co.Name & vbTab & "chart" & vbTab & ch.ChartType & vbTab & co.Width & vbTab & co.Height
    Err.Clear
    t = ch.PlotArea.InsideLeft & vbTab & ch.PlotArea.InsideTop & vbTab & ch.PlotArea.InsideWidth & vbTab & ch.PlotArea.InsideHeight
    If Err.Number = 0 Then f.WriteLine ws.Name & vbTab & co.Name & vbTab & "plot" & vbTab & t
    For k = 1 To 2
      Err.Clear
      Set ax = Nothing
      Set ax = ch.Axes(k, 1)
      If Err.Number = 0 And Not ax Is Nothing Then
        Err.Clear
        t = ax.MinimumScale & vbTab & ax.MaximumScale & vbTab & ax.MajorUnit
        If Err.Number <> 0 Then t = "-"
        Err.Clear
        g = ax.BinsType & vbTab & ax.BinWidthValue & vbTab & ax.BinsCountValue
        If Err.Number <> 0 Then g = "-"
        f.WriteLine ws.Name & vbTab & co.Name & vbTab & "axis" & vbTab & k & vbTab & t & vbTab & "bins " & g
      End If
    Next
    For s = 1 To ch.FullSeriesCollection.Count
      Set p = ch.FullSeriesCollection(s)
      Err.Clear
      f.WriteLine ws.Name & vbTab & co.Name & vbTab & "series" & vbTab & s & vbTab & p.Name & vbTab & p.Points.Count & vbTab & Hex(p.Format.Fill.ForeColor.RGB)
      For j = 1 To p.Points.Count
        Err.Clear
        t = ""
        t = p.Points(j).Left & vbTab & p.Points(j).Top & vbTab & p.Points(j).Width & vbTab & p.Points(j).Height
        If Err.Number <> 0 Then t = "-"
        Err.Clear
        g = ""
        g = Replace(p.Points(j).DataLabel.Text, vbLf, "\n")
        f.WriteLine ws.Name & vbTab & co.Name & vbTab & "point" & vbTab & s & vbTab & j & vbTab & t & vbTab & g & vbTab & Hex(p.Points(j).Format.Fill.ForeColor.RGB)
      Next
    Next
    Err.Clear
    ch.Export pngDir & "\" & ws.Name & "-" & co.Name & ".png"
  Next
Next
On Error GoTo 0
f.Close
wb.Close False
x.Quit
