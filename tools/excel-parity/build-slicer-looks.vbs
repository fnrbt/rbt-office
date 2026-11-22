' cscript build-slicer-looks.vbs OUT.tsv PNGDIR: builds synthetic-slicer-looks.xlsx in Excel (saved into PNGDIR) - slicers and
' timelines made by Excel itself, for how Excel 2019 orders a slicer's items (sorted up / down / as the data has them),
' where and how it shows the items with no data (crossFilter) and its header's buttons. Pictures with shoot-sheets.ps1, compared in tools/xlsx-browser `slicers-excel-parity`.
' OUT gets one line a slicer: sheet, name, caption, sort, cross filter, its items in Excel's order
' (name:selected:has data).
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Do While wb.Worksheets.Count < 3: wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count): Loop
Set data = wb.Worksheets(1): data.Name = "Data"
Set rep = wb.Worksheets(2): rep.Name = "Slicers"
Set tl = wb.Worksheets(3): tl.Name = "Timelines"
data.Range("A1:D1").Value = Array("Date", "Region", "Product", "Amount")
' Regions in data order North, South, East, West, Central; West and Central sell no Apples.
regions = Array("North", "South", "East", "West", "Central")
products = Array("Pears", "Apples", "Plums", "Figs")
r = 2
For m = 0 To 17
  For k = 0 To 4
    p = (m + k) Mod 4
    If (k = 3 Or k = 4) And p = 1 Then p = 0
    data.Cells(r, 1).Value = DateSerial(2012, 1 + m, 1 + (k * 5))
    data.Cells(r, 2).Value = regions(k)
    data.Cells(r, 3).Value = products(p)
    data.Cells(r, 4).Value = 10 + m + k
    r = r + 1
  Next
Next
data.ListObjects.Add 1, data.Range("A1:D" & (r - 1)), , 1
data.ListObjects(1).Name = "Sales"
Set pc = wb.PivotCaches.Create(1, "Sales")
Dim pts(9)
For i = 0 To 3
  Set pts(i) = pc.CreatePivotTable(rep.Cells(1 + i * 12, 20), "PT" & (i + 1))
  pts(i).PivotFields("Region").Orientation = 1
  pts(i).AddDataField pts(i).PivotFields("Amount"), "Sum of Amount", -4157
Next
For i = 4 To 9
  Set pts(i) = pc.CreatePivotTable(tl.Cells(1 + (i - 4) * 12, 20), "PT" & (i + 1))
  pts(i).PivotFields("Region").Orientation = 1
  pts(i).AddDataField pts(i).PivotFields("Amount"), "Sum of Amount", -4157
Next
' Product: Apples only, on PT1 .. PT4 - West and Central have no data in the Region slicers.
Set prod = wb.SlicerCaches.Add2(pts(0), "Product")
prod.PivotTables.AddPivotTable pts(1)
prod.PivotTables.AddPivotTable pts(2)
prod.PivotTables.AddPivotTable pts(3)
Set s = prod.Slicers.Add(rep, , "ProductS", "Product", 10, 10, 144, 190)
For Each it In prod.SlicerItems
  If it.Name <> "Apples" Then it.Selected = False
Next
names = Array("RegionUp", "RegionDown", "RegionInPlace", "RegionHide")
captions = Array("Sorted up", "Sorted down", "No data in place", "Hide no data")
sorts = Array(2, 3, 2, 2)
crosses = Array(2, 2, 3, 4)
For i = 0 To 3
  Set sc = wb.SlicerCaches.Add2(pts(i), "Region")
  ' (Not every way of setting these is taken: the cache's own, else its level's. Data source order is for OLAP only.)
  On Error Resume Next
  sc.SortItems = CLng(sorts(i))
  If Err.Number <> 0 Then Err.Clear: sc.SlicerCacheLevels(1).SortItems = CLng(sorts(i))
  If Err.Number <> 0 Then out.WriteLine "sort " & names(i) & vbTab & Err.Description: Err.Clear
  sc.CrossFilterType = CLng(crosses(i))
  If Err.Number <> 0 Then Err.Clear: sc.SlicerCacheLevels(1).CrossFilterType = CLng(crosses(i))
  If Err.Number <> 0 Then out.WriteLine "cross " & names(i) & vbTab & Err.Description: Err.Clear
  On Error GoTo 0
  Set s = sc.Slicers.Add(rep, , names(i), captions(i), 10 + 200 * ((i + 1) \ 3), 10 + 160 * ((i + 1) Mod 3), 144, 190)
  If i = 0 Then sc.SlicerItems("East").Selected = False
Next
' A slicer whose items are all selected (the clear-filter button greyed), in two columns (PT5's regions).
Set sc = wb.SlicerCaches.Add2(pts(4), "Region")
Set s = sc.Slicers.Add(rep, , "ProductAll", "Two columns", 410, 330, 200, 120)
s.NumberOfColumns = 2
' (Timelines: SlicerCaches.Add2(..., xlTimeline) is refused by Excel 2019 from a script - 'Invalid procedure call' -
' so the timeline cases are variants of synthetic-slicers.xlsx's timeline, make-slicers-fixture.py --timelines.)
For Each sc In wb.SlicerCaches
  For Each s In sc.Slicers
    line = s.Parent.Name & vbTab & s.Name & vbTab & s.Caption
    On Error Resume Next
    line = line & vbTab & sc.SortItems & vbTab & sc.CrossFilterType
    items = ""
    For Each it In sc.SlicerItems
      items = items & it.Name & ":" & it.Selected & ":" & it.HasData & "|"
    Next
    line = line & vbTab & items
    On Error GoTo 0
    out.WriteLine line
  Next
Next
rep.Activate
wb.SaveAs pngDir & "\synthetic-slicer-looks.xlsx", 51
out.Close
wb.Close False: x.Quit
