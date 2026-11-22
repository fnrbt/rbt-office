' cscript build-pivot-styles.vbs OUT.tsv PNGDIR: builds synthetic-pivot-styles.xlsx in Excel (saved into PNGDIR) - pivot
' tables of one data sheet in built-in pivot styles, layouts (compact, outline, tabular), with page fields, subtotals
' at the top and the bottom, two levels of rows and of columns, banded rows / columns, row / column headers off - and
' writes what Excel shows of every cell around each (Range.DisplayFormat: fill, font, the four edges), the reference
' for how a pivot style is laid over a pivot table (Model.pivotCellLook; docs/excel-parity-tracker.md "Pivot tables").
' Lines: sheet, address, text, fill RRGGBB or "-", bold, italic, font RRGGBB, then top, bottom, left, right edges as
' linestyle/weight/RRGGBB or "-"; a sheet's first line "#", its pivot's TableRange1, TableRange2 and style.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Do While wb.Worksheets.Count > 1: wb.Worksheets(wb.Worksheets.Count).Delete: Loop
Set data = wb.Worksheets(1): data.Name = "Data"
heads = Array("Region", "Product", "Year", "Quarter", "Sales", "Units")
For i = 0 To 5: data.Cells(1, i + 1).Value = heads(i): Next
regions = Array("East", "North", "South", "West"): products = Array("Apples", "Pears", "Plums")
r = 2
For y = 2023 To 2024
  For q = 1 To 4
    For a = 0 To 3
      For p = 0 To 2
        data.Cells(r, 1).Value = regions(a): data.Cells(r, 2).Value = products(p)
        data.Cells(r, 3).Value = y: data.Cells(r, 4).Value = "Q" & q
        data.Cells(r, 5).Value = 100 + ((r * 37) Mod 400): data.Cells(r, 6).Value = 1 + ((r * 13) Mod 20)
        r = r + 1
      Next
    Next
  Next
Next
Set src = data.Range("A1:F" & (r - 1))

' name, style, layout (0 compact, 1 tabular, 2 outline), row fields, column fields, page field, data fields,
' row stripes, column stripes, row headers, column headers, subtotals at bottom
cases = Array( _
  Array("Compact", "PivotStyleLight16", 0, "Year,Quarter", "Region", "Product", "Sales", False, False, True, True, False), _
  Array("Banded", "PivotStyleMedium9", 0, "Year,Quarter", "Region", "Product", "Sales", True, True, True, True, False), _
  Array("Outline", "PivotStyleMedium2", 2, "Region,Product", "Year", "", "Sales", False, False, True, True, True), _
  Array("Tabular", "PivotStyleDark2", 1, "Region,Product", "", "", "Sales,Units", False, False, True, True, True), _
  Array("Flat", "PivotStyleLight8", 0, "Product", "", "", "Sales", False, False, False, False, False), _
  Array("Deep", "PivotStyleMedium4", 0, "Region,Product,Quarter", "", "Year", "Sales", True, False, True, True, False), _
  Array("ColDeep", "PivotStyleMedium14", 0, "Region", "Year,Quarter", "", "Sales", False, True, True, True, False), _
  Array("Dark7", "PivotStyleDark7", 0, "Year,Quarter", "Region", "", "Sales", False, False, True, True, False), _
  Array("Light1", "PivotStyleLight1", 0, "Year,Quarter", "Region", "Product", "Sales", True, False, True, True, False), _
  Array("Light20", "PivotStyleLight20", 2, "Region,Product", "Year", "", "Sales", False, False, True, True, False), _
  Array("Medium28", "PivotStyleMedium28", 0, "Year,Quarter", "Region", "", "Sales,Units", False, False, True, True, False), _
  Array("Dark12", "PivotStyleDark12", 1, "Year,Quarter", "Region", "Product", "Sales", True, True, True, True, True))
Set pc = wb.PivotCaches.Create(1, src)
For Each c In cases
  Set sh = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count)): sh.Name = c(0)
  Set pt = pc.CreatePivotTable(sh.Range("B4"), "P" & c(0))
  pt.ManualUpdate = True
  If c(5) <> "" Then pt.PivotFields(c(5)).Orientation = 3
  pos = 1
  For Each f In Split(c(3), ","): pt.PivotFields(f).Orientation = 1: pt.PivotFields(f).Position = pos: pos = pos + 1: Next
  pos = 1
  If c(4) <> "" Then
    For Each f In Split(c(4), ","): pt.PivotFields(f).Orientation = 2: pt.PivotFields(f).Position = pos: pos = pos + 1: Next
  End If
  For Each f In Split(c(6), ","): pt.AddDataField pt.PivotFields(f), "Sum of " & f, -4157: Next
  pt.ManualUpdate = False
  pt.RowAxisLayout c(2)
  If c(11) Then pt.SubtotalLocation 2 Else pt.SubtotalLocation 1
  pt.TableStyle2 = c(1)
  pt.ShowTableStyleRowStripes = c(7): pt.ShowTableStyleColumnStripes = c(8)
  pt.ShowTableStyleRowHeaders = c(9): pt.ShowTableStyleColumnHeaders = c(10)
  pt.RefreshTable
  Set t2 = pt.TableRange2
  out.WriteLine sh.Name & vbTab & "#" & vbTab & pt.TableRange1.Address(False, False) & vbTab & t2.Address(False, False) & vbTab & pt.TableStyle2
  For rr = t2.Row - 1 To t2.Row + t2.Rows.Count
    For cc = t2.Column - 1 To t2.Column + t2.Columns.Count
      Set cell = sh.Cells(rr, cc): Set d = cell.DisplayFormat
      line = sh.Name & vbTab & cell.Address(False, False) & vbTab & cell.Text & vbTab
      If d.Interior.Pattern = -4142 Then line = line & "-" Else line = line & hexOf(d.Interior.Color)
      line = line & vbTab & d.Font.Bold & vbTab & d.Font.Italic & vbTab & hexOf(d.Font.Color)
      For Each e In Array(8, 9, 7, 10)
        Set b = d.Borders(e)
        If b.LineStyle = -4142 Then line = line & vbTab & "-" Else line = line & vbTab & b.LineStyle & "/" & b.Weight & "/" & hexOf(b.Color)
      Next
      out.WriteLine line
    Next
  Next
Next
data.Activate
wb.SaveAs pngDir & "\synthetic-pivot-styles.xlsx", 51
out.Close
wb.Close False: x.Quit

Function hexOf(v)
  If IsNull(v) Then hexOf = "?": Exit Function
  hexOf = Right("0" & Hex(v And 255), 2) & Right("0" & Hex((v \ 256) And 255), 2) & Right("0" & Hex((v \ 65536) And 255), 2)
End Function
