' cscript probe-pivot-style-order.vbs OUT.tsv PNGDIR: which elements of a pivot style cover which cells of a pivot table,
' and which of them wins where they meet - in Excel 2019. A custom pivot style gives each element its own fill colour
' and its own lines (the edges of its block in one colour, the lines inside it in another); the
' cells' DisplayFormat names the element shown. First each element alone (what it fills, which edges it draws), then
' every two elements that fill one cell or draw one edge together, alone: which of them Excel shows. Three pivot
' tables of one data sheet (compact with the subtotals at the top; outline with the subtotals at the bottom and blank
' lines; tabular), each with two page fields, four row fields and three column fields, every option of the style on.
' Lines: pivot, "cov", element, cell, its fill, its bottom and right edges ("e<id>" an element's edge colour, "i<id>"
' its inside colour, "-" none); pivot, "fill" / "line", the two elements (a line's as id:e or id:i), the cell (and
' edge: 9 bottom, 10 right), the one shown. Element ids are Excel's XlTableStyleElementType. The workbook is saved as
' PNGDIR\pivot-style-order.xlsx.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Do While wb.Worksheets.Count > 1: wb.Worksheets(wb.Worksheets.Count).Delete: Loop
Set data = wb.Worksheets(1): data.Name = "Data"
For f = 1 To 9: data.Cells(1, f).Value = "F" & f: Next
data.Cells(1, 10).Value = "Sales"
r = 2
For n = 0 To 127
  For f = 1 To 9
    data.Cells(r, f).Value = "F" & f & Chr(65 + ((n \ (2 ^ ((f - 1) Mod 7))) Mod 2))
  Next
  data.Cells(r, 10).Value = 10 + (n * 7) Mod 50
  r = r + 1
Next
Set pc = wb.PivotCaches.Create(1, data.Range("A1:J" & (r - 1)))
ids = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27)
Set ts = wb.TableStyles.Add("PeelStyle")
ts.ShowAsAvailablePivotTableStyle = True
ts.ShowAsAvailableTableStyle = False

layouts = Array(Array("PeelCompact", 0, False, False), Array("PeelOutline", 2, True, True), Array("PeelTabular", 1, True, False))
For Each lay In layouts
  Set sh = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count)): sh.Name = lay(0)
  Set pt = pc.CreatePivotTable(sh.Range("B6"), lay(0))
  pt.ManualUpdate = True
  pt.PivotFields("F8").Orientation = 3: pt.PivotFields("F9").Orientation = 3
  For f = 1 To 4: pt.PivotFields("F" & f).Orientation = 1: pt.PivotFields("F" & f).Position = f: Next
  For f = 5 To 7: pt.PivotFields("F" & f).Orientation = 2: pt.PivotFields("F" & f).Position = f - 4: Next
  pt.AddDataField pt.PivotFields("Sales"), "Sum of Sales", -4157
  pt.ManualUpdate = False
  pt.RowAxisLayout lay(1)
  If lay(2) Then pt.SubtotalLocation 2 Else pt.SubtotalLocation 1
  If lay(3) Then pt.PivotFields("F1").LayoutBlankLine = True: pt.PivotFields("F2").LayoutBlankLine = True
  pt.TableStyle2 = "PeelStyle"
  pt.ShowTableStyleRowStripes = True: pt.ShowTableStyleColumnStripes = True
  pt.ShowTableStyleRowHeaders = True: pt.ShowTableStyleColumnHeaders = True: pt.ShowTableStyleLastColumn = True
  pt.RefreshTable
  Set t2 = pt.TableRange2
  out.WriteLine lay(0) & vbTab & "#" & vbTab & pt.TableRange1.Address(False, False) & vbTab & t2.Address(False, False)
  ' Each element alone: the cells it fills and the edges it draws (its block's edges "e", the lines inside it "i").
  Set fills = CreateObject("Scripting.Dictionary"): Set edges = CreateObject("Scripting.Dictionary")
  For Each e In ids
    clearAll
    paint e, True, True
    restyle
    For rr = t2.Row - 1 To t2.Row + t2.Rows.Count
      For cc = t2.Column - 1 To t2.Column + t2.Columns.Count
        Set d = sh.Cells(rr, cc).DisplayFormat
        a = sh.Cells(rr, cc).Address(False, False)
        line = lay(0) & vbTab & "cov" & vbTab & e & vbTab & a
        If d.Interior.Pattern = -4142 Then k = "-" Else k = nameOf(d.Interior.Color): fills(a) = fills(a) & " " & e
        line = line & vbTab & k
        For Each side In Array(9, 10)
          If d.Borders(side).LineStyle = -4142 Then k = "-" Else k = nameOf(d.Borders(side).Color): edges(a & "|" & side) = edges(a & "|" & side) & " " & e & ":" & Left(k, 1)
          line = line & vbTab & k
        Next
        out.WriteLine line
      Next
    Next
  Next
  ' Two elements at a time, on a cell both fill / an edge both draw: which one Excel shows.
  Set done = CreateObject("Scripting.Dictionary")
  For Each a In fills.Keys
    els = Split(Trim(fills(a)), " ")
    For i = 0 To UBound(els) - 1
      For j = i + 1 To UBound(els)
        key = "f" & els(i) & "/" & els(j)
        If Not done.Exists(key) Then
          done(key) = 1
          clearAll: paint CInt(els(i)), True, False: paint CInt(els(j)), True, False: restyle
          out.WriteLine lay(0) & vbTab & "fill" & vbTab & els(i) & vbTab & els(j) & vbTab & a & vbTab & nameOf(sh.Range(a).DisplayFormat.Interior.Color)
        End If
      Next
    Next
  Next
  For Each a In edges.Keys
    parts = Split(a, "|")
    els = Split(Trim(edges(a)), " ")
    For i = 0 To UBound(els) - 1
      For j = i + 1 To UBound(els)
        key = "l" & els(i) & "/" & els(j)
        ei = CInt(Split(els(i), ":")(0)): ej = CInt(Split(els(j), ":")(0))
        If Not done.Exists(key) Then
          done(key) = 1
          clearAll: paint ei, False, True: paint ej, False, True: restyle
          Set b = sh.Range(parts(0)).DisplayFormat.Borders(CInt(parts(1)))
          If b.LineStyle = -4142 Then k = "-" Else k = nameOf(b.Color)
          out.WriteLine lay(0) & vbTab & "line" & vbTab & els(i) & vbTab & els(j) & vbTab & a & vbTab & k
        End If
      Next
    Next
  Next
Next
data.Activate
wb.SaveAs pngDir & "\pivot-style-order.xlsx", 51
out.Close
wb.Close False: x.Quit

' (A pivot table takes its style's changed elements only when it is given the style again.)
Sub restyle()
  pt.TableStyle2 = "PivotStyleLight16": pt.TableStyle2 = "PeelStyle"
End Sub

Sub clearAll()
  Dim e
  For Each e In ids: ts.TableStyleElements(e).Clear: Next
End Sub

Sub paint(e, withFill, withLines)
  Dim el, side
  Set el = ts.TableStyleElements(e)
  If withFill Then el.Interior.Color = colorOf(e, 0)
  If withLines Then
    For Each side In Array(7, 8, 9, 10)
      el.Borders(side).LineStyle = 1: el.Borders(side).Weight = 2: el.Borders(side).Color = colorOf(e, 1)
    Next
    For Each side In Array(11, 12)
      el.Borders(side).LineStyle = 1: el.Borders(side).Weight = 2: el.Borders(side).Color = colorOf(e, 2)
    Next
  End If
End Sub

Function colorOf(e, kind)
  colorOf = RGB(20 + e * 8, 3 + kind * 40, 250 - e * 7)
End Function

Function nameOf(v)
  Dim e
  nameOf = "?" & v
  For Each e In ids
    If colorOf(e, 0) = v Or colorOf(e, 1) = v Then nameOf = "e" & e: Exit Function
    If colorOf(e, 2) = v Then nameOf = "i" & e: Exit Function
  Next
End Function
