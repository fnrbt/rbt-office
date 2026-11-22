' cscript build-slicer-cut.vbs OUT.tsv DIR: builds synthetic-slicer-cut.xlsx and synthetic-slicer-cut-twcen.xlsx in Excel
' (saved into DIR) - for where Excel 2019 cuts a slicer's button text and its caption, and how. A table "Items" of
' three labels ("00000000", "Central", "Ab Cd Efghij") with a table slicer cache on its column; the same cache shown
' by slicers one pixel wider each: sheets "B1" / "B2" (caption "C") 40 .. 139 px, sheets "C1" / "C2" (caption
' "00000000") 100 .. 199 px. The second book is the same in a theme whose body font is Tw Cen MT (whether the
' slicer's text is in the theme's body font and is cut by its widths). Pictures with shoot-sheets.ps1 at 100 %.
' OUT gets one line a slicer: book, sheet, name, its width asked (px), Left, Top, Width, Height (points) as Excel
' has them.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
dir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
books = Array("synthetic-slicer-cut.xlsx", "synthetic-slicer-cut-twcen.xlsx")
For b = 0 To 1
  Set wb = x.Workbooks.Add()
  If b = 1 Then wb.Theme.ThemeFontScheme.MinorFont.Item(1).Name = "Tw Cen MT"
  Set data = wb.Worksheets(1): data.Name = "Data"
  data.Range("A1").Value = "Item"
  data.Range("A2").Value = "'00000000"
  data.Range("A3").Value = "Central"
  data.Range("A4").Value = "Ab Cd Efghij"
  Set lo = data.ListObjects.Add(1, data.Range("A1:A4"), , 1)
  lo.Name = "Items"
  Set sc = wb.SlicerCaches.Add2(lo, "Item")
  sheets = Array("B1", "B2", "C1", "C2")
  firsts = Array(40, 90, 100, 150)
  For s = 0 To 3
    Set ws = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheets(s)
    For i = 0 To 49
      w = firsts(s) + i
      nm = sheets(s) & "_" & w
      If s < 2 Then cap = "C" Else cap = "00000000"
      ' (10 a row, 160 px apart; rows 135 px apart. Points are 3/4 of a pixel.)
      l = (10 + (i Mod 10) * 160) * 0.75
      t = (10 + (i \ 10) * 135) * 0.75
      Set sl = sc.Slicers.Add(ws, , nm, cap, t, l, w * 0.75, 125 * 0.75)
      sl.Style = "SlicerStyleLight1"
      out.WriteLine books(b) & vbTab & ws.Name & vbTab & nm & vbTab & w & vbTab & sl.Left & vbTab & sl.Top & vbTab & sl.Width & vbTab & sl.Height
    Next
  Next
  wb.Worksheets("B1").Activate
  wb.SaveAs dir & "\" & books(b), 51
  wb.Close False
Next
out.Close
x.Quit
