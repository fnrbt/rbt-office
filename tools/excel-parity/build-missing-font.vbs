' cscript build-missing-font.vbs OUT.tsv DIR: builds synthetic-missing-font.xlsx and synthetic-missing-font-twcen.xlsx in
' Excel (saved into DIR) - for the face Excel 2019 draws a cell's text in when the cell names a font Windows does not
' have (the e-commerce dashboard's KPI YOY cells in Lato bold, shown in its linked pictures), and how a linked
' picture (camera) of such cells is drawn. Sheet "Fonts": column A the font asked for, B its text in that font
' regular, C bold; rows of Lato, a made-up name, and installed faces to compare (Calibri, Arial, Tw Cen MT). Sheet
' "Camera": linked pictures of Fonts!B2:C2 (Lato) and Fonts!B5:C5 (Calibri) at their own size and squeezed to 0.85 of
' their width. The second book is the same with Tw Cen MT as the theme's body font (its Normal style's font).
' OUT gets one line a book: the Normal font, and a line a picture: its range, left, top, width, height (points).
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
dir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = True: x.DisplayAlerts = False
books = Array("synthetic-missing-font.xlsx", "synthetic-missing-font-twcen.xlsx")
fonts = Array("Lato", "Rbt No Such Face", "Lato", "Calibri", "Arial", "Tw Cen MT")
For b = 0 To 1
  Set wb = x.Workbooks.Add()
  If b = 1 Then wb.Theme.ThemeFontScheme.MinorFont.Item(1).Name = "Tw Cen MT"
  Set ws = wb.Worksheets(1): ws.Name = "Fonts"
  ws.Columns("A").ColumnWidth = 18
  ws.Columns("B:C").ColumnWidth = 22
  ws.Range("A1:C1").Value = Array("Font", "Regular", "Bold")
  For i = 0 To 5
    r = i + 2
    ws.Cells(r, 1).Value = fonts(i)
    For c = 2 To 3
      ws.Cells(r, c).Value = ChrW(9650) & "20.62% Hamburgevons 0123456789"
      ws.Cells(r, c).Font.Name = fonts(i)
      ws.Cells(r, c).Font.Size = 11
      ws.Cells(r, c).Font.Bold = (c = 3)
    Next
  Next
  ' (Row 4 again Lato, 16 pt: the size kept.)
  ws.Range("B4:C4").Font.Size = 16
  Set cam = wb.Worksheets.Add(, ws): cam.Name = "Camera"
  cam.Activate
  ranges = Array("B2:C2", "B5:C5", "B2:C2", "B5:C5")
  For k = 0 To 3
    ws.Range(ranges(k)).Copy
    cam.Range("B" & (2 + 4 * k)).Select
    cam.Pictures.Paste(True).Select
    Set p = cam.Shapes(cam.Shapes.Count)
    p.LockAspectRatio = 0
    If k >= 2 Then p.Width = p.Width * 0.85
    out.WriteLine books(b) & vbTab & "Fonts!" & ranges(k) & vbTab & p.Left & vbTab & p.Top & vbTab & p.Width & vbTab & p.Height
  Next
  x.CutCopyMode = False
  out.WriteLine books(b) & vbTab & "Normal" & vbTab & wb.Styles("Normal").Font.Name
  ws.Activate
  wb.SaveAs dir & "\" & books(b), 51
  wb.Close False
Next
out.Close
x.Quit
