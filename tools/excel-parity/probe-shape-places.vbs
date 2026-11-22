' cscript probe-shape-places.vbs OUT.tsv BOOK.xlsx ...: where Excel 2019 has every shape (form controls too) of each
' workbook's sheets: book, sheet, name, placement, left, top, width, height (points). A book whose name has
' ".original." gets a row inserted above row 2 and a column before column B first, as xlsx-browser's keep mode does to
' its ".insert." save (run.mjs --keep): Excel's places of its shapes then, against the editor's saved file's.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
For a = 1 To WScript.Arguments.Count - 1
  path = WScript.Arguments(a)
  Set wb = x.Workbooks.Open(path, False, False)
  For Each sh In wb.Worksheets
    If InStr(path, ".original.") > 0 Then
      sh.Rows(2).Insert
      sh.Columns("B").Insert
    End If
    For Each shp In sh.Shapes
      out.WriteLine fso.GetFileName(path) & vbTab & sh.Name & vbTab & shp.Name & vbTab & shp.Placement & vbTab & shp.Left & vbTab & shp.Top & vbTab & shp.Width & vbTab & shp.Height
    Next
  Next
  wb.Close False
Next
out.Close
x.Quit
