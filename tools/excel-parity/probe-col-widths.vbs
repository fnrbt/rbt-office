' cscript probe-col-widths.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel 2019 and writes, for columns A..AD of every
' visible worksheet: sheet, column, ColumnWidth (characters), Width (points) and Hidden, as Excel has them on opening.
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = True : x.DisplayAlerts = False : x.AutomationSecurity = 3
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
For Each s In wb.Worksheets
  If s.Visible = -1 Then
    For c = 1 To 30
      out.WriteLine s.Name & vbTab & c & vbTab & s.Columns(c).ColumnWidth & vbTab & s.Columns(c).Width & vbTab & s.Columns(c).Hidden
    Next
  End If
Next
out.Close
wb.Close False
x.Quit
