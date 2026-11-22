' cscript probe-wrap-fit.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel and writes, a line per row of each sheet's
' used range: sheet, row, Rows(r).Height as Excel fitted it on opening, after Rows.AutoFit with the sheet shown at
' its own zoom, and after Rows.AutoFit at 100 % (make-wrap-fit-fixture.py). Run it in the desktop session
' (run-vbs-in-session.ps1).
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = True : x.DisplayAlerts = False : x.AutomationSecurity = 3
x.ScreenUpdating = True
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
For Each s In wb.Worksheets
  n = s.UsedRange.Row + s.UsedRange.Rows.Count - 1
  ReDim h0(n), h1(n), h2(n)
  For r = 1 To n : h0(r) = s.Rows(r).Height : Next
  s.Activate
  z = x.ActiveWindow.Zoom
  s.Rows("1:" & n).AutoFit
  For r = 1 To n : h1(r) = s.Rows(r).Height : Next
  x.ActiveWindow.Zoom = 100
  s.Rows("1:" & n).AutoFit
  For r = 1 To n : h2(r) = s.Rows(r).Height : Next
  x.ActiveWindow.Zoom = z
  For r = 1 To n : out.WriteLine s.Name & vbTab & r & vbTab & h0(r) & vbTab & h1(r) & vbTab & h2(r) & vbTab & z : Next
Next
out.Close
wb.Close False
x.Quit
