' cscript probe-open-heights.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel and writes, a line per row of each sheet's
' used range, sheet, row and Rows(r).Height as Excel fitted it on opening - nothing activated, nothing fitted again.
' Run it in the desktop session (run-vbs-in-session.ps1).
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = True : x.DisplayAlerts = False : x.AutomationSecurity = 3
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
For Each s In wb.Worksheets
  n = s.UsedRange.Row + s.UsedRange.Rows.Count - 1
  For r = 1 To n : out.WriteLine s.Name & vbTab & r & vbTab & s.Rows(r).Height : Next
Next
out.Close
wb.Close False
x.Quit
