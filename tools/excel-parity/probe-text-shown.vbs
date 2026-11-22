' cscript probe-text-shown.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel (shown), and for each sheet - activated, at
' its own zoom - writes a line per cell of its used range: sheet, address, Value2, Text (what the cell shows: `#`s
' where its number does not fit). Run it in the desktop session (run-vbs-in-session.ps1).
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = True : x.DisplayAlerts = False : x.AutomationSecurity = 3
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
For Each s In wb.Worksheets
  s.Activate
  For Each c In s.UsedRange.Cells
    out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & CStr(c.Value2) & vbTab & c.Text & vbTab & x.ActiveWindow.Zoom
  Next
Next
out.Close
wb.Close False
x.Quit
