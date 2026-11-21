' cscript probe-display.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel (read-only), recalculates fully and writes
' to OUT.tsv (UTF-16) one line per cell of each sheet's used range: sheet, address, Value2, Text, and what
' conditional formats make of it - DisplayFormat.Interior.Color and DisplayFormat.Font.Color (BGR integers).
' Run it in the desktop session (run-vbs-in-session.ps1).
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
x.CalculateFull
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
For Each s In wb.Worksheets
  For Each c In s.UsedRange.Cells
    v = c.Value2
    If VarType(v) = 10 Then v = c.Text
    out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & CStr(v) & vbTab & c.Text & vbTab & CStr(c.DisplayFormat.Interior.Color) & vbTab & CStr(c.DisplayFormat.Font.Color)
  Next
Next
out.Close
wb.Close False
x.Quit
