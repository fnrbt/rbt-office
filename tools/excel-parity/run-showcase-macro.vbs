' cscript run-showcase-macro.vbs NAME__ENTRY.xlsm OUT.tsv: opens the macro-enabled workbook in Excel with
' macros allowed, runs ENTRY (the part of the file name after "__"), and writes OUT.tsv (UTF-16):
' a line "status<TAB>ok|error number description<TAB>seconds", then one line per cell of the active
' sheet's used range that has a value or a fill: address, Value2 as text, fill colour (BGR long, or -).
' For the showcase macros (tests/macros/showcase): tests/macros/showcase.fsx --dump writes the same for ours.
' Run through run-excel-vbs-on-vm.sh (a scheduled task in the desktop session).
Set fso = CreateObject("Scripting.FileSystemObject")
inPath = WScript.Arguments(0)
outPath = WScript.Arguments(1)
base = fso.GetBaseName(inPath)
entry = "Main"
If InStr(base, "__") > 0 Then entry = Mid(base, InStr(base, "__") + 2)
Set x = CreateObject("Excel.Application")
x.Visible = False : x.DisplayAlerts = False
x.AutomationSecurity = 1  ' msoAutomationSecurityLow: the workbook's macros may run
Set wb = x.Workbooks.Open(inPath)
started = Timer
On Error Resume Next
x.Run "'" & wb.Name & "'!" & entry
status = "ok"
If Err.Number <> 0 Then status = "error " & Err.Number & " " & Err.Description
Err.Clear
On Error GoTo 0
elapsed = Timer - started
Set out = fso.CreateTextFile(outPath, True, True)
out.WriteLine "status" & vbTab & status & vbTab & FormatNumber(elapsed, 2)
Set s = x.ActiveSheet
For Each c In s.UsedRange.Cells
  v = c.Value2
  If VarType(v) = 10 Then
    t = c.Text
  ElseIf IsEmpty(v) Then
    t = ""
  Else
    t = Replace(Replace(CStr(v), vbTab, " "), vbLf, " ")
  End If
  If c.Interior.Pattern = -4142 Then f = "-" Else f = CStr(c.Interior.Color)
  If t <> "" Or f <> "-" Then out.WriteLine c.Address(False, False) & vbTab & t & vbTab & f
Next
out.Close
wb.Close False
x.Quit
