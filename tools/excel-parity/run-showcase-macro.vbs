' cscript run-showcase-macro.vbs NAME__ENTRY[__BEFORE].xlsm OUT.tsv: opens the macro-enabled workbook in Excel
' with macros allowed, runs BEFORE (a setup Sub) then ENTRY, and writes OUT.tsv (UTF-16):
' a line "status<TAB>ok|error number description<TAB>seconds", then one line per cell of the active
' sheet's used range that has a value or a fill: address, Value2 as text, fill colour (BGR long, or -).
' For the showcase macros (tests/macros/showcase): tests/macros/showcase.fsx --dump writes the same for ours.
' Run through run-excel-vbs-on-vm.sh (a scheduled task in the desktop session).
Set fso = CreateObject("Scripting.FileSystemObject")
inPath = WScript.Arguments(0)
outPath = WScript.Arguments(1)
base = fso.GetBaseName(inPath)
parts = Split(base, "__")
entry = "Main" : before = ""
If UBound(parts) >= 1 Then entry = parts(1)
If UBound(parts) >= 2 Then before = parts(2)
Set x = CreateObject("Excel.Application")
' (Visible: a macro that sets ActiveWindow.Zoom needs a window; an invisible Excel has none and stops on error 91.)
x.Visible = True : x.DisplayAlerts = False
x.AutomationSecurity = 1  ' msoAutomationSecurityLow: the workbook's macros may run
Set wb = x.Workbooks.Open(inPath)
status = ""
started = Timer
On Error Resume Next
If before <> "" Then RunMacro before
If status = "" Then RunMacro entry
If status = "" Then status = "ok"
Err.Clear
On Error GoTo 0
elapsed = Timer - started
Set out = fso.CreateTextFile(outPath, True, True)
out.WriteLine "status" & vbTab & status & vbTab & FormatNumber(elapsed, 2)
Set s = x.ActiveSheet
' (Each cell is a few COM calls: a large picture is sampled - every cell of a range of up to 4000, else a
' diagonal grid of about 4000 cells - so the dump ends within the task's time.)
' (Within A1:GR250: a sheet formatted whole has a used range of every cell.)
Set area = x.Intersect(s.UsedRange, s.Range("A1:GR250"))
total = area.Cells.Count
stepSize = 1
If total > 4000 Then stepSize = Int(total / 4000) + 1
If stepSize > 1 Then out.WriteLine "sampled" & vbTab & stepSize & vbTab & "-"
For Each c In area.Cells
  If stepSize = 1 Or ((c.Row * 7 + c.Column) Mod stepSize) = 0 Then
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
  End If
Next
out.Close
wb.Close False
x.Quit

' Runs a Sub by name; a Sub in a module of the same name (Main in Main) needs Module.Sub.
Sub RunMacro(name)
  Err.Clear
  x.Run "'" & wb.Name & "'!" & name
  If Err.Number = 1004 Then
    Err.Clear
    x.Run "'" & wb.Name & "'!" & name & "." & name
  End If
  If Err.Number <> 0 Then status = "error " & Err.Number & " " & Err.Description & " (" & name & ")"
End Sub
