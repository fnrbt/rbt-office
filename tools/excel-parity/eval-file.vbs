' cscript eval-file.vbs BOOK.xlsx OUT.tsv [MODE]: opens BOOK.xlsx in Excel (read-only, links not updated)
' and writes to OUT.tsv (UTF-16) one line per formula cell: sheet, address, type (n number, s text,
' b boolean, e error, x empty), value (numbers with 15 significant digits), what the cell shows, and for a
' number its exact value "hi:lo:e" ((hi*2^26 + lo) * 2^e, as eval-cells.vbs); then "CIRCULAR <sheet>
' <address>" for each sheet whose CircularReference Excel reports, and "ITERATION <on> <count> <change>".
' MODE: "full" (the default) recalculates fully (CalculateFull) before reading, "none" reads the values
' as opened, "calc" recalculates what needs it (Calculate). Run it in the desktop session
' (run-vbs-in-session.ps1; eval-file-on-vm.sh does all of it).
Set fso = CreateObject("Scripting.FileSystemObject")
inPath = WScript.Arguments(0)
outPath = WScript.Arguments(1)
mode = "full"
If WScript.Arguments.Count > 2 Then mode = WScript.Arguments(2)
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
x.AskToUpdateLinks = False
Set wb = x.Workbooks.Open(inPath, 0, True)
If mode = "full" Then x.CalculateFull
If mode = "calc" Then x.Calculate
Set out = fso.OpenTextFile(outPath, 2, True, -1)
For Each s In wb.Worksheets
  Set r = Nothing
  On Error Resume Next
  Set r = s.UsedRange.SpecialCells(-4123)
  If Err.Number <> 0 Then Err.Clear : Set r = Nothing
  On Error GoTo 0
  If Not r Is Nothing Then
    For Each c In r.Cells
      v = c.Value2
      If VarType(v) = 10 Then
        out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & "e" & vbTab & c.Text & vbTab & c.Text
      ElseIf IsEmpty(v) Then
        out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & "x" & vbTab & "" & vbTab & ""
      ElseIf VarType(v) = vbBoolean Then
        out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & "b" & vbTab & CStr(v) & vbTab & c.Text
      ElseIf VarType(v) = vbString Then
        out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & "s" & vbTab & Replace(Replace(Replace(v, vbTab, " "), vbCr, " "), vbLf, " ") & vbTab & c.Text
      Else
        out.WriteLine s.Name & vbTab & c.Address(False, False) & vbTab & "n" & vbTab & x.WorksheetFunction.Text(v, "0.00000000000000E+00") & vbTab & c.Text & vbTab & Exact(CDbl(v))
      End If
    Next
  End If
  Set cr = Nothing
  On Error Resume Next
  Set cr = s.CircularReference
  On Error GoTo 0
  If Not cr Is Nothing Then out.WriteLine "CIRCULAR" & vbTab & s.Name & vbTab & cr.Address(False, False)
Next
out.WriteLine "ITERATION" & vbTab & CStr(x.Iteration) & vbTab & CStr(x.MaxIterations) & vbTab & CStr(x.MaxChange)
out.Close
wb.Close False
x.Quit

' v as hi:lo:e with v = (hi*2^26 + lo) * 2^e exactly (as eval-cells.vbs).
Function Exact(v)
  Dim m, e, hi, lo, sg
  If v = 0 Then Exact = "0:0:0" : Exit Function
  sg = "" : m = v
  If m < 0 Then sg = "-" : m = -m
  e = 0
  Do While m <> Fix(m) : m = m * 2 : e = e - 1 : Loop
  Do While m >= 9007199254740992 : m = m / 2 : e = e + 1 : Loop
  hi = Fix(m / 67108864) : lo = m - hi * 67108864
  Exact = sg & CStr(hi) & ":" & CStr(lo) & ":" & CStr(e)
End Function
