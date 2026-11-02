' cscript dump-values.vbs DIR: opens every workbook of DIR\docs in Excel, recalculates it fully
' (CalculateFull) and appends to DIR\values.tsv one line per formula cell:
'   workbook, sheet, address, type (n number, s text, b boolean, e error, x empty), value
' (numbers with 15 significant digits). Run it in the desktop session (a scheduled task with /IT).
' tools/excel-parity/compare-values.fsx compares it with what our engine computes.
Set fso = CreateObject("Scripting.FileSystemObject")
base = WScript.Arguments(0)
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
x.AskToUpdateLinks = False
For Each f In fso.GetFolder(base & "\docs").Files
  Err.Clear
  On Error Resume Next
  Set wb = x.Workbooks.Open(f.Path, 0, True)
  If Err.Number <> 0 Then
    Say f.Name & vbTab & "-" & vbTab & "-" & vbTab & "ERROR" & vbTab & Err.Number
    Err.Clear
  Else
    x.CalculateFull
    For Each s In wb.Worksheets
      Set r = Nothing
      Set r = s.UsedRange.SpecialCells(-4123)  ' xlCellTypeFormulas
      If Err.Number <> 0 Then
        Err.Clear
      Else
        buf = ""
        For Each c In r.Cells
          v = c.Value2
          If VarType(v) = 10 Then  ' (vbError: VBScript has no IsError)
            t = "e" : s2 = c.Text
          ElseIf IsEmpty(v) Then
            t = "x" : s2 = ""
          ElseIf VarType(v) = vbBoolean Then
            t = "b" : s2 = CStr(v)
          ElseIf VarType(v) = vbString Then
            t = "s" : s2 = Replace(Replace(Replace(v, vbTab, " "), vbCr, " "), vbLf, " ")
          Else
            t = "n" : s2 = x.WorksheetFunction.Text(v, "0.00000000000000E+00")
          End If
          buf = buf & f.Name & vbTab & s.Name & vbTab & c.Address(False, False) & vbTab & t & vbTab & s2 & vbCrLf
        Next
        If Len(buf) > 0 Then Say Left(buf, Len(buf) - 2)
      End If
    Next
    wb.Close False
  End If
  On Error GoTo 0
Next
x.Quit
Say "DONE"

Sub Say(line)
  Set out = fso.OpenTextFile(base & "\values.tsv", 8, True, -1)
  out.WriteLine line
  out.Close
End Sub
