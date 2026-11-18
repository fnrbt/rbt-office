' cscript run-cases.vbs DOCS OUT.tsv: for every .bas of DOCS (a case's module with RbtParityMain, the case's
' Main inside an error trap, written by tools/vba-parity/parity.fsx) adds a blank workbook in Excel, puts the
' module into its VBA project (needs "Trust access to the VBA project object model": the per-user registry
' value HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1), runs it, and writes
' to OUT.tsv (UTF-16) one line per cell that has a value or a format of its own, of every sheet:
'   case, sheet, address, type (n s b e), value (numbers as 0.00000000000000E+00), text, number format,
'   bold, fill (Interior.Color or none), font colour
' and per case a line "case #RUN status". Run it in the desktop session (a scheduled task with /IT).
Set fso = CreateObject("Scripting.FileSystemObject")
docs = WScript.Arguments(0)
logPath = WScript.Arguments(1) & ".progress.txt"
Sub Progress(text)
  Set lf = fso.OpenTextFile(logPath, 8, True)
  lf.WriteLine Now & " " & text
  lf.Close
End Sub
Progress "start"
Set out = fso.CreateTextFile(WScript.Arguments(1), True, True)
Set x = CreateObject("Excel.Application") : x.Visible = True : x.DisplayAlerts = False  ' (visible: a dialog that stops a case can be seen on the VM)
Progress "excel started"
x.AutomationSecurity = 1  ' msoAutomationSecurityLow: the cases' macros run
For Each f In fso.GetFolder(docs).Files
  If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
    name = fso.GetBaseName(f.Name)
    code = fso.OpenTextFile(f.Path, 1).ReadAll
    Progress name & " add"
    Err.Clear
    On Error Resume Next
    Set wb = x.Workbooks.Add(-4167)  ' one worksheet
    If Err.Number = 0 Then wb.VBProject.VBComponents.Add(1).CodeModule.AddFromString code
    If Err.Number <> 0 Then
      out.WriteLine name & vbTab & "#RUN" & vbTab & "open-error " & Err.Number & " " & Err.Description
      Err.Clear
    Else
      Progress name & " run"
      x.Run "RbtParityMain"
      Progress name & " ran " & Err.Number
      If Err.Number <> 0 Then
        out.WriteLine name & vbTab & "#RUN" & vbTab & "run-error " & Err.Number & " " & Err.Description
        Err.Clear
      Else
        out.WriteLine name & vbTab & "#RUN" & vbTab & "ok"
      End If
      For Each s In wb.Worksheets
        For Each c In s.UsedRange.Cells
          v = c.Value2
          fill = "none" : If c.Interior.ColorIndex <> -4142 Then fill = CStr(c.Interior.Color)
          If Not (IsEmpty(v) And fill = "none" And Not c.Font.Bold And c.Font.Color = 0) Then
            If VarType(v) = 10 Then
              t = "e" : sv = c.Text
            ElseIf IsEmpty(v) Then
              t = "x" : sv = ""
            ElseIf VarType(v) = vbBoolean Then
              t = "b" : sv = CStr(v)
            ElseIf VarType(v) = vbString Then
              t = "s" : sv = Replace(Replace(Replace(v, vbTab, "\t"), vbCr, "\r"), vbLf, "\n")
            Else
              t = "n" : sv = x.WorksheetFunction.Text(v, "0.00000000000000E+00")
            End If
            out.WriteLine name & vbTab & s.Name & vbTab & c.Address(False, False) & vbTab & t & vbTab & sv & vbTab & Replace(c.Text, vbTab, " ") & vbTab & c.NumberFormat & vbTab & CStr(CBool(c.Font.Bold)) & vbTab & fill & vbTab & CStr(c.Font.Color)
          End If
        Next
      Next
      Progress name & " dumped"
      wb.Close False
    End If
    On Error GoTo 0
  End If
Next
x.Quit
out.WriteLine "#DONE"
out.Close
