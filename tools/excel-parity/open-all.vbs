' cscript open-all.vbs DIR: opens every .xlsx of DIR\docs in Excel, read only, and appends to DIR\opened.tsv:
' name, ok, sheets, used cells, formulas, characters of text, sum of numbers (or ERROR, number, message; a
' workbook Excel had to repair says "repaired"). Run it in the desktop session (a scheduled task with /IT).
' Compare each .save, .edit and .insert.xlsx that `xlsx-browser/run.mjs --keep` keeps with its .original.xlsx.
Set fso = CreateObject("Scripting.FileSystemObject")
base = WScript.Arguments(0)
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
x.AskToUpdateLinks = False
For Each f In fso.GetFolder(base & "\docs").Files
  Err.Clear
  On Error Resume Next
  Set wb = x.Workbooks.Open(f.Path, 0, True)
  If Err.Number <> 0 Then
    Say f.Name & vbTab & "ERROR" & vbTab & Err.Number & vbTab & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ")
    Err.Clear
  Else
    state = "ok" : If InStr(1, x.Caption & wb.Name, "Repaired", 1) > 0 Then state = "repaired"
    used = 0 : formulas = 0 : chars = 0 : total = 0
    For Each s In wb.Worksheets
      Set r = s.UsedRange
      used = used + r.Cells.Count
      If r.Cells.Count > 1 And r.Cells.Count < 400000 Then
        v = r.Value2 : g = r.HasFormula
        For a = 1 To UBound(v, 1) : For b = 1 To UBound(v, 2)
          If VarType(v(a, b)) = 8 Then chars = chars + Len(v(a, b))
          If VarType(v(a, b)) = 5 Then total = total + v(a, b)
        Next : Next
      End If
      fc = 0 : Err.Clear : fc = r.SpecialCells(-4123).Count : If Err.Number = 0 Then formulas = formulas + fc
      Err.Clear
    Next
    Say f.Name & vbTab & state & vbTab & wb.Worksheets.Count & vbTab & used & vbTab & formulas & vbTab & chars & vbTab & Round(total, 2)
    wb.Close False
  End If
  On Error Goto 0
Next
Say "DONE"
x.Quit

Sub Say(text)
  Set sink = fso.OpenTextFile(base & "\opened.tsv", 8, True)
  sink.WriteLine text
  sink.Close
End Sub
