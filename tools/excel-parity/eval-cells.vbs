' cscript eval-cells.vbs CELLS.tsv OUT.tsv: a probe of how Excel calculates. CELLS.tsv (UTF-16) has one
' line per cell: sheet, address, what is typed into it (a formula starting with "=", a value, or
' "{=...}" for an array formula), tab-separated. The cells are typed into a new workbook (its sheets
' made as named), the workbook recalculated fully (CalculateFull), and OUT.tsv (UTF-16) gets one line
' per cell: sheet, address, type (n number, s text, b boolean, e error, x empty), value (numbers with
' 15 significant digits), what the cell shows; then a line "CIRCULAR <sheet> <address>" for each
' sheet whose CircularReference Excel reports. Run it in the desktop session (run-vbs-in-session.ps1).
Set fso = CreateObject("Scripting.FileSystemObject")
inPath = WScript.Arguments(0)
outPath = WScript.Arguments(1)
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False
Set wb = x.Workbooks.Add
Do While wb.Worksheets.Count > 1 : wb.Worksheets(wb.Worksheets.Count).Delete : Loop
first = True
Set lines = CreateObject("Scripting.Dictionary")
Set f = fso.OpenTextFile(inPath, 1, False, -1)
n = 0
Do While Not f.AtEndOfStream
  line = f.ReadLine
  If Len(line) > 0 Then
    parts = Split(line, vbTab, 3)
    If UBound(parts) = 2 Then
      Set s = Nothing
      On Error Resume Next
      Set s = wb.Worksheets(parts(0))
      On Error GoTo 0
      If s Is Nothing Then
        If first Then
          Set s = wb.Worksheets(1)
        Else
          Set s = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
        End If
        s.Name = parts(0)
      End If
      first = False
      If Left(parts(2), 2) = "{=" Then
        s.Range(parts(1)).FormulaArray = Mid(parts(2), 2, Len(parts(2)) - 2)
      Else
        s.Range(parts(1)).Formula = parts(2)
      End If
      lines(n) = Array(parts(0), parts(1)) : n = n + 1
    End If
  End If
Loop
f.Close
x.CalculateFull
Set out = fso.OpenTextFile(outPath, 2, True, -1)
For i = 0 To n - 1
  Set c = wb.Worksheets(lines(i)(0)).Range(lines(i)(1))
  v = c.Value2
  If VarType(v) = 10 Then
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
  out.WriteLine lines(i)(0) & vbTab & lines(i)(1) & vbTab & t & vbTab & s2 & vbTab & c.Text
Next
For Each s In wb.Worksheets
  Set r = Nothing
  On Error Resume Next
  Set r = s.CircularReference
  On Error GoTo 0
  If Not r Is Nothing Then out.WriteLine "CIRCULAR" & vbTab & s.Name & vbTab & r.Address(False, False)
Next
out.Close
wb.Close False
x.Quit
