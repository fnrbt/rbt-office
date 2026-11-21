' cscript eval-cells.vbs CELLS.tsv OUT.tsv: a probe of how Excel calculates. CELLS.tsv (UTF-16) has one
' line per cell: sheet, address, what is typed into it (a formula starting with "=", a value, or
' "{=...}" for an array formula), tab-separated. The cells are typed into a new workbook (its sheets
' made as named), the workbook recalculated fully (CalculateFull), and OUT.tsv (UTF-16) gets one line
' per cell: sheet, address, type (n number, s text, b boolean, e error, x empty), value (numbers with
' 15 significant digits), what the cell shows; then a line "CIRCULAR <sheet> <address>" for each
' sheet whose CircularReference Excel reports. Run it in the desktop session (run-vbs-in-session.ps1).
' A number's line ends with its exact value, "m*2^e" (m an integer of up to 53 bits, as hi*2^26+lo):
' "hi:lo:e". A line "#ITERATE<tab>count<tab>delta" first turns iterative calculation on.
' A line "#NOCALC" leaves out the full recalculation (the values as typed in, one by one).
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
  If Left(line, 8) = "#ITERATE" Then
    it = Split(line, vbTab)
    x.Iteration = True : x.MaxIterations = CInt(it(1)) : x.MaxChange = CDbl(it(2))
  ElseIf line = "#NOCALC" Then
    noCalc = True
  ElseIf Len(line) > 0 Then
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
If Not noCalc Then x.CalculateFull
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
  If t = "n" Then
    out.WriteLine lines(i)(0) & vbTab & lines(i)(1) & vbTab & t & vbTab & s2 & vbTab & c.Text & vbTab & Exact(CDbl(v))
  Else
    out.WriteLine lines(i)(0) & vbTab & lines(i)(1) & vbTab & t & vbTab & s2 & vbTab & c.Text
  End If
Next
For Each s In wb.Worksheets
  Set r = Nothing
  On Error Resume Next
  Set r = s.CircularReference
  On Error GoTo 0
  If Not r Is Nothing Then out.WriteLine "CIRCULAR" & vbTab & s.Name & vbTab & r.Address(False, False)
Next
out.Close
x.Iteration = False
wb.Close False
x.Quit

' v as hi:lo:e with v = (hi*2^26 + lo) * 2^e exactly (every step a multiplication by 2: exact).
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
