' Run in Excel's desktop session: probe-cache-open.vbs BOOK.xlsx OUT.tsv
' Record typed values after opening, ordinary/full calculation, and breaking each cycle.
' Reproduce:
'   python3 tools/excel-parity/make-cache-open-probe.py /tmp/cache-open-probe
'   BOOKFIRST=1 tools/excel-parity/probe-on-vm.sh tools/excel-parity/probe-cache-open.vbs /tmp/noiter.tsv /tmp/cache-open-probe/cache-open-noiter.xlsx
'   BOOKFIRST=1 tools/excel-parity/probe-on-vm.sh tools/excel-parity/probe-cache-open.vbs /tmp/iter.tsv /tmp/cache-open-probe/cache-open-iter.xlsx
' Excel 2019, auto calculation (-4105), iteration count 2, max change 0.001:
' iteration off: on open, Calculate, and CalculateFull, A is String "" for absent
' or empty v (also t=str), Double 0 for zero v, Double 7 for nonzero v.
' B and dependent C retain cached Double 2 and 99; acyclic D is Double 3.
' iteration on: all cache variants match: A/B/C are Doubles 2.5/1.25/12.5
' on open, 2.125/1.0625/12.125 after Calculate, and
' 2.03125/1.015625/12.03125 after CalculateFull; D is Double 3.
' With either setting, replacing each B with 4 breaks its cycle and produces
' Double A=5, B=4, C=15, D=3.
' For an empty numeric cache on every dependent C, add uncached-dependent as
' the generator's second argument. Iteration off: C is String "" on open,
' Calculate, CalculateFull, and reopen; A's empty strings also survive reopen.
' Excel saves A1, A4, C1, C4 as t="str" with <v/>. Raw saved workbook can
' be retrieved from the VM copy after the probe (which saves before reopen).
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application")
x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
x.AskToUpdateLinks = False
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, False)
Set s = wb.Worksheets(1)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
Dump "open"
x.Calculate
Dump "calculate"
x.CalculateFull
Dump "full"
wb.Save
wb.Close False
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, False)
Set s = wb.Worksheets(1)
Dump "reopen"
For row = 1 To 6
  s.Range("B" & row).Value2 = 4
Next
Dump "break"
out.Close
wb.Close False
x.Quit

Sub Dump(stage)
  out.WriteLine "SETTINGS" & vbTab & stage & vbTab & x.Calculation & vbTab & x.Iteration & vbTab & x.MaxIterations & vbTab & x.MaxChange
  For r = 1 To 6
    For col = 1 To 4
      Set cell = s.Cells(r, col)
      v = cell.Value2
      If VarType(v) = 10 Then
        value = "ERROR:" & cell.Text
      ElseIf IsEmpty(v) Then
        value = "EMPTY"
      ElseIf VarType(v) = vbString Then
        value = "STRING:[" & v & "]"
      Else
        value = TypeName(v) & ":[" & CStr(v) & "]"
      End If
      out.WriteLine stage & vbTab & s.Cells(r, 5).Value2 & vbTab & cell.Address(False, False) & vbTab & value
    Next
  Next
End Sub
