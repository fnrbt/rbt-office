' error-probe.vbs WORKBOOK OUT.tsv [RANGE]: the background error checks Excel 2019 finds in every sheet's cells
' (RANGE of each sheet, default A1:AD80): one line per cell and check that holds (Range.Errors(n).Value), with
' the check's number - 1 evaluates to an error, 2 text date with a two-digit year, 3 number stored as text,
' 4 inconsistent formula, 5 formula omits adjacent cells, 6 unlocked formula, 7 empty cell reference, 8 list data
' validation, 9 inconsistent table formula - and first the checks Excel is set to run (ErrorCheckingOptions).
' Run it in the desktop session (run-vbs-in-session.ps1).
Option Explicit
Dim book, out, area, fso, f, xl, wb, ws, c, n, used, o
book = WScript.Arguments(0)
out = WScript.Arguments(1)
area = "A1:AD80"
If WScript.Arguments.Count > 2 Then area = WScript.Arguments(2)
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.CreateTextFile(out, True, True)
Set xl = CreateObject("Excel.Application")
xl.Visible = True
xl.DisplayAlerts = False
xl.AskToUpdateLinks = False
xl.EnableEvents = False
Set o = xl.ErrorCheckingOptions
f.WriteLine "options" & vbTab & "background " & xl.ErrorCheckingOptions.BackgroundChecking & " evaluate " & o.EvaluateToError & " textDate " & o.TextDate & " numberAsText " & o.NumberAsText & " inconsistent " & o.InconsistentFormula & " omitted " & o.OmittedCells & " unlocked " & o.UnlockedFormulaCells & " empty " & o.EmptyCellReferences & " list " & o.ListDataValidation & " table " & o.InconsistentTableFormula
Set wb = xl.Workbooks.Open(book, False, True)
On Error Resume Next
For Each ws In wb.Worksheets
  Set used = Nothing
  Set used = xl.Intersect(ws.UsedRange, ws.Range(area))
  If Not used Is Nothing Then
    For Each c In used.Cells
      For n = 1 To 9
        Err.Clear
        If c.Errors(n).Value And Not c.Errors(n).Ignore Then
          If Err.Number = 0 Then f.WriteLine ws.Name & vbTab & c.Address(False, False) & vbTab & n
        End If
      Next
    Next
  End If
Next
On Error GoTo 0
f.Close
wb.Close False
xl.Quit
