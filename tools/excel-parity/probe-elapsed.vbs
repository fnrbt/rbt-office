' probe-elapsed.vbs
'
' Sweeps Excel's elapsed-time number-format output ([h], [m], [s] and
' their padded / fractional variants) across a grid of time serials by
' setting each (serial, format) pair into a fresh cell and reading back
' `cell.Text`. Elapsed formats show TOTAL hours/minutes/seconds rather
' than wrapping at a clock boundary, and .NET cannot express them with a
' stock DateTime format string, so Excel's own output is our ground truth.
'
' Usage:
'   cscript //nologo probe-elapsed.vbs <unused-input> <output.tsv>
'
' Args layout matches run-excel-vbs-on-vm.sh's contract (input ignored).

Option Explicit

Const xlCalculationManual = -4135

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo probe-elapsed.vbs <input> <output.tsv>"
  WScript.Quit 2
End If

outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, workbook, sheet, outFile
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
excel.EnableEvents = False
On Error Resume Next
excel.Calculation = xlCalculationManual
On Error GoTo 0

Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)

Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine "serial" & vbTab & "numFmt" & vbTab & "displayedText"

' Time serials (fraction of a day). Cover whole/partial hours, the >24h
' wraparound case, sub-minute, and sub-second fractions.
' Serial -> intended clock value:
'   0.5=12:00:00  0.0416..=1:00:00  1.0625=25:30:00 (>24h, elapsed)
'   2.5208..=60:30:00  0.000694..=0:01:00  0.0001157..=0:00:10
'   0.052083..=1:15:00  1.521261..=36:30:37  0.0003761..=0:00:32.5
'   0.043831..=1:03:07
Dim serials, i
serials = Array( _
  0.5, _
  0.041666666666667, _
  1.0625, _
  2.5208333333333, _
  0.0006944444444, _
  0.0001157407407, _
  0.052083333333, _
  1.5212615740741, _
  0.00037615740741, _
  0.043831018518519 )

Dim fmts, j
fmts = Array( _
  "[h]:mm:ss", _
  "[h]:mm", _
  "[hh]:mm:ss", _
  "[h]", _
  "[m]:ss", _
  "[m]", _
  "[mm]:ss", _
  "[s]", _
  "[ss]", _
  "[s].000", _
  "[h]:mm:ss.0", _
  "h:mm:ss" )

Dim cell, shown
Set cell = sheet.Range("A1")

For i = 0 To UBound(serials)
  For j = 0 To UBound(fmts)
    cell.NumberFormat = fmts(j)
    cell.Value = serials(i)
    shown = cell.Text
    outFile.WriteLine _
      CStr(serials(i)) & vbTab & _
      fmts(j) & vbTab & _
      shown
  Next
Next

outFile.Close
workbook.Close False
excel.Quit
Set excel = Nothing
