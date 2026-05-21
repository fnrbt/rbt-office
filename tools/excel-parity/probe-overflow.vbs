' probe-overflow.vbs
'
' Sweeps Excel's hash-overflow decision across a grid of (text, font,
' size, column width) cases by writing each into a fresh workbook,
' setting an exact column width, and reading back `cell.Text` and
' `cell.Width`. Excel returns "########" when the formatted number/date
' overflows the column, otherwise it returns the formatted display. We
' record both — that's our ground truth for the fit threshold.
'
' Usage:
'   cscript //nologo probe-overflow.vbs <unused-input.xlsx> <output.tsv>
'
' Args layout matches run-excel-vbs-on-vm.sh's contract (script needs
' to accept an input and an output path). The input is ignored.

Option Explicit

Const xlNormalView = 1
Const xlCalculationManual = -4135
Const xlAlertsNone = False
Const xlDoNotSaveChanges = 2

Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo probe-overflow.vbs <input> <output.tsv>"
  WScript.Quit 2
End If

outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, workbook, sheet, outFile
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = xlAlertsNone
excel.ScreenUpdating = False
excel.EnableEvents = False
excel.AskToUpdateLinks = False
On Error Resume Next
excel.Calculation = xlCalculationManual
On Error GoTo 0

Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)

Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine _
  "fontName" & vbTab & _
  "fontSize" & vbTab & _
  "storedCharWidth" & vbTab & _
  "cellWidthPt" & vbTab & _
  "valueSerial" & vbTab & _
  "numFmt" & vbTab & _
  "displayedText" & vbTab & _
  "hashed"

' Date serials. 36528 = 2000-01-03 → "1/3/2000 0:00" (13 chars).
' 36535 = 2000-01-10 → "1/10/2000 0:00" (14 chars). The two cases that
' bracket the narrow-date diff.
Dim serials, serial_count, i
serials = Array(36528, 36535, 36901, 37000)  ' 13ch, 14ch, 14ch, 14ch
serial_count = UBound(serials) + 1

Dim fonts, sizes, fIdx, sIdx, sName, sSize
fonts = Array("Calibri", "Arial")
sizes = Array(10, 11, 12, 14, 18)

' Char-width steps. Excel's stored width is in MDW char units;
' we sweep tightly around the transition zone. 11..16 chars covers
' the boundary for 13-14-char date strings under both fonts.
Dim charWidths, cwIdx, cw
charWidths = Array(11.0, 11.25, 11.5, 11.75, 12.0, 12.25, 12.5, 12.75, _
                   13.0, 13.25, 13.5, 13.75, 14.0, 14.25, 14.5, 14.75, _
                   15.0, 15.5, 16.0)

' Pick a single workbook style anchor — we change the cell's font and
' size for each probe to ensure each combo is exercised against Excel's
' actual measurement.
Dim numFmt
numFmt = "m/d/yyyy h:mm"

Dim cell
Set cell = sheet.Range("A1")
cell.NumberFormat = numFmt

For fIdx = 0 To UBound(fonts)
  sName = fonts(fIdx)
  cell.Font.Name = sName
  For sIdx = 0 To UBound(sizes)
    sSize = sizes(sIdx)
    cell.Font.Size = sSize
    For cwIdx = 0 To UBound(charWidths)
      cw = charWidths(cwIdx)
      sheet.Columns(1).ColumnWidth = cw
      For i = 0 To serial_count - 1
        cell.Value = serials(i)
        Dim shown, hashed, widthPt
        shown = cell.Text
        widthPt = cell.Width
        hashed = (InStr(shown, "#") > 0)
        outFile.WriteLine _
          sName & vbTab & _
          sSize & vbTab & _
          cw & vbTab & _
          widthPt & vbTab & _
          serials(i) & vbTab & _
          numFmt & vbTab & _
          shown & vbTab & _
          CStr(hashed)
      Next
    Next
  Next
Next

outFile.Close
workbook.Close False
excel.Quit
Set excel = Nothing
