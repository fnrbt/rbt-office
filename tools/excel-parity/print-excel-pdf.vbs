Option Explicit

Const xlTypePDF = 0
Const xlQualityStandard = 0
Const xlDoNotSaveChanges = 2
Const xlAlertsNone = False

Dim fso, inputPath, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo print-excel-pdf.vbs input.xlsx output.pdf"
  WScript.Quit 2
End If

inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

If Not fso.FileExists(inputPath) Then
  WScript.Echo "input not found: " & inputPath
  WScript.Quit 2
End If

Dim excel, workbook
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = xlAlertsNone
excel.ScreenUpdating = False
excel.EnableEvents = False
excel.AskToUpdateLinks = False

Set workbook = excel.Workbooks.Open(inputPath, False, True)

' Print Excel's row/column headers and gridlines so the PDF matches what a
' user sees on screen. Without these, the PDF is just bare cell content
' with no axis to align our renderer against.
Dim sheet
For Each sheet In workbook.Worksheets
  On Error Resume Next
  sheet.PageSetup.PrintHeadings = True
  sheet.PageSetup.PrintGridlines = True
  On Error GoTo 0
Next

workbook.ExportAsFixedFormat xlTypePDF, outputPath, xlQualityStandard, True, False, , , False

workbook.Close xlDoNotSaveChanges
excel.Quit
