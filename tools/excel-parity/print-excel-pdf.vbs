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

' Print exactly what the fixture asks for — its <printOptions
' gridLines/headings> drive the chrome. Earlier versions force-enabled both
' as a workaround for our renderer not supporting chrome; now that the
' chrome path lands the fixture's preferences, forcing them on hides real
' parity bugs (e.g. a fixture with headings="0" should print no chrome).

workbook.ExportAsFixedFormat xlTypePDF, outputPath, xlQualityStandard, True, False, , , False

workbook.Close xlDoNotSaveChanges
excel.Quit
