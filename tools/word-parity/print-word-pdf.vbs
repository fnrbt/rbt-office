Option Explicit

Const wdDoNotSaveChanges = 0
Const wdExportFormatPDF = 17
Const wdExportOptimizeForPrint = 0
Const wdExportAllDocument = 0
Const wdExportDocumentContent = 0
Const wdExportCreateHeadingBookmarks = 1

Dim fso, inputPath, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo print-word-pdf.vbs input.docx output.pdf"
  WScript.Quit 2
End If

inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

If Not fso.FileExists(inputPath) Then
  WScript.Echo "input not found: " & inputPath
  WScript.Quit 2
End If

EnsureParentFolder outputPath

Dim word, doc
Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
' (Not ScreenUpdating = False: with it, ExportAsFixedFormat drops paragraph
' borders (pBdr) from the PDF, though Word has read them. Checked 2026-11-14.)

On Error Resume Next
word.Options.Pagination = True
On Error GoTo 0

Set doc = word.Documents.Open(inputPath, False, True, False)
doc.Repaginate

doc.ExportAsFixedFormat _
  outputPath, _
  wdExportFormatPDF, _
  False, _
  wdExportOptimizeForPrint, _
  wdExportAllDocument, _
  1, _
  1, _
  wdExportDocumentContent, _
  True, _
  True, _
  wdExportCreateHeadingBookmarks, _
  True, _
  True, _
  False

doc.Close wdDoNotSaveChanges
word.Quit

WScript.Echo outputPath

Sub EnsureParentFolder(path)
  Dim parent
  parent = fso.GetParentFolderName(path)
  If parent <> "" And Not fso.FolderExists(parent) Then
    EnsureParentFolder parent
    fso.CreateFolder parent
  End If
End Sub
