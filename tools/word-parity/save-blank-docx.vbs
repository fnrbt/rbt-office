Option Explicit
' cscript //nologo save-blank-docx.vbs ignored-input output.docx
' Saves Word's new blank document (Normal.dotm's styles) as output.docx: the base for macro documents.

Const wdFormatXMLDocument = 12
Dim fso, outputPath, word, doc
Set fso = CreateObject("Scripting.FileSystemObject")
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
Set doc = word.Documents.Add
doc.SaveAs2 outputPath, wdFormatXMLDocument
doc.Close 0
word.Quit
WScript.Echo outputPath
