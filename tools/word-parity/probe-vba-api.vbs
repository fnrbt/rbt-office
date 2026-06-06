Option Explicit

Dim inputPath, outputPath
inputPath = WScript.Arguments(0)
outputPath = WScript.Arguments(1)

Dim fso, out
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(outputPath, True, True)

Sub Emit(name, value)
  out.WriteLine name & vbTab & CStr(value)
End Sub

Function CleanText(value)
  CleanText = Replace(Replace(CStr(value), vbCr, "<CR>"), Chr(7), "<BEL>")
End Function

Dim word, doc, range
Set word = CreateObject("Word.Application")
word.Visible = False

On Error GoTo 0
Set doc = word.Documents.Add()
doc.Range.Text = "Hello world" & vbCr & "Second line"
Set range = doc.Range(0, 5)

Emit "WordVersion", word.Version
Emit "RangeText", CleanText(range.Text)
Emit "Start", range.Start
Emit "End", range.End
Emit "Words", range.ComputeStatistics(0)
Emit "Characters", range.ComputeStatistics(3)
Emit "CharactersWithSpaces", range.ComputeStatistics(5)
Emit "ActiveEndPage", range.Information(3)

doc.Close False
word.Quit
out.Close
