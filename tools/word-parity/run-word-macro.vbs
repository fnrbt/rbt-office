Option Explicit
' cscript //nologo run-word-macro.vbs input.docm <Macro>.tsv
'
' Opens a macro-enabled document in Word with macros allowed (AutomationSecurity = 1), runs the macro
' the output file is named after (Main.tsv runs "Main", Module1.Main.tsv runs "Module1.Main"), and
' writes what the document then is: <Macro>.tsv (paragraphs with their text and formatting, tables,
' counts, the document variable "log" a probe macro writes) and <Macro>.pdf (the pages).
' A run-time error of the macro is written as the first line ("error<TAB>number<TAB>description").

Const wdDoNotSaveChanges = 0
Const wdExportFormatPDF = 17

Dim fso, inputPath, outputPath, pdfPath, macroName, word, doc, out
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
macroName = fso.GetBaseName(outputPath)
pdfPath = fso.BuildPath(fso.GetParentFolderName(outputPath), macroName & ".pdf")

Set word = CreateObject("Word.Application")
word.Visible = True
word.DisplayAlerts = 0
word.AutomationSecurity = 1
Set doc = word.Documents.Open(inputPath, False, False, False)
doc.Activate

Set out = fso.CreateTextFile(outputPath, True, True)
On Error Resume Next
word.Run macroName
If Err.Number <> 0 Then
  out.WriteLine "error" & vbTab & Err.Number & vbTab & Esc(Err.Description)
  Err.Clear
End If
On Error GoTo 0

Dim d
Set d = word.ActiveDocument
Dump d
d.Repaginate
d.ExportAsFixedFormat pdfPath, wdExportFormatPDF
out.Close
d.Close wdDoNotSaveChanges
On Error Resume Next
doc.Close wdDoNotSaveChanges
On Error GoTo 0
word.Quit
WScript.Echo outputPath

Function Esc(s)
  Dim t
  t = Replace(CStr(s), "\", "\\")
  t = Replace(t, vbTab, "\t")
  t = Replace(t, vbCr, "\r")
  t = Replace(t, vbLf, "\n")
  t = Replace(t, Chr(7), "\a")
  t = Replace(t, Chr(11), "\v")
  t = Replace(t, Chr(12), "\f")
  Esc = t
End Function

Sub Dump(d)
  Dim p, i, r, f, t, ti, ri, ci, v
  out.WriteLine "counts" & vbTab & d.Paragraphs.Count & vbTab & d.Words.Count & vbTab & d.Characters.Count & vbTab & d.Tables.Count & vbTab & d.Content.End & vbTab & d.ComputeStatistics(0) & vbTab & d.ComputeStatistics(2)
  out.WriteLine "selection" & vbTab & word.Selection.Start & vbTab & word.Selection.End
  On Error Resume Next
  v = d.Variables("log").Value
  If Err.Number = 0 Then out.WriteLine "log" & vbTab & Esc(v)
  Err.Clear
  On Error GoTo 0
  i = 0
  For Each p In d.Paragraphs
    i = i + 1
    Set r = p.Range
    Set f = r.Characters(1).Font
    out.WriteLine "p" & vbTab & i & vbTab & r.Start & vbTab & r.End & vbTab & Esc(r.Text) & vbTab & p.Style & vbTab & p.Alignment & vbTab & _
      p.SpaceBefore & vbTab & p.SpaceAfter & vbTab & p.LeftIndent & vbTab & r.Information(12) & vbTab & _
      f.Name & vbTab & f.Size & vbTab & f.Bold & vbTab & f.Italic & vbTab & f.Underline & vbTab & f.Color & vbTab & r.Characters(1).HighlightColorIndex & vbTab & _
      r.Font.Bold & vbTab & r.Font.Size & vbTab & r.Font.Color & vbTab & r.ListFormat.ListType & vbTab & p.Shading.BackgroundPatternColor
  Next
  ti = 0
  For Each t In d.Tables
    ti = ti + 1
    out.WriteLine "table" & vbTab & ti & vbTab & t.Rows.Count & vbTab & t.Columns.Count & vbTab & t.Range.Start & vbTab & t.Range.End & vbTab & t.Borders.Enable
    For ri = 1 To t.Rows.Count
      For ci = 1 To t.Rows(ri).Cells.Count
        Set r = t.Cell(ri, ci).Range
        out.WriteLine "cell" & vbTab & ti & vbTab & ri & vbTab & ci & vbTab & Esc(r.Text) & vbTab & t.Cell(ri, ci).Shading.BackgroundPatternColor & vbTab & r.Characters(1).Font.Bold & vbTab & r.Characters(1).Font.Color
      Next
    Next
  Next
End Sub
