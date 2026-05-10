Option Explicit

Const wdActiveEndPageNumber = 3
Const wdHorizontalPositionRelativeToPage = 5
Const wdVerticalPositionRelativeToPage = 6
Const wdFirstCharacterLineNumber = 10
Const wdPrintView = 3
Const wdDoNotSaveChanges = 0

Dim fso, inputPath, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo extract-word-layout.vbs input.docx output.tsv"
  WScript.Quit 2
End If

inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

If Not fso.FileExists(inputPath) Then
  WScript.Echo "input not found: " & inputPath
  WScript.Quit 2
End If

Dim word, doc, outFile
Set word = CreateObject("Word.Application")
word.Visible = False
word.ScreenUpdating = False

On Error Resume Next
word.Options.Pagination = True
On Error GoTo 0

Set doc = word.Documents.Open(inputPath, False, True, False)
word.ActiveWindow.View.Type = wdPrintView
doc.Repaginate

Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine "doc" & vbTab & "page" & vbTab & "paragraph" & vbTab & "line" & vbTab & "start" & vbTab & "end" & vbTab & "x" & vbTab & "y" & vbTab & "text"

Dim paraIndex
paraIndex = 0

Dim para
For Each para In doc.Paragraphs
  paraIndex = paraIndex + 1
  EmitParagraphLines doc, inputPath, outFile, para, paraIndex
Next

outFile.Close
doc.Close wdDoNotSaveChanges
word.Quit

Sub EmitParagraphLines(docObj, docPath, fileObj, paragraph, paragraphIndex)
  Dim baseStart, baseEnd, charCount
  baseStart = paragraph.Range.Start
  baseEnd = paragraph.Range.End
  charCount = baseEnd - baseStart

  If charCount <= 0 Then
    Exit Sub
  End If

  Dim currentPage, currentLine, currentX, currentY, currentStart, currentEnd, currentText
  currentPage = Empty
  currentLine = Empty
  currentX = Empty
  currentY = Empty
  currentStart = 0
  currentEnd = 0
  currentText = ""

  Dim i, r, pageNum, lineNum, x, y, ch
  For i = 0 To charCount - 1
    Set r = docObj.Range(baseStart + i, baseStart + i + 1)
    ch = r.Text

    If ch <> Chr(13) And ch <> Chr(7) Then
      pageNum = SafeInfo(r, wdActiveEndPageNumber)
      lineNum = SafeInfo(r, wdFirstCharacterLineNumber)
      x = SafeInfo(r, wdHorizontalPositionRelativeToPage)
      y = SafeInfo(r, wdVerticalPositionRelativeToPage)

      If IsEmpty(currentLine) Then
        currentPage = pageNum
        currentLine = lineNum
        currentX = x
        currentY = y
        currentStart = i
        currentEnd = i + 1
        currentText = ch
      ElseIf pageNum = currentPage And lineNum = currentLine Then
        currentEnd = i + 1
        currentText = currentText & ch
      Else
        WriteLine fileObj, docPath, currentPage, paragraphIndex, currentLine, currentStart, currentEnd, currentX, currentY, currentText
        currentPage = pageNum
        currentLine = lineNum
        currentX = x
        currentY = y
        currentStart = i
        currentEnd = i + 1
        currentText = ch
      End If
    End If
  Next

  If Not IsEmpty(currentLine) Then
    WriteLine fileObj, docPath, currentPage, paragraphIndex, currentLine, currentStart, currentEnd, currentX, currentY, currentText
  End If
End Sub

Function SafeInfo(rangeObj, infoKind)
  On Error Resume Next
  SafeInfo = rangeObj.Information(infoKind)
  If Err.Number <> 0 Then
    SafeInfo = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Sub WriteLine(fileObj, docPath, pageNum, paragraphIndex, lineNum, startOffset, endOffset, x, y, lineText)
  fileObj.WriteLine _
    Tsv(docPath) & vbTab & _
    Tsv(CStr(pageNum)) & vbTab & _
    Tsv(CStr(paragraphIndex)) & vbTab & _
    Tsv(CStr(lineNum)) & vbTab & _
    Tsv(CStr(startOffset)) & vbTab & _
    Tsv(CStr(endOffset)) & vbTab & _
    Tsv(CStr(x)) & vbTab & _
    Tsv(CStr(y)) & vbTab & _
    Tsv(lineText)
End Sub

Function Tsv(value)
  Dim text
  text = CStr(value)
  text = Replace(text, vbTab, " ")
  text = Replace(text, vbCr, " ")
  text = Replace(text, vbLf, " ")
  Tsv = text
End Function
