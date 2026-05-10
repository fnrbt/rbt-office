Option Explicit

Const wdActiveEndPageNumber = 3
Const wdHorizontalPositionRelativeToPage = 5
Const wdVerticalPositionRelativeToPage = 6
Const wdFirstCharacterLineNumber = 10
Const wdWithInTable = 12
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
outFile.WriteLine _
  "doc" & vbTab & _
  "page" & vbTab & _
  "paragraph" & vbTab & _
  "line" & vbTab & _
  "start" & vbTab & _
  "end" & vbTab & _
  "x" & vbTab & _
  "y" & vbTab & _
  "lastX" & vbTab & _
  "lastY" & vbTab & _
  "inTable" & vbTab & _
  "table" & vbTab & _
  "row" & vbTab & _
  "column" & vbTab & _
  "style" & vbTab & _
  "text"

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

  Dim currentPage, currentLine, currentX, currentY, currentLastX, currentLastY, currentStart, currentEnd, currentText
  currentPage = Empty
  currentLine = Empty
  currentX = Empty
  currentY = Empty
  currentLastX = Empty
  currentLastY = Empty
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
        currentLastX = x
        currentLastY = y
        currentStart = i
        currentEnd = i + 1
        currentText = ch
      ElseIf pageNum = currentPage And lineNum = currentLine Then
        currentEnd = i + 1
        currentLastX = x
        currentLastY = y
        currentText = currentText & ch
      Else
        WriteLine fileObj, docObj, docPath, paragraph, paragraphIndex, currentPage, currentLine, currentStart, currentEnd, currentX, currentY, currentLastX, currentLastY, currentText
        currentPage = pageNum
        currentLine = lineNum
        currentX = x
        currentY = y
        currentLastX = x
        currentLastY = y
        currentStart = i
        currentEnd = i + 1
        currentText = ch
      End If
    End If
  Next

  If Not IsEmpty(currentLine) Then
    WriteLine fileObj, docObj, docPath, paragraph, paragraphIndex, currentPage, currentLine, currentStart, currentEnd, currentX, currentY, currentLastX, currentLastY, currentText
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

Function SafeStyleName(paragraph)
  On Error Resume Next
  SafeStyleName = paragraph.Range.Style.NameLocal
  If Err.Number <> 0 Then
    SafeStyleName = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function TableIndexForParagraph(docObj, paragraph)
  On Error Resume Next
  TableIndexForParagraph = ""
  If paragraph.Range.Tables.Count > 0 Then
    Dim paraStart, i
    paraStart = paragraph.Range.Start
    For i = 1 To docObj.Tables.Count
      If docObj.Tables(i).Range.Start <= paraStart And docObj.Tables(i).Range.End >= paraStart Then
        TableIndexForParagraph = CStr(i)
        Exit For
      End If
    Next
  End If
  If Err.Number <> 0 Then
    TableIndexForParagraph = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function TableRowForParagraph(paragraph)
  On Error Resume Next
  TableRowForParagraph = ""
  If paragraph.Range.Cells.Count > 0 Then
    TableRowForParagraph = CStr(paragraph.Range.Cells(1).RowIndex)
  End If
  If Err.Number <> 0 Then
    TableRowForParagraph = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function TableColumnForParagraph(paragraph)
  On Error Resume Next
  TableColumnForParagraph = ""
  If paragraph.Range.Cells.Count > 0 Then
    TableColumnForParagraph = CStr(paragraph.Range.Cells(1).ColumnIndex)
  End If
  If Err.Number <> 0 Then
    TableColumnForParagraph = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Sub WriteLine(fileObj, docObj, docPath, paragraph, paragraphIndex, pageNum, lineNum, startOffset, endOffset, x, y, lastX, lastY, lineText)
  Dim inTable, tableIndex, rowIndex, columnIndex, styleName
  inTable = SafeInfo(paragraph.Range, wdWithInTable)
  tableIndex = TableIndexForParagraph(docObj, paragraph)
  rowIndex = TableRowForParagraph(paragraph)
  columnIndex = TableColumnForParagraph(paragraph)
  styleName = SafeStyleName(paragraph)

  fileObj.WriteLine _
    Tsv(docPath) & vbTab & _
    Tsv(CStr(pageNum)) & vbTab & _
    Tsv(CStr(paragraphIndex)) & vbTab & _
    Tsv(CStr(lineNum)) & vbTab & _
    Tsv(CStr(startOffset)) & vbTab & _
    Tsv(CStr(endOffset)) & vbTab & _
    Tsv(CStr(x)) & vbTab & _
    Tsv(CStr(y)) & vbTab & _
    Tsv(CStr(lastX)) & vbTab & _
    Tsv(CStr(lastY)) & vbTab & _
    Tsv(CStr(inTable)) & vbTab & _
    Tsv(tableIndex) & vbTab & _
    Tsv(rowIndex) & vbTab & _
    Tsv(columnIndex) & vbTab & _
    Tsv(styleName) & vbTab & _
    Tsv(lineText)
End Sub

Function Tsv(value)
  Dim text
  text = CStr(value)
  text = Replace(text, vbTab, " ")
  text = Replace(text, vbCr, " ")
  text = Replace(text, vbLf, " ")
  text = Replace(text, Chr(7), "")
  Tsv = text
End Function
