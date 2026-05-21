Option Explicit

Const wdActiveEndPageNumber = 3
Const wdActiveEndAdjustedPageNumber = 1
Const wdActiveEndSectionNumber = 2
Const wdNumberOfPagesInDocument = 4
Const wdHorizontalPositionRelativeToPage = 5
Const wdVerticalPositionRelativeToPage = 6
Const wdHorizontalPositionRelativeToTextBoundary = 7
Const wdVerticalPositionRelativeToTextBoundary = 8
Const wdFirstCharacterColumnNumber = 9
Const wdFirstCharacterLineNumber = 10
Const wdWithInTable = 12
Const wdStartOfRangeRowNumber = 13
Const wdEndOfRangeRowNumber = 14
Const wdMaximumNumberOfRows = 15
Const wdStartOfRangeColumnNumber = 16
Const wdEndOfRangeColumnNumber = 17
Const wdMaximumNumberOfColumns = 18
Const wdAtEndOfRowMarker = 31
Const wdInContentControl = 46
Const wdInFieldCode = 44
Const wdInFieldResult = 45
Const wdStatisticLines = 1
Const wdStatisticPages = 2
Const wdPrintView = 3
Const wdDoNotSaveChanges = 0
Const wdAlertsNone = 0

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
word.DisplayAlerts = wdAlertsNone

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
  "adjustedPage" & vbTab & _
  "section" & vbTab & _
  "pageCount" & vbTab & _
  "textBoundaryX" & vbTab & _
  "textBoundaryY" & vbTab & _
  "firstCharColumn" & vbTab & _
  "startRow" & vbTab & _
  "endRow" & vbTab & _
  "maxRows" & vbTab & _
  "startColumn" & vbTab & _
  "endColumn" & vbTab & _
  "maxColumns" & vbTab & _
  "atEndOfRowMarker" & vbTab & _
  "inContentControl" & vbTab & _
  "inFieldCode" & vbTab & _
  "inFieldResult" & vbTab & _
  "rangeLines" & vbTab & _
  "rangePages" & vbTab & _
  "inTable" & vbTab & _
  "table" & vbTab & _
  "row" & vbTab & _
  "column" & vbTab & _
  "style" & vbTab & _
  "text"

Dim paraIndex, paraTotal, startSeconds
paraIndex = 0
paraTotal = doc.Paragraphs.Count
startSeconds = Timer
WScript.StdErr.WriteLine "progress: " & paraTotal & " paragraphs; t=0.0s"

Dim para
For Each para In doc.Paragraphs
  paraIndex = paraIndex + 1
  EmitParagraphLines doc, inputPath, outFile, para, paraIndex
  ' Progress to stderr (captured in the task log) so a stall is visible at the
  ' paragraph it stops on instead of just timing out opaquely.
  If paraIndex Mod 10 = 0 Or paraIndex = paraTotal Then
    WScript.StdErr.WriteLine "progress: paragraph " & paraIndex & "/" & paraTotal & " t=" & FormatNumber(Timer - startSeconds, 1) & "s"
  End If
Next

WScript.StdErr.WriteLine "progress: done t=" & FormatNumber(Timer - startSeconds, 1) & "s"

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

  Dim offset
  offset = NextVisibleOffset(docObj, baseEnd, baseStart)

  Do While offset < baseEnd
    Dim r, pageNum, lineNum, x, y, lineEnd, textEnd, lineText, lastRange, lastX, lastY
    Set r = docObj.Range(offset, offset + 1)
    pageNum = SafeInfo(r, wdActiveEndPageNumber)
    lineNum = SafeInfo(r, wdFirstCharacterLineNumber)
    x = SafeInfo(r, wdHorizontalPositionRelativeToPage)
    y = SafeInfo(r, wdVerticalPositionRelativeToPage)

    If pageNum = "" Or lineNum = "" Then
      offset = NextVisibleOffset(docObj, baseEnd, offset + 1)
    Else
      lineEnd = FindLineEnd(docObj, baseEnd, offset, pageNum, lineNum)
      textEnd = TrimTrailingSkipped(docObj, offset, lineEnd)

      If textEnd > offset Then
        Set lastRange = docObj.Range(textEnd - 1, textEnd)
        lastX = SafeInfo(lastRange, wdHorizontalPositionRelativeToPage)
        lastY = SafeInfo(lastRange, wdVerticalPositionRelativeToPage)
        lineText = docObj.Range(offset, textEnd).Text

        WriteLine fileObj, docObj, docPath, paragraph, paragraphIndex, pageNum, lineNum, offset - baseStart, textEnd - baseStart, x, y, lastX, lastY, lineText
      End If

      offset = NextVisibleOffset(docObj, baseEnd, lineEnd)
    End If
  Loop
End Sub

Function FindLineEnd(docObj, baseEnd, startOffset, pageNum, lineNum)
  Dim lo, hi, stepSize
  lo = startOffset
  hi = startOffset + 1
  stepSize = 1

  Do While hi < baseEnd And OffsetOnLine(docObj, hi, pageNum, lineNum)
    lo = hi
    stepSize = stepSize * 2
    hi = startOffset + stepSize
    If hi > baseEnd Then
      hi = baseEnd
    End If
  Loop

  If hi >= baseEnd Then
    If OffsetOnLine(docObj, baseEnd - 1, pageNum, lineNum) Then
      FindLineEnd = baseEnd
      Exit Function
    End If
  End If

  Dim low, high, mid
  low = lo + 1
  high = hi

  Do While low < high
    mid = (low + high) \ 2
    If OffsetOnLine(docObj, mid, pageNum, lineNum) Then
      low = mid + 1
    Else
      high = mid
    End If
  Loop

  FindLineEnd = low
End Function

Function OffsetOnLine(docObj, offset, pageNum, lineNum)
  Dim r, ch
  Set r = docObj.Range(offset, offset + 1)
  ch = r.Text

  If IsSkippedLayoutChar(ch) Then
    OffsetOnLine = False
    Exit Function
  End If

  OffsetOnLine = (SafeInfo(r, wdActiveEndPageNumber) = pageNum And SafeInfo(r, wdFirstCharacterLineNumber) = lineNum)
End Function

Function NextVisibleOffset(docObj, baseEnd, offset)
  Dim pos, ch
  pos = offset

  Do While pos < baseEnd
    ch = docObj.Range(pos, pos + 1).Text
    If Not IsSkippedLayoutChar(ch) Then
      Exit Do
    End If
    pos = pos + 1
  Loop

  NextVisibleOffset = pos
End Function

Function TrimTrailingSkipped(docObj, startOffset, endOffset)
  Dim pos, ch
  pos = endOffset

  Do While pos > startOffset
    ch = docObj.Range(pos - 1, pos).Text
    If Not IsSkippedLayoutChar(ch) Then
      Exit Do
    End If
    pos = pos - 1
  Loop

  TrimTrailingSkipped = pos
End Function

Function IsSkippedLayoutChar(ch)
  IsSkippedLayoutChar = (ch = Chr(13) Or ch = Chr(7))
End Function

Function SafeInfo(rangeObj, infoKind)
  On Error Resume Next
  SafeInfo = rangeObj.Information(infoKind)
  If Err.Number <> 0 Then
    SafeInfo = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeStatistic(rangeObj, statisticKind)
  On Error Resume Next
  SafeStatistic = rangeObj.ComputeStatistics(statisticKind)
  If Err.Number <> 0 Then
    SafeStatistic = ""
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
  Dim lineRange
  Set lineRange = docObj.Range(paragraph.Range.Start + startOffset, paragraph.Range.Start + endOffset)

  Dim adjustedPage, sectionNumber, pageCount, textBoundaryX, textBoundaryY, firstCharColumn
  adjustedPage = SafeInfo(lineRange, wdActiveEndAdjustedPageNumber)
  sectionNumber = SafeInfo(lineRange, wdActiveEndSectionNumber)
  pageCount = SafeInfo(lineRange, wdNumberOfPagesInDocument)
  textBoundaryX = SafeInfo(lineRange, wdHorizontalPositionRelativeToTextBoundary)
  textBoundaryY = SafeInfo(lineRange, wdVerticalPositionRelativeToTextBoundary)
  firstCharColumn = SafeInfo(lineRange, wdFirstCharacterColumnNumber)

  Dim startRow, endRow, maxRows, startColumn, endColumn, maxColumns
  startRow = SafeInfo(lineRange, wdStartOfRangeRowNumber)
  endRow = SafeInfo(lineRange, wdEndOfRangeRowNumber)
  maxRows = SafeInfo(lineRange, wdMaximumNumberOfRows)
  startColumn = SafeInfo(lineRange, wdStartOfRangeColumnNumber)
  endColumn = SafeInfo(lineRange, wdEndOfRangeColumnNumber)
  maxColumns = SafeInfo(lineRange, wdMaximumNumberOfColumns)

  Dim atEndOfRowMarker, inContentControl, inFieldCode, inFieldResult, rangeLines, rangePages
  atEndOfRowMarker = SafeInfo(lineRange, wdAtEndOfRowMarker)
  inContentControl = SafeInfo(lineRange, wdInContentControl)
  inFieldCode = SafeInfo(lineRange, wdInFieldCode)
  inFieldResult = SafeInfo(lineRange, wdInFieldResult)
  rangeLines = SafeStatistic(lineRange, wdStatisticLines)
  rangePages = SafeStatistic(lineRange, wdStatisticPages)

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
    Tsv(CStr(adjustedPage)) & vbTab & _
    Tsv(CStr(sectionNumber)) & vbTab & _
    Tsv(CStr(pageCount)) & vbTab & _
    Tsv(CStr(textBoundaryX)) & vbTab & _
    Tsv(CStr(textBoundaryY)) & vbTab & _
    Tsv(CStr(firstCharColumn)) & vbTab & _
    Tsv(CStr(startRow)) & vbTab & _
    Tsv(CStr(endRow)) & vbTab & _
    Tsv(CStr(maxRows)) & vbTab & _
    Tsv(CStr(startColumn)) & vbTab & _
    Tsv(CStr(endColumn)) & vbTab & _
    Tsv(CStr(maxColumns)) & vbTab & _
    Tsv(CStr(atEndOfRowMarker)) & vbTab & _
    Tsv(CStr(inContentControl)) & vbTab & _
    Tsv(CStr(inFieldCode)) & vbTab & _
    Tsv(CStr(inFieldResult)) & vbTab & _
    Tsv(CStr(rangeLines)) & vbTab & _
    Tsv(CStr(rangePages)) & vbTab & _
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
