Option Explicit

Const wdPrintView = 3
Const wdDoNotSaveChanges = 0
Const wdAlertsNone = 0
Const wdHorizontalPositionRelativeToPage = 5
Const wdVerticalPositionRelativeToPage = 6

Const wdBorderTop = -1
Const wdBorderLeft = -2
Const wdBorderBottom = -3
Const wdBorderRight = -4

Dim fso, inputPath, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo extract-word-table-geometry.vbs input.docx output.tsv"
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
  "kind" & vbTab & _
  "doc" & vbTab & _
  "table" & vbTab & _
  "seq" & vbTab & _
  "row" & vbTab & _
  "column" & vbTab & _
  "start" & vbTab & _
  "end" & vbTab & _
  "x" & vbTab & _
  "y" & vbTab & _
  "width" & vbTab & _
  "height" & vbTab & _
  "topPadding" & vbTab & _
  "rightPadding" & vbTab & _
  "bottomPadding" & vbTab & _
  "leftPadding" & vbTab & _
  "topStyle" & vbTab & _
  "topWidth" & vbTab & _
  "topColor" & vbTab & _
  "rightStyle" & vbTab & _
  "rightWidth" & vbTab & _
  "rightColor" & vbTab & _
  "bottomStyle" & vbTab & _
  "bottomWidth" & vbTab & _
  "bottomColor" & vbTab & _
  "leftStyle" & vbTab & _
  "leftWidth" & vbTab & _
  "leftColor"

Dim tableIndex, tableObj
tableIndex = 0

For Each tableObj In doc.Tables
  tableIndex = tableIndex + 1
  EmitTable outFile, inputPath, tableObj, tableIndex
Next

outFile.Close
doc.Close wdDoNotSaveChanges
word.Quit

Sub EmitTable(fileObj, docPath, tableObj, tableIndex)
  fileObj.WriteLine _
    Tsv("table") & vbTab & _
    Tsv(docPath) & vbTab & _
    Tsv(CStr(tableIndex)) & vbTab & _
    Tsv("") & vbTab & _
    Tsv("") & vbTab & _
    Tsv("") & vbTab & _
    Tsv(CStr(SafeRangeStart(tableObj.Range))) & vbTab & _
    Tsv(CStr(SafeRangeEnd(tableObj.Range))) & vbTab & _
    Tsv(CStr(SafeInfo(tableObj.Range, wdHorizontalPositionRelativeToPage))) & vbTab & _
    Tsv(CStr(SafeInfo(tableObj.Range, wdVerticalPositionRelativeToPage))) & vbTab & _
    Tsv(CStr(SafeTableWidth(tableObj))) & vbTab & _
    Tsv(CStr(SafeTableHeight(tableObj))) & vbTab & _
    Tsv("") & vbTab & Tsv("") & vbTab & Tsv("") & vbTab & Tsv("") & vbTab & _
    Tsv("") & vbTab & Tsv("") & vbTab & Tsv("") & vbTab & _
    Tsv("") & vbTab & Tsv("") & vbTab & Tsv("") & vbTab & _
    Tsv("") & vbTab & Tsv("") & vbTab & Tsv("") & vbTab & _
    Tsv("") & vbTab & Tsv("") & vbTab & Tsv("")

  Dim seq, cellObj
  seq = 0
  For Each cellObj In tableObj.Range.Cells
    seq = seq + 1
    EmitCell fileObj, docPath, tableIndex, seq, cellObj
  Next
End Sub

Sub EmitCell(fileObj, docPath, tableIndex, seq, cellObj)
  fileObj.WriteLine _
    Tsv("cell") & vbTab & _
    Tsv(docPath) & vbTab & _
    Tsv(CStr(tableIndex)) & vbTab & _
    Tsv(CStr(seq)) & vbTab & _
    Tsv(CStr(SafeCellRow(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellColumn(cellObj))) & vbTab & _
    Tsv(CStr(SafeRangeStart(cellObj.Range))) & vbTab & _
    Tsv(CStr(SafeRangeEnd(cellObj.Range))) & vbTab & _
    Tsv(CStr(SafeInfo(cellObj.Range, wdHorizontalPositionRelativeToPage))) & vbTab & _
    Tsv(CStr(SafeInfo(cellObj.Range, wdVerticalPositionRelativeToPage))) & vbTab & _
    Tsv(CStr(SafeCellWidth(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellHeight(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellTopPadding(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellRightPadding(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellBottomPadding(cellObj))) & vbTab & _
    Tsv(CStr(SafeCellLeftPadding(cellObj))) & vbTab & _
    Tsv(CStr(SafeBorderStyle(cellObj, wdBorderTop))) & vbTab & _
    Tsv(CStr(SafeBorderWidth(cellObj, wdBorderTop))) & vbTab & _
    Tsv(CStr(SafeBorderColor(cellObj, wdBorderTop))) & vbTab & _
    Tsv(CStr(SafeBorderStyle(cellObj, wdBorderRight))) & vbTab & _
    Tsv(CStr(SafeBorderWidth(cellObj, wdBorderRight))) & vbTab & _
    Tsv(CStr(SafeBorderColor(cellObj, wdBorderRight))) & vbTab & _
    Tsv(CStr(SafeBorderStyle(cellObj, wdBorderBottom))) & vbTab & _
    Tsv(CStr(SafeBorderWidth(cellObj, wdBorderBottom))) & vbTab & _
    Tsv(CStr(SafeBorderColor(cellObj, wdBorderBottom))) & vbTab & _
    Tsv(CStr(SafeBorderStyle(cellObj, wdBorderLeft))) & vbTab & _
    Tsv(CStr(SafeBorderWidth(cellObj, wdBorderLeft))) & vbTab & _
    Tsv(CStr(SafeBorderColor(cellObj, wdBorderLeft)))
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

Function SafeRangeStart(rangeObj)
  On Error Resume Next
  SafeRangeStart = rangeObj.Start
  If Err.Number <> 0 Then SafeRangeStart = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeRangeEnd(rangeObj)
  On Error Resume Next
  SafeRangeEnd = rangeObj.End
  If Err.Number <> 0 Then SafeRangeEnd = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeTableWidth(tableObj)
  On Error Resume Next
  SafeTableWidth = tableObj.Range.Cells(1).Width * tableObj.Columns.Count
  If Err.Number <> 0 Then SafeTableWidth = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeTableHeight(tableObj)
  On Error Resume Next
  Dim total, rowObj
  total = 0
  For Each rowObj In tableObj.Rows
    total = total + rowObj.Height
  Next
  SafeTableHeight = total
  If Err.Number <> 0 Then SafeTableHeight = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellRow(cellObj)
  On Error Resume Next
  SafeCellRow = cellObj.RowIndex
  If Err.Number <> 0 Then SafeCellRow = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellColumn(cellObj)
  On Error Resume Next
  SafeCellColumn = cellObj.ColumnIndex
  If Err.Number <> 0 Then SafeCellColumn = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellWidth(cellObj)
  On Error Resume Next
  SafeCellWidth = cellObj.Width
  If Err.Number <> 0 Then SafeCellWidth = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellHeight(cellObj)
  On Error Resume Next
  SafeCellHeight = cellObj.Height
  If Err.Number <> 0 Then SafeCellHeight = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellTopPadding(cellObj)
  On Error Resume Next
  SafeCellTopPadding = cellObj.TopPadding
  If Err.Number <> 0 Then SafeCellTopPadding = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellRightPadding(cellObj)
  On Error Resume Next
  SafeCellRightPadding = cellObj.RightPadding
  If Err.Number <> 0 Then SafeCellRightPadding = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellBottomPadding(cellObj)
  On Error Resume Next
  SafeCellBottomPadding = cellObj.BottomPadding
  If Err.Number <> 0 Then SafeCellBottomPadding = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeCellLeftPadding(cellObj)
  On Error Resume Next
  SafeCellLeftPadding = cellObj.LeftPadding
  If Err.Number <> 0 Then SafeCellLeftPadding = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeBorderStyle(cellObj, borderKind)
  On Error Resume Next
  SafeBorderStyle = cellObj.Borders(borderKind).LineStyle
  If Err.Number <> 0 Then SafeBorderStyle = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeBorderWidth(cellObj, borderKind)
  On Error Resume Next
  SafeBorderWidth = cellObj.Borders(borderKind).LineWidth
  If Err.Number <> 0 Then SafeBorderWidth = "": Err.Clear
  On Error GoTo 0
End Function

Function SafeBorderColor(cellObj, borderKind)
  On Error Resume Next
  SafeBorderColor = cellObj.Borders(borderKind).Color
  If Err.Number <> 0 Then SafeBorderColor = "": Err.Clear
  On Error GoTo 0
End Function

Function Tsv(value)
  Dim text
  text = CStr(value)
  text = Replace(text, vbTab, " ")
  text = Replace(text, vbCr, " ")
  text = Replace(text, vbLf, " ")
  text = Replace(text, Chr(7), "")
  Tsv = text
End Function
