Option Explicit

Const xlNormalView = 1
Const xlPrintView = 2
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105
Const xlDoNotSaveChanges = 2
Const xlAlertsNone = False

Dim fso, inputPath, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo extract-excel-layout.vbs input.xlsx output.tsv"
  WScript.Quit 2
End If

inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

If Not fso.FileExists(inputPath) Then
  WScript.Echo "input not found: " & inputPath
  WScript.Quit 2
End If

Dim excel, workbook, outFile
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = xlAlertsNone
excel.ScreenUpdating = False
excel.EnableEvents = False
excel.AskToUpdateLinks = False

On Error Resume Next
excel.Calculation = xlCalculationManual
On Error GoTo 0

Set workbook = excel.Workbooks.Open(inputPath, False, True)

Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine _
  "doc" & vbTab & _
  "sheet" & vbTab & _
  "sheetIndex" & vbTab & _
  "row" & vbTab & _
  "column" & vbTab & _
  "address" & vbTab & _
  "x" & vbTab & _
  "y" & vbTab & _
  "width" & vbTab & _
  "height" & vbTab & _
  "mergedRows" & vbTab & _
  "mergedCols" & vbTab & _
  "rowHeight" & vbTab & _
  "columnWidth" & vbTab & _
  "wrapText" & vbTab & _
  "hAlign" & vbTab & _
  "vAlign" & vbTab & _
  "fontName" & vbTab & _
  "fontSize" & vbTab & _
  "bold" & vbTab & _
  "italic" & vbTab & _
  "numberFormat" & vbTab & _
  "style" & vbTab & _
  "formula" & vbTab & _
  "value" & vbTab & _
  "text"

Dim sheet, sheetIndex
sheetIndex = 0
For Each sheet In workbook.Worksheets
  sheetIndex = sheetIndex + 1
  If sheet.Visible = True Then
    EmitSheet workbook, sheet, sheetIndex, inputPath, outFile
  End If
Next

outFile.Close
workbook.Close xlDoNotSaveChanges

On Error Resume Next
excel.Calculation = xlCalculationAutomatic
On Error GoTo 0

excel.Quit

Sub EmitSheet(wb, ws, sheetIndex, docPath, fileObj)
  Dim used, cell, address, mergedRows, mergedCols
  Dim rowHeight, columnWidth, hAlign, vAlign, fontName, fontSize, bold, italic, numFmt, styleName, formula, valueText, displayText

  Set used = ws.UsedRange
  If used Is Nothing Then Exit Sub
  If used.Cells.Count = 0 Then Exit Sub

  For Each cell In used.Cells
    If Not IsEmptyCell(cell) Then
      address = SafeAddress(cell)
      mergedRows = 1
      mergedCols = 1
      If cell.MergeCells Then
        mergedRows = cell.MergeArea.Rows.Count
        mergedCols = cell.MergeArea.Columns.Count
      End If

      rowHeight = SafeDouble(cell.RowHeight)
      columnWidth = SafeDouble(cell.ColumnWidth)
      hAlign = SafeHAlign(cell.HorizontalAlignment)
      vAlign = SafeVAlign(cell.VerticalAlignment)
      fontName = SafeString(cell.Font.Name)
      fontSize = SafeDouble(cell.Font.Size)
      bold = SafeBool(cell.Font.Bold)
      italic = SafeBool(cell.Font.Italic)
      numFmt = SafeString(cell.NumberFormat)
      styleName = SafeStyleName(cell)
      formula = SafeFormula(cell)
      valueText = SafeValue(cell)
      displayText = SafeText(cell)

      fileObj.WriteLine _
        Tsv(docPath) & vbTab & _
        Tsv(ws.Name) & vbTab & _
        Tsv(CStr(sheetIndex)) & vbTab & _
        Tsv(CStr(cell.Row)) & vbTab & _
        Tsv(CStr(cell.Column)) & vbTab & _
        Tsv(address) & vbTab & _
        Tsv(CStr(cell.Left)) & vbTab & _
        Tsv(CStr(cell.Top)) & vbTab & _
        Tsv(CStr(cell.Width)) & vbTab & _
        Tsv(CStr(cell.Height)) & vbTab & _
        Tsv(CStr(mergedRows)) & vbTab & _
        Tsv(CStr(mergedCols)) & vbTab & _
        Tsv(rowHeight) & vbTab & _
        Tsv(columnWidth) & vbTab & _
        Tsv(SafeBool(cell.WrapText)) & vbTab & _
        Tsv(hAlign) & vbTab & _
        Tsv(vAlign) & vbTab & _
        Tsv(fontName) & vbTab & _
        Tsv(fontSize) & vbTab & _
        Tsv(bold) & vbTab & _
        Tsv(italic) & vbTab & _
        Tsv(numFmt) & vbTab & _
        Tsv(styleName) & vbTab & _
        Tsv(formula) & vbTab & _
        Tsv(valueText) & vbTab & _
        Tsv(displayText)
    End If
  Next
End Sub

Function IsEmptyCell(cell)
  Dim v, formulaText
  formulaText = ""
  On Error Resume Next
  formulaText = cell.Formula
  v = cell.Value
  On Error GoTo 0

  If formulaText <> "" Then
    IsEmptyCell = False
    Exit Function
  End If

  If IsEmpty(v) Then
    IsEmptyCell = True
  ElseIf VarType(v) = vbString And v = "" Then
    IsEmptyCell = True
  Else
    IsEmptyCell = False
  End If
End Function

Function SafeAddress(cell)
  On Error Resume Next
  SafeAddress = cell.Address(False, False, 1, False)
  If Err.Number <> 0 Then
    SafeAddress = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeStyleName(cell)
  On Error Resume Next
  SafeStyleName = cell.Style.Name
  If Err.Number <> 0 Then
    SafeStyleName = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeFormula(cell)
  On Error Resume Next
  If cell.HasFormula Then
    SafeFormula = cell.Formula
  Else
    SafeFormula = ""
  End If
  If Err.Number <> 0 Then
    SafeFormula = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeValue(cell)
  Dim v
  On Error Resume Next
  v = cell.Value2
  If Err.Number <> 0 Then
    SafeValue = ""
    Err.Clear
    On Error GoTo 0
    Exit Function
  End If
  On Error GoTo 0

  If IsEmpty(v) Then
    SafeValue = ""
  ElseIf VarType(v) = vbDate Then
    SafeValue = CStr(CDbl(v))
  Else
    SafeValue = CStr(v)
  End If
End Function

Function SafeText(cell)
  On Error Resume Next
  SafeText = CStr(cell.Text)
  If Err.Number <> 0 Then
    SafeText = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeDouble(value)
  On Error Resume Next
  SafeDouble = CStr(CDbl(value))
  If Err.Number <> 0 Then
    SafeDouble = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeBool(value)
  On Error Resume Next
  If IsNull(value) Then
    SafeBool = ""
  ElseIf value Then
    SafeBool = "true"
  Else
    SafeBool = "false"
  End If
  If Err.Number <> 0 Then
    SafeBool = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeString(value)
  On Error Resume Next
  SafeString = CStr(value)
  If Err.Number <> 0 Then
    SafeString = ""
    Err.Clear
  End If
  On Error GoTo 0
End Function

Function SafeHAlign(value)
  Select Case value
    Case 1
      SafeHAlign = "general"
    Case -4131
      SafeHAlign = "left"
    Case -4108
      SafeHAlign = "center"
    Case -4152
      SafeHAlign = "right"
    Case -4130
      SafeHAlign = "fill"
    Case -4117
      SafeHAlign = "justify"
    Case 7
      SafeHAlign = "centerAcrossSelection"
    Case 8
      SafeHAlign = "distributed"
    Case Else
      SafeHAlign = CStr(value)
  End Select
End Function

Function SafeVAlign(value)
  Select Case value
    Case -4160
      SafeVAlign = "top"
    Case -4108
      SafeVAlign = "center"
    Case -4107
      SafeVAlign = "bottom"
    Case -4117
      SafeVAlign = "justify"
    Case -4117
      SafeVAlign = "distributed"
    Case Else
      SafeVAlign = CStr(value)
  End Select
End Function

Function Tsv(value)
  Dim text
  text = CStr(value)
  text = Replace(text, vbTab, " ")
  text = Replace(text, vbCr, " ")
  text = Replace(text, vbLf, " ")
  Tsv = text
End Function
