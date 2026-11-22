Option Explicit
' Run with fixtures/column-edit.py and run-word-vbs-on-vm.sh.
' Each case starts from the original document. x is the start of cell text
' relative to the page; zero cell margins make it the cell's left edge.
Const wdPrintView = 3
Const wdHorizontalPositionRelativeToPage = 5
Const wdDoNotSaveChanges = 0
Dim fso, word, outFile, inputPath, outputPath, col, direction
Set fso = CreateObject("Scripting.FileSystemObject")
inputPath = fso.GetAbsolutePathName(WScript.Arguments(0))
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))
Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine "case" & vbTab & "phase" & vbTab & "row" & vbTab & "column" & vbTab & "text" & vbTab & "x_pt" & vbTab & "width_pt" & vbTab & "fill_color"
Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
For col = 1 To 2
  For Each direction In Array("right", "left")
    RunCase col, direction
  Next
Next
outFile.Close
word.Quit

Sub RunCase(sourceColumn, side)
  Dim doc, table, insertedColumn, name, tempPath
  name = side & "-of-" & sourceColumn
  Set doc = word.Documents.Open(inputPath, False, True, False)
  tempPath = outputPath & "-" & name & ".docx"
  doc.SaveAs2 tempPath, 16
  doc.Close wdDoNotSaveChanges
  Set doc = word.Documents.Open(tempPath, False, False, False)
  word.ActiveWindow.View.Type = wdPrintView
  Set table = doc.Tables(1)
  Emit doc, table, name, "before"
  table.Cell(1, sourceColumn).Range.Select
  If side = "right" Then
    word.Selection.InsertColumnsRight
    insertedColumn = sourceColumn + 1
  Else
    word.Selection.InsertColumns
    insertedColumn = sourceColumn
  End If
  doc.Save
  Emit doc, table, name, "inserted"
  table.Cell(1, insertedColumn).Range.Select
  word.Selection.Cells(1).Column.Delete
  doc.Save
  Emit doc, table, name, "deleted"
  doc.Close wdDoNotSaveChanges
  fso.DeleteFile tempPath
End Sub

Sub Emit(doc, table, name, phase)
  Dim cell, text, xml, nodes, node, widths, kind, startRange
  doc.Repaginate
  Set xml = CreateObject("MSXML2.DOMDocument.6.0")
  xml.async = False
  If Not xml.loadXML(doc.WordOpenXML) Then Err.Raise 1, , xml.parseError.reason
  xml.setProperty "SelectionNamespaces", "xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'"
  For Each kind In Array("tblW", "gridCol", "tcW")
    Set nodes = xml.selectNodes("(//w:tbl)[1]//w:" & kind)
    widths = ""
    For Each node In nodes
      If Len(widths) > 0 Then widths = widths & ","
      widths = widths & node.getAttribute("w:w") & ":" & node.getAttribute("w:type")
    Next
    outFile.WriteLine name & vbTab & phase & vbTab & "xml" & vbTab & kind & vbTab & widths & vbTab & vbTab
  Next
  For Each cell In table.Range.Cells
    text = Replace(Replace(Replace(cell.Range.Text, Chr(13), ""), Chr(7), ""), vbTab, " ")
    Set startRange = cell.Range.Duplicate
    startRange.Collapse 1
    outFile.WriteLine name & vbTab & phase & vbTab & cell.RowIndex & vbTab & cell.ColumnIndex & vbTab & text & vbTab & startRange.Information(wdHorizontalPositionRelativeToPage) & vbTab & cell.Width & vbTab & cell.Shading.BackgroundPatternColor
  Next
End Sub
