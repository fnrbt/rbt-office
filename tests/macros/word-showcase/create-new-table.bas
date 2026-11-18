Attribute VB_Name = "CreateTable"
' Name: Create a table, insert text, apply formatting
' Source: https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-tables
'         ("Creating a table, inserting text, and applying formatting")
' Author: Microsoft (MicrosoftDocs/VBA-Docs)
' Licence: code samples MIT (MicrosoftDocs/VBA-Docs LICENSE-CODE), text CC BY 4.0
' What it does: a 4-column, 3-row table at the start of the document, "Cell 1" .. "Cell 12" in its cells,
'               then Table.AutoFormat with wdTableFormatColorful2.
' Adaptations: none to CreateNewTable; Main (added) runs it and writes a caption under the table (the
'              table is at the very start, so text inserted at position 0 would go into its first cell).
' Entry point: Main
Sub CreateNewTable()
 Dim docActive As Document
 Dim tblNew As Table
 Dim celTable As Cell
 Dim intCount As Integer

 Set docActive = ActiveDocument
 Set tblNew = docActive.Tables.Add( _
 Range:=docActive.Range(Start:=0, End:=0), NumRows:=3, _
 NumColumns:=4)
 intCount = 1

 For Each celTable In tblNew.Range.Cells
  celTable.Range.InsertAfter "Cell " & intCount
  intCount = intCount + 1
 Next celTable

 tblNew.AutoFormat Format:=wdTableFormatColorful2, _
 ApplyBorders:=True, ApplyFont:=True, ApplyColor:=True
End Sub

Sub Main()
    CreateNewTable
    ActiveDocument.Paragraphs.Last.Range.Text = "A table from Microsoft's VBA documentation, AutoFormat wdTableFormatColorful2"
    ActiveDocument.Paragraphs.Last.Style = wdStyleCaption
End Sub
