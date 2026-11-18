Attribute VB_Name = "Prep"
' (Not from the source: what its workbook has - square cells for the board, and the "Toggle" button, a shape
' whose text the game sets, clicked to pause - set up here, the button drawn with Shapes.AddShape.)
Sub Setup()
    Columns("A:AA").ColumnWidth = 2.14
    Rows("1:30").RowHeight = 15
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Range("AC2").Left, Range("AC2").Top, 110, 34)
        .Name = "Toggle"
        .TextFrame.Characters.Text = "Start"
        .OnAction = "TogglePause"
    End With
End Sub
