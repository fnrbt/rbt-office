Attribute VB_Name = "Board"
' (Written for the rbt-office demo, MIT: what the game's own workbook has - square cells for the board and
' its buttons, shapes whose OnAction runs a macro.)
Sub MakeBoard()
    Columns("A:AA").ColumnWidth = 2.14
    Rows("1:30").RowHeight = 15
    AddButton "NewGame", "New game", "StartGame", Range("AC2").Top
    AddButton "Toggle", "Pause", "TogglePause", Range("AC5").Top
    Range("AC8").Value = "Arrow keys steer."
    ResetGame
    ActiveSheet.Shapes("Toggle").TextFrame.Characters.Text = "Pause"
End Sub

Private Sub AddButton(name As String, caption As String, macro As String, top As Double)
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, Range("AC2").Left, top, 110, 34)
        .Name = name
        .TextFrame.Characters.Text = caption
        .OnAction = macro
    End With
End Sub
