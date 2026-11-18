Attribute VB_Name = "Prep"
' (Not from the source: the "Toggle" button its workbook has - a shape whose text the game sets, clicked
' to pause - drawn here with Shapes.AddShape.)
Sub Setup()
    With ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 560, 20, 100, 32)
        .Name = "Toggle"
        .TextFrame.Characters.Text = "Start"
        .OnAction = "TogglePause"
    End With
End Sub
