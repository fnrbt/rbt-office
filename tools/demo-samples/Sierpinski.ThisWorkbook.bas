' (Written for the rbt-office demo, MIT.) When macros are enabled: a button that draws it.
Private Sub Workbook_Open()
    Dim s As Object
    On Error Resume Next
    Set s = Sheets(1).Shapes("Draw")
    On Error GoTo 0
    If s Is Nothing Then
        With Sheets(1).Shapes.AddShape(msoShapeRoundedRectangle, 980, 16, 130, 36)
            .Name = "Draw"
            .TextFrame.Characters.Text = "Draw"
            .OnAction = "main"
        End With
    End If
End Sub
