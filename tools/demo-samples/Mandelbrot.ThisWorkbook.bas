' When macros are enabled: a button that draws it (a shape whose OnAction is the macro).
Private Sub Workbook_Open()
    Dim s As Shape
    On Error Resume Next
    Set s = Sheets(1).Shapes("Draw")
    On Error GoTo 0
    If s Is Nothing Then
        ' (The picture's square cells first: the button goes beside them.)
        Range(Columns(1), Columns(96)).ColumnWidth = 1
        Range(Rows(1), Rows(64)).RowHeight = 9
        With Sheets(1).Shapes.AddShape(msoShapeRoundedRectangle, Sheets(1).Cells(2, 99).Left, 16, 130, 36)
            .Name = "Draw"
            .TextFrame.Characters.Text = "Draw"
            .OnAction = "DrawMandelbrot"
        End With
    End If
End Sub
