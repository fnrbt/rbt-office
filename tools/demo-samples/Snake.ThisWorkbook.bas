Private Sub Workbook_Open()
    Dim s As Object
    On Error Resume Next
    Set s = Sheets(1).Shapes("NewGame")
    On Error GoTo 0
    If s Is Nothing Then MakeBoard
End Sub
