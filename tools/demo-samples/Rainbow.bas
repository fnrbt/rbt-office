Attribute VB_Name = "Rainbow"
' Every character of the document a colour of a rainbow, the title larger. Written for the rbt-office demo (MIT).
Sub Rainbow()
    Dim i As Long, k As Double
    For i = 1 To ActiveDocument.Characters.Count
        k = i / 4
        ActiveDocument.Characters(i).Font.Color = RGB(128 + 127 * Sin(k), 128 + 127 * Sin(k + 2.1), 128 + 127 * Sin(k + 4.2))
    Next
    ActiveDocument.Paragraphs(1).Range.Font.Size = 36
    ActiveDocument.Paragraphs(1).Range.Font.Bold = True
End Sub

Sub BackToBlack()
    ActiveDocument.Content.Font.Color = wdColorAutomatic
End Sub
