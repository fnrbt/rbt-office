Attribute VB_Name = "FormLetters"
' Name: Form letters from data (mail merge without a data source)
' Source: none - written for this showcase (Selection.TypeText / TypeParagraph / InsertBreak, as in
'         https://learn.microsoft.com/en-us/office/vba/api/word.selection.typetext)
' Author: written for this showcase
' Licence: MIT (this repository)
' What it does: one letter per customer from arrays in the macro, typed with the Selection: a right-
'               aligned sender block, the date, a bold greeting, a paragraph with the customer's own figures
'               in bold, and a signature, with a page break between letters; then each customer's name is
'               found and coloured (Find loop).
' Adaptations: not applicable
' Entry point: Main
Sub Main()
    Dim names, cities, points, i As Integer, r As Range
    names = Array("Ada Lovelace", "Grace Hopper", "Alan Turing")
    cities = Array("London", "Arlington", "Manchester")
    points = Array(1250, 980, 2310)
    ActiveDocument.Content.Delete
    Selection.HomeKey wdStory
    For i = 0 To UBound(names)
        If i > 0 Then Selection.InsertBreak wdPageBreak
        Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
        Selection.Font.Size = 10
        Selection.Font.Color = RGB(89, 89, 89)
        Selection.TypeText "Rbt Office Club" & vbCr & "1 Macro Street" & vbCr & "Browserville"
        Selection.TypeParagraph
        Selection.TypeText Format(DateSerial(2026, 11, 18), "d mmmm yyyy")
        Selection.TypeParagraph
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Selection.Font.Size = 12
        Selection.Font.Color = wdColorAutomatic
        Selection.TypeParagraph
        Selection.Font.Bold = True
        Selection.TypeText "Dear " & names(i) & ","
        Selection.Font.Bold = False
        Selection.TypeParagraph
        Selection.TypeText "Thank you for another year with the club in " & cities(i) & ". You have collected "
        Selection.Font.Bold = True
        Selection.Font.Color = RGB(0, 112, 192)
        Selection.TypeText Format(points(i), "#,##0") & " points"
        Selection.Font.Bold = False
        Selection.Font.Color = wdColorAutomatic
        Selection.TypeText ", which you can spend on any of our courses."
        Selection.TypeParagraph
        Selection.TypeText "With best wishes,"
        Selection.TypeParagraph
        Selection.Font.Italic = True
        Selection.TypeText "The Rbt Office team"
        Selection.Font.Italic = False
    Next i
    For i = 0 To UBound(names)
        Set r = ActiveDocument.Content
        With r.Find
            .Text = names(i)
            Do While .Execute
                r.Font.Color = RGB(192, 0, 0)
                r.Collapse wdCollapseEnd
            Loop
        End With
    Next i
End Sub
