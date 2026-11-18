Attribute VB_Name = "Multiplication"
' Name: Multiplication table
' Source: https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-tables
'         (Tables.Add and filling cells through Cell(row, column).Range, MIT-licensed samples)
' Author: written for this showcase on that page's pattern
' Licence: MIT (this repository)
' What it does: a 10 x 10 multiplication table (Tables.Add, Cell(r, c).Range.Text), the header row and
'               column shaded dark with white bold text, the squares on the diagonal gold, the cells
'               shaded on a gradient by the size of the product (RGB), centred numbers.
' Adaptations: not applicable
' Entry point: Main
Sub Main()
    Dim doc As Document, t As Table, r As Integer, c As Integer, v As Integer, shade As Integer
    Set doc = ActiveDocument
    doc.Content.Delete
    doc.Content.Text = "Times tables"
    doc.Paragraphs(1).Style = wdStyleTitle
    doc.Content.InsertParagraphAfter
    Set t = doc.Tables.Add(doc.Paragraphs.Last.Range, 11, 11, wdWord9TableBehavior)
    For r = 1 To 11
        For c = 1 To 11
            With t.Cell(r, c)
                If r = 1 And c = 1 Then
                    .Range.Text = "x"
                ElseIf r = 1 Then
                    .Range.Text = c - 1
                ElseIf c = 1 Then
                    .Range.Text = r - 1
                Else
                    v = (r - 1) * (c - 1)
                    .Range.Text = v
                    shade = 255 - v * 1.5
                    .Shading.BackgroundPatternColor = RGB(shade, shade, 255)
                    If r = c Then
                        .Shading.BackgroundPatternColor = RGB(255, 192, 0)
                        .Range.Font.Bold = True
                    End If
                End If
                If r = 1 Or c = 1 Then
                    .Shading.BackgroundPatternColor = RGB(31, 56, 100)
                    .Range.Font.Color = wdColorWhite
                    .Range.Font.Bold = True
                End If
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Range.ParagraphFormat.SpaceAfter = 0
            End With
        Next c
    Next r
    t.Range.Font.Size = 13
    t.Borders.InsideColor = wdColorWhite
    t.Borders.OutsideColor = RGB(31, 56, 100)
End Sub
