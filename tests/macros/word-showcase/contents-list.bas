Attribute VB_Name = "ContentsList"
' Name: A contents list built from the headings
' Source: https://www.oreilly.com/library/view/writing-word-macros/9781565927254/ch12s11.html
'         (Steven Roman, "Writing Word Macros", "Example: Printing Document Headings": loop over the
'         paragraphs and take those whose style is like "Heading [0-9]")
' Author: written for this showcase after that example's approach
' Licence: the book's code is not copied; this module is the repository's (MIT)
' What it does: walks ActiveDocument.Paragraphs, collects the Heading 1-3 paragraphs (Style Like
'               "Heading #", level 3 at most), and puts a "Contents" block at the top: one line per heading, indented
'               by its level (LeftIndent), numbered 1, 1.1, 1.1.1, with the paragraph's position, shaded
'               and framed with a border.
' Adaptations: it writes the list into the document instead of printing it.
' Needs: a document with headings (Setup writes a short report with Heading 1-3 paragraphs)
' Entry point: Main
Sub Setup()
    Dim parts, i As Integer, d As Document
    Set d = ActiveDocument
    parts = Array( _
        "1|Introduction", "0|This report looks at how the team ships software.", _
        "2|Goals", "0|Faster releases, fewer regressions.", _
        "2|Scope", "0|The web client and the desktop host.", _
        "1|Findings", "0|Three things stood out.", _
        "2|Testing", "0|Coverage grew by a third.", _
        "3|Browser checks", "0|Screenshots compared on every change.", _
        "3|Native checks", "0|The same suites on .NET.", _
        "2|Releases", "0|One a week since March.", _
        "1|Next steps", "0|Keep the checks fast.")
    d.Content.Delete
    For i = 0 To UBound(parts)
        If i > 0 Then d.Content.InsertParagraphAfter
        d.Content.InsertAfter Split(parts(i), "|")(1)
        If Left(parts(i), 1) <> "0" Then
            d.Paragraphs.Last.Style = d.Styles("Heading " & Left(parts(i), 1))
        Else
            d.Paragraphs.Last.Style = wdStyleNormal
        End If
    Next i
End Sub

Sub Main()
    Dim p As Paragraph, lines As String, level As Integer, n(3) As Integer, number As String
    Dim r As Range, count As Integer, i As Integer
    Setup
    For Each p In ActiveDocument.Paragraphs
        ' (Not Like "Heading [1-3]": Rbt.Vb 0.4.0 does not take ranges in Like's brackets.)
        If p.Style Like "Heading #" And Val(Right(p.Style, 1)) <= 3 Then
            level = Val(Right(p.Style, 1))
            n(level) = n(level) + 1
            For i = level + 1 To 3
                n(i) = 0
            Next i
            number = n(1)
            If level >= 2 Then number = number & "." & n(2)
            If level >= 3 Then number = number & "." & n(3)
            lines = lines & level & "|" & number & "   " & Replace(p.Range.Text, vbCr, "") & vbTab & "@" & p.Range.Start & vbCr
            count = count + 1
        End If
    Next p
    Set r = ActiveDocument.Range(0, 0)
    r.InsertBefore "Contents" & vbCr & Left(lines, Len(lines) - 1) & vbCr
    ActiveDocument.Paragraphs(1).Style = wdStyleTitle
    For i = 2 To count + 1
        With ActiveDocument.Paragraphs(i)
            level = Val(Left(.Range.Text, 1))
            .Range.Characters(1).Delete
            .Range.Characters(1).Delete
            .LeftIndent = (level - 1) * 24
            .SpaceAfter = 2
            .Shading.BackgroundPatternColor = RGB(242, 242, 242)
            If level = 1 Then .Range.Font.Bold = True
            .Range.Font.Color = RGB(31, 56, 100)
        End With
    Next i
    With ActiveDocument.Paragraphs(count + 1).Borders(wdBorderBottom)
        .LineStyle = wdLineStyleSingle
        .LineWidth = wdLineWidth150pt
        .Color = RGB(31, 56, 100)
    End With
End Sub
