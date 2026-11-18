Attribute VB_Name = "SalesReport"
' Name: A formatted sales report generated from data in the macro
' Source: none - written for this showcase (it uses the patterns of the Microsoft Learn samples:
'         https://learn.microsoft.com/en-us/office/vba/word/concepts/working-with-word/working-with-tables)
' Author: written for this showcase
' Licence: MIT (this repository)
' What it does: a one-page report: Title and Subtitle, a Heading 1 per section, body text with bold and
'               coloured figures, a bulleted list of highlights (ListFormat.ApplyBulletDefault), a data
'               table built with ConvertToTable and styled "Grid Table 4 - Accent 1" with right-aligned
'               numbers and a total row, and a Quote at the end.
' Adaptations: not applicable
' Entry point: Main
Sub Add(text As String, Optional style As Variant)
    Dim d As Document
    Set d = ActiveDocument
    If Len(d.Content.Text) > 1 Then d.Content.InsertParagraphAfter
    d.Content.InsertAfter text
    If Not IsMissing(style) Then d.Paragraphs.Last.Style = style Else d.Paragraphs.Last.Style = wdStyleNormal
End Sub

Sub Main()
    Dim d As Document, regions, sales, i As Integer, total As Double, t As Table, r As Range, first As Long
    Set d = ActiveDocument
    d.Content.Delete
    regions = Array("North", "South", "East", "West", "Online")
    sales = Array(182500, 141200, 96750, 120300, 233900)
    For i = 0 To UBound(sales)
        total = total + sales(i)
    Next i
    Add "Quarterly Sales Report", wdStyleTitle
    Add "Third quarter " & Year(Date) & " - prepared by a VBA macro", wdStyleSubtitle
    Add "Summary", wdStyleHeading1
    Add "Sales reached " & Format(total, "$#,##0") & " this quarter, up 12% on the last. " & _
        "Online sales passed the North region for the first time."
    Set r = d.Paragraphs.Last.Range
    With r.Find
        .Text = Format(total, "$#,##0")
        If .Execute Then
            r.Font.Bold = True
            r.Font.Color = RGB(0, 112, 192)
        End If
    End With
    Add "Highlights", wdStyleHeading1
    first = d.Paragraphs.Count + 1
    Add "Online grew 31% and is now the largest channel."
    Add "The East region opened two new stores."
    Add "Returns fell below 2% for the first time."
    d.Range(d.Paragraphs(first).Range.Start, d.Paragraphs.Last.Range.End).ListFormat.ApplyBulletDefault
    Add "Sales by region", wdStyleHeading1
    first = d.Paragraphs.Count + 1
    Add "Region" & vbTab & "Sales" & vbTab & "Share"
    For i = 0 To UBound(regions)
        Add regions(i) & vbTab & Format(sales(i), "$#,##0") & vbTab & Format(sales(i) / total, "0.0%")
    Next i
    Add "Total" & vbTab & Format(total, "$#,##0") & vbTab & "100.0%"
    Set t = d.Range(d.Paragraphs(first).Range.Start, d.Paragraphs.Last.Range.End).ConvertToTable(Separator:=vbTab)
    t.Style = "Grid Table 4 - Accent 1"
    For i = 2 To t.Rows.Count
        t.Cell(i, 2).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
        t.Cell(i, 3).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
    Next i
    t.Rows(t.Rows.Count).Range.Font.Bold = True
    t.Rows(t.Rows.Count).Shading.BackgroundPatternColor = RGB(180, 198, 231)
    d.Content.InsertParagraphAfter
    Add "Good numbers are the result of many small decisions made well.", wdStyleQuote
End Sub
