Attribute VB_Name = "MonthCalendar"
' Name: Month calendars as Word tables
' Source: https://thelongestcode.blogspot.com/2025/11 ("Word VBA Macro create calendars for boox pdf",
'         Anthony, 3 November 2025: CreateMonthlyCalendar_Landscape / AKCreateMonthlyCalendar)
' Author: written for this showcase after that post's approach
' Licence: the post states none, so its code is not copied; this module is the repository's (MIT)
' What it does: for each month a Heading 1 title, then a table with a Mon..Sun header row and the days in
'               their weekday columns (Weekday(date, vbMonday), DateSerial), weekend cells shaded, today's
'               date highlighted, a page break between months.
' Adaptations: the year and the number of months are constants instead of an InputBox; no page margins.
' Entry point: Main
Const FirstMonth As Integer = 1
Const MonthCount As Integer = 3

Sub Main()
    Dim doc As Document, t As Table, m As Integer, y As Integer
    Dim first As Date, d As Date, rows As Integer, r As Integer, c As Integer, i As Integer
    Dim names
    names = Array("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")
    y = 2026
    Set doc = ActiveDocument
    doc.Content.Delete
    For m = FirstMonth To FirstMonth + MonthCount - 1
        first = DateSerial(y, m, 1)
        If m > FirstMonth Then doc.Paragraphs.Last.Range.InsertBreak wdPageBreak
        doc.Paragraphs.Last.Range.Text = Format(first, "mmmm yyyy")
        doc.Paragraphs.Last.Style = wdStyleTitle
        doc.Paragraphs.Last.Alignment = wdAlignParagraphCenter
        doc.Content.InsertParagraphAfter
        rows = 1 + (Weekday(first, vbMonday) - 1 + Day(DateSerial(y, m + 1, 0)) + 6) \ 7
        Set t = doc.Tables.Add(Range:=doc.Paragraphs.Last.Range, NumRows:=rows, NumColumns:=7)
        t.Borders.Enable = True
        t.Borders.OutsideLineWidth = wdLineWidth150pt
        For c = 1 To 7
            With t.Cell(1, c)
                .Range.Text = names(c - 1)
                .Range.Font.Bold = True
                .Range.Font.Color = wdColorWhite
                .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
                .Shading.BackgroundPatternColor = RGB(46, 84, 150)
            End With
        Next c
        r = 2
        c = Weekday(first, vbMonday)
        For i = 1 To Day(DateSerial(y, m + 1, 0))
            d = DateSerial(y, m, i)
            With t.Cell(r, c)
                .Range.Text = i
                .Range.Font.Size = 14
                If c > 5 Then
                    .Shading.BackgroundPatternColor = RGB(222, 235, 247)
                    .Range.Font.Color = RGB(192, 0, 0)
                End If
                If d = DateSerial(2026, 1, 1) Then .Range.HighlightColorIndex = wdYellow
            End With
            c = c + 1
            If c > 7 Then
                c = 1
                r = r + 1
            End If
        Next i
        t.Rows.Height = 40
    Next m
End Sub
