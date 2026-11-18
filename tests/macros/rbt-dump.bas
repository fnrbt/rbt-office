Attribute VB_Name = "RbtDump"
' The document as tools/word-parity/run-word-macro.vbs dumps it from Word, written with Debug.Print by
' our object model: the same columns, for comparing a macro's result in Word and in ours.

Private Function Esc(s)
    Dim t
    t = Replace(CStr(s), "\", "\\")
    t = Replace(t, vbTab, "\t")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, Chr(7), "\a")
    t = Replace(t, Chr(11), "\v")
    t = Replace(t, Chr(12), "\f")
    Esc = t
End Function

Public Sub Dump()
    Dim d, p, i, r, f, t, ti, ri, ci
    Set d = ActiveDocument
    Debug.Print "counts" & vbTab & d.Paragraphs.Count & vbTab & d.Words.Count & vbTab & d.Characters.Count & vbTab & d.Tables.Count & vbTab & d.Content.End & vbTab & d.ComputeStatistics(0) & vbTab & d.ComputeStatistics(2)
    Debug.Print "selection" & vbTab & Selection.Start & vbTab & Selection.End
    i = 0
    For Each p In d.Paragraphs
        i = i + 1
        Set r = p.Range
        Set f = r.Characters(1).Font
        Debug.Print "p" & vbTab & i & vbTab & r.Start & vbTab & r.End & vbTab & Esc(r.Text) & vbTab & p.Style & vbTab & p.Alignment & vbTab & _
            p.SpaceBefore & vbTab & p.SpaceAfter & vbTab & p.LeftIndent & vbTab & r.Information(12) & vbTab & _
            f.Name & vbTab & f.Size & vbTab & f.Bold & vbTab & f.Italic & vbTab & f.Underline & vbTab & f.Color & vbTab & r.Characters(1).HighlightColorIndex & vbTab & _
            r.Font.Bold & vbTab & r.Font.Size & vbTab & r.Font.Color & vbTab & r.ListFormat.ListType & vbTab & p.Shading.BackgroundPatternColor
    Next
    ti = 0
    For Each t In d.Tables
        ti = ti + 1
        Debug.Print "table" & vbTab & ti & vbTab & t.Rows.Count & vbTab & t.Columns.Count & vbTab & t.Range.Start & vbTab & t.Range.End & vbTab & t.Borders.Enable
        For ri = 1 To t.Rows.Count
            For ci = 1 To t.Rows(ri).Cells.Count
                Set r = t.Cell(ri, ci).Range
                Debug.Print "cell" & vbTab & ti & vbTab & ri & vbTab & ci & vbTab & Esc(r.Text) & vbTab & t.Cell(ri, ci).Shading.BackgroundPatternColor & vbTab & r.Characters(1).Font.Bold & vbTab & r.Characters(1).Font.Color
            Next
        Next
    Next
End Sub
