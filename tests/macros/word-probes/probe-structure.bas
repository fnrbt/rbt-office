Attribute VB_Name = "Probe"
' Positions and structure as Word has them: what a collapsed Content range is, where text goes at the
' very end, the style after InsertParagraphAfter, vbLf, the formatting a new paragraph's text gets from
' the mark, a table's positions and borders, Paragraphs.Count with a table, Words. Run in Word
' (tools/word-parity/word-macro-on-vm.sh) and in ours (run-macro-native.fsx) and compare the logs.
Dim logText As String

Sub L(s)
    logText = logText & s & "|"
End Sub

Sub Main()
    Dim r As Range, t As Table, d As Document
    Set d = ActiveDocument
    L d.Content.End
    d.Content.InsertAfter "Hello"
    L d.Content.End & "," & d.Paragraphs.Count
    Set r = d.Content
    r.Collapse wdCollapseEnd
    L r.Start & "," & r.End
    r.Text = "X"
    L r.Start & "," & r.End & "," & Replace(d.Content.Text, vbCr, "/")
    d.Paragraphs(1).Style = wdStyleHeading1
    d.Paragraphs(1).Range.InsertParagraphAfter
    L d.Paragraphs.Count & "," & d.Paragraphs(2).Style
    d.Paragraphs(2).Range.InsertBefore "A" & vbLf & "B"
    L d.Paragraphs.Count & "," & Len(d.Content.Text)
    Set r = d.Paragraphs(d.Paragraphs.Count).Range
    r.Font.Bold = True
    r.InsertParagraphAfter
    r.InsertAfter "after"
    L r.Start & "," & r.End & "," & d.Paragraphs(d.Paragraphs.Count).Range.Font.Bold
    d.Content.InsertParagraphAfter
    Set t = d.Tables.Add(d.Paragraphs(d.Paragraphs.Count).Range, 2, 2)
    L d.Paragraphs.Count & "," & t.Borders.Enable & "," & t.Range.Start & "," & t.Range.End & "," & d.Content.End
    L Len(t.Cell(1, 1).Range.Text) & "," & (t.Cell(1, 1).Range.End - t.Cell(1, 1).Range.Start) & "," & (t.Rows(1).Range.End - t.Rows(1).Range.Start)
    t.Cell(1, 1).Range.Text = "a"
    L t.Cell(1, 1).Range.Text = "a" & vbCr & Chr(7)
    L d.Words.Count & "," & d.Characters.Count & "," & d.Sentences.Count
    Set r = d.Range(0, 0)
    r.InsertAfter "Lead "
    L r.Start & "," & r.End
    r.InsertBefore ">"
    L r.Start & "," & r.End & "," & r.Text
    Selection.HomeKey wdStory
    Selection.TypeText "S"
    L Selection.Start & "," & Selection.End & "," & Selection.Type
    Selection.EndKey wdStory
    L Selection.Start
    Debug.Print logText
    d.Variables.Add "log", logText
End Sub
