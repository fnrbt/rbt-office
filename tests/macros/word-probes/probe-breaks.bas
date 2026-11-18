Attribute VB_Name = "ProbeBreaks"
' Where InsertBreak puts a page break: on a whole paragraph's range, a collapsed range in a paragraph's
' text, the Selection in text, after a table; each case's paragraphs logged.
Dim logText As String

Sub Show(label As String)
    Dim p, s As String
    For Each p In ActiveDocument.Paragraphs
        s = s & "[" & Replace(Replace(p.Range.Text, vbCr, "/"), Chr(12), "F") & "]"
    Next
    logText = logText & label & ":" & s & "|"
End Sub

Sub Main()
    Dim d As Document, r As Range, t As Table
    Set d = ActiveDocument
    d.Content.Text = "one" & vbCr & "two"
    d.Paragraphs.Last.Range.InsertBreak wdPageBreak
    Show "a-lastrange"
    d.Content.Text = "one" & vbCr & "two"
    Set r = d.Range(1, 1)
    r.InsertBreak wdPageBreak
    Show "b-mid"
    L2 "b-range:" & r.Start & "," & r.End
    d.Content.Text = "one" & vbCr & "two"
    Set r = d.Range(4, 4)
    r.InsertBreak wdPageBreak
    Show "c-parastart"
    d.Content.Text = "one" & vbCr & "two"
    Selection.SetRange 5, 5
    Selection.InsertBreak wdPageBreak
    Show "d-selection"
    L2 "d-sel:" & Selection.Start
    d.Content.Text = "x"
    Set t = d.Tables.Add(d.Range(0, 0), 1, 1)
    d.Paragraphs.Last.Range.InsertBreak wdPageBreak
    Show "e-aftertable"
    d.Paragraphs.Last.Range.Text = "Feb"
    Show "e2"
    d.Content.Text = "abc"
    d.Content.InsertAfter "x" & Chr(12) & "y"
    Show "f-chr12"
    Debug.Print logText
    d.Variables.Add "log", logText
End Sub

Sub L2(s)
    logText = logText & s & "|"
End Sub
