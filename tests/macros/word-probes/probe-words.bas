Attribute VB_Name = "ProbeWords"
' Word's Words and Sentences on the showcase texts: every word logged between |.
Sub Main()
    Dim w, s As String
    ActiveDocument.Content.Text = "hi there. The sky is blue, the sea is blue, and BOLD words stand out." & vbCr & _
        "C and C++ run them. Java and C# power it; F# and OCaml too. It's 1,250 points (2.5%) at $1,000.00 - ok?" & vbCr & _
        "e.g. Mr. Smith's well-known end"
    For Each w In ActiveDocument.Words
        s = s & Replace(w.Text, vbCr, "\r") & "|"
    Next
    s = s & "#" & ActiveDocument.Sentences.Count
    For Each w In ActiveDocument.Sentences
        s = s & "{" & Replace(w.Text, vbCr, "\r") & "}"
    Next
    Debug.Print s
    ActiveDocument.Variables.Add "log", s
End Sub
