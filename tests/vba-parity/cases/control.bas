Sub Main()
    P "select is", Classify(15)
    P "select to", Classify(5)
    P "select list", Classify(42)
    P "select else", Classify(-1)
    P "select string", Grade("b")
    Dim i, n
    Do While i < 5
        i = i + 1
    Loop
    P "do while", i
    i = 0
    Do
        i = i + 2
    Loop Until i >= 7
    P "loop until", i
    For i = 10 To 1 Step -3
        n = n & i & ","
    Next
    P "for step", n
    P "for after", i
    For i = 1 To 10
        If i = 4 Then Exit For
    Next
    P "exit for", i
    P "iif", IIf(1 > 2, "yes", "no")
    P "choose", Choose(2, "a", "b", "c")
    P "switch", Switch(False, 1, True, 2)
    Dim c As New Collection
    c.Add "first"
    c.Add "second", "k2"
    c.Add "zero", Before:=1
    P "collection count", c.Count
    P "collection item", c(1) & "|" & c("k2")
    Dim t, item
    For Each item In c
        t = t & item & ";"
    Next
    P "collection for each", t
    c.Remove "k2"
    P "collection remove", c.Count
    On Error Resume Next
    P "collection missing", c("nope")
    P "collection missing err", Err.Number
    Err.Clear
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    d("b") = 2: d("a") = 1
    d.Add "c", 3
    P "dict keys", Join(d.Keys, ",")
    P "dict exists", d.Exists("a")
    P "dict count", d.Count
    d.Add "a", 9
    P "dict duplicate err", Err.Number
    Err.Clear
    d.CompareMode = vbTextCompare
    P "dict comparemode err", Err.Number
    Err.Clear
    P "gosub", ViaGoSub()
    With Worksheets(1)
        .Range("F1").Value = "with"
        With .Range("F2")
            .Value = "inner"
        End With
        .Range("F3").Value = .Name
    End With
End Sub
Function Classify(x)
    Select Case x
        Case Is > 10: Classify = "big"
        Case 1 To 9: Classify = "small"
        Case 42, 43: Classify = "answer"
        Case Else: Classify = "other"
    End Select
End Function
Function Grade(g)
    Select Case UCase(g)
        Case "A" To "C": Grade = "good"
        Case Else: Grade = "bad"
    End Select
End Function
Function ViaGoSub()
    Dim r
    GoSub inner
    ViaGoSub = r
    Exit Function
inner:
    r = "returned"
    Return
End Function
