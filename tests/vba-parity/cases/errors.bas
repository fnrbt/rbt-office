Sub Main()
    On Error Resume Next
    Dim x, o As Object
    x = 1 / 0: P "div by zero", Err.Number: P "div text", Err.Description: Err.Clear
    x = "a" * 2: P "type mismatch", Err.Number: Err.Clear
    Dim arr: arr = Array(1): x = arr(5): P "subscript", Err.Number: Err.Clear
    x = o.Name: P "object not set", Err.Number: Err.Clear
    x = CInt(99999): P "overflow", Err.Number: Err.Clear
    x = Mid("abc", 0): P "invalid call", Err.Number: Err.Clear
    Err.Raise vbObjectError + 513, "MySource", "custom"
    P "raise number", Err.Number - vbObjectError
    P "raise source", Err.Source
    P "raise description", Err.Description
    Err.Clear
    P "cleared", Err.Number
    On Error GoTo 0
    P "handler resume next", Handled()
    P "handler resume label", Retried()
    P "nested handler", Outer()
    P "error in handler caller", CallsFailing()
    P "typename err", TypeName(Err)
    P "erl", WithLines()
End Sub
Function Handled()
    On Error GoTo h
    Dim y
    y = 1 / 0
    Handled = "after:" & y
    Exit Function
h:
    y = -1
    Resume Next
End Function
Function Retried()
    Dim tries
    On Error GoTo h
    FailBelow tries, 3
    Retried = "tries:" & tries
    Exit Function
h:
    tries = tries + 1
    Resume
End Function
Sub FailBelow(n, limit)
    If n < limit Then Err.Raise 5
End Sub
Function Outer()
    On Error GoTo h
    Inner
    Outer = "no error"
    Exit Function
h:
    Outer = "caught " & Err.Number & " from inner"
End Function
Sub Inner()
    Err.Raise 1004, , "deep"
End Sub
Function CallsFailing()
    On Error Resume Next
    Inner
    CallsFailing = Err.Number & " " & Err.Description
End Function
Function WithLines()
    On Error GoTo h
10  Dim a
20  a = 1 / 0
30  WithLines = "none"
    Exit Function
h:
    WithLines = "line " & Erl
End Function
