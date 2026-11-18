Sub Main()
    Dim a(1 To 3) As Long
    a(1) = 10: a(2) = 20: a(3) = 30
    P "lbound", LBound(a)
    P "ubound", UBound(a)
    Dim b() As String
    ReDim b(2)
    b(0) = "x": b(2) = "z"
    ReDim Preserve b(4)
    P "preserve", b(2) & "|" & b(4) & "|" & UBound(b)
    Dim m(1 To 2, 1 To 3)
    m(2, 3) = 7
    P "2d ubound dim 2", UBound(m, 2)
    P "2d value", m(2, 3)
    Dim v
    v = Array(1, "two", 3.5)
    P "array typename", TypeName(v)
    P "array item", v(1)
    P "isarray", IsArray(v)
    Erase a
    P "erase fixed", a(2)
    Dim s
    s = Split("p q r")
    P "split default", s(2)
    Double3 a
    P "byref array", a(1)
    Dim e()
    P "empty dynamic ubound", "unset"
    On Error Resume Next
    P "ubound unallocated", UBound(e)
    P "ubound unallocated err", Err.Number
    Err.Clear
    P "out of range", a(9)
    P "out of range err", Err.Number
End Sub
Sub Double3(arr)
    arr(1) = 99
End Sub
