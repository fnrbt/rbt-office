Type Point
    X As Double
    Y As Double
End Type
Enum Colour
    Red = 1
    Green
    Blue = 10
End Enum
Private counter As Long
Sub Main()
    Dim pt As Point, q As Point
    pt.X = 3: pt.Y = 4
    q = pt
    q.X = 100
    P "udt copy", pt.X & "," & q.X
    P "udt calc", Dist(pt)
    P "enum", Green + Blue
    P "static", Bump() & Bump() & Bump()
    P "module var", counter
    P "optional default", Opt(1)
    P "optional given", Opt(1, 5)
    P "ismissing", Missing()
    P "paramarray", Total(1, 2, 3, 4)
    P "named args", Opt(b:=7, a:=1)
    P "byval", ByValTest()
    P "recursion", Fact(10)
    P "function as array", Squares()(3)
End Sub
Function Dist(p As Point) As Double
    Dist = Sqr(p.X ^ 2 + p.Y ^ 2)
End Function
Function Bump()
    Static n
    n = n + 1
    counter = counter + 1
    Bump = n
End Function
Function Opt(a, Optional b = 10)
    Opt = a + b
End Function
Function Missing(Optional v)
    Missing = IsMissing(v)
End Function
Function Total(ParamArray xs())
    Dim x, t
    For Each x In xs
        t = t + x
    Next
    Total = t
End Function
Function ByValTest()
    Dim v
    v = 1
    Change v
    ChangeVal v
    ByValTest = v
End Function
Sub Change(ByRef x)
    x = x + 10
End Sub
Sub ChangeVal(ByVal x)
    x = x + 100
End Sub
Function Fact(n)
    If n <= 1 Then Fact = 1 Else Fact = n * Fact(n - 1)
End Function
Function Squares()
    Dim r(1 To 5), i
    For i = 1 To 5: r(i) = i * i: Next
    Squares = r
End Function
