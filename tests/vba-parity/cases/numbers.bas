Sub Main()
    P "int div", 7 \ 2
    P "int div negative", -7 \ 2
    P "mod", 7 Mod 3
    P "mod negative", -7 Mod 3
    P "mod double", 7.5 Mod 2
    P "pow", 2 ^ 10
    P "pow type", TypeName(2 ^ 2)
    P "round half even", Round(2.5)
    P "round 3.5", Round(3.5)
    P "round digits", Round(1.005, 2)
    P "cint 2.5", CInt(2.5)
    P "cint 3.5", CInt(3.5)
    P "clng -2.5", CLng(-2.5)
    P "int negative", Int(-2.5)
    P "fix negative", Fix(-2.5)
    P "val", Val("  12 34abc")
    P "val hex", Val("&HFF")
    P "val dot", Val("1.5e2x")
    P "cdbl", CDbl("3.25")
    P "hex", Hex(255)
    P "hex negative", Hex(-1)
    P "oct", Oct(8)
    P "sgn", Sgn(-3)
    P "abs", Abs(-2.5)
    P "sqr", Sqr(2)
    P "precedence", 2 + 3 * 4 ^ 2 / 8 - 1
    P "unary", -2 ^ 2
    P "integer literal type", TypeName(5)
    P "long literal type", TypeName(70000)
    P "double literal type", TypeName(1.5)
    P "integer sum type", TypeName(30000 + 30000)
    P "and bitwise", 12 And 10
    P "or bitwise", 12 Or 3
    P "xor", 5 Xor 3
    P "not", Not 0
    P "true value", CInt(True)
    P "cbool", CBool("True")
    P "cstr double", CStr(0.1 + 0.2)
    P "cstr big", CStr(1E+16)
    P "cstr small", CStr(0.000001)
    P "large mult", 123456789 * 1000
    Dim i As Integer
    On Error Resume Next
    i = 40000
    P "integer overflow err", Err.Number
    Err.Clear
    Dim x
    x = 1 / 0
    P "div zero err", Err.Number
    Err.Clear
    x = 5 \ 0
    P "int div zero err", Err.Number
End Sub
