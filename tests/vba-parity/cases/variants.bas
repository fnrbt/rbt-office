Sub Main()
    Dim e, n
    n = Null
    P "typename empty", TypeName(e)
    P "typename null", TypeName(n)
    P "vartype empty", VarType(e)
    P "vartype string", VarType("x")
    P "vartype array", VarType(Array())
    P "typename nothing", TypeName(Nothing)
    P "typename range", TypeName(Range("A1"))
    P "typename sheet", TypeName(ActiveSheet)
    P "typename app", TypeName(Application)
    P "isempty", IsEmpty(e)
    P "isnull", IsNull(n)
    P "null plus", IsNull(n + 1)
    P "null concat", "a" & n & "b"
    P "empty plus", e + 1
    P "empty concat", "[" & e & "]"
    P "empty equals zero", e = 0
    P "empty equals string", e = ""
    P "isnumeric string", IsNumeric("12.5")
    P "isnumeric blank", IsNumeric("")
    P "isnumeric empty", IsNumeric(e)
    P "isnumeric hex", IsNumeric("&H10")
    P "isnumeric spaces", IsNumeric(" 5 ")
    P "isobject", IsObject(Range("A1"))
    P "is nothing", Range("A1") Is Nothing
    P "cvar", TypeName(CVar(5))
    P "cvErr", TypeName(CVErr(2042))
    P "iserror", IsError(CVErr(2007))
    P "currency type", TypeName(CCur(1.5))
    P "currency math", CCur(0.1) * 3
    P "byte", CByte(255)
    P "single", TypeName(CSng(1.5))
    P "string to date compare", "2" < 10
End Sub
