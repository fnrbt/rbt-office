' Members of Excel's objects reached late-bound (As Object / Variant): what an unknown member, a
' read-only property written, the wrong number of arguments or a wrong argument name raise (number and
' Err.Description), CallByName, default members. (Early-bound mistakes are compile errors in Excel:
' they stop the module before it runs, so they are not in a case.)
' Err is read into n / d before P: a procedure with an On Error statement clears Err when it runs one.
Sub Main()
    On Error Resume Next
    Dim o As Object, w As Object, a As Object, x, n, d
    Set o = Range("H20")
    Set w = ActiveSheet
    Set a = Application
    o.Value = 7

    x = o.Nope: n = Err.Number: d = Err.Description: Err.Clear: P "read unknown", n: P "read unknown text", d
    o.Nope = 1: n = Err.Number: d = Err.Description: Err.Clear: P "write unknown", n: P "write unknown text", d
    o.Nope 1, 2: n = Err.Number: d = Err.Description: Err.Clear: P "call unknown", n: P "call unknown text", d
    Set o.Nope = Nothing: n = Err.Number: Err.Clear: P "set unknown", n
    x = w.Nope: n = Err.Number: d = Err.Description: Err.Clear: P "sheet unknown", n: P "sheet unknown text", d
    x = a.Nope: n = Err.Number: Err.Clear: P "application unknown", n
    x = a.NopeFunction(1, 2): n = Err.Number: Err.Clear: P "application unknown call", n

    o.Row = 5: n = Err.Number: d = Err.Description: Err.Clear: P "write read-only", n: P "write read-only text", d
    w.Index = 3: n = Err.Number: Err.Clear: P "write read-only index", n
    Set w.UsedRange = o: n = Err.Number: d = Err.Description: Err.Clear: P "set read-only object", n: P "set read-only object text", d
    x = o.Offset(1, 2, 3): n = Err.Number: d = Err.Description: Err.Clear: P "too many arguments", n: P "too many arguments text", d
    x = o.Offset(Bogus:=1).Address: n = Err.Number: d = Err.Description: Err.Clear: P "unknown named argument", n: P "unknown named argument text", d
    x = o.Offset(ColumnOffset:=2).Address: Err.Clear: P "named argument", x
    x = o.Resize(): n = Err.Number: Err.Clear: P "method without its arguments", n

    x = CallByName(o, "Value", VbGet): Err.Clear: P "callbyname get", x
    CallByName o, "Value", VbLet, 9: Err.Clear: P "callbyname let", o.Value
    x = CallByName(o, "Offset", VbGet, 1, 1).Address: n = Err.Number: Err.Clear: P "callbyname property with arguments", x: P "callbyname property with arguments error", n
    x = CallByName(o, "Offset", VbMethod, 1, 1).Address: n = Err.Number: Err.Clear: P "callbyname property as method", n
    CallByName o, "ClearContents", VbMethod: n = Err.Number: Err.Clear: P "callbyname method", n: P "callbyname method cleared", IsEmpty(o.Value)
    x = CallByName(o, "Nope", VbGet): n = Err.Number: d = Err.Description: Err.Clear: P "callbyname unknown", n: P "callbyname unknown text", d
    x = CallByName(o, "Row", VbLet, 3): n = Err.Number: Err.Clear: P "callbyname read-only", n

    o.Value = 7
    P "default member", CStr(o)
    P "typename", TypeName(o) & " " & TypeName(w) & " " & TypeName(a) & " " & TypeName(o.Font) & " " & TypeName(o.Cells)
    P "property with arguments", o.Cells(2, 2).Address
    P "count read as call", o.Count()
    x = o.Value(1, 2, 3): n = Err.Number: Err.Clear: P "value with arguments", n
    x = Range("ZZZ999999").Value: n = Err.Number: d = Err.Description: Err.Clear: P "bad range", n: P "bad range text", d
    x = Worksheets("NoSuchSheet").Name: n = Err.Number: d = Err.Description: Err.Clear: P "bad sheet", n: P "bad sheet text", d
    x = a.Sum(1, 2): Err.Clear: P "application function", x
    x = a.VLookup(5, Array(1, 2), 2, False): Err.Clear: P "application function error", IsError(x)
    x = a.WorksheetFunction.VLookup(5, Array(1, 2), 2, False): n = Err.Number: d = Err.Description: Err.Clear: P "worksheetfunction error", n: P "worksheetfunction error text", d
End Sub
