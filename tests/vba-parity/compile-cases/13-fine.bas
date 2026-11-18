Option Explicit
Function F() As Range
End Function
Sub Test()
    Dim x
    Dim r As Range
    x = r.Offset(ColumnOffset:=1).Value + r.Cells(1, 1).Value
End Sub
