Sub Main()
    Dim ws
    Set ws = Worksheets(1)
    ws.Range("H1:H5").Value = 1
    ws.Range("H3").Value = Empty
    P "end down", ws.Range("H1").End(xlDown).Address
    P "end up", ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    P "end right empty row", ws.Range("H10").End(xlToRight).Column
    P "rows count", ws.Rows.Count
    P "columns count", ws.Columns.Count
    ws.Range("J1:K2").Value = 5
    P "current region", ws.Range("J1").CurrentRegion.Address
    P "offset resize", ws.Range("J1").Offset(1, 1).Resize(2, 3).Address
    P "address relative", ws.Range("B3").Address(False, False)
    P "address row abs", ws.Range("B3").Address(True, False)
    ' (After the workbook's name, which an unsaved workbook gets from its environment.)
    Dim external
    external = ws.Range("B3").Address(External:=True)
    P "address external", Mid(external, InStr(external, "]") + 1)
    P "address r1c1", ws.Range("B3").Address(ReferenceStyle:=xlR1C1)
    P "cells letter", ws.Cells(2, "C").Address
    P "count", ws.Range("A1:C4").Count
    P "rows of range", ws.Range("A1:C4").Rows.Count
    P "intersect", Application.Intersect(ws.Range("A1:C4"), ws.Range("B2:E9")).Address
    P "union", Application.Union(ws.Range("A1"), ws.Range("C3")).Address
    P "item index", ws.Range("A1:C3")(5).Address
    P "entire row", ws.Range("B2").EntireRow.Address
    ws.Range("L1").Value = "=SUM(J1:K2)"
    P "formula", ws.Range("L1").Formula
    P "formula value", ws.Range("L1").Value
    P "formula r1c1", ws.Range("L1").FormulaR1C1
    P "hasformula", ws.Range("L1").HasFormula
    ws.Range("L2").Value = "2024-01-31"
    P "entered date type", TypeName(ws.Range("L2").Value)
    P "entered date value2", ws.Range("L2").Value2
    ws.Range("L3").Value = "'0123"
    P "forced text", ws.Range("L3").Value
    ws.Range("L4").Value = "12%"
    P "percent entry", ws.Range("L4").Value
    P "percent format", ws.Range("L4").NumberFormat
    ws.Range("L5").Value = 0.5
    ws.Range("L5").NumberFormat = "0.0%"
    P "text formatted", ws.Range("L5").Text
    P "find", ws.Range("H1:L5").Find(What:="5", LookAt:=xlWhole).Address
    P "usedrange", ws.UsedRange.Address
    ws.Range("N1:N3").Value = WorksheetFunction.Transpose(Array(3, 1, 2))
    P "transpose", ws.Range("N3").Value
    P "vlookup", WorksheetFunction.VLookup(5, ws.Range("J1:K2"), 2, False)
    P "match", WorksheetFunction.Match(2, ws.Range("N1:N3"), 0)
    P "countif", WorksheetFunction.CountIf(ws.Range("H1:H5"), 1)
    On Error Resume Next
    P "worksheetfunction error", WorksheetFunction.Match(99, ws.Range("N1:N3"), 0)
    P "worksheetfunction error number", Err.Number
    Err.Clear
    P "application function error", TypeName(Application.Match(99, ws.Range("N1:N3"), 0))
    P "evaluate", Evaluate("2*21")
    P "evaluate range", Evaluate("SUM(N1:N3)")
End Sub
