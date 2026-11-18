Sub Main()
    Dim s, names
    Worksheets.Add After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "Second"
    Worksheets.Add Before:=Worksheets(1)
    ActiveSheet.Name = "First"
    For Each s In Worksheets
        names = names & s.Name & ","
    Next
    Worksheets("Sheet1").Activate
    P "order", names
    P "count", Worksheets.Count
    P "index", Worksheets("Second").Index
    Worksheets("Second").Range("A1").Value = 42
    P "other sheet", Worksheets("Second").Range("A1").Value
    P "cross sheet formula", Evaluate("Second!A1*2")
    Worksheets("Sheet1").Range("Z1").Formula = "=Second!A1+1"
    P "cross formula value", Worksheets("Sheet1").Range("Z1").Value
    ActiveWorkbook.Names.Add Name:="Answer", RefersTo:="=Second!$A$1"
    P "name value", Evaluate("Answer")
    P "name refers", ActiveWorkbook.Names("Answer").RefersTo
    On Error Resume Next
    Worksheets("Second").Name = "Sheet1"
    P "duplicate name err", Err.Number
    Err.Clear
    Worksheets("First").Delete
    P "after delete", Worksheets.Count
    P "next sheet", Worksheets("Sheet1").Next.Name
End Sub
