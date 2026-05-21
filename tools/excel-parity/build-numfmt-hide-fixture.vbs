Option Explicit
Const xlOpenXMLWorkbook = 51
Dim fso,o : Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh : Set x=CreateObject("Excel.Application") : x.Visible=False : x.DisplayAlerts=False
Set wb=x.Workbooks.Add() : Set sh=wb.Worksheets(1) : sh.Name="HideText"
sh.Cells(1,1).Value="value" : sh.Cells(1,2).Value="0.00;;  (hide neg/zero)"
sh.Cells(1,3).Value="0;;  (hide zero)" : sh.Cells(1,4).Value="@  (text)"
Dim vals,i : vals=Array(5, 0, -5, 3.5)
For i=0 To UBound(vals)
  Dim r : r=i+2
  sh.Cells(r,1).Value=vals(i)
  sh.Cells(r,2).Value=vals(i) : sh.Cells(r,2).NumberFormat="0.00;;"
  sh.Cells(r,3).Value=vals(i) : sh.Cells(r,3).NumberFormat="0;;"
  sh.Cells(r,4).Value=vals(i) : sh.Cells(r,4).NumberFormat="@"
Next
sh.Columns("A:D").ColumnWidth=18
wb.SaveAs o, xlOpenXMLWorkbook : wb.Close False : x.Quit
