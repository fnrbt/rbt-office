Option Explicit
Const xlOpenXMLWorkbook = 51
Dim fso,o : Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh : Set x=CreateObject("Excel.Application") : x.Visible=False : x.DisplayAlerts=False
Set wb=x.Workbooks.Add() : Set sh=wb.Worksheets(1) : sh.Name="Scale"
sh.Cells(1,1).Value="value" : sh.Cells(1,2).Value="#,##0," : sh.Cells(1,3).Value="#,##0.0,"
sh.Cells(1,4).Value="0.0,," : sh.Cells(1,5).Value="#,##0.0,,""M"""
Dim vals,i : vals=Array(1234567,2500000,12345678)
For i=0 To UBound(vals)
  Dim r : r=i+2
  sh.Cells(r,1).Value=vals(i)
  sh.Cells(r,2).Value=vals(i) : sh.Cells(r,2).NumberFormat="#,##0,"
  sh.Cells(r,3).Value=vals(i) : sh.Cells(r,3).NumberFormat="#,##0.0,"
  sh.Cells(r,4).Value=vals(i) : sh.Cells(r,4).NumberFormat="0.0,,"
  sh.Cells(r,5).Value=vals(i) : sh.Cells(r,5).NumberFormat="#,##0.0,,""M"""
Next
sh.Columns("A:E").ColumnWidth=14
wb.SaveAs o, xlOpenXMLWorkbook : wb.Close False : x.Quit
