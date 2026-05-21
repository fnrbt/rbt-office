Option Explicit
Const xlOpenXMLWorkbook=51, xlCellValue=1, xlGreater=5
Dim fso,o: Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh: Set x=CreateObject("Excel.Application"): x.Visible=False: x.DisplayAlerts=False
Set wb=x.Workbooks.Add(): Set sh=wb.Worksheets(1): sh.Name="CFCellRef"
sh.Cells(1,2).Value="Threshold:": sh.Cells(1,3).Value=10
Dim vals,i: vals=Array(5,12,8,20,3)
For i=0 To UBound(vals)
  sh.Cells(i+2,1).Value=vals(i)
Next
Dim rng,fc: Set rng=sh.Range("A2:A6")
Set fc=rng.FormatConditions.Add(xlCellValue, xlGreater, "=$C$1")
fc.Interior.Color=RGB(255,199,206)
fc.Font.Color=RGB(156,0,6)
sh.Columns("A:C").ColumnWidth=14
wb.SaveAs o, xlOpenXMLWorkbook: wb.Close False: x.Quit
