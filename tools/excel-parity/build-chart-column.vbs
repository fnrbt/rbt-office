Option Explicit
Const xlOpenXMLWorkbook=51, xlColumnClustered=51
Dim fso,o: Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh: Set x=CreateObject("Excel.Application"): x.Visible=False: x.DisplayAlerts=False
Set wb=x.Workbooks.Add(): Set sh=wb.Worksheets(1): sh.Name="ChartCol"
sh.Cells(1,1).Value="Quarter": sh.Cells(1,2).Value="Sales"
Dim cats,vals,i: cats=Array("Q1","Q2","Q3","Q4"): vals=Array(10,25,15,30)
For i=0 To 3
  sh.Cells(i+2,1).Value=cats(i)
  sh.Cells(i+2,2).Value=vals(i)
Next
Dim co,ch
Set co=sh.Shapes.AddChart2(-1, xlColumnClustered, 200, 10, 360, 240)
Set ch=co.Chart
ch.SetSourceData sh.Range("A1:B5")
ch.HasTitle=True: ch.ChartTitle.Text="Quarterly Sales"
ch.HasLegend=False
wb.SaveAs o, xlOpenXMLWorkbook: wb.Close False: x.Quit
