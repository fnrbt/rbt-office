Option Explicit
Const xlOpenXMLWorkbook=51
Const xlColumnClustered=51, xlBarClustered=57, xlLineMarkers=65, xlPie=5, xlArea=1, xlXYScatter=-4169
Dim fso,o: Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh: Set x=CreateObject("Excel.Application"): x.Visible=False: x.DisplayAlerts=False
Set wb=x.Workbooks.Add(): Set sh=wb.Worksheets(1): sh.Name="Zoo"
' category data
sh.Cells(1,1).Value="Quarter": sh.Cells(1,2).Value="Sales"
Dim cats,vals,i: cats=Array("Q1","Q2","Q3","Q4"): vals=Array(10,25,15,30)
For i=0 To 3
  sh.Cells(i+2,1).Value=cats(i): sh.Cells(i+2,2).Value=vals(i)
Next
' scatter data (X,Y)
sh.Cells(1,4).Value="X": sh.Cells(1,5).Value="Y"
Dim xs,ys: xs=Array(1,2,3,4,5): ys=Array(3,8,5,12,9)
For i=0 To 4
  sh.Cells(i+2,4).Value=xs(i): sh.Cells(i+2,5).Value=ys(i)
Next

Sub MakeChart(t, l, tp, titleText, rng)
  Dim co,ch
  Set co=sh.Shapes.AddChart2(-1, t, l, tp, 300, 200)
  Set ch=co.Chart
  ch.SetSourceData sh.Range(rng)
  ch.HasTitle=True: ch.ChartTitle.Text=titleText
  ch.HasLegend=False
End Sub

MakeChart xlColumnClustered, 220, 10, "Column", "A1:B5"
MakeChart xlBarClustered, 540, 10, "Bar", "A1:B5"
MakeChart xlLineMarkers, 220, 220, "Line", "A1:B5"
MakeChart xlPie, 540, 220, "Pie", "A1:B5"
MakeChart xlArea, 220, 430, "Area", "A1:B5"
MakeChart xlXYScatter, 540, 430, "Scatter", "D1:E6"

wb.SaveAs o, xlOpenXMLWorkbook: wb.Close False: x.Quit
