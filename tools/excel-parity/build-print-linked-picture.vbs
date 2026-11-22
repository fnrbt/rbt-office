Option Explicit
' Synthetic camera-style linked picture crossing a vertical print page boundary.
Dim xl, wb, source, report, pic, fso, output, i, mode, box
Set fso = CreateObject("Scripting.FileSystemObject")
output = fso.GetAbsolutePathName(WScript.Arguments(1))
' Output names containing "under" or "over" select held-out z-order variants.
mode = LCase(fso.GetFileName(output))
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add()
Do While wb.Worksheets.Count < 2
  wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count)
Loop
Do While wb.Worksheets.Count > 2
  wb.Worksheets(wb.Worksheets.Count).Delete
Loop
Set source = wb.Worksheets(1)
Set report = wb.Worksheets(2)
source.Name = "Source"
report.Name = "Report"
source.Columns("A:B").ColumnWidth = 16
source.Range("A1").Value = "LEFTTOP"
source.Range("B1").Value = "RIGHTTOP"
source.Range("A2").Value = "LEFTBOTTOM"
source.Range("B2").Value = "RIGHTBOTTOM"
source.Range("A1:B2").Font.Size = 14
report.Columns("A:J").ColumnWidth = 10
report.Range("A1").Value = "FIRSTPAGE"
report.Range("J1").Value = "SECONDPAGE"
If InStr(mode, "under") > 0 Then
  Set box = report.Shapes.AddShape(1, report.Range("F3").Left + 90, report.Range("F3").Top, 110, 48)
  box.Fill.ForeColor.RGB = 255
  box.Line.Visible = False
End If
source.Range("A1:B2").Copy
report.Activate
Set pic = report.Pictures.Paste(True)
pic.Left = report.Range("F3").Left
pic.Top = report.Range("F3").Top
If InStr(mode, "over") > 0 Then
  Set box = report.Shapes.AddShape(1, report.Range("F3").Left + 90, report.Range("F3").Top, 110, 48)
  box.Fill.ForeColor.RGB = 255
  box.Line.Visible = False
End If
source.Visible = 0
With report.PageSetup
  .Orientation = 1
  .PaperSize = 1
  .LeftMargin = 50.4
  .RightMargin = 50.4
  .TopMargin = 54
  .BottomMargin = 54
  .Zoom = 100
  .PrintArea = "$A$1:$J$8"
End With
report.VPageBreaks.Add report.Range("G1")
wb.SaveAs output, 51
wb.Close False
xl.Quit
