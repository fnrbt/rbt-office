Option Explicit
' Create public, synthetic controls for horizontal page coverage. Each sheet has
' identical uniquely labelled cells; only its print setup differs.
Dim fso, output, xl, wb, sh, i, c, r, widths
Set fso = CreateObject("Scripting.FileSystemObject")
output = fso.GetAbsolutePathName(WScript.Arguments(1))
Set xl = CreateObject("Excel.Application")
xl.Visible = False
xl.DisplayAlerts = False
Set wb = xl.Workbooks.Add()
Do While wb.Worksheets.Count < 3
  wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count)
Loop
Do While wb.Worksheets.Count > 3
  wb.Worksheets(wb.Worksheets.Count).Delete
Loop
widths = Array(8, 13, 18, 10, 16, 8, 14, 19, 11, 15, 9, 17)
For i = 1 To 3
  Set sh = wb.Worksheets(i)
  Select Case i
    Case 1: sh.Name = "Zoom"
    Case 2: sh.Name = "FitWide"
    Case 3: sh.Name = "PrintArea"
  End Select
  For c = 1 To 30
    sh.Columns(c).ColumnWidth = widths((c - 1) Mod 12)
    For r = 1 To 3
      sh.Cells(r, c).Value = "C" & Right("00" & c, 2) & "R" & r
    Next
  Next
  With sh.PageSetup
    .Orientation = 1
    .PaperSize = 1
    .LeftMargin = 50.4
    .RightMargin = 50.4
    .TopMargin = 54
    .BottomMargin = 54
    If i = 2 Then
      .Zoom = False
      .FitToPagesWide = 1
      .FitToPagesTall = False
    Else
      .Zoom = 100
    End If
    If i = 3 Then .PrintArea = "$A$1:$T$3"
  End With
Next
wb.SaveAs output, 51
wb.Close False
xl.Quit
