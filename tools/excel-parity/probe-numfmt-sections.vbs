' probe-numfmt-sections.vbs
'
' Sweeps Excel number-format SECTION selection across (format, value) pairs:
' which section Excel applies for conditional (`[>=100]...`) and sign-based
' formats, plus the colour it shows. Reads cell.Text and the effective font
' colour (DisplayFormat.Font.Color) so we can match Excel's section-picking
' and colour rules.
'
' Usage: cscript //nologo probe-numfmt-sections.vbs <unused-input> <out.tsv>

Option Explicit
Dim fso, outputPath
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: probe-numfmt-sections.vbs <input> <out.tsv>"
  WScript.Quit 2
End If
outputPath = fso.GetAbsolutePathName(WScript.Arguments(1))

Dim excel, wb, sh, outFile
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
Set wb = excel.Workbooks.Add()
Set sh = wb.Worksheets(1)
Set outFile = fso.CreateTextFile(outputPath, True, True)
outFile.WriteLine "numFmt" & vbTab & "value" & vbTab & "text" & vbTab & "fontColorBGR"

' Formats exercising conditional sections, sign sections, and colours.
Dim fmts, fi
fmts = Array( _
  "[>=100]""High"";[>=50]""Mid"";""Low""", _
  "[>=100]0;[>=50]0;0", _
  "[<0][Red]0.0;[Blue]0.0", _
  "0.0;[Red]-0.0;""zero""", _
  "[Green]0;[Red]\-0;0", _
  "[>5]0""+"";0""-""", _
  "[Red][>100]0;0", _
  "#,##0.00;[Red](#,##0.00)" )

Dim vals, vi
vals = Array(-10, 0, 5, 30, 50, 70, 99, 100, 150)

Dim cell
Set cell = sh.Range("A1")
For fi = 0 To UBound(fmts)
  cell.NumberFormat = fmts(fi)
  For vi = 0 To UBound(vals)
    cell.Value = vals(vi)
    Dim shown, col
    shown = cell.Text
    On Error Resume Next
    col = cell.DisplayFormat.Font.Color
    If Err.Number <> 0 Then col = cell.Font.Color
    On Error GoTo 0
    outFile.WriteLine fmts(fi) & vbTab & CStr(vals(vi)) & vbTab & shown & vbTab & CStr(col)
  Next
Next

outFile.Close
wb.Close False
excel.Quit
Set excel = Nothing
