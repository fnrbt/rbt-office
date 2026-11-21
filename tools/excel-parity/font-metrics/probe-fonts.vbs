' probe-fonts.vbs OUTDIR (run-vbs-in-session.ps1 on the VM) : row heights of cells of many fonts in sheets of a few Normal fonts (AutoFit), the
' Normal-font sweep at fractional sizes (StandardHeight, column widths), and what Font.Size keeps of a fraction.
Option Explicit
Dim fso, f, g, h, x, wb, sh, faces, sizes, normals, fn, sz, n, r, b, it, lines, out, k, v
out = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application"): x.Visible = True: x.DisplayAlerts = False
x.ScreenUpdating = False
faces = Array("Garamond", "Century Gothic", "Trebuchet MS", "Arial Narrow", "Consolas", "Candara", "Corbel", "Book Antiqua", "Lucida Sans", "Franklin Gothic Book", "Gill Sans MT", "Palatino Linotype", "Constantia", "Roboto", "Myriad Pro", "Segoe UI Black", "MS Sans Serif", "Aptos", "Aptos Narrow", "Aptos Display", "Calibri", "Arial", "Times New Roman", "Cambria", "Verdana", "Tahoma", "Segoe UI", "Courier New", "Georgia", "Calibri Light")
sizes = Array(6, 7, 7.5, 8, 8.5, 9, 9.5, 10, 10.5, 11, 11.5, 12, 12.5, 13, 13.5, 14, 15, 16, 16.5, 17, 18, 19, 20, 21, 22, 24, 26, 28, 32, 36)
normals = Array("Arial|6", "Calibri|11", "Arial|10", "Tahoma|10")

' 1. Font.Size of a fraction
Set g = fso.CreateTextFile(out & "\fontsize.tsv", True)
Set wb = x.Workbooks.Add: Set sh = wb.Worksheets(1)
For Each v In Array(10.3, 10.25, 10.2, 10.75, 10.6, 10.5, 10.1, 10.05, 10.01, 8.33, 409, 1, 1.2)
  sh.Cells(1, 1).Font.Size = v
  g.WriteLine v & vbTab & sh.Cells(1, 1).Font.Size
Next
g.Close
wb.Close False

' 2. Cells of each font in sheets of each Normal font: a new workbook a face (a workbook holds 512 fonts)
Set f = fso.CreateTextFile(out & "\cellfonts.tsv", True)
For Each n In normals
  For Each fn In faces
  Set wb = x.Workbooks.Add: Set sh = wb.Worksheets(1)
  wb.Styles("Normal").Font.Name = Split(n, "|")(0)
  wb.Styles("Normal").Font.Size = CDbl(Split(n, "|")(1))
  sh.Columns(1).ColumnWidth = 40
  r = 1
    For Each sz In sizes
      For k = 0 To 3
        b = (k = 1 Or k = 3): it = (k = 2 Or k = 3)
        If Not (Left(fn, 5) = "Aptos" And k > 0) Then
          For lines = 1 To 2
            If lines = 1 Then
              sh.Cells(r, 1).Value = "Xg"
            Else
              sh.Cells(r, 1).WrapText = True
              sh.Cells(r, 1).Value = "Xg" & vbLf & "Xg"
            End If
            On Error Resume Next
            Err.Clear
            With sh.Cells(r, 1).Font
              .Name = fn: .Size = sz: .Bold = b: .Italic = it
            End With
            If Err.Number = 0 Then
              sh.Rows(r).AutoFit
              f.WriteLine n & vbTab & fn & vbTab & sz & vbTab & k & vbTab & lines & vbTab & sh.Rows(r).Height
            Else
              f.WriteLine n & vbTab & fn & vbTab & sz & vbTab & k & vbTab & lines & vbTab & "ERR" & vbTab & Err.Description
            End If
            Err.Clear
            On Error GoTo 0
            r = r + 1
          Next
        End If
      Next
    Next
  wb.Close False
  Next
Next
f.Close


' 3. Normal fonts at fractional sizes
Set h = fso.CreateTextFile(out & "\normals.tsv", True)
Set wb = x.Workbooks.Add: Set sh = wb.Worksheets(1)
For Each fn In faces
  For Each sz In Array(7.5, 8.5, 9.5, 10.5, 11.5, 12.5, 13.5, 14.5, 15, 17, 22)
    wb.Styles("Normal").Font.Name = fn
    wb.Styles("Normal").Font.Size = sz
    sh.Columns(1).ColumnWidth = 10: v = sh.Columns(1).Width
    sh.Columns(1).ColumnWidth = 20
    h.WriteLine fn & vbTab & sz & vbTab & v & vbTab & sh.Columns(1).Width & vbTab & sh.StandardWidth & vbTab & sh.Columns(5).Width & vbTab & sh.StandardHeight
  Next
Next
h.Close
wb.Close False
x.Quit
