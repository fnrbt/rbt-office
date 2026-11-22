' probe-font-faces.vbs OUT.tsv FACES.txt (run-vbs-in-session.ps1 on the VM; probe-on-vm.sh probe-font-faces.vbs OUT FACES.txt):
' probe-fonts.vbs's measures for faces of one's choosing, a line each in FACES.txt - an Office cloud font Excel draws though GDI
' has it not (Lato), a face the tables lack (Tw Cen MT): "normal" lines, each face as the workbook's Normal font at 6 - 48 pt
' (then a line of excel-normal-fonts.tsv: face in lower case, points, the widest of the ten figures - ColumnWidth 20 less 10,
' a tenth, in px -, StandardHeight in px), and "cell" lines, the row of a cell of each face and style in an Arial 6 sheet
' (AutoFit; then a line of excel-cell-fonts.tsv).
Option Explicit
Dim fso, f, h, x, wb, sh, faces, fn, sz, r, k, b, it, w10, w20
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.OpenTextFile(WScript.Arguments(1))
faces = Replace(f.ReadAll, vbCr, "")
Do While Right(faces, 1) = vbLf: faces = Left(faces, Len(faces) - 1): Loop
faces = Split(faces, vbLf)
f.Close
Set x = CreateObject("Excel.Application"): x.Visible = True: x.DisplayAlerts = False
x.ScreenUpdating = False
Set h = fso.CreateTextFile(WScript.Arguments(0), True)
For Each fn In faces
  Set wb = x.Workbooks.Add: Set sh = wb.Worksheets(1)
  For sz = 6 To 48 Step 0.5
    wb.Styles("Normal").Font.Name = fn
    wb.Styles("Normal").Font.Size = sz
    sh.Columns(1).ColumnWidth = 10: w10 = sh.Columns(1).Width
    sh.Columns(1).ColumnWidth = 20: w20 = sh.Columns(1).Width
    h.WriteLine "normal" & vbTab & LCase(fn) & vbTab & sz & vbTab & Round((w20 - w10) / 0.75 / 10) & vbTab & Round(sh.StandardHeight / 0.75)
  Next
  wb.Close False
  Set wb = x.Workbooks.Add: Set sh = wb.Worksheets(1)
  wb.Styles("Normal").Font.Name = "Arial"
  wb.Styles("Normal").Font.Size = 6
  r = 1
  For sz = 6 To 48 Step 0.5
    For k = 0 To 3
      b = (k = 1 Or k = 3): it = (k = 2 Or k = 3)
      sh.Cells(r, 1).Value = "Xg"
      With sh.Cells(r, 1).Font
        .Name = fn: .Size = sz: .Bold = b: .Italic = it
      End With
      sh.Rows(r).AutoFit
      h.WriteLine "cell" & vbTab & "Arial|6" & vbTab & fn & vbTab & sz & vbTab & k & vbTab & 1 & vbTab & sh.Rows(r).Height
      r = r + 1
    Next
  Next
  wb.Close False
Next
h.Close
x.Quit
