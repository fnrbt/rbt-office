' cscript probe-print-pages.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel 2019 and, for every visible worksheet,
' exports it on its own to PDF (Worksheet.ExportAsFixedFormat, as print-excel-pdf.vbs does the workbook) as
' OUT-<n>.pdf next to OUT.tsv (n the sheet's 1-based index among the workbook's worksheets) and writes its page setup
' and its page breaks (Page Break Preview: automatic and manual) to OUT.tsv, a line per fact:
'   sheet <tab> n <tab> key <tab> value
' keys: setup.* (PageSetup's properties), hbreak / vbreak (row / column before which a page starts, and its type:
' manual or auto), pdf (the PDF's name, or "error ..."). Run it in the desktop session (excel-print-pages.sh).
Option Explicit
Dim fso, x, wb, s, out, base, n, i, b, ps, pdf, t
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application")
x.Visible = True : x.DisplayAlerts = False : x.AutomationSecurity = 3 : x.ScreenUpdating = True
x.AskToUpdateLinks = False
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
base = fso.BuildPath(fso.GetParentFolderName(WScript.Arguments(1)), fso.GetBaseName(WScript.Arguments(1)))

Sub W(key, value)
  out.WriteLine s.Name & vbTab & n & vbTab & key & vbTab & value
End Sub

Function V(f)
  On Error Resume Next
  V = "?"
  V = CStr(f)
End Function

n = 0
For Each s In wb.Worksheets
  n = n + 1
  If s.Visible = -1 Then
    s.Activate
    Set ps = s.PageSetup
    On Error Resume Next
    W "setup.orientation", ps.Orientation
    W "setup.paper", ps.PaperSize
    W "setup.zoom", V(ps.Zoom)
    W "setup.fitwide", V(ps.FitToPagesWide)
    W "setup.fittall", V(ps.FitToPagesTall)
    W "setup.margins", ps.LeftMargin & " " & ps.RightMargin & " " & ps.TopMargin & " " & ps.BottomMargin & " " & ps.HeaderMargin & " " & ps.FooterMargin
    W "setup.center", ps.CenterHorizontally & " " & ps.CenterVertically
    W "setup.printarea", ps.PrintArea
    W "setup.titlerows", ps.PrintTitleRows
    W "setup.titlecols", ps.PrintTitleColumns
    W "setup.order", ps.Order
    W "setup.firstpage", ps.FirstPageNumber
    W "setup.gridlines", ps.PrintGridlines
    W "setup.headings", ps.PrintHeadings
    W "setup.bw", ps.BlackAndWhite
    W "setup.draft", ps.Draft
    W "setup.errors", ps.PrintErrors
    W "setup.header", ps.LeftHeader & "|" & ps.CenterHeader & "|" & ps.RightHeader
    W "setup.footer", ps.LeftFooter & "|" & ps.CenterFooter & "|" & ps.RightFooter
    W "used", s.UsedRange.Address
    x.ActiveWindow.View = 2
    For i = 1 To s.HPageBreaks.Count
      Set b = s.HPageBreaks(i)
      t = "auto" : If b.Type = -4135 Then t = "manual"
      W "hbreak", b.Location.Row & " " & t
    Next
    For i = 1 To s.VPageBreaks.Count
      Set b = s.VPageBreaks(i)
      t = "auto" : If b.Type = -4135 Then t = "manual"
      W "vbreak", b.Location.Column & " " & t
    Next
    x.ActiveWindow.View = 1
    Err.Clear
    pdf = base & "-" & n & ".pdf"
    s.ExportAsFixedFormat 0, pdf, 0, True, False, , , False
    If Err.Number <> 0 Then
      W "pdf", "error " & Err.Description
    Else
      W "pdf", fso.GetFileName(pdf)
    End If
    Err.Clear
    On Error GoTo 0
  End If
Next
out.Close
wb.Close False
x.Quit
