' cscript export-ranges.vbs LIST: for each line of LIST (tab separated: workbook path, sheet, range, output),
' Excel's own picture of the range with the objects over it. An output ending in .png is the range as on
' screen (CopyPicture xlScreen, xlBitmap, pasted into an empty chart of the same size and exported: 96 dpi,
' one pixel a screen pixel at 100%); one ending in .pdf is the range printed (its print area, 100%, no
' gridlines or headings). Works from an SSH session (no desktop needed for the PDF; the PNG needs the
' clipboard). Writes "ok" / "ERROR" lines to stdout.
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application")
x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3 : x.AskToUpdateLinks = False
Set list = fso.OpenTextFile(WScript.Arguments(0), 1)
Do Until list.AtEndOfStream
  line = list.ReadLine
  If Len(Trim(line)) > 0 Then
    parts = Split(line, vbTab)
    On Error Resume Next
    Err.Clear
    Set wb = x.Workbooks.Open(parts(0), 0, True)
    If Err.Number <> 0 Then
      WScript.Echo "ERROR" & vbTab & parts(3) & vbTab & "open: " & Err.Description
    Else
      Set ws = wb.Worksheets(parts(1))
      Set r = ws.Range(parts(2))
      If LCase(Right(parts(3), 4)) = ".pdf" Then
        ws.PageSetup.PrintArea = r.Address
        ws.PageSetup.Zoom = 100
        ws.PageSetup.PrintGridlines = False
        ws.PageSetup.PrintHeadings = False
        ws.PageSetup.LeftMargin = 0 : ws.PageSetup.RightMargin = 0 : ws.PageSetup.TopMargin = 0 : ws.PageSetup.BottomMargin = 0
        ws.PageSetup.HeaderMargin = 0 : ws.PageSetup.FooterMargin = 0
        ws.PageSetup.CenterHeader = "" : ws.PageSetup.CenterFooter = ""
        ws.ExportAsFixedFormat 0, parts(3), 0, True, False
      Else
        ws.Activate
        x.ActiveWindow.Zoom = 100
        r.CopyPicture 1, 2
        Set holder = ws.ChartObjects.Add(0, 0, r.Width, r.Height)
        holder.Activate
        holder.Chart.ChartArea.Format.Line.Visible = False
        holder.Chart.Paste
        holder.Chart.Export parts(3), "PNG"
        holder.Delete
      End If
      If Err.Number <> 0 Then
        WScript.Echo "ERROR" & vbTab & parts(3) & vbTab & Err.Description
      Else
        WScript.Echo "ok" & vbTab & parts(3) & vbTab & r.Width & " x " & r.Height & " pt"
      End If
      wb.Close False
    End If
    On Error Goto 0
  End If
Loop
x.Quit
