Option Explicit
' Word's pages of every .docx in a folder, for tools/word-parity/page-parity.mjs.
'
'   cscript //nologo page-lines.vbs <dir with .docx> <out dir> [pdf]
'
' Writes <out dir>\<name>.tsv per document:
'   pages <n>                 (ComputeStatistics(wdStatisticPages) after Repaginate)
'   <page> TAB <start> TAB <first line> TAB <last line> TAB <first text line> TAB <last text line>
' where <start> is the character position of the page's first character in the
' main story, the lines are the text of the page's first and last line
' (Selection's HomeKey / EndKey of wdLine; in a table's cell, see LineRange) - tabs,
' paragraph, cell and line marks as spaces - and the text lines the first and last with a letter or number (up
' to 12 lines in from either end; empty when none).
' With "pdf" it also exports <out dir>\<name>.pdf. Read-only, macros off, links
' not updated; the document is closed unsaved.

Const wdGoToPage = 1
Const wdGoToAbsolute = 1
Const wdStatisticPages = 2
Const wdPrintView = 3
Const wdDoNotSaveChanges = 0
Const wdExportFormatPDF = 17

Dim fso, inDir, outDir, wantPdf, w, f
Set fso = CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count < 2 Then
  WScript.Echo "usage: cscript //nologo page-lines.vbs <docx dir> <out dir> [pdf]"
  WScript.Quit 2
End If
inDir = fso.GetAbsolutePathName(WScript.Arguments(0))
outDir = fso.GetAbsolutePathName(WScript.Arguments(1))
wantPdf = WScript.Arguments.Count > 2
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

' A Word of its own for each document: in one session, a document opened after others can be set in
' a stand-in for its embedded fonts (the showcase's exhibition-leaflet came out in Nirmala UI on one
' page in a batch, in Poppins on two alone) - a session's state, not the document's layout.
Sub StartWord()
  Set w = CreateObject("Word.Application")
  w.Visible = False
  w.DisplayAlerts = 0
  w.AutomationSecurity = 3
  w.Options.UpdateLinksAtOpen = False
  On Error Resume Next
  w.Options.Pagination = True
  On Error GoTo 0
End Sub

StartWord
WScript.Echo "Word " & w.Version & " build " & w.Build
w.Quit
For Each f In fso.GetFolder(inDir).Files
  If LCase(fso.GetExtensionName(f.Name)) = "docx" Then
    StartWord
    ProcessFile f.Path
    w.Quit
  End If
Next
WScript.Echo "DONE"

Function Clean(s)
  Dim t
  t = Replace(s, vbTab, " ")
  t = Replace(t, vbCr, " ")
  t = Replace(t, vbLf, " ")
  t = Replace(t, Chr(7), " ")
  t = Replace(t, Chr(11), " ")
  t = Replace(t, Chr(12), " ")
  t = Replace(t, Chr(14), " ")
  Clean = Trim(t)
End Function

' The line at pos. Outside a table: Selection's HomeKey / EndKey of wdLine. In a table's cell those
' go by the row and the paragraph, not the line: the whole of a cell paragraph a page break splits,
' from its start on the page before (docx4j-headers-footers p16 gave "Error! Reference source not
' found." for the "found." atop the page), a cell paragraph's lines all as one, a row's line across its
' cells ("Panel: who pays for the wires?Room B - 50 min"); the predefined bookmark \Line is no better
' there (at a cell's end it starts a character late: "9" of "29"). In a cell the line is the characters
' of pos's paragraph on pos's page at pos's height (Information wdActiveEndPageNumber,
' wdVerticalPositionRelativeToPage): those only go down along the paragraph, so a binary search either
' way finds where the line starts and ends.
Function LineKey(d, i)
  Dim r
  Set r = d.Range(i, i)
  LineKey = r.Information(3) * 100000 + r.Information(6)
End Function

Function LineRange(d, pos)
  Dim sel, para, ps, pe, k, lo, hi, m
  Set sel = d.ActiveWindow.Selection
  sel.SetRange pos, pos
  If sel.Information(12) Then
    Set para = d.Range(pos, pos).Paragraphs(1).Range
    ps = para.Start
    pe = para.End - 1
    If pos > pe Then pos = pe
    If pos < ps Then pos = ps
    k = LineKey(d, pos)
    lo = ps
    hi = pos
    Do While lo < hi
      m = (lo + hi) \ 2
      If LineKey(d, m) < k Then lo = m + 1 Else hi = m
    Loop
    ps = lo
    lo = pos
    hi = pe
    Do While lo < hi
      m = (lo + hi + 1) \ 2
      If LineKey(d, m) > k Then hi = m - 1 Else lo = m
    Loop
    Set LineRange = d.Range(ps, lo + 1)
  Else
    sel.HomeKey 5
    sel.EndKey 5, 1
    Set LineRange = d.Range(sel.Start, sel.End)
  End If
End Function

Function LineAt(d, pos)
  LineAt = Clean(LineRange(d, pos).Text)
End Function

Function HasText(s)
  Dim i, c
  HasText = False
  For i = 1 To Len(s)
    c = AscW(Mid(s, i, 1))
    If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c > 191 Then
      HasText = True
      Exit Function
    End If
  Next
End Function

' The first (dir = 1) or last (dir = -1) line between a and b with text.
Function TextLine(d, a, b, dir)
  Dim sel, i, t, prev, r
  Set sel = d.ActiveWindow.Selection
  TextLine = ""
  If dir = 1 Then sel.SetRange a, a Else sel.SetRange b, b
  For i = 1 To 12
    Set r = LineRange(d, sel.Start)
    If r.Start > b Or r.End < a Then Exit Function
    t = Clean(r.Text)
    If HasText(t) Then
      TextLine = t
      Exit Function
    End If
    prev = r.Start
    sel.SetRange r.Start, r.Start
    If dir = 1 Then sel.MoveDown 5, 1 Else sel.MoveUp 5, 1
    If sel.Start = prev Then Exit Function
  Next
End Function

Sub ProcessFile(path)
  Dim d, name, out, n, p, starts(), lastEnd, a, b, ok
  name = fso.GetBaseName(path)
  On Error Resume Next
  Set d = w.Documents.Open(path, False, True, False, , , , , , , , True, , False)
  If Err.Number <> 0 Then
    WScript.Echo name & " open ERROR " & Err.Number & ": " & Err.Description
    Err.Clear
    Exit Sub
  End If
  d.ActiveWindow.View.Type = wdPrintView
  d.Repaginate
  n = d.ComputeStatistics(wdStatisticPages)
  ' (A line inside a field reads as the whole field - a hyperlink across a page break gave its page's
  ' first line as the whole address. The main story's fields become their results, in this copy only
  ' (opened read-only, never saved); if that moved a page, they are put back.)
  If d.Fields.Count > 0 Then
    d.Fields.Unlink
    d.Repaginate
    If d.ComputeStatistics(wdStatisticPages) <> n Then
      d.Undo
      d.Repaginate
    End If
  End If
  Set out = fso.CreateTextFile(fso.BuildPath(outDir, name & ".tsv"), True, True)
  out.WriteLine "pages" & vbTab & n
  ReDim starts(n + 1)
  For p = 1 To n
    starts(p) = d.GoTo(wdGoToPage, wdGoToAbsolute, p).Start
  Next
  lastEnd = d.Content.End
  For p = 1 To n
    a = starts(p)
    If p < n Then b = starts(p + 1) Else b = lastEnd
    If b > a Then b = b - 1
    out.WriteLine p & vbTab & a & vbTab & LineAt(d, a) & vbTab & LineAt(d, b) & vbTab & TextLine(d, a, b, 1) & vbTab & TextLine(d, a, b, -1)
    If Err.Number <> 0 Then
      out.WriteLine "error" & vbTab & p & vbTab & Err.Number & vbTab & Clean(Err.Description)
      Err.Clear
    End If
  Next
  out.Close
  If wantPdf Then
    d.ExportAsFixedFormat fso.BuildPath(outDir, name & ".pdf"), wdExportFormatPDF, False, 0, 0, 1, 1, 0, True, True, 1, True, True, False
    If Err.Number <> 0 Then
      WScript.Echo name & " pdf ERROR " & Err.Number & ": " & Err.Description
      Err.Clear
    End If
  End If
  d.Close wdDoNotSaveChanges
  Err.Clear
  WScript.Echo name & vbTab & n
  On Error GoTo 0
End Sub
