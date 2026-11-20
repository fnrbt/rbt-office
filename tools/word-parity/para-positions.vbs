' Word's vertical position of every paragraph of the main story, and of every line of every
' footnote, for checking line-height and footnote rules against fixtures (tools/word-parity/fixtures).
'
'   cscript //nologo para-positions.vbs <dir with .docx> <out dir>
'
' Writes <out dir>\<name>.tsv:
'   <paragraph index> TAB <page> TAB <top of its first line, points from the page's top> TAB <its text>
' then for each footnote <i>:
'   ref<i> TAB <page> TAB <top of the line of its reference mark>
'   fn<i> TAB <page> TAB <top of a line of the note> TAB <the line's first word>   (one per line, as
'      the note's words change page or top)
' Information(6) is in whole pixels at 96 dpi (multiples of 0.75pt). Read-only, macros off; the
' document is closed unsaved.
Option Explicit
Dim fso, word, folder, file, doc, outDir, out, i, p, r, n, w, lastPage, lastTop, pg, tp
Set fso = CreateObject("Scripting.FileSystemObject")
' (A Word of its own for each document: see page-lines.vbs - a session can lose a document's embedded fonts.)
outDir = WScript.Arguments(1)
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
Set folder = fso.GetFolder(WScript.Arguments(0))
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "docx" Then
        Set word = CreateObject("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        word.AutomationSecurity = 3
        Set doc = word.Documents.Open(file.Path, False, True, False)
        doc.ActiveWindow.View.Type = 3
        doc.Repaginate
        Set out = fso.CreateTextFile(fso.BuildPath(outDir, fso.GetBaseName(file.Name) & ".tsv"), True, True)
        i = 0
        For Each p In doc.Paragraphs
            i = i + 1
            Set r = p.Range
            r.Collapse 1
            out.WriteLine i & vbTab & r.Information(3) & vbTab & r.Information(6) & vbTab & Replace(Replace(Left(p.Range.Text, 60), vbCr, ""), vbTab, " ")
        Next
        For n = 1 To doc.Footnotes.Count
            Set r = doc.Footnotes(n).Reference
            out.WriteLine "ref" & n & vbTab & r.Information(3) & vbTab & r.Information(6)
            lastPage = -1
            lastTop = -1
            For Each w In doc.Footnotes(n).Range.Words
                pg = w.Information(3)
                tp = w.Information(6)
                If pg <> lastPage Or tp <> lastTop Then
                    out.WriteLine "fn" & n & vbTab & pg & vbTab & tp & vbTab & Replace(Replace(w.Text, vbCr, ""), vbTab, " ")
                    lastPage = pg
                    lastTop = tp
                End If
            Next
        Next
        out.Close
        doc.Close 0
        word.Quit
    End If
Next
