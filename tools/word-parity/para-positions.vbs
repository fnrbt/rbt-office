' Word's vertical position of every paragraph of the main story, for checking line-height rules
' against fixtures (tools/word-parity/fixtures).
'
'   cscript //nologo para-positions.vbs <dir with .docx> <out dir>
'
' Writes <out dir>\<name>.tsv: <paragraph index> TAB <page> TAB <top of its first line, points from
' the page's top> TAB <its text>. Read-only, macros off; the document is closed unsaved.
Option Explicit
Dim fso, word, folder, file, doc, outDir, out, i, p, r
Set fso = CreateObject("Scripting.FileSystemObject")
Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
word.AutomationSecurity = 3
outDir = WScript.Arguments(1)
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir
Set folder = fso.GetFolder(WScript.Arguments(0))
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Name)) = "docx" Then
        Set doc = word.Documents.Open(file.Path, False, True, False)
        doc.Repaginate
        Set out = fso.CreateTextFile(fso.BuildPath(outDir, fso.GetBaseName(file.Name) & ".tsv"), True, True)
        i = 0
        For Each p In doc.Paragraphs
            i = i + 1
            Set r = p.Range
            r.Collapse 1
            out.WriteLine i & vbTab & r.Information(3) & vbTab & r.Information(6) & vbTab & Replace(Replace(Left(p.Range.Text, 60), vbCr, ""), vbTab, " ")
        Next
        out.Close
        doc.Close 0
    End If
Next
word.Quit
