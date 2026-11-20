Option Explicit
' What Word shows in the content controls of every .docx in a folder, and what it saves
' (tools/word-parity/fixtures/data-binding.py).
'
'   cscript //nologo content-controls.vbs <dir with .docx> <out dir>
'
' Writes, per document <name>, in <out dir>:
'   <name>.tsv          (UTF-8) one row per paragraph of every story:
'                         para TAB <story type> TAB <n> TAB <text>
'                       and one per content control:
'                         cc TAB <story type> TAB <n> TAB <type> TAB <showing placeholder> TAB <mapped> TAB <xpath>
'                            TAB <checked> TAB <bold> TAB <italic> TAB <color> TAB <font> TAB <size> TAB <text>
'                       (text: Cr as \r, vertical tab / line break as \v, Tab as \t; formatting of the first character)
'   <name>.saved.docx   the document saved as Word saves it after opening it
'   <name>.edited.docx  each mapped control's text set to "edit <n>" (a check box toggled), then saved
' A Word of its own for each document and for each of the two passes.

Const wdFormatXMLDocument = 12
Const wdDoNotSaveChanges = 0
Const wdContentControlCheckBox = 8

Dim fso, inDir, outDir, w, f
Set fso = CreateObject("Scripting.FileSystemObject")
inDir = fso.GetAbsolutePathName(WScript.Arguments(0))
outDir = fso.GetAbsolutePathName(WScript.Arguments(1))
If Not fso.FolderExists(outDir) Then fso.CreateFolder outDir

Sub StartWord()
  Set w = CreateObject("Word.Application")
  w.Visible = False
  w.DisplayAlerts = 0
  w.AutomationSecurity = 3
  w.Options.UpdateLinksAtOpen = False
End Sub

Function Shown(s)
  s = Replace(s, vbCr, "\r")
  s = Replace(s, Chr(11), "\v")
  s = Replace(s, vbTab, "\t")
  s = Replace(s, Chr(7), "\a")
  Shown = s
End Function

Function OrEmpty(v)
  On Error Resume Next
  OrEmpty = ""
  OrEmpty = CStr(v)
End Function

Sub Describe(doc, out)
  Dim story, p, n, cc, x, checked
  For Each story In doc.StoryRanges
    Do While Not story Is Nothing
      n = 0
      For Each p In story.Paragraphs
        n = n + 1
        out.WriteText "para" & vbTab & story.StoryType & vbTab & n & vbTab & Shown(p.Range.Text) & vbLf
      Next
      n = 0
      For Each cc In story.ContentControls
        n = n + 1
        x = ""
        If cc.XMLMapping.IsMapped Then x = cc.XMLMapping.XPath
        checked = ""
        If cc.Type = wdContentControlCheckBox Then checked = OrEmpty(cc.Checked)
        out.WriteText "cc" & vbTab & story.StoryType & vbTab & n & vbTab & cc.Type & vbTab & cc.ShowingPlaceholderText & vbTab & _
          cc.XMLMapping.IsMapped & vbTab & x & vbTab & checked & vbTab & OrEmpty(cc.Range.Characters(1).Bold) & vbTab & _
          OrEmpty(cc.Range.Characters(1).Italic) & vbTab & OrEmpty(cc.Range.Characters(1).Font.Color) & vbTab & _
          OrEmpty(cc.Range.Characters(1).Font.Name) & vbTab & OrEmpty(cc.Range.Characters(1).Font.Size) & vbTab & Shown(cc.Range.Text) & vbLf
      Next
      Set story = story.NextStoryRange
    Loop
  Next
End Sub

Sub Edit(doc)
  Dim story, cc, n
  n = 0
  For Each story In doc.StoryRanges
    Do While Not story Is Nothing
      For Each cc In story.ContentControls
        If cc.XMLMapping.IsMapped Then
          n = n + 1
          On Error Resume Next
          If cc.Type = wdContentControlCheckBox Then
            cc.Checked = Not cc.Checked
          Else
            cc.Range.Text = "edit " & n
          End If
          If Err.Number <> 0 Then WScript.Echo "  edit " & n & ": " & Err.Description
          Err.Clear
          On Error GoTo 0
        End If
      Next
      Set story = story.NextStoryRange
    Loop
  Next
End Sub

Sub ProcessFile(path)
  Dim name, doc, out
  name = fso.GetBaseName(path)
  WScript.Echo name
  StartWord
  Set doc = w.Documents.Open(path, False, False, False)
  Set out = CreateObject("ADODB.Stream")
  out.Type = 2
  out.Charset = "utf-8"
  out.Open
  Describe doc, out
  out.SaveToFile outDir & "\" & name & ".tsv", 2
  out.Close
  doc.SaveAs2 outDir & "\" & name & ".saved.docx", wdFormatXMLDocument
  doc.Close wdDoNotSaveChanges
  w.Quit

  StartWord
  Set doc = w.Documents.Open(path, False, False, False)
  Edit doc
  doc.SaveAs2 outDir & "\" & name & ".edited.docx", wdFormatXMLDocument
  doc.Close wdDoNotSaveChanges
  w.Quit
End Sub

For Each f In fso.GetFolder(inDir).Files
  If LCase(fso.GetExtensionName(f.Name)) = "docx" Then ProcessFile f.Path
Next
WScript.Echo "DONE"
