' cscript openall.vbs DIR: opens every .docx of DIR\docs in Word, read only, and appends to DIR\opened.tsv: name, ok, pages, characters (or ERROR, number, message). Run it in the desktop session (a scheduled task with /IT): Windows refuses Word to a script over SSH. Compare each .passN.docx and .model.docx that `roundtrip.mjs --out` keeps with its .original.docx.
Set fso = CreateObject("Scripting.FileSystemObject")
base = WScript.Arguments(0)
Set w = CreateObject("Word.Application") : w.Visible = False : w.DisplayAlerts = 0 : w.AutomationSecurity = 3
For Each f In fso.GetFolder(base & "\docs").Files
  Err.Clear
  On Error Resume Next
  Set d = w.Documents.Open(f.Path, False, True, False)
  If Err.Number <> 0 Then
    Say f.Name & vbTab & "ERROR" & vbTab & Err.Number & vbTab & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ")
    Err.Clear
  Else
    n = d.Content.End : p = d.ComputeStatistics(2) : t = Len(d.Content.Text)
    Say f.Name & vbTab & "ok" & vbTab & p & vbTab & t
    d.Close 0
  End If
  On Error Goto 0
Next
Say "DONE"
w.Quit

Sub Say(text)
  Set sink = fso.OpenTextFile(base & "\opened.tsv", 8, True)
  sink.WriteLine text
  sink.Close
End Sub
