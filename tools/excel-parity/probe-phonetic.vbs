' cscript probe-phonetic.vbs BOOK.xlsx OUT.tsv: opens BOOK.xlsx in Excel and writes, for each cell of column A
' of the first sheet, the furigana Excel keeps for it (Phonetics.Count, each Phonetic's Start, Length and Text)
' and what PHONETIC() of it gives, then Excel's language settings (LanguageSettings: install, UI and whether
' the East Asian editing languages are enabled). Run it with run-excel-vbs-on-vm.sh. (In the VM's Excel 2019 -
' English, no East Asian editing language - Phonetics is not there: every cell's line says "Object required";
' the language lines are phonetic-vm-language.tsv.)
Set fso = CreateObject("Scripting.FileSystemObject")
Set x = CreateObject("Excel.Application") : x.Visible = False : x.DisplayAlerts = False : x.AutomationSecurity = 3
Set wb = x.Workbooks.Open(WScript.Arguments(0), 0, True)
Set out = fso.OpenTextFile(WScript.Arguments(1), 2, True, -1)
Set s = wb.Worksheets(1)
For r = 1 To 14
  Set c = s.Cells(r, 1)
  line = c.Address(False, False) & vbTab & CStr(c.Value2)
  On Error Resume Next
  n = c.Phonetics.Count
  If Err.Number <> 0 Then n = "err " & Err.Description : Err.Clear
  line = line & vbTab & n
  For k = 1 To c.Phonetics.Count
    line = line & vbTab & c.Phonetics(k).Start & ":" & c.Phonetics(k).Length & ":" & c.Phonetics(k).Text
  Next
  line = line & vbTab & "evaluate=" & CStr(s.Evaluate("PHONETIC(" & c.Address & ")"))
  If Err.Number <> 0 Then line = line & " err " & Err.Description : Err.Clear
  On Error GoTo 0
  out.WriteLine line
Next
Set ls = x.LanguageSettings
On Error Resume Next
out.WriteLine "install" & vbTab & ls.LanguageID(1) & vbTab & "ui" & vbTab & ls.LanguageID(2) & vbTab & "help" & vbTab & ls.LanguageID(3)
out.WriteLine "japanese-enabled" & vbTab & ls.LanguagePreferredForEditing(1041) & vbTab & "chinese-simplified" & vbTab & ls.LanguagePreferredForEditing(2052) & vbTab & "korean" & vbTab & ls.LanguagePreferredForEditing(1042)
out.Close
wb.Close False
x.Quit
