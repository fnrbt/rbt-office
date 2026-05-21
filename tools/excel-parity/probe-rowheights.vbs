Option Explicit
Dim fso,o: Set fso=CreateObject("Scripting.FileSystemObject")
o=fso.GetAbsolutePathName(WScript.Arguments(1))
Dim x,wb,sh,f: Set x=CreateObject("Excel.Application"): x.Visible=False: x.DisplayAlerts=False
Set wb=x.Workbooks.Open(fso.GetAbsolutePathName(WScript.Arguments(0)))
Set sh=wb.Worksheets(1): Set f=fso.CreateTextFile(o,True)
Dim r
For r=1 To 145
  f.WriteLine r & vbTab & sh.Rows(r).Height
Next
f.Close: wb.Close False: x.Quit
