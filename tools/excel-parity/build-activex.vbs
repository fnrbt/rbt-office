' cscript build-activex.vbs OUT.tsv PNGDIR: builds synthetic-activex.xlsx in Excel 2019 (saved into PNGDIR) - one ActiveX
' control of each kind of the Forms library (command button, check box, option buttons, toggle button, combo box, list
' box, text box, label, spin button, scroll bars, image), some with a caption font of their own or a value, for the
' pictures (EMF) Excel keeps of them and draws until they are activated: Emf.svg against Excel's picture
' (shoot-sheets.ps1). OUT gets a line a control: name, ProgID, left, top, width, height (points), or the error.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1): sh.Name = "ActiveX"
For i = 0 To 5: sh.Cells(i + 1, 14).Value = Array("North", "South", "East", "West", "Central", "Online")(i): Next
Sub Add(progId, l, t, w, h)
  On Error Resume Next
  Set o = sh.OLEObjects.Add(progId, , , , , , , l, t, w, h)
  If Err.Number <> 0 Then
    out.WriteLine progId & vbTab & "error " & Err.Description: Err.Clear
  Else
    out.WriteLine o.Name & vbTab & progId & vbTab & o.Left & vbTab & o.Top & vbTab & o.Width & vbTab & o.Height
  End If
  On Error GoTo 0
End Sub
Add "Forms.CommandButton.1", 10, 10, 100, 30
Add "Forms.CheckBox.1", 10, 50, 110, 20
Add "Forms.OptionButton.1", 10, 80, 110, 20
Add "Forms.OptionButton.1", 10, 105, 110, 20
Add "Forms.ToggleButton.1", 10, 135, 100, 26
Add "Forms.ComboBox.1", 140, 10, 120, 20
Add "Forms.ListBox.1", 140, 40, 120, 80
Add "Forms.TextBox.1", 140, 130, 120, 22
Add "Forms.Label.1", 140, 165, 120, 20
Add "Forms.SpinButton.1", 280, 10, 18, 36
Add "Forms.ScrollBar.1", 310, 10, 150, 18
Add "Forms.ScrollBar.1", 280, 60, 18, 120
Add "Forms.Image.1", 310, 40, 60, 40
Add "Forms.CommandButton.1", 310, 100, 140, 40
On Error Resume Next
Set objs = sh.OLEObjects
objs(1).Object.Caption = "Run report"
objs(2).Object.Caption = "Include tax": objs(2).Object.Value = True
objs(3).Object.Caption = "Monthly": objs(3).Object.Value = True
objs(4).Object.Caption = "Yearly"
objs(5).Object.Caption = "Toggle": objs(5).Object.Value = True
objs(6).ListFillRange = "$N$1:$N$6": objs(6).Object.ListIndex = 2
objs(7).ListFillRange = "$N$1:$N$6": objs(7).Object.ListIndex = 1
objs(8).Object.Text = "Typed text"
objs(9).Object.Caption = "A label": objs(9).Object.BackColor = RGB(255, 242, 204)
objs(11).Object.Value = 30
objs(12).Object.Value = 70
objs(14).Object.Caption = "Big bold": objs(14).Object.Font.Size = 14: objs(14).Object.Font.Bold = True: objs(14).Object.ForeColor = RGB(192, 0, 0)
On Error GoTo 0
sh.Range("A1").Select
' (Excel keeps the pictures it last drew of the controls: shown, so that they are drawn with their captions and values.)
x.Visible = True: x.ScreenUpdating = True
sh.Activate: WScript.Sleep 1500
x.ActiveWindow.SmallScroll 10: WScript.Sleep 500: x.ActiveWindow.ScrollRow = 1: WScript.Sleep 1500
wb.SaveAs pngDir & "\synthetic-activex.xlsx", 51
out.Close
wb.Close False: x.Quit
