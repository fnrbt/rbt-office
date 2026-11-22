' cscript build-form-controls.vbs OUT.tsv PNGDIR: builds synthetic-form-controls.xlsx in Excel (saved into PNGDIR) -
' every kind of form control (check boxes on / off / mixed, option buttons in a group box, a combo box, a list box,
' a spinner, scroll bars across and down, a label, a button), flat on sheet "Flat" and with 3-D shading on sheet
' "Shaded", each linked to a cell, plus an ActiveX command button and check box on "ActiveX" - for how Excel 2019
' draws a form control and what it reads of its linked cell: pictures with shoot-sheets.ps1, compared in
' tools/xlsx-browser `controls-excel-parity`. OUT gets one line a control: sheet, name, kind, left, top, width,
' height (points), linked cell, its value.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Do While wb.Worksheets.Count < 3: wb.Worksheets.Add , wb.Worksheets(wb.Worksheets.Count): Loop
names = Array("Flat", "Shaded")
For s = 0 To 1
  Set sh = wb.Worksheets(s + 1): sh.Name = names(s)
  ' input range of the combo and list boxes
  items = Array("North", "South", "East", "West", "Central", "Overseas", "Online", "Other")
  For i = 0 To 7: sh.Cells(i + 1, 12).Value = items(i): Next
  Set c = sh.CheckBoxes.Add(10, 10, 100, 17): c.Caption = "Checked": c.LinkedCell = "$J$1"
  Set c = sh.CheckBoxes.Add(10, 32, 100, 17): c.Caption = "Unchecked": c.LinkedCell = "$J$2"
  Set c = sh.CheckBoxes.Add(10, 54, 100, 17): c.Caption = "Mixed": c.Value = 2
  Set g = sh.GroupBoxes.Add(130, 8, 120, 80): g.Caption = "Group box"
  Set c = sh.OptionButtons.Add(140, 24, 90, 17): c.Caption = "Option one": c.LinkedCell = "$J$4"
  Set c = sh.OptionButtons.Add(140, 44, 90, 17): c.Caption = "Option two"
  Set c = sh.OptionButtons.Add(140, 64, 90, 17): c.Caption = "Option three"
  Set c = sh.DropDowns.Add(270, 10, 110, 18): c.ListFillRange = "$L$1:$L$8": c.LinkedCell = "$J$5": c.DropDownLines = 8
  Set c = sh.ListBoxes.Add(270, 40, 110, 70): c.ListFillRange = "$L$1:$L$8": c.LinkedCell = "$J$6"
  Set c = sh.Spinners.Add(400, 10, 18, 34): c.Min = 0: c.Max = 100: c.SmallChange = 5: c.LinkedCell = "$J$7"
  Set c = sh.ScrollBars.Add(10, 120, 200, 17): c.Min = 0: c.Max = 100: c.SmallChange = 1: c.LargeChange = 10: c.LinkedCell = "$J$8"
  Set c = sh.ScrollBars.Add(440, 10, 17, 120): c.Min = 0: c.Max = 10: c.LinkedCell = "$J$9"
  Set c = sh.Labels.Add(10, 150, 120, 17): c.Caption = "A label"
  Set c = sh.Buttons.Add(150, 145, 90, 26): c.Caption = "Button 1"
  Set c = sh.CheckBoxes.Add(10, 76, 100, 17): c.Caption = "On, no link": c.Value = 1
  ' (the linked cells' values, set after linking: linking writes the control's state into the cell)
  sh.Range("J1").Value = True: sh.Range("J2").Value = False
  sh.Range("J4").Value = 2: sh.Range("J5").Value = 3: sh.Range("J6").Value = 5
  sh.Range("J7").Value = 40: sh.Range("J8").Value = 70: sh.Range("J9").Value = 7
  If s = 1 Then
    For Each k In sh.CheckBoxes: k.Display3DShading = True: Next
    For Each k In sh.OptionButtons: k.Display3DShading = True: Next
    For Each k In sh.GroupBoxes: k.Display3DShading = True: Next
    For Each k In sh.DropDowns: k.Display3DShading = True: Next
    For Each k In sh.ListBoxes: k.Display3DShading = True: Next
    For Each k In sh.Spinners: k.Display3DShading = True: Next
    For Each k In sh.ScrollBars: k.Display3DShading = True: Next
  End If
  For Each shp In sh.Shapes
    linked = "": v = ""
    On Error Resume Next
    linked = shp.ControlFormat.LinkedCell
    v = shp.ControlFormat.Value
    On Error GoTo 0
    out.WriteLine sh.Name & vbTab & shp.Name & vbTab & shp.FormControlType & vbTab & shp.Left & vbTab & shp.Top & vbTab & shp.Width & vbTab & shp.Height & vbTab & linked & vbTab & v
  Next
Next
Set sh = wb.Worksheets(3): sh.Name = "ActiveX"
On Error Resume Next
Set o = sh.OLEObjects.Add("Forms.CommandButton.1", , , , , , , 10, 10, 100, 30)
Set o = sh.OLEObjects.Add("Forms.CheckBox.1", , , , , , , 10, 50, 100, 20)
o.LinkedCell = "$J$1"
If Err.Number <> 0 Then out.WriteLine "ActiveX" & vbTab & "error " & Err.Description
On Error GoTo 0
wb.Worksheets(1).Activate
wb.SaveAs pngDir & "\synthetic-form-controls.xlsx", 51
out.Close
wb.Close False: x.Quit
