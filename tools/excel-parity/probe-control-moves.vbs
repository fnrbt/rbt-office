' cscript probe-control-moves.vbs OUT.tsv PNGDIR: how Excel 2019 moves and sizes form controls (and, for comparison,
' a rectangle) when rows and columns are inserted, deleted and resized, by their Placement (move and size / move /
' free floating). Builds a sheet of controls in Excel, saves it as control-moves-before.xlsx (into PNGDIR), then does
' six edits one after the other - insert rows 1:2, insert column A, row 6 40 pt high, column D 25 characters wide,
' delete rows 7:8, delete column C - and saves control-moves-after.xlsx. OUT gets a line a shape after each step:
' step, name, control type, placement, left, top, width, height (points), top-left cell, bottom-right cell.
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(WScript.Arguments(0), True, True)
pngDir = WScript.Arguments(1)
Set x = CreateObject("Excel.Application"): x.Visible = False: x.DisplayAlerts = False
Set wb = x.Workbooks.Add()
Set sh = wb.Worksheets(1): sh.Name = "Moves"
For i = 0 To 3: sh.Cells(i + 1, 14).Value = "Item " & (i + 1): Next
Sub Named(shape, name)
  shape.Name = name
End Sub
' One of each kind at the placement Excel gives it, then check boxes and buttons at each placement.
Set c = sh.CheckBoxes.Add(60, 40, 90, 17): c.Caption = "Default": Named c, "cbDefault"
Set c = sh.CheckBoxes.Add(60, 70, 90, 17): c.Caption = "MoveSize": Named c, "cbMoveSize": c.Placement = 1
Set c = sh.CheckBoxes.Add(60, 100, 90, 17): c.Caption = "Move": Named c, "cbMove": c.Placement = 2
Set c = sh.CheckBoxes.Add(60, 130, 90, 17): c.Caption = "Free": Named c, "cbFree": c.Placement = 3
Set c = sh.Buttons.Add(200, 40, 90, 30): c.Caption = "Default": Named c, "btnDefault"
Set c = sh.Buttons.Add(200, 80, 90, 30): c.Caption = "MoveSize": Named c, "btnMoveSize": c.Placement = 1
Set c = sh.Buttons.Add(200, 120, 90, 30): c.Caption = "Move": Named c, "btnMove": c.Placement = 2
Set c = sh.Buttons.Add(200, 160, 90, 30): c.Caption = "Free": Named c, "btnFree": c.Placement = 3
Set c = sh.ScrollBars.Add(330, 40, 100, 17): Named c, "scrollDefault"
Set c = sh.DropDowns.Add(330, 70, 100, 18): c.ListFillRange = "$N$1:$N$4": Named c, "comboDefault"
Set c = sh.ListBoxes.Add(330, 100, 100, 50): c.ListFillRange = "$N$1:$N$4": Named c, "listDefault"
Set c = sh.GroupBoxes.Add(330, 160, 100, 50): c.Caption = "Group": Named c, "groupDefault"
Set c = sh.Spinners.Add(450, 40, 17, 34): Named c, "spinDefault"
Set c = sh.OptionButtons.Add(450, 90, 80, 17): c.Caption = "Option": Named c, "optionDefault"
Set c = sh.Labels.Add(450, 120, 80, 17): c.Caption = "Label": Named c, "labelDefault"
Set c = sh.Shapes.AddShape(1, 450, 160, 60, 40): c.Name = "rectDefault"
Sub Dump(stepName)
  For Each shp In sh.Shapes
    t = -1
    On Error Resume Next
    t = shp.FormControlType
    On Error GoTo 0
    out.WriteLine stepName & vbTab & shp.Name & vbTab & t & vbTab & shp.Placement & vbTab & shp.Left & vbTab & shp.Top & vbTab & shp.Width & vbTab & shp.Height & vbTab & shp.TopLeftCell.Address(False, False) & vbTab & shp.BottomRightCell.Address(False, False)
  Next
End Sub
Dump "0-built"
wb.SaveAs pngDir & "\control-moves-before.xlsx", 51
sh.Rows("1:2").Insert: Dump "1-insert-rows-1-2"
sh.Columns("A").Insert: Dump "2-insert-col-A"
sh.Rows(6).RowHeight = 40: Dump "3-row-6-40pt"
sh.Columns("D").ColumnWidth = 25: Dump "4-col-D-25"
sh.Rows("7:8").Delete: Dump "5-delete-rows-7-8"
sh.Columns("C").Delete: Dump "6-delete-col-C"
wb.SaveAs pngDir & "\control-moves-after.xlsx", 51
out.Close
wb.Close False: x.Quit
