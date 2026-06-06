Option Explicit

Dim inputPath, outputPath
inputPath = WScript.Arguments(0)
outputPath = WScript.Arguments(1)

Dim fso, out
Set fso = CreateObject("Scripting.FileSystemObject")
Set out = fso.CreateTextFile(outputPath, True, True)

Sub Emit(name, value)
  out.WriteLine name & vbTab & CStr(value)
End Sub

Dim excel, workbook, sheet, cell, range
Set excel = CreateObject("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

Set workbook = excel.Workbooks.Add()
Set sheet = workbook.Worksheets(1)
sheet.Range("A1").Value2 = 123
sheet.Range("B2").Value2 = "Hello"
Set cell = sheet.Range("A1")
Set range = sheet.Range("A1:B2")

Emit "ExcelVersion", excel.Version
Emit "AddressDefault", cell.Address
Emit "AddressRelative", cell.Address(False, False, 1, False)
Emit "RangeAddressRelative", range.Address(False, False, 1, False)
Emit "Value2", cell.Value2
Emit "Text", cell.Text
Emit "Left", Round(cell.Left, 2)
Emit "Top", Round(cell.Top, 2)
Emit "Width", Round(cell.Width, 2)
Emit "Height", Round(cell.Height, 2)
Emit "RowsCount", range.Rows.Count
Emit "ColumnsCount", range.Columns.Count
Emit "CellsAddressRelative", sheet.Cells(1, 1).Address(False, False, 1, False)

workbook.Close False
excel.Quit
out.Close
