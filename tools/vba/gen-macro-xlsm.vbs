' Create a macro-enabled workbook with a known VBA module, used to validate our
' xl/vbaProject.bin reader (VbaProject.fs) against REAL Excel output.
'
' Driven by tools/excel-parity/run-excel-vbs-on-vm.sh:
'   cscript gen-macro-xlsm.vbs <input.xlsx (unused)> <output.xlsm>
'
' Requires "Trust access to the VBA project object model" in Excel's Trust
' Center (the parity VM has it enabled for the macro tooling).

Option Explicit
Dim outPath
outPath = WScript.Arguments(1)

Dim xl, wb, comp, code
Set xl = CreateObject("Excel.Application")
xl.DisplayAlerts = False

Set wb = xl.Workbooks.Add()

code = "Sub Demo()" & vbCrLf & _
       "  Range(""A1"").Value = 42" & vbCrLf & _
       "  Range(""A2"").Value = Range(""A1"").Value * 2" & vbCrLf & _
       "  Cells(3, 1).Value = ""hello""" & vbCrLf & _
       "End Sub"

On Error Resume Next
Set comp = wb.VBProject.VBComponents.Add(1)   ' 1 = vbext_ct_StdModule
comp.Name = "Module1"
comp.CodeModule.AddFromString code
If Err.Number <> 0 Then
  WScript.Echo "VBProject access failed (enable Trust access to the VBA project object model): " & Err.Description
End If
On Error GoTo 0

' 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
wb.SaveAs outPath, 52
wb.Close False
xl.Quit
