' A module-level declaration after the first procedure is a compile error in VBA ("Only comments may
' appear after End Sub, End Function, or End Property"); Rbt.Vb 0.4.0 accepts it. (Excel stopped with
' that dialog in the parity run of 2026-11-18.)
Sub Main()
End Sub
Private afterwards As Long
