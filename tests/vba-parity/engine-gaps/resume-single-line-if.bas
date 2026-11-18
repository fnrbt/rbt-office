' Resume re-executes the statement that raised the error. In a single-line If that is the statement after
' Then, not the If: real VBA raises again forever here; Rbt.Vb 0.4.0 re-evaluates the If and finishes.
' (Excel hung on it in the parity run of 2026-11-18.) Expected: Excel's behaviour - an endless loop.
Function Retried()
    Dim tries
    On Error GoTo h
    If tries < 3 Then Err.Raise 5
    Retried = "tries:" & tries
    Exit Function
h:
    tries = tries + 1
    Resume
End Function
