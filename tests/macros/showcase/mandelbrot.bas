' Source: https://slicker.me/fractals/excel.htm ("The Mandelbrot set in Excel Visual Basic")
' Author: slicker.me (no person named)
' Licence: none stated
' What: paints the Mandelbrot set into 200x200 cells, blue channel = iteration count (RGB(10,10,i*10), so it saturates).
' Entry: Worksheet_Activate
' Before: Setup
' The page gives only the loop body and says to paste it into Sheet1's Worksheet_Activate event; the
' Private Sub Worksheet_Activate() / End Sub lines are the ones the VBA editor adds for that step. The rest is verbatim.
' Uses Interior.Color, VBA.RGB with a component over 255, implicit Variants (no Option Explicit).
Private Sub Worksheet_Activate()
For y = 1 To 200
    For x = 1 To 200
        i = 0
        zx = 0
        zy = 0
        cx = -2 + x / 50
        cy = -2 + y / 50
        Do Until (i = 255) Or (zx * zx + zy * zy) >= 4
            xt = zx * zy
            zx = zx * zx - zy * zy + cx
            zy = 2 * xt + cy
            i = i + 1
        Loop
        Cells(y, x).Interior.Color = VBA.RGB(10, 10, i * 10)
    Next x
Next y
End Sub
