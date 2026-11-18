Attribute VB_Name = "Mandelbrot"
' The Mandelbrot set, one cell a pixel: 96 x 64 square cells, each coloured by how many steps of
' z = z^2 + c it takes to leave the circle of radius 2 (black: it never does). Written for the rbt-office demo (MIT).
Option Explicit

Sub DrawMandelbrot()
    Const W As Long = 96, H As Long = 64, Steps As Long = 80
    Dim x As Long, y As Long, i As Long
    Dim cr As Double, ci As Double, zr As Double, zi As Double, t As Double, k As Double
    Range(Columns(1), Columns(W)).ColumnWidth = 1
    Range(Rows(1), Rows(H)).RowHeight = 9
    For y = 1 To H
        ci = (y - (H + 1) / 2) / (H / 2) * 1.15
        For x = 1 To W
            cr = -0.65 + (x - (W + 1) / 2) / (W / 2) * 1.75
            zr = 0: zi = 0: i = 0
            Do While i < Steps And zr * zr + zi * zi < 4
                t = zr * zr - zi * zi + cr
                zi = 2 * zr * zi + ci
                zr = t
                i = i + 1
            Loop
            If i = Steps Then
                Cells(y, x).Interior.Color = RGB(8, 8, 20)
            Else
                k = Sqr(i / Steps)
                Cells(y, x).Interior.Color = RGB(255 * k, 60 + 160 * k * (1 - k) * 2, 255 * (1 - k) * 0.9 + 30)
            End If
        Next
    Next
End Sub

Sub ClearSheet()
    Cells.Clear
End Sub
