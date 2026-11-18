' Source: https://gist.github.com/fabienbk/3a0daae4ee9f9b876336906cdf1b26f4
' Raw:    https://gist.githubusercontent.com/fabienbk/3a0daae4ee9f9b876336906cdf1b26f4/raw
' Author: fabienbk (GitHub), gist from April 2019
' Licence: none stated
' What: the "Doom fire" effect in a 78x26 block of cells. A white-hot bottom row heats the rows above with random cooling and
' drift for 100 frames, then the source goes out and the fire dies down over 100 more. Each cell holds its heat 0-35 as its value.
' Entry: doom_fire
' Before: Setup
' Setup: NOT fully self-contained. read_colors takes the 35-colour palette from A28:AI28 (DisplayFormat.Interior.Color), so
' fill those cells first with a black -> dark red -> orange -> yellow -> white ramp (the gist has no workbook).
' Without it every heat level reads as the default fill and the fire cannot be seen.
' Uses Public Const, a Public module array Colors(36), Range.DisplayFormat.Interior.Color, Interior.Color = RGB(...),
' Rnd/Randomize, Cells().Value as the heat buffer, implicit Variants. Animates only if the screen repaints inside the loop
' (it has no DoEvents or Wait).
' Code below is verbatim from the source.
Public Const Width = 78
Public Const Height = 26
Public Colors(36)

Sub spread_fire(x, y)
    rand = Int(3 * Rnd)
    randcool = Int(6 * Rnd)
    
    xshift = x - rand + 1
    If xshift <= 0 Then xshift = x
    If xshift > Width Then xshift = x
    
    newValue = Cells(y + 1, xshift).Value - (randcool)
    
    If newValue < 0 Then newValue = 0
    
    Cells(y, x).Value = newValue
    Cells(y, x).Interior.Color = Colors(newValue)
End Sub

Sub do_fire()
For y = 1 To Height - 1
For x = 1 To Width
   Call spread_fire(x, y)
Next x
Next y
End Sub

Sub read_colors()
For x = 1 To 35
Colors(x - 1) = Cells(28, x).DisplayFormat.Interior.Color
Next x
Colors(35) = RGB(255, 255, 255)
End Sub

Sub init_fire()
For y = 1 To Height - 1
For x = 1 To Width
   Cells(y, x).Interior.Color = RGB(0, 0, 0)
   Cells(y, x).Value = 0
Next x
Next y

For x = 1 To Width
Cells(26, x).Interior.Color = RGB(255, 255, 255)
Cells(26, x).Value = 35
Next x
End Sub

Sub end_fire()
For x = 1 To Width
Cells(26, x).Interior.Color = RGB(0, 0, 0)
Cells(26, x).Value = 0
Next x
End Sub


Sub doom_fire()
Randomize

Call read_colors
Call init_fire

For x = 1 To 100
Call do_fire
Next x

Call end_fire

For x = 1 To 100
Call do_fire
Next x

End Sub