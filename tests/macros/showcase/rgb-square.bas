' Source: https://codereview.stackexchange.com/questions/149112 ("Making an RBG chart on Excel")
' Author: Moacir (https://codereview.stackexchange.com/users/124341/moacir)
' Licence: CC BY-SA 3.0 (Stack Exchange user content, posted 2016)
' What: fills a 48x48 block of cells with a two-axis colour gradient (red fades, green grows to the right, blue grows downward).
' Entry: paint
' Before: Setup
' Run: narrow the columns so the block is square. Small, but a quick colour-gradient demo.
' Uses Cells(Row, col).Interior.Color = RGB(...), Int, lower-case "sub"/"end sub", implicit Variants. No animation.
' Code below is verbatim from the source.
sub paint()
red = 255
blue = 0
green = 0
Max = 48
For col = 1 To Max
    For Row = 1 To Max
        Cells(Row, col).Interior.Color = RGB(red - col * Int(256 / (Max * 2)) - Row * Int(256 / (Max * 2)), green + col * Int(256 / Max), blue + Row * Int(256 / Max))
    Next Row
Next col
end sub
