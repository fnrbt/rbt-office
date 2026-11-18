' Written for the showcase (not from the source): the start pattern a user would colour in before
' running GameOfLife - an R-pentomino, which in this bounded 50x50 grid settles after 259 generations.
Sub Seed()
    Dim cell
    For Each cell In Array("AD21", "AE21", "AC22", "AD22", "AD23")
        Range(cell).Interior.ColorIndex = 1
    Next
End Sub
