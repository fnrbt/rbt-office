' Written for the showcase (not from the source): small square cells, as the page's picture shows them -
' the macro paints one cell per pixel and leaves the cell size to the user.
Sub Setup()
    Columns("A:GR").ColumnWidth = 0.42
    Rows("1:200").RowHeight = 3.75
End Sub
