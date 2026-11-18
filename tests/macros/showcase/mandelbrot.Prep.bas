' Written for the showcase (not from the source): small square cells, as the page's picture shows them -
' the macro paints one cell per pixel and leaves the cell size to the user (zoomed out to see all 200 rows).
Sub Setup()
    Columns("A:GR").ColumnWidth = 0.42
    Rows("1:200").RowHeight = 3.75
    ActiveWindow.Zoom = 70
End Sub
