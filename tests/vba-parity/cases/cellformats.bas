Sub Main()
    Dim ws
    Set ws = Worksheets(1)
    With ws.Range("F1:G2")
        .Value = "x"
        .Font.Bold = True
        .Font.Color = RGB(200, 0, 0)
        .Interior.Color = RGB(255, 255, 0)
    End With
    ws.Range("F3").Interior.ColorIndex = 6
    ws.Range("F4").Font.ColorIndex = 3
    ws.Range("F5").Value = 1234.5
    ws.Range("F5").NumberFormat = "#,##0.00"
    P "interior color", ws.Range("F1").Interior.Color
    P "interior colorindex", ws.Range("F3").Interior.Color
    P "fresh interior", ws.Range("Z9").Interior.Color
    P "fresh colorindex", ws.Range("Z9").Interior.ColorIndex
    P "font color", ws.Range("F1").Font.Color
    P "font colorindex", ws.Range("F4").Font.Color
    P "bold", ws.Range("G2").Font.Bold
    P "font name", ws.Range("Z9").Font.Name
    P "font size", ws.Range("Z9").Font.Size
    P "text", ws.Range("F5").Text
    P "numberformat general", ws.Range("Z9").NumberFormat
    P "qbcolor", QBColor(12)
    P "rgb", RGB(1, 2, 3)
    ws.Range("F6").Value = "merge"
    ws.Range("F6:G6").Merge
    P "mergecells", ws.Range("G6").MergeCells
    ws.Range("F7").HorizontalAlignment = xlCenter
    P "alignment", ws.Range("F7").HorizontalAlignment
    ws.Columns("F").ColumnWidth = 20
    P "column width", ws.Columns("F").ColumnWidth
    ws.Rows(8).RowHeight = 30
    P "row height", ws.Rows(8).RowHeight
End Sub
