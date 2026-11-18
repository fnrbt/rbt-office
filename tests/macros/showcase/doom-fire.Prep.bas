' Written for the showcase (not from the source): what the gist's workbook had - the fire palette in A28:AI28
' (the first 35 colours of Doom's fire palette, black to white-hot, as Fabien Sanglard lists them) - and
' small square cells for the 78 x 26 fire, its values hidden.
Sub Setup()
    Dim palette, i
    palette = Array("070707", "1F0707", "2F0F07", "470F07", "571707", "671F07", "771F07", "8F2707", "9F2F07", _
        "AF3F07", "BF4707", "C74707", "DF4F07", "DF5707", "DF5707", "D75F07", "D75F07", "D7670F", "CF6F0F", _
        "CF770F", "CF7F0F", "CF8717", "C78717", "C78F17", "C7971F", "BF9F1F", "BF9F1F", "BFA727", "BFA727", _
        "BFAF2F", "B7AF2F", "B7B72F", "B7B737", "CFCF6F", "DFDF9F")
    For i = 0 To 34
        Cells(28, i + 1).Interior.Color = RGB(CLng("&H" & Mid(palette(i), 1, 2)), CLng("&H" & Mid(palette(i), 3, 2)), CLng("&H" & Mid(palette(i), 5, 2)))
    Next
    Columns("A:BZ").ColumnWidth = 1.29
    Rows("1:28").RowHeight = 11.25
    ' (The heat each cell holds is its value: hidden, as the fire is in its colours.)
    Range("A1:BZ26").NumberFormat = ";;;"
End Sub
