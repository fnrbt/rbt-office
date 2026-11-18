' Source: https://github.com/MThelic/Maze-Generator-in-VBA (files Generate, Decision, Wall, Boundry)
' Author: MThelic (GitHub)
' Licence: none stated (no LICENSE file in the repo)
' What: depth-first maze generator drawn with thick cell borders; size from AP3/AQ3 (defaults 10x10), marks start I and finish F.
' Entry: GenerateMaze
' Then run HeatMap (module maze.Solver.bas, from the companion repo) to flood-fill distances as a 3-colour scale and draw the path.
' Each repo file holds one procedure; they are concatenated here, separated by a '===== file: comment line. Otherwise verbatim.
' Uses Borders(xlEdge*), FormatConditions (cell value rules, ThemeColor), WorksheetFunction.CountIf/Min, [AQ3] evaluation, Rnd.
' ===== file: Generate =====
Sub GenerateMaze()
    L = [AQ3]     'This is the Cell where you input the Width of the maze
    If IsEmpty(L) Then
        L = 10
        [AQ3] = 10
    End If
    C = [AP3]     'This is the Cell where you input the Height of the maze
    If IsEmpty(C) Then
        C = 10
        [AP3] = 10
    End If
    Set mapa = Range("A1", Cells(L, C))
    Range("A1:AN40").Clear  'This Range you need to put the maximum size of your maze, so it will clear everything from the previous maze
    mapa.Value = 0
    With mapa
        .Borders.LineStyle = xlContinuous
        .Borders.ColorIndex = 0
        .Borders.TintAndShade = 0
        .Borders.Weight = xlThick
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    li = Int(L * Rnd + 1)
    ili = li
    co = 1
    ori = 2
    mapa.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=0"
    With mapa.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
    End With
    mapa.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1"
    With mapa.FormatConditions(2)
        .Font.ThemeColor = xlThemeColorAccent4
        .Interior.ThemeColor = xlThemeColorAccent4
    End With
    mapa.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=2"
    With mapa.FormatConditions(3)
        .Font.ThemeColor = xlThemeColorAccent6
        .Interior.ThemeColor = xlThemeColorAccent6
    End With
    While WorksheetFunction.CountIf(mapa, 0) >= 1
        tip = Cells(li, co)
        des = Decision((li), (co), (ori), (L), (C))
        If des > 4 Then
            des = des - 4
        Else
            Call Wall((li), (co), (ori), (des), (tip))
        End If
        If Cells(li, co) = 0 Then
            Cells(li, co) = 1
        Else
            Cells(li, co) = 2
        End If
        If des = 1 Then
            li = li - 1
        Else
            If des = 2 Then
                co = co - 1
            Else
                If des = 3 Then
                    li = li + 1
                Else
                    co = co + 1
                End If
            End If
        End If
        If des < 3 Then
            ori = des + 2
        Else
            ori = des - 2
        End If
    Wend
    mapa.FormatConditions.Delete
    mapa.ClearContents
    Cells(ili, 1) = "I"
    With Cells(ili, 1)
        .Interior.ThemeColor = xlThemeColorAccent1
        .Font.ThemeColor = xlThemeColorLight1
        .Font.Bold = True
        .Font.Size = 12
    End With
    fli = Int(L * Rnd + 1)
    Cells(fli, C) = "F"
    With Cells(fli, C)
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Interior.ThemeColor = xlThemeColorAccent2
        .Font.ThemeColor = xlThemeColorLight1
        .Font.Bold = True
        .Font.Size = 12
    End With
End Sub
' ===== file: Decision =====
Function Decision(li As Integer, co As Integer, ori As Integer, L As Integer, C As Integer) As Integer
    Dim ale(4), v(4), vt(4) As Integer
    If Boundry((li) - 1, (co), (L), (C)) Then
        v(1) = Cells(li - 1, co)
        If Cells(li, co).Borders(xlEdgeTop).LineStyle = xlNone Then
            vt(1) = Cells(li - 1, co)
        Else
            vt(1) = 3
        End If
    Else
        v(1) = 3
        vt(1) = 3
    End If
    If Boundry((li), (co) - 1, (L), (C)) Then
        v(2) = Cells(li, co - 1)
        If Cells(li, co).Borders(xlEdgeLeft).LineStyle = xlNone Then
            vt(2) = Cells(li, co - 1)
        Else
            vt(2) = 3
        End If
    Else
        v(2) = 3
        vt(2) = 3
    End If
    If Boundry((li) + 1, (co), (L), (C)) Then
        v(3) = Cells(li + 1, co)
        If Cells(li, co).Borders(xlEdgeBottom).LineStyle = xlNone Then
            vt(3) = Cells(li + 1, co)
        Else
            vt(3) = 3
        End If
    Else
        v(3) = 3
        vt(3) = 3
    End If
    If Boundry((li), (co) + 1, (L), (C)) Then
        v(4) = Cells(li, co + 1)
        If Cells(li, co).Borders(xlEdgeRight).LineStyle = xlNone Then
            vt(4) = Cells(li, co + 1)
        Else
            vt(4) = 3
        End If
    Else
        v(4) = 3
        vt(4) = 3
    End If
    mini = WorksheetFunction.Min(v(1), v(2), v(3), v(4))
    minit = WorksheetFunction.Min(vt(1), vt(2), vt(3), vt(4))
    i = 0
    If mini = 0 Then
        For it = 1 To 4
            If v(it) = 0 Then
                i = i + 1
                ale(i) = it
            End If
        Next
        Decision = ale(Int(i * Rnd + 1))
    Else
        For it = 1 To 4
            If vt(it) = minit Then
                i = i + 1
                ale(i) = it
            End If
        Next
        Decision = ale(Int(i * Rnd + 1)) + 4
    End If
End Function
' ===== file: Wall =====
Sub Wall(li As Integer, co As Integer, ori As Integer, des As Integer, tip As Integer)
    If tip > 1 Then
        If des = 1 Then
            Cells(li, co).Borders(xlEdgeTop).LineStyle = xlNone
        Else
            If des = 2 Then
                Cells(li, co).Borders(xlEdgeLeft).LineStyle = xlNone
            Else
                If des = 3 Then
                    Cells(li, co).Borders(xlEdgeBottom).LineStyle = xlNone
                Else
                    Cells(li, co).Borders(xlEdgeRight).LineStyle = xlNone
                End If
            End If
        End If
    Else
        If des = 1 Or ori = 1 Then
            Cells(li, co).Borders(xlEdgeTop).LineStyle = xlNone
        End If
        If des = 2 Or ori = 2 Then
            Cells(li, co).Borders(xlEdgeLeft).LineStyle = xlNone
        End If
        If des = 3 Or ori = 3 Then
            Cells(li, co).Borders(xlEdgeBottom).LineStyle = xlNone
        End If
        If des = 4 Or ori = 4 Then
            Cells(li, co).Borders(xlEdgeRight).LineStyle = xlNone
        End If
    End If
End Sub
' ===== file: Boundry =====
Function Boundry(li As Integer, co As Integer, L As Integer, C As Integer) As Boolean
    If li <= L And li > 0 Then
        If co <= C And co > 0 Then
            Boundry = True
        Else
            Boundry = False
        End If
    Else
        Boundry = False
    End If
End Function
