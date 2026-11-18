' Source: https://github.com/MThelic/Maze-Solver-in-VBA (files HeatMap, PathFinder, Fim, Topo, Baixo, Direita, Esquerda)
' Author: MThelic (GitHub)
' Licence: none stated (no LICENSE file in the repo)
' What: solves the maze GenerateMaze drew: breadth-first distance numbers shown as a 3-colour ColorScale heat map, then the path back from F.
' Sub to run after GenerateMaze: HeatMap (no arguments). Part of the maze group, see maze.Generator.bas.
' Each repo file holds one procedure; they are concatenated here, separated by a '===== file: comment line. Otherwise verbatim.
' Uses FormatConditions.AddColorScale, ColorScaleCriteria, Borders(...).LineStyle reads, WorksheetFunction.Max(Cells).
' ===== file: HeatMap =====
Sub HeatMap()
    L = [AQ3]   'This is the Cell where you input the Width of the maze
    If IsEmpty(L) Then
        L = 10
        [AQ3] = 10
    End If
    C = [AP3]   'This is the Cell where you input the Height of the maze
    If IsEmpty(C) Then
        C = 10  
        [AP3] = 10
    End If
    For li = 1 To L
        For co = 1 To C
            If IsNumeric(Cells(li, co)) Then
                Cells(li, co).ClearContents
                Cells(li, co).Interior.Pattern = xlNone
                Cells(li, co).Font.ThemeColor = xlThemeColorDark1
            End If
        Next
    Next
    Set mapa = Range("A1", Cells(L, C))
    mapa.FormatConditions.AddColorScale ColorScaleType:=3
    mapa.FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    With mapa.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
    End With
    mapa.FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
    mapa.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With mapa.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 16635892
        .TintAndShade = 0
    End With
    mapa.FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    With mapa.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
    End With
    For li = 1 To L
        For co = 1 To C
            If Cells(li, co) = "I" Then
                If topo((li), (co)) Then
                    Cells(li - 1, co) = 1
                End If
                If esquerda((li), (co)) Then
                    Cells(li, co - 1) = 1
                End If
                If baixo((li), (co), (L)) Then
                    Cells(li + 1, co) = 1
                End If
                If direita((li), (co), (C)) Then
                    Cells(li, co + 1) = 1
                End If
            End If
            If Cells(li, co) = "F" Then
                fli = li
                fco = co
            End If
        Next
    Next
    While Fim((fli), (fco), (L), (C))
        For li = 1 To L
            For co = 1 To C
                If Not IsEmpty(Cells(li, co)) And IsNumeric(Cells(li, co)) Then
                    If topo((li), (co)) Then
                        If (Cells(li - 1, co) > Cells(li, co) + 1 Or IsEmpty(Cells(li - 1, co))) And Cells(li - 1, co) <> "F" Then
                            Cells(li - 1, co) = Cells(li, co) + 1
                        End If
                    End If
                    If esquerda((li), (co)) Then
                        If (Cells(li, co - 1) > Cells(li, co) + 1 Or IsEmpty(Cells(li, co - 1))) And Cells(li, co - 1) <> "F" Then
                            Cells(li, co - 1) = Cells(li, co) + 1
                        End If
                    End If
                    If baixo((li), (co), (L)) Then
                        If (Cells(li + 1, co) > Cells(li, co) + 1 Or IsEmpty(Cells(li + 1, co))) And Cells(li + 1, co) <> "F" Then
                            Cells(li + 1, co) = Cells(li, co) + 1
                        End If
                    End If
                    If direita((li), (co), (C)) Then
                        If (Cells(li, co + 1) > Cells(li, co) + 1 Or IsEmpty(Cells(li, co + 1))) And Cells(li, co + 1) <> "F" Then
                            Cells(li, co + 1) = Cells(li, co) + 1
                        End If
                    End If
                End If
            Next
        Next
    Wend
    Call PathFinder((fli), (fco), (L), (C))
    For li = 1 To L
        For co = 1 To C
            If IsNumeric(Cells(li, co)) Then
                Cells(li, co).ClearContents
            End If
        Next
    Next
    mapa.FormatConditions.Delete
End Sub
' ===== file: PathFinder =====
Sub PathFinder(fli As Integer, fco As Integer, L As Integer, C As Integer)
    For i = 1 To Application.WorksheetFunction.Max(Cells)
        For j = -1 To 1
            For jt = -1 To 1
                If (j = 0 Or jt = 0) And fli + j > 0 And fco + jt > 0 Then
                    If Cells(fli + j, fco + jt) = "I" Then
                        If j = -1 Then
                            If topo((fli), (fco)) Then
                                Exit Sub
                            End If
                        End If
                        If jt = -1 Then
                            If esquerda((fli), (fco)) Then
                                Exit Sub
                            End If
                        End If
                        If j = 1 Then
                            If baixo((fli), (fco), L) Then
                                Exit Sub
                            End If
                        End If
                        If jt = 1 Then
                            If direita((fli), (fco), C) Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
        Next
        If topo((fli), (fco)) Then
            If Not IsEmpty(Cells(fli - 1, fco)) And IsNumeric(Cells(fli - 1, fco)) Then
                vtop = Cells(fli - 1, fco)
            Else
                vtop = 99999
            End If
        Else
            vtop = 99999
        End If
        If esquerda((fli), (fco)) Then
            If Not IsEmpty(Cells(fli, fco - 1)) And IsNumeric(Cells(fli, fco - 1)) Then
                vesq = Cells(fli, fco - 1)
            Else
                vesq = 99999
            End If
        Else
            vesq = 99999
        End If
        If baixo((fli), (fco), (L)) Then
            If Not IsEmpty(Cells(fli + 1, fco)) And IsNumeric(Cells(fli + 1, fco)) Then
                vbai = Cells(fli + 1, fco)
            Else
                vbai = 99999
            End If
        Else
            vbai = 99999
        End If
        If direita((fli), (fco), (C)) Then
            If Not IsEmpty(Cells(fli, fco + 1)) And IsNumeric(Cells(fli, fco + 1)) Then
                vdir = Cells(fli, fco + 1)
            Else
                vdir = 99999
            End If
        Else
            vdir = 99999
        End If
        vmin = Application.WorksheetFunction.Min(vtop, vesq, vbai, vdir)
        If vmin = 99999 Then
            Exit Sub
        End If
        If vmin = vtop Then
            fli = fli - 1
        Else
            If vmin = vesq Then
                fco = fco - 1
            Else
                If vmin = vbai Then
                    fli = fli + 1
                Else
                    fco = fco + 1
                End If
            End If
        End If
        Cells(fli, fco).Interior.Color = 13750992
        With Cells(fli, fco).Font
            .Color = 0
            .Bold = True
            .Size = 11
        End With
    Next
End Sub
' ===== file: Fim =====
Function Fim(fli As Integer, fco As Integer, L As Integer, C As Integer) As Boolean
    If topo((fli), (fco)) Then
        If Not IsEmpty(Cells(fli - 1, fco)) Then
            Fim = False
            Exit Function
        End If
    End If
    If esquerda((fli), (fco)) Then
        If Not IsEmpty(Cells(fli, fco - 1)) Then
            Fim = False
            Exit Function
        End If
    End If
    If baixo((fli), (fco), (L)) Then
        If Not IsEmpty(Cells(fli + 1, fco)) Then
            Fim = False
            Exit Function
        End If
    End If
    If direita((fli), (fco), (C)) Then
        If Not IsEmpty(Cells(fli, fco + 1)) Then
            Fim = False
            Exit Function
        End If
    End If
    Fim = True
End Function
' ===== file: Topo =====
Function topo(li As Integer, co As Integer) As Boolean
    If li > 1 Then
        If Cells(li, co).Borders(xlEdgeTop).LineStyle = xlNone And Cells(li - 1, co) <> "I" Then 'And Cells(li - 1, co) <> "F" Then
            topo = True
        Else
            topo = False
        End If
    Else
        topo = False
    End If
End Function
' ===== file: Baixo =====
Function baixo(li As Integer, co As Integer, L As Integer) As Boolean
    If li < L Then
        If Cells(li, co).Borders(xlEdgeBottom).LineStyle = xlNone And Cells(li + 1, co) <> "I" Then 'And Cells(li + 1, co) <> "F" Then
            baixo = True
        Else
            baixo = False
        End If
    Else
        baixo = False
    End If
End Function
' ===== file: Direita =====
Function direita(li As Integer, co As Integer, C As Integer) As Boolean
    If co < C Then
        If Cells(li, co).Borders(xlEdgeRight).LineStyle = xlNone And Cells(li, co + 1) <> "I" Then 'And Cells(li, co + 1) <> "F" Then
            direita = True
        Else
            direita = False
        End If
    Else
        direita = False
    End If
End Function
' ===== file: Esquerda =====
Function esquerda(li As Integer, co As Integer) As Boolean
    If co > 1 Then
        If Cells(li, co).Borders(xlEdgeLeft).LineStyle = xlNone And Cells(li, co - 1) <> "I" Then 'And Cells(li, co - 1) <> "F" Then
            esquerda = True
        Else
            esquerda = False
        End If
    Else
        esquerda = False
    End If
End Function
