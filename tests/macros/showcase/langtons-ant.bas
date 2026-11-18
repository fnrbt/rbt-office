' Source: https://github.com/MonDieuJeTadore/Excel_fourmi_Langton (README.md, section "VBA code")
' Raw:    https://raw.githubusercontent.com/MonDieuJeTadore/Excel_fourmi_Langton/master/README.md
' Author: Quentin Domon (2016), per the README
' Licence: none stated (the repo has no LICENSE file)
' What: Langton's ant, 20,000 steps on black/white cell fills; each cell also counts the ant's visits (red digits).
' Entry: Langton
' The README copy is used: the repo's LangtonAntAlgo.bas adds a kernel32 Sleep Declare for slow motion, the README copy has none.
' Starts at Z30 heading "Gauche". A simulation of the code shows all 20,000 steps stay inside rows 1-241, columns D-HF
' (row 1 is reached but never crossed), ending in the diagonal "highway".
' Uses Cells.Clear, whole-sheet Cells.Interior.Color / Font.Color, ActiveWindow.Zoom, Columns.ColumnWidth, Rows.RowHeight,
' Range.Offset, Select Case on Interior.Color against vbWhite/vbBlack, implicit Variants (no Option Explicit).
' ScreenUpdating stays on, so it animates. Comments are in French.
' Code below is verbatim from the source.
Sub Langton()
    Application.ScreenUpdating = True 'on s'assure que chaque mouvement de la fourmi est affiché immédiatement
    'LANCER LA MACRO AVEC F5
    Dim Fourmi As Range 'La fourmi est une cellule Excel
    ActiveWindow.Zoom = 40 'on dézoome
    Cells.Clear 'on efface les cellules et le tracé précédent
    Cells.Interior.Color = vbWhite 'on met les cellules en noir pour éviter de voir le quadrillage
    Cells.Font.Color = vbRed 'on met la police des cellules en rouge pour la rendre facilement lisible
    
    'on rend les cellules à peu près carrées
    Columns.ColumnWidth = 2
    Rows.RowHeight = 14
    
    Set Fourmi = Range("Z30") 'On part d'une cellule Excel un peu au hasard, la "Z30"
    DirectionFourmi = "Gauche" 'on part d'une des 4 directions possibles
    
    For i = 1 To 20000 'Nombre de déplacements de la souris
        Fourmi.Value = Fourmi.Value + 1 'On écrit dans la cellule sur laquelle est la fourmi le nombre déjà écrit + 1 (pour savoir combien de fois la fourmi passe dans cette cellule)
    
        Select Case Fourmi.Interior.Color 'on regarde la couleur de la cellule sur laquelle est la fourmi
            Case vbWhite 'si la cellule est blanche
            Fourmi.Interior.Color = vbBlack 'alors elle devient noire
            Select Case DirectionFourmi 'on regarde également la direction
                Case "Gauche" 'si elle allait vers la gauche
                    DirectionFourmi = "Bas" 'alors elle va maintenant vers le bas
                    Set Fourmi = Fourmi.Offset(-1, 0) 'et elle se déplace d'une case vers le bas
                Case "Droite"
                    DirectionFourmi = "Haut"
                    Set Fourmi = Fourmi.Offset(1, 0)
                Case "Haut"
                    DirectionFourmi = "Gauche"
                    Set Fourmi = Fourmi.Offset(0, -1)
                Case "Bas" 'si elle allait vers le bas
                    DirectionFourmi = "Droite" 'alors elle va maintenant vers la droite
                    Set Fourmi = Fourmi.Offset(0, 1) 'et elle se déplace d'une case vers la droite
            End Select
    
            Case vbBlack 'si la cellule n'est pas blanche mais est noire, on fait la même chose que dans le cas précédent, mais en tournant vers la droite
            Fourmi.Interior.Color = vbWhite
            Select Case DirectionFourmi
                Case "Gauche"
                    DirectionFourmi = "Haut"
                    Set Fourmi = Fourmi.Offset(1, 0)
                Case "Droite"
                    DirectionFourmi = "Bas"
                    Set Fourmi = Fourmi.Offset(-1, 0)
                Case "Haut"
                    DirectionFourmi = "Droite"
                    Set Fourmi = Fourmi.Offset(0, 1)
                Case "Bas"
                    DirectionFourmi = "Gauche"
                    Set Fourmi = Fourmi.Offset(0, -1)
            End Select
        End Select
    Next i 'on relance la boucle jusqu'à atteindre le nombre d'itérations spécifié
End Sub
