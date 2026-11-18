' Source: https://codereview.stackexchange.com/questions/233533 ("Chess game: Part 1 - Creating the layout")
' Author: TinMan (https://codereview.stackexchange.com/users/171419/tinman)
' Licence: CC BY-SA 4.0 (Stack Exchange user content, posted 2019); reuse needs attribution and share-alike.
' What: builds a chessboard: merged 4-row squares, cyan/no-fill checkering, rank and file labels, Unicode pieces at 36pt.
' Entry: CreateChessBoard
' Uses Cells.Delete, Range.Resize/Offset/EntireColumn, Merge, BorderAround xlSolid/xlMedium, Interior.Color = xlNone via IIf,
' Range.Name (defines names _A8.._H1), Font.Size, Rows(n).Value = Array(...), ChrW(9812..9823) chess symbols, Public Enum,
' and a Do While loop on Range.Width/Height that nudges ColumnWidth by 0.1 until the squares are square, with DoEvents.
' No animation. This is the question's full module (its second code block); the first block repeats FitColumnsToRangeHeight.
' Code below is verbatim from the source.
Option Explicit
Public Enum PieceType
    King
    Queen
    Rook
    Bishop
    Knight
    Pawn
End Enum

Public Enum PieceColor
    Black = 9818
    White = 9812
End Enum

Public Sub CreateChessBoard()

    Application.ScreenUpdating = False
    Const RowHeight As Double = 15, RowsPerSquare As Long = 4
    Const TopLeftAddress = "B3"

    Dim Squares As Range

    Rem Reset ActiveWorksheet
    Cells.Delete

    With Range(TopLeftAddress).Resize(8 * RowsPerSquare, 9).Offset(0, -1).EntireColumn
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With

    Set Squares = Range(TopLeftAddress).Resize(8 * RowsPerSquare, 8)

    FitColumnsToRangeHeight Squares, RowHeight

    Squares.BorderAround xlSolid, xlMedium

    Dim r As Long, c As Long, n As Long

    For n = 1 To 8
        r = (n - 1) * RowsPerSquare + 1
        For c = 1 To 8
            With Squares.Cells(r, c).Resize(RowsPerSquare)
                .Merge
                .Interior.Color = IIf((n + c) Mod 2 = 0, xlNone, vbCyan)
                .Name = "_" & Chr(64 + c) & (9 - n)
            End With
        Next

        With Squares.Cells(r, 0).Resize(RowsPerSquare)
            .Merge
            .Value = Array(8, 7, 6, 5, 4, 3, 2, 1)(n - 1)
            .Font.Size = 18
        End With
    Next

    For c = 1 To 8
        With Squares(Squares.Count + c).Resize(2)
            .Merge
        End With
    Next

    With Squares
        .Font.Size = 36
    End With

    With Squares.Rows(Squares.Rows.Count + 1)
        .Font.Size = 20
        .Value = Array("A", "B", "C", "D", "E", "F", "G", "H")
    End With

    With Squares
        .Rows(1).Value = Array(ChrW(Black + Rook), ChrW(Black + Knight), ChrW(Black + Bishop), ChrW(Black + Queen), ChrW(Black + King), ChrW(Black + Bishop), ChrW(Black + Knight), ChrW(Black + Rook))
        .Rows(RowsPerSquare + 1).Value = Array(ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn), ChrW(Black + Pawn))
        .Rows(RowsPerSquare * 6 + 1).Value = Array(ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn), ChrW(White + Pawn))
        .Rows(RowsPerSquare * 7 + 1).Value = Array(ChrW(White + Rook), ChrW(White + Knight), ChrW(White + Bishop), ChrW(White + Queen), ChrW(White + King), ChrW(White + Bishop), ChrW(White + Knight), ChrW(White + Rook))
    End With

End Sub

Private Sub FitColumnsToRangeHeight(ByRef Target As Range, ByVal RowHeight As Double)
    Const Precision As Double = 0.1

    With Target
        .RowHeight = RowHeight

        Do While .Width < .Height
            .ColumnWidth = .ColumnWidth + Precision
            DoEvents
        Loop

        Do While .Width > .Height
            .ColumnWidth = .ColumnWidth - Precision
            DoEvents
        Loop

    End With
End Sub
