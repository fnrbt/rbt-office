' Source: https://github.com/u1f992/maze (src/maze.xlsm/Main.bas)
' Author: u1f992 (GitHub)
' Licence: none stated (no LICENSE file in the repo)
' What: wall-extension maze generator (red walls growing, then black) and A* shortest-path search shown as cell colours.
' Entry: Main
' Split over modules: maze-astar.Main.bas, maze-astar.Declarations.bas, maze-astar.Make.bas, maze-astar.Solve.bas, maze-astar.Utility.bas.
' The repo's ThisWorkbook.dcm (a build hook that shells out to cscript via WshShell) is left out.
' Note: SIZE = 501, so the maze is 501x501 cells and grows/deletes Range arrays one element at a time - very slow even in Excel.
' Uses Interior.Color reads and writes as the maze state, Cells.Value as distances, RowHeight/ColumnWidth, StatusBar, ReDim Preserve of Range arrays.
' Code is verbatim; the source files are Shift-JIS (Japanese comments and StatusBar text) and are stored here as UTF-8.
Attribute VB_Name = "Main"
Option Explicit

Sub Main()
    
    DefaultHeightWidth
    
    ReDim TempSearch(0 To 0)
    
    Set Corner = Cells(2, 2)
    Set RangeMaze = Range(Corner, Cells(SIZE + Corner.Row - 1, SIZE + Corner.Column - 1))
    
    Application.StatusBar = "迷路を生成しています..."
    
    MakeMaze
    
    Application.StatusBar = "スタート/ゴール地点を設定しています..."

    Set START = RangeMaze.Cells(2, 2) '左上端
    START.Interior.Color = RGB(0, 255, 0)
    
    Set GOAL = RangeMaze.Cells(RangeMaze.Rows.Count - 1, RangeMaze.Columns.Count - 1) '右下端
    GOAL.Interior.Color = RGB(255, 0, 0)
    
    Application.StatusBar = "最短経路探索を行います..."
        
    SearchSet START
    Do While SearchSet(SearchGet) = False
        DoEvents
    Loop

    SolveSet GOAL
    Do While SolveSet(SolveGet) = False
        DoEvents
    Loop
    
    Application.StatusBar = "最短経路を探索しました。"
    
End Sub

Function DefaultHeightWidth()
    ActiveSheet.Cells.Clear
    ActiveSheet.Rows.RowHeight = 18.75
    ActiveSheet.Columns.ColumnWidth = 8.38
    ActiveSheet.Range(Cells(1, 1), Cells(Rows.Count, Columns.Count)).ClearFormats
    ActiveSheet.Cells(1, 1).Select
End Function
