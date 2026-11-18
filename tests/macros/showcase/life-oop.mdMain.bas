' Source: https://github.com/zq99/vba-conway-game-of-life (file mdMain.bas)
' Author: Zaid Qureshi (zq99)
' Licence: MIT (repo LICENSE, Copyright (c) 2025 Zaid Qureshi)
' What: object-oriented Conway's Game of Life: random seed, renders black/white cell fills, repaints only changed cells.
' Entry: RunSimulation
' Split over modules: life-oop.mdMain.bas plus classes life-oop.clsGameOfLife.cls, life-oop.clsGridRenderer.cls,
' life-oop.clsSimulation.cls, life-oop.clsSimulationConfig.cls and interfaces life-oop.ICellularAutomaton.cls, life-oop.IRenderer.cls.
' Uses Implements, Byte arrays, Interior.Color, ScreenUpdating pulses, EnableCancelKey, Calculation, DoEvents.
' Code below is verbatim from the source (the class files were cp1252 and are stored here as UTF-8).
Attribute VB_Name = "mdMain"
' mdMain - Run the simulation demo
' Purpose: Starts the simulation


Option Explicit

Public Sub RunSimulation()
    Dim config As New clsSimulationConfig
    
    ' (optional) tweak defaults:
    'cfg.Width = 80: cfg.Height = 50
    'cfg.AliveProb = 0.25
    'cfg.PaintEvery = 1
    'cfg.WrapEdges = True

    Dim simulation As New clsSimulation
  
    simulation.Init config, New clsGameOfLife, New clsGridRenderer
    simulation.Run topLeft:=ActiveSheet.Cells(1, 1)
    
    Set simulation = Nothing
    Set config = Nothing

End Sub

