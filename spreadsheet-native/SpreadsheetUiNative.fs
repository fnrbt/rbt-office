module Composit.Ui.SpreadsheetUiNative

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

module OfficeShell = Composit.Ui.OfficeShell
module SpreadsheetUi = RedBat.Spreadsheet.Ui
module SpreadsheetGrid = Composit.Ui.SpreadsheetGrid

open RedBat.Spreadsheet.Model

type ShellRenderState = OfficeShell.ShellRenderState

let closed = OfficeShell.closed
let withOpenMenu = OfficeShell.withOpenMenu

let private palette = OfficeShell.palette SpreadsheetUi.chrome

let private formulaBar (state: SpreadsheetUi.ShellState) =
    let cellRefBox =
        sizedW 64.0 (
            stack [
                box palette.Input empty
                border palette.ChromeMid 1.0 0.0 (
                    padXY 6.0 3.0 (textStyled Font.sm palette.Text (state.ActiveCellRef())))
            ])
    let fxLabel = padXY 6.0 4.0 (textBoldStyled Font.sm palette.Text "fx")
    let formulaInput =
        sizedW (palette.DocumentTrackWidth - 64.0 - 32.0) (
            stack [
                box palette.Input empty
                border palette.ChromeMid 1.0 0.0 (
                    padXY 6.0 3.0 (textStyled Font.sm palette.Text (state.FormulaText())))
            ])
    sizedH 22.0 (
        rowAlign MStart CCenter 4.0 [ cellRefBox; fxLabel; formulaInput ])

let private sheetTabs (state: SpreadsheetUi.ShellState) =
    let tabs = state.Sheets()
    let active = state.ActiveSheet()
    let renderTab (choice: RedBat.Office.Ui.ChoiceSpec) =
        let isActive = choice.Value = active
        let bg = if isActive then palette.Input else palette.Chrome
        sized 84.0 22.0 (
            stack [
                box bg empty
                border palette.ChromeMid 1.0 0.0 (
                    padXY 8.0 3.0 (textStyled Font.sm palette.Text choice.Label))
            ])
    sizedH 22.0 (
        rowAlign MStart CCenter 1.0 (tabs |> List.map renderTab))

/// Build the spreadsheet workspace widget: formula bar above the grid, sheet
/// tabs below.
let workspace (state: SpreadsheetUi.ShellState) (model: SpreadsheetModel) =
    let gridHeight = 360.0
    column [
        formulaBar state
        sized palette.DocumentTrackWidth gridHeight (
            border palette.Border 1.0 0.0 (SpreadsheetGrid.widget palette.DocumentTrackWidth gridHeight model))
        sheetTabs state
    ]

let view (renderState: ShellRenderState)
         (model: SpreadsheetModel)
         (state: SpreadsheetUi.ShellState)
         (actions: SpreadsheetUi.ShellActions) =
    let config : OfficeShell.ShellConfig = {
        Chrome = SpreadsheetUi.chrome
        Title = "RedBat Calc - Book1"
        StatusText = state.SelectionSummary
        Menus = SpreadsheetUi.menus state actions
        ToolbarRows = SpreadsheetUi.toolbarRows state actions
        ToolbarGlyph = SpreadsheetUi.toolbarGlyph
        MenuLeft = SpreadsheetUi.menuLeft
        ShowToolbarPressed = SpreadsheetUi.showToolbarPressed
        Workspace = workspace state model
    }
    OfficeShell.view renderState config
