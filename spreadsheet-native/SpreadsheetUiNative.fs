module Composit.Ui.SpreadsheetUiNative

open Renderer
open Composit.Ui.Reactive
open Composit.Ui.Theme
open Composit.Ui.Widget

module OfficeShell = Composit.Ui.OfficeShell
module SpreadsheetUi = RedBat.Spreadsheet.Ui
module SpreadsheetGrid = Composit.Ui.SpreadsheetGrid
module SpreadsheetEditor = Composit.Ui.SpreadsheetEditor

open RedBat.Spreadsheet.Model

type ShellRenderState = OfficeShell.ShellRenderState

let closed = OfficeShell.closed
let withOpenMenu = OfficeShell.withOpenMenu

let private palette = OfficeShell.palette SpreadsheetUi.chrome
let private accent = OfficeShell.parseHexColor SpreadsheetUi.accentGreen
let private mutedText = OfficeShell.parseHexColor "#5F6368"

let inline private rect x y w h = Renderer.Types.rect x y w h

/// Flat formula bar — one white band, a left "name box" showing the
/// active cell ref, a hairline divider, an "fx" mark, and the formula /
/// value text filling the rest. A single hairline runs along the
/// bottom. No nested boxes-in-borders; the band is the surface.
let formulaBar (editor: SpreadsheetEditor.SpreadsheetEditorState) =
    let height = 26.0
    let nameBoxWidth = 96.0
    let chromeLayer =
        canvas palette.WindowWidth height (fun w h ->
            [ FillRect(rect 0.0 0.0 w h, palette.Input)
              // bottom hairline separates the bar from the grid
              FillRect(rect 0.0 (h - 1.0) w 1.0, palette.ChromeMid)
              // divider between the name box and the formula field
              FillRect(rect nameBoxWidth 5.0 1.0 (h - 10.0), palette.ChromeMid) ])
    let nameBox =
        sizedW nameBoxWidth (
            padXY 12.0 0.0 (
                dynamic (fun () ->
                    textStyled Font.sm palette.Text (SpreadsheetEditor.activeCellRef editor))))
    let fxMark =
        sizedW 28.0 (center (textStyled Font.sm mutedText "fx"))
    let formulaField =
        expanded (
            padXY 6.0 0.0 (
                dynamic (fun () ->
                    textStyled Font.sm palette.Text (SpreadsheetEditor.formulaText editor))))
    sizedH height (
        stack [
            chromeLayer
            rowAlign MStart CCenter 0.0 [ nameBox; fxMark; formulaField ]
        ])

/// Flat sheet-tab strip. The whole band sits on a faint desk tint with a
/// hairline along the top. The active tab is a white chip with a 2px
/// green accent on its top edge; inactive tabs are plain muted text that
/// picks up a soft grey on hover. The trailing "+" is a flat glyph.
let sheetTabs (editor: SpreadsheetEditor.SpreadsheetEditorState) (onInsert: unit -> unit) =
    let height = 30.0
    let tabWidth = 96.0
    dynamic (fun () ->
        let names = editor.SheetNames.Value
        let active = editor.ActiveSheet.Value
        let renderTab i (name: string) =
            let isActive = i = active
            let surface =
                if isActive then
                    stack [
                        box Palette.white empty
                        canvas tabWidth height (fun w _ ->
                            [ FillRect(rect 0.0 0.0 w 2.0, accent) ])
                    ]
                else empty
            let labelColor = if isActive then palette.Text else mutedText
            clickable (fun () -> SpreadsheetEditor.switchSheet editor i) (
                sized tabWidth height (
                    stack [
                        surface
                        center (textStyled Font.sm labelColor name)
                    ]))
        let addButton =
            clickable onInsert (
                sized 30.0 height (
                    center (textBoldStyled Font.lg mutedText "+")))
        let strip =
            canvas palette.WindowWidth height (fun w h ->
                [ FillRect(rect 0.0 0.0 w h, palette.PaperDesk)
                  FillRect(rect 0.0 0.0 w 1.0, palette.ChromeMid) ])
        sizedH height (
            stack [
                strip
                rowAlign MStart CCenter 0.0
                    [ for i, name in Seq.indexed names -> renderTab i name
                      yield addButton ]
            ]))

/// Static placeholder workspace — a plain grid surface on the desk.
let placeholderWorkspace =
    box palette.PaperDesk (
        center (textStyled Font.lg mutedText "No workbook open"))

/// Workspace = flat formula bar, full-bleed grid, flat sheet tabs. No
/// border boxes the grid in; the formula bar's bottom hairline and the
/// tab strip's top hairline already frame it.
let workspace
        (editor: SpreadsheetEditor.SpreadsheetEditorState)
        (gridHeight: float)
        (onInsertSheet: unit -> unit) =
    column [
        formulaBar editor
        expanded (SpreadsheetEditor.scrollWidget palette.WindowWidth gridHeight editor)
        sheetTabs editor onInsertSheet
    ]

let viewWith
        (renderState: ShellRenderState)
        (title: string)
        (statusText: unit -> string)
        (state: SpreadsheetUi.ShellState)
        (actions: SpreadsheetUi.ShellActions)
        (workspace: Widget) =
    let config : OfficeShell.ShellConfig = {
        Chrome = SpreadsheetUi.chrome
        Title = title
        StatusText = statusText
        Menus = SpreadsheetUi.menus state actions
        ToolbarRows = SpreadsheetUi.toolbarRows state actions
        ToolbarGlyph = SpreadsheetUi.toolbarGlyph
        MenuLeft = SpreadsheetUi.menuLeft
        ShowToolbarPressed = SpreadsheetUi.showToolbarPressed
        Workspace = workspace
    }
    OfficeShell.view renderState config

/// Backwards-compatible static view used by snapshot tools. Renders the
/// shell with a single sheet model and no formula bar interactivity.
let view
        (renderState: ShellRenderState)
        (model: SpreadsheetModel)
        (state: SpreadsheetUi.ShellState)
        (actions: SpreadsheetUi.ShellActions) =
    let staticWorkspace =
        expanded (SpreadsheetGrid.widget palette.WindowWidth 360.0 model)
    viewWith renderState "RedBat Calc - Book1" state.SelectionSummary state actions staticWorkspace
