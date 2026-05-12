module Composit.Ui.WordUiNative

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

module OfficeShell = Composit.Ui.OfficeShell
module WordUi = RedBat.Word.Ui

type ShellRenderState = OfficeShell.ShellRenderState

let closed = OfficeShell.closed
let withOpenMenu = OfficeShell.withOpenMenu

let private palette = OfficeShell.palette WordUi.chrome

/// Word's document workspace — paper desk + page frame + ruler. The shared
/// shell wraps this in titlebar/menubar/toolbar/status chrome.
let private ruler =
    sizedH 18.0 (
        stack [
            box (Color.rgb 250uy 248uy 242uy) empty
            row [
                sizedW 38.0 (box palette.Chrome empty)
                canvas palette.DocumentTrackWidth 18.0 (fun w _ ->
                    [ for x in 0.0 .. 8.0 .. w -> FillRect(Renderer.Types.rect x 8.0 1.0 5.0, Color.rgb 200uy 194uy 183uy)
                      for x in 0.0 .. 40.0 .. w -> FillRect(Renderer.Types.rect x 3.0 1.0 12.0, Color.rgb 140uy 140uy 140uy) ])
            ]
        ])

let private documentWorkspace =
    box palette.PaperDesk (
        pad 20.0 (
            center (
                filtered [] (
                    border palette.Border 1.0 0.0 (
                        sized palette.DocumentTrackWidth 360.0 (box Palette.white empty))))))

let private workspace = column [ ruler; documentWorkspace ]

let view (renderState: ShellRenderState) (state: WordUi.ShellState) (actions: WordUi.ShellActions) =
    let config : OfficeShell.ShellConfig = {
        Chrome = WordUi.chrome
        Title = "RedBat Word - Document1"
        StatusText = fun () -> "Ready"
        Menus = WordUi.menus state actions
        ToolbarRows = WordUi.toolbarRows state actions
        ToolbarGlyph = WordUi.toolbarGlyph
        MenuLeft = WordUi.menuLeft
        ShowToolbarPressed = WordUi.showToolbarPressed
        Workspace = workspace
    }
    OfficeShell.view renderState config
