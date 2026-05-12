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

/// Word's classic horizontal ruler — the strip above the document area.
let ruler =
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

/// Static placeholder workspace — a blank page on the paper desk. Used when
/// no real document widget has been provided.
let placeholderWorkspace =
    box palette.PaperDesk (
        pad 20.0 (
            center (
                filtered [] (
                    border palette.Border 1.0 0.0 (
                        sized palette.DocumentTrackWidth 360.0 (box Palette.white empty))))))

/// Wrap an editor surface in the paper-desk background so it matches the
/// rest of the Word chrome aesthetic.
let documentWorkspace (editor: Widget) =
    box palette.PaperDesk (pad 12.0 (center editor))

let viewWith (renderState: ShellRenderState)
             (title: string)
             (statusText: unit -> string)
             (state: WordUi.ShellState)
             (actions: WordUi.ShellActions)
             (workspace: Widget) =
    let config : OfficeShell.ShellConfig = {
        Chrome = WordUi.chrome
        Title = title
        StatusText = statusText
        Menus = WordUi.menus state actions
        ToolbarRows = WordUi.toolbarRows state actions
        ToolbarGlyph = WordUi.toolbarGlyph
        MenuLeft = WordUi.menuLeft
        ShowToolbarPressed = WordUi.showToolbarPressed
        Workspace = column [ ruler; workspace ]
    }
    OfficeShell.view renderState config

let view (renderState: ShellRenderState) (state: WordUi.ShellState) (actions: WordUi.ShellActions) =
    viewWith renderState "RedBat Word - Document1" (fun () -> "Ready") state actions placeholderWorkspace
