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

let private tickMark = Color.rgb 154uy 160uy 166uy
let private subTick  = Color.rgb 218uy 220uy 224uy

/// Light, flat ruler with subtle ticks. Matches the airy chrome aesthetic.
let ruler =
    sizedH 22.0 (
        stack [
            box palette.Chrome empty
            canvas palette.WindowWidth 22.0 (fun _ h ->
                [ FillRect(Renderer.Types.rect 0.0 (h - 1.0) palette.WindowWidth 1.0, palette.ChromeMid) ])
            row [
                sizedW ((palette.WindowWidth - palette.DocumentTrackWidth) / 2.0) (box palette.Chrome empty)
                canvas palette.DocumentTrackWidth 22.0 (fun w _ ->
                    [ for x in 0.0 .. 8.0 .. w  -> FillRect(Renderer.Types.rect x 12.0 1.0 4.0, subTick)
                      for x in 0.0 .. 40.0 .. w -> FillRect(Renderer.Types.rect x 8.0 1.0 9.0, tickMark) ])
            ]
        ])

/// Static placeholder workspace — a blank page on the paper desk.
let placeholderWorkspace =
    box palette.PaperDesk (
        pad 24.0 (
            center (
                filtered [DropShadow(0.0, 1.0, 4, palette.Shadow)] (
                    border palette.Border 1.0 0.0 (
                        sized palette.DocumentTrackWidth 1056.0 (box Palette.white empty))))))

/// Wrap an editor surface in the paper-desk background. Adds the soft shadow
/// and hairline border that make the page feel like a document floating on
/// the desk (Google-Docs style).
let documentWorkspace (editor: Widget) =
    box palette.PaperDesk (
        pad 24.0 (
            center (
                filtered [DropShadow(0.0, 1.0, 4, palette.Shadow)] (
                    border palette.Border 1.0 0.0 editor))))

let viewWith (renderState: ShellRenderState)
             (title: string)
             (statusText: unit -> string)
             (state: WordUi.ShellState)
             (actions: WordUi.ShellActions)
             (extraMenus: RedBat.Office.Ui.MenuSpec list)
             (workspace: Widget) =
    let config : OfficeShell.ShellConfig = {
        Chrome = WordUi.chrome
        Title = title
        StatusText = statusText
        // Extra menus tack onto the end of the standard WordUi set so
        // the host can register app-specific entries (Tools, Help,
        // About…) without rewriting WordUi.ShellActions.
        Menus = WordUi.menus state actions @ extraMenus
        ToolbarRows = WordUi.toolbarRows state actions
        ToolbarGlyph = WordUi.toolbarGlyph
        MenuLeft = WordUi.menuLeft
        ShowToolbarPressed = WordUi.showToolbarPressed
        Workspace = column [ ruler; workspace ]
    }
    OfficeShell.view renderState config

let view (renderState: ShellRenderState) (state: WordUi.ShellState) (actions: WordUi.ShellActions) =
    viewWith renderState "RedBat Word - Document1" (fun () -> "Ready") state actions [] placeholderWorkspace
