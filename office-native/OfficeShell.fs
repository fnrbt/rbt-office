module Composit.Ui.OfficeShell

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget
open RedBat.Office.Ui

/// Shared chrome renderer used by every native Office app entry point.
/// Apps provide a ChromeSpec + menus + toolbar rows + (testId -> glyph) +
/// (testId -> float) menu-x-offsets + (testId -> bool -> bool) pressed
/// predicate + title string + status text getter + document workspace widget.

type ShellRenderState = {
    OpenMenuTestId: string option
}

let closed = { OpenMenuTestId = None }
let withOpenMenu testId = { OpenMenuTestId = Some testId }

let parseHexColor (value: string) =
    let parseByte (s: string) = System.Convert.ToByte(s, 16)
    let s = if value.StartsWith("#") then value.Substring(1) else value
    if s.Length = 6 then Color.rgb (parseByte s.[0..1]) (parseByte s.[2..3]) (parseByte s.[4..5])
    else Color.rgb 17uy 17uy 17uy

/// Pre-resolved chrome colors so the renderer doesn't reparse hex on every
/// repaint. Apps usually wrap this in a private module via `Palette.from`.
type ChromePalette = {
    WindowWidth: float
    MenuWidth: float
    MenuGutterWidth: float
    DocumentTrackWidth: float
    TitleTop: Color
    TitleBottom: Color
    Chrome: Color
    ChromeTop: Color
    ChromeBottom: Color
    ChromeLight: Color
    ChromeMid: Color
    ChromeDark: Color
    Pressed: Color
    HotTop: Color
    HotBottom: Color
    Input: Color
    MenuPanel: Color
    MenuGutter: Color
    PaperDesk: Color
    Border: Color
    Text: Color
    Shadow: Color
}

let palette (spec: ChromeSpec) : ChromePalette = {
    WindowWidth = spec.WindowWidth
    MenuWidth = spec.MenuWidth
    MenuGutterWidth = spec.MenuGutterWidth
    DocumentTrackWidth = spec.DocumentTrackWidth
    TitleTop = parseHexColor spec.TitleTop
    TitleBottom = parseHexColor spec.TitleBottom
    Chrome = parseHexColor spec.Chrome
    ChromeTop = parseHexColor spec.ChromeTop
    ChromeBottom = parseHexColor spec.ChromeBottom
    ChromeLight = parseHexColor spec.ChromeLight
    ChromeMid = parseHexColor spec.ChromeMid
    ChromeDark = parseHexColor spec.ChromeDark
    Pressed = parseHexColor spec.Pressed
    HotTop = parseHexColor spec.HotTop
    HotBottom = parseHexColor spec.HotBottom
    Input = parseHexColor spec.Input
    MenuPanel = parseHexColor spec.MenuPanel
    MenuGutter = parseHexColor spec.MenuGutter
    PaperDesk = parseHexColor spec.PaperDesk
    Border = parseHexColor spec.Border
    Text = parseHexColor spec.Text
    Shadow = Color.rgba 0uy 0uy 0uy 65uy
}

let private chromeBand (p: ChromePalette) top bottom height child =
    sizedW p.WindowWidth (
        stack [
            canvas p.WindowWidth height (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), top)
                  FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), bottom) ])
            child
        ])

let private toolCellBg (p: ChromePalette) isActive =
    canvas 24.0 22.0 (fun w h ->
        let fillTop, fillBottom, line =
            if isActive then
                p.Pressed, Color.rgb 195uy 217uy 247uy, Color.rgb 126uy 166uy 222uy
            else
                let clear = Color.rgba 255uy 255uy 255uy 0uy
                clear, clear, clear
        [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), fillTop)
          FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), fillBottom)
          FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, line)
          FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, line)
          FillRect(Renderer.Types.rect (w - 1.0) 0.0 1.0 h, line)
          FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, line) ])

let private menuButtonBg (p: ChromePalette) isOpen =
    canvas 1.0 21.0 (fun w h ->
        if isOpen then
            let line = Color.rgb 126uy 166uy 222uy
            [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), p.HotTop)
              FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), p.HotBottom)
              FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, line)
              FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, line)
              FillRect(Renderer.Types.rect (w - 1.0) 0.0 1.0 h, line) ]
        else
            [ FillRect(Renderer.Types.rect 0.0 0.0 w h, Palette.transparent) ])

let private flatToolbarCell (p: ChromePalette) isActive content action =
    clickable action (
        sized 24.0 22.0 (
            stack [
                toolCellBg p isActive
                center content
            ]))

let private toolbarTextIcon (p: ChromePalette) value = textStyled Font.sm p.Text value

let private miniIcon draw = canvas 16.0 16.0 draw

let private pictureIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 3.0 12.0 10.0, Color.rgb 255uy 255uy 255uy)
          FillRect(Renderer.Types.rect 2.0 3.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 12.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 3.0 1.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 13.0 3.0 1.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 4.0 9.0 3.0 2.0, Color.rgb 45uy 125uy 45uy)
          FillRect(Renderer.Types.rect 7.0 7.0 4.0 4.0, Color.rgb 45uy 125uy 45uy)
          FillRect(Renderer.Types.rect 10.0 5.0 2.0 2.0, Color.rgb 220uy 170uy 35uy) ])

let private tableIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 3.0 12.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 4.0 10.0 8.0, Color.rgb 255uy 255uy 255uy)
          for x in [ 6.0; 10.0 ] do FillRect(Renderer.Types.rect x 4.0 1.0 8.0, p.ChromeDark)
          for y in [ 7.0; 10.0 ] do FillRect(Renderer.Types.rect 3.0 y 10.0 1.0, p.ChromeDark) ])

let private documentIcon (p: ChromePalette) label =
    stack [
        miniIcon (fun _ _ ->
            [ FillRect(Renderer.Types.rect 4.0 2.0 8.0 12.0, Color.rgb 255uy 255uy 255uy)
              FillRect(Renderer.Types.rect 4.0 2.0 8.0 1.0, p.Border)
              FillRect(Renderer.Types.rect 4.0 13.0 8.0 1.0, p.Border)
              FillRect(Renderer.Types.rect 4.0 2.0 1.0 12.0, p.Border)
              FillRect(Renderer.Types.rect 11.0 2.0 1.0 12.0, p.Border)
              FillRect(Renderer.Types.rect 6.0 6.0 5.0 1.0, p.ChromeDark)
              FillRect(Renderer.Types.rect 6.0 8.0 5.0 1.0, p.ChromeDark) ])
        center (textStyled Font.sm p.Text label)
    ]

let private linkIcon =
    let blue = Color.rgb 20uy 78uy 160uy
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 7.0 5.0 2.0, blue)
          FillRect(Renderer.Types.rect 8.0 7.0 5.0 2.0, blue)
          FillRect(Renderer.Types.rect 6.0 6.0 4.0 1.0, blue)
          FillRect(Renderer.Types.rect 6.0 9.0 4.0 1.0, blue) ])

let private paragraphIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 4.0 3.0 2.0 10.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 3.0 2.0 10.0, p.Text)
          FillRect(Renderer.Types.rect 4.0 3.0 5.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 4.0 7.0 5.0 1.0, p.Text) ])

let private alignIcon (p: ChromePalette) mode =
    miniIcon (fun _ _ ->
        let widths =
            match mode with
            | "center" -> [ 8.0; 12.0; 8.0; 10.0 ]
            | "right" -> [ 8.0; 12.0; 10.0; 7.0 ]
            | "justify" -> [ 12.0; 12.0; 12.0; 12.0 ]
            | _ -> [ 12.0; 8.0; 11.0; 7.0 ]
        [ for i, w in widths |> List.indexed do
            let x =
                match mode with
                | "center" -> 8.0 - w / 2.0
                | "right" -> 14.0 - w
                | _ -> 2.0
            FillRect(Renderer.Types.rect x (3.0 + float i * 3.0) w 1.0, p.Text) ])

let private listIcon (p: ChromePalette) numbered =
    miniIcon (fun _ _ ->
        [ for i in 0..2 do
            let y = 4.0 + float i * 4.0
            if numbered then FillRect(Renderer.Types.rect 2.0 y 2.0 1.0, p.Text)
            else FillRect(Renderer.Types.rect 2.0 y 2.0 2.0, p.Text)
            FillRect(Renderer.Types.rect 6.0 y 8.0 1.0, p.Text) ])

let private filterIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 3.0 10.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 5.0 6.0 6.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 9.0 2.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 11.0 2.0 1.0, p.Text) ])

let private chartIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 11.0 2.0 3.0, Color.rgb 45uy 125uy 45uy)
          FillRect(Renderer.Types.rect 6.0 8.0 2.0 6.0, Color.rgb 200uy 100uy 20uy)
          FillRect(Renderer.Types.rect 9.0 5.0 2.0 9.0, Color.rgb 20uy 78uy 160uy)
          FillRect(Renderer.Types.rect 12.0 2.0 2.0 12.0, Color.rgb 130uy 36uy 152uy) ])

let private pivotIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 3.0 12.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 4.0 10.0 8.0, Color.rgb 240uy 240uy 240uy)
          FillRect(Renderer.Types.rect 3.0 4.0 10.0 2.0, Color.rgb 200uy 215uy 240uy)
          FillRect(Renderer.Types.rect 3.0 4.0 4.0 8.0, Color.rgb 200uy 215uy 240uy)
          for x in [ 7.0; 10.0 ] do FillRect(Renderer.Types.rect x 4.0 1.0 8.0, p.ChromeDark)
          for y in [ 6.0; 8.0; 10.0 ] do FillRect(Renderer.Types.rect 3.0 y 10.0 1.0, p.ChromeDark) ])

let private mergeIcon (p: ChromePalette) =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 4.0 12.0 8.0, p.Input)
          FillRect(Renderer.Types.rect 2.0 4.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 11.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 4.0 1.0 8.0, p.Border)
          FillRect(Renderer.Types.rect 13.0 4.0 1.0 8.0, p.Border)
          FillRect(Renderer.Types.rect 7.0 6.0 2.0 4.0, p.Text) ])

/// Built-in glyph table. Apps return glyph names from their own testId map;
/// new icons can be added here once and reused everywhere.
let private renderGlyph (p: ChromePalette) (label: string) (glyph: string) =
    match glyph with
    | "search" -> Icon.search 14.0 p.Text
    | "table" -> tableIcon p
    | "bold" -> textBoldStyled Font.sm p.Text "B"
    | "italic" -> textStyled Font.sm p.Text "I"
    | "underline" -> textStyled Font.sm p.Text "U"
    | "strike" -> toolbarTextIcon p "S"
    | "clear-formatting" -> toolbarTextIcon p "Tx"
    | "superscript" -> toolbarTextIcon p "x²"
    | "subscript" -> toolbarTextIcon p "x₂"
    | "align-left" -> alignIcon p "left"
    | "align-center" -> alignIcon p "center"
    | "align-right" -> alignIcon p "right"
    | "align-justify" -> alignIcon p "justify"
    | "valign-top" -> toolbarTextIcon p "↑"
    | "valign-middle" -> toolbarTextIcon p "↕"
    | "valign-bottom" -> toolbarTextIcon p "↓"
    | "wrap-text" -> toolbarTextIcon p "↵"
    | "bullet-list" -> listIcon p false
    | "numbered-list" -> listIcon p true
    | "indent-decrease" -> Icon.arrowLeft 13.0 p.Text
    | "indent-increase" -> Icon.arrowRight 13.0 p.Text
    | "heading-1" -> toolbarTextIcon p "H1"
    | "heading-2" -> toolbarTextIcon p "H2"
    | "heading-3" -> toolbarTextIcon p "H3"
    | "normal" -> toolbarTextIcon p "N"
    | "document" -> documentIcon p ""
    | "picture" -> pictureIcon p
    | "zoom-100" -> toolbarTextIcon p "100"
    | "date-time" -> toolbarTextIcon p "Dt"
    | "link" -> linkIcon
    | "row-add" -> toolbarTextIcon p "+R"
    | "column-add" -> toolbarTextIcon p "+C"
    | "row-delete" -> toolbarTextIcon p "-R"
    | "column-delete" -> toolbarTextIcon p "-C"
    | "paragraph" -> paragraphIcon p
    | "filter" -> filterIcon p
    | "filter-clear" -> toolbarTextIcon p "⌀"
    | "chart" -> chartIcon p
    | "pivot" -> pivotIcon p
    | "merge" -> mergeIcon p
    | "unmerge" -> toolbarTextIcon p "↔"
    | "sort-asc" -> toolbarTextIcon p "A↓"
    | "sort-desc" -> toolbarTextIcon p "Z↓"
    | "format-number" -> toolbarTextIcon p "1.2"
    | "format-currency" -> toolbarTextIcon p "$"
    | "format-percent" -> toolbarTextIcon p "%"
    | "format-date" -> toolbarTextIcon p "Dt"
    | "decimal-plus" -> toolbarTextIcon p ".0+"
    | "decimal-minus" -> toolbarTextIcon p ".0-"
    | "cut" -> toolbarTextIcon p "✂"
    | "copy" -> toolbarTextIcon p "⎘"
    | "paste" -> toolbarTextIcon p "⎗"
    | "undo" -> Icon.arrowLeft 13.0 p.Text
    | "redo" -> Icon.arrowRight 13.0 p.Text
    | "border" -> toolbarTextIcon p "▦"
    | "freeze" -> toolbarTextIcon p "❄"
    | "split" -> toolbarTextIcon p "⊟"
    | "sheet-add" -> toolbarTextIcon p "+S"
    | "sheet-delete" -> toolbarTextIcon p "-S"
    | "sheet-rename" -> toolbarTextIcon p "Sr"
    | "text-color" -> toolbarTextIcon p "A"
    | "highlight-color" -> toolbarTextIcon p "▮"
    | "all-caps" -> toolbarTextIcon p "AA"
    | "small-caps" -> toolbarTextIcon p "Aa"
    | _ -> toolbarTextIcon p label

let private menuButton (p: ChromePalette) label isOpen action =
    clickable action (
        stack [
            menuButtonBg p isOpen
            padXY 7.0 4.0 (textStyled Font.sm p.Text label)
        ])

let private comboControl (p: ChromePalette) label action =
    clickable action (
        stack [
            canvas 1.0 22.0 (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, p.Input)
                  FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, p.ChromeMid)
                  FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, p.ChromeMid)
                  FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, p.ChromeMid) ])
            border p.ChromeDark 1.0 0.0 (
                padXY 5.0 3.0 (textStyled Font.sm p.Text label))
        ])

let private colorControl (p: ChromePalette) value action =
    clickable action (
        stack [
            canvas 1.0 20.0 (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, p.Input)
                  FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, p.ChromeMid)
                  FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, p.ChromeMid)
                  FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, p.ChromeMid) ])
            border p.ChromeDark 1.0 0.0 (
                padXY 5.0 4.0 (
                    sized 20.0 9.0 (box (parseHexColor value) empty)))
        ])

/// Per-app configuration handed to `view`. ToolbarGlyph maps a testId to one
/// of the glyph names known to `renderGlyph` (or to a fallback string that
/// becomes the button label).
type ShellConfig = {
    Chrome: ChromeSpec
    Title: string
    StatusText: unit -> string
    Menus: MenuSpec list
    ToolbarRows: ToolbarRowSpec list
    ToolbarGlyph: string -> string
    MenuLeft: string -> float
    ShowToolbarPressed: string -> bool -> bool
    /// Vertical workspace widget — the document canvas, grid, etc. Sized by
    /// the caller to fit within the window body.
    Workspace: Widget
}

let private renderToolbarControl (p: ChromePalette) (config: ShellConfig) = function
    | ToolbarButton spec ->
        let isActive = spec.IsActive |> Option.map (fun active -> active()) |> Option.defaultValue false
        let pressed = config.ShowToolbarPressed spec.TestId isActive
        let icon = renderGlyph p spec.Label (config.ToolbarGlyph spec.TestId)
        flatToolbarCell p pressed icon spec.Action
    | ToolbarSelect spec ->
        let label =
            spec.Options()
            |> List.tryFind (fun option -> option.Value = spec.Value())
            |> Option.map _.Label
            |> Option.defaultValue (spec.Value())
        comboControl p label (fun () -> spec.Action(spec.Value()))
    | ToolbarInput spec ->
        match spec.Kind with
        | NumberInput -> comboControl p (spec.Value()) (fun () -> spec.Action(spec.Value()))
        | ColorInput -> colorControl p (spec.Value()) (fun () -> spec.Action(spec.Value()))
        | TextInput -> comboControl p (spec.Value()) (fun () -> spec.Action(spec.Value()))

let private toolbarGrip (p: ChromePalette) =
    sized 7.0 18.0 (
        stack [
            box p.Chrome empty
            canvas 7.0 18.0 (fun _ _ ->
                [ FillRect(Renderer.Types.rect 1.0 2.0 1.0 14.0, p.ChromeLight)
                  FillRect(Renderer.Types.rect 3.0 2.0 1.0 14.0, p.ChromeMid)
                  for y in 4.0 .. 4.0 .. 14.0 do FillRect(Renderer.Types.rect 1.0 y 4.0 1.0, p.ChromeDark) ])
        ])

let private renderToolbarRow (p: ChromePalette) (config: ShellConfig) (row: ToolbarRowSpec) =
    sizedH 24.0 (
        rowAlign MStart CCenter 1.0 (toolbarGrip p :: (row.Controls |> List.map (renderToolbarControl p config))))

let private renderMenuButton (p: ChromePalette) (renderState: ShellRenderState) (spec: MenuSpec) =
    let isOpen = renderState.OpenMenuTestId = Some spec.TestId
    menuButton p spec.Label isOpen ignore

let private renderMenuPanel (p: ChromePalette) (spec: MenuSpec) =
    let gutterWidth = p.MenuGutterWidth

    let renderItem (item: MenuItemSpec) =
        let isChecked = item.IsChecked |> Option.map (fun check -> check()) |> Option.defaultValue false
        clickable item.Action (
            sizedH 25.0 (
                stack [
                    empty
                    rowAlign MStart CCenter 0.0 [
                        sizedW gutterWidth (
                            stack [
                                box p.MenuGutter empty
                                center (if isChecked then Icon.check 12.0 p.Text else empty)
                            ])
                        sizedW (p.MenuWidth - gutterWidth - 8.0) (
                            padXY 10.0 0.0 (textStyled Font.sm p.Text item.Label))
                    ]
                ]))

    let menuContents = spec.Items |> List.map renderItem
    let panelHeight = float spec.Items.Length * 25.0 + 4.0

    filtered [DropShadow(3.0, 4.0, 10, p.Shadow)] (
        rounded 0.0 p.MenuPanel (
            stack [
                canvas p.MenuWidth panelHeight (fun w h ->
                    [ FillRect(Renderer.Types.rect 0.0 0.0 w h, p.MenuPanel)
                      FillRect(Renderer.Types.rect 0.0 0.0 gutterWidth h, p.MenuGutter) ])
                border p.ChromeLight 1.0 0.0 empty
                border p.Border 1.0 0.0 (
                    pad 2.0 (columnSpaced 0.0 menuContents))
                canvas p.MenuWidth panelHeight (fun w h ->
                    [ FillRect(Renderer.Types.rect 1.0 1.0 (w - 2.0) 1.0, Color.rgba 255uy 255uy 255uy 160uy)
                      FillRect(Renderer.Types.rect gutterWidth 4.0 1.0 (h - 8.0), Color.rgb 226uy 232uy 242uy) ])
            ]))

/// Assemble the chrome around the caller's workspace widget.
let view (renderState: ShellRenderState) (config: ShellConfig) =
    let p = palette config.Chrome
    let openMenu = renderState.OpenMenuTestId |> Option.bind (fun testId -> config.Menus |> List.tryFind (fun menu -> menu.TestId = testId))
    let menuOverlay =
        match openMenu with
        | Some menu -> overlay (offset (config.MenuLeft menu.TestId) 46.0 (sizedW p.MenuWidth (renderMenuPanel p menu)))
        | None -> empty
    let titleBar =
        chromeBand p p.TitleTop p.TitleBottom 22.0 (
            sizedH 22.0 (
                padXY 6.0 0.0 (
                    rowAlign MStart CCenter 0.0 [ textBoldStyled Font.sm Palette.white config.Title ])))
    let menuBar =
        chromeBand p (Color.rgb 255uy 255uy 255uy) (Color.rgb 236uy 243uy 252uy) 24.0 (
            sizedH 24.0 (
                padXY 2.0 2.0 (
                    rowAlign MStart CCenter 1.0 (config.Menus |> List.map (renderMenuButton p renderState)))))
    let toolbarBand =
        chromeBand p (Color.rgb 255uy 255uy 255uy) (Color.rgb 226uy 238uy 251uy) 52.0 (
            padXY 2.0 2.0 (
                column (config.ToolbarRows |> List.map (renderToolbarRow p config))))
    let statusBar =
        sizedH 24.0 (
            chromeBand p p.Chrome p.Chrome 24.0 (
                padXY 6.0 4.0 (textStyled Font.sm p.Text (config.StatusText()))))
    let windowBody =
        filtered [DropShadow(0.0, 16.0, 38, p.Shadow)] (
            rounded 0.0 p.Chrome (
                border p.Border 1.0 0.0 (
                    column [
                        titleBar
                        menuBar
                        toolbarBand
                        config.Workspace
                        statusBar
                    ])))

    sizedW p.WindowWidth (
        stack [
            windowBody
            menuOverlay
        ])
