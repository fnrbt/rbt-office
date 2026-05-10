module Composit.Ui.WordUiNative

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget

module WordUi = RedBat.Word.Ui

type ShellRenderState = {
    OpenMenuTestId: string option
}

let closed = { OpenMenuTestId = None }

let withOpenMenu testId = { OpenMenuTestId = Some testId }

let private colorFromHex (value: string) =
    let parseByte (s: string) = System.Convert.ToByte(s, 16)
    let s = if value.StartsWith("#") then value.Substring(1) else value
    if s.Length = 6 then Color.rgb (parseByte s.[0..1]) (parseByte s.[2..3]) (parseByte s.[4..5])
    else Color.rgb 17uy 17uy 17uy

module private WordChrome =
    let spec = WordUi.chrome
    let windowWidth = spec.WindowWidth
    let menuWidth = spec.MenuWidth
    let menuGutterWidth = spec.MenuGutterWidth
    let rulerTrackWidth = spec.RulerTrackWidth
    let titleTop = colorFromHex spec.TitleTop
    let titleBottom = colorFromHex spec.TitleBottom
    let chrome = colorFromHex spec.Chrome
    let chromeTop = colorFromHex spec.ChromeTop
    let chromeBottom = colorFromHex spec.ChromeBottom
    let chromeLight = colorFromHex spec.ChromeLight
    let chromeMid = colorFromHex spec.ChromeMid
    let chromeDark = colorFromHex spec.ChromeDark
    let pressed = colorFromHex spec.Pressed
    let hotTop = colorFromHex spec.HotTop
    let hotBottom = colorFromHex spec.HotBottom
    let input = colorFromHex spec.Input
    let menuPanel = colorFromHex spec.MenuPanel
    let menuGutter = colorFromHex spec.MenuGutter
    let paperDesk = colorFromHex spec.PaperDesk
    let border = colorFromHex spec.Border
    let borderLight = chromeDark
    let text = colorFromHex spec.Text
    let selected = Color.rgb 10uy 36uy 106uy
    let shadow = Color.rgba 0uy 0uy 0uy 65uy

let private chromeBand top bottom height child =
    sizedW WordChrome.windowWidth (
        stack [
            canvas WordChrome.windowWidth height (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), top)
                  FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), bottom) ])
            child
        ])

let private toolCellBg isActive =
    canvas 24.0 22.0 (fun w h ->
        let fillTop, fillBottom, topLine, leftLine, rightLine, bottomLine =
            if isActive then
                WordChrome.pressed, Color.rgb 195uy 217uy 247uy, Color.rgb 126uy 166uy 222uy, Color.rgb 126uy 166uy 222uy, Color.rgb 126uy 166uy 222uy, Color.rgb 126uy 166uy 222uy
            else
                Color.rgba 255uy 255uy 255uy 0uy, Color.rgba 255uy 255uy 255uy 0uy, Color.rgba 255uy 255uy 255uy 0uy, Color.rgba 255uy 255uy 255uy 0uy, Color.rgba 255uy 255uy 255uy 0uy, Color.rgba 255uy 255uy 255uy 0uy
        [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), fillTop)
          FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), fillBottom)
          FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, topLine)
          FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, leftLine)
          FillRect(Renderer.Types.rect (w - 1.0) 0.0 1.0 h, rightLine)
          FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, bottomLine) ])

let private menuButtonBg isOpen =
    canvas 1.0 21.0 (fun w h ->
        if isOpen then
            [ FillRect(Renderer.Types.rect 0.0 0.0 w (h / 2.0), WordChrome.hotTop)
              FillRect(Renderer.Types.rect 0.0 (h / 2.0) w (h / 2.0), WordChrome.hotBottom)
              FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, Color.rgb 126uy 166uy 222uy)
              FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, Color.rgb 126uy 166uy 222uy)
              FillRect(Renderer.Types.rect (w - 1.0) 0.0 1.0 h, Color.rgb 126uy 166uy 222uy) ]
        else
            [ FillRect(Renderer.Types.rect 0.0 0.0 w h, Palette.transparent) ])

let private flatToolbarCell isActive content action =
    clickable action (
        sized 24.0 22.0 (
            stack [
                toolCellBg isActive
                center content
            ]))

let private toolbarTextIcon value = textStyled Font.sm WordChrome.text value

let private miniIcon draw = canvas 16.0 16.0 draw

let private pictureIcon =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 3.0 12.0 10.0, Color.rgb 255uy 255uy 255uy)
          FillRect(Renderer.Types.rect 2.0 3.0 12.0 1.0, WordChrome.border)
          FillRect(Renderer.Types.rect 2.0 12.0 12.0 1.0, WordChrome.border)
          FillRect(Renderer.Types.rect 2.0 3.0 1.0 10.0, WordChrome.border)
          FillRect(Renderer.Types.rect 13.0 3.0 1.0 10.0, WordChrome.border)
          FillRect(Renderer.Types.rect 4.0 9.0 3.0 2.0, Color.rgb 45uy 125uy 45uy)
          FillRect(Renderer.Types.rect 7.0 7.0 4.0 4.0, Color.rgb 45uy 125uy 45uy)
          FillRect(Renderer.Types.rect 10.0 5.0 2.0 2.0, Color.rgb 220uy 170uy 35uy) ])

let private tableIcon =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 3.0 12.0 10.0, WordChrome.border)
          FillRect(Renderer.Types.rect 3.0 4.0 10.0 8.0, Color.rgb 255uy 255uy 255uy)
          for x in [ 6.0; 10.0 ] do FillRect(Renderer.Types.rect x 4.0 1.0 8.0, WordChrome.chromeDark)
          for y in [ 7.0; 10.0 ] do FillRect(Renderer.Types.rect 3.0 y 10.0 1.0, WordChrome.chromeDark) ])

let private documentIcon label =
    stack [
        miniIcon (fun _ _ ->
            [ FillRect(Renderer.Types.rect 4.0 2.0 8.0 12.0, Color.rgb 255uy 255uy 255uy)
              FillRect(Renderer.Types.rect 4.0 2.0 8.0 1.0, WordChrome.border)
              FillRect(Renderer.Types.rect 4.0 13.0 8.0 1.0, WordChrome.border)
              FillRect(Renderer.Types.rect 4.0 2.0 1.0 12.0, WordChrome.border)
              FillRect(Renderer.Types.rect 11.0 2.0 1.0 12.0, WordChrome.border)
              FillRect(Renderer.Types.rect 6.0 6.0 5.0 1.0, WordChrome.chromeDark)
              FillRect(Renderer.Types.rect 6.0 8.0 5.0 1.0, WordChrome.chromeDark) ])
        center (textStyled Font.sm WordChrome.text label)
    ]

let private linkIcon =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 7.0 5.0 2.0, Color.rgb 20uy 78uy 160uy)
          FillRect(Renderer.Types.rect 8.0 7.0 5.0 2.0, Color.rgb 20uy 78uy 160uy)
          FillRect(Renderer.Types.rect 6.0 6.0 4.0 1.0, Color.rgb 20uy 78uy 160uy)
          FillRect(Renderer.Types.rect 6.0 9.0 4.0 1.0, Color.rgb 20uy 78uy 160uy) ])

let private paragraphIcon =
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 4.0 3.0 2.0 10.0, WordChrome.text)
          FillRect(Renderer.Types.rect 7.0 3.0 2.0 10.0, WordChrome.text)
          FillRect(Renderer.Types.rect 4.0 3.0 5.0 1.0, WordChrome.text)
          FillRect(Renderer.Types.rect 4.0 7.0 5.0 1.0, WordChrome.text) ])

let private alignIcon mode =
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
            FillRect(Renderer.Types.rect x (3.0 + float i * 3.0) w 1.0, WordChrome.text) ])

let private listIcon numbered =
    miniIcon (fun _ _ ->
        [ for i in 0..2 do
            let y = 4.0 + float i * 4.0
            if numbered then FillRect(Renderer.Types.rect 2.0 y 2.0 1.0, WordChrome.text)
            else FillRect(Renderer.Types.rect 2.0 y 2.0 2.0, WordChrome.text)
            FillRect(Renderer.Types.rect 6.0 y 8.0 1.0, WordChrome.text) ])

let private toolbarIcon (spec: WordUi.ButtonSpec) =
    match WordUi.toolbarGlyph spec.TestId with
    | "search" -> Icon.search 14.0 WordChrome.text
    | "table" -> tableIcon
    | "bold" -> textBoldStyled Font.sm WordChrome.text "B"
    | "italic" -> textStyled Font.sm WordChrome.text "I"
    | "underline" -> textStyled Font.sm WordChrome.text "U"
    | "strike" -> toolbarTextIcon "S"
    | "clear-formatting" -> toolbarTextIcon "Tx"
    | "superscript" -> toolbarTextIcon "x2"
    | "subscript" -> toolbarTextIcon "x2"
    | "align-left" -> alignIcon "left"
    | "align-center" -> alignIcon "center"
    | "align-right" -> alignIcon "right"
    | "align-justify" -> alignIcon "justify"
    | "bullet-list" -> listIcon false
    | "numbered-list" -> listIcon true
    | "indent-decrease" -> Icon.arrowLeft 13.0 WordChrome.text
    | "indent-increase" -> Icon.arrowRight 13.0 WordChrome.text
    | "heading-1" -> toolbarTextIcon "H1"
    | "heading-2" -> toolbarTextIcon "H2"
    | "heading-3" -> toolbarTextIcon "H3"
    | "normal" -> toolbarTextIcon "N"
    | "document" -> documentIcon ""
    | "picture" -> pictureIcon
    | "zoom-100" -> toolbarTextIcon "100"
    | "date-time" -> toolbarTextIcon "Dt"
    | "link" -> linkIcon
    | "row-add" -> toolbarTextIcon "+R"
    | "column-add" -> toolbarTextIcon "+C"
    | "row-delete" -> toolbarTextIcon "-R"
    | "column-delete" -> toolbarTextIcon "-C"
    | "paragraph" -> paragraphIcon
    | _ -> toolbarTextIcon spec.Label

let private shouldShowPressed (spec: WordUi.ButtonSpec) isActive =
    WordUi.showToolbarPressed spec.TestId isActive

let private menuButton label isOpen action =
    clickable action (
        stack [
            menuButtonBg isOpen
            padXY 7.0 4.0 (textStyled Font.sm WordChrome.text label)
        ])

let private comboControl label action =
    clickable action (
        stack [
            canvas 1.0 22.0 (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, WordChrome.input)
                  FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, WordChrome.chromeMid)
                  FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, WordChrome.chromeMid)
                  FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, WordChrome.chromeMid) ])
            border WordChrome.chromeDark 1.0 0.0 (
                padXY 5.0 3.0 (textStyled Font.sm WordChrome.text label))
        ])

let private parseHexColor (value: string) =
    colorFromHex value

let private colorControl value action =
    clickable action (
        stack [
            canvas 1.0 20.0 (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, WordChrome.input)
                  FillRect(Renderer.Types.rect 0.0 0.0 w 1.0, WordChrome.chromeMid)
                  FillRect(Renderer.Types.rect 0.0 0.0 1.0 h, WordChrome.chromeMid)
                  FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, WordChrome.chromeMid) ])
            border WordChrome.chromeDark 1.0 0.0 (
                padXY 5.0 4.0 (
                    sized 20.0 9.0 (box (parseHexColor value) empty)))
        ])

let private renderToolbarControl = function
    | WordUi.ToolbarButton spec ->
        let isActive = spec.IsActive |> Option.map (fun active -> active()) |> Option.defaultValue false
        flatToolbarCell (shouldShowPressed spec isActive) (toolbarIcon spec) spec.Action
    | WordUi.ToolbarSelect spec ->
        let label =
            spec.Options()
            |> List.tryFind (fun option -> option.Value = spec.Value())
            |> Option.map _.Label
            |> Option.defaultValue (spec.Value())
        comboControl label (fun () -> spec.Action(spec.Value()))
    | WordUi.ToolbarInput spec ->
        match spec.Kind with
        | WordUi.NumberInput -> comboControl (spec.Value()) (fun () -> spec.Action(spec.Value()))
        | WordUi.ColorInput -> colorControl (spec.Value()) (fun () -> spec.Action(spec.Value()))

let private toolbarGrip =
    sized 7.0 18.0 (
        stack [
            box WordChrome.chrome empty
            canvas 7.0 18.0 (fun _ _ ->
                [ FillRect(Renderer.Types.rect 1.0 2.0 1.0 14.0, WordChrome.chromeLight)
                  FillRect(Renderer.Types.rect 3.0 2.0 1.0 14.0, WordChrome.chromeMid)
                  for y in 4.0 .. 4.0 .. 14.0 do FillRect(Renderer.Types.rect 1.0 y 4.0 1.0, WordChrome.chromeDark) ])
        ])

let private renderToolbarRow (row: WordUi.ToolbarRowSpec) =
    sizedH 24.0 (
        rowAlign MStart CCenter 1.0 (toolbarGrip :: (row.Controls |> List.map renderToolbarControl)))

let private renderMenuButton (renderState: ShellRenderState) (spec: WordUi.MenuSpec) =
    let isOpen = renderState.OpenMenuTestId = Some spec.TestId
    menuButton spec.Label isOpen ignore

let private renderMenuPanel (spec: WordUi.MenuSpec) =
    let gutterWidth = WordChrome.menuGutterWidth

    let renderItem (item: WordUi.MenuItemSpec) =
        let isChecked = item.IsChecked |> Option.map (fun check -> check()) |> Option.defaultValue false
        clickable item.Action (
            sizedH 25.0 (
                stack [
                    empty
                    rowAlign MStart CCenter 0.0 [
                    sizedW gutterWidth (
                        stack [
                            box WordChrome.menuGutter empty
                            center (if isChecked then Icon.check 12.0 WordChrome.text else empty)
                        ])
                    sizedW (WordChrome.menuWidth - gutterWidth - 8.0) (
                        padXY 10.0 0.0 (textStyled Font.sm WordChrome.text item.Label))
                    ]
                ]))

    let menuContents = spec.Items |> List.map renderItem
    let panelHeight = float spec.Items.Length * 25.0 + 4.0

    filtered [DropShadow(3.0, 4.0, 10, WordChrome.shadow)] (
        rounded 0.0 WordChrome.menuPanel (
            stack [
                canvas WordChrome.menuWidth panelHeight (fun w h ->
                    [ FillRect(Renderer.Types.rect 0.0 0.0 w h, WordChrome.menuPanel)
                      FillRect(Renderer.Types.rect 0.0 0.0 gutterWidth h, WordChrome.menuGutter) ])
                border WordChrome.chromeLight 1.0 0.0 empty
                border WordChrome.border 1.0 0.0 (
                    pad 2.0 (columnSpaced 0.0 menuContents))
                canvas WordChrome.menuWidth panelHeight (fun w h ->
                    [ FillRect(Renderer.Types.rect 1.0 1.0 (w - 2.0) 1.0, Color.rgba 255uy 255uy 255uy 160uy)
                      FillRect(Renderer.Types.rect gutterWidth 4.0 1.0 (h - 8.0), Color.rgb 226uy 232uy 242uy) ])
            ]))

let private documentWorkspace =
    box WordChrome.paperDesk (
        pad 20.0 (
            center (
                filtered [] (
                    border WordChrome.border 1.0 0.0 (
                        sized WordChrome.rulerTrackWidth 360.0 (box Palette.white empty))))))

let private ruler =
    sizedH 18.0 (
        stack [
            box (Color.rgb 250uy 248uy 242uy) empty
            row [
                sizedW 38.0 (box WordChrome.chrome empty)
                canvas WordChrome.rulerTrackWidth 18.0 (fun w h ->
                    [ for x in 0.0 .. 8.0 .. w -> FillRect(Renderer.Types.rect x 8.0 1.0 5.0, Color.rgb 200uy 194uy 183uy)
                      for x in 0.0 .. 40.0 .. w -> FillRect(Renderer.Types.rect x 3.0 1.0 12.0, Color.rgb 140uy 140uy 140uy) ])
            ]
        ])

let view renderState state actions =
    let menus = WordUi.menus state actions
    let toolbarRows = WordUi.toolbarRows state actions
    let openMenu = renderState.OpenMenuTestId |> Option.bind (fun testId -> menus |> List.tryFind (fun menu -> menu.TestId = testId))
    let menuOverlay =
        match openMenu with
        | Some menu -> overlay (offset (WordUi.menuLeft menu.TestId) 46.0 (sizedW WordChrome.menuWidth (renderMenuPanel menu)))
        | None -> empty
    let titleBar =
        chromeBand WordChrome.titleTop WordChrome.titleBottom 22.0 (
            sizedH 22.0 (
                padXY 6.0 0.0 (
                    rowAlign MStart CCenter 0.0 [ textBoldStyled Font.sm Palette.white "RedBat Word - Document1" ])))
    let menuBar =
        chromeBand (Color.rgb 255uy 255uy 255uy) (Color.rgb 236uy 243uy 252uy) 24.0 (
            sizedH 24.0 (
                padXY 2.0 2.0 (
                    rowAlign MStart CCenter 1.0 (menus |> List.map (renderMenuButton renderState)))))
    let toolbarBand =
        chromeBand (Color.rgb 255uy 255uy 255uy) (Color.rgb 226uy 238uy 251uy) 52.0 (
            padXY 2.0 2.0 (
                column (toolbarRows |> List.map renderToolbarRow)))
    let statusBar =
        sizedH 24.0 (
            chromeBand WordChrome.chrome WordChrome.chrome 24.0 (
                padXY 6.0 4.0 (textStyled Font.sm WordChrome.text "Ready")))
    let windowBody =
        filtered [DropShadow(0.0, 16.0, 38, WordChrome.shadow)] (
            rounded 0.0 WordChrome.chrome (
                border WordChrome.border 1.0 0.0 (
                    column [
                        titleBar
                        menuBar
                        toolbarBand
                        ruler
                        documentWorkspace
                        statusBar
                    ])))

    sizedW WordChrome.windowWidth (
        stack [
            windowBody
            menuOverlay
        ])
