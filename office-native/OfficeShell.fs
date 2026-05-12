module Composit.Ui.OfficeShell

open Renderer
open Composit.Ui.Theme
open Composit.Ui.Widget
open Composit.Ui.Reactive
open RedBat.Office.Ui

/// Shared chrome renderer used by every native Office app entry point.
/// Apps provide a ChromeSpec + menus + toolbar rows + (testId -> glyph) +
/// (testId -> float) menu-x-offsets + (testId -> bool -> bool) pressed
/// predicate + title string + status text getter + document workspace widget.

/// Captures the open dialog plus per-field value signals so the renderer
/// can read and write field values reactively. Defined as a class so it
/// uses reference equality (the underlying Dictionary and function-valued
/// DialogButton fields don't admit structural equality).
type OpenDialogState(spec: DialogSpec, values: System.Collections.Generic.Dictionary<string, Signal<string>>) =
    member _.Spec = spec
    member _.Values = values

/// Mutable, signal-backed state owned by the chrome itself. Apps construct
/// one with `createState ()` and pass it to `view`; the chrome reads and
/// writes the open-menu / open-dialog signals in response to user input.
type ShellRenderState = {
    OpenMenu: Signal<string option>
    /// Currently-open dialog, if any. Apps set this with `openDialog`.
    OpenDialog: Signal<OpenDialogState option>
}

let createState () : ShellRenderState =
    { OpenMenu = signal None
      OpenDialog = signal None }

let openDialog (state: ShellRenderState) (spec: DialogSpec) =
    let values = System.Collections.Generic.Dictionary<string, Signal<string>>()
    for field in spec.Fields do
        values.[field.Id] <- signal field.Initial
    state.OpenMenu.Value <- None
    state.OpenDialog.Value <- Some (OpenDialogState(spec, values))

let closeDialog (state: ShellRenderState) =
    state.OpenDialog.Value <- None

/// Back-compat constants used by older callers that built ShellRenderState
/// as a static record. These now create fresh signal-backed states.
let closed : ShellRenderState = createState ()
let withOpenMenu (testId: string) : ShellRenderState =
    let state = createState ()
    state.OpenMenu.Value <- Some testId
    state

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

/// Flat chrome band. Uses the palette's ChromeTop as the fill (assumes the
/// app supplied matching Top/Bottom for a single-tone look); the separator
/// hairline is drawn by `chromeBandWithBorder` below.
let private chromeBand (p: ChromePalette) top _bottom height child =
    sizedW p.WindowWidth (
        stack [
            canvas p.WindowWidth height (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, top) ])
            child
        ])

/// Same as chromeBand but draws a 1px hairline along the bottom edge.
let private chromeBandWithBorder (p: ChromePalette) fill height child =
    sizedW p.WindowWidth (
        stack [
            canvas p.WindowWidth height (fun w h ->
                [ FillRect(Renderer.Types.rect 0.0 0.0 w h, fill)
                  FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, p.ChromeMid) ])
            child
        ])

let private flatToolbarCell (p: ChromePalette) isActive content action =
    let bg = if isActive then rounded 4.0 p.Pressed empty else empty
    clickable action (
        sized 28.0 28.0 (
            stack [
                bg
                center content
            ]))

let private menuButtonChrome (p: ChromePalette) isOpen =
    if isOpen then rounded 4.0 p.Pressed empty else empty

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
            menuButtonChrome p isOpen
            padXY 12.0 6.0 (textStyled Font.sm p.Text label)
        ])

let private comboControl (p: ChromePalette) label action =
    clickable action (
        stack [
            box p.Input empty
            border p.Border 1.0 4.0 (
                padXY 8.0 5.0 (textStyled Font.sm p.Text label))
        ])

let private colorControl (p: ChromePalette) value action =
    clickable action (
        stack [
            box p.Input empty
            border p.Border 1.0 4.0 (
                padXY 4.0 4.0 (
                    sized 20.0 16.0 (box (parseHexColor value) empty)))
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

/// Thin vertical separator between toolbar groups.
let private toolbarSeparator (p: ChromePalette) =
    sized 1.0 18.0 (box p.ChromeMid empty)

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
    | ToolbarSeparator -> padXY 4.0 0.0 (toolbarSeparator p)

let private renderToolbarRow (p: ChromePalette) (config: ShellConfig) (row: ToolbarRowSpec) =
    sizedH 32.0 (
        padXY 4.0 2.0 (
            rowAlign MStart CCenter 2.0 (row.Controls |> List.map (renderToolbarControl p config))))

let private renderMenuButton (p: ChromePalette) (renderState: ShellRenderState) (spec: MenuSpec) =
    let isOpen = renderState.OpenMenu.Value = Some spec.TestId
    let toggle () =
        if isOpen then renderState.OpenMenu.Value <- None
        else renderState.OpenMenu.Value <- Some spec.TestId
    menuButton p spec.Label isOpen toggle

let private renderMenuPanel (p: ChromePalette) (renderState: ShellRenderState) (spec: MenuSpec) =
    let gutterWidth = p.MenuGutterWidth
    let itemHeight = 30.0

    let renderItem (item: MenuItemSpec) =
        let isChecked = item.IsChecked |> Option.map (fun check -> check()) |> Option.defaultValue false
        let activate () =
            item.Action ()
            renderState.OpenMenu.Value <- None
        clickable activate (
            sizedH itemHeight (
                stack [
                    box p.MenuPanel empty
                    rowAlign MStart CCenter 0.0 [
                        sizedW gutterWidth (
                            center (if isChecked then Icon.check 12.0 p.Text else empty))
                        sizedW (p.MenuWidth - gutterWidth - 8.0) (
                            padXY 8.0 0.0 (textStyled Font.sm p.Text item.Label))
                    ]
                ]))

    let menuContents = spec.Items |> List.map renderItem
    let panelHeight = float spec.Items.Length * itemHeight + 8.0

    filtered [DropShadow(0.0, 4.0, 12, p.Shadow)] (
        rounded 4.0 p.MenuPanel (
            border p.Border 1.0 4.0 (
                pad 4.0 (columnSpaced 0.0 menuContents))))

// ── Dialog renderer ───────────────────────────────────────────

let private dialogFieldWidth = 360.0

let private renderDialogField (p: ChromePalette) (values: System.Collections.Generic.Dictionary<string, Signal<string>>) (field: DialogField) : Widget =
    let valueSignal = values.[field.Id]
    let labelLine = textStyled Font.sm (parseHexColor "#5F6368") field.Label
    let inputBox =
        match field.Kind with
        | DialogTextField placeholder ->
            sized dialogFieldWidth 30.0 (
                stack [
                    rounded 4.0 p.Input empty
                    border p.Border 1.0 4.0 (
                        padXY 8.0 6.0 (textInput valueSignal placeholder))
                ])
        | DialogNumberField(_, _) ->
            sized dialogFieldWidth 30.0 (
                stack [
                    rounded 4.0 p.Input empty
                    border p.Border 1.0 4.0 (
                        padXY 8.0 6.0 (textInput valueSignal ""))
                ])
        | DialogCheckboxField ->
            let toggle () =
                valueSignal.Value <- if valueSignal.Value = "true" then "false" else "true"
            clickable toggle (
                rowAlign MStart CCenter 8.0 [
                    dynamic (fun () ->
                        let isChecked = valueSignal.Value = "true"
                        sized 16.0 16.0 (
                            stack [
                                rounded 3.0 (if isChecked then parseHexColor "#1A73E8" else p.Input) empty
                                border (if isChecked then parseHexColor "#1A73E8" else p.Border) 1.0 3.0 (
                                    if isChecked then center (Icon.check 10.0 Palette.white)
                                    else empty)
                            ]))
                    textStyled Font.sm p.Text field.Label
                ])
        | DialogSelectField options ->
            let cycle () =
                let current = valueSignal.Value
                let idx = options |> List.tryFindIndex (fun o -> o.Value = current) |> Option.defaultValue -1
                let next = (idx + 1) % options.Length
                valueSignal.Value <- options.[next].Value
            clickable cycle (
                sized dialogFieldWidth 30.0 (
                    stack [
                        rounded 4.0 p.Input empty
                        border p.Border 1.0 4.0 (
                            padXY 8.0 6.0 (
                                dynamic (fun () ->
                                    let label =
                                        options
                                        |> List.tryFind (fun o -> o.Value = valueSignal.Value)
                                        |> Option.map (fun o -> o.Label)
                                        |> Option.defaultValue valueSignal.Value
                                    rowAlign MSpaceBetween CCenter 0.0 [
                                        textStyled Font.sm p.Text label
                                        textStyled Font.sm (parseHexColor "#5F6368") "▾"
                                    ])))
                    ]))
        | DialogStatusField ->
            dynamic (fun () ->
                textStyled Font.sm (parseHexColor "#5F6368") valueSignal.Value)

    match field.Kind with
    | DialogCheckboxField ->
        // Checkbox has its label inline; skip the separate label line.
        inputBox
    | _ ->
        columnSpaced 4.0 [ labelLine; inputBox ]

let private renderDialogButton (p: ChromePalette) (renderState: ShellRenderState) (values: System.Collections.Generic.Dictionary<string, Signal<string>>) (btn: DialogButton) : Widget =
    let activate () =
        let snapshot = values |> Seq.map (fun kv -> kv.Key, kv.Value.Value) |> Map.ofSeq
        let shouldClose = btn.Action snapshot
        if shouldClose then renderState.OpenDialog.Value <- None
    let textColor = if btn.IsPrimary then Palette.white else p.Text
    let bgColor = if btn.IsPrimary then parseHexColor "#1A73E8" else p.Input
    let borderColor = if btn.IsPrimary then parseHexColor "#1A73E8" else p.Border
    clickable activate (
        sizedH 32.0 (
            stack [
                rounded 4.0 bgColor empty
                border borderColor 1.0 4.0 (
                    padXY 16.0 6.0 (textStyled Font.sm textColor btn.Label))
            ]))

let private renderDialog (p: ChromePalette) (renderState: ShellRenderState) (state: OpenDialogState) : Widget =
    let spec = state.Spec
    let title = textBoldStyled Font.lg p.Text spec.Title
    let subtitle =
        match spec.Subtitle with
        | Some s -> textStyled Font.sm (parseHexColor "#5F6368") s
        | None -> empty
    let header = columnSpaced 4.0 [ title; subtitle ]

    let fields =
        spec.Fields
        |> List.map (renderDialogField p state.Values)
        |> columnSpaced 14.0

    let buttons =
        spec.Buttons
        |> List.map (renderDialogButton p renderState state.Values)
        |> rowAlign MEnd CCenter 8.0

    let panel =
        sizedW (dialogFieldWidth + 56.0) (
            filtered [DropShadow(0.0, 8.0, 24, Color.rgba 0uy 0uy 0uy 96uy)] (
                rounded 8.0 p.MenuPanel (
                    border p.Border 1.0 8.0 (
                        pad 24.0 (
                            columnSpaced 18.0 [
                                header
                                fields
                                buttons
                            ])))))
    center panel

let private dialogScrim () =
    // Half-transparent dark layer over the whole window; clicking it does
    // nothing (force the user to use Cancel / OK to dismiss).
    sized 4000.0 4000.0 (box (Color.rgba 0uy 0uy 0uy 64uy) empty)

/// Assemble the chrome around the caller's workspace widget. Layout:
///
///   ┌─────────────────────────────────┐
///   │  Title bar         (32, white)  │
///   ├─────────────────────────────────┤  ← hairline
///   │  Menu bar          (32, white)  │
///   ├─────────────────────────────────┤  ← hairline
///   │  Toolbar (n rows)  (32 each)    │
///   ├─────────────────────────────────┤  ← hairline
///   │  Workspace                       │
///   ├─────────────────────────────────┤  ← hairline
///   │  Status bar        (22, white)  │
///   └─────────────────────────────────┘
let view (renderState: ShellRenderState) (config: ShellConfig) =
    let p = palette config.Chrome
    let titleHeight = 32.0
    let menuHeight  = 32.0
    let toolbarRowHeight = 32.0
    let statusHeight = 22.0
    let menuBarTop = titleHeight

    let openMenu () =
        renderState.OpenMenu.Value
        |> Option.bind (fun testId -> config.Menus |> List.tryFind (fun menu -> menu.TestId = testId))

    let closeMenu () = renderState.OpenMenu.Value <- None

    let menuOverlay =
        dynamic (fun () ->
            match openMenu () with
            | Some menu ->
                stack [
                    // Click-anywhere-outside catcher (drawn behind the panel).
                    overlay (clickable closeMenu (sized p.WindowWidth 2000.0 empty))
                    // The actual menu panel, positioned under the menu button.
                    overlay (offset (config.MenuLeft menu.TestId) (menuBarTop + menuHeight) (sizedW p.MenuWidth (renderMenuPanel p renderState menu)))
                ]
            | None -> empty)

    let dialogOverlay =
        dynamic (fun () ->
            match renderState.OpenDialog.Value with
            | Some state ->
                stack [
                    overlay (dialogScrim ())
                    overlay (renderDialog p renderState state)
                ]
            | None -> empty)

    let menuBar =
        dynamic (fun () ->
            chromeBandWithBorder p p.ChromeTop menuHeight (
                sizedH menuHeight (
                    padXY 6.0 0.0 (
                        rowAlign MStart CCenter 2.0 (config.Menus |> List.map (renderMenuButton p renderState))))))

    let titleBar =
        chromeBandWithBorder p p.ChromeTop titleHeight (
            sizedH titleHeight (
                padXY 14.0 0.0 (
                    rowAlign MStart CCenter 0.0 [ textBoldStyled Font.lg p.Text config.Title ])))
    let toolbarRowCount = max 1 (List.length config.ToolbarRows)
    let toolbarHeight = float toolbarRowCount * toolbarRowHeight
    let toolbarBand =
        chromeBandWithBorder p p.ChromeTop toolbarHeight (
            column (config.ToolbarRows |> List.map (renderToolbarRow p config)))
    let statusBar =
        chromeBandWithBorder p p.ChromeTop statusHeight (
            sizedH statusHeight (
                dynamic (fun () ->
                    padXY 12.0 4.0 (textStyled Font.sm p.Text (config.StatusText())))))
    let windowBody =
        column [
            titleBar
            menuBar
            toolbarBand
            expanded config.Workspace
            statusBar
        ]

    sizedW p.WindowWidth (
        stack [
            windowBody
            menuOverlay
            dialogOverlay
        ])
