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
type OpenDialogState(
    spec: DialogSpec,
    values: System.Collections.Generic.Dictionary<string, Signal<string>>,
    focusIds: System.Collections.Generic.Dictionary<string, int>,
    activeTab: Signal<int>) =
    member _.Spec = spec
    member _.Values = values
    member _.FocusIds = focusIds
    /// Index into `Spec.Tabs` of the currently-active tab. Ignored when
    /// the spec uses the un-tabbed Fields layout.
    member _.ActiveTab = activeTab

/// Captures an open context (right-click) menu plus its anchor coordinates
/// in screen space. Reference equality (the MenuItemSpec list contains
/// function-valued Action fields).
type OpenContextMenuState(x: float, y: float, items: MenuItemSpec list) =
    member _.X = x
    member _.Y = y
    member _.Items = items

/// Captures an open toolbar/dialog dropdown anchored to a sibling widget
/// (combobox trigger, color swatch). The anchor's bounds determine where
/// the panel renders — no caller-supplied (x, y).
type OpenAnchoredDropdownState(anchor: AnchorRef, items: MenuItemSpec list) =
    member _.Anchor = anchor
    member _.Items = items

/// Mutable, signal-backed state owned by the chrome itself. Apps construct
/// one with `createState ()` and pass it to `view`; the chrome reads and
/// writes these signals in response to user input.
type ShellRenderState = {
    OpenMenu: Signal<string option>
    /// Anchor refs per widget testId. Menu buttons, combo triggers, and
    /// color swatches wrap themselves in an `anchorRegion` keyed by
    /// testId; their popovers (dropdowns, panels) read the anchor's
    /// recorded bounds via `popoverBelow`. No hand-computed offsets.
    Anchors: System.Collections.Generic.Dictionary<string, AnchorRef>
    /// Currently-open dialog, if any. Apps set this with `openDialog`.
    OpenDialog: Signal<OpenDialogState option>
    /// Currently-open right-click menu (cursor-anchored). Apps set this
    /// with `openContextMenu`.
    OpenContextMenu: Signal<OpenContextMenuState option>
    /// Currently-open toolbar/dialog dropdown (widget-anchored). Opened
    /// by `openAnchoredDropdown` from combo/color controls.
    OpenAnchoredDropdown: Signal<OpenAnchoredDropdownState option>
    /// The testId of the toolbar control whose dropdown (combo or color
    /// swatches) is currently open. Mirrors `OpenMenu`'s role for menu
    /// buttons: `comboControl`/`colorControl` render the trigger in its
    /// "Pressed" highlight while their dropdown is up.
    OpenDropdownTrigger: Signal<string option>
    /// Identifier of the cell currently under the cursor (toolbar button,
    /// menu item, etc). Cells stamp this when their hoverable fires; the
    /// chrome reads it to render a hover background.
    HoveredId: Signal<string option>
    /// When a hover begins on a cell, the chrome stamps `(id, t)` here
    /// with `t = Platform.now()`. The view's status-bar tooltip path
    /// reads it to know when to surface the cell's tooltip text (after a
    /// short delay, so brushing past doesn't flicker descriptions in and
    /// out of the status bar).
    HoverStart: Signal<(string * float) option>
}

let createState () : ShellRenderState =
    { OpenMenu = signal None
      Anchors = System.Collections.Generic.Dictionary<string, AnchorRef>()
      OpenDialog = signal None
      OpenContextMenu = signal None
      OpenAnchoredDropdown = signal None
      OpenDropdownTrigger = signal None
      HoveredId = signal None
      HoverStart = signal None }

/// Delay (seconds) between hover-enter and the tooltip appearing in the
/// status bar. Long enough that you can sweep past a toolbar without the
/// status line flickering; short enough to feel responsive when you stop.
let private tooltipDelaySec = 0.5

/// Wrap a child so it stamps `HoveredId` with `id` while the cursor is
/// inside and clears it on exit (only if it's still the current id, so
/// that fast moves between cells don't clobber the new one). Also records
/// the hover-start time and schedules a redraw after `tooltipDelaySec`
/// so the status-bar tooltip path has something to fire on.
let private withHover (renderState: ShellRenderState) (id: string) (child: Widget) : Widget =
    hoverable (fun h ->
        if h then
            renderState.HoveredId.Value <- Some id
            renderState.HoverStart.Value <- Some (id, Platform.now())
            // Trigger a redraw after the delay so the tooltip-elapsed
            // check in the status bar re-evaluates even when no other
            // signal has changed in the interim.
            Platform.setTimeout (int (tooltipDelaySec * 1000.0)) (fun () ->
                Platform.requestRedraw()) |> ignore
        elif renderState.HoveredId.Value = Some id then
            renderState.HoveredId.Value <- None
            renderState.HoverStart.Value <- None)
        child

/// Get-or-create the anchor ref for a widget keyed by testId. Refs live
/// on the render state so they persist across re-renders; the layout
/// pass writes resolved bounds into them and popovers read them back.
let anchorFor (state: ShellRenderState) (testId: string) : AnchorRef =
    match state.Anchors.TryGetValue testId with
    | true, a -> a
    | _ ->
        let a = anchorRef ()
        state.Anchors.[testId] <- a
        a

let openContextMenu (state: ShellRenderState) (x: float) (y: float) (items: MenuItemSpec list) =
    state.OpenMenu.Value <- None
    state.OpenContextMenu.Value <- Some (OpenContextMenuState(x, y, items))

let closeContextMenu (state: ShellRenderState) =
    state.OpenContextMenu.Value <- None
    state.OpenDropdownTrigger.Value <- None

/// Open a dropdown anchored to the trigger widget's recorded bounds. The
/// trigger widget must have wrapped itself in `anchorRegion (anchorFor
/// state triggerId) …`. The `triggerId` is also stored on
/// `OpenDropdownTrigger` so the trigger can render its "open" highlight
/// while the dropdown is up.
let openAnchoredDropdown (state: ShellRenderState) (triggerId: string) (items: MenuItemSpec list) =
    state.OpenMenu.Value <- None
    state.OpenContextMenu.Value <- None
    state.OpenDropdownTrigger.Value <- Some triggerId
    state.OpenAnchoredDropdown.Value <- Some (OpenAnchoredDropdownState(anchorFor state triggerId, items))

let closeAnchoredDropdown (state: ShellRenderState) =
    state.OpenAnchoredDropdown.Value <- None
    state.OpenDropdownTrigger.Value <- None

let openDialog (state: ShellRenderState) (spec: DialogSpec) =
    let values = System.Collections.Generic.Dictionary<string, Signal<string>>()
    let focusIds = System.Collections.Generic.Dictionary<string, int>()
    let initField (field: DialogField) =
        values.[field.Id] <- signal field.Initial
        match field.Kind with
        | DialogTextField _ | DialogNumberField _ | DialogTextAreaField _ ->
            focusIds.[field.Id] <- Focus.allocId()
        | _ -> ()
    // Init both the flat Fields list and every tab's fields, so submit
    // callbacks see values from any tab regardless of which is currently
    // visible.
    for field in spec.Fields do initField field
    for _, tabFields in spec.Tabs do
        for field in tabFields do initField field
    let activeTab = signal 0
    state.OpenMenu.Value <- None
    state.OpenDialog.Value <- Some (OpenDialogState(spec, values, focusIds, activeTab))

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

let private flatToolbarCell
        (p: ChromePalette) (renderState: ShellRenderState) (id: string)
        (isActive: bool) (isEnabled: bool) (content: Widget) (action: unit -> unit) =
    let bg =
        dynamic (fun () ->
            if not isEnabled then empty
            elif isActive then rounded 4.0 p.Pressed empty
            elif renderState.HoveredId.Value = Some id then rounded 4.0 p.HotTop empty
            else empty)
    let body =
        sized 28.0 28.0 (
            stack [
                bg
                center content
            ])
    // Disabled cells skip both the hover stamp (no tooltip / highlight) and
    // the click action; rendered at 40% opacity to match the Google-style
    // greyed look. Enabled cells retain the full interactive surface.
    if isEnabled then
        withHover renderState id (clickable action body)
    else
        opacity 0.4 body

let private menuButtonChrome (p: ChromePalette) (renderState: ShellRenderState) (id: string) isOpen =
    dynamic (fun () ->
        if isOpen then rounded 4.0 p.Pressed empty
        elif renderState.HoveredId.Value = Some id then rounded 4.0 p.HotTop empty
        else empty)

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

let private unmergeIcon (p: ChromePalette) =
    // Two cells side by side with a divider running between them.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 4.0 12.0 8.0, p.Input)
          FillRect(Renderer.Types.rect 2.0 4.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 11.0 12.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 2.0 4.0 1.0 8.0, p.Border)
          FillRect(Renderer.Types.rect 13.0 4.0 1.0 8.0, p.Border)
          FillRect(Renderer.Types.rect 7.0 4.0 1.0 8.0, p.Border) ])

let private cutIcon (p: ChromePalette) =
    // Two circular handles (drawn as 2x2 squares) with crossing blades.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 3.0 2.0 2.0, p.Text)
          FillRect(Renderer.Types.rect 3.0 11.0 2.0 2.0, p.Text)
          FillRect(Renderer.Types.rect 5.0 5.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 6.0 6.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 7.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 8.0 8.0 2.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 10.0 9.0 4.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 5.0 10.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 6.0 9.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 8.0 1.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 10.0 7.0 4.0 1.0, p.Text) ])

let private copyIcon (p: ChromePalette) =
    // Two overlapping document rectangles.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 3.0 7.0 9.0, p.Input)
          FillRect(Renderer.Types.rect 3.0 3.0 7.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 11.0 7.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 3.0 1.0 9.0, p.Border)
          FillRect(Renderer.Types.rect 9.0 3.0 1.0 9.0, p.Border)
          FillRect(Renderer.Types.rect 6.0 5.0 7.0 9.0, p.Input)
          FillRect(Renderer.Types.rect 6.0 5.0 7.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 6.0 13.0 7.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 6.0 5.0 1.0 9.0, p.Border)
          FillRect(Renderer.Types.rect 12.0 5.0 1.0 9.0, p.Border) ])

let private pasteIcon (p: ChromePalette) =
    // Clipboard outline with a tab at top and a sheet inside.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 4.0 10.0 10.0, p.Input)
          FillRect(Renderer.Types.rect 3.0 4.0 10.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 13.0 10.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 3.0 4.0 1.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 12.0 4.0 1.0 10.0, p.Border)
          FillRect(Renderer.Types.rect 6.0 2.0 4.0 3.0, p.Text)
          FillRect(Renderer.Types.rect 5.0 7.0 6.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 5.0 9.0 6.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 5.0 11.0 4.0 1.0, p.ChromeDark) ])

let private borderIcon (p: ChromePalette) =
    // 2x2 grid of cells.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 2.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 13.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 2.0 1.0 12.0, p.Text)
          FillRect(Renderer.Types.rect 13.0 2.0 1.0 12.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 7.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 2.0 1.0 12.0, p.Text) ])

let private freezeIcon (p: ChromePalette) =
    // Cell grid with a thicker dividing line marking the frozen split.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 2.0 12.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 2.0 13.0 12.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 2.0 2.0 1.0 12.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 13.0 2.0 1.0 12.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 2.0 8.0 1.0 6.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 13.0 8.0 1.0 6.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 8.0 2.0 1.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 8.0 13.0 1.0 1.0, p.ChromeDark)
          FillRect(Renderer.Types.rect 2.0 7.0 12.0 2.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 2.0 2.0 6.0, p.Text) ])

let private splitIcon (p: ChromePalette) =
    // Four panes with a cross divider through the middle.
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 2.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 13.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 2.0 1.0 12.0, p.Text)
          FillRect(Renderer.Types.rect 13.0 2.0 1.0 12.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 2.0 2.0 12.0, p.Text)
          FillRect(Renderer.Types.rect 2.0 7.0 12.0 2.0, p.Text) ])

let private highlightIcon (p: ChromePalette) =
    // Solid yellow swipe under a baseline.
    let yellow = Color.rgb 250uy 215uy 75uy
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 2.0 9.0 12.0 3.0, yellow)
          FillRect(Renderer.Types.rect 2.0 12.0 12.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 4.0 4.0 8.0 4.0, p.Input)
          FillRect(Renderer.Types.rect 4.0 4.0 8.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 4.0 8.0 8.0 1.0, p.Border)
          FillRect(Renderer.Types.rect 4.0 4.0 1.0 4.0, p.Border)
          FillRect(Renderer.Types.rect 11.0 4.0 1.0 4.0, p.Border) ])

let private filterClearIcon (p: ChromePalette) =
    // Funnel shape crossed by a diagonal X.
    let red = Color.rgb 200uy 65uy 65uy
    miniIcon (fun _ _ ->
        [ FillRect(Renderer.Types.rect 3.0 3.0 10.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 5.0 6.0 6.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 9.0 2.0 1.0, p.Text)
          FillRect(Renderer.Types.rect 7.0 11.0 2.0 1.0, p.Text)
          for i in 0..6 do FillRect(Renderer.Types.rect (3.0 + float i) (4.0 + float i) 1.0 1.0, red)
          for i in 0..6 do FillRect(Renderer.Types.rect (9.0 - float i) (4.0 + float i) 1.0 1.0, red) ])

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
    | "filter-clear" -> filterClearIcon p
    | "chart" -> chartIcon p
    | "pivot" -> pivotIcon p
    | "merge" -> mergeIcon p
    | "unmerge" -> unmergeIcon p
    | "sort-asc" -> toolbarTextIcon p "A-Z"
    | "sort-desc" -> toolbarTextIcon p "Z-A"
    | "format-number" -> toolbarTextIcon p "1.2"
    | "format-currency" -> toolbarTextIcon p "$"
    | "format-percent" -> toolbarTextIcon p "%"
    | "format-date" -> toolbarTextIcon p "Dt"
    | "decimal-plus" -> toolbarTextIcon p ".0+"
    | "decimal-minus" -> toolbarTextIcon p ".0-"
    | "cut" -> cutIcon p
    | "copy" -> copyIcon p
    | "paste" -> pasteIcon p
    | "undo" -> Icon.arrowLeft 13.0 p.Text
    | "redo" -> Icon.arrowRight 13.0 p.Text
    | "border" -> borderIcon p
    | "freeze" -> freezeIcon p
    | "split" -> splitIcon p
    | "sheet-add" -> toolbarTextIcon p "+S"
    | "sheet-delete" -> toolbarTextIcon p "-S"
    | "sheet-rename" -> toolbarTextIcon p "Sr"
    | "text-color" -> toolbarTextIcon p "A"
    | "highlight-color" -> highlightIcon p
    | "all-caps" -> toolbarTextIcon p "AA"
    | "small-caps" -> toolbarTextIcon p "Aa"
    | _ -> toolbarTextIcon p label

let private menuButton (p: ChromePalette) (renderState: ShellRenderState) (id: string) label isOpen action =
    withHover renderState id (
        clickable action (
            stack [
                menuButtonChrome p renderState id isOpen
                padXY 12.0 6.0 (textStyled Font.sm p.Text label)
            ]))

let private comboControl
        (p: ChromePalette) (renderState: ShellRenderState) (id: string)
        (label: string) (action: unit -> unit) =
    // Matches widgets/Widgets.fs `select` styling: 28-30px tall, rounded
    // hairline border, label + chevron on the right. Reads its own hover
    // and open state from `renderState` so the trigger tracks the same
    // pressed/hot/idle states as `flatToolbarCell` and `menuButton`.
    let surface =
        dynamic (fun () ->
            let isOpen = renderState.OpenDropdownTrigger.Value = Some id
            let isHover = renderState.HoveredId.Value = Some id
            let fill =
                if isOpen then p.Pressed
                elif isHover then p.HotTop
                else p.Input
            rounded 4.0 fill empty)
    withHover renderState id (
        clickable action (
            sizedH 28.0 (
                stack [
                    surface
                    border p.Border 1.0 4.0 (
                        padXY 10.0 0.0 (
                            rowAlign MSpaceBetween CCenter 0.0 [
                                textStyled Font.sm p.Text label
                                padXY 6.0 0.0 (textStyled Font.sm (parseHexColor "#5F6368") "▾")
                            ]))
                ])))

/// Standard 8x5 Word-style color palette. Apps may swap this for a richer
/// picker later.
let standardColors : (string * string) list = [
    "Auto",        "#000000"
    "Black",       "#000000"
    "White",       "#FFFFFF"
    "Light Grey",  "#D9D9D9"
    "Grey",        "#808080"
    "Dark Grey",   "#404040"
    "Red",         "#C00000"
    "Orange",      "#ED7D31"
    "Yellow",      "#FFC000"
    "Light Yellow","#FFF2CC"
    "Green",       "#00B050"
    "Teal",        "#00B0F0"
    "Blue",        "#0070C0"
    "Dark Blue",   "#002060"
    "Purple",      "#7030A0"
    "Pink",        "#FFB6C1"
]

let private colorControl
        (p: ChromePalette) (renderState: ShellRenderState) (id: string)
        (value: string) (action: unit -> unit) =
    let surface =
        dynamic (fun () ->
            let isOpen = renderState.OpenDropdownTrigger.Value = Some id
            let isHover = renderState.HoveredId.Value = Some id
            let fill =
                if isOpen then p.Pressed
                elif isHover then p.HotTop
                else p.Input
            box fill empty)
    withHover renderState id (
        clickable action (
            stack [
                surface
                border p.Border 1.0 4.0 (
                    padXY 4.0 4.0 (
                        sized 20.0 16.0 (box (parseHexColor value) empty)))
            ]))

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
    ShowToolbarPressed: string -> bool -> bool
    /// Vertical workspace widget — the document canvas, grid, etc. Sized by
    /// the caller to fit within the window body.
    Workspace: Widget
}

/// Thin vertical separator between toolbar groups.
let private toolbarSeparator (p: ChromePalette) =
    sized 1.0 18.0 (box p.ChromeMid empty)

let private renderToolbarControl (p: ChromePalette) (renderState: ShellRenderState) (config: ShellConfig) = function
    | ToolbarButton spec ->
        let isActive = spec.IsActive |> Option.map (fun active -> active()) |> Option.defaultValue false
        let isEnabled = spec.IsEnabled |> Option.map (fun en -> en()) |> Option.defaultValue true
        let pressed = config.ShowToolbarPressed spec.TestId isActive
        let icon = renderGlyph p spec.Label (config.ToolbarGlyph spec.TestId)
        flatToolbarCell p renderState spec.TestId pressed isEnabled icon spec.Action
    | ToolbarSelect spec ->
        let label =
            spec.Options()
            |> List.tryFind (fun option -> option.Value = spec.Value())
            |> Option.map _.Label
            |> Option.defaultValue (spec.Value())
        // Open a dropdown of the available choices anchored to this combo
        // trigger's recorded bounds. `popoverBelow` (in the shell-wide
        // dropdown overlay) reads the anchor that `anchorRegion` writes
        // below — no manual coordinate math.
        let openDropdown () =
            let items =
                spec.Options()
                |> List.map (fun choice ->
                    menuItem
                        (sprintf "%s-%s" spec.TestId choice.Value)
                        choice.Label
                        (Some (fun () -> choice.Value = spec.Value()))
                        (fun () -> spec.Action(choice.Value)))
            openAnchoredDropdown renderState spec.TestId items
        anchorRegion (anchorFor renderState spec.TestId)
            (comboControl p renderState spec.TestId label openDropdown)
    | ToolbarInput spec ->
        match spec.Kind with
        | NumberInput ->
            comboControl p renderState spec.TestId (spec.Value()) (fun () -> spec.Action(spec.Value()))
        | ColorInput ->
            let openSwatches () =
                let items =
                    standardColors
                    |> List.map (fun (name, hex) ->
                        menuItem
                            (sprintf "%s-%s" spec.TestId hex)
                            (sprintf "● %s" name)
                            (Some (fun () -> spec.Value() = hex))
                            (fun () -> spec.Action(hex)))
                openAnchoredDropdown renderState spec.TestId items
            anchorRegion (anchorFor renderState spec.TestId)
                (colorControl p renderState spec.TestId (spec.Value()) openSwatches)
        | TextInput ->
            comboControl p renderState spec.TestId (spec.Value()) (fun () -> spec.Action(spec.Value()))
    | ToolbarSeparator -> padXY 4.0 0.0 (toolbarSeparator p)

let private renderToolbarRow (p: ChromePalette) (renderState: ShellRenderState) (config: ShellConfig) (row: ToolbarRowSpec) =
    sizedH 32.0 (
        padXY 4.0 2.0 (
            rowAlign MStart CCenter 2.0 (row.Controls |> List.map (renderToolbarControl p renderState config))))

let private renderMenuButton (p: ChromePalette) (renderState: ShellRenderState) (spec: MenuSpec) =
    let isOpen = renderState.OpenMenu.Value = Some spec.TestId
    let toggle () =
        if isOpen then renderState.OpenMenu.Value <- None
        else renderState.OpenMenu.Value <- Some spec.TestId
    // The anchor records this button's resolved bounds so the dropdown
    // panel can attach to it via `popoverBelow` — no manual x math.
    anchorRegion (anchorFor renderState spec.TestId)
        (menuButton p renderState spec.TestId spec.Label isOpen toggle)

let private renderMenuPanel (p: ChromePalette) (renderState: ShellRenderState) (spec: MenuSpec) =
    let gutterWidth = p.MenuGutterWidth
    let itemHeight = 34.0

    let renderItem (item: MenuItemSpec) =
        let isChecked = item.IsChecked |> Option.map (fun check -> check()) |> Option.defaultValue false
        let isEnabled = item.IsEnabled |> Option.map (fun en -> en()) |> Option.defaultValue true
        let activate () =
            if isEnabled then
                item.Action ()
                renderState.OpenMenu.Value <- None
        // Row backgrounds, by priority: hover > checked > idle. The checked
        // tint makes the active option in radio-style groups (paragraph
        // style, font family) visible at a glance instead of the gutter
        // check being the only signal.
        let bg =
            dynamic (fun () ->
                if not isEnabled then box p.MenuPanel empty
                elif renderState.HoveredId.Value = Some item.TestId then box p.HotTop empty
                elif isChecked then box p.Pressed empty
                else box p.MenuPanel empty)
        let labelColor = if isEnabled then p.Text else p.ChromeDark
        let shortcutColor = if isEnabled then parseHexColor "#5F6368" else p.ChromeDark
        let shortcutWidget =
            match item.Shortcut with
            | Some s -> padXY 12.0 0.0 (textStyled Font.sm shortcutColor s)
            | None -> empty
        let checkIcon = if isChecked && isEnabled then Icon.check 12.0 p.Text else empty
        let row =
            sized p.MenuWidth itemHeight (
                stack [
                    bg
                    rowAlign MStart CCenter 0.0 [
                        sizedW gutterWidth (center checkIcon)
                        padXY 4.0 0.0 (textStyled Font.sm labelColor item.Label)
                        expanded empty
                        shortcutWidget
                    ]
                ])
        if isEnabled then
            withHover renderState item.TestId (clickable activate row)
        else
            row

    let menuContents = spec.Items |> List.map renderItem
    let panelHeight = float spec.Items.Length * itemHeight + 8.0

    filtered [DropShadow(0.0, 4.0, 12, p.Shadow)] (
        rounded 4.0 p.MenuPanel (
            border p.Border 1.0 4.0 (
                pad 4.0 (columnSpaced 0.0 menuContents))))

// ── Dialog renderer ───────────────────────────────────────────

let private dialogFieldWidth = 360.0

let private renderDialogField (p: ChromePalette) (renderState: ShellRenderState) (state: OpenDialogState) (field: DialogField) : Widget =
    let valueSignal = state.Values.[field.Id]
    let labelColor = parseHexColor "#5F6368"
    let accent = parseHexColor "#1A73E8"
    let inputHeight = 36.0
    let textInputWithId (placeholder: string) =
        // Allocate-once FocusId stored in the OpenDialogState — keeps the
        // input focused across renders. TextInput renders its own border,
        // padding, and focus state so we just size it.
        let fid = state.FocusIds.[field.Id]
        sized dialogFieldWidth inputHeight (
            Composit.Ui.Widget.TextInput(createTextInputState valueSignal fid placeholder 12.5))
    let labelLine = textStyled Font.sm labelColor field.Label
    let inputBox =
        match field.Kind with
        | DialogTextField placeholder -> textInputWithId placeholder
        | DialogNumberField(_, _) -> textInputWithId ""
        | DialogTextAreaField(placeholder, lines) ->
            // Multi-line text area. Same TextInput widget but sized to N
            // rows so wrapped content has room. Lines * 18 px is the
            // line-height the TextInput uses internally at fontSize 12.5.
            let fid = state.FocusIds.[field.Id]
            sized dialogFieldWidth (float lines * 18.0 + 16.0) (
                Composit.Ui.Widget.TextInput(createTextInputState valueSignal fid placeholder 12.5))
        | DialogCheckboxField ->
            let toggle () =
                valueSignal.Value <- if valueSignal.Value = "true" then "false" else "true"
            clickable toggle (
                sizedH inputHeight (
                    rowAlign MStart CCenter 10.0 [
                        dynamic (fun () ->
                            let isChecked = valueSignal.Value = "true"
                            sized 18.0 18.0 (
                                stack [
                                    rounded 3.0 (if isChecked then accent else p.Input) empty
                                    border (if isChecked then accent else p.Border) 1.5 3.0 (
                                        if isChecked then center (Icon.check 12.0 Palette.white)
                                        else empty)
                                ]))
                        textStyled Font.sm p.Text field.Label
                    ]))
        | DialogSelectField options ->
            let cycle () =
                let current = valueSignal.Value
                let idx = options |> List.tryFindIndex (fun o -> o.Value = current) |> Option.defaultValue -1
                let next = (idx + 1) % options.Length
                valueSignal.Value <- options.[next].Value
            clickable cycle (
                sized dialogFieldWidth inputHeight (
                    stack [
                        rounded 4.0 p.Input empty
                        border p.Border 1.0 4.0 (
                            padXY 12.0 0.0 (
                                dynamic (fun () ->
                                    let label =
                                        options
                                        |> List.tryFind (fun o -> o.Value = valueSignal.Value)
                                        |> Option.map (fun o -> o.Label)
                                        |> Option.defaultValue valueSignal.Value
                                    rowAlign MSpaceBetween CCenter 0.0 [
                                        textStyled Font.sm p.Text label
                                        textStyled Font.sm labelColor "▾"
                                    ])))
                    ]))
        | DialogStatusField ->
            dynamic (fun () ->
                let v = valueSignal.Value
                if v = "" then empty
                else textStyled Font.sm labelColor v)
        | DialogSectionHeader ->
            // Bold section title with a 1px hairline beneath it. Not
            // interactive; lets long forms breathe with named groups.
            sizedH 28.0 (
                stack [
                    canvas dialogFieldWidth 28.0 (fun w h ->
                        [ Renderer.FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, p.ChromeMid) ])
                    rowAlign MStart CCenter 0.0 [
                        textBoldStyled Font.sm p.Text field.Label
                    ]
                ])
        | DialogColorField ->
            // Color swatch: 28×28 chip with a hairline border. Clicking
            // opens the standard 16-color picker anchored to the swatch.
            let openPicker () =
                let items =
                    standardColors
                    |> List.map (fun (name, hex) ->
                        menuItem
                            (sprintf "%s-%s" field.Id hex)
                            (sprintf "● %s" name)
                            (Some (fun () -> valueSignal.Value = hex))
                            (fun () -> valueSignal.Value <- hex))
                openAnchoredDropdown renderState field.Id items
            anchorRegion (anchorFor renderState field.Id) (
                withHover renderState field.Id (
                    clickable openPicker (
                        sizedH inputHeight (
                            rowAlign MStart CCenter 12.0 [
                                dynamic (fun () ->
                                    sized 28.0 28.0 (
                                        stack [
                                            rounded 4.0 (parseHexColor valueSignal.Value) empty
                                            border p.Border 1.0 4.0 empty
                                        ]))
                                dynamic (fun () -> textStyled Font.sm labelColor valueSignal.Value)
                            ]))))
        | DialogSliderField(minV, maxV) ->
            // Click-anywhere-on-the-track to step through the range. The
            // step is (max-min)/10 and wraps to min when it would
            // overshoot — same behavior as the standalone widgets slider.
            // Drag isn't wired yet; clicking the track is enough for
            // most settings (zoom %, autosave interval, etc.).
            let trackWidth = dialogFieldWidth - 60.0
            sizedH inputHeight (
                rowAlign MStart CCenter 12.0 [
                    dynamic (fun () ->
                        let v =
                            match System.Double.TryParse(valueSignal.Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
                            | true, parsed -> max minV (min maxV parsed)
                            | _ -> minV
                        let pct = (v - minV) / (max 0.0001 (maxV - minV))
                        let thumbX = max 8.0 (min (trackWidth - 8.0) (trackWidth * pct))
                        let step = (maxV - minV) / 10.0
                        let advance () =
                            let next = if v + step > maxV then minV else v + step
                            valueSignal.Value <- sprintf "%g" next
                        clickable advance (
                            sized trackWidth 24.0 (
                                canvas trackWidth 24.0 (fun w _ ->
                                    [ Renderer.FillRoundedRect(Renderer.Types.rect 0.0 10.0 w 4.0, LanguagePrimitives.FloatWithMeasure 2.0, p.Border)
                                      Renderer.FillRoundedRect(Renderer.Types.rect 0.0 10.0 (w * pct) 4.0, LanguagePrimitives.FloatWithMeasure 2.0, accent)
                                      Renderer.FillCircle(Renderer.Types.vec2 thumbX 12.0, LanguagePrimitives.FloatWithMeasure 9.0, Color.rgba 0uy 0uy 0uy 18uy)
                                      Renderer.FillCircle(Renderer.Types.vec2 thumbX 11.0, LanguagePrimitives.FloatWithMeasure 8.0, Palette.white)
                                      Renderer.FillCircle(Renderer.Types.vec2 thumbX 11.0, LanguagePrimitives.FloatWithMeasure 4.0, accent) ]))))
                    dynamic (fun () ->
                        // Right-side numeric readout so the user knows
                        // the current value at a glance.
                        textStyled Font.sm p.Text valueSignal.Value)
                ])
        | DialogRadioField options ->
            // Column of radio buttons. Outer circle with `accent` border
            // when selected, plain border otherwise; inner dot when the
            // option matches valueSignal.
            columnSpaced 8.0 [
                for opt in options ->
                    let isSelected () = valueSignal.Value = opt.Value
                    let select () = valueSignal.Value <- opt.Value
                    clickable select (
                        sizedH 24.0 (
                            rowAlign MStart CCenter 10.0 [
                                dynamic (fun () ->
                                    let sel = isSelected ()
                                    sized 16.0 16.0 (
                                        stack [
                                            rounded 8.0 p.Input empty
                                            border (if sel then accent else p.Border) 1.5 8.0 (
                                                if sel then
                                                    center (
                                                        sized 6.0 6.0 (rounded 3.0 accent empty))
                                                else empty)
                                        ]))
                                textStyled Font.sm p.Text opt.Label
                            ]))
            ]

    match field.Kind with
    | DialogCheckboxField | DialogStatusField | DialogSectionHeader -> inputBox
    | _ -> columnSpaced 6.0 [ labelLine; inputBox ]

let private renderDialogButton (p: ChromePalette) (renderState: ShellRenderState) (values: System.Collections.Generic.Dictionary<string, Signal<string>>) (btn: DialogButton) : Widget =
    let activate () =
        let snapshot = values |> Seq.map (fun kv -> kv.Key, kv.Value.Value) |> Map.ofSeq
        let shouldClose = btn.Action snapshot
        if shouldClose then renderState.OpenDialog.Value <- None
    let textColor = if btn.IsPrimary then Palette.white else p.Text
    let bgColor = if btn.IsPrimary then parseHexColor "#1A73E8" else Palette.transparent
    let borderColor = if btn.IsPrimary then parseHexColor "#1A73E8" else p.Border
    // Buttons are ~natural-width — `padXY 18 0` gives the text its breathing
    // room while the sizedH pins the row height.
    clickable activate (
        sizedH 36.0 (
            stack [
                rounded 4.0 bgColor empty
                border borderColor 1.0 4.0 (
                    padXY 18.0 0.0 (
                        center (textBoldStyled Font.sm textColor btn.Label)))
            ]))

let private renderDialog (p: ChromePalette) (renderState: ShellRenderState) (state: OpenDialogState) : Widget =
    let spec = state.Spec
    let title = textBoldStyled Font.xl p.Text spec.Title
    let subtitle =
        match spec.Subtitle with
        | Some s -> textStyled Font.sm (parseHexColor "#5F6368") s
        | None -> empty
    let header =
        match spec.Subtitle with
        | Some _ -> columnSpaced 6.0 [ title; subtitle ]
        | None -> title

    // Tab bar — only rendered when the spec actually has tabs. Each tab
    // is a natural-width clickable label whose footprint is just its
    // text + horizontal padding; an active tab adds a 2 px accent strip
    // along the bottom edge of its own column. Building it as
    // `column [pad text; underline]` (rather than a fixed-width canvas
    // overlay) lets the row of tabs auto-fit the panel — a 5-tab dialog
    // like Format Cells no longer overflows.
    let accent = parseHexColor "#1A73E8"
    let renderTab (idx: int) (label: string) =
        let tabId = sprintf "%s-tab-%d" spec.TestId idx
        dynamic (fun () ->
            let isActive = state.ActiveTab.Value = idx
            let labelColor = if isActive then p.Text else parseHexColor "#5F6368"
            let weight = if isActive then textBoldStyled else textStyled
            let underlineColor = if isActive then accent else Palette.transparent
            let hoverBg =
                dynamic (fun () ->
                    if renderState.HoveredId.Value = Some tabId && not isActive
                    then box p.HotTop empty
                    else empty)
            withHover renderState tabId (
                clickable (fun () -> state.ActiveTab.Value <- idx) (
                    stack [
                        hoverBg
                        column [
                            padXY 14.0 8.0 (
                                rowAlign MStart CCenter 0.0 [ weight Font.sm labelColor label ])
                            sizedH 2.0 (box underlineColor empty)
                        ]
                    ])))
    let tabBar =
        if List.isEmpty spec.Tabs then empty
        else
            stack [
                canvas (dialogFieldWidth + 72.0) 34.0 (fun w h ->
                    [ Renderer.FillRect(Renderer.Types.rect 0.0 (h - 1.0) w 1.0, p.ChromeMid) ])
                rowAlign MStart CCenter 4.0 (spec.Tabs |> List.mapi (fun i (label, _) -> renderTab i label))
            ]

    // Field list for the current view: when there are tabs, use the
    // active tab's fields; otherwise fall back to Fields.
    let renderFields fields =
        fields
        |> List.map (renderDialogField p renderState state)
        |> columnSpaced 18.0
    let fields =
        if List.isEmpty spec.Tabs then renderFields spec.Fields
        else
            dynamic (fun () ->
                let idx = max 0 (min (spec.Tabs.Length - 1) state.ActiveTab.Value)
                let _, tabFields = spec.Tabs.[idx]
                renderFields tabFields)

    let buttons =
        spec.Buttons
        |> List.map (renderDialogButton p renderState state.Values)
        |> rowAlign MEnd CCenter 8.0

    // Body of the dialog. Wrap header / tabs / fields / buttons in their
    // own groups with generous vertical breathing room.
    let body =
        columnSpaced 20.0 [
            header
            if not (List.isEmpty spec.Tabs) then tabBar
            fields
            buttons
        ]

    let panel =
        sizedW (dialogFieldWidth + 72.0) (
            filtered [DropShadow(0.0, 10.0, 30, Color.rgba 0uy 0uy 0uy 70uy)] (
                rounded 10.0 p.MenuPanel (
                    border p.Border 1.0 10.0 (
                        pad 28.0 body))))
    center panel

/// Assemble the chrome around the caller's workspace widget. Layout:
///
///   ┌─────────────────────────────────┐
///   │  Menu bar          (32, white)  │
///   ├─────────────────────────────────┤  ← hairline
///   │  Toolbar (n rows)  (32 each)    │
///   ├─────────────────────────────────┤  ← hairline
///   │  Workspace                       │
///   ├─────────────────────────────────┤  ← hairline
///   │  Status bar        (22, white)  │
///   └─────────────────────────────────┘
///
/// Render a dialog spec on its own (no surrounding shell) — for headless
/// preview / screenshot testing of a `DialogSpec`. Opens it on a fresh state
/// and returns just the dialog widget.
let dialogPreview (chromeSpec: ChromeSpec) (spec: DialogSpec) : Widget =
    let state = createState ()
    openDialog state spec
    match state.OpenDialog.Value with
    | Some ds -> renderDialog (palette chromeSpec) state ds
    | None -> empty

/// The OS window decoration carries the document title; there's no
/// in-app title strip.
let view (renderState: ShellRenderState) (config: ShellConfig) =
    let p = palette config.Chrome
    let menuHeight  = 32.0
    let toolbarRowHeight = 32.0
    let statusHeight = 22.0

    let openMenu () =
        renderState.OpenMenu.Value
        |> Option.bind (fun testId -> config.Menus |> List.tryFind (fun menu -> menu.TestId = testId))

    let closeMenu () = renderState.OpenMenu.Value <- None

    let menuOverlay =
        dynamic (fun () ->
            match openMenu () with
            | Some menu ->
                stack [
                    modalScrim None (Some closeMenu)
                    popoverBelow (anchorFor renderState menu.TestId)
                        (sizedW p.MenuWidth (renderMenuPanel p renderState menu))
                ]
            | None -> empty)

    let dialogOverlay =
        dynamic (fun () ->
            match renderState.OpenDialog.Value with
            | Some state ->
                stack [
                    // Half-transparent dark backdrop; absorbs clicks so the
                    // workspace doesn't react while a dialog is up. The
                    // user dismisses via Cancel/OK only.
                    modalScrim (Some (Color.rgba 0uy 0uy 0uy 64uy)) None
                    overlay (renderDialog p renderState state)
                ]
            | None -> empty)

    let closeCtxMenu () =
        renderState.OpenContextMenu.Value <- None

    let renderPanelItem (closeFn: unit -> unit) (item: MenuItemSpec) =
        let isChecked = item.IsChecked |> Option.map (fun check -> check()) |> Option.defaultValue false
        let isEnabled = item.IsEnabled |> Option.map (fun en -> en()) |> Option.defaultValue true
        let activate () =
            if isEnabled then
                item.Action ()
                closeFn ()
        let bg =
            dynamic (fun () ->
                if not isEnabled then box p.MenuPanel empty
                elif renderState.HoveredId.Value = Some item.TestId then box p.HotTop empty
                elif isChecked then box p.Pressed empty
                else box p.MenuPanel empty)
        let labelColor = if isEnabled then p.Text else p.ChromeDark
        let checkIcon = if isChecked && isEnabled then Icon.check 12.0 p.Text else empty
        let row =
            sizedH 28.0 (
                stack [
                    bg
                    rowAlign MStart CCenter 0.0 [
                        sizedW p.MenuGutterWidth (center checkIcon)
                        padXY 8.0 0.0 (textStyled Font.sm labelColor item.Label)
                    ]
                ])
        if isEnabled then
            withHover renderState item.TestId (clickable activate row)
        else
            row

    let panelFrame (items: Widget list) =
        filtered [DropShadow(0.0, 4.0, 12, p.Shadow)] (
            rounded 4.0 p.MenuPanel (
                border p.Border 1.0 4.0 (
                    pad 4.0 (columnSpaced 0.0 items))))

    let renderContextMenuPanel (state: OpenContextMenuState) =
        panelFrame (state.Items |> List.map (renderPanelItem closeCtxMenu))

    let contextMenuOverlay =
        dynamic (fun () ->
            match renderState.OpenContextMenu.Value with
            | Some state ->
                stack [
                    modalScrim None (Some closeCtxMenu)
                    popoverAt state.X state.Y
                        (sizedW p.MenuWidth (renderContextMenuPanel state))
                ]
            | None -> empty)

    // Combo / color-picker dropdowns reuse the same panel rendering as
    // right-click context menus, but anchor to a sibling widget's bounds
    // via `popoverBelow` rather than the cursor position.
    let anchoredDropdownOverlay =
        dynamic (fun () ->
            match renderState.OpenAnchoredDropdown.Value with
            | Some state ->
                let dismiss () = closeAnchoredDropdown renderState
                let panel = panelFrame (state.Items |> List.map (renderPanelItem dismiss))
                stack [
                    modalScrim None (Some dismiss)
                    popoverBelow state.Anchor (sizedW p.MenuWidth panel)
                ]
            | None -> empty)

    let menuBar =
        dynamic (fun () ->
            chromeBandWithBorder p p.ChromeTop menuHeight (
                sizedH menuHeight (
                    padXY 6.0 0.0 (
                        rowAlign MStart CCenter 2.0 (config.Menus |> List.map (renderMenuButton p renderState))))))

    let toolbarRowCount = max 1 (List.length config.ToolbarRows)
    let toolbarHeight = float toolbarRowCount * toolbarRowHeight
    let toolbarBand =
        chromeBandWithBorder p p.ChromeTop toolbarHeight (
            column (config.ToolbarRows |> List.map (renderToolbarRow p renderState config)))
    // Precompute id → tooltip lookup so the per-frame status-bar path is
    // O(1) instead of scanning every toolbar row on each tick.
    let tooltipFor : string -> string option =
        let table = System.Collections.Generic.Dictionary<string, string>()
        for row in config.ToolbarRows do
            for control in row.Controls do
                match control with
                | ToolbarButton spec ->
                    match spec.Tooltip with
                    | Some t -> table.[spec.TestId] <- t
                    | None -> ()
                | _ -> ()
        fun id -> match table.TryGetValue id with true, t -> Some t | _ -> None
    let statusBar =
        chromeBandWithBorder p p.ChromeTop statusHeight (
            sizedH statusHeight (
                dynamic (fun () ->
                    // After a 500 ms hover settles on a toolbar button with
                    // a Tooltip, the status bar surfaces it. Otherwise the
                    // host's StatusText() wins. This avoids inventing a
                    // floating tooltip layer and matches how classic Word
                    // surfaces action descriptions in its status line.
                    let label =
                        match renderState.HoverStart.Value with
                        | Some (id, t) when Platform.now() - t >= tooltipDelaySec ->
                            tooltipFor id |> Option.defaultValue (config.StatusText())
                        | _ -> config.StatusText()
                    padXY 12.0 4.0 (textStyled Font.sm p.Text label))))
    let windowBody =
        column [
            menuBar
            toolbarBand
            expanded config.Workspace
            statusBar
        ]

    sizedW p.WindowWidth (
        stack [
            windowBody
            menuOverlay
            anchoredDropdownOverlay
            contextMenuOverlay
            dialogOverlay
        ])
