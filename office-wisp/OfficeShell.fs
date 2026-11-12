module Wisp.OfficeShell

// The Wisp backend for the Office UI abstraction. Renders Rbt.Office.Ui specs (menus, toolbars,
// dialogs) to Wisp `Element`s, atrium-themed — the `wisp + fableui` replacement for the Composit
// `OfficeShell`. The specs and the apps are untouched; this is just a new renderer behind the seam.
// FableUI is the same SolidJS-style model as Wisp, so this is written in the natural signal/effect
// style rather than the old rebuild-the-tree way.

open System.Collections.Generic
open Rbt.Renderer
open Wisp.Reactive
open Wisp.Layout
open Rbt.Office.Ui

module Ui = Wisp.Atrium

// ── shell state (Wisp-native: the open menu + the active dialog + its field signals) ──
type ShellState =
    { OpenMenu: Signal<int>
      Dialog: RefSignal<DialogSpec option>
      ActiveTab: Signal<int>
      Fields: Dictionary<string, Signal<string>> }

let create () : ShellState =
    { OpenMenu = signal -1; Dialog = refSignal None; ActiveTab = signal 0; Fields = Dictionary() }

let openDialog (st: ShellState) (spec: DialogSpec) =
    st.Fields.Clear()
    st.ActiveTab.Value <- 0
    for f in spec.Fields @ (spec.Tabs |> List.collect snd) do st.Fields.[f.Id] <- signal f.Initial
    st.Dialog.Value <- Some spec

let closeDialog (st: ShellState) =
    // Removed text fields must not keep consuming the editor's keyboard input.
    Focus.blur()
    st.Dialog.Value <- None

let private fieldSig (st: ShellState) (f: DialogField) =
    match st.Fields.TryGetValue f.Id with
    | true, s -> s
    | _ -> let s = signal f.Initial in st.Fields.[f.Id] <- s; s

let private currentMap (st: ShellState) : Map<string, string> =
    st.Fields |> Seq.map (fun kv -> kv.Key, kv.Value.Value) |> Map.ofSeq

let private parseHex (s: string) : Color =
    let h = s.TrimStart('#')
    try
        if h.Length >= 6 then Color.rgb (System.Convert.ToByte(h.Substring(0, 2), 16)) (System.Convert.ToByte(h.Substring(2, 2), 16)) (System.Convert.ToByte(h.Substring(4, 2), 16))
        else Ui.P.mute
    with _ -> Ui.P.mute

// ── menus: convert the Office spec to the atrium menu and reuse Ui.menuBar/menuOverlay ──
let private toMenu (m: MenuSpec) : Ui.Menu =
    { Title = m.Label
      Items =
        m.Items
        |> List.map (fun it ->
            let enabled () = it.IsEnabled |> Option.map (fun f -> f ()) |> Option.defaultValue true
            { Ui.Label = it.Label
              Ui.Action = (fun () -> if enabled () then it.Action ())
              Ui.Checked = it.IsChecked
              Ui.Shortcut = it.Shortcut }) }

// ── toolbar controls ──
let private tbButton (b: ButtonSpec) : Element =
    let enabled () = b.IsEnabled |> Option.map (fun f -> f ()) |> Option.defaultValue true
    let active () = b.IsActive |> Option.map (fun f -> f ()) |> Option.defaultValue false
    // active-state in the paint (re-read each frame) rather than `dyn` — a dyn whose key reads no
    // signal (an always-inactive button) isn't built during a one-shot render.
    clickable (fun () -> if enabled () then b.Action ())
        (Ui.surfaceBox
            (fun x y w h -> if active () then [ FillRoundedRect(rect x y w h, Ui.radius * 1.0<px>, Ui.P.sel) ] else [])
            (padXY 9.0 5.0 (text 12.5 Ui.P.text b.Label)))

let private chevronDown : Element =
    canvas 10.0 12.0 (fun x y _ _ -> [ Group(Matrix3x2.translate x y, 1.0, Ui.Glyph.down 10.0 Ui.P.dim) ])

let private tbSelect (s: SelectSpec) : Element =
    let cycle () =
        let opts = s.Options () |> List.filter (fun item -> item.IsEnabled)
        if not opts.IsEmpty then
            let i = opts |> List.tryFindIndex (fun o -> o.Value = s.Value ()) |> Option.defaultValue -1
            s.Action opts.[(i + 1) % opts.Length].Value
    let label () =
        let v = s.Value ()
        s.Options () |> List.tryFind (fun o -> o.Value = v) |> Option.map (fun o -> o.Label) |> Option.defaultValue v
    clickable cycle
        (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Control x y w h Ui.radius)
            (padXY 8.0 4.0 (rowA MStart CCenter 6.0 [ sizedBox 84.0 16.0 (textf 12.0 Ui.P.text false label); chevronDown ])))

let private tbInput (i: InputSpec) : Element =
    match i.Kind with
    | ColorInput ->
        Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Well x y w h Ui.radius)
            (padXY 4.0 4.0 (canvas 24.0 12.0 (fun x y _ _ -> [ FillRoundedRect(rect (x + 1.0) (y + 1.0) 22.0 10.0, 2.0<px>, parseHex (i.Value ())) ])))
    | _ ->
        sizedBox 60.0 24.0
            (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Well x y w h Ui.radius)
                (padXY 8.0 4.0 (textf 12.0 Ui.P.text false (fun () -> i.Value ()))))

let private tbControl = function
    | ToolbarButton b -> tbButton b
    | ToolbarSelect s -> tbSelect s
    | ToolbarInput i -> tbInput i
    | ToolbarSeparator -> canvas 9.0 24.0 (fun x y _ _ -> [ FillRect(rect (x + 4.0) (y + 4.0) 1.0 16.0, Ui.P.border) ])

let toolbar (rows: ToolbarRowSpec list) : Element =
    columnA MStart CStretch 0.0
        [ for r in rows -> heightBox 38.0 (padXY Ui.pad 0.0 (rowA MStart CCenter 6.0 (r.Controls |> List.map tbControl))) ]

// ── dialog fields ──
let private radioRow (sg: Signal<string>) (changed: string -> unit) (o: ChoiceSpec) : Element =
    let ring = { Width = 1.0<px>; Cap = Round; Join = RoundJoin; Paint = Solid Ui.P.border; DashArray = []; DashOffset = 0.0<px>; MiterLimit = 4.0 }
    clickable (fun () -> if o.IsEnabled then changed o.Value)
        (rowA MStart CCenter 7.0
            [ canvas 14.0 14.0 (fun x y _ _ ->
                [ StrokeRoundedRect(rect (x + 2.0) (y + 2.0) 10.0 10.0, 5.0<px>, ring) ]
                @ (if sg.Peek() = o.Value then [ FillRoundedRect(rect (x + 4.5) (y + 4.5) 5.0 5.0, 2.5<px>, Ui.P.accent) ] else []))
              text 12.5 Ui.P.text o.Label ])

let private selectChip (sg: Signal<string>) (changed: string -> unit) (opts: ChoiceSpec list) : Element =
    let cycle () =
        let opts = opts |> List.filter (fun choice -> choice.IsEnabled)
        if not opts.IsEmpty then
            let i = opts |> List.tryFindIndex (fun o -> o.Value = sg.Peek()) |> Option.defaultValue -1
            changed opts.[(i + 1) % opts.Length].Value
    clickable cycle
        (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Control x y w h Ui.radius)
            (padXY 8.0 5.0 (rowA MStart CCenter 8.0
                [ expanded 1 (textf 12.0 Ui.P.text false (fun () -> opts |> List.tryFind (fun o -> o.Value = sg.Value) |> Option.map (fun o -> o.Label) |> Option.defaultValue (sg.Value)))
                  chevronDown ])))

let private renderField (st: ShellState) (f: DialogField) : Element =
    let sg = fieldSig st f
    let kind, update =
        match f.Kind with
        | DialogUpdatingField (kind, update) -> kind, Some update
        | kind -> kind, None
    let changed value =
        match update with
        | None -> sg.Value <- value
        | Some update ->
            for KeyValue(id, text) in update (currentMap st) value do
                match st.Fields.TryGetValue id with
                | true, target when target.Peek() <> text -> target.Value <- text
                | _ -> ()
    let labeled (control: Element) = rowA MStart CCenter Ui.gap [ sizedBox 124.0 18.0 (Ui.labelDim f.Label); expanded 1 control ]
    match kind with
    | DialogUpdatingField _ -> failwith "Nested dialog update is not supported"
    | DialogSectionHeader -> columnA MStart CStretch 4.0 [ textBold 12.5 Ui.P.text f.Label; Ui.divider () ]
    | DialogStatusField -> labeled (textf 12.0 Ui.P.dim false (fun () -> sg.Value))
    | DialogTextField ph -> labeled (Ui.textField sg ph (fun _ -> changed (sg.Peek())))
    | DialogNumberField _ -> labeled (sizedBox 90.0 27.0 (Ui.textField sg "" (fun _ -> changed (sg.Peek()))))
    | DialogTextAreaField (ph, lines) -> labeled (heightBox (float (max 2 lines) * 18.0 + 12.0) (Ui.textField sg ph (fun _ -> changed (sg.Peek()))))
    | DialogCheckboxField ->
        rowA MStart CCenter Ui.gap
            [ Ui.checkbox (fun () -> sg.Peek() = "true") (fun () -> changed (if sg.Peek() = "true" then "false" else "true"))
              text 12.5 Ui.P.text f.Label ]
    | DialogSelectField opts -> labeled (selectChip sg changed opts)
    | DialogRadioField opts -> labeled (columnA MStart CStretch 4.0 [ for o in opts -> radioRow sg changed o ])
    | DialogColorField ->
        labeled (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Well x y w h Ui.radius)
            (padXY 5.0 4.0 (canvas 40.0 14.0 (fun x y _ _ -> [ FillRoundedRect(rect x y 40.0 14.0, 2.0<px>, parseHex (sg.Peek())) ]))))
    | DialogSliderField (mn, mx) ->
        labeled (rowA MStart CCenter 8.0
            [ Ui.slider 150.0 mn mx (fun () -> (try float (sg.Peek()) with _ -> mn)) (fun v -> changed (sprintf "%g" v))
              sizedBox 40.0 16.0 (textf 12.0 Ui.P.text false (fun () -> sg.Value)) ])
    | DialogComputedField compute -> labeled (dyn (fun () -> compute (currentMap st)) (fun () -> Ui.wrapText 12.0 Ui.P.dim (compute (currentMap st))))

let private dialogView (st: ShellState) (spec: DialogSpec) : Element =
    let fieldsView (fields: DialogField list) = columnA MStart CStretch 10.0 [ for f in fields -> renderField st f ]
    let body =
        if not spec.Tabs.IsEmpty then
            columnA MStart CStretch 12.0
                [ heightBox 28.0 (Ui.segmented (spec.Tabs |> List.map fst) (fun () -> st.ActiveTab.Peek()) (fun i -> st.ActiveTab.Value <- i))
                  dyn (fun () -> st.ActiveTab.Value) (fun () -> fieldsView (snd spec.Tabs.[st.ActiveTab.Peek()])) ]
        else fieldsView spec.Fields
    let runBtn (b: DialogButton) =
        let act () = if b.Action (currentMap st) then closeDialog st
        if b.IsPrimary then Ui.btnPrimary b.Label act else Ui.btn b.Label act
    Wisp.Widgets.modal (fun () -> closeDialog st)
        (fixedWidth 470.0
            (Ui.surfaceBox (fun x y w h -> Ui.modalCmds x y w h 8.0)
                (padding 20.0
                    (columnA MStart CStretch 14.0
                        [ textBold 15.0 Ui.P.text spec.Title
                          (match spec.Subtitle with Some s -> text 12.0 Ui.P.dim s | None -> Ui.blank ())
                          Ui.divider ()
                          body
                          Ui.divider ()
                          rowA MEnd CCenter Ui.gap [ for b in spec.Buttons -> runBtn b ] ]))))

/// Render a dialog spec on its own for headless preview / screenshot tests.
/// The chrome spec is accepted for parity with the Composit OfficeShell API;
/// Wisp uses its own Atrium theme for these standalone previews.
let dialogPreview (_chromeSpec: ChromeSpec) (spec: DialogSpec) : Element =
    let st = create ()
    openDialog st spec
    dialogView st spec

// ── shell composition: menu bar + toolbar + workspace, with menu/dialog overlays ──
let chrome (st: ShellState) (title: string) (menus: MenuSpec list) (toolbars: ToolbarRowSpec list) (workspace: Element) : Element =
    let ms = menus |> List.map toMenu
    let panel child = Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Panel x y w h 0.0) child
    stack
        [ Ui.windowSurface
            (columnA MStart CStretch 0.0
                [ panel (rowA MStart CCenter 0.0 [ Ui.menuBar st.OpenMenu ms; expanded 1 (Ui.blank ()); padXY Ui.pad 0.0 (textf 11.5 Ui.P.mute false (fun () -> title)) ])
                  Ui.divider ()
                  panel (toolbar toolbars)
                  Ui.divider ()
                  expanded 1 workspace ])
          Ui.menuOverlay st.OpenMenu ms
          show (fun () -> st.Dialog.Value.IsSome) (fun () -> match st.Dialog.Peek() with Some s -> dialogView st s | None -> Ui.blank ()) ]
