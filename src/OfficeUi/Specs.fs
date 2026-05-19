namespace RedBat.Office

/// Backend-agnostic UI spec types shared by Word, Spreadsheet, and any other
/// Office-style apps. Concrete apps define their own chrome palette, menu
/// list, toolbar rows, and ShellState/ShellActions records using these types,
/// then hand them to a shared Fable renderer or a shared native (../widgets)
/// renderer.
module Ui =
    type ChoiceSpec = {
        Value: string
        Label: string
    }

    type ButtonSpec = {
        TestId: string
        Label: string
        IsActive: (unit -> bool) option
        /// When Some, evaluated each frame; false renders the button greyed
        /// and swallows the click. None = always enabled. Wire to undo-stack
        /// presence, clipboard contents, selection state, etc.
        IsEnabled: (unit -> bool) option
        /// Optional hover-tooltip text. Toolbar buttons are icon-only, so
        /// the chrome surfaces this in the status bar after a short hover
        /// delay rather than inventing a floating tooltip layer.
        Tooltip: string option
        Action: unit -> unit
    }

    type SelectSpec = {
        TestId: string
        Value: unit -> string
        Options: unit -> ChoiceSpec list
        Action: string -> unit
    }

    /// Free-form input kinds. ColorInput renders a swatch; NumberInput a
    /// numeric textbox; TextInput a plain textbox.
    type InputKind =
        | NumberInput
        | ColorInput
        | TextInput

    type InputSpec = {
        TestId: string
        Kind: InputKind
        Value: unit -> string
        Min: string option
        Max: string option
        Action: string -> unit
    }

    type ToolbarControlSpec =
        | ToolbarButton of ButtonSpec
        | ToolbarSelect of SelectSpec
        | ToolbarInput of InputSpec
        | ToolbarSeparator

    type ToolbarRowSpec = {
        Label: string
        Controls: ToolbarControlSpec list
    }

    type MenuItemSpec = {
        TestId: string
        Label: string
        /// Optional keyboard shortcut hint (e.g. "Ctrl+S") rendered right-
        /// aligned in muted text. The shortcut isn't bound by this spec —
        /// the app's key handler still has to do the dispatching.
        Shortcut: string option
        IsChecked: (unit -> bool) option
        /// Same semantics as ButtonSpec.IsEnabled — false greys the row and
        /// swallows the click.
        IsEnabled: (unit -> bool) option
        Action: unit -> unit
    }

    type MenuSpec = {
        TestId: string
        PanelTestId: string
        Label: string
        Items: MenuItemSpec list
    }

    /// Generic chrome palette used by both native and Fable shells. Apps
    /// supply concrete values; native renderers parse the hex strings and
    /// Fable renderers emit them as CSS variables.
    type ChromeSpec = {
        WindowWidth: float
        MenuWidth: float
        MenuGutterWidth: float
        DocumentTrackWidth: float
        TitleTop: string
        TitleBottom: string
        Chrome: string
        ChromeTop: string
        ChromeBottom: string
        ChromeLight: string
        ChromeMid: string
        ChromeDark: string
        Pressed: string
        HotTop: string
        HotBottom: string
        Input: string
        MenuPanel: string
        MenuGutter: string
        PaperDesk: string
        Border: string
        Text: string
    }

    let choice value label = { Value = value; Label = label }

    let menuItem testId label isChecked action = {
        TestId = testId
        Label = label
        Shortcut = None
        IsChecked = isChecked
        IsEnabled = None
        Action = action
    }

    let menuItemShortcut testId label shortcut isChecked action = {
        TestId = testId
        Label = label
        Shortcut = Some shortcut
        IsChecked = isChecked
        IsEnabled = None
        Action = action
    }

    /// Pipe-friendly menu-item decorator: `menuItem ... |> menuItemEnabled state.CanX`.
    let menuItemEnabled (predicate: unit -> bool) (item: MenuItemSpec) =
        { item with IsEnabled = Some predicate }

    let toolbarButton testId label active action =
        ToolbarButton {
            TestId = testId; Label = label
            IsActive = active; IsEnabled = None; Tooltip = None
            Action = action
        }

    /// Pipe-friendly button decorators. Use as
    ///   toolbarButton "x" "B" None act
    ///   |> withTooltip "Bold (Ctrl+B)"
    ///   |> withEnabled state.HasSelection
    let withTooltip (text: string) = function
        | ToolbarButton b -> ToolbarButton { b with Tooltip = Some text }
        | other -> other

    let withEnabled (predicate: unit -> bool) = function
        | ToolbarButton b -> ToolbarButton { b with IsEnabled = Some predicate }
        | other -> other

    let toolbarSelect testId value options action =
        ToolbarSelect { TestId = testId; Value = value; Options = options; Action = action }

    let toolbarNumberInput testId value minValue maxValue action =
        ToolbarInput { TestId = testId; Kind = NumberInput; Value = value; Min = minValue; Max = maxValue; Action = action }

    let toolbarColorInput testId value action =
        ToolbarInput { TestId = testId; Kind = ColorInput; Value = value; Min = None; Max = None; Action = action }

    let toolbarTextInput testId value action =
        ToolbarInput { TestId = testId; Kind = TextInput; Value = value; Min = None; Max = None; Action = action }

    let toolbarSep = ToolbarSeparator

    // ── Dialogs ──────────────────────────────────────────────
    //
    // Backend-agnostic dialog spec. Apps build a DialogSpec value and hand
    // it to the shell, which renders it (native = widget overlay; Fable =
    // DOM dialog). The submit callback receives a `Map<fieldId, value>` so
    // the renderer can carry field state internally without the app
    // pre-creating Signals.

    type DialogFieldKind =
        | DialogTextField of placeholder: string
        | DialogNumberField of minValue: string option * maxValue: string option
        | DialogCheckboxField
        | DialogSelectField of options: ChoiceSpec list
        /// Multi-line read-only label, e.g. a result/status line.
        | DialogStatusField

    type DialogField = {
        /// Stable id used as the key in the submit callback's value map.
        Id: string
        /// Display label.
        Label: string
        Kind: DialogFieldKind
        /// Initial value. Booleans serialize as "true"/"false".
        Initial: string
    }

    type DialogButton = {
        /// `data-testid` and visible button label.
        TestId: string
        Label: string
        /// Primary button gets accent styling and is triggered by Enter.
        IsPrimary: bool
        /// Receives the latest field values keyed by field id. Return true to
        /// close the dialog, false to leave it open (e.g. validation error).
        Action: Map<string, string> -> bool
    }

    type DialogSpec = {
        /// `data-testid` for the dialog root element.
        TestId: string
        Title: string
        /// Optional one-line description shown under the title.
        Subtitle: string option
        Fields: DialogField list
        Buttons: DialogButton list
    }

    let dialogText id label placeholder initial =
        { Id = id; Label = label; Kind = DialogTextField placeholder; Initial = initial }

    let dialogNumber id label minValue maxValue initial =
        { Id = id; Label = label; Kind = DialogNumberField(minValue, maxValue); Initial = initial }

    let dialogCheckbox id label initial =
        { Id = id; Label = label; Kind = DialogCheckboxField; Initial = (if initial then "true" else "false") }

    let dialogSelect id label options initial =
        { Id = id; Label = label; Kind = DialogSelectField options; Initial = initial }

    let dialogStatus id label initial =
        { Id = id; Label = label; Kind = DialogStatusField; Initial = initial }

    let dialogButton testId label isPrimary action =
        { TestId = testId; Label = label; IsPrimary = isPrimary; Action = action }

    /// Standard Cancel button — closes the dialog without doing anything.
    let dialogCancel = dialogButton "dlg-cancel" "Cancel" false (fun _ -> true)
