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

    type ToolbarRowSpec = {
        Label: string
        Controls: ToolbarControlSpec list
    }

    type MenuItemSpec = {
        TestId: string
        Label: string
        IsChecked: (unit -> bool) option
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
        IsChecked = isChecked
        Action = action
    }

    let toolbarButton testId label active action =
        ToolbarButton { TestId = testId; Label = label; IsActive = active; Action = action }

    let toolbarSelect testId value options action =
        ToolbarSelect { TestId = testId; Value = value; Options = options; Action = action }

    let toolbarNumberInput testId value minValue maxValue action =
        ToolbarInput { TestId = testId; Kind = NumberInput; Value = value; Min = minValue; Max = maxValue; Action = action }

    let toolbarColorInput testId value action =
        ToolbarInput { TestId = testId; Kind = ColorInput; Value = value; Min = None; Max = None; Action = action }

    let toolbarTextInput testId value action =
        ToolbarInput { TestId = testId; Kind = TextInput; Value = value; Min = None; Max = None; Action = action }
