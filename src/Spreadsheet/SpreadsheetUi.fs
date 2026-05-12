namespace RedBat.Spreadsheet

/// Spreadsheet-specific UI specs. Generic types live in RedBat.Office.Ui; this
/// module supplies the Excel-style chrome palette, menus, toolbar rows, and
/// ShellState/ShellActions records that both the Fable and native shell
/// renderers consume.
module Ui =
    open RedBat.Office.Ui

    // Re-export the generic spec types so consumers that read them off the
    // Spreadsheet UI module keep working.
    type ChoiceSpec = RedBat.Office.Ui.ChoiceSpec
    type ButtonSpec = RedBat.Office.Ui.ButtonSpec
    type SelectSpec = RedBat.Office.Ui.SelectSpec
    type InputSpec = RedBat.Office.Ui.InputSpec
    type InputKind = RedBat.Office.Ui.InputKind
    type ToolbarControlSpec = RedBat.Office.Ui.ToolbarControlSpec
    type ToolbarRowSpec = RedBat.Office.Ui.ToolbarRowSpec
    type MenuItemSpec = RedBat.Office.Ui.MenuItemSpec
    type MenuSpec = RedBat.Office.Ui.MenuSpec
    type ChromeSpec = RedBat.Office.Ui.ChromeSpec

    let choice = RedBat.Office.Ui.choice

    /// Excel-style chrome palette. The DocumentTrackWidth field is reused for
    /// the formula-bar width.
    let chrome : ChromeSpec = {
        WindowWidth = 1080.0
        MenuWidth = 210.0
        MenuGutterWidth = 30.0
        DocumentTrackWidth = 1000.0
        TitleTop = "#1F5E2C"
        TitleBottom = "#3A8C4E"
        Chrome = "#F4F7FB"
        ChromeTop = "#FFFFFF"
        ChromeBottom = "#E8EFF8"
        ChromeLight = "#FFFFFF"
        ChromeMid = "#BFCDE0"
        ChromeDark = "#7487A0"
        Pressed = "#D3E3F9"
        HotTop = "#E5F1FF"
        HotBottom = "#CDE0FB"
        Input = "#FFFFFF"
        MenuPanel = "#FFFFFF"
        MenuGutter = "#F6F8FC"
        PaperDesk = "#E8EEF6"
        Border = "#8FA2BC"
        Text = "#111111"
    }

    let chromeCssVariables = [
        "--xl-title-top", chrome.TitleTop
        "--xl-title-bottom", chrome.TitleBottom
        "--xl-chrome", chrome.Chrome
        "--xl-chrome-top", chrome.ChromeTop
        "--xl-chrome-bottom", chrome.ChromeBottom
        "--xl-chrome-light", chrome.ChromeLight
        "--xl-chrome-mid", chrome.ChromeMid
        "--xl-chrome-dark", chrome.ChromeDark
        "--xl-pressed", chrome.Pressed
        "--xl-hot-top", chrome.HotTop
        "--xl-hot-bottom", chrome.HotBottom
        "--xl-input", chrome.Input
        "--xl-menu-panel", chrome.MenuPanel
        "--xl-menu-gutter", chrome.MenuGutter
        "--xl-paper-desk", chrome.PaperDesk
        "--xl-border", chrome.Border
        "--xl-text", chrome.Text
        // Spreadsheet-specific palette extensions live alongside the shared
        // chrome variables so a single root <style> applies both.
        "--xl-header-fill", "#F2F2F2"
        "--xl-grid-line", "#D0D7E2"
        "--xl-selection-fill", "#E8F5E9"
        "--xl-selection-stroke", "#217346"
    ]

    /// Per-menu x offsets for the native shell, mirroring WordUi.menuLeft.
    let menuLeft testId =
        match testId with
        | "xl-file-menu" -> 3.0
        | "xl-edit-menu" -> 36.0
        | "xl-view-menu" -> 72.0
        | "xl-insert-menu" -> 108.0
        | "xl-format-menu" -> 152.0
        | "xl-data-menu" -> 205.0
        | "xl-table-menu" -> 244.0
        | _ -> 3.0

    /// Map test ids to abstract glyph names so backends can pick icons.
    let toolbarGlyph testId =
        match testId with
        | "xl-find-replace" -> "search"
        | "xl-bold" -> "bold"
        | "xl-italic" -> "italic"
        | "xl-underline" -> "underline"
        | "xl-strike" -> "strike"
        | "xl-clear-formatting" -> "clear-formatting"
        | "xl-align-left" -> "align-left"
        | "xl-align-center" -> "align-center"
        | "xl-align-right" -> "align-right"
        | "xl-valign-top" -> "valign-top"
        | "xl-valign-middle" -> "valign-middle"
        | "xl-valign-bottom" -> "valign-bottom"
        | "xl-wrap-text" -> "wrap-text"
        | "xl-merge-cells" -> "merge"
        | "xl-unmerge-cells" -> "unmerge"
        | "xl-format-number" -> "format-number"
        | "xl-format-currency" -> "format-currency"
        | "xl-format-percent" -> "format-percent"
        | "xl-format-date" -> "format-date"
        | "xl-decimal-increase" -> "decimal-plus"
        | "xl-decimal-decrease" -> "decimal-minus"
        | "xl-insert-row" -> "row-add"
        | "xl-insert-column" -> "column-add"
        | "xl-delete-row" -> "row-delete"
        | "xl-delete-column" -> "column-delete"
        | "xl-sort-asc" -> "sort-asc"
        | "xl-sort-desc" -> "sort-desc"
        | "xl-filter" -> "filter"
        | "xl-clear-filter" -> "filter-clear"
        | "xl-cut" -> "cut"
        | "xl-copy" -> "copy"
        | "xl-paste" -> "paste"
        | "xl-undo" -> "undo"
        | "xl-redo" -> "redo"
        | "xl-text-color" -> "text-color"
        | "xl-bg-color" -> "highlight-color"
        | "xl-border-style" -> "border"
        | "xl-insert-chart" -> "chart"
        | "xl-insert-pivot" -> "pivot"
        | "xl-insert-table" -> "table"
        | "xl-insert-sheet" -> "sheet-add"
        | "xl-delete-sheet" -> "sheet-delete"
        | "xl-rename-sheet" -> "sheet-rename"
        | "xl-print-pdf" -> "document"
        | "xl-freeze-panes" -> "freeze"
        | "xl-split-panes" -> "split"
        | "xl-zoom-100" -> "zoom-100"
        | _ -> "text"

    let showToolbarPressed testId isActive =
        isActive
        && testId <> "xl-align-left"
        && testId <> "xl-valign-bottom"

    type ShellState = {
        ActiveCellRef: unit -> string
        FormulaText: unit -> string
        IsBoldActive: unit -> bool
        IsItalicActive: unit -> bool
        IsUnderlineActive: unit -> bool
        IsStrikeActive: unit -> bool
        IsWrapTextActive: unit -> bool
        IsMergedActive: unit -> bool
        IsFilterActive: unit -> bool
        IsFreezePanesActive: unit -> bool
        HorizontalAlign: unit -> string
        VerticalAlign: unit -> string
        NumberFormat: unit -> string
        TextColor: unit -> string
        BgColor: unit -> string
        FontFamily: unit -> string
        FontSize: unit -> string
        BorderStyle: unit -> string
        ZoomLevel: unit -> string
        ActiveSheet: unit -> string
        Sheets: unit -> ChoiceSpec list
        SelectionSummary: unit -> string
    }

    type ShellActions = {
        PrintPdf: unit -> unit
        OpenFindReplace: unit -> unit
        Cut: unit -> unit
        Copy: unit -> unit
        Paste: unit -> unit
        Undo: unit -> unit
        Redo: unit -> unit
        ToggleBold: unit -> unit
        ToggleItalic: unit -> unit
        ToggleUnderline: unit -> unit
        ToggleStrike: unit -> unit
        ToggleWrapText: unit -> unit
        ClearFormatting: unit -> unit
        ApplyHorizontalAlign: string -> unit
        ApplyVerticalAlign: string -> unit
        ApplyNumberFormat: string -> unit
        IncreaseDecimals: unit -> unit
        DecreaseDecimals: unit -> unit
        ApplyFontFamily: string -> unit
        ApplyFontSize: string -> unit
        ApplyTextColor: string -> unit
        ApplyBgColor: string -> unit
        ApplyBorderStyle: string -> unit
        ApplyZoom: string -> unit
        MergeCells: unit -> unit
        UnmergeCells: unit -> unit
        InsertRow: unit -> unit
        InsertColumn: unit -> unit
        DeleteRow: unit -> unit
        DeleteColumn: unit -> unit
        SortAscending: unit -> unit
        SortDescending: unit -> unit
        ToggleFilter: unit -> unit
        ClearFilter: unit -> unit
        ToggleFreezePanes: unit -> unit
        SplitPanes: unit -> unit
        InsertChart: unit -> unit
        InsertPivot: unit -> unit
        InsertTable: unit -> unit
        InsertSheet: unit -> unit
        DeleteSheet: unit -> unit
        RenameSheet: unit -> unit
        SwitchSheet: string -> unit
        EditFormula: string -> unit
        SubmitFormula: unit -> unit
    }

    let fontFamilies = [
        choice "Calibri" "Calibri"
        choice "Arial" "Arial"
        choice "Calibri Light" "Calibri Light"
        choice "Cambria" "Cambria"
        choice "Consolas" "Consolas"
        choice "Courier New" "Courier New"
        choice "Georgia" "Georgia"
        choice "Helvetica" "Helvetica"
        choice "Segoe UI" "Segoe UI"
        choice "Times New Roman" "Times New Roman"
        choice "Verdana" "Verdana"
    ]

    let fontSizes =
        [ "8"; "9"; "10"; "11"; "12"; "14"; "16"; "18"; "20"; "24"; "28"; "36" ]
        |> List.map (fun value -> choice value value)

    let numberFormats = [
        choice "general" "General"
        choice "number" "Number (1,234.56)"
        choice "currency" "Currency ($)"
        choice "percent" "Percent (%)"
        choice "date" "Date"
        choice "time" "Time"
        choice "scientific" "Scientific"
        choice "text" "Text"
    ]

    let borderStyles = [
        choice "none" "No border"
        choice "all" "All borders"
        choice "outside" "Outside border"
        choice "thick-outside" "Thick outside"
        choice "top" "Top border"
        choice "bottom" "Bottom border"
        choice "left" "Left border"
        choice "right" "Right border"
    ]

    let zoomLevels =
        [ "50"; "75"; "100"; "125"; "150"; "200" ]
        |> List.map (fun value -> choice value (value + "%"))

    let private item = menuItem
    let private button = toolbarButton

    let menus (state: ShellState) (actions: ShellActions) = [
        {
            TestId = "xl-file-menu"
            PanelTestId = "xl-file-menu-panel"
            Label = "File"
            Items = [
                item "xl-menu-print-pdf" "Print PDF" None actions.PrintPdf
                item "xl-menu-insert-sheet" "New sheet" None actions.InsertSheet
                item "xl-menu-delete-sheet" "Delete sheet" None actions.DeleteSheet
                item "xl-menu-rename-sheet" "Rename sheet..." None actions.RenameSheet
            ]
        }
        {
            TestId = "xl-edit-menu"
            PanelTestId = "xl-edit-menu-panel"
            Label = "Edit"
            Items = [
                item "xl-menu-undo" "Undo" None actions.Undo
                item "xl-menu-redo" "Redo" None actions.Redo
                item "xl-menu-cut" "Cut" None actions.Cut
                item "xl-menu-copy" "Copy" None actions.Copy
                item "xl-menu-paste" "Paste" None actions.Paste
                item "xl-menu-find-replace" "Find and replace..." None actions.OpenFindReplace
            ]
        }
        {
            TestId = "xl-view-menu"
            PanelTestId = "xl-view-menu-panel"
            Label = "View"
            Items = [
                item "xl-menu-zoom-100" "Zoom 100%" (Some (fun () -> state.ZoomLevel() = "100")) (fun () -> actions.ApplyZoom "100")
                item "xl-menu-zoom-150" "Zoom 150%" (Some (fun () -> state.ZoomLevel() = "150")) (fun () -> actions.ApplyZoom "150")
                item "xl-menu-freeze-panes" "Freeze panes" (Some state.IsFreezePanesActive) actions.ToggleFreezePanes
                item "xl-menu-split-panes" "Split panes" None actions.SplitPanes
            ]
        }
        {
            TestId = "xl-insert-menu"
            PanelTestId = "xl-insert-menu-panel"
            Label = "Insert"
            Items = [
                item "xl-menu-insert-row" "Insert row" None actions.InsertRow
                item "xl-menu-insert-column" "Insert column" None actions.InsertColumn
                item "xl-menu-insert-chart" "Chart..." None actions.InsertChart
                item "xl-menu-insert-pivot" "Pivot table..." None actions.InsertPivot
                item "xl-menu-insert-table" "Table..." None actions.InsertTable
                item "xl-menu-insert-sheet" "Worksheet" None actions.InsertSheet
            ]
        }
        {
            TestId = "xl-format-menu"
            PanelTestId = "xl-format-menu-panel"
            Label = "Format"
            Items = [
                item "xl-menu-bold" "Bold" (Some state.IsBoldActive) actions.ToggleBold
                item "xl-menu-italic" "Italic" (Some state.IsItalicActive) actions.ToggleItalic
                item "xl-menu-underline" "Underline" (Some state.IsUnderlineActive) actions.ToggleUnderline
                item "xl-menu-strike" "Strikethrough" (Some state.IsStrikeActive) actions.ToggleStrike
                item "xl-menu-wrap-text" "Wrap text" (Some state.IsWrapTextActive) actions.ToggleWrapText
                item "xl-menu-merge-cells" "Merge cells" (Some state.IsMergedActive) actions.MergeCells
                item "xl-menu-unmerge-cells" "Unmerge cells" None actions.UnmergeCells
                item "xl-menu-clear-formatting" "Clear formatting" None actions.ClearFormatting
                item "xl-menu-format-number" "Number (1,234.56)" (Some (fun () -> state.NumberFormat() = "number")) (fun () -> actions.ApplyNumberFormat "number")
                item "xl-menu-format-currency" "Currency ($)" (Some (fun () -> state.NumberFormat() = "currency")) (fun () -> actions.ApplyNumberFormat "currency")
                item "xl-menu-format-percent" "Percent (%)" (Some (fun () -> state.NumberFormat() = "percent")) (fun () -> actions.ApplyNumberFormat "percent")
                item "xl-menu-format-date" "Date" (Some (fun () -> state.NumberFormat() = "date")) (fun () -> actions.ApplyNumberFormat "date")
                item "xl-menu-text-color-red" "Red text" (Some (fun () -> state.TextColor() = "#C00000")) (fun () -> actions.ApplyTextColor "#C00000")
                item "xl-menu-bg-color-yellow" "Yellow fill" (Some (fun () -> state.BgColor() = "#FFF2CC")) (fun () -> actions.ApplyBgColor "#FFF2CC")
            ]
        }
        {
            TestId = "xl-data-menu"
            PanelTestId = "xl-data-menu-panel"
            Label = "Data"
            Items = [
                item "xl-menu-sort-asc" "Sort A to Z" None actions.SortAscending
                item "xl-menu-sort-desc" "Sort Z to A" None actions.SortDescending
                item "xl-menu-filter" "Toggle filter" (Some state.IsFilterActive) actions.ToggleFilter
                item "xl-menu-clear-filter" "Clear filter" None actions.ClearFilter
            ]
        }
        {
            TestId = "xl-table-menu"
            PanelTestId = "xl-table-menu-panel"
            Label = "Table"
            Items = [
                item "xl-menu-insert-table" "Convert range to table" None actions.InsertTable
                item "xl-menu-insert-row" "Insert row" None actions.InsertRow
                item "xl-menu-insert-column" "Insert column" None actions.InsertColumn
                item "xl-menu-delete-row" "Delete row" None actions.DeleteRow
                item "xl-menu-delete-column" "Delete column" None actions.DeleteColumn
            ]
        }
    ]

    let toolbarRows (state: ShellState) (actions: ShellActions) = [
        { Label = "Standard"; Controls = [
            button "xl-undo" "Undo" None actions.Undo
            button "xl-redo" "Redo" None actions.Redo
            button "xl-cut" "Cut" None actions.Cut
            button "xl-copy" "Copy" None actions.Copy
            button "xl-paste" "Paste" None actions.Paste
            button "xl-find-replace" "Find" None actions.OpenFindReplace
            button "xl-print-pdf" "PDF" None actions.PrintPdf
            button "xl-insert-row" "+Row" None actions.InsertRow
            button "xl-insert-column" "+Col" None actions.InsertColumn
            button "xl-delete-row" "-Row" None actions.DeleteRow
            button "xl-delete-column" "-Col" None actions.DeleteColumn
            button "xl-sort-asc" "A→Z" None actions.SortAscending
            button "xl-sort-desc" "Z→A" None actions.SortDescending
            button "xl-filter" "Filter" (Some state.IsFilterActive) actions.ToggleFilter
            button "xl-clear-filter" "Clear" None actions.ClearFilter
            button "xl-insert-chart" "Chart" None actions.InsertChart
            button "xl-insert-pivot" "Pivot" None actions.InsertPivot
            button "xl-insert-table" "Table" None actions.InsertTable
            button "xl-freeze-panes" "Freeze" (Some state.IsFreezePanesActive) actions.ToggleFreezePanes
        ] }
        { Label = "Formatting"; Controls = [
            toolbarSelect "xl-font-family" state.FontFamily (fun () -> fontFamilies) actions.ApplyFontFamily
            toolbarSelect "xl-font-size" state.FontSize (fun () -> fontSizes) actions.ApplyFontSize
            button "xl-bold" "B" (Some state.IsBoldActive) actions.ToggleBold
            button "xl-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic
            button "xl-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline
            button "xl-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike
            button "xl-clear-formatting" "Clear" None actions.ClearFormatting
            toolbarColorInput "xl-text-color" state.TextColor actions.ApplyTextColor
            toolbarColorInput "xl-bg-color" state.BgColor actions.ApplyBgColor
            toolbarSelect "xl-border-style" state.BorderStyle (fun () -> borderStyles) actions.ApplyBorderStyle
            button "xl-align-left" "Left" (Some (fun () -> state.HorizontalAlign() = "left")) (fun () -> actions.ApplyHorizontalAlign "left")
            button "xl-align-center" "Center" (Some (fun () -> state.HorizontalAlign() = "center")) (fun () -> actions.ApplyHorizontalAlign "center")
            button "xl-align-right" "Right" (Some (fun () -> state.HorizontalAlign() = "right")) (fun () -> actions.ApplyHorizontalAlign "right")
            button "xl-valign-top" "Top" (Some (fun () -> state.VerticalAlign() = "top")) (fun () -> actions.ApplyVerticalAlign "top")
            button "xl-valign-middle" "Middle" (Some (fun () -> state.VerticalAlign() = "middle")) (fun () -> actions.ApplyVerticalAlign "middle")
            button "xl-valign-bottom" "Bottom" (Some (fun () -> state.VerticalAlign() = "bottom")) (fun () -> actions.ApplyVerticalAlign "bottom")
            button "xl-wrap-text" "Wrap" (Some state.IsWrapTextActive) actions.ToggleWrapText
            button "xl-merge-cells" "Merge" (Some state.IsMergedActive) actions.MergeCells
            button "xl-unmerge-cells" "Unmerge" None actions.UnmergeCells
        ] }
        { Label = "Number"; Controls = [
            toolbarSelect "xl-number-format" state.NumberFormat (fun () -> numberFormats) actions.ApplyNumberFormat
            button "xl-format-currency" "$" (Some (fun () -> state.NumberFormat() = "currency")) (fun () -> actions.ApplyNumberFormat "currency")
            button "xl-format-percent" "%" (Some (fun () -> state.NumberFormat() = "percent")) (fun () -> actions.ApplyNumberFormat "percent")
            button "xl-format-number" "1,234" (Some (fun () -> state.NumberFormat() = "number")) (fun () -> actions.ApplyNumberFormat "number")
            button "xl-decimal-increase" ".0+" None actions.IncreaseDecimals
            button "xl-decimal-decrease" ".0-" None actions.DecreaseDecimals
            toolbarSelect "xl-zoom-level" state.ZoomLevel (fun () -> zoomLevels) actions.ApplyZoom
        ] }
    ]

    /// Formula bar spec: the cell reference is read-only, the formula edit
    /// drives EditFormula on every keystroke and SubmitFormula on Enter.
    type FormulaBarSpec = {
        CellRef: unit -> string
        Formula: unit -> string
        OnEdit: string -> unit
        OnSubmit: unit -> unit
    }

    let formulaBar (state: ShellState) (actions: ShellActions) = {
        CellRef = state.ActiveCellRef
        Formula = state.FormulaText
        OnEdit = actions.EditFormula
        OnSubmit = actions.SubmitFormula
    }

    /// Sheet tab bar spec for the bottom of the window.
    type SheetTabSpec = {
        Tabs: unit -> ChoiceSpec list
        Active: unit -> string
        OnSwitch: string -> unit
        OnInsert: unit -> unit
        OnDelete: unit -> unit
        OnRename: unit -> unit
    }

    let sheetTabs (state: ShellState) (actions: ShellActions) = {
        Tabs = state.Sheets
        Active = state.ActiveSheet
        OnSwitch = actions.SwitchSheet
        OnInsert = actions.InsertSheet
        OnDelete = actions.DeleteSheet
        OnRename = actions.RenameSheet
    }
