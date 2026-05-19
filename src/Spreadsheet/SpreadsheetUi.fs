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

    /// Flat, white-first chrome palette — same Google-Sheets / OnlyOffice
    /// language the Word shell adopted in the chrome refresh. Top == Bottom
    /// on every band so there's no gradient seam; the only separators are
    /// 1px ChromeMid hairlines. The spreadsheet keeps a green identity, but
    /// the green now lives in accents (active sheet tab, selection stroke)
    /// rather than a heavy title gradient. DocumentTrackWidth == WindowWidth
    /// so the formula bar and grid span the full window.
    let chrome : ChromeSpec = {
        WindowWidth = 1100.0
        MenuWidth = 240.0
        MenuGutterWidth = 24.0
        DocumentTrackWidth = 1100.0
        TitleTop = "#FFFFFF"
        TitleBottom = "#FFFFFF"
        Chrome = "#FFFFFF"
        ChromeTop = "#FFFFFF"
        ChromeBottom = "#FFFFFF"
        ChromeLight = "#FFFFFF"
        ChromeMid = "#E1E5EA"      // hairline separators
        ChromeDark = "#B9C0C8"     // muted icon/tick grey
        Pressed = "#E6F4EA"         // soft green tint (active)
        HotTop = "#F1F3F4"          // Google-grey hover
        HotBottom = "#F1F3F4"
        Input = "#FFFFFF"
        MenuPanel = "#FFFFFF"
        MenuGutter = "#FFFFFF"      // no gutter color difference
        PaperDesk = "#F8F9FA"       // very light desk
        Border = "#DADCE0"          // Google border grey
        Text = "#202124"
    }

    /// Excel/Sheets green — used for the active sheet-tab accent and the
    /// cell-selection stroke. Not part of the shared ChromeSpec because no
    /// other Office app needs it.
    let accentGreen = "#217346"

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
        /// True iff the editor's undo stack is non-empty. The chrome wires
        /// this to the Undo button's IsEnabled so the button greys out
        /// when there's nothing to undo, instead of looking clickable but
        /// being a no-op.
        CanUndo: unit -> bool
        /// True iff the editor's redo stack is non-empty.
        CanRedo: unit -> bool
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
                item "xl-menu-undo" "Undo" None actions.Undo |> menuItemEnabled state.CanUndo
                item "xl-menu-redo" "Redo" None actions.Redo |> menuItemEnabled state.CanRedo
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
            |> withTooltip "Undo (Ctrl+Z)" |> withEnabled state.CanUndo
            button "xl-redo" "Redo" None actions.Redo
            |> withTooltip "Redo (Ctrl+Y)" |> withEnabled state.CanRedo
            button "xl-cut" "Cut" None actions.Cut |> withTooltip "Cut (Ctrl+X)"
            button "xl-copy" "Copy" None actions.Copy |> withTooltip "Copy (Ctrl+C)"
            button "xl-paste" "Paste" None actions.Paste |> withTooltip "Paste (Ctrl+V)"
            button "xl-find-replace" "Find" None actions.OpenFindReplace |> withTooltip "Find and replace (Ctrl+F)"
            button "xl-print-pdf" "PDF" None actions.PrintPdf |> withTooltip "Export to PDF"
            button "xl-insert-row" "+Row" None actions.InsertRow |> withTooltip "Insert row above"
            button "xl-insert-column" "+Col" None actions.InsertColumn |> withTooltip "Insert column left"
            button "xl-delete-row" "-Row" None actions.DeleteRow |> withTooltip "Delete row"
            button "xl-delete-column" "-Col" None actions.DeleteColumn |> withTooltip "Delete column"
            button "xl-sort-asc" "A→Z" None actions.SortAscending |> withTooltip "Sort ascending"
            button "xl-sort-desc" "Z→A" None actions.SortDescending |> withTooltip "Sort descending"
            button "xl-filter" "Filter" (Some state.IsFilterActive) actions.ToggleFilter |> withTooltip "Toggle filter"
            button "xl-clear-filter" "Clear" None actions.ClearFilter |> withTooltip "Clear filters"
            button "xl-insert-chart" "Chart" None actions.InsertChart |> withTooltip "Insert chart"
            button "xl-insert-pivot" "Pivot" None actions.InsertPivot |> withTooltip "Insert pivot table"
            button "xl-insert-table" "Table" None actions.InsertTable |> withTooltip "Insert table"
            button "xl-freeze-panes" "Freeze" (Some state.IsFreezePanesActive) actions.ToggleFreezePanes |> withTooltip "Freeze panes"
        ] }
        { Label = "Formatting"; Controls = [
            toolbarSelect "xl-font-family" state.FontFamily (fun () -> fontFamilies) actions.ApplyFontFamily
            toolbarSelect "xl-font-size" state.FontSize (fun () -> fontSizes) actions.ApplyFontSize
            button "xl-bold" "B" (Some state.IsBoldActive) actions.ToggleBold |> withTooltip "Bold (Ctrl+B)"
            button "xl-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic |> withTooltip "Italic (Ctrl+I)"
            button "xl-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline |> withTooltip "Underline (Ctrl+U)"
            button "xl-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike |> withTooltip "Strikethrough"
            button "xl-clear-formatting" "Clear" None actions.ClearFormatting |> withTooltip "Clear formatting"
            toolbarColorInput "xl-text-color" state.TextColor actions.ApplyTextColor
            toolbarColorInput "xl-bg-color" state.BgColor actions.ApplyBgColor
            toolbarSelect "xl-border-style" state.BorderStyle (fun () -> borderStyles) actions.ApplyBorderStyle
            button "xl-align-left" "Left" (Some (fun () -> state.HorizontalAlign() = "left")) (fun () -> actions.ApplyHorizontalAlign "left") |> withTooltip "Align left"
            button "xl-align-center" "Center" (Some (fun () -> state.HorizontalAlign() = "center")) (fun () -> actions.ApplyHorizontalAlign "center") |> withTooltip "Align center"
            button "xl-align-right" "Right" (Some (fun () -> state.HorizontalAlign() = "right")) (fun () -> actions.ApplyHorizontalAlign "right") |> withTooltip "Align right"
            button "xl-valign-top" "Top" (Some (fun () -> state.VerticalAlign() = "top")) (fun () -> actions.ApplyVerticalAlign "top") |> withTooltip "Align top"
            button "xl-valign-middle" "Middle" (Some (fun () -> state.VerticalAlign() = "middle")) (fun () -> actions.ApplyVerticalAlign "middle") |> withTooltip "Align middle"
            button "xl-valign-bottom" "Bottom" (Some (fun () -> state.VerticalAlign() = "bottom")) (fun () -> actions.ApplyVerticalAlign "bottom") |> withTooltip "Align bottom"
            button "xl-wrap-text" "Wrap" (Some state.IsWrapTextActive) actions.ToggleWrapText |> withTooltip "Wrap text"
            button "xl-merge-cells" "Merge" (Some state.IsMergedActive) actions.MergeCells |> withTooltip "Merge cells"
            button "xl-unmerge-cells" "Unmerge" None actions.UnmergeCells |> withTooltip "Unmerge cells"
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

    // ── Dialog specs ─────────────────────────────────────────
    //
    // Backend-agnostic dialog descriptors that mirror the Word side. Apps
    // pass callbacks that perform the underlying workbook mutation; the
    // spec describes the form and binds the buttons to those callbacks.

    /// Find / Replace.
    ///   `onFindNext q matchCase`        → status string ("Match 1 of 5").
    ///   `onReplaceAll q r matchCase`    → status string ("Replaced 3").
    let findReplaceDialog
            (onFindNext: string -> bool -> string)
            (onReplaceAll: string -> string -> bool -> string) : DialogSpec =
        {
            TestId = "xl-dlg-find-replace"
            Title = "Find and replace"
            Subtitle = None
            Fields = [
                dialogText "find" "Find" "Search text" ""
                dialogText "replace" "Replace with" "Replacement text" ""
                dialogCheckbox "case" "Match case" false
                dialogStatus "result" "Result" ""
            ]
            Buttons = [
                dialogCancel
                dialogButton "dlg-find-next" "Find next" false (fun values ->
                    let q = Map.find "find" values
                    let mc = (Map.find "case" values) = "true"
                    let _ = onFindNext q mc
                    false)
                dialogButton "dlg-replace-all" "Replace all" true (fun values ->
                    let q = Map.find "find" values
                    let r = Map.find "replace" values
                    let mc = (Map.find "case" values) = "true"
                    let _ = onReplaceAll q r mc
                    true)
            ]
        }

    /// Insert a new worksheet. `apply name` is called with the trimmed
    /// name; the workbook decides what to do if the name collides.
    let insertSheetDialog (apply: string -> unit) : DialogSpec = {
        TestId = "xl-dlg-insert-sheet"
        Title = "New worksheet"
        Subtitle = None
        Fields = [
            dialogText "name" "Sheet name" "Sheet" ""
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-insert-sheet-ok" "Add" true (fun values ->
                let name = (Map.find "name" values).Trim()
                if name = "" then false
                else
                    apply name
                    true)
        ]
    }

    /// Rename the active worksheet.
    let renameSheetDialog (initialName: string) (apply: string -> unit) : DialogSpec = {
        TestId = "xl-dlg-rename-sheet"
        Title = "Rename worksheet"
        Subtitle = None
        Fields = [
            dialogText "name" "New name" "Sheet" initialName
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-rename-sheet-ok" "Rename" true (fun values ->
                let name = (Map.find "name" values).Trim()
                if name = "" then false
                else
                    apply name
                    true)
        ]
    }

    /// Insert a table (rows x cols) at the active cell.
    let insertTableDialog (apply: int -> int -> unit) : DialogSpec = {
        TestId = "xl-dlg-insert-table"
        Title = "Insert table"
        Subtitle = None
        Fields = [
            dialogNumber "rows" "Rows" (Some "1") (Some "1000") "3"
            dialogNumber "cols" "Columns" (Some "1") (Some "100") "3"
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-insert-table-ok" "Insert" true (fun values ->
                let rows = System.Int32.TryParse(Map.find "rows" values) |> snd
                let cols = System.Int32.TryParse(Map.find "cols" values) |> snd
                if rows > 0 && cols > 0 then
                    apply rows cols
                    true
                else
                    false)
        ]
    }

    /// Format the selection's number / font / fills in one go. v1 only
    /// covers number format + a foreground/background hex pair; richer
    /// controls (borders, alignment) live on the toolbar already.
    let formatCellsDialog
            (initialFormat: string)
            (initialTextColor: string)
            (initialBgColor: string)
            (apply: string * string * string -> unit) : DialogSpec = {
        TestId = "xl-dlg-format-cells"
        Title = "Format cells"
        Subtitle = Some "Number format codes match Excel's custom-format syntax."
        Fields = [
            dialogSelect "format" "Number format" numberFormats initialFormat
            dialogText "text" "Text color (#hex)" "#000000" initialTextColor
            dialogText "bg" "Fill color (#hex)" "#FFFFFF" initialBgColor
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-format-cells-ok" "Apply" true (fun values ->
                let fmt = Map.find "format" values
                let text = (Map.find "text" values).Trim()
                let bg = (Map.find "bg" values).Trim()
                apply (fmt, text, bg)
                true)
        ]
    }

    /// Page setup for printing / PDF export. Orientation + margins.
    let pageSetupDialog
            (initialOrientation: string)
            (initialMarginLeft: float)
            (initialMarginRight: float)
            (initialMarginTop: float)
            (initialMarginBottom: float)
            (apply: string * float * float * float * float -> unit) : DialogSpec =
        let orientations = [
            choice "portrait" "Portrait"
            choice "landscape" "Landscape"
        ]
        {
            TestId = "xl-dlg-page-setup"
            Title = "Page setup"
            Subtitle = Some "Margins are in points (72 pt = 1 inch)."
            Fields = [
                dialogSelect "orientation" "Orientation" orientations initialOrientation
                dialogNumber "marginLeft" "Left margin" (Some "0") (Some "500") (sprintf "%g" initialMarginLeft)
                dialogNumber "marginRight" "Right margin" (Some "0") (Some "500") (sprintf "%g" initialMarginRight)
                dialogNumber "marginTop" "Top margin" (Some "0") (Some "500") (sprintf "%g" initialMarginTop)
                dialogNumber "marginBottom" "Bottom margin" (Some "0") (Some "500") (sprintf "%g" initialMarginBottom)
            ]
            Buttons = [
                dialogCancel
                dialogButton "dlg-page-setup-ok" "Apply" true (fun values ->
                    let parse k =
                        match System.Double.TryParse(Map.find k values) with
                        | true, v -> v
                        | _ -> nan
                    let orient = Map.find "orientation" values
                    let ml = parse "marginLeft"
                    let mr = parse "marginRight"
                    let mt = parse "marginTop"
                    let mb = parse "marginBottom"
                    if System.Double.IsNaN ml || System.Double.IsNaN mr || System.Double.IsNaN mt || System.Double.IsNaN mb then
                        false
                    else
                        apply (orient, ml, mr, mt, mb)
                        true)
            ]
        }

    /// Open a workbook from disk.
    let openFileDialog (apply: string -> unit) : DialogSpec = {
        TestId = "xl-dlg-open-file"
        Title = "Open workbook"
        Subtitle = Some "Path to an .xlsx file."
        Fields = [
            dialogText "path" "File path" "/path/to/book.xlsx" ""
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-open-file-ok" "Open" true (fun values ->
                let path = (Map.find "path" values).Trim()
                if path = "" then false
                else
                    apply path
                    true)
        ]
    }

    /// Save the current workbook to a path.
    let saveFileDialog (initialPath: string) (apply: string -> unit) : DialogSpec = {
        TestId = "xl-dlg-save-file"
        Title = "Save workbook"
        Subtitle = Some "Path to write an .xlsx file."
        Fields = [
            dialogText "path" "File path" "/path/to/book.xlsx" initialPath
        ]
        Buttons = [
            dialogCancel
            dialogButton "dlg-save-file-ok" "Save" true (fun values ->
                let path = (Map.find "path" values).Trim()
                if path = "" then false
                else
                    apply path
                    true)
        ]
    }
