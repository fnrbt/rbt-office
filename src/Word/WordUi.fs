namespace RedBat.Word

/// Word-specific UI specs. Generic types live in RedBat.Office.Ui; this module
/// supplies Word's chrome palette, menus, toolbar rows, and the
/// ShellState/ShellActions records that the shell renderers consume.
module Ui =
    open RedBat.Office.Ui

    // Re-export the generic spec types so consumers that look them up via
    // RedBat.Word.Ui keep working. RedBat.Office.Ui is the authoritative home.
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

    /// Helper re-exports so external code keeps using `WordUi.choice`.
    let choice = RedBat.Office.Ui.choice

    /// Generic chrome uses DocumentTrackWidth; Word's older paths called this
    /// the ruler track width.
    let inline rulerTrackWidth (spec: ChromeSpec) = spec.DocumentTrackWidth
    /// Older Word code referenced `chrome.RulerTrackWidth`. Alias kept for the
    /// existing call sites that read it through `WordUi.chrome.RulerTrackWidth`.
    let inline RulerTrackWidth (spec: ChromeSpec) = spec.DocumentTrackWidth

    let chrome : ChromeSpec = {
        WindowWidth = 1100.0
        MenuWidth = 280.0
        MenuGutterWidth = 24.0
        DocumentTrackWidth = 816.0
        // Flat, white-first palette. Top == Bottom on each band so the chrome
        // looks like a single sheet of paper rather than gradient banding.
        TitleTop = "#FFFFFF"
        TitleBottom = "#FFFFFF"
        Chrome = "#FFFFFF"
        ChromeTop = "#FFFFFF"
        ChromeBottom = "#FFFFFF"
        ChromeLight = "#FFFFFF"
        ChromeMid = "#E1E5EA"      // hairline separators
        ChromeDark = "#B9C0C8"     // subtler than before
        Pressed = "#E8F0FE"         // soft blue tint (Google Docs active)
        HotTop = "#F1F3F4"          // Google-grey hover
        HotBottom = "#F1F3F4"
        Input = "#FFFFFF"
        MenuPanel = "#FFFFFF"
        MenuGutter = "#FFFFFF"      // no gutter color difference
        PaperDesk = "#F8F9FA"       // very light desk
        Border = "#DADCE0"          // Google border grey
        Text = "#202124"            // Google text grey
    }

    let chromeCssVariables = [
        "--word-title-top", chrome.TitleTop
        "--word-title-bottom", chrome.TitleBottom
        "--word-chrome", chrome.Chrome
        "--word-chrome-top", chrome.ChromeTop
        "--word-chrome-bottom", chrome.ChromeBottom
        "--word-chrome-light", chrome.ChromeLight
        "--word-chrome-mid", chrome.ChromeMid
        "--word-chrome-dark", chrome.ChromeDark
        "--word-pressed", chrome.Pressed
        "--word-hot-top", chrome.HotTop
        "--word-hot-bottom", chrome.HotBottom
        "--word-input", chrome.Input
        "--word-menu-panel", chrome.MenuPanel
        "--word-menu-gutter", chrome.MenuGutter
        "--word-paper-desk", chrome.PaperDesk
        "--word-border", chrome.Border
        "--word-text", chrome.Text
    ]

    let menuLeft testId =
        match testId with
        | "word-file-menu" -> 3.0
        | "word-edit-menu" -> 36.0
        | "word-insert-menu" -> 72.0
        | "word-format-menu" -> 116.0
        | "word-table-menu" -> 169.0
        | _ -> 3.0

    let toolbarGlyph testId =
        match testId with
        | "word-find-replace" -> "search"
        | "word-insert-image" -> "picture"
        | "word-image-size" -> "zoom-100"
        | "word-page-setup" -> "document"
        | "word-header-footer" -> "document"
        | "word-insert-date-time" -> "date-time"
        | "word-insert-page-number" -> "document"
        | "word-insert-textbox" -> "document"
        | "word-textbox-style" -> "zoom-100"
        | "word-insert-caption" -> "document"
        | "word-insert-cross-reference" -> "link"
        | "word-bookmark" -> "link"
        | "word-content-control" -> "document"
        | "word-insert-field" -> "document"
        | "word-refresh-fields" -> "document"
        | "word-insert-equation" -> "document"
        | "word-revision" -> "document"
        | "word-track-changes" -> "document"
        | "word-next-revision" -> "document"
        | "word-previous-revision" -> "document"
        | "word-accept-revision" -> "document"
        | "word-reject-revision" -> "document"
        | "word-accept-all-revisions" -> "document"
        | "word-reject-all-revisions" -> "document"
        | "word-insert-link" -> "link"
        | "word-insert-comment" -> "document"
        | "word-insert-note" -> "document"
        | "word-insert-table" -> "table"
        | "word-insert-row-after" -> "row-add"
        | "word-insert-column-after" -> "column-add"
        | "word-delete-row" -> "row-delete"
        | "word-delete-column" -> "column-delete"
        | "word-table-cell-style" -> "table"
        | "word-print-pdf" -> "document"
        | "word-bold" -> "bold"
        | "word-italic" -> "italic"
        | "word-underline" -> "underline"
        | "word-all-caps" -> "all-caps"
        | "word-small-caps" -> "small-caps"
        | "word-strike" -> "strike"
        | "word-clear-formatting" -> "clear-formatting"
        | "word-superscript" -> "superscript"
        | "word-subscript" -> "subscript"
        | "word-text-color" -> "text-color"
        | "word-highlight-color" -> "highlight-color"
        | "word-align-left" -> "align-left"
        | "word-align-center" -> "align-center"
        | "word-align-right" -> "align-right"
        | "word-align-justify" -> "align-justify"
        | "word-bullet-list" -> "bullet-list"
        | "word-numbered-list" -> "numbered-list"
        | "word-decrease-indent" -> "indent-decrease"
        | "word-increase-indent" -> "indent-increase"
        | "word-heading-1" -> "heading-1"
        | "word-heading-2" -> "heading-2"
        | "word-heading-3" -> "heading-3"
        | "word-normal-paragraph" -> "normal"
        | "word-paragraph-style" -> "paragraph"
        | _ -> "text"

    let showToolbarPressed testId isActive =
        isActive
        && testId <> "word-normal-paragraph"
        && testId <> "word-align-left"

    type ShellState = {
        IsBoldActive: unit -> bool
        IsItalicActive: unit -> bool
        IsUnderlineActive: unit -> bool
        IsAllCapsActive: unit -> bool
        IsSmallCapsActive: unit -> bool
        IsStrikeActive: unit -> bool
        Alignment: unit -> string
        BlockStyle: unit -> string
        ListStyle: unit -> string
        NumberFormat: unit -> string
        VertAlign: unit -> string
        TextColor: unit -> string
        HighlightColor: unit -> string
        FontFamily: unit -> string
        FontSize: unit -> string
        IsTrackChangesActive: unit -> bool
        ParagraphStyles: unit -> ChoiceSpec list
        CharacterStyles: unit -> ChoiceSpec list
        /// Recently-opened file paths, newest first. The File menu displays
        /// these inline below Save. Empty list = no Recent section.
        RecentFiles: unit -> string list
        /// Open a path from the Recent list. Receives the absolute path.
        OpenRecentFile: string -> unit
    }

    type ShellActions = {
        OpenFile: unit -> unit
        SaveFile: unit -> unit
        PrintPdf: unit -> unit
        OpenFindReplace: unit -> unit
        OpenImagePicker: unit -> unit
        OpenImageSize: unit -> unit
        OpenTextBoxStyle: unit -> unit
        OpenPageSetup: unit -> unit
        OpenHeaderFooter: unit -> unit
        InsertDateTime: unit -> unit
        InsertPageNumber: unit -> unit
        InsertTextBox: unit -> unit
        OpenCaption: unit -> unit
        OpenCrossReference: unit -> unit
        OpenBookmark: unit -> unit
        OpenContentControl: unit -> unit
        OpenField: unit -> unit
        RefreshFields: unit -> unit
        OpenEquation: unit -> unit
        OpenRevision: unit -> unit
        ToggleTrackChanges: unit -> unit
        SelectNextRevision: unit -> unit
        SelectPreviousRevision: unit -> unit
        AcceptSelectedRevision: unit -> unit
        RejectSelectedRevision: unit -> unit
        AcceptAllRevisions: unit -> unit
        RejectAllRevisions: unit -> unit
        InsertSectionBreak: string -> unit
        OpenAnnotation: string -> unit
        OpenParagraphStyle: unit -> unit
        OpenTableCellStyle: unit -> unit
        OpenTableMargins: unit -> unit
        OpenInsertTable: unit -> unit
        InsertDefaultTable: unit -> unit
        InsertTableRow: unit -> unit
        InsertTableColumn: unit -> unit
        DeleteTableRow: unit -> unit
        DeleteTableColumn: unit -> unit
        ApplyHeading: int -> unit
        ApplyNormalParagraph: unit -> unit
        ApplyParagraphStyle: string -> unit
        ApplyCharacterStyle: string -> unit
        ToggleBullet: unit -> unit
        ToggleNumbered: unit -> unit
        ApplyNumberFormat: string -> unit
        IncreaseListIndent: unit -> unit
        DecreaseListIndent: unit -> unit
        ToggleBold: unit -> unit
        ToggleItalic: unit -> unit
        ToggleUnderline: unit -> unit
        ToggleAllCaps: unit -> unit
        ToggleSmallCaps: unit -> unit
        ToggleStrike: unit -> unit
        ClearFormatting: unit -> unit
        ApplyVertAlign: string -> unit
        ApplyAlignment: string -> unit
        ApplyFontFamily: string -> unit
        ApplyFontSize: string -> unit
        ApplyTextColor: string -> unit
        ApplyHighlight: string -> unit
    }

    let fontFamilies = [
        choice "Calibri" "Calibri"
        choice "Aptos" "Aptos"
        choice "Aptos Display" "Aptos Display"
        choice "Arial" "Arial"
        choice "Calibri Light" "Calibri Light"
        choice "Cambria" "Cambria"
        choice "Consolas" "Consolas"
        choice "Courier New" "Courier New"
        choice "Georgia" "Georgia"
        choice "Helvetica" "Helvetica"
        choice "Inter" "Inter"
        choice "Segoe UI" "Segoe UI"
        choice "Times New Roman" "Times New Roman"
    ]

    let paragraphStyles = [
        choice "normal" "Normal"
        choice "heading-1" "Heading 1"
        choice "heading-2" "Heading 2"
        choice "heading-3" "Heading 3"
    ]

    let numberFormats = [
        choice "decimal" "1, 2, 3"
        choice "lowerLetter" "a, b, c"
        choice "upperLetter" "A, B, C"
        choice "lowerRoman" "i, ii, iii"
        choice "upperRoman" "I, II, III"
    ]

    let private item = menuItem
    let private itemKey = menuItemShortcut
    let private button = toolbarButton

    let private recentItems (state: ShellState) (actions: ShellActions) : MenuItemSpec list =
        let basename (path: string) =
            let i = max (path.LastIndexOf '/') (path.LastIndexOf '\\')
            if i < 0 then path else path.Substring(i + 1)
        state.RecentFiles ()
        |> List.truncate 6
        |> List.mapi (fun i path ->
            item (sprintf "word-menu-recent-%d" i) (basename path) None (fun () -> state.OpenRecentFile path))

    let menus state actions = [
        {
            TestId = "word-file-menu"
            PanelTestId = "word-file-menu-panel"
            Label = "File"
            Items = [
                itemKey "word-menu-open" "Open..." "Ctrl+O" None actions.OpenFile
                itemKey "word-menu-save" "Save" "Ctrl+S" None actions.SaveFile
                yield! recentItems state actions
                item "word-menu-page-setup" "Page setup..." None actions.OpenPageSetup
                item "word-menu-header-footer" "Headers and footers..." None actions.OpenHeaderFooter
                item "word-menu-print-pdf" "Print PDF" None actions.PrintPdf
            ]
        }
        {
            TestId = "word-edit-menu"
            PanelTestId = "word-edit-menu-panel"
            Label = "Edit"
            Items = [ itemKey "word-menu-find-replace" "Find and replace..." "Ctrl+F" None actions.OpenFindReplace ]
        }
        {
            TestId = "word-insert-menu"
            PanelTestId = "word-insert-menu-panel"
            Label = "Insert"
            Items = [
                item "word-menu-insert-image" "Image..." None actions.OpenImagePicker
                item "word-menu-insert-table" "Table..." None actions.OpenInsertTable
                itemKey "word-menu-insert-link" "Hyperlink..." "Ctrl+K" None (fun () -> actions.OpenAnnotation "hyperlink")
                item "word-menu-insert-comment" "Comment..." None (fun () -> actions.OpenAnnotation "comment")
                item "word-menu-insert-note" "Footnote / endnote..." None (fun () -> actions.OpenAnnotation "note")
                item "word-menu-insert-date-time" "Date and time" None actions.InsertDateTime
                item "word-menu-insert-page-number" "Page number" None actions.InsertPageNumber
                item "word-menu-insert-textbox" "Text box" None actions.InsertTextBox
                item "word-menu-insert-caption" "Caption..." None actions.OpenCaption
                item "word-menu-bookmark" "Bookmark..." None actions.OpenBookmark
                item "word-menu-insert-field" "Field..." None actions.OpenField
                item "word-menu-insert-equation" "Equation..." None actions.OpenEquation
                item "word-menu-section-break-next" "Section break" None (fun () -> actions.InsertSectionBreak "nextPage")
            ]
        }
        {
            TestId = "word-format-menu"
            PanelTestId = "word-format-menu-panel"
            Label = "Format"
            Items = [
                item "word-menu-heading-1" "Heading 1" (Some (fun () -> state.BlockStyle() = "heading-1")) (fun () -> actions.ApplyHeading 1)
                item "word-menu-heading-2" "Heading 2" (Some (fun () -> state.BlockStyle() = "heading-2")) (fun () -> actions.ApplyHeading 2)
                item "word-menu-heading-3" "Heading 3" (Some (fun () -> state.BlockStyle() = "heading-3")) (fun () -> actions.ApplyHeading 3)
                item "word-menu-normal-paragraph" "Normal paragraph" (Some (fun () -> state.BlockStyle() = "normal")) actions.ApplyNormalParagraph
                item "word-menu-bullet-list" "Bullet list" (Some (fun () -> state.ListStyle() = "bullet")) actions.ToggleBullet
                item "word-menu-numbered-list" "Numbered list" (Some (fun () -> state.ListStyle() = "numbered")) actions.ToggleNumbered
                item "word-menu-increase-indent" "Increase indent" None actions.IncreaseListIndent
                item "word-menu-decrease-indent" "Decrease indent" None actions.DecreaseListIndent
                item "word-menu-all-caps" "All caps" (Some state.IsAllCapsActive) actions.ToggleAllCaps
                item "word-menu-small-caps" "Small caps" (Some state.IsSmallCapsActive) actions.ToggleSmallCaps
                item "word-menu-strike" "Strikethrough" (Some state.IsStrikeActive) actions.ToggleStrike
                item "word-menu-superscript" "Superscript" (Some (fun () -> state.VertAlign() = "superscript")) (fun () -> actions.ApplyVertAlign "superscript")
                item "word-menu-subscript" "Subscript" (Some (fun () -> state.VertAlign() = "subscript")) (fun () -> actions.ApplyVertAlign "subscript")
                item "word-menu-clear-formatting" "Clear formatting" None actions.ClearFormatting
                item "word-menu-paragraph-style" "Paragraph style..." None actions.OpenParagraphStyle
            ]
        }
        {
            TestId = "word-table-menu"
            PanelTestId = "word-table-menu-panel"
            Label = "Table"
            Items = [
                item "word-menu-insert-table" "Insert 2 x 3 table" None actions.InsertDefaultTable
                item "word-menu-insert-custom-table" "Insert table..." None actions.OpenInsertTable
                item "word-menu-insert-row-after" "Insert row after" None actions.InsertTableRow
                item "word-menu-insert-column-after" "Insert column after" None actions.InsertTableColumn
                item "word-menu-delete-row" "Delete row" None actions.DeleteTableRow
                item "word-menu-delete-column" "Delete column" None actions.DeleteTableColumn
                item "word-menu-table-cell-style" "Cell style..." None actions.OpenTableCellStyle
                item "word-menu-table-margins" "Cell margins..." None actions.OpenTableMargins
            ]
        }
    ]

    let toolbar state actions = [
        button "word-bold" "B" (Some state.IsBoldActive) actions.ToggleBold |> withTooltip "Bold (Ctrl+B)"
        button "word-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic |> withTooltip "Italic (Ctrl+I)"
        button "word-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline |> withTooltip "Underline (Ctrl+U)"
        button "word-all-caps" "AA" (Some state.IsAllCapsActive) actions.ToggleAllCaps |> withTooltip "All caps"
        button "word-small-caps" "Aa" (Some state.IsSmallCapsActive) actions.ToggleSmallCaps |> withTooltip "Small caps"
        button "word-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike |> withTooltip "Strikethrough"
        button "word-clear-formatting" "Clear" None actions.ClearFormatting |> withTooltip "Clear formatting"
        button "word-superscript" "Sup" (Some (fun () -> state.VertAlign() = "superscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "superscript" then "baseline" else "superscript")) |> withTooltip "Superscript"
        button "word-subscript" "Sub" (Some (fun () -> state.VertAlign() = "subscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "subscript" then "baseline" else "subscript")) |> withTooltip "Subscript"
        button "word-align-left" "Left" (Some (fun () -> state.Alignment() = "left")) (fun () -> actions.ApplyAlignment "left") |> withTooltip "Align left (Ctrl+Shift+L)"
        button "word-align-center" "Center" (Some (fun () -> state.Alignment() = "center")) (fun () -> actions.ApplyAlignment "center") |> withTooltip "Align center (Ctrl+Shift+E)"
        button "word-align-right" "Right" (Some (fun () -> state.Alignment() = "right")) (fun () -> actions.ApplyAlignment "right") |> withTooltip "Align right (Ctrl+Shift+R)"
        button "word-align-justify" "Justify" (Some (fun () -> state.Alignment() = "justify")) (fun () -> actions.ApplyAlignment "justify") |> withTooltip "Justify (Ctrl+Shift+J)"
        button "word-heading-1" "H1" (Some (fun () -> state.BlockStyle() = "heading-1")) (fun () -> actions.ApplyHeading 1) |> withTooltip "Heading 1"
        button "word-heading-2" "H2" (Some (fun () -> state.BlockStyle() = "heading-2")) (fun () -> actions.ApplyHeading 2) |> withTooltip "Heading 2"
        button "word-heading-3" "H3" (Some (fun () -> state.BlockStyle() = "heading-3")) (fun () -> actions.ApplyHeading 3) |> withTooltip "Heading 3"
        button "word-normal-paragraph" "Normal" (Some (fun () -> state.BlockStyle() = "normal")) actions.ApplyNormalParagraph |> withTooltip "Normal paragraph"
        button "word-bullet-list" "Bullet" (Some (fun () -> state.ListStyle() = "bullet")) actions.ToggleBullet |> withTooltip "Bulleted list"
        button "word-numbered-list" "Numbered" (Some (fun () -> state.ListStyle() = "numbered")) actions.ToggleNumbered |> withTooltip "Numbered list"
        button "word-decrease-indent" "Out" None actions.DecreaseListIndent |> withTooltip "Decrease indent"
        button "word-increase-indent" "In" None actions.IncreaseListIndent |> withTooltip "Increase indent"
        toolbarSelect "word-font-family" state.FontFamily (fun () -> fontFamilies) actions.ApplyFontFamily
        toolbarNumberInput "word-font-size" state.FontSize (Some "8") (Some "72") actions.ApplyFontSize
        toolbarColorInput "word-text-color" state.TextColor actions.ApplyTextColor
        toolbarColorInput "word-highlight-color" state.HighlightColor actions.ApplyHighlight
        button "word-print-pdf" "PDF" None actions.PrintPdf |> withTooltip "Export to PDF"
        button "word-find-replace" "Find" None actions.OpenFindReplace |> withTooltip "Find and replace (Ctrl+F)"
        button "word-insert-image" "Image" None actions.OpenImagePicker |> withTooltip "Insert image"
        button "word-image-size" "Img size" None actions.OpenImageSize
        button "word-insert-date-time" "Date" None actions.InsertDateTime
        button "word-insert-page-number" "Page #" None actions.InsertPageNumber
        button "word-insert-textbox" "Text box" None actions.InsertTextBox
        button "word-textbox-style" "Box style" None actions.OpenTextBoxStyle
        button "word-insert-caption" "Caption" None actions.OpenCaption
        button "word-insert-cross-reference" "X-ref" None actions.OpenCrossReference
        button "word-bookmark" "Bookmark" None actions.OpenBookmark
        button "word-content-control" "Control" None actions.OpenContentControl
        button "word-insert-field" "Field" None actions.OpenField
        button "word-refresh-fields" "Refresh fields" None actions.RefreshFields
        button "word-insert-equation" "Equation" None actions.OpenEquation
        button "word-revision" "Revision" None actions.OpenRevision
        button "word-track-changes" "Track" (Some state.IsTrackChangesActive) actions.ToggleTrackChanges
        button "word-next-revision" "Next rev" None actions.SelectNextRevision
        button "word-previous-revision" "Prev rev" None actions.SelectPreviousRevision
        button "word-accept-revision" "Accept rev" None actions.AcceptSelectedRevision
        button "word-reject-revision" "Reject rev" None actions.RejectSelectedRevision
        button "word-accept-all-revisions" "Accept all" None actions.AcceptAllRevisions
        button "word-reject-all-revisions" "Reject all" None actions.RejectAllRevisions
        button "word-insert-link" "Link" None (fun () -> actions.OpenAnnotation "hyperlink")
        button "word-insert-comment" "Comment" None (fun () -> actions.OpenAnnotation "comment")
        button "word-insert-note" "Note" None (fun () -> actions.OpenAnnotation "note")
        button "word-insert-table" "Table" None actions.InsertDefaultTable
        button "word-insert-row-after" "+Row" None actions.InsertTableRow
        button "word-insert-column-after" "+Col" None actions.InsertTableColumn
        button "word-delete-row" "-Row" None actions.DeleteTableRow
        button "word-delete-column" "-Col" None actions.DeleteTableColumn
        button "word-table-cell-style" "Cell" None actions.OpenTableCellStyle
        button "word-table-margins" "Margins" None actions.OpenTableMargins
        button "word-paragraph-style" "Paragraph" None actions.OpenParagraphStyle
    ]

    let private sep = toolbarSep

    /// Single, focused toolbar. Niche actions live in menus.
    let toolbarRows state actions = [
        { Label = "Main"; Controls = [
            toolbarSelect "word-paragraph-style-select" state.BlockStyle state.ParagraphStyles (fun value ->
                match value with
                | "heading-1" -> actions.ApplyHeading 1
                | "heading-2" -> actions.ApplyHeading 2
                | "heading-3" -> actions.ApplyHeading 3
                | "normal" -> actions.ApplyNormalParagraph ()
                | _ -> actions.ApplyParagraphStyle value)
            toolbarSelect "word-font-family" state.FontFamily (fun () -> fontFamilies) actions.ApplyFontFamily
            toolbarNumberInput "word-font-size" state.FontSize (Some "8") (Some "72") actions.ApplyFontSize
            sep
            button "word-bold" "B" (Some state.IsBoldActive) actions.ToggleBold
            button "word-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic
            button "word-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline
            button "word-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike
            sep
            toolbarColorInput "word-text-color" state.TextColor actions.ApplyTextColor
            toolbarColorInput "word-highlight-color" state.HighlightColor actions.ApplyHighlight
            sep
            button "word-align-left" "L" (Some (fun () -> state.Alignment() = "left")) (fun () -> actions.ApplyAlignment "left")
            button "word-align-center" "C" (Some (fun () -> state.Alignment() = "center")) (fun () -> actions.ApplyAlignment "center")
            button "word-align-right" "R" (Some (fun () -> state.Alignment() = "right")) (fun () -> actions.ApplyAlignment "right")
            button "word-align-justify" "J" (Some (fun () -> state.Alignment() = "justify")) (fun () -> actions.ApplyAlignment "justify")
            sep
            button "word-bullet-list" "•" (Some (fun () -> state.ListStyle() = "bullet")) actions.ToggleBullet
            button "word-numbered-list" "#" (Some (fun () -> state.ListStyle() = "numbered")) actions.ToggleNumbered
            button "word-decrease-indent" "<" None actions.DecreaseListIndent
            button "word-increase-indent" ">" None actions.IncreaseListIndent
            sep
            button "word-insert-image" "Img" None actions.OpenImagePicker
            button "word-insert-table" "Tbl" None actions.InsertDefaultTable
            button "word-insert-link" "Lnk" None (fun () -> actions.OpenAnnotation "hyperlink")
            button "word-insert-comment" "Cmt" None (fun () -> actions.OpenAnnotation "comment")
            sep
            button "word-find-replace" "Q" None actions.OpenFindReplace
        ] }
    ]

    // ── Dialog specs ─────────────────────────────────────────
    //
    // Backend-agnostic dialog descriptors. The app passes in callbacks that
    // perform the underlying document mutation; the spec describes the form
    // and binds the buttons to those callbacks.

    /// Find / Replace.
    ///
    /// `onFindNext q matchCase`      → status string to show (e.g. "1 of 5").
    /// `onReplaceAll q r matchCase`  → status string (e.g. "Replaced 3").
    let findReplaceDialog (onFindNext: string -> bool -> string) (onReplaceAll: string -> string -> bool -> string) : DialogSpec =
        let statusFieldId = "result"
        {
            TestId = "word-dlg-find-replace"
            Title = "Find and replace"
            Subtitle = None
            Fields = [
                dialogText "find" "Find" "Search text" ""
                dialogText "replace" "Replace with" "Replacement text" ""
                dialogCheckbox "case" "Match case" false
                dialogStatus statusFieldId "Result" ""
            ]
            Tabs = []
            Buttons = [
                dialogCancel
                dialogButton "dlg-find-next" "Find next" false (fun values ->
                    let q = Map.find "find" values
                    let mc = (Map.find "case" values) = "true"
                    let _ = onFindNext q mc
                    // keep open (status updated externally via callback writing to the renderer)
                    false)
                dialogButton "dlg-replace-all" "Replace all" true (fun values ->
                    let q = Map.find "find" values
                    let r = Map.find "replace" values
                    let mc = (Map.find "case" values) = "true"
                    let _ = onReplaceAll q r mc
                    true)
            ]
        }

    /// Insert hyperlink. Submit invokes `apply` with the URL value.
    let insertLinkDialog (apply: string -> unit) : DialogSpec = {
        TestId = "word-dlg-insert-link"
        Title = "Insert link"
        Subtitle = Some "The URL is applied to the current selection."
        Fields = [
            dialogText "url" "URL" "https://" ""
        ]
        Tabs = []
        Buttons = [
            dialogCancel
            dialogButton "dlg-insert-link-ok" "Insert" true (fun values ->
                apply (Map.find "url" values)
                true)
        ]
    }

    /// Insert table. Rows and columns are integers, defaults 2x3.
    let insertTableDialog (apply: int -> int -> unit) : DialogSpec = {
        TestId = "word-dlg-insert-table"
        Title = "Insert table"
        Subtitle = None
        Fields = [
            dialogNumber "rows" "Rows" (Some "1") (Some "100") "2"
            dialogNumber "cols" "Columns" (Some "1") (Some "20") "3"
        ]
        Tabs = []
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

    /// Page setup. Values are in points; margins typically 36..144 (0.5..2 inches).
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
            TestId = "word-dlg-page-setup"
            Title = "Page setup"
            Subtitle = Some "Margins are in points (72 pt = 1 inch)."
            Fields = [
                dialogSelect "orientation" "Orientation" orientations initialOrientation
                dialogNumber "marginLeft" "Left margin" (Some "0") (Some "500") (sprintf "%g" initialMarginLeft)
                dialogNumber "marginRight" "Right margin" (Some "0") (Some "500") (sprintf "%g" initialMarginRight)
                dialogNumber "marginTop" "Top margin" (Some "0") (Some "500") (sprintf "%g" initialMarginTop)
                dialogNumber "marginBottom" "Bottom margin" (Some "0") (Some "500") (sprintf "%g" initialMarginBottom)
            ]
            Tabs = []
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

    /// Insert image. v1 takes only a URL / src string; native side may
    /// later swap this for a file picker.
    let insertImageDialog (apply: string -> unit) : DialogSpec = {
        TestId = "word-dlg-insert-image"
        Title = "Insert image"
        Subtitle = Some "Paste a file path or URL."
        Fields = [
            dialogText "src" "Image source" "/path/or/url" ""
        ]
        Tabs = []
        Buttons = [
            dialogCancel
            dialogButton "dlg-insert-image-ok" "Insert" true (fun values ->
                let src = (Map.find "src" values).Trim()
                if src = "" then false
                else
                    apply src
                    true)
        ]
    }

    /// Open a document from disk. Native shells pass the path to a loader;
    /// web shells can override the action entirely with a file picker.
    let openFileDialog (apply: string -> unit) : DialogSpec = {
        TestId = "word-dlg-open-file"
        Title = "Open document"
        Subtitle = Some "Path to a .docx file."
        Fields = [
            dialogText "path" "File path" "/path/to/doc.docx" ""
        ]
        Tabs = []
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

    /// Save the current document to a path. Initial value is the current
    /// document's path (or empty for a new doc).
    let saveFileDialog (initialPath: string) (apply: string -> unit) : DialogSpec = {
        TestId = "word-dlg-save-file"
        Title = "Save document"
        Subtitle = Some "Path to write a .docx file."
        Fields = [
            dialogText "path" "File path" "/path/to/doc.docx" initialPath
        ]
        Tabs = []
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

    // ── Advanced widget showcase dialogs ─────────────────────────
    //
    // The following builders exist primarily to demonstrate the new dialog
    // field kinds (section header, color, slider, radio, tabs, text area)
    // and to give the editor a more professional surface area. The submit
    // callbacks just record the picked values via `apply`; the host
    // decides what to do with them (typically a status-line confirmation).

    /// Application preferences. Tabbed, with one tab per concern area:
    /// General, Editor, Display, Files. Demonstrates the full kit —
    /// sections, sliders, color picker, radio buttons, text area, plus
    /// the existing select/checkbox/text inputs.
    let preferencesDialog (apply: Map<string, string> -> unit) : DialogSpec =
        let themes = [ choice "light" "Light"; choice "dark" "Dark"; choice "auto" "Match system" ]
        let saveModes = [
            choice "docx"  "Word document (.docx)"
            choice "pdf"   "PDF (read-only)"
            choice "html"  "Web page (.html)"
            choice "txt"   "Plain text (.txt)"
        ]
        let backupChoices = [
            choice "every-edit" "Snapshot every edit"
            choice "every-min"  "Every minute"
            choice "every-5min" "Every five minutes"
            choice "off"        "Off"
        ]
        {
            TestId = "word-dlg-preferences"
            Title = "Preferences"
            Subtitle = Some "Configure how the editor looks and behaves."
            Fields = []
            Tabs = [
                "General", [
                    dialogSection "Identity"
                    dialogText "name" "Your name" "Jane Doe" ""
                    dialogText "initials" "Initials" "JD" ""
                    dialogText "email" "Email" "jane@example.com" ""
                    dialogSection "Behaviour"
                    dialogCheckbox "spellcheck" "Check spelling as you type" true
                    dialogCheckbox "autocorrect" "Auto-correct common typos" true
                    dialogCheckbox "autocomplete" "Suggest completions while typing" false
                    dialogSection "Recents"
                    dialogSlider "recent-count" "Number of recent documents" 0.0 30.0 12.0
                ]
                "Appearance", [
                    dialogSection "Theme"
                    dialogRadio "theme" "Theme" themes "light"
                    dialogSection "Accent"
                    dialogColor "accent-color" "Accent colour" "#1A73E8"
                    dialogColor "highlight-color" "Default highlight" "#FFF2CC"
                    dialogSection "Type"
                    dialogSelect "ui-font" "Interface font" [
                        choice "calibri" "Calibri"
                        choice "segoe-ui" "Segoe UI"
                        choice "ibm-plex" "IBM Plex Sans"
                        choice "system" "System default"
                    ] "calibri"
                    dialogSlider "ui-font-size" "Interface font size (pt)" 9.0 18.0 11.0
                ]
                "Editor", [
                    dialogSection "Composition"
                    dialogCheckbox "smart-quotes" "Use smart quotes" true
                    dialogCheckbox "smart-dashes" "Convert -- to em-dash" true
                    dialogCheckbox "auto-bullets" "Auto-start bullet lists from - or *" true
                    dialogCheckbox "show-invisibles" "Show paragraph marks and tabs" false
                    dialogSection "Defaults"
                    dialogSlider "default-font-size" "Default font size (pt)" 8.0 24.0 12.0
                    dialogSlider "line-spacing" "Default line spacing (×)" 1.0 3.0 1.15
                    dialogSlider "para-spacing" "Default spacing after paragraph (pt)" 0.0 24.0 8.0
                    dialogSection "Comments"
                    dialogTextArea "comment-template" "Comment signature" "-- Jane Doe" 3 ""
                ]
                "Files", [
                    dialogSection "Save"
                    dialogSelect "default-format" "Default save format" saveModes "docx"
                    dialogText "autosave-folder" "Auto-save folder" "~/Documents/Word" "~/Documents/Word"
                    dialogCheckbox "save-on-blur" "Save when the window loses focus" false
                    dialogSection "Backup"
                    dialogRadio "backup-mode" "Backup snapshots" backupChoices "every-min"
                    dialogSlider "backup-keep" "Maximum backups to retain" 1.0 100.0 20.0
                    dialogSection "Recovery"
                    dialogCheckbox "recover-on-open" "Offer recovery on next open after crash" true
                ]
            ]
            Buttons = [
                dialogCancel
                dialogButton "dlg-prefs-apply" "Apply" true (fun values ->
                    apply values
                    true)
            ]
        }

    /// Paragraph formatting — alignment / indents / spacing / line break
    /// behaviour. Demonstrates radio groups, sliders, and section
    /// headers grouping logically related fields.
    let paragraphDialog (apply: Map<string, string> -> unit) : DialogSpec =
        let alignments = [
            choice "left"    "Left"
            choice "center"  "Centre"
            choice "right"   "Right"
            choice "justify" "Justify"
        ]
        let lineRules = [
            choice "single"   "Single"
            choice "1.15"     "1.15"
            choice "1.5"      "1.5"
            choice "double"   "Double"
            choice "exactly"  "Exactly (pt)"
            choice "multiple" "Multiple (×)"
        ]
        let outlineLevels = [
            choice "body"     "Body text"
            choice "level1"   "Level 1"
            choice "level2"   "Level 2"
            choice "level3"   "Level 3"
            choice "level4"   "Level 4"
        ]
        {
            TestId = "word-dlg-paragraph"
            Title = "Paragraph"
            Subtitle = Some "Indent, spacing, and line-break behaviour for the selected paragraphs."
            Fields = [
                dialogSection "Alignment"
                dialogRadio "alignment" "Alignment" alignments "left"
                dialogSelect "outline-level" "Outline level" outlineLevels "body"
                dialogSection "Indentation"
                dialogSlider "indent-left"  "Left indent (pt)"  0.0 144.0 0.0
                dialogSlider "indent-right" "Right indent (pt)" 0.0 144.0 0.0
                dialogSlider "indent-first" "First-line indent (pt)" 0.0 72.0 0.0
                dialogCheckbox "mirror-indents" "Mirror indents on facing pages" false
                dialogSection "Spacing"
                dialogSlider "space-before" "Before (pt)" 0.0 36.0 0.0
                dialogSlider "space-after"  "After  (pt)" 0.0 36.0 8.0
                dialogSelect "line-rule"    "Line spacing" lineRules "1.15"
                dialogSlider "line-amount"  "Line spacing amount" 1.0 4.0 1.15
                dialogSection "Pagination"
                dialogCheckbox "keep-with-next" "Keep with next paragraph" false
                dialogCheckbox "keep-together"  "Keep lines together" false
                dialogCheckbox "page-before"    "Page break before" false
                dialogCheckbox "widow-orphan"   "Widow/orphan control" true
            ]
            Tabs = []
            Buttons = [
                dialogCancel
                dialogButton "dlg-paragraph-apply" "Apply" true (fun values ->
                    apply values
                    true)
            ]
        }

    /// Character formatting — font, weight, colour, position, effects.
    /// Showcases the color picker alongside sliders for tracking +
    /// position offsets, plus toggles for visual effects.
    let characterDialog (apply: Map<string, string> -> unit) : DialogSpec =
        let weights = [
            choice "regular"    "Regular"
            choice "light"      "Light"
            choice "medium"     "Medium"
            choice "semibold"   "Semi-bold"
            choice "bold"       "Bold"
            choice "black"      "Black"
        ]
        let underlines = [
            choice "none"   "(none)"
            choice "single" "Single"
            choice "double" "Double"
            choice "thick"  "Thick"
            choice "dotted" "Dotted"
            choice "wavy"   "Wavy"
        ]
        {
            TestId = "word-dlg-character"
            Title = "Character"
            Subtitle = Some "Font, weight, colour, and visual effects for the selected text."
            Fields = [
                dialogSection "Font"
                dialogSelect "font-family" "Family" fontFamilies "Calibri"
                dialogSelect "font-weight" "Weight" weights "regular"
                dialogSlider "font-size" "Size (pt)" 6.0 96.0 11.0
                dialogSection "Colour"
                dialogColor "font-color" "Text colour" "#000000"
                dialogColor "highlight" "Highlight" "#FFFFFF"
                dialogSection "Decoration"
                dialogSelect "underline" "Underline" underlines "none"
                dialogColor "underline-color" "Underline colour" "#000000"
                dialogCheckbox "strike" "Strikethrough" false
                dialogCheckbox "double-strike" "Double strikethrough" false
                dialogCheckbox "small-caps" "Small caps" false
                dialogCheckbox "all-caps" "All caps" false
                dialogCheckbox "hidden" "Hidden text" false
                dialogSection "Position"
                dialogSlider "tracking" "Character spacing (1/100 em)" -20.0 50.0 0.0
                dialogSlider "raise" "Vertical position (pt, +up/-down)" -12.0 12.0 0.0
                dialogSlider "scale" "Horizontal scale (%)" 50.0 200.0 100.0
            ]
            Tabs = []
            Buttons = [
                dialogCancel
                dialogButton "dlg-character-apply" "Apply" true (fun values ->
                    apply values
                    true)
            ]
        }

    /// About box. Demo of section headers paired with the read-only
    /// status field for a polished "About" panel.
    let aboutDialog (version: string) (build: string) : DialogSpec =
        let creditsField = dialogStatus "credits" "Credits" ""
        let creditsLong =
            "RedBat Word native shell\n" +
            "  • Composit UI widgets\n" +
            "  • RedBat.Office.Ui spec layer\n" +
            "  • RedBat.Word document model\n\n" +
            "Built on Text.Engine, OpenType.Core, Bidi, Shaper.Core."
        {
            TestId = "word-dlg-about"
            Title = "About RedBat Word"
            Subtitle = Some "A native, F#-first Word-compatible editor."
            Fields = [
                dialogSection "Version"
                dialogStatus "version" "Version" version
                dialogStatus "build"   "Build"   build
                dialogSection "Project"
                dialogStatus "homepage" "Homepage" "https://example.com/redbat-word"
                dialogStatus "license"  "License"  "MIT"
                dialogSection "Components"
                { creditsField with Initial = creditsLong }
            ]
            Tabs = []
            Buttons = [
                dialogButton "dlg-about-close" "Close" true (fun _ -> true)
            ]
        }
