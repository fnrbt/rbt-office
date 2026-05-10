namespace RedBat.Word

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

    type InputKind =
        | NumberInput
        | ColorInput

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

    type ChromeSpec = {
        WindowWidth: float
        MenuWidth: float
        MenuGutterWidth: float
        RulerTrackWidth: float
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

    let chrome = {
        WindowWidth = 980.0
        MenuWidth = 190.0
        MenuGutterWidth = 30.0
        RulerTrackWidth = 900.0
        TitleTop = "#2159AA"
        TitleBottom = "#407ECE"
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
    }

    type ShellActions = {
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

    let choice value label = { Value = value; Label = label }

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

    let private item testId label isChecked action = {
        TestId = testId
        Label = label
        IsChecked = isChecked
        Action = action
    }

    let private button testId label active action =
        ToolbarButton { TestId = testId; Label = label; IsActive = active; Action = action }

    let menus state actions = [
        {
            TestId = "word-file-menu"
            PanelTestId = "word-file-menu-panel"
            Label = "File"
            Items = [
                item "word-menu-page-setup" "Page setup..." None actions.OpenPageSetup
                item "word-menu-header-footer" "Headers and footers..." None actions.OpenHeaderFooter
                item "word-menu-print-pdf" "Print PDF" None actions.PrintPdf
            ]
        }
        {
            TestId = "word-edit-menu"
            PanelTestId = "word-edit-menu-panel"
            Label = "Edit"
            Items = [ item "word-menu-find-replace" "Find and replace..." None actions.OpenFindReplace ]
        }
        {
            TestId = "word-insert-menu"
            PanelTestId = "word-insert-menu-panel"
            Label = "Insert"
            Items = [
                item "word-menu-insert-image" "Image..." None actions.OpenImagePicker
                item "word-menu-image-size" "Image size..." None actions.OpenImageSize
                item "word-menu-insert-date-time" "Date and time" None actions.InsertDateTime
                item "word-menu-insert-page-number" "Page number" None actions.InsertPageNumber
                item "word-menu-insert-textbox" "Text box" None actions.InsertTextBox
                item "word-menu-textbox-style" "Text box style..." None actions.OpenTextBoxStyle
                item "word-menu-insert-caption" "Caption..." None actions.OpenCaption
                item "word-menu-insert-cross-reference" "Cross-reference..." None actions.OpenCrossReference
                item "word-menu-bookmark" "Bookmark..." None actions.OpenBookmark
                item "word-menu-content-control" "Content control..." None actions.OpenContentControl
                item "word-menu-insert-field" "Field..." None actions.OpenField
                item "word-menu-insert-equation" "Equation..." None actions.OpenEquation
                item "word-menu-revision" "Revision..." None actions.OpenRevision
                item "word-menu-track-changes" "Track changes" (Some state.IsTrackChangesActive) actions.ToggleTrackChanges
                item "word-menu-next-revision" "Next revision" None actions.SelectNextRevision
                item "word-menu-previous-revision" "Previous revision" None actions.SelectPreviousRevision
                item "word-menu-accept-revision" "Accept selected revision" None actions.AcceptSelectedRevision
                item "word-menu-reject-revision" "Reject selected revision" None actions.RejectSelectedRevision
                item "word-menu-accept-all-revisions" "Accept all revisions" None actions.AcceptAllRevisions
                item "word-menu-reject-all-revisions" "Reject all revisions" None actions.RejectAllRevisions
                item "word-menu-section-break-next" "Section break: next page" None (fun () -> actions.InsertSectionBreak "nextPage")
                item "word-menu-section-break-continuous" "Section break: continuous" None (fun () -> actions.InsertSectionBreak "continuous")
                item "word-menu-insert-link" "Hyperlink..." None (fun () -> actions.OpenAnnotation "hyperlink")
                item "word-menu-insert-comment" "Comment..." None (fun () -> actions.OpenAnnotation "comment")
                item "word-menu-insert-note" "Footnote / endnote..." None (fun () -> actions.OpenAnnotation "note")
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
                item "word-menu-number-format-decimal" "Numbering: 1, 2, 3" (Some (fun () -> state.NumberFormat() = "decimal")) (fun () -> actions.ApplyNumberFormat "decimal")
                item "word-menu-number-format-lower-letter" "Numbering: a, b, c" (Some (fun () -> state.NumberFormat() = "lowerLetter")) (fun () -> actions.ApplyNumberFormat "lowerLetter")
                item "word-menu-number-format-upper-roman" "Numbering: I, II, III" (Some (fun () -> state.NumberFormat() = "upperRoman")) (fun () -> actions.ApplyNumberFormat "upperRoman")
                item "word-menu-increase-indent" "Increase indent" None actions.IncreaseListIndent
                item "word-menu-decrease-indent" "Decrease indent" None actions.DecreaseListIndent
                item "word-menu-font-georgia" "Georgia font" (Some (fun () -> state.FontFamily() = "Georgia")) (fun () -> actions.ApplyFontFamily "Georgia")
                item "word-menu-font-size-16" "16 pt" (Some (fun () -> state.FontSize() = "16")) (fun () -> actions.ApplyFontSize "16")
                item "word-menu-all-caps" "All caps" (Some state.IsAllCapsActive) actions.ToggleAllCaps
                item "word-menu-small-caps" "Small caps" (Some state.IsSmallCapsActive) actions.ToggleSmallCaps
                item "word-menu-strike" "Strikethrough" (Some state.IsStrikeActive) actions.ToggleStrike
                item "word-menu-clear-formatting" "Clear formatting" None actions.ClearFormatting
                item "word-menu-superscript" "Superscript" (Some (fun () -> state.VertAlign() = "superscript")) (fun () -> actions.ApplyVertAlign "superscript")
                item "word-menu-subscript" "Subscript" (Some (fun () -> state.VertAlign() = "subscript")) (fun () -> actions.ApplyVertAlign "subscript")
                item "word-menu-text-color-red" "Red text" (Some (fun () -> state.TextColor() = "#C00000")) (fun () -> actions.ApplyTextColor "#C00000")
                item "word-menu-highlight-yellow" "Yellow highlight" (Some (fun () -> state.HighlightColor() = "#FFF2CC")) (fun () -> actions.ApplyHighlight "#FFF2CC")
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
            ]
        }
    ]

    let toolbar state actions = [
        button "word-bold" "B" (Some state.IsBoldActive) actions.ToggleBold
        button "word-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic
        button "word-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline
        button "word-all-caps" "AA" (Some state.IsAllCapsActive) actions.ToggleAllCaps
        button "word-small-caps" "Aa" (Some state.IsSmallCapsActive) actions.ToggleSmallCaps
        button "word-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike
        button "word-clear-formatting" "Clear" None actions.ClearFormatting
        button "word-superscript" "Sup" (Some (fun () -> state.VertAlign() = "superscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "superscript" then "baseline" else "superscript"))
        button "word-subscript" "Sub" (Some (fun () -> state.VertAlign() = "subscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "subscript" then "baseline" else "subscript"))
        button "word-align-left" "Left" (Some (fun () -> state.Alignment() = "left")) (fun () -> actions.ApplyAlignment "left")
        button "word-align-center" "Center" (Some (fun () -> state.Alignment() = "center")) (fun () -> actions.ApplyAlignment "center")
        button "word-align-right" "Right" (Some (fun () -> state.Alignment() = "right")) (fun () -> actions.ApplyAlignment "right")
        button "word-align-justify" "Justify" (Some (fun () -> state.Alignment() = "justify")) (fun () -> actions.ApplyAlignment "justify")
        button "word-heading-1" "H1" (Some (fun () -> state.BlockStyle() = "heading-1")) (fun () -> actions.ApplyHeading 1)
        button "word-heading-2" "H2" (Some (fun () -> state.BlockStyle() = "heading-2")) (fun () -> actions.ApplyHeading 2)
        button "word-heading-3" "H3" (Some (fun () -> state.BlockStyle() = "heading-3")) (fun () -> actions.ApplyHeading 3)
        button "word-normal-paragraph" "Normal" (Some (fun () -> state.BlockStyle() = "normal")) actions.ApplyNormalParagraph
        button "word-bullet-list" "Bullet" (Some (fun () -> state.ListStyle() = "bullet")) actions.ToggleBullet
        button "word-numbered-list" "Numbered" (Some (fun () -> state.ListStyle() = "numbered")) actions.ToggleNumbered
        button "word-decrease-indent" "Out" None actions.DecreaseListIndent
        button "word-increase-indent" "In" None actions.IncreaseListIndent
        ToolbarSelect { TestId = "word-font-family"; Value = state.FontFamily; Options = (fun () -> fontFamilies); Action = actions.ApplyFontFamily }
        ToolbarInput { TestId = "word-font-size"; Kind = NumberInput; Value = state.FontSize; Min = Some "8"; Max = Some "72"; Action = actions.ApplyFontSize }
        ToolbarInput { TestId = "word-text-color"; Kind = ColorInput; Value = state.TextColor; Min = None; Max = None; Action = actions.ApplyTextColor }
        ToolbarInput { TestId = "word-highlight-color"; Kind = ColorInput; Value = state.HighlightColor; Min = None; Max = None; Action = actions.ApplyHighlight }
        button "word-print-pdf" "PDF" None actions.PrintPdf
        button "word-find-replace" "Find" None actions.OpenFindReplace
        button "word-insert-image" "Image" None actions.OpenImagePicker
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
        button "word-paragraph-style" "Paragraph" None actions.OpenParagraphStyle
    ]

    let toolbarRows state actions = [
        { Label = "Standard"; Controls = [
            button "word-find-replace" "Find" None actions.OpenFindReplace
            button "word-insert-image" "Picture" None actions.OpenImagePicker
            button "word-image-size" "Size" None actions.OpenImageSize
            button "word-page-setup" "Page" None actions.OpenPageSetup
            button "word-header-footer" "Head" None actions.OpenHeaderFooter
            button "word-insert-date-time" "Date" None actions.InsertDateTime
            button "word-insert-page-number" "Page #" None actions.InsertPageNumber
            button "word-insert-textbox" "Box" None actions.InsertTextBox
            button "word-textbox-style" "Box fmt" None actions.OpenTextBoxStyle
            button "word-insert-caption" "Caption" None actions.OpenCaption
            button "word-insert-cross-reference" "X-ref" None actions.OpenCrossReference
            button "word-bookmark" "Mark" None actions.OpenBookmark
            button "word-content-control" "Ctrl" None actions.OpenContentControl
            button "word-insert-field" "Field" None actions.OpenField
            button "word-refresh-fields" "Refresh" None actions.RefreshFields
            button "word-insert-equation" "Eq" None actions.OpenEquation
            button "word-revision" "Rev" None actions.OpenRevision
            button "word-track-changes" "Track" (Some state.IsTrackChangesActive) actions.ToggleTrackChanges
            button "word-next-revision" "Next" None actions.SelectNextRevision
            button "word-previous-revision" "Prev" None actions.SelectPreviousRevision
            button "word-accept-revision" "Accept" None actions.AcceptSelectedRevision
            button "word-reject-revision" "Reject" None actions.RejectSelectedRevision
            button "word-accept-all-revisions" "All+" None actions.AcceptAllRevisions
            button "word-reject-all-revisions" "All-" None actions.RejectAllRevisions
            button "word-insert-link" "Link" None (fun () -> actions.OpenAnnotation "hyperlink")
            button "word-insert-comment" "Comment" None (fun () -> actions.OpenAnnotation "comment")
            button "word-insert-note" "Note" None (fun () -> actions.OpenAnnotation "note")
            button "word-insert-table" "Table" None actions.InsertDefaultTable
            button "word-insert-row-after" "+Row" None actions.InsertTableRow
            button "word-insert-column-after" "+Col" None actions.InsertTableColumn
            button "word-delete-row" "-Row" None actions.DeleteTableRow
            button "word-delete-column" "-Col" None actions.DeleteTableColumn
            button "word-table-cell-style" "Cell" None actions.OpenTableCellStyle
            button "word-print-pdf" "PDF" None actions.PrintPdf
        ] }
        { Label = "Formatting"; Controls = [
            ToolbarSelect {
                TestId = "word-paragraph-style-select"
                Value = state.BlockStyle
                Options = state.ParagraphStyles
                Action = fun value ->
                    match value with
                    | "heading-1" -> actions.ApplyHeading 1
                    | "heading-2" -> actions.ApplyHeading 2
                    | "heading-3" -> actions.ApplyHeading 3
                    | "normal" -> actions.ApplyNormalParagraph ()
                    | _ -> actions.ApplyParagraphStyle value
            }
            ToolbarSelect {
                TestId = "word-character-style-select"
                Value = fun () -> ""
                Options = state.CharacterStyles
                Action = actions.ApplyCharacterStyle
            }
            ToolbarSelect { TestId = "word-font-family"; Value = state.FontFamily; Options = (fun () -> fontFamilies); Action = actions.ApplyFontFamily }
            ToolbarInput { TestId = "word-font-size"; Kind = NumberInput; Value = state.FontSize; Min = Some "8"; Max = Some "72"; Action = actions.ApplyFontSize }
            button "word-bold" "B" (Some state.IsBoldActive) actions.ToggleBold
            button "word-italic" "I" (Some state.IsItalicActive) actions.ToggleItalic
            button "word-underline" "U" (Some state.IsUnderlineActive) actions.ToggleUnderline
            button "word-all-caps" "AA" (Some state.IsAllCapsActive) actions.ToggleAllCaps
            button "word-small-caps" "Aa" (Some state.IsSmallCapsActive) actions.ToggleSmallCaps
            button "word-strike" "S" (Some state.IsStrikeActive) actions.ToggleStrike
            button "word-clear-formatting" "Clear" None actions.ClearFormatting
            button "word-superscript" "x²" (Some (fun () -> state.VertAlign() = "superscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "superscript" then "baseline" else "superscript"))
            button "word-subscript" "x₂" (Some (fun () -> state.VertAlign() = "subscript")) (fun () -> actions.ApplyVertAlign (if state.VertAlign() = "subscript" then "baseline" else "subscript"))
            ToolbarInput { TestId = "word-text-color"; Kind = ColorInput; Value = state.TextColor; Min = None; Max = None; Action = actions.ApplyTextColor }
            ToolbarInput { TestId = "word-highlight-color"; Kind = ColorInput; Value = state.HighlightColor; Min = None; Max = None; Action = actions.ApplyHighlight }
            button "word-align-left" "Left" (Some (fun () -> state.Alignment() = "left")) (fun () -> actions.ApplyAlignment "left")
            button "word-align-center" "Center" (Some (fun () -> state.Alignment() = "center")) (fun () -> actions.ApplyAlignment "center")
            button "word-align-right" "Right" (Some (fun () -> state.Alignment() = "right")) (fun () -> actions.ApplyAlignment "right")
            button "word-align-justify" "Justify" (Some (fun () -> state.Alignment() = "justify")) (fun () -> actions.ApplyAlignment "justify")
            button "word-bullet-list" "Bullet" (Some (fun () -> state.ListStyle() = "bullet")) actions.ToggleBullet
            button "word-numbered-list" "Numbered" (Some (fun () -> state.ListStyle() = "numbered")) actions.ToggleNumbered
            ToolbarSelect { TestId = "word-number-format"; Value = state.NumberFormat; Options = (fun () -> numberFormats); Action = actions.ApplyNumberFormat }
            button "word-decrease-indent" "Out" None actions.DecreaseListIndent
            button "word-increase-indent" "In" None actions.IncreaseListIndent
            button "word-heading-1" "H1" (Some (fun () -> state.BlockStyle() = "heading-1")) (fun () -> actions.ApplyHeading 1)
            button "word-heading-2" "H2" (Some (fun () -> state.BlockStyle() = "heading-2")) (fun () -> actions.ApplyHeading 2)
            button "word-heading-3" "H3" (Some (fun () -> state.BlockStyle() = "heading-3")) (fun () -> actions.ApplyHeading 3)
            button "word-normal-paragraph" "Normal" (Some (fun () -> state.BlockStyle() = "normal")) actions.ApplyNormalParagraph
            button "word-paragraph-style" "Paragraph" None actions.OpenParagraphStyle
        ] }
    ]
