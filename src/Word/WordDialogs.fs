namespace Rbt.Word

#if FABLE_COMPILER

module Dialogs =
    open System.Collections.Generic
    open Rbt.Office.Ui
    open Rbt.FableUI.Reactive

    /// String-valued controls let mixed values remain distinct from explicit
    /// "none" values until the user chooses a replacement.
    type FormatDialogState = {
        IsOpen: Signal<bool>
        Error: Signal<string>
        Fields: Signal<string> array
    }

    let formatDialogState count = {
        IsOpen = signal false
        Error = signal ""
        Fields = Array.init count (fun _ -> signal "")
    }

    type ParagraphStyleState = {
        IsOpen: Signal<bool>
        Alignment: Signal<string>
        IndentLeft: Signal<string>
        IndentRight: Signal<string>
        SpacingBefore: Signal<string>
        SpacingAfter: Signal<string>
        LineSpacing: Signal<string>
        BackgroundColor: Signal<string>
        FontSize: Signal<string>
        SpecialIndent: Signal<string>
        SpecialIndentBy: Signal<string>
    }

    type FindReplaceState = {
        IsOpen: Signal<bool>
        Query: Signal<string>
        Replacement: Signal<string>
        MatchCase: Signal<string>
        Result: Signal<string>
    }

    type TableCellStyleState = {
        IsOpen: Signal<bool>
        Error: Signal<string>
        BackgroundColor: Signal<string>
        ShadingPattern: Signal<string>
        ShadingForegroundColor: Signal<string>
        Width: Signal<string>
        BorderColor: Signal<string>
        BorderWidth: Signal<string>
        VerticalAlignment: Signal<string>
        PaddingTop: Signal<string>
        PaddingRight: Signal<string>
        PaddingBottom: Signal<string>
        PaddingLeft: Signal<string>
    }

    type TableMarginsState = {
        IsOpen: Signal<bool>
        Error: Signal<string>
        PaddingTop: Signal<string>
        PaddingRight: Signal<string>
        PaddingBottom: Signal<string>
        PaddingLeft: Signal<string>
        CellSpacing: Signal<string>
        Fields: Signal<string> array
    }

    type InsertTableState = {
        IsOpen: Signal<bool>
        Error: Signal<string>
        Rows: Signal<string>
        Columns: Signal<string>
    }

    type ImageSizeState = {
        IsOpen: Signal<bool>
        Width: Signal<string>
        Height: Signal<string>
    }

    type TextBoxStyleState = {
        IsOpen: Signal<bool>
        X: Signal<string>
        Y: Signal<string>
        Width: Signal<string>
        Height: Signal<string>
        FillColor: Signal<string>
        BorderColor: Signal<string>
        BorderWidth: Signal<string>
    }

    type CaptionState = {
        IsOpen: Signal<bool>
        Label: Signal<string>
        Text: Signal<string>
    }

    type CrossReferenceState = {
        IsOpen: Signal<bool>
        Label: Signal<string>
    }

    type BookmarkState = {
        IsOpen: Signal<bool>
        Name: Signal<string>
    }

    type ContentControlState = {
        IsOpen: Signal<bool>
        Id: Signal<string>
        Tag: Signal<string>
        Alias: Signal<string>
    }

    type FieldState = {
        IsOpen: Signal<bool>
        Code: Signal<string>
        Result: Signal<string>
    }

    type EquationState = {
        IsOpen: Signal<bool>
        Text: Signal<string>
    }

    type RevisionState = {
        IsOpen: Signal<bool>
        RevisionType: Signal<string>
        RevisionId: Signal<string>
        Author: Signal<string>
        Date: Signal<string>
    }

    type PageSetupState = {
        IsOpen: Signal<bool>
        Preset: Signal<string>
        Orientation: Signal<string>
        Error: Signal<string>
        PageWidth: Signal<string>
        PageHeight: Signal<string>
        MarginLeft: Signal<string>
        MarginRight: Signal<string>
        MarginTop: Signal<string>
        MarginBottom: Signal<string>
        HeaderDistance: Signal<string>
        FooterDistance: Signal<string>
        DifferentFirstPage: Signal<string>
        EvenAndOddHeaders: Signal<string>
    }

    type HeaderFooterState = {
        IsOpen: Signal<bool>
        DefaultHeader: Signal<string>
        DefaultFooter: Signal<string>
        DefaultHeaderTableLeft: Signal<string>
        DefaultHeaderTableRight: Signal<string>
        DefaultFooterTableLeft: Signal<string>
        DefaultFooterTableRight: Signal<string>
        DefaultHeaderImageSrc: Signal<string>
        DefaultHeaderImageAlt: Signal<string>
        DefaultHeaderImageWidth: Signal<string>
        DefaultHeaderImageHeight: Signal<string>
        DefaultFooterImageSrc: Signal<string>
        DefaultFooterImageAlt: Signal<string>
        DefaultFooterImageWidth: Signal<string>
        DefaultFooterImageHeight: Signal<string>
        FirstHeader: Signal<string>
        FirstFooter: Signal<string>
        EvenHeader: Signal<string>
        EvenFooter: Signal<string>
        OverrideSection: Signal<string>
        LinkSectionToPrevious: Signal<string>
        SectionHeader: Signal<string>
        SectionFooter: Signal<string>
    }

    type AnnotationState = {
        IsOpen: Signal<bool>
        Kind: Signal<string>
        Target: Signal<string>
        LinkText: Signal<string>
        HasLink: Signal<string>
        Error: Signal<string>
        AnnotationId: Signal<string>
        NoteType: Signal<string>
        Text: Signal<string>
        Author: Signal<string>
        Date: Signal<string>
    }

    let paragraphStyleState () = {
        IsOpen = signal false
        Alignment = signal "left"
        IndentLeft = signal "0"
        IndentRight = signal "0"
        SpacingBefore = signal "0"
        SpacingAfter = signal "10"
        LineSpacing = signal "1.5"
        BackgroundColor = signal "#EAF2F8"
        FontSize = signal "11"
        SpecialIndent = signal "none"
        SpecialIndentBy = signal "0"
    }

    let findReplaceState () = {
        IsOpen = signal false
        Query = signal ""
        Replacement = signal ""
        MatchCase = signal "false"
        Result = signal "No search yet"
    }

    let tableCellStyleState () = {
        IsOpen = signal false
        Error = signal ""
        BackgroundColor = signal "#D9EAF7"
        ShadingPattern = signal "clear"
        ShadingForegroundColor = signal "#6FB31A"
        Width = signal "180"
        BorderColor = signal "#C00000"
        BorderWidth = signal "2"
        VerticalAlignment = signal "center"
        PaddingTop = signal "10"
        PaddingRight = signal "18"
        PaddingBottom = signal "10"
        PaddingLeft = signal "18"
    }

    let tableMarginsState () = {
        IsOpen = signal false
        Error = signal ""
        PaddingTop = signal "0"
        PaddingRight = signal "7.2"
        PaddingBottom = signal "0"
        PaddingLeft = signal "7.2"
        CellSpacing = signal "0"
        Fields = Array.init 14 (fun _ -> signal "")
    }

    let insertTableState () = {
        IsOpen = signal false
        Error = signal ""
        Rows = signal "2"
        Columns = signal "3"
    }

    let imageSizeState () = {
        IsOpen = signal false
        Width = signal "240"
        Height = signal "160"
    }

    let textBoxStyleState () = {
        IsOpen = signal false
        X = signal "96"
        Y = signal "96"
        Width = signal "240"
        Height = signal "96"
        FillColor = signal "#FFFFFF"
        BorderColor = signal "#808080"
        BorderWidth = signal "1"
    }

    let captionState () = {
        IsOpen = signal false
        Label = signal "Figure"
        Text = signal "Caption"
    }

    let crossReferenceState () = {
        IsOpen = signal false
        Label = signal "Figure"
    }

    /// The VBA editor: the module shown and its code, the macro to run, an Immediate line, the output;
    /// the project's modules (name, kind, code) and macros as they were when the editor was (re)opened.
    type VbaEditorState = {
        IsOpen: Signal<bool>
        Module: Signal<string>
        Code: Signal<string>
        Macro: Signal<string>
        Immediate: Signal<string>
        Output: Signal<string>
        mutable Modules: (string * string * string) list
        mutable Macros: string list
    }

    let vbaEditorState () = {
        IsOpen = signal false
        Module = signal "Module1"
        Code = signal ""
        Macro = signal ""
        Immediate = signal "?ActiveDocument.Paragraphs.Count"
        Output = signal ""
        Modules = []
        Macros = []
    }

    let bookmarkState () = {
        IsOpen = signal false
        Name = signal "Bookmark1"
    }

    let contentControlState () = {
        IsOpen = signal false
        Id = signal "1"
        Tag = signal "content-control"
        Alias = signal "Content Control"
    }

    let fieldState () = {
        IsOpen = signal false
        Code = signal " AUTHOR "
        Result = signal "Author"
    }

    let equationState () = {
        IsOpen = signal false
        Text = signal "x+1"
    }

    let revisionState () = {
        IsOpen = signal false
        RevisionType = signal "insert"
        RevisionId = signal "1"
        Author = signal "Reviewer"
        Date = signal "2026-05-01T00:00:00Z"
    }

    let pageSetupState () = {
        IsOpen = signal false
        Preset = signal "custom"
        Orientation = signal "portrait"
        Error = signal ""
        PageWidth = signal "8.5"
        PageHeight = signal "11"
        MarginLeft = signal "1"
        MarginRight = signal "1"
        MarginTop = signal "1"
        MarginBottom = signal "1"
        HeaderDistance = signal "0.5"
        FooterDistance = signal "0.5"
        DifferentFirstPage = signal "false"
        EvenAndOddHeaders = signal "false"
    }

    let headerFooterState () = {
        IsOpen = signal false
        DefaultHeader = signal ""
        DefaultFooter = signal ""
        DefaultHeaderTableLeft = signal ""
        DefaultHeaderTableRight = signal ""
        DefaultFooterTableLeft = signal ""
        DefaultFooterTableRight = signal ""
        DefaultHeaderImageSrc = signal ""
        DefaultHeaderImageAlt = signal "header image"
        DefaultHeaderImageWidth = signal "96"
        DefaultHeaderImageHeight = signal "48"
        DefaultFooterImageSrc = signal ""
        DefaultFooterImageAlt = signal "footer image"
        DefaultFooterImageWidth = signal "96"
        DefaultFooterImageHeight = signal "48"
        FirstHeader = signal ""
        FirstFooter = signal ""
        EvenHeader = signal ""
        EvenFooter = signal ""
        OverrideSection = signal "false"
        LinkSectionToPrevious = signal "false"
        SectionHeader = signal ""
        SectionFooter = signal ""
    }

    let annotationState () = {
        IsOpen = signal false
        Kind = signal "hyperlink"
        Target = signal "https://example.com"
        LinkText = signal ""
        HasLink = signal "false"
        Error = signal ""
        AnnotationId = signal "note-1"
        NoteType = signal "footnote"
        Text = signal "Annotation text"
        Author = signal "Reviewer"
        Date = signal "2026-05-01T00:00:00Z"
    }

    // The host owns the editing transaction and the field signals. This module
    // describes fields and buttons; OfficeUiDialog owns the actual controls.
    type Form = DialogSpec * Dictionary<string, Signal<string>>

    let private textField id label (value: Signal<string>) =
        (dialogText id label "" value.Value |> dialogTestId id), value

    let private mixedTextField id label (value: Signal<string>) =
        (dialogText id label "Mixed" value.Value |> dialogTestId id), value

    let private numberField id label (value: Signal<string>) =
        (dialogNumber id label None None value.Value |> dialogTestId id), value

    let private colorField id label (value: Signal<string>) =
        (dialogColor id label value.Value |> dialogTestId id), value

    let private checkboxField id label (value: Signal<string>) =
        (dialogCheckbox id label (value.Value = "true") |> dialogTestId id), value

    let private choices values = values |> List.map (fun (value, label) -> if value = "" then disabledChoice value label else choice value label)

    let private selectField id label options (value: Signal<string>) =
        let options =
            if options |> List.exists (fun (candidate, _) -> candidate = value.Value) then options
            else (value.Value, value.Value + " (current)") :: options
        (dialogSelect id label (choices options) value.Value |> dialogTestId id), value

    let private statusField id (value: Signal<string>) =
        (dialogStatus id "" value.Value |> dialogTestId id), value

    let private note id content =
        let value = signal content
        (dialogStatus id "" content |> dialogTestId id), value

    let private form id title (fields: (DialogField * Signal<string>) list) (buttons: DialogButton list) : Form =
        let values = Dictionary<string, Signal<string>>()
        for field, value in fields do values.[field.Id] <- value
        { TestId = id; Title = title; Subtitle = None
          Fields = fields |> List.map fst; Tabs = []; Buttons = buttons }, values

    let private cancel id (isOpen: Signal<bool>) =
        dialogButton id "Cancel" false (fun _ -> isOpen.Value <- false; true)

    let private apply id label (isOpen: Signal<bool>) action =
        dialogButton id label true (fun _ -> action (); not isOpen.Value)

    let private mixed options = ("", "Mixed") :: options

    let fontDialog (state: FormatDialogState) onApply =
        let field = state.Fields
        let t i label = mixedTextField ("word-font-" + string i) label field.[i]
        let n i label = mixedTextField ("word-font-" + string i) label field.[i]
        let s i label options = selectField ("word-font-" + string i) label (mixed options) field.[i]
        let yesNo = [ "true", "On"; "false", "Off" ]
        form "word-font-dialog" "Font" [
            statusField "word-font-error" state.Error
            t 0 "Font family"; n 1 "Size (pt)"
            s 2 "Bold" yesNo; s 3 "Italic" yesNo
            s 4 "Underline" [ "none", "None"; "single", "Single"; "double", "Double"; "thick", "Thick"; "dotted", "Dotted"; "dash", "Dash"; "dotDash", "Dot dash"; "dotDotDash", "Dot dot dash"; "wave", "Wave"; "wavyDouble", "Double wave" ]
            s 5 "Strikethrough" yesNo; s 6 "All caps" yesNo; s 7 "Small caps" yesNo
            s 8 "Position" [ "baseline", "Baseline"; "superscript", "Superscript"; "subscript", "Subscript" ]
            t 9 "Text color (#RRGGBB or auto)"; s 11 "Hidden" yesNo; t 10 "Highlight (#RRGGBB or none)"
        ] [ cancel "word-font-cancel" state.IsOpen; apply "word-font-apply" "Apply" state.IsOpen onApply ]

    let listsDialog (state: FormatDialogState) onApply =
        let field = state.Fields
        let s i label options = selectField ("word-lists-" + string i) label (mixed options) field.[i]
        form "word-lists-dialog" "Bullets and numbering" [
            statusField "word-lists-error" state.Error
            s 0 "List kind" [ "none", "Ordinary paragraphs"; "bullet", "Bullets"; "numbered", "Numbering" ]
            s 1 "Bullet marker" [ "•", "•"; "◦", "◦"; "▪", "▪"; "■", "■"; "–", "–" ]
            s 2 "Number format" [ "decimal", "1, 2, 3"; "lowerLetter", "a, b, c"; "upperLetter", "A, B, C"; "lowerRoman", "i, ii, iii"; "upperRoman", "I, II, III" ]
            s 3 "List level" [ for level in 1 .. 9 -> string level, string level ]
            s 4 "Numbering sequence" [ "keep", "Keep existing"; "restart", "Restart at"; "continue", "Continue preceding list" ]
            mixedTextField "word-lists-5" "Restart at (1–2147483647)" field.[5]
            note "word-lists-help" "Restart and continue affect the contiguous list from the first selected item. Continue can cross ordinary paragraphs."
        ] [ cancel "word-lists-cancel" state.IsOpen; apply "word-lists-apply" "Apply" state.IsOpen onApply ]

    let paragraphBordersDialog (state: FormatDialogState) onApply =
        let field = state.Fields
        let t i label = mixedTextField ("word-paragraph-borders-" + string i) label field.[i]
        let n i label = mixedTextField ("word-paragraph-borders-" + string i) label field.[i]
        let s i label options = selectField ("word-paragraph-borders-" + string i) label (mixed options) field.[i]
        let styles = [ "none", "None"; "single", "Single"; "dotted", "Dotted"; "dashed", "Dashed"; "dotDash", "Dot dash" ]
        form "word-paragraph-borders-dialog" "Paragraph borders and shading" [
            statusField "word-paragraph-borders-error" state.Error
            for side, offset in [ "Top", 0; "Right", 4; "Bottom", 8; "Left", 12 ] do
                s offset (side + " border") styles
                t (offset + 1) (side + " color (#RRGGBB or auto)")
                n (offset + 2) (side + " width (px or unset)")
                n (offset + 3) (side + " spacing (px or unset)")
            t 16 "Background (#RRGGBB or none)"
            s 17 "Shading pattern" [ "none", "No shading"; "clear", "Clear fill"; "solid", "Solid"; "pct5", "5%"; "pct10", "10%"; "pct20", "20%"; "horzStripe", "Horizontal stripe"; "vertStripe", "Vertical stripe"; "diagStripe", "Diagonal stripe"; "reverseDiagStripe", "Reverse diagonal stripe"; "diagCross", "Diagonal cross" ]
            t 18 "Pattern color (#RRGGBB or auto)"
        ] [ cancel "word-paragraph-borders-cancel" state.IsOpen; apply "word-paragraph-borders-apply" "Apply" state.IsOpen onApply ]

    let paragraphStyleDialog dialogTestId (state: ParagraphStyleState) onApply =
        form dialogTestId "Paragraph style" [
            selectField "word-paragraph-align" "Alignment" [ "", "Mixed / inherited"; "left", "Left"; "center", "Center"; "right", "Right"; "justify", "Justify" ] state.Alignment
            numberField "word-paragraph-spacing-after" "Spacing after" state.SpacingAfter
            numberField "word-paragraph-spacing-before" "Spacing before" state.SpacingBefore
            numberField "word-paragraph-line-spacing" "Line spacing" state.LineSpacing
            (dialogText "word-paragraph-background" "Background" "Mixed / inherited" state.BackgroundColor.Value |> Rbt.Office.Ui.dialogTestId "word-paragraph-background"), state.BackgroundColor
            numberField "word-paragraph-indent-left" "Left indent" state.IndentLeft
            numberField "word-paragraph-indent-right" "Right indent" state.IndentRight
            numberField "word-paragraph-font-size" "Font size" state.FontSize
            selectField "word-paragraph-special-indent" "Special indent" [ "", "Mixed / inherited"; "none", "None"; "firstLine", "First line"; "hanging", "Hanging" ] state.SpecialIndent
            numberField "word-paragraph-special-indent-by" "By" state.SpecialIndentBy
        ] [ cancel "word-paragraph-cancel" state.IsOpen; apply "word-paragraph-apply" "Apply" state.IsOpen onApply ]

    let findReplaceDialog dialogTestId (state: FindReplaceState) onFindNext onReplaceAll =
        form dialogTestId "Find and replace" [
            textField "word-find-query" "Find" state.Query
            textField "word-find-replacement" "Replace with" state.Replacement
            checkboxField "word-find-match-case" "Match case" state.MatchCase
            statusField "word-find-result" state.Result
        ] [ dialogButton "word-find-close" "Close" false (fun _ -> state.IsOpen.Value <- false; true)
            dialogButton "word-find-next" "Find next" true (fun _ -> onFindNext(); false)
            dialogButton "word-replace-all" "Replace all" false (fun _ -> onReplaceAll(); false) ]

    let tableCellStyleDialog dialogTestId (state: TableCellStyleState) onApply =
        form dialogTestId "Table cell style" [
            statusField "word-table-cell-style-error" state.Error
            colorField "word-table-cell-background" "Background color" state.BackgroundColor
            selectField "word-table-cell-shading-pattern" "Shading pattern" [ "clear", "Clear fill"; "none", "No fill"; "solid", "Solid pattern"; "pct5", "5% pattern"; "pct10", "10% pattern"; "pct20", "20% pattern"; "horzStripe", "Horizontal stripe"; "vertStripe", "Vertical stripe"; "diagStripe", "Diagonal stripe"; "reverseDiagStripe", "Reverse diagonal stripe"; "diagCross", "Diagonal cross" ] state.ShadingPattern
            colorField "word-table-cell-shading-color" "Pattern color" state.ShadingForegroundColor
            numberField "word-table-cell-width" "Width override (twips)" state.Width
            textField "word-table-cell-border-color" "Border color" state.BorderColor
            numberField "word-table-cell-border-width" "Border width" state.BorderWidth
            selectField "word-table-cell-valign" "Vertical alignment" [ "top", "Top"; "center", "Center"; "bottom", "Bottom" ] state.VerticalAlignment
            numberField "word-table-cell-padding-top" "Cell padding top" state.PaddingTop
            numberField "word-table-cell-padding-right" "Cell padding right" state.PaddingRight
            numberField "word-table-cell-padding-bottom" "Cell padding bottom" state.PaddingBottom
            numberField "word-table-cell-padding-left" "Cell padding left" state.PaddingLeft
        ] [ cancel "word-table-cell-style-cancel" state.IsOpen; apply "word-table-cell-style-apply" "Apply" state.IsOpen onApply ]

    let private tablePatternChoices =
        [ "none", "No shading"; "clear", "Clear fill"; "solid", "Solid"; "pct5", "5%"; "pct10", "10%"; "pct20", "20%"; "horzStripe", "Horizontal stripe"; "vertStripe", "Vertical stripe"; "diagStripe", "Diagonal stripe"; "reverseDiagStripe", "Reverse diagonal stripe"; "diagCross", "Diagonal cross" ]

    let tableMarginsDialog dialogTestId (state: TableMarginsState) onApply =
        let fields = state.Fields
        let s id index label options = selectField id label options fields.[index]
        let n id index label = numberField id label fields.[index]
        let t id index label = textField id label fields.[index]
        let update field change = let spec, value = field in dialogUpdating change spec, value
        let set key newValue values = values |> Map.add key newValue
        form dialogTestId "Table properties" [
            statusField "word-table-margins-error" state.Error
            note "word-table-description" "Table (width in twips or percent; indent and margins in pixels)"
            s "word-table-width-kind" 0 "Preferred width unit" [ "", "Inherited"; "auto", "Auto"; "dxa", "Twips"; "pct", "Percent"; "pt", "Pixels" ]
            n "word-table-width" 1 "Preferred width"
            s "word-table-alignment" 2 "Alignment" [ "", "Inherited"; "left", "Left"; "center", "Center"; "right", "Right"; "start", "Start"; "end", "End" ]
            n "word-table-indent" 3 "Indent (left/start alignment)"
            note "word-table-row-description" "Current row (only leading header rows repeat)"
            n "word-table-row-height" 5 "Minimum row height (twips; changing replaces any exact height)"
            s "word-table-row-split" 6 "Allow row to split across pages" [ "", "Inherited"; "true", "Yes"; "false", "No" ]
            s "word-table-row-header" 7 "Repeat as header" [ "true", "Yes"; "false", "No" ]
            t "word-table-background" 8 "Table background color (blank for none)"
            s "word-table-shading-pattern" 12 "Table shading pattern" tablePatternChoices
            t "word-table-shading-color" 13 "Table pattern color (#RRGGBB)"
            |> fun field -> update field (fun values value ->
                values |> set "word-table-shading-color" value
                       |> (fun values -> if values.["word-table-shading-pattern"] = "none" || values.["word-table-shading-pattern"] = "clear" then set "word-table-shading-pattern" "solid" values else values))
            note "word-table-border-description" "Table borders (cell and row overrides remain)"
            s "word-table-borders" 9 "Table borders" [ "keep", "Keep existing"; "grid", "Replace with grid"; "none", "Remove" ]
            t "word-table-border-color" 10 "Grid border color (#RRGGBB)"
            |> fun field -> update field (fun values value -> values |> set "word-table-border-color" value |> set "word-table-borders" "grid")
            n "word-table-border-width" 11 "Grid border width (pixels)"
            |> fun field -> update field (fun values value -> values |> set "word-table-border-width" value |> set "word-table-borders" "grid")
            note "word-table-cell-description" "Default cell margins and spacing"
            numberField "word-table-margin-top" "Default cell margin top" state.PaddingTop
            numberField "word-table-margin-right" "Default cell margin right" state.PaddingRight
            numberField "word-table-margin-bottom" "Default cell margin bottom" state.PaddingBottom
            numberField "word-table-margin-left" "Default cell margin left" state.PaddingLeft
            numberField "word-table-cell-spacing" "Cell spacing" state.CellSpacing
        ] [ cancel "word-table-margins-cancel" state.IsOpen; apply "word-table-margins-apply" "Apply" state.IsOpen onApply ]

    let insertTableDialog dialogTestId (state: InsertTableState) onApply =
        form dialogTestId "Insert table" [
            statusField "word-insert-table-error" state.Error
            numberField "word-insert-table-rows" "Rows" state.Rows
            numberField "word-insert-table-columns" "Columns" state.Columns
        ] [ cancel "word-insert-table-cancel" state.IsOpen; apply "word-insert-table-apply" "Insert" state.IsOpen onApply ]

    let imageSizeDialog dialogTestId (state: ImageSizeState) onApply =
        form dialogTestId "Image size" [
            numberField "word-image-width" "Width" state.Width
            numberField "word-image-height" "Height" state.Height
        ] [ cancel "word-image-size-cancel" state.IsOpen; apply "word-image-size-apply" "Apply" state.IsOpen onApply ]

    let textBoxStyleDialog dialogTestId (state: TextBoxStyleState) onApply =
        form dialogTestId "Text box style" [
            numberField "word-textbox-x" "X" state.X; numberField "word-textbox-y" "Y" state.Y
            numberField "word-textbox-width" "Width" state.Width; numberField "word-textbox-height" "Height" state.Height
            colorField "word-textbox-fill" "Fill" state.FillColor
            colorField "word-textbox-border-color" "Border" state.BorderColor
            numberField "word-textbox-border-width" "Border width" state.BorderWidth
        ] [ cancel "word-textbox-style-cancel" state.IsOpen; apply "word-textbox-style-apply" "Apply" state.IsOpen onApply ]

    let private captionChoices = [ "Figure", "Figure"; "Table", "Table" ]

    let captionDialog dialogTestId (state: CaptionState) onApply =
        form dialogTestId "Caption" [
            selectField "word-caption-label" "Label" captionChoices state.Label
            textField "word-caption-text" "Text" state.Text
        ] [ cancel "word-caption-cancel" state.IsOpen; apply "word-caption-apply" "Insert" state.IsOpen onApply ]

    let crossReferenceDialog dialogTestId (state: CrossReferenceState) onApply =
        form dialogTestId "Cross-reference" [
            selectField "word-cross-reference-label" "Reference type" captionChoices state.Label
            note "word-cross-reference-help" "Inserts a reference to the latest caption of this type."
        ] [ cancel "word-cross-reference-cancel" state.IsOpen; apply "word-cross-reference-apply" "Insert" state.IsOpen onApply ]

    /// The VBA editor (the spreadsheet's Ui.macroEditor in the Word shell's terms): a module's code to read
    /// and change (Apply compiles it), the macros to run on the document (one undo step), an Immediate line,
    /// and the output. The dialog stays open; the host shows it again when the project changed.
    let vbaEditorDialog testId (state: VbaEditorState) (onApply: unit -> unit) (onRun: unit -> unit) (onImmediate: unit -> unit) =
        let keepOpen id label primary action = dialogButton id label primary (fun _ -> action (); false)
        let moduleField, moduleValue =
            selectField "word-vba-module" "Module" [ for name, kind, _ in state.Modules -> name, sprintf "%s (%s)" name kind ] state.Module
        let codeField = dialogTextArea "word-vba-code" "Code" "Sub Main()\n    ActiveDocument.Content.InsertAfter \"Hello\"\nEnd Sub" 16 state.Code.Value |> dialogTestId "word-vba-code"
        let spec, values =
            form testId "VBA" [
                (moduleField |> dialogUpdating (fun values chosen ->
                    let code = state.Modules |> List.tryPick (fun (n, _, c) -> if n = chosen then Some c else None) |> Option.defaultValue ""
                    values |> Map.add "word-vba-module" chosen |> Map.add "word-vba-code" code)), moduleValue
                codeField, state.Code
                selectField "word-vba-macro" "Macro" [ for name in state.Macros -> name, name ] state.Macro
                textField "word-vba-immediate" "Immediate (?expression, or a statement)" state.Immediate
                statusField "word-vba-output" state.Output
            ] [
                dialogButton "word-vba-close" "Close" false (fun _ -> state.IsOpen.Value <- false; true)
                keepOpen "word-vba-apply" "Apply" false onApply
                keepOpen "word-vba-immediate-run" "Run line" false onImmediate
                keepOpen "word-vba-run" "Run macro" true onRun
            ]
        { spec with Subtitle = Some "Apply compiles a module's changes; Run runs a macro on the document as one undo step." }, values

    let bookmarkDialog dialogTestId (state: BookmarkState) onApply =
        form dialogTestId "Bookmark" [
            textField "word-bookmark-name" "Name" state.Name
        ] [ cancel "word-bookmark-cancel" state.IsOpen; apply "word-bookmark-apply" "Apply" state.IsOpen onApply ]

    let contentControlDialog dialogTestId (state: ContentControlState) onApply =
        form dialogTestId "Content control" [
            textField "word-content-control-id" "ID" state.Id
            textField "word-content-control-tag" "Tag" state.Tag
            textField "word-content-control-alias" "Alias" state.Alias
        ] [ cancel "word-content-control-cancel" state.IsOpen; apply "word-content-control-apply" "Apply" state.IsOpen onApply ]

    let fieldDialog dialogTestId (state: FieldState) onApply =
        form dialogTestId "Field" [
            textField "word-field-code" "Code" state.Code; textField "word-field-result" "Result" state.Result
        ] [ cancel "word-field-cancel" state.IsOpen; apply "word-field-apply" "Apply" state.IsOpen onApply ]

    let equationDialog dialogTestId (state: EquationState) onApply =
        form dialogTestId "Equation" [
            textField "word-equation-text" "Equation text" state.Text
            note "word-equation-help" "Stores this as opaque Word OMML for DOCX export."
        ] [ cancel "word-equation-cancel" state.IsOpen; apply "word-equation-apply" "Insert" state.IsOpen onApply ]

    let revisionDialog dialogTestId (state: RevisionState) onApply =
        form dialogTestId "Revision" [
            selectField "word-revision-type" "Type" [ "insert", "Inserted"; "delete", "Deleted" ] state.RevisionType
            textField "word-revision-id" "ID" state.RevisionId
            textField "word-revision-author" "Author" state.Author
            textField "word-revision-date" "Date" state.Date
        ] [ cancel "word-revision-cancel" state.IsOpen; apply "word-revision-apply" "Apply" state.IsOpen onApply ]

    let private pageSizes =
        [ "letter", (8.5, 11.0); "legal", (8.5, 14.0)
          "a4", (210.0 / 25.4, 297.0 / 25.4)
          "a5", (148.0 / 25.4, 210.0 / 25.4) ]

    let pageSetupDialog dialogTestId (state: PageSetupState) onApply =
        let update field changed = let spec, value = field in dialogUpdating changed spec, value
        let width, height, preset, orientation = "word-page-width", "word-page-height", "word-page-preset", "word-page-orientation"
        let dimensionsChanged id values value =
            let values = values |> Map.add id value |> Map.add preset "custom"
            match System.Double.TryParse values.[width], System.Double.TryParse values.[height] with
            | (true, w), (true, h) when w > 0.0 && h > 0.0 ->
                values |> Map.add orientation (if w > h then "landscape" else "portrait")
            | _ -> values
        form dialogTestId "Page setup" [
            statusField "word-page-setup-error" state.Error
            note "word-page-setup-description" "Applies to the section containing the selection's focus."
            selectField preset "Paper size" [ "custom", "Custom"; "letter", "Letter (8.5 × 11 in)"; "legal", "Legal (8.5 × 14 in)"; "a4", "A4 (210 × 297 mm)"; "a5", "A5 (148 × 210 mm)" ] state.Preset
            |> fun field -> update field (fun values value ->
                let values = values |> Map.add preset value
                match pageSizes |> List.tryFind (fun (name, _) -> name = value) with
                | None -> values
                | Some (_, (w, h)) ->
                    let w, h = if values.[orientation] = "landscape" then h, w else w, h
                    values |> Map.add width (string w) |> Map.add height (string h))
            selectField orientation "Orientation" [ "portrait", "Portrait"; "landscape", "Landscape" ] state.Orientation
            |> fun field -> update field (fun values value ->
                if values.[orientation] = value then values
                else values |> Map.add orientation value |> Map.add width values.[height] |> Map.add height values.[width])
            numberField width "Page width (in)" state.PageWidth |> fun field -> update field (dimensionsChanged width)
            numberField height "Page height (in)" state.PageHeight |> fun field -> update field (dimensionsChanged height)
            numberField "word-margin-top" "Top margin (in)" state.MarginTop
            numberField "word-margin-bottom" "Bottom margin (in)" state.MarginBottom
            numberField "word-margin-left" "Left margin (in)" state.MarginLeft
            numberField "word-margin-right" "Right margin (in)" state.MarginRight
            numberField "word-header-distance" "Header from top (in)" state.HeaderDistance
            numberField "word-footer-distance" "Footer from bottom (in)" state.FooterDistance
            checkboxField "word-different-first-page" "Different first page" state.DifferentFirstPage
        ] [ cancel "word-page-setup-cancel" state.IsOpen; apply "word-page-setup-apply" "Apply" state.IsOpen onApply ]

    let headerFooterDialog dialogTestId (state: HeaderFooterState) onApply =
        form dialogTestId "Headers and footers" [
            textField "word-default-header" "Default header" state.DefaultHeader
            textField "word-default-footer" "Default footer" state.DefaultFooter
            textField "word-default-header-table-left" "Default header table left" state.DefaultHeaderTableLeft
            textField "word-default-header-table-right" "Default header table right" state.DefaultHeaderTableRight
            textField "word-default-footer-table-left" "Default footer table left" state.DefaultFooterTableLeft
            textField "word-default-footer-table-right" "Default footer table right" state.DefaultFooterTableRight
            textField "word-default-header-image-src" "Default header image URL" state.DefaultHeaderImageSrc
            textField "word-default-header-image-alt" "Default header image alt" state.DefaultHeaderImageAlt
            numberField "word-default-header-image-width" "Default header image width" state.DefaultHeaderImageWidth
            numberField "word-default-header-image-height" "Default header image height" state.DefaultHeaderImageHeight
            textField "word-default-footer-image-src" "Default footer image URL" state.DefaultFooterImageSrc
            textField "word-default-footer-image-alt" "Default footer image alt" state.DefaultFooterImageAlt
            numberField "word-default-footer-image-width" "Default footer image width" state.DefaultFooterImageWidth
            numberField "word-default-footer-image-height" "Default footer image height" state.DefaultFooterImageHeight
            textField "word-first-header" "First-page header" state.FirstHeader
            textField "word-first-footer" "First-page footer" state.FirstFooter
            textField "word-even-header" "Even-page header" state.EvenHeader
            textField "word-even-footer" "Even-page footer" state.EvenFooter
            checkboxField "word-override-section-header-footer" "Override current section" state.OverrideSection
            checkboxField "word-link-section-header-footer" "Link current section to previous" state.LinkSectionToPrevious
            textField "word-section-header" "Current section header" state.SectionHeader
            textField "word-section-footer" "Current section footer" state.SectionFooter
        ] [ cancel "word-header-footer-cancel" state.IsOpen; apply "word-header-footer-apply" "Apply" state.IsOpen onApply ]

    let annotationDialog dialogTestId (state: AnnotationState) onApply =
        let title =
            match state.Kind.Value with
            | "comment" -> "Comment"
            | "note" -> "Footnote / endnote"
            | _ -> "Hyperlink"
        form dialogTestId title [
            note "word-annotation-kind" title
            statusField "word-annotation-error" state.Error
            match state.Kind.Value with
            | "hyperlink" ->
                textField "word-annotation-target" "Target" state.Target
                if state.HasLink.Value <> "true" then
                    textField "word-link-text" "Text (used when no text is selected)" state.LinkText
            | "note" ->
                textField "word-annotation-id" "ID" state.AnnotationId
                selectField "word-annotation-note-type" "Note type" [ "footnote", "Footnote"; "endnote", "Endnote" ] state.NoteType
                textField "word-annotation-text" "Text" state.Text
            | _ ->
                textField "word-annotation-id" "ID" state.AnnotationId
                textField "word-annotation-text" "Text" state.Text
                textField "word-annotation-author" "Author" state.Author
                textField "word-annotation-date" "Date" state.Date
        ] [
            if state.Kind.Value = "hyperlink" && state.HasLink.Value = "true" then
                dialogButton "word-link-remove" "Remove link" false (fun _ ->
                    state.Target.Value <- ""
                    onApply ()
                    not state.IsOpen.Value)
            cancel "word-annotation-cancel" state.IsOpen
            apply "word-annotation-apply" "Apply" state.IsOpen onApply
        ]

#endif
