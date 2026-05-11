namespace RedBat.Word

#if FABLE_COMPILER

module Dialogs =
    open Browser.Types
    open RedBat.FableUI.Reactive
    open RedBat.FableUI.Dom
    open RedBat.FableUI.UI

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
        MatchCase: Signal<bool>
        Result: Signal<string>
    }

    type TableCellStyleState = {
        IsOpen: Signal<bool>
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

    type InsertTableState = {
        IsOpen: Signal<bool>
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
        PageWidth: Signal<string>
        PageHeight: Signal<string>
        MarginLeft: Signal<string>
        MarginRight: Signal<string>
        MarginTop: Signal<string>
        MarginBottom: Signal<string>
        HeaderDistance: Signal<string>
        FooterDistance: Signal<string>
        DifferentFirstPage: Signal<bool>
        EvenAndOddHeaders: Signal<bool>
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
        OverrideSection: Signal<bool>
        LinkSectionToPrevious: Signal<bool>
        SectionHeader: Signal<string>
        SectionFooter: Signal<string>
    }

    type AnnotationState = {
        IsOpen: Signal<bool>
        Kind: Signal<string>
        Target: Signal<string>
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
        MatchCase = signal false
        Result = signal "No search yet"
    }

    let tableCellStyleState () = {
        IsOpen = signal false
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

    let insertTableState () = {
        IsOpen = signal false
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
        PageWidth = signal "8.5"
        PageHeight = signal "11"
        MarginLeft = signal "1"
        MarginRight = signal "1"
        MarginTop = signal "1"
        MarginBottom = signal "1"
        HeaderDistance = signal "0.5"
        FooterDistance = signal "0.5"
        DifferentFirstPage = signal false
        EvenAndOddHeaders = signal false
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
        OverrideSection = signal false
        LinkSectionToPrevious = signal false
        SectionHeader = signal ""
        SectionFooter = signal ""
    }

    let annotationState () = {
        IsOpen = signal false
        Kind = signal "hyperlink"
        Target = signal "https://example.com"
        AnnotationId = signal "note-1"
        NoteType = signal "footnote"
        Text = signal "Annotation text"
        Author = signal "Reviewer"
        Date = signal "2026-05-01T00:00:00Z"
    }

    let private selectValue (event: Event) =
        (event.target :?> HTMLSelectElement).value

    let private checkedValue (event: Event) =
        (event.target :?> HTMLInputElement).``checked``

    let paragraphStyleDialog overlayTestId dialogTestId (state: ParagraphStyleState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Paragraph style" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Alignment" ]
                    select [ attr "data-testid" "word-paragraph-align"; propS "value" state.Alignment; onChange (fun event -> state.Alignment.Value <- selectValue event) ] [
                        option [ attr "value" "left" ] [ text "Left" ]
                        option [ attr "value" "center" ] [ text "Center" ]
                        option [ attr "value" "right" ] [ text "Right" ]
                        option [ attr "value" "justify" ] [ text "Justify" ]
                    ]
                ]
                label [] [
                    span [] [ text "Spacing after" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-spacing-after"; valueS state.SpacingAfter; onInput (fun value -> state.SpacingAfter.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Spacing before" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-spacing-before"; valueS state.SpacingBefore; onInput (fun value -> state.SpacingBefore.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Line spacing" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-line-spacing"; attr "step" "0.1"; valueS state.LineSpacing; onInput (fun value -> state.LineSpacing.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Background" ]
                    Input [ typ "color"; attr "data-testid" "word-paragraph-background"; valueS state.BackgroundColor; onInput (fun value -> state.BackgroundColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Left indent" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-indent-left"; valueS state.IndentLeft; onInput (fun value -> state.IndentLeft.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Right indent" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-indent-right"; valueS state.IndentRight; onInput (fun value -> state.IndentRight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Font size" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-font-size"; valueS state.FontSize; onInput (fun value -> state.FontSize.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Special indent" ]
                    select [ attr "data-testid" "word-paragraph-special-indent"; propS "value" state.SpecialIndent; onChange (fun event -> state.SpecialIndent.Value <- selectValue event) ] [
                        option [ attr "value" "none" ] [ text "None" ]
                        option [ attr "value" "firstLine" ] [ text "First line" ]
                        option [ attr "value" "hanging" ] [ text "Hanging" ]
                    ]
                ]
                label [] [
                    span [] [ text "By" ]
                    Input [ typ "number"; attr "data-testid" "word-paragraph-special-indent-by"; valueS state.SpecialIndentBy; onInput (fun value -> state.SpecialIndentBy.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-paragraph-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-paragraph-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let findReplaceDialog overlayTestId dialogTestId (state: FindReplaceState) onFindNext onReplaceAll =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Find and replace" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Find" ]
                    Input [ attr "data-testid" "word-find-query"; valueS state.Query; onInput (fun value -> state.Query.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Replace with" ]
                    Input [ attr "data-testid" "word-find-replacement"; valueS state.Replacement; onInput (fun value -> state.Replacement.Value <- value) ]
                ]
                label [ cls "dialog-inline-field" ] [
                    input [ typ "checkbox"; attr "data-testid" "word-find-match-case"; propS "checked" state.MatchCase; onChange (fun event -> state.MatchCase.Value <- checkedValue event) ]
                    span [] [ text "Match case" ]
                ]
                p [ attr "data-testid" "word-find-result" ] [ textR (fun () -> state.Result.Value) ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-find-close"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Close" ]
                Button [ attr "data-testid" "word-find-next"; onClick (fun _ -> onFindNext()) ] [ text "Find next" ]
                Button [ attr "data-testid" "word-replace-all"; onClick (fun _ -> onReplaceAll()) ] [ text "Replace all" ]
            ]
        ]

    let tableCellStyleDialog overlayTestId dialogTestId (state: TableCellStyleState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Table cell style" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Background color" ]
                    Input [ typ "color"; attr "data-testid" "word-table-cell-background"; valueS state.BackgroundColor; onInput (fun value -> state.BackgroundColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Shading pattern" ]
                    select [ attr "data-testid" "word-table-cell-shading-pattern"; propS "value" state.ShadingPattern; onChange (fun event -> state.ShadingPattern.Value <- selectValue event) ] [
                        option [ attr "value" "clear" ] [ text "Clear fill" ]
                        option [ attr "value" "none" ] [ text "No fill" ]
                        option [ attr "value" "solid" ] [ text "Solid pattern" ]
                        option [ attr "value" "pct5" ] [ text "5% pattern" ]
                        option [ attr "value" "pct10" ] [ text "10% pattern" ]
                        option [ attr "value" "pct20" ] [ text "20% pattern" ]
                        option [ attr "value" "horzStripe" ] [ text "Horizontal stripe" ]
                        option [ attr "value" "vertStripe" ] [ text "Vertical stripe" ]
                        option [ attr "value" "diagStripe" ] [ text "Diagonal stripe" ]
                        option [ attr "value" "reverseDiagStripe" ] [ text "Reverse diagonal stripe" ]
                        option [ attr "value" "diagCross" ] [ text "Diagonal cross" ]
                    ]
                ]
                label [] [
                    span [] [ text "Pattern color" ]
                    Input [ typ "color"; attr "data-testid" "word-table-cell-shading-color"; valueS state.ShadingForegroundColor; onInput (fun value -> state.ShadingForegroundColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Width" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-width"; valueS state.Width; onInput (fun value -> state.Width.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Border color" ]
                    Input [ attr "data-testid" "word-table-cell-border-color"; valueS state.BorderColor; onInput (fun value -> state.BorderColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Border width" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-border-width"; valueS state.BorderWidth; onInput (fun value -> state.BorderWidth.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Vertical alignment" ]
                    select [ attr "data-testid" "word-table-cell-valign"; propS "value" state.VerticalAlignment; onChange (fun event -> state.VerticalAlignment.Value <- selectValue event) ] [
                        option [ attr "value" "top" ] [ text "Top" ]
                        option [ attr "value" "center" ] [ text "Center" ]
                        option [ attr "value" "bottom" ] [ text "Bottom" ]
                    ]
                ]
                label [] [
                    span [] [ text "Cell padding top" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-padding-top"; valueS state.PaddingTop; onInput (fun value -> state.PaddingTop.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Cell padding right" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-padding-right"; valueS state.PaddingRight; onInput (fun value -> state.PaddingRight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Cell padding bottom" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-padding-bottom"; valueS state.PaddingBottom; onInput (fun value -> state.PaddingBottom.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Cell padding left" ]
                    Input [ typ "number"; attr "data-testid" "word-table-cell-padding-left"; valueS state.PaddingLeft; onInput (fun value -> state.PaddingLeft.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-table-cell-style-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-table-cell-style-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let insertTableDialog overlayTestId dialogTestId (state: InsertTableState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Insert table" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Rows" ]
                    Input [ typ "number"; attr "data-testid" "word-insert-table-rows"; valueS state.Rows; onInput (fun value -> state.Rows.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Columns" ]
                    Input [ typ "number"; attr "data-testid" "word-insert-table-columns"; valueS state.Columns; onInput (fun value -> state.Columns.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-insert-table-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-insert-table-apply"; onClick (fun _ -> onApply()) ] [ text "Insert" ]
            ]
        ]

    let imageSizeDialog overlayTestId dialogTestId (state: ImageSizeState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Image size" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Width" ]
                    Input [ typ "number"; attr "data-testid" "word-image-width"; valueS state.Width; onInput (fun value -> state.Width.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Height" ]
                    Input [ typ "number"; attr "data-testid" "word-image-height"; valueS state.Height; onInput (fun value -> state.Height.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-image-size-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-image-size-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let textBoxStyleDialog overlayTestId dialogTestId (state: TextBoxStyleState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Text box style" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "X" ]
                    Input [ typ "number"; attr "data-testid" "word-textbox-x"; valueS state.X; onInput (fun value -> state.X.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Y" ]
                    Input [ typ "number"; attr "data-testid" "word-textbox-y"; valueS state.Y; onInput (fun value -> state.Y.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Width" ]
                    Input [ typ "number"; attr "data-testid" "word-textbox-width"; valueS state.Width; onInput (fun value -> state.Width.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Height" ]
                    Input [ typ "number"; attr "data-testid" "word-textbox-height"; valueS state.Height; onInput (fun value -> state.Height.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Fill" ]
                    Input [ typ "color"; attr "data-testid" "word-textbox-fill"; valueS state.FillColor; onInput (fun value -> state.FillColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Border" ]
                    Input [ typ "color"; attr "data-testid" "word-textbox-border-color"; valueS state.BorderColor; onInput (fun value -> state.BorderColor.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Border width" ]
                    Input [ typ "number"; attr "data-testid" "word-textbox-border-width"; valueS state.BorderWidth; onInput (fun value -> state.BorderWidth.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-textbox-style-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-textbox-style-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let captionDialog overlayTestId dialogTestId (state: CaptionState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Caption" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Label" ]
                    select [ attr "data-testid" "word-caption-label"; propS "value" state.Label; onChange (fun event -> state.Label.Value <- selectValue event) ] [
                        option [ attr "value" "Figure" ] [ text "Figure" ]
                        option [ attr "value" "Table" ] [ text "Table" ]
                    ]
                ]
                label [] [
                    span [] [ text "Text" ]
                    Input [ attr "data-testid" "word-caption-text"; valueS state.Text; onInput (fun value -> state.Text.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-caption-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-caption-apply"; onClick (fun _ -> onApply()) ] [ text "Insert" ]
            ]
        ]

    let crossReferenceDialog overlayTestId dialogTestId (state: CrossReferenceState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Cross-reference" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Reference type" ]
                    select [ attr "data-testid" "word-cross-reference-label"; propS "value" state.Label; onChange (fun event -> state.Label.Value <- selectValue event) ] [
                        option [ attr "value" "Figure" ] [ text "Figure" ]
                        option [ attr "value" "Table" ] [ text "Table" ]
                    ]
                ]
                p [ attr "data-testid" "word-cross-reference-help" ] [ text "Inserts a reference to the latest caption of this type." ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-cross-reference-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-cross-reference-apply"; onClick (fun _ -> onApply()) ] [ text "Insert" ]
            ]
        ]

    let bookmarkDialog overlayTestId dialogTestId (state: BookmarkState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Bookmark" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Name" ]
                    Input [ attr "data-testid" "word-bookmark-name"; valueS state.Name; onInput (fun value -> state.Name.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-bookmark-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-bookmark-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let contentControlDialog overlayTestId dialogTestId (state: ContentControlState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Content control" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "ID" ]
                    Input [ attr "data-testid" "word-content-control-id"; valueS state.Id; onInput (fun value -> state.Id.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Tag" ]
                    Input [ attr "data-testid" "word-content-control-tag"; valueS state.Tag; onInput (fun value -> state.Tag.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Alias" ]
                    Input [ attr "data-testid" "word-content-control-alias"; valueS state.Alias; onInput (fun value -> state.Alias.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-content-control-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-content-control-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let fieldDialog overlayTestId dialogTestId (state: FieldState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Field" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Code" ]
                    Input [ attr "data-testid" "word-field-code"; valueS state.Code; onInput (fun value -> state.Code.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Result" ]
                    Input [ attr "data-testid" "word-field-result"; valueS state.Result; onInput (fun value -> state.Result.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-field-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-field-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let equationDialog overlayTestId dialogTestId (state: EquationState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Equation" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Equation text" ]
                    Input [ attr "data-testid" "word-equation-text"; valueS state.Text; onInput (fun value -> state.Text.Value <- value) ]
                ]
                p [ attr "data-testid" "word-equation-help" ] [ text "Stores this as opaque Word OMML for DOCX export." ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-equation-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-equation-apply"; onClick (fun _ -> onApply()) ] [ text "Insert" ]
            ]
        ]

    let revisionDialog overlayTestId dialogTestId (state: RevisionState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Revision" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Type" ]
                    select [ attr "data-testid" "word-revision-type"; propS "value" state.RevisionType; onChange (fun event -> state.RevisionType.Value <- selectValue event) ] [
                        option [ attr "value" "insert" ] [ text "Inserted" ]
                        option [ attr "value" "delete" ] [ text "Deleted" ]
                    ]
                ]
                label [] [
                    span [] [ text "ID" ]
                    Input [ attr "data-testid" "word-revision-id"; valueS state.RevisionId; onInput (fun value -> state.RevisionId.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Author" ]
                    Input [ attr "data-testid" "word-revision-author"; valueS state.Author; onInput (fun value -> state.Author.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Date" ]
                    Input [ attr "data-testid" "word-revision-date"; valueS state.Date; onInput (fun value -> state.Date.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-revision-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-revision-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let pageSetupDialog overlayTestId dialogTestId (state: PageSetupState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Page setup" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Page width (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-page-width"; valueS state.PageWidth; onInput (fun value -> state.PageWidth.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Page height (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-page-height"; valueS state.PageHeight; onInput (fun value -> state.PageHeight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Top margin (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-margin-top"; valueS state.MarginTop; onInput (fun value -> state.MarginTop.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Bottom margin (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-margin-bottom"; valueS state.MarginBottom; onInput (fun value -> state.MarginBottom.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Left margin (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-margin-left"; valueS state.MarginLeft; onInput (fun value -> state.MarginLeft.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Right margin (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-margin-right"; valueS state.MarginRight; onInput (fun value -> state.MarginRight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Header from top (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-header-distance"; valueS state.HeaderDistance; onInput (fun value -> state.HeaderDistance.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Footer from bottom (in)" ]
                    Input [ typ "number"; attr "step" "0.1"; attr "data-testid" "word-footer-distance"; valueS state.FooterDistance; onInput (fun value -> state.FooterDistance.Value <- value) ]
                ]
                label [ cls "dialog-inline-field" ] [
                    input [ typ "checkbox"; attr "data-testid" "word-different-first-page"; propS "checked" state.DifferentFirstPage; onChange (fun event -> state.DifferentFirstPage.Value <- checkedValue event) ]
                    span [] [ text "Different first page" ]
                ]
                label [ cls "dialog-inline-field" ] [
                    input [ typ "checkbox"; attr "data-testid" "word-even-and-odd-headers"; propS "checked" state.EvenAndOddHeaders; onChange (fun event -> state.EvenAndOddHeaders.Value <- checkedValue event) ]
                    span [] [ text "Different odd and even pages" ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-page-setup-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-page-setup-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let headerFooterDialog overlayTestId dialogTestId (state: HeaderFooterState) onApply =
        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ text "Headers and footers" ] ]
            DialogBody [] [
                label [] [
                    span [] [ text "Default header" ]
                    Input [ attr "data-testid" "word-default-header"; valueS state.DefaultHeader; onInput (fun value -> state.DefaultHeader.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer" ]
                    Input [ attr "data-testid" "word-default-footer"; valueS state.DefaultFooter; onInput (fun value -> state.DefaultFooter.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header table left" ]
                    Input [ attr "data-testid" "word-default-header-table-left"; valueS state.DefaultHeaderTableLeft; onInput (fun value -> state.DefaultHeaderTableLeft.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header table right" ]
                    Input [ attr "data-testid" "word-default-header-table-right"; valueS state.DefaultHeaderTableRight; onInput (fun value -> state.DefaultHeaderTableRight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer table left" ]
                    Input [ attr "data-testid" "word-default-footer-table-left"; valueS state.DefaultFooterTableLeft; onInput (fun value -> state.DefaultFooterTableLeft.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer table right" ]
                    Input [ attr "data-testid" "word-default-footer-table-right"; valueS state.DefaultFooterTableRight; onInput (fun value -> state.DefaultFooterTableRight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header image URL" ]
                    Input [ attr "data-testid" "word-default-header-image-src"; valueS state.DefaultHeaderImageSrc; onInput (fun value -> state.DefaultHeaderImageSrc.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header image alt" ]
                    Input [ attr "data-testid" "word-default-header-image-alt"; valueS state.DefaultHeaderImageAlt; onInput (fun value -> state.DefaultHeaderImageAlt.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header image width" ]
                    Input [ typ "number"; attr "data-testid" "word-default-header-image-width"; valueS state.DefaultHeaderImageWidth; onInput (fun value -> state.DefaultHeaderImageWidth.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default header image height" ]
                    Input [ typ "number"; attr "data-testid" "word-default-header-image-height"; valueS state.DefaultHeaderImageHeight; onInput (fun value -> state.DefaultHeaderImageHeight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer image URL" ]
                    Input [ attr "data-testid" "word-default-footer-image-src"; valueS state.DefaultFooterImageSrc; onInput (fun value -> state.DefaultFooterImageSrc.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer image alt" ]
                    Input [ attr "data-testid" "word-default-footer-image-alt"; valueS state.DefaultFooterImageAlt; onInput (fun value -> state.DefaultFooterImageAlt.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer image width" ]
                    Input [ typ "number"; attr "data-testid" "word-default-footer-image-width"; valueS state.DefaultFooterImageWidth; onInput (fun value -> state.DefaultFooterImageWidth.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Default footer image height" ]
                    Input [ typ "number"; attr "data-testid" "word-default-footer-image-height"; valueS state.DefaultFooterImageHeight; onInput (fun value -> state.DefaultFooterImageHeight.Value <- value) ]
                ]
                label [] [
                    span [] [ text "First-page header" ]
                    Input [ attr "data-testid" "word-first-header"; valueS state.FirstHeader; onInput (fun value -> state.FirstHeader.Value <- value) ]
                ]
                label [] [
                    span [] [ text "First-page footer" ]
                    Input [ attr "data-testid" "word-first-footer"; valueS state.FirstFooter; onInput (fun value -> state.FirstFooter.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Even-page header" ]
                    Input [ attr "data-testid" "word-even-header"; valueS state.EvenHeader; onInput (fun value -> state.EvenHeader.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Even-page footer" ]
                    Input [ attr "data-testid" "word-even-footer"; valueS state.EvenFooter; onInput (fun value -> state.EvenFooter.Value <- value) ]
                ]
                label [ cls "dialog-inline-field" ] [
                    input [ typ "checkbox"; attr "data-testid" "word-override-section-header-footer"; propS "checked" state.OverrideSection; onChange (fun event -> state.OverrideSection.Value <- checkedValue event) ]
                    span [] [ text "Override current section" ]
                ]
                label [ cls "dialog-inline-field" ] [
                    input [ typ "checkbox"; attr "data-testid" "word-link-section-header-footer"; propS "checked" state.LinkSectionToPrevious; onChange (fun event -> state.LinkSectionToPrevious.Value <- checkedValue event) ]
                    span [] [ text "Link current section to previous" ]
                ]
                label [] [
                    span [] [ text "Current section header" ]
                    Input [ attr "data-testid" "word-section-header"; valueS state.SectionHeader; onInput (fun value -> state.SectionHeader.Value <- value) ]
                ]
                label [] [
                    span [] [ text "Current section footer" ]
                    Input [ attr "data-testid" "word-section-footer"; valueS state.SectionFooter; onInput (fun value -> state.SectionFooter.Value <- value) ]
                ]
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-header-footer-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-header-footer-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

    let annotationDialog overlayTestId dialogTestId (state: AnnotationState) onApply =
        let title () =
            match state.Kind.Value with
            | "comment" -> "Comment"
            | "note" -> "Footnote / endnote"
            | _ -> "Hyperlink"

        DialogWithTestIds overlayTestId dialogTestId state.IsOpen [
            DialogHeader [] [ DialogTitle [] [ textR title ] ]
            DialogBody [] [
                p [ attr "data-testid" "word-annotation-kind" ] [ textR title ]
                showIf (fun () -> state.Kind.Value = "hyperlink")
                    (label [] [
                        span [] [ text "Target" ]
                        Input [ attr "data-testid" "word-annotation-target"; valueS state.Target; onInput (fun value -> state.Target.Value <- value) ]
                    ])
                showIf (fun () -> state.Kind.Value <> "hyperlink")
                    (label [] [
                        span [] [ text "ID" ]
                        Input [ attr "data-testid" "word-annotation-id"; valueS state.AnnotationId; onInput (fun value -> state.AnnotationId.Value <- value) ]
                    ])
                showIf (fun () -> state.Kind.Value = "note")
                    (label [] [
                        span [] [ text "Note type" ]
                        select [ attr "data-testid" "word-annotation-note-type"; propS "value" state.NoteType; onChange (fun event -> state.NoteType.Value <- selectValue event) ] [
                            option [ attr "value" "footnote" ] [ text "Footnote" ]
                            option [ attr "value" "endnote" ] [ text "Endnote" ]
                        ]
                    ])
                showIf (fun () -> state.Kind.Value <> "hyperlink")
                    (label [] [
                        span [] [ text "Text" ]
                        Input [ attr "data-testid" "word-annotation-text"; valueS state.Text; onInput (fun value -> state.Text.Value <- value) ]
                    ])
                showIf (fun () -> state.Kind.Value = "comment")
                    (label [] [
                        span [] [ text "Author" ]
                        Input [ attr "data-testid" "word-annotation-author"; valueS state.Author; onInput (fun value -> state.Author.Value <- value) ]
                    ])
                showIf (fun () -> state.Kind.Value = "comment")
                    (label [] [
                        span [] [ text "Date" ]
                        Input [ attr "data-testid" "word-annotation-date"; valueS state.Date; onInput (fun value -> state.Date.Value <- value) ]
                    ])
            ]
            DialogFooter [] [
                Button [ attr "data-testid" "word-annotation-cancel"; onClick (fun _ -> state.IsOpen.Value <- false) ] [ text "Cancel" ]
                Button [ attr "data-testid" "word-annotation-apply"; onClick (fun _ -> onApply()) ] [ text "Apply" ]
            ]
        ]

#endif
