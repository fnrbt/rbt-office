namespace RedBat.Word

#if FABLE_COMPILER

module Shell =
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types
    open Browser.Dom
    open RedBat.FableUI.Reactive
    open RedBat.FableUI.Dom
    open RedBat.Office.Ui

    module WordApi = RedBat.Word.API
    module WordHost = RedBat.Word.Host
    module WordUi = RedBat.Word.Ui

    let private initialText = "FableUI Word shell proof. Select text or use toolbar commands, then open Format > Paragraph style."

    let private runCommand host command = WordHost.runCommand host command

    let private runAllTextCommand host command =
        WordHost.runCommand host (fun editor ->
            WordApi.selectAll editor
            command editor)

    let private onOff value = if value then "true" else "false"

    let private activeClass value = if value then "rb-word-tool-button-active" else ""

    let private menuActiveClass value = if value then "rb-word-menu-item-active" else ""

    let private parseFloat fallback (value: string) =
        match System.Double.TryParse(value) with
        | true, parsed -> parsed
        | false, _ -> fallback

    let private selectValue (event: Event) = (event.target :?> HTMLSelectElement).value

    let private renderToolbarButton (buttonSpec: WordUi.ButtonSpec) =
        let isActive () = buttonSpec.IsActive |> Option.map (fun active -> active()) |> Option.defaultValue false
        button [ cls "word-tool-button"
                 clsR (fun () -> activeClass (isActive()))
                 attr "data-testid" buttonSpec.TestId
                 attr "data-word-icon" (WordUi.toolbarGlyph buttonSpec.TestId)
                 attrR "aria-pressed" (fun () -> onOff (isActive()))
                 onClick (fun _ -> buttonSpec.Action()) ] [ text buttonSpec.Label ]

    let private renderToolbarSelect (selectSpec: WordUi.SelectSpec) =
        select [ cls "word-tool-select"
                 attr "data-testid" selectSpec.TestId
                 propR "value" (fun () -> box (selectSpec.Value()))
                 onChange (fun event -> selectSpec.Action(selectValue event)) ] [
            children (fun () ->
                selectSpec.Options()
                |> List.map (fun choice -> option [ attr "value" choice.Value ] [ text choice.Label ]))
        ]

    let private renderToolbarInput (inputSpec: WordUi.InputSpec) =
        let inputType =
            match inputSpec.Kind with
            | NumberInput -> "number"
            | ColorInput -> "color"
            | TextInput -> "text"

        let attrs =
            [ cls (if inputSpec.Kind = ColorInput then "word-tool-color" else "word-tool-input")
              typ inputType
              attr "data-testid" inputSpec.TestId
              propR "value" (fun () -> box (inputSpec.Value()))
              onInput inputSpec.Action ]
            @ (inputSpec.Min |> Option.map (fun value -> attr "min" value) |> Option.toList)
            @ (inputSpec.Max |> Option.map (fun value -> attr "max" value) |> Option.toList)

        input attrs

    let private renderToolbarControl = function
        | ToolbarButton buttonSpec -> renderToolbarButton buttonSpec
        | ToolbarSelect selectSpec -> renderToolbarSelect selectSpec
        | ToolbarInput inputSpec -> renderToolbarInput inputSpec
        | ToolbarSeparator -> div [ cls "word-toolbar-separator" ] []

    let private renderToolbarRow (rowSpec: WordUi.ToolbarRowSpec) =
        div [ cls "word-toolbar-row"; attr "data-toolbar-row" rowSpec.Label ] [
            div [ cls "word-toolbar-grip" ] []
            yield! rowSpec.Controls |> List.map renderToolbarControl
        ]

    let private renderMenu (isOpen: Signal<bool>) closeAll (menuSpec: WordUi.MenuSpec) =
        let toggleMenu () =
            let next = not isOpen.Value
            closeAll()
            isOpen.Value <- next

        div [ cls "word-menu-slot" ] [
            button [ cls "word-menu-button"; attr "data-testid" menuSpec.TestId; onClick (fun _ -> toggleMenu()) ] [ text menuSpec.Label ]
            showIf (fun () -> isOpen.Value)
                (div [ cls "word-menu"; attr "data-testid" menuSpec.PanelTestId ] [
                    for itemSpec in menuSpec.Items ->
                        let isChecked () = itemSpec.IsChecked |> Option.map (fun isChecked -> isChecked()) |> Option.defaultValue false
                        button [ cls "word-menu-item"
                                 clsR (fun () -> menuActiveClass (isChecked()))
                                 attr "data-testid" itemSpec.TestId
                                 attrR "aria-checked" (fun () -> onOff (isChecked()))
                                 onClick (fun _ -> itemSpec.Action(); isOpen.Value <- false) ] [ text itemSpec.Label ]
                ])
        ]

    [<Emit("$0.files && $0.files.length > 0 ? $0.files[0] : null")>]
    let private selectedFile (input: HTMLInputElement) : File = jsNative

    [<Emit("new FileReader()")>]
    let private createFileReader () : obj = jsNative

    [<Emit("$0.onload = $1")>]
    let private setReaderOnLoad (reader: obj) (handler: Event -> unit) : unit = jsNative

    [<Emit("$0.onerror = $1")>]
    let private setReaderOnError (reader: obj) (handler: Event -> unit) : unit = jsNative

    [<Emit("$0.readAsDataURL($1)")>]
    let private readAsDataUrl (reader: obj) (file: File) : unit = jsNative

    [<Emit("$0.result")>]
    let private readerResult (reader: obj) : string = jsNative

    type ImageInsertAsset = {
        Src: string
        Width: float
        Height: float
        Alt: string
    }

    type ImagePickerAdapter = {
        OpenPicker: unit -> unit
        PickerView: (ImageInsertAsset -> unit) -> (string -> unit) -> View
    }

    let private browserImagePicker inputId = {
        OpenPicker = fun () ->
            match document.getElementById(inputId) with
            | null -> ()
            | element -> (element :?> HTMLInputElement).click()
        PickerView = fun onPicked onError ->
            let insertImageFile (event: Event) =
                let input = event.target :?> HTMLInputElement
                let file = selectedFile input
                if isNull (box file) then
                    onError "No image selected"
                else
                    let reader = createFileReader ()
                    setReaderOnLoad reader (fun _ ->
                        onPicked {
                            Src = readerResult reader
                            Width = 240.0
                            Height = 160.0
                            Alt = file.name
                        }
                        input.value <- "")
                    setReaderOnError reader (fun _ -> onError (sprintf "Could not read image: %s" file.name))
                    readAsDataUrl reader file

            input [ typ "file"; attr "id" inputId; attr "data-testid" inputId; attr "accept" "image/*"; attr "hidden" "hidden"; onChange insertImageFile ]
    }

    let viewWithAdapters (adapters: ImagePickerAdapter) (host: WordHost.Host) =
        let isFormatMenuOpen = signal false
        let isEditMenuOpen = signal false
        let isFileMenuOpen = signal false
        let isInsertMenuOpen = signal false
        let isTableMenuOpen = signal false
        let isBoldActive = signal false
        let isItalicActive = signal false
        let isUnderlineActive = signal false
        let isAllCapsActive = signal false
        let isSmallCapsActive = signal false
        let isStrikeActive = signal false
        let alignment = signal "left"
        let blockStyle = signal "normal"
        let listStyle = signal "none"
        let numberFormat = signal "decimal"
        let vertAlign = signal "baseline"
        let textColor = signal "#000000"
        let highlightColor = signal "#FFF2CC"
        let fontFamily = signal "Calibri"
        let fontSize = signal "11"

        let paragraphStyleOptions () =
            let _ = host.StyleCatalogVersion.Value
            let catalogStyles =
                match host.Editor with
                | None -> []
                | Some editor ->
                    editor.Doc.Styles
                    |> List.filter (fun style -> style.StyleType = "paragraph")
                    |> List.map (fun style -> WordUi.choice style.Id (style.Name |> Option.defaultValue style.Id))
            (WordUi.paragraphStyles @ catalogStyles)
            |> List.distinctBy (fun choice -> choice.Value)

        let characterStyleOptions () =
            let _ = host.StyleCatalogVersion.Value
            let catalogStyles =
                match host.Editor with
                | None -> []
                | Some editor ->
                    editor.Doc.Styles
                    |> List.filter (fun style -> style.StyleType = "character")
                    |> List.map (fun style -> WordUi.choice style.Id (style.Name |> Option.defaultValue style.Id))
            (WordUi.choice "" "Character style" :: catalogStyles)
            |> List.distinctBy (fun choice -> choice.Value)

        let renderMarkupNavigator () =
            let bookmarkItems () =
                let _ = host.Status.Value
                match host.Editor with
                | None -> [||]
                | Some editor -> WordApi.getBookmarkInfos editor

            let contentControlItems () =
                let _ = host.Status.Value
                match host.Editor with
                | None -> [||]
                | Some editor -> WordApi.getContentControlInfos editor

            let selectBookmarkById bookmarkId bookmarkName =
                runCommand host (fun editor -> WordApi.selectBookmark editor bookmarkId)
                host.Status.Value <- sprintf "Bookmark selected: %s" bookmarkName

            let selectContentControlById controlId label =
                runCommand host (fun editor -> WordApi.selectContentControl editor controlId)
                host.Status.Value <- sprintf "Content control selected: %s" label

            let removeBookmarkById bookmarkId bookmarkName =
                runCommand host (fun editor -> WordApi.removeBookmark editor bookmarkId)
                host.Status.Value <- sprintf "Bookmark removed: %s" bookmarkName

            let removeContentControlById controlId label =
                runCommand host (fun editor -> WordApi.removeContentControl editor controlId)
                host.Status.Value <- sprintf "Content control removed: %s" label

            let editBookmark bookmarkId bookmarkName =
                WordHost.openBookmarkEdit host bookmarkId bookmarkName
                host.Status.Value <- sprintf "Editing bookmark: %s" bookmarkName

            let editContentControl controlId tag alias =
                WordHost.openContentControlEdit host controlId tag alias
                host.Status.Value <- sprintf "Editing content control: %s" (if alias = "" then tag else alias)

            div [ cls "word-markup-navigator"; attr "data-testid" "word-markup-navigator" ] [
                div [ cls "word-markup-navigator-section" ] [
                    div [ cls "word-markup-navigator-title" ] [ text "Bookmarks" ]
                    div [ cls "word-markup-navigator-list"; attr "data-testid" "word-bookmark-list" ] [
                        children (fun () ->
                            let items = bookmarkItems ()
                            if items.Length = 0 then
                                [ div [ cls "word-markup-navigator-empty"; attr "data-testid" "word-bookmark-empty" ] [ text "None" ] ]
                            else
                                items
                                |> Array.toList
                                |> List.map (fun item ->
                                    let bookmarkId = item?Id |> string
                                    let bookmarkName = item?Name |> string
                                    div [ cls "word-markup-navigator-item" ] [
                                        button [ cls "word-markup-navigator-select"; attr "data-testid" "word-bookmark-item"; onClick (fun _ -> selectBookmarkById bookmarkId bookmarkName) ] [ text bookmarkName ]
                                        button [ cls "word-markup-navigator-edit"; attr "data-testid" "word-bookmark-edit"; onClick (fun _ -> editBookmark bookmarkId bookmarkName) ] [ text "Edit" ]
                                        button [ cls "word-markup-navigator-remove"; attr "data-testid" "word-bookmark-remove"; onClick (fun _ -> removeBookmarkById bookmarkId bookmarkName) ] [ text "Remove" ]
                                    ]))
                    ]
                ]
                div [ cls "word-markup-navigator-section" ] [
                    div [ cls "word-markup-navigator-title" ] [ text "Content controls" ]
                    div [ cls "word-markup-navigator-list"; attr "data-testid" "word-content-control-list" ] [
                        children (fun () ->
                            let items = contentControlItems ()
                            if items.Length = 0 then
                                [ div [ cls "word-markup-navigator-empty"; attr "data-testid" "word-content-control-empty" ] [ text "None" ] ]
                            else
                                items
                                |> Array.toList
                                |> List.map (fun item ->
                                    let controlId = item?Id |> string
                                    let tag = item?Tag |> string
                                    let alias = item?Alias |> string
                                    let label = if alias = "" then tag else alias
                                    div [ cls "word-markup-navigator-item" ] [
                                        button [ cls "word-markup-navigator-select"; attr "data-testid" "word-content-control-item"; onClick (fun _ -> selectContentControlById controlId label) ] [ text label ]
                                        button [ cls "word-markup-navigator-edit"; attr "data-testid" "word-content-control-edit"; onClick (fun _ -> editContentControl controlId tag alias) ] [ text "Edit" ]
                                        button [ cls "word-markup-navigator-remove"; attr "data-testid" "word-content-control-remove"; onClick (fun _ -> removeContentControlById controlId label) ] [ text "Remove" ]
                                    ]))
                    ]
                ]
            ]

        let runAlignment value command status =
            alignment.Value <- value
            host.ParagraphStyle.Alignment.Value <- value
            runCommand host command
            host.Status.Value <- status

        let applyHeading level =
            blockStyle.Value <- sprintf "heading-%d" level
            runCommand host (fun editor -> WordApi.setHeadingLevel editor level)
            host.Status.Value <- sprintf "Heading %d applied" level

        let applyHeading1 () = applyHeading 1

        let applyHeading2 () = applyHeading 2

        let applyHeading3 () = applyHeading 3

        let applyNormalParagraph () =
            blockStyle.Value <- "normal"
            runCommand host WordApi.clearHeading
            host.Status.Value <- "Normal paragraph applied"

        let applyParagraphStyle styleId =
            blockStyle.Value <- styleId
            runCommand host (fun editor -> WordApi.setParagraphStyleFor editor styleId)
            host.Status.Value <- sprintf "Paragraph style: %s" styleId

        let applyCharacterStyle styleId =
            if styleId <> "" then
                runAllTextCommand host (fun editor -> WordApi.setCharacterStyleFor editor styleId)
                host.Status.Value <- sprintf "Character style: %s" styleId

        let toggleBullet () =
            listStyle.Value <- if listStyle.Value = "bullet" then "none" else "bullet"
            runCommand host WordApi.toggleBulletList
            host.Status.Value <- "Bullet list toggled"

        let toggleNumbered () =
            listStyle.Value <- if listStyle.Value = "numbered" then "none" else "numbered"
            runCommand host WordApi.toggleNumberedList
            host.Status.Value <- "Numbered list toggled"

        let applyNumberFormat value =
            numberFormat.Value <- value
            listStyle.Value <- "numbered"
            runCommand host (fun editor -> WordApi.setNumberFormat editor value)
            host.Status.Value <- sprintf "Number format: %s" value

        let increaseListIndent () =
            runCommand host WordApi.increaseIndent
            host.Status.Value <- "List indent increased"

        let decreaseListIndent () =
            runCommand host WordApi.decreaseIndent
            host.Status.Value <- "List indent decreased"

        let insertTableRow () =
            runCommand host WordApi.insertTableRowAfter
            host.Status.Value <- "Table row inserted"

        let insertTableColumn () =
            runCommand host WordApi.insertTableColumnAfter
            host.Status.Value <- "Table column inserted"

        let deleteTableRow () =
            runCommand host WordApi.deleteTableRow
            host.Status.Value <- "Table row deleted"

        let deleteTableColumn () =
            runCommand host WordApi.deleteTableColumn
            host.Status.Value <- "Table column deleted"

        let openImagePicker () = adapters.OpenPicker()

        let insertImageAsset asset =
            runCommand host (fun editor -> WordApi.insertImage editor asset.Src asset.Width asset.Height asset.Alt)
            host.Status.Value <- sprintf "Image inserted: %s" asset.Alt

        let reportImageError message = host.Status.Value <- message

        let insertDateTime () =
            let value = System.DateTime.Now.ToString("yyyy-MM-dd HH:mm")
            runCommand host (fun editor -> WordApi.insertFieldAt editor " DATE \\@ \"yyyy-MM-dd HH:mm\" " value)
            host.Status.Value <- sprintf "Date/time inserted: %s" value

        let insertPageNumber () =
            runCommand host (fun editor -> WordApi.insertFieldAt editor " PAGE " "1")
            host.Status.Value <- "Page number field inserted"

        let refreshFields () =
            runCommand host WordApi.refreshFields
            host.Status.Value <- "Fields refreshed"

        let insertTextBox () =
            runCommand host (fun editor -> WordApi.insertTextBox editor "Text box")
            host.Status.Value <- "Text box inserted"

        let insertSectionBreak kind =
            runCommand host (fun editor -> WordApi.setSectionBreak editor kind)
            host.Status.Value <-
                match kind with
                | "continuous" -> "Continuous section break inserted"
                | _ -> "Next-page section break inserted"

        let openAnnotation kind =
            runCommand host WordApi.selectAll
            WordHost.openAnnotation kind host
            host.Status.Value <-
                match kind with
                | "comment" -> "Comment dialog opened"
                | "note" -> "Note dialog opened"
                | _ -> "Hyperlink dialog opened"

        let applyFontFamily family =
            fontFamily.Value <- family
            runAllTextCommand host (fun editor -> WordApi.setFontFamilyFor editor family)
            host.Status.Value <- sprintf "Font family: %s" family

        let applyFontSize value =
            fontSize.Value <- value
            let size = parseFloat 11.0 value
            runAllTextCommand host (fun editor -> WordApi.setFontSizeFor editor size)
            host.Status.Value <- sprintf "Font size: %spt" value

        let applyVertAlign value =
            vertAlign.Value <- value
            runAllTextCommand host (fun editor -> WordApi.setVertAlignFor editor value)
            host.Status.Value <-
                match value with
                | "superscript" -> "Superscript applied"
                | "subscript" -> "Subscript applied"
                | _ -> "Baseline applied"

        let applyTextColor value =
            textColor.Value <- value
            runAllTextCommand host (fun editor -> WordApi.setColorFor editor value)
            host.Status.Value <- sprintf "Text color: %s" value

        let applyHighlight value =
            highlightColor.Value <- value
            runAllTextCommand host (fun editor -> WordApi.setHighlightFor editor value)
            host.Status.Value <- sprintf "Highlight: %s" value

        let toggleBold () =
            isBoldActive.Value <- not isBoldActive.Value
            runAllTextCommand host WordApi.toggleBold
            host.Status.Value <- "Bold applied"

        let toggleItalic () =
            isItalicActive.Value <- not isItalicActive.Value
            runAllTextCommand host WordApi.toggleItalic
            host.Status.Value <- "Italic applied"

        let toggleUnderline () =
            isUnderlineActive.Value <- not isUnderlineActive.Value
            runAllTextCommand host WordApi.toggleUnderline
            host.Status.Value <- "Underline applied"

        let toggleAllCaps () =
            isAllCapsActive.Value <- not isAllCapsActive.Value
            if isAllCapsActive.Value then isSmallCapsActive.Value <- false
            runAllTextCommand host WordApi.toggleAllCaps
            host.Status.Value <- "All caps toggled"

        let toggleSmallCaps () =
            isSmallCapsActive.Value <- not isSmallCapsActive.Value
            if isSmallCapsActive.Value then isAllCapsActive.Value <- false
            runAllTextCommand host WordApi.toggleSmallCaps
            host.Status.Value <- "Small caps toggled"

        let toggleStrike () =
            isStrikeActive.Value <- not isStrikeActive.Value
            runAllTextCommand host WordApi.toggleStrike
            host.Status.Value <- "Strikethrough applied"

        let clearFormatting () =
            isBoldActive.Value <- false
            isItalicActive.Value <- false
            isUnderlineActive.Value <- false
            isAllCapsActive.Value <- false
            isSmallCapsActive.Value <- false
            isStrikeActive.Value <- false
            vertAlign.Value <- "baseline"
            textColor.Value <- "#000000"
            highlightColor.Value <- "#FFF2CC"
            fontFamily.Value <- "Calibri"
            fontSize.Value <- "11"
            runAllTextCommand host WordApi.clearFormatting
            host.Status.Value <- "Formatting cleared"

        let applyAlignment value =
            match value with
            | "center" -> runAlignment "center" WordApi.alignCenter "Aligned center"
            | "right" -> runAlignment "right" WordApi.alignRight "Aligned right"
            | "justify" -> runAlignment "justify" WordApi.alignJustify "Aligned justify"
            | _ -> runAlignment "left" WordApi.alignLeft "Aligned left"

        let shellState: WordUi.ShellState = {
            IsBoldActive = fun () -> isBoldActive.Value
            IsItalicActive = fun () -> isItalicActive.Value
            IsUnderlineActive = fun () -> isUnderlineActive.Value
            IsAllCapsActive = fun () -> isAllCapsActive.Value
            IsSmallCapsActive = fun () -> isSmallCapsActive.Value
            IsStrikeActive = fun () -> isStrikeActive.Value
            Alignment = fun () -> alignment.Value
            BlockStyle = fun () -> blockStyle.Value
            ListStyle = fun () -> listStyle.Value
            NumberFormat = fun () -> numberFormat.Value
            VertAlign = fun () -> vertAlign.Value
            TextColor = fun () -> textColor.Value
            HighlightColor = fun () -> highlightColor.Value
            FontFamily = fun () -> fontFamily.Value
            FontSize = fun () -> fontSize.Value
            IsTrackChangesActive = fun () -> host.TrackChanges.Value
            ParagraphStyles = paragraphStyleOptions
            CharacterStyles = characterStyleOptions
            RecentFiles = fun () -> []
            OpenRecentFile = fun _ -> ()
        }

        let shellActions: WordUi.ShellActions = {
            OpenFile = fun () -> host.Status.Value <- "Open file: use the file input above"
            SaveFile = fun () -> host.Status.Value <- "Save: not yet wired (use export buttons)"
            PrintPdf = fun () -> WordHost.printPdf host
            OpenFindReplace = fun () -> WordHost.openFindReplace host; host.Status.Value <- "Find and replace opened"
            OpenImagePicker = openImagePicker
            OpenImageSize = fun () -> WordHost.openImageSize host
            OpenPageSetup = fun () -> WordHost.openPageSetup host; host.Status.Value <- "Page setup opened"
            OpenHeaderFooter = fun () -> WordHost.openHeaderFooter host; host.Status.Value <- "Headers and footers opened"
            InsertDateTime = insertDateTime
            InsertPageNumber = insertPageNumber
            InsertTextBox = insertTextBox
            OpenCaption = fun () -> WordHost.openCaption host
            OpenCrossReference = fun () -> WordHost.openCrossReference host
            OpenBookmark = fun () -> WordHost.openBookmark host
            OpenContentControl = fun () -> WordHost.openContentControl host
            OpenField = fun () -> WordHost.openField host
            RefreshFields = refreshFields
            OpenEquation = fun () -> WordHost.openEquation host
            OpenRevision = fun () -> WordHost.openRevision host
            ToggleTrackChanges = fun () -> WordHost.toggleTrackChanges host
            SelectNextRevision = fun () -> WordHost.selectNextRevision host
            SelectPreviousRevision = fun () -> WordHost.selectPreviousRevision host
            AcceptSelectedRevision = fun () -> WordHost.acceptSelectedRevision host
            RejectSelectedRevision = fun () -> WordHost.rejectSelectedRevision host
            AcceptAllRevisions = fun () -> WordHost.acceptAllRevisions host
            RejectAllRevisions = fun () -> WordHost.rejectAllRevisions host
            OpenTextBoxStyle = fun () -> WordHost.openTextBoxStyle host
            InsertSectionBreak = insertSectionBreak
            OpenAnnotation = openAnnotation
            OpenParagraphStyle = fun () -> WordHost.openParagraphStyle host
            OpenTableCellStyle = fun () -> WordHost.openTableCellStyle host
            OpenTableMargins = fun () -> WordHost.openTableMargins host
            OpenInsertTable = fun () -> WordHost.openInsertTable host
            InsertDefaultTable = fun () -> WordHost.insertTable host
            InsertTableRow = insertTableRow
            InsertTableColumn = insertTableColumn
            DeleteTableRow = deleteTableRow
            DeleteTableColumn = deleteTableColumn
            ApplyHeading = applyHeading
            ApplyNormalParagraph = applyNormalParagraph
            ApplyParagraphStyle = applyParagraphStyle
            ApplyCharacterStyle = applyCharacterStyle
            ToggleBullet = toggleBullet
            ToggleNumbered = toggleNumbered
            ApplyNumberFormat = applyNumberFormat
            IncreaseListIndent = increaseListIndent
            DecreaseListIndent = decreaseListIndent
            ToggleBold = toggleBold
            ToggleItalic = toggleItalic
            ToggleUnderline = toggleUnderline
            ToggleAllCaps = toggleAllCaps
            ToggleSmallCaps = toggleSmallCaps
            ToggleStrike = toggleStrike
            ClearFormatting = clearFormatting
            ApplyVertAlign = applyVertAlign
            ApplyAlignment = applyAlignment
            ApplyFontFamily = applyFontFamily
            ApplyFontSize = applyFontSize
            ApplyTextColor = applyTextColor
            ApplyHighlight = applyHighlight
        }

        let menuSpecs = WordUi.menus shellState shellActions
        let menuOpenStates = menuSpecs |> List.map (fun _ -> signal false)
        let closeAllMenus () = menuOpenStates |> List.iter (fun state -> state.Value <- false)
        let toolbarRows = WordUi.toolbarRows shellState shellActions

        let shellAttrs =
            [ cls "word-shell"; attr "data-testid" "word-shell" ]
            @ (WordUi.chromeCssVariables |> List.map (fun (name, value) -> style name value))

        div shellAttrs [
            div [ cls "word-titlebar" ] [
                div [ cls "word-titlebar-app" ] [ text "RedBat Word - Document1" ]
                div [ cls "word-titlebar-document" ] [ text "" ]
            ]
            div [ cls "word-menubar" ] [
                yield! List.zip menuOpenStates menuSpecs |> List.map (fun (isOpen, menuSpec) -> renderMenu isOpen closeAllMenus menuSpec)
            ]

            div [ cls "word-toolbar"; attr "data-testid" "word-toolbar" ] [
                yield! toolbarRows |> List.map renderToolbarRow
                yield adapters.PickerView insertImageAsset reportImageError
            ]

            div [ cls "word-ruler" ] [
                div [ cls "word-ruler-margin" ] []
                div [ cls "word-ruler-track" ] []
            ]

            WordHost.view host

            renderMarkupNavigator ()

            div [ cls "word-status"; attr "data-testid" "word-status" ] [ textR (fun () -> host.Status.Value) ]
        ]

    let view host = viewWithAdapters (browserImagePicker "word-image-file-input") host

    let mutable private mountedHost: WordHost.Host option = None

    let mount selector =
        let host = WordHost.create initialText "word-shell"
        mountedHost <- Some host
        mountApp selector (view host)

    let exportDocument () : obj array =
        mountedHost |> Option.map WordHost.exportDocument |> Option.defaultValue [||]

    let exportDocumentPackage () =
        mountedHost |> Option.bind WordHost.exportDocumentPackage |> Option.defaultValue null

    let importDocumentPackage package =
        mountedHost
        |> Option.iter (fun host ->
            WordHost.runCommand host (fun editor -> WordApi.importFromDocxPackage editor package)
            host.StyleCatalogVersion.Value <- host.StyleCatalogVersion.Value + 1
            host.Status.Value <- "Document package imported")

    let exportPageSetup () =
        mountedHost |> Option.bind WordHost.exportPageSetup |> Option.defaultValue null

    let exportHeaderFooter () =
        mountedHost |> Option.bind WordHost.exportHeaderFooter |> Option.defaultValue null

    let exportCommentAnchors () =
        mountedHost |> Option.map WordHost.exportCommentAnchors |> Option.defaultValue [||]

    let exportRevisionInfos () =
        mountedHost |> Option.map WordHost.exportRevisionInfos |> Option.defaultValue [||]

    let exportSelection () =
        mountedHost
        |> Option.bind (fun host -> host.Editor |> Option.map WordApi.getSelection)
        |> Option.defaultValue null

    let printPdf () =
        mountedHost |> Option.iter WordHost.printPdf

#endif
