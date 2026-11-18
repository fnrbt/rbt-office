namespace Rbt.Word

#if FABLE_COMPILER

module Shell =
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types
    open Browser.Dom
    open Rbt.FableUI.Reactive
    open Rbt.FableUI.Dom
    open Rbt.Office.Ui

    module WordApi = Rbt.Word.API
    module WordHost = Rbt.Word.Host
    module WordUi = Rbt.Word.Ui

    let private initialText = "FableUI Word shell proof. Select text or use toolbar commands, then open Format > Paragraph style."

    let private shellCss = """
.rb-word-shell { display: flex; flex-direction: column; min-height: 0; border: 1px solid var(--office-border); }
.rb-word-shell .rb-word-titlebar { display: flex; align-items: baseline; gap: 14px; padding: 10px 16px 5px; min-width: 0; }
.rb-word-shell .rb-word-titlebar-app { color: var(--office-accent); font-weight: 700; white-space: nowrap; }
.rb-word-shell .rb-word-titlebar-document { color: var(--office-muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.rb-word-shell .rb-word-menubar { display: flex; flex-wrap: wrap; gap: 2px; padding: 0 10px 4px; border-bottom: 1px solid var(--office-border); }
.rb-word-shell .rb-word-menu-slot { position: relative; }
.rb-word-shell .rb-word-menu-button[aria-expanded="true"] { background: var(--office-selected); color: var(--office-accent); }
.rb-word-shell .rb-word-menu { position: absolute; z-index: 20; top: 100%; left: 0; width: max-content; min-width: 210px; max-width: min(280px, calc(100vw - 24px)); max-height: min(65vh, 480px); overflow: auto; padding: 5px; border: 1px solid var(--office-border); border-radius: var(--office-radius); background: var(--office-surface); box-shadow: 0 5px 18px rgb(0 0 0 / 16%); }
.rb-word-shell .rb-word-menu-item { display: block; width: 100%; text-align: left; }
.rb-word-shell .rb-word-toolbar { padding: 6px 10px; border-bottom: 1px solid var(--office-border); }
.rb-word-shell .rb-word-toolbar-row { display: flex; flex-wrap: wrap; align-items: center; gap: 3px; min-width: 0; }
.rb-word-shell .rb-word-toolbar-grip { display: none; }
.rb-word-shell .rb-word-tool-button { min-width: 34px; padding-inline: 7px; }
.rb-word-shell .rb-word-tool-select { max-width: 150px; }
.rb-word-shell .rb-word-tool-input { width: 65px; }
.rb-word-shell .rb-word-toolbar-separator { align-self: stretch; width: 1px; min-height: 24px; margin: 3px 4px; background: var(--office-border); }
.rb-word-shell .rb-word-ruler { height: 8px; border-bottom: 1px solid var(--office-border); background: var(--office-desk); }
/* The document area takes what the shell's height leaves (70% of the window when the page does not give the shell a height). */
.rb-word-shell .rb-word-host { flex: 1 1 70vh; min-height: 240px; min-width: 0; max-width: 100%; display: flex; flex-direction: column; background: var(--office-desk); }
/* The workspace scrolls; the canvas is pinned to its visible area and the engine paints the pages under the scroll position (setViewport). */
.rb-word-shell .rb-word-workspace { position: relative; flex: 1 1 auto; min-height: 0; overflow: auto; background: var(--office-desk); }
.rb-word-shell .rb-word-pin { position: sticky; top: 0; left: 0; width: 0; height: 0; }
.rb-word-shell .rb-word-canvas { position: absolute; top: 0; left: 0; display: block; max-width: none; }
.rb-word-shell .rb-word-spacer { pointer-events: none; }
.rb-word-shell .rb-word-markup-navigator, .rb-word-shell .rb-word-revision-pane { display: flex; flex-wrap: wrap; gap: 12px; padding: 10px 16px; border-top: 1px solid var(--office-border); background: var(--office-surface); }
.rb-word-shell .rb-word-markup-navigator-section { flex: 1 1 240px; min-width: 0; }
.rb-word-shell .rb-word-markup-navigator-title { font-weight: 600; }
.rb-word-shell .rb-word-markup-navigator-empty { color: var(--office-muted); }
.rb-word-shell .rb-word-markup-navigator-item { display: flex; flex-wrap: wrap; align-items: center; }
.rb-word-shell .rb-word-markup-navigator-select { flex: 1 1 100px; min-width: 0; text-align: left; overflow-wrap: anywhere; }
.rb-word-shell .rb-word-styles-pane { min-width: 0; max-width: 100%; box-sizing: border-box; padding: 10px 16px; border-top: 1px solid var(--office-border); background: var(--office-surface); }
.rb-word-shell .rb-word-styles-header { display: flex; justify-content: space-between; align-items: center; gap: 8px; }
.rb-word-shell .rb-word-styles-search { box-sizing: border-box; width: 100%; max-width: 320px; }
.rb-word-shell .rb-word-styles-pane > label { display: flex; align-items: center; gap: 6px; margin: 8px 0; }
.rb-word-shell .rb-word-styles-list { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(100%, 210px), 1fr)); gap: 4px; max-height: 260px; overflow: auto; }
.rb-word-shell .rb-word-style-choice { min-width: 0; text-align: left; overflow-wrap: anywhere; }
.rb-word-shell .rb-word-style-choice[aria-pressed="true"] { background: var(--office-selected); }
.rb-word-shell .rb-word-style-info { display: block; color: var(--office-muted); font-size: 12px; }
.rb-word-shell .rb-word-status { display: flex; flex-wrap: wrap; gap: 4px 16px; min-width: 0; overflow-wrap: anywhere; padding: 6px 16px; border-top: 1px solid var(--office-border); color: var(--office-muted); font-size: 12px; }
@media (max-width: 600px) {
  .rb-word-shell .rb-word-titlebar { padding-inline: 10px; }
  .rb-word-shell .rb-word-menu-slot:nth-last-child(-n+2) .rb-word-menu { left: auto; right: 0; }
}
"""

    let mutable private shellSequence = 0

    [<Emit("document.addEventListener('pointerdown', $0, true)")>]
    let private addPointerCapture (_: Event -> unit) : unit = jsNative

    [<Emit("document.removeEventListener('pointerdown', $0, true)")>]
    let private removePointerCapture (_: Event -> unit) : unit = jsNative

    [<Emit("document.addEventListener('focusin', $0, true)")>]
    let private addFocusCapture (_: Event -> unit) : unit = jsNative

    [<Emit("document.removeEventListener('focusin', $0, true)")>]
    let private removeFocusCapture (_: Event -> unit) : unit = jsNative

    [<Emit("$0 && $0.focus()")>]
    let private focusElement (_: Element) : unit = jsNative

    [<Emit("queueMicrotask($0)")>]
    let private afterChromeUpdate (_: unit -> unit) : unit = jsNative

    [<Emit("!!($0 && $0.closest && $0.closest('.rb-word-menu-slot') && $1.contains($0.closest('.rb-word-menu-slot')))")>]
    let private inMenuSlot (_: EventTarget) (_: Element) : bool = jsNative

    let private runCommand host command = WordHost.runCommand host command


    let private onOff value = if value then "true" else "false"

    [<Emit("$0.isComposing || $0.keyCode === 229")>]
    let private keyBelongsToComposition (_: KeyboardEvent) : bool = jsNative

    [<Emit("typeof $0 === 'boolean' ? $0 : null")>]
    let private booleanValue (_: obj) : bool option = jsNative

    [<Emit("$0 == null || $0 === 'mixed' ? '' : String($0)")>]
    let private formatString (_: obj) : string = jsNative

    [<Emit("$0 === 'mixed'")>]
    let private isMixed (_: obj) : bool = jsNative

    [<Emit("$0.defaultPrevented")>]
    let private alreadyHandled (_: Event) : bool = jsNative

    [<Emit("!!($0.closest && $0.closest('input,textarea,select,[contenteditable=true],[role=dialog]'))")>]
    let private isEditingChrome (_: EventTarget) : bool = jsNative
    [<Emit("document.addEventListener('keydown', $0, true)")>]
    let private addKeyCapture (_: Event -> unit) : unit = jsNative

    [<Emit("document.removeEventListener('keydown', $0, true)")>]
    let private removeKeyCapture (_: Event -> unit) : unit = jsNative

    [<Emit("$0.target === $1")>]
    let private isEventTarget (_: Event) (_: obj) : bool = jsNative
    [<Emit("!!($0 && !$0.readOnly && !$0.disabled)")>]
    let private inputEditable (_: obj) : bool = jsNative

    let private activeClass value = if value then "rb-word-tool-button-active" else ""

    let private menuActiveClass value = if value then "rb-word-menu-item-active" else ""


    let private selectValue (event: Event) = (event.target :?> HTMLSelectElement).value

    let private renderToolbarButton pressed (buttonSpec: ButtonSpec) =
        let isActive () = buttonSpec.IsActive |> Option.map (fun active -> active()) |> Option.defaultValue false
        let label = buttonSpec.Tooltip |> Option.defaultValue buttonSpec.Label
        let attrs =
            [ cls "word-tool-button"
              clsR (fun () -> activeClass (isActive()))
              attr "data-testid" buttonSpec.TestId
              attr "data-word-icon" (WordUi.toolbarGlyph buttonSpec.TestId)
              attr "aria-label" label; attr "title" label
              onClick (fun _ -> buttonSpec.Action()) ]
            @ (buttonSpec.IsActive |> Option.map (fun _ -> attrR "aria-pressed" (fun () -> pressed buttonSpec.TestId (isActive()))) |> Option.toList)
            @ (buttonSpec.IsEnabled |> Option.map (fun enabled -> propR "disabled" (fun () -> box (not (enabled())))) |> Option.toList)
        button attrs [ text buttonSpec.Label ]

    let private renderToolbarSelect (selectSpec: SelectSpec) =
        let label =
            match selectSpec.TestId with
            | "word-paragraph-style-select" -> "Paragraph style"
            | "word-font-family" -> "Font family"
            | _ -> "Style"
        select [ cls "word-tool-select"
                 attr "data-testid" selectSpec.TestId
                 attr "aria-label" label; attr "title" label
                 propR "value" (fun () -> box (selectSpec.Value()))
                 onChange (fun event -> selectSpec.Action(selectValue event)) ] [
            children (fun () ->
                let selected = selectSpec.Value()
                let choices = selectSpec.Options()
                let values =
                    if selected <> "" && not (choices |> List.exists (fun choice -> choice.Value = selected)) then
                        choice selected selected :: choices
                    else choices
                let values = choice "" "—" :: (values |> List.filter (fun item -> item.Value <> ""))
                values |> List.map (fun item ->
                    option ([ attr "value" item.Value ]
                            @ (if item.Value = selected then [ attr "selected" "selected" ] else [])
                            @ (if item.IsEnabled then [] else [ attr "disabled" "disabled" ])) [ text item.Label ]))
        ]

    let private renderToolbarInput (inputSpec: InputSpec) =
        let inputType =
            match inputSpec.Kind with
            | NumberInput -> "number"
            | ColorInput -> "color"
            | TextInput -> "text"
        let label =
            match inputSpec.TestId with
            | "word-font-size" -> "Font size (points)"
            | "word-text-color" -> "Text color"
            | "word-highlight-color" -> "Highlight color"
            | _ -> "Formatting value"

        let attrs =
            [ cls (if inputSpec.Kind = ColorInput then "word-tool-color" else "word-tool-input")
              typ inputType
              attr "data-testid" inputSpec.TestId
              attr "aria-label" label; attr "title" label
              propR "value" (fun () -> box (inputSpec.Value()))
              onInput inputSpec.Action ]
            @ (inputSpec.Min |> Option.map (fun value -> attr "min" value) |> Option.toList)
            @ (inputSpec.Max |> Option.map (fun value -> attr "max" value) |> Option.toList)

        input attrs

    let private renderToolbarControl pressed = function
        | ToolbarButton buttonSpec -> renderToolbarButton pressed buttonSpec
        | ToolbarSelect selectSpec -> renderToolbarSelect selectSpec
        | ToolbarInput inputSpec -> renderToolbarInput inputSpec
        | ToolbarSeparator -> div [ cls "word-toolbar-separator"; attr "aria-hidden" "true" ] []

    let private renderToolbarRow pressed (rowSpec: ToolbarRowSpec) =
        div [ cls "word-toolbar-row"; attr "data-toolbar-row" rowSpec.Label; attr "role" "group"; attr "aria-label" "Formatting controls" ] [
            div [ cls "word-toolbar-grip" ] []
            yield! rowSpec.Controls |> List.map (renderToolbarControl pressed)
        ]

    let private renderMenu panelId (isOpen: Signal<bool>) closeAll (menuSpec: MenuSpec) =
        let mutable trigger: Element = null
        let toggleMenu () =
            let next = not isOpen.Value
            closeAll()
            isOpen.Value <- next

        div [ cls "word-menu-slot" ] [
            button [ cls "word-menu-button"; attr "data-testid" menuSpec.TestId
                     refEffect (fun element -> trigger <- element; fun () -> trigger <- null)
                     attr "aria-controls" panelId; attrR "aria-expanded" (fun () -> onOff isOpen.Value)
                     onClick (fun _ -> toggleMenu()) ] [ text menuSpec.Label ]
            div [ cls "word-menu"; attr "id" panelId; attr "role" "group"; attr "aria-label" (menuSpec.Label + " actions")
                  attr "data-testid" menuSpec.PanelTestId; styleR "display" (fun () -> if isOpen.Value then "block" else "none") ] [
                    for itemSpec in menuSpec.Items ->
                        let isChecked () = itemSpec.IsChecked |> Option.map (fun isChecked -> isChecked()) |> Option.defaultValue false
                        let attrs =
                            [ cls "word-menu-item"
                              clsR (fun () -> menuActiveClass (isChecked()))
                              attr "data-testid" itemSpec.TestId
                              onClick (fun _ -> isOpen.Value <- false; focusElement trigger; itemSpec.Action()) ]
                            @ (itemSpec.IsChecked |> Option.map (fun _ -> attrR "aria-pressed" (fun () -> onOff (isChecked()))) |> Option.toList)
                            @ (itemSpec.IsEnabled |> Option.map (fun enabled ->
                                propR "disabled" (fun () ->
                                    let _ = isOpen.Value
                                    box (not (enabled())))) |> Option.toList)
                        button attrs [ text itemSpec.Label ]
                ]
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
        let stylesOpen = signal false
        let styleSearch = signal ""
        let showHiddenStyles = signal false
        let mutable styleSearchElement: Element = null

        let mutable documentPicker: HTMLInputElement option = None
        let openDocumentPicker () =
            WordHost.beginOpen host
            documentPicker |> Option.iter (fun picker -> picker.click())

        let documentSelected (event: Event) =
            let picker = event.target :?> HTMLInputElement
            let file = selectedFile picker
            picker.value <- ""
            if not (isNull (box file)) then
                WordHost.beginOpen host
                WordHost.openFile host file

        let paragraphStyleOptions () =
            let _ = host.StyleCatalogVersion.Value
            let catalogStyles =
                match host.Editor with
                | None -> []
                | Some editor ->
                    editor.Doc.Styles
                    |> List.filter (fun style -> style.StyleType = "paragraph")
                    |> List.map (fun style -> choice style.Id (style.Name |> Option.defaultValue style.Id))
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
                    |> List.map (fun style -> choice style.Id (style.Name |> Option.defaultValue style.Id))
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

            div [ cls "word-markup-navigator"; attr "role" "region"; attr "aria-label" "Document markup"; attr "data-testid" "word-markup-navigator" ] [
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
                                        button [ cls "word-markup-navigator-select"; attr "aria-label" ("Select bookmark " + bookmarkName); attr "data-testid" "word-bookmark-item"; onClick (fun _ -> selectBookmarkById bookmarkId bookmarkName) ] [ text bookmarkName ]
                                        button [ cls "word-markup-navigator-edit"; attr "aria-label" ("Edit bookmark " + bookmarkName); attr "data-testid" "word-bookmark-edit"; onClick (fun _ -> editBookmark bookmarkId bookmarkName) ] [ text "Edit" ]
                                        button [ cls "word-markup-navigator-remove"; attr "aria-label" ("Remove bookmark " + bookmarkName); attr "data-testid" "word-bookmark-remove"; onClick (fun _ -> removeBookmarkById bookmarkId bookmarkName) ] [ text "Remove" ]
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
                                        button [ cls "word-markup-navigator-select"; attr "aria-label" ("Select content control " + label); attr "data-testid" "word-content-control-item"; onClick (fun _ -> selectContentControlById controlId label) ] [ text label ]
                                        button [ cls "word-markup-navigator-edit"; attr "aria-label" ("Edit content control " + label); attr "data-testid" "word-content-control-edit"; onClick (fun _ -> editContentControl controlId tag alias) ] [ text "Edit" ]
                                        button [ cls "word-markup-navigator-remove"; attr "aria-label" ("Remove content control " + label); attr "data-testid" "word-content-control-remove"; onClick (fun _ -> removeContentControlById controlId label) ] [ text "Remove" ]
                                    ]))
                    ]
                ]
            ]

        let runAlignment value command status =
            runCommand host command
            host.Status.Value <- status

        let applyHeading level =
            runCommand host (fun editor -> WordApi.setHeadingLevel editor level)
            host.Status.Value <- sprintf "Heading %d applied" level


        let applyNormalParagraph () =
            runCommand host (fun editor -> WordApi.setParagraphStyleFor editor "")
            host.Status.Value <- "Normal paragraph applied"

        let applyParagraphStyle styleId =
            runCommand host (fun editor -> WordApi.setParagraphStyleFor editor styleId)
            host.Status.Value <- sprintf "Paragraph style: %s" styleId

        let applyCharacterStyle styleId =
            if styleId <> "" then
                runCommand host (fun editor -> WordApi.setCharacterStyleFor editor styleId)
                host.Status.Value <- sprintf "Character style: %s" styleId

        let toggleBullet () =
            runCommand host WordApi.toggleBulletList
            host.Status.Value <- "Bullet list toggled"

        let toggleNumbered () =
            runCommand host WordApi.toggleNumberedList
            host.Status.Value <- "Numbered list toggled"

        let applyNumberFormat value =
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
            runCommand host (fun editor -> WordApi.insertTextBox editor "Text box" |> ignore)
            host.Status.Value <- "Text box inserted"

        let insertSectionBreak kind =
            runCommand host (fun editor -> WordApi.setSectionBreak editor kind)
            host.Status.Value <-
                match kind with
                | "continuous" -> "Continuous section break inserted"
                | _ -> "Next-page section break inserted"

        let openAnnotation kind =
            WordHost.openAnnotation kind host
            host.Status.Value <-
                match kind with
                | "comment" -> "Comment dialog opened"
                | "note" -> "Note dialog opened"
                | _ -> "Hyperlink dialog opened"

        let applyFontFamily family =
            runCommand host (fun editor -> WordApi.setFontFamilyFor editor family)

        let applyFontSize (value: string) =
            match System.Double.TryParse value with
            | true, size when size > 0.0 -> runCommand host (fun editor -> WordApi.setFontSizePt editor size)
            | _ -> ()

        let applyVertAlign value = runCommand host (fun editor -> WordApi.setVertAlignFor editor value)
        let applyTextColor value = runCommand host (fun editor -> WordApi.setColorFor editor value)
        let applyHighlight value = runCommand host (fun editor -> WordApi.setHighlightFor editor value)
        let toggleBold () = runCommand host WordApi.toggleBold
        let toggleItalic () = runCommand host WordApi.toggleItalic
        let toggleUnderline () = runCommand host WordApi.toggleUnderline
        let toggleAllCaps () = runCommand host WordApi.toggleAllCaps
        let toggleSmallCaps () = runCommand host WordApi.toggleSmallCaps
        let toggleStrike () = runCommand host WordApi.toggleStrike
        let clearFormatting () = runCommand host WordApi.clearFormatting

        let applyAlignment value =
            match value with
            | "center" -> runAlignment "center" WordApi.alignCenter "Aligned center"
            | "right" -> runAlignment "right" WordApi.alignRight "Aligned right"
            | "justify" -> runAlignment "justify" WordApi.alignJustify "Aligned justify"
            | _ -> runAlignment "left" WordApi.alignLeft "Aligned left"
        let mutable cachedEditor: WordApi.Editor option = None
        let mutable cachedVersion = -1
        let mutable cachedFormat: obj = null
        let selection () =
            let version = host.SelectionVersion.Value
            let editor = host.Editor
            let sameEditor =
                match cachedEditor, editor with
                | Some previous, Some current -> obj.ReferenceEquals(previous, current)
                | None, None -> true
                | _ -> false
            if version <> cachedVersion || not sameEditor then
                cachedEditor <- editor
                cachedVersion <- version
                cachedFormat <- editor |> Option.map WordApi.getSelectionFormat |> Option.defaultValue null
            cachedFormat

        let field name =
            let format = selection ()
            if isNull format then null else format?(name)

        let paragraphField name =
            let format = selection ()
            if isNull format then null else
                let paragraph: obj = format?Paragraph
                if isNull paragraph then null else paragraph?(name)

        let checkedField name () = booleanValue (field name)
        let fieldText name = field name |> formatString
        let paragraphText name = paragraphField name |> formatString
        let registry = Rbt.Office.Commands.CommandRegistry()
        let enabled () = host.Editor.IsSome
        let canEdit () =
            let _ = host.SelectionVersion.Value
            host.Editor |> Option.exists (fun editor -> inputEditable (WordApi.getInputElement editor) && editor.Doc.Selection.IsSome)
        let inTable () =
            canEdit () && (host.Editor |> Option.exists (fun editor -> (Rbt.Word.Document.caretContext editor.Doc).InTable))
        let add id label shortcut action =
            registry.Add(Rbt.Office.Commands.command id label shortcut enabled (fun _ -> action()))
        let addValue id label action =
            registry.Add(Rbt.Office.Commands.command id label None enabled (fun value -> action (unbox<string> value)))
        let addToggle id label shortcut fieldName action =
            registry.Add(Rbt.Office.Commands.toggle id label shortcut enabled (checkedField fieldName) (fun _ -> action()))
        add "file.open" "Open" (Some "Ctrl+O") openDocumentPicker
        add "file.save" "Save" (Some "Ctrl+S") (fun () -> WordHost.saveFile host)
        add "file.print" "Print" (Some "Ctrl+P") (fun () -> WordHost.printPdf host)
        add "edit.find" "Find" (Some "Ctrl+F") (fun () -> WordHost.openFindReplace host)
        registry.Add(Rbt.Office.Commands.command "insert.link" "Hyperlink" (Some "Ctrl+K") canEdit (fun _ -> openAnnotation "hyperlink"))
        add "edit.replace" "Replace" (Some "Ctrl+H") (fun () -> WordHost.openFindReplace host)
        let addHistory id label shortcut canRun action =
            registry.Add(Rbt.Office.Commands.command id label (Some shortcut)
                (fun () -> host.Editor |> Option.exists canRun)
                (fun _ -> runCommand host (fun editor -> action editor |> ignore)))
        addHistory "edit.undo" "Undo" "Ctrl+Z" WordApi.canUndo WordApi.undo
        addHistory "edit.redo" "Redo" "Ctrl+Y" WordApi.canRedo WordApi.redo
        addToggle "format.bold" "Bold" (Some "Ctrl+B") "Bold" toggleBold
        addToggle "format.italic" "Italic" (Some "Ctrl+I") "Italic" toggleItalic
        addToggle "format.underline" "Underline" (Some "Ctrl+U") "Underline" toggleUnderline
        addHistory "edit.redoAlternate" "Redo" "Ctrl+Shift+Z" WordApi.canRedo WordApi.redo
        addToggle "format.allCaps" "All caps" None "AllCaps" toggleAllCaps
        addToggle "format.smallCaps" "Small caps" None "SmallCaps" toggleSmallCaps
        addToggle "format.strike" "Strike" None "Strike" toggleStrike
        add "format.clear" "Clear formatting" None clearFormatting
        addValue "format.fontFamily" "Font family" applyFontFamily
        addValue "format.fontSize" "Font size" applyFontSize
        addValue "format.color" "Color" applyTextColor
        addValue "format.highlight" "Highlight" applyHighlight
        addValue "format.vertAlign" "Vertical alignment" applyVertAlign
        addValue "format.alignment" "Alignment" applyAlignment
        add "format.paragraphDialog" "Paragraph style" None (fun () -> WordHost.openParagraphStyle host)
        registry.Add(Rbt.Office.Commands.command "format.fontDialog" "Font" None canEdit (fun _ -> WordHost.openFontDialog host))
        registry.Add(Rbt.Office.Commands.command "tools.vba" "VBA editor" None canEdit (fun _ -> WordHost.openVbaEditor host))
        registry.Add(Rbt.Office.Commands.command "format.paragraphBorders" "Borders and shading" None canEdit (fun _ -> WordHost.openParagraphBorders host))
        registry.Add(Rbt.Office.Commands.command "format.listsDialog" "Bullets and numbering" None canEdit (fun _ -> WordHost.openListsDialog host))
        add "format.normal" "Normal paragraph" None applyNormalParagraph
        add "format.bullet" "Bullet list" None toggleBullet
        add "format.numbered" "Numbered list" None toggleNumbered
        add "format.increaseIndent" "Increase indent" None increaseListIndent
        add "format.decreaseIndent" "Decrease indent" None decreaseListIndent
        registry.Add(Rbt.Office.Commands.command "format.paragraphStyle" "Paragraph style" None canEdit (fun value -> applyParagraphStyle (unbox<string> value)))
        addValue "format.numberFormat" "Number format" applyNumberFormat
        registry.Add(Rbt.Office.Commands.command "format.heading" "Heading" None enabled (fun value -> applyHeading (unbox<int> value)))
        registry.Add(Rbt.Office.Commands.command "format.characterStyle" "Character style" None canEdit (fun value -> applyCharacterStyle (unbox<string> value)))
        registry.Add(Rbt.Office.Commands.command "format.stylesPane" "Styles pane" None enabled (fun _ ->
            stylesOpen.Value <- not stylesOpen.Value
            if stylesOpen.Value then afterChromeUpdate (fun () -> if stylesOpen.Value then focusElement styleSearchElement)))
        let run id () = registry.Run id |> ignore
        let runValue id value = registry.Run(id, argument = box value) |> ignore
        let isChecked id =
            let _ = host.SelectionVersion.Value
            registry.State id |> Option.bind (fun state -> state.Checked) |> Option.defaultValue false

        let shellState: WordUi.ShellState = {
            IsBoldActive = fun () -> isChecked "format.bold"
            IsItalicActive = fun () -> isChecked "format.italic"
            IsUnderlineActive = fun () -> isChecked "format.underline"
            IsAllCapsActive = fun () -> isChecked "format.allCaps"
            IsSmallCapsActive = fun () -> isChecked "format.smallCaps"
            IsStrikeActive = fun () -> isChecked "format.strike"
            Alignment = fun () -> paragraphText "Alignment"
            BlockStyle = fun () ->
                let style = paragraphText "StyleId"
                let heading = paragraphText "HeadingLevel"
                if isMixed (paragraphField "HeadingLevel") || isMixed (paragraphField "StyleId") then "" elif style <> "" then style
                else
                    match host.Editor |> Option.bind (fun editor -> (WordStyles.styleContext editor.Doc).DefaultParagraphStyleId) with
                    | Some defaultId -> defaultId
                    | None when heading <> "" && heading <> "0" -> "heading-" + heading
                    | None -> "normal"
            ListStyle = fun () -> paragraphText "ListKind"
            NumberFormat = fun () -> paragraphText "ListNumberFormat"
            VertAlign = fun () -> let value = fieldText "VertAlign" in if value = "" && not (isMixed (field "VertAlign")) then "baseline" else value
            TextColor = fun () -> fieldText "Color"
            HighlightColor = fun () -> fieldText "Highlight"
            FontFamily = fun () -> fieldText "FontFamily"
            FontSize = fun () -> fieldText "FontSizePt"
            IsTrackChangesActive = fun () -> host.TrackChanges.Value
            IsInTable = inTable
            CanEdit = canEdit
            ParagraphStyles = paragraphStyleOptions
            CharacterStyles = characterStyleOptions
            RecentFiles = fun () -> []
            OpenRecentFile = fun _ -> ()
        }

        let baseActions: WordUi.ShellActions = {
            OpenFile = run "file.open"
            SaveFile = run "file.save"
            PrintPdf = run "file.print"
            OpenFindReplace = run "edit.find"
            OpenImagePicker = openImagePicker
            OpenImageSize = fun () -> WordHost.openImageSize host
            OpenPageSetup = fun () -> WordHost.openPageSetup host
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
            OpenParagraphStyle = run "format.paragraphDialog"
            OpenFontDialog = Some (run "format.fontDialog")
            OpenVbaEditor = Some (run "tools.vba")
            OpenParagraphBorders = Some (run "format.paragraphBorders")
            OpenListsDialog = Some (run "format.listsDialog")
            OpenTableCellStyle = fun () -> WordHost.openTableCellStyle host
            OpenTableMargins = fun () -> WordHost.openTableMargins host
            OpenInsertTable = fun () -> WordHost.openInsertTable host
            InsertDefaultTable = fun () -> WordHost.insertTable host
            InsertTableRow = insertTableRow
            InsertTableColumn = insertTableColumn
            DeleteTableRow = deleteTableRow
            DeleteTableColumn = deleteTableColumn
            ApplyHeading = runValue "format.heading"
            ApplyNormalParagraph = run "format.normal"
            ApplyParagraphStyle = runValue "format.paragraphStyle"
            ApplyCharacterStyle = runValue "format.characterStyle"
            ToggleBullet = run "format.bullet"
            ToggleNumbered = run "format.numbered"
            ApplyNumberFormat = runValue "format.numberFormat"
            IncreaseListIndent = run "format.increaseIndent"
            DecreaseListIndent = run "format.decreaseIndent"
            ToggleBold = run "format.bold"
            ToggleItalic = run "format.italic"
            ToggleUnderline = run "format.underline"
            ToggleAllCaps = run "format.allCaps"
            ToggleSmallCaps = run "format.smallCaps"
            ToggleStrike = run "format.strike"
            ClearFormatting = run "format.clear"
            ApplyVertAlign = runValue "format.vertAlign"
            ApplyAlignment = runValue "format.alignment"
            ApplyFontFamily = runValue "format.fontFamily"
            ApplyFontSize = runValue "format.fontSize"
            ApplyTextColor = runValue "format.color"
            ApplyHighlight = runValue "format.highlight"
        }
        // The view specifications invoke the same registry as keyboard entry.
        let otherCommands = [
            "insert.image", baseActions.OpenImagePicker
            "image.size", baseActions.OpenImageSize
            "page.setup", baseActions.OpenPageSetup
            "page.headerFooter", baseActions.OpenHeaderFooter
            "insert.dateTime", baseActions.InsertDateTime
            "insert.pageNumber", baseActions.InsertPageNumber
            "insert.textBox", baseActions.InsertTextBox
            "insert.caption", baseActions.OpenCaption
            "insert.crossReference", baseActions.OpenCrossReference
            "insert.bookmark", baseActions.OpenBookmark
            "insert.contentControl", baseActions.OpenContentControl
            "insert.field", baseActions.OpenField
            "edit.refreshFields", baseActions.RefreshFields
            "insert.equation", baseActions.OpenEquation
            "review.revision", baseActions.OpenRevision
            "review.trackChanges", baseActions.ToggleTrackChanges
            "review.next", baseActions.SelectNextRevision
            "review.previous", baseActions.SelectPreviousRevision
            "review.accept", baseActions.AcceptSelectedRevision
            "review.reject", baseActions.RejectSelectedRevision
            "review.acceptAll", baseActions.AcceptAllRevisions
            "review.rejectAll", baseActions.RejectAllRevisions
            "textBox.style", baseActions.OpenTextBoxStyle
        ]
        for id, action in otherCommands do add id id None action
        for id, action in [
            "table.cellStyle", baseActions.OpenTableCellStyle
            "table.margins", baseActions.OpenTableMargins
            "table.insertRow", baseActions.InsertTableRow
            "table.insertColumn", baseActions.InsertTableColumn
            "table.deleteRow", baseActions.DeleteTableRow
            "table.deleteColumn", baseActions.DeleteTableColumn
        ] do registry.Add(Rbt.Office.Commands.command id id None inTable (fun _ -> action()))
        for id, action in [ "table.insertDialog", baseActions.OpenInsertTable; "table.insert", baseActions.InsertDefaultTable ] do
            registry.Add(Rbt.Office.Commands.command id id None canEdit (fun _ -> action()))
        addValue "insert.sectionBreak" "Section break" baseActions.InsertSectionBreak
        addValue "insert.annotation" "Annotation" baseActions.OpenAnnotation
        let shellActions =
            { baseActions with
                OpenImagePicker = run "insert.image"
                OpenImageSize = run "image.size"
                OpenPageSetup = run "page.setup"
                OpenHeaderFooter = run "page.headerFooter"
                InsertDateTime = run "insert.dateTime"
                InsertPageNumber = run "insert.pageNumber"
                InsertTextBox = run "insert.textBox"
                OpenCaption = run "insert.caption"
                OpenCrossReference = run "insert.crossReference"
                OpenBookmark = run "insert.bookmark"
                OpenContentControl = run "insert.contentControl"
                OpenField = run "insert.field"
                RefreshFields = run "edit.refreshFields"
                OpenEquation = run "insert.equation"
                OpenRevision = run "review.revision"
                ToggleTrackChanges = run "review.trackChanges"
                SelectNextRevision = run "review.next"
                SelectPreviousRevision = run "review.previous"
                AcceptSelectedRevision = run "review.accept"
                RejectSelectedRevision = run "review.reject"
                AcceptAllRevisions = run "review.acceptAll"
                RejectAllRevisions = run "review.rejectAll"
                OpenTextBoxStyle = run "textBox.style"
                InsertSectionBreak = runValue "insert.sectionBreak"
                OpenAnnotation = fun kind -> if kind = "hyperlink" then run "insert.link" () else runValue "insert.annotation" kind
                OpenTableCellStyle = run "table.cellStyle"
                OpenTableMargins = run "table.margins"
                OpenInsertTable = run "table.insertDialog"
                InsertDefaultTable = run "table.insert"
                InsertTableRow = run "table.insertRow"
                InsertTableColumn = run "table.insertColumn"
                DeleteTableRow = run "table.deleteRow"
                DeleteTableColumn = run "table.deleteColumn" }

        let menuSpecs = WordUi.menus shellState shellActions
        shellSequence <- shellSequence + 1
        let instanceId = shellSequence
        let menuOpenStates = menuSpecs |> List.map (fun _ -> signal false)
        let closeAllMenus () = menuOpenStates |> List.iter (fun state -> state.Value <- false)
        let toolbarRows = WordUi.toolbarRows shellState shellActions
        let pressed testId active =
            let id =
                match testId with
                | "word-bold" -> "format.bold"
                | "word-italic" -> "format.italic"
                | "word-underline" -> "format.underline"
                | "word-strike" -> "format.strike"
                | "word-all-caps" -> "format.allCaps"
                | "word-small-caps" -> "format.smallCaps"
                | _ -> ""
            let _ = host.SelectionVersion.Value
            let mixedParagraph =
                match testId with
                | "word-align-left" | "word-align-center" | "word-align-right" | "word-align-justify" ->
                    isMixed (paragraphField "Alignment")
                | "word-bullet-list" | "word-numbered-list" -> isMixed (paragraphField "ListKind")
                | _ -> false
            if mixedParagraph || (id <> "" && (registry.State id |> Option.exists (fun state -> state.Mixed))) then "mixed"
            else onOff active
        let renderStylesPane () =
            let currentParagraph () =
                let id = paragraphText "StyleId"
                if id <> "" then id
                elif shellState.BlockStyle().StartsWith("heading-") && (host.Editor |> Option.exists (fun editor -> (WordStyles.styleContext editor.Doc).DefaultParagraphStyleId.IsNone)) then shellState.BlockStyle()
                else
                    match host.Editor with
                    | Some editor ->
                        let context = WordStyles.styleContext editor.Doc
                        context.DefaultParagraphStyleId |> Option.defaultValue (shellState.BlockStyle())
                    | None -> ""
            let currentCharacter () = fieldText "CharacterStyleId"
            let stateText () =
                let paragraph = if isMixed (paragraphField "StyleId") || isMixed (paragraphField "HeadingLevel") then "mixed" else currentParagraph ()
                let character = if isMixed (field "CharacterStyleId") then "mixed" else currentCharacter ()
                sprintf "Paragraph: %s; Character: %s" (if paragraph = "" then "default" else paragraph) (if character = "" then "none" else character)
            let close () =
                stylesOpen.Value <- false
                let trigger = document.querySelector(sprintf "[data-word-styles-trigger='%d']" instanceId)
                focusElement trigger
            div [ cls "word-styles-pane"; attr "role" "region"; attr "aria-label" "Styles"
                  attr "data-testid" "word-styles-pane"; styleR "display" (fun () -> if stylesOpen.Value then "block" else "none") ] [
                div [ cls "word-styles-header" ] [
                    div [ cls "word-markup-navigator-title" ] [ text "Styles" ]
                    button [ attr "type" "button"; attr "data-testid" "word-styles-close"; onClick (fun _ -> close()) ] [ text "Close" ]
                ]
                p [] [ text "Apply existing paragraph and character styles. Style definitions cannot be edited here." ]
                input [ cls "word-styles-search"; typ "search"; attr "aria-label" "Search styles"; attr "placeholder" "Search styles"; attr "data-testid" "word-styles-search"
                        propR "value" (fun () -> box styleSearch.Value); onInput (fun value -> styleSearch.Value <- value)
                        refEffect (fun element -> styleSearchElement <- element; fun () -> styleSearchElement <- null) ]
                label [] [
                    input [ typ "checkbox"; attr "data-testid" "word-styles-show-hidden"; propR "checked" (fun () -> box showHiddenStyles.Value)
                            onChange (fun event -> showHiddenStyles.Value <- (event.target :?> HTMLInputElement).``checked``) ]
                    text "Show hidden styles"
                ]
                p [ attr "data-testid" "word-styles-current"; attr "aria-live" "polite" ] [ textR stateText ]
                div [ cls "word-styles-list"; attr "data-testid" "word-styles-list" ] [
                    children (fun () ->
                        let _ = host.StyleCatalogVersion.Value
                        let query = styleSearch.Value.Trim().ToLowerInvariant()
                        let matches (name: string) (id: string) = query = "" || name.ToLowerInvariant().Contains(query) || id.ToLowerInvariant().Contains(query)
                        let catalog =
                            match host.Editor with
                            | None -> []
                            | Some editor ->
                                let context = WordStyles.styleContext editor.Doc
                                editor.Doc.Styles
                                |> List.filter (fun item -> (item.StyleType = "paragraph" || item.StyleType = "character") && not (WordStyles.isInternalStyleId item.Id))
                                |> List.choose (fun item ->
                                    let name = item.Name |> Option.defaultValue item.Id
                                    let hidden = item.SemiHidden && not item.UnhideWhenUsed
                                    let current = if item.StyleType = "paragraph" then currentParagraph() = item.Id && not (isMixed (paragraphField "StyleId")) else currentCharacter() = item.Id && not (isMixed (field "CharacterStyleId"))
                                    if not (matches name item.Id) || (hidden && not showHiddenStyles.Value && not current) then None
                                    else
                                        let resolved = if item.StyleType = "paragraph" then Map.tryFind item.Id context.ResolvedParagraphStyles else Map.tryFind item.Id context.ResolvedCharacterStyles
                                        let font = resolved |> Option.bind (fun value -> value.Text) |> Option.bind (fun value -> value.FontFamily)
                                        let info =
                                            [ yield item.StyleType
                                              if hidden then yield "hidden"
                                              if item.IsDefault then yield "default"
                                              match item.BasedOn with Some id -> yield "based on " + id | None -> ()
                                              match font with Some family -> yield "font " + family | None -> () ]
                                            |> String.concat " · "
                                        Some (button [ cls "word-style-choice"; attr "type" "button"; attr "data-testid" "word-style-choice"; attr "data-style-id" item.Id; attr "data-style-type" item.StyleType
                                                       attr "aria-pressed" (onOff current); propR "disabled" (fun () -> box (not (canEdit())))
                                                       onClick (fun _ -> if item.StyleType = "paragraph" then runValue "format.paragraphStyle" item.Id else runValue "format.characterStyle" item.Id) ] [
                                            text name; span [ cls "word-style-info" ] [ text info ]
                                        ]))
                        let builtins =
                            if host.Editor |> Option.exists (fun editor -> editor.Doc.Styles |> List.exists (fun item -> item.StyleType = "paragraph" && not (WordStyles.isInternalStyleId item.Id))) then []
                            else
                                WordUi.paragraphStyles |> List.filter (fun item -> matches item.Label item.Value) |> List.map (fun item ->
                                    button [ cls "word-style-choice"; attr "type" "button"; attr "data-testid" "word-style-choice"; attr "data-style-id" item.Value; attr "data-style-type" "paragraph"
                                             attr "aria-pressed" (onOff (shellState.BlockStyle() = item.Value)); propR "disabled" (fun () -> box (not (canEdit())))
                                             onClick (fun _ -> match item.Value with "heading-1" -> runValue "format.heading" 1 | "heading-2" -> runValue "format.heading" 2 | "heading-3" -> runValue "format.heading" 3 | _ -> run "format.normal" ()) ] [
                                        text item.Label; span [ cls "word-style-info" ] [ text "paragraph · built-in" ]
                                    ])
                        builtins @ catalog)
                ]
            ]


        let shellAttrs =
            [ cls "office-shell"; cls "word-shell"; attr "data-testid" "word-shell"
              refEffect (fun root ->
                  Rbt.FableUI.Chrome.installStyles "word" shellCss
                  let stateSubscription = effect (fun () ->
                      let _ = host.SelectionVersion.Value
                      registry.NotifyChanged())
                  let disposeContextMenu = ContextMenu.attach root host registry
                  let handleKeyDown (event: Event) =
                      let key = event :?> KeyboardEvent
                      let editorTarget = host.Editor |> Option.exists (fun editor -> isEventTarget event (WordApi.getInputElement editor))
                      let ownsTarget = root.contains(event.target :?> Node) || editorTarget
                      if ownsTarget && key.key = "Escape" && stylesOpen.Value && root.querySelector("[data-testid='word-styles-pane']").contains(event.target :?> Node) then
                          stylesOpen.Value <- false
                          focusElement (root.querySelector("[data-testid='word-styles-toggle']"))
                          key.preventDefault()
                          key.stopPropagation()
                      elif ownsTarget && key.key = "Escape" && not (isEditingChrome event.target)
                         && (menuOpenStates |> List.exists (fun state -> state.Value)) then
                          let trigger = root.querySelector(".rb-word-menu-button[aria-expanded='true']")
                          closeAllMenus()
                          focusElement trigger
                          key.preventDefault()
                          key.stopPropagation()
                      elif ownsTarget && (editorTarget || not (isEditingChrome event.target))
                         && not (alreadyHandled event) && not (keyBelongsToComposition key) then
                          let chord = Rbt.Office.Commands.chordOf key.key (key.ctrlKey || key.metaKey) key.shiftKey key.altKey
                          match registry.ForChord chord with
                          | Some command when (command.Id <> "insert.link" || editorTarget) && command.IsEnabled() ->
                              if not key.repeat then registry.HandleKey chord |> ignore
                              key.preventDefault()
                              key.stopPropagation()
                          | _ -> ()
                  let handlePointerDown (event: Event) =
                      if not (inMenuSlot event.target root) then closeAllMenus()
                  addKeyCapture handleKeyDown
                  addPointerCapture handlePointerDown
                  addFocusCapture handlePointerDown
                  fun () ->
                      stateSubscription.Dispose()
                      disposeContextMenu()
                      removeKeyCapture handleKeyDown
                      removePointerCapture handlePointerDown
                      removeFocusCapture handlePointerDown) ]
            @ (WordUi.chromeCssVariables |> List.map (fun (name, value) -> style name value))

        div shellAttrs [
            div [ cls "word-titlebar" ] [
                div [ cls "word-titlebar-app" ] [ text "rbt-office" ]
                div [ cls "word-titlebar-document"; attr "data-testid" "word-document-title" ] [
                    textR (fun () -> host.FileName.Value + (if host.Dirty.Value then " *" else ""))
                ]
            ]
            div [ cls "word-menubar"; attr "role" "group"; attr "aria-label" "Document actions" ] [
                yield! List.zip menuOpenStates menuSpecs |> List.mapi (fun index (isOpen, menuSpec) -> renderMenu (sprintf "rb-word-%d-menu-%d" instanceId index) isOpen closeAllMenus menuSpec)
            ]

            div [ cls "word-toolbar"; attr "data-testid" "word-toolbar" ] [
                yield! toolbarRows |> List.map (renderToolbarRow pressed)
                yield button [ cls "word-tool-button"; attr "type" "button"; attr "data-testid" "word-styles-toggle"; attr "data-word-styles-trigger" (string instanceId)
                               attrR "aria-expanded" (fun () -> onOff stylesOpen.Value); onClick (fun _ -> run "format.stylesPane" ()) ] [ text "Styles" ]
                yield adapters.PickerView insertImageAsset reportImageError
                yield input [ typ "file"; attr "data-testid" "word-document-file-input"
                              attr "accept" ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                              attr "hidden" "hidden"; onChange documentSelected
                              refEffect (fun element ->
                                  documentPicker <- Some (element :?> HTMLInputElement)
                                  fun () -> documentPicker <- None) ]
            ]

            renderStylesPane ()

            div [ cls "word-ruler" ] [
                div [ cls "word-ruler-margin" ] []
                div [ cls "word-ruler-track" ] []
            ]

            WordHost.view host


            renderMarkupNavigator ()

            div [ cls "word-status" ] [
                span [ attr "role" "status"; attr "data-testid" "word-count" ] [
                    textR (fun () ->
                        let _ = host.SelectionVersion.Value
                        match host.Editor with
                        | None -> "Body: 0 words"
                        | Some editor ->
                            let counts = WordApi.getWordCounts editor
                            let body: int = counts?BodyWords
                            let selected: int = counts?SelectedWords
                            let hasSelection: bool = counts?HasSelection
                            let label = if WordApi.isLoadComplete editor then "Body" else "Body (loaded)"
                            let bodyUnit = if body = 1 then "word" else "words"
                            let selectionUnit = if selected = 1 then "word" else "words"
                            if hasSelection then sprintf "%s: %d %s · Selection: %d %s" label body bodyUnit selected selectionUnit
                            else sprintf "%s: %d %s" label body bodyUnit)
                ]
                span [ attr "role" "status"; attr "data-testid" "word-status" ] [ textR (fun () -> host.Status.Value) ]
            ]
        ]

    let view host = viewWithAdapters (browserImagePicker "word-image-file-input") host

    let mutable private mountedHost: WordHost.Host option = None

    let mount selector =
        let host = WordHost.create initialText "word-shell"
        mountedHost <- Some host
        mountApp selector (view host)

    let exportDocumentPackage () =
        mountedHost |> Option.bind WordHost.exportDocumentPackage |> Option.defaultValue null

    let importDocumentPackage package fileName =
        mountedHost |> Option.iter (fun host -> WordHost.importDocumentPackage host package fileName)

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
    let setPdfFonts fonts =
        mountedHost |> Option.iter (fun host -> WordHost.setPdfFonts host fonts)


    let printPdf () =
        mountedHost |> Option.iter WordHost.printPdf

#endif
