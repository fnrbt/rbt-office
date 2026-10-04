namespace Rbt.Word

#if FABLE_COMPILER

module Host =
    open Browser.Types
    open Fable.Core
    open Fable.Core.JsInterop
    open Rbt.FableUI.Reactive
    open Rbt.FableUI.Dom

    module WordApi = Rbt.Word.API
    module WordDialogs = Rbt.Word.Dialogs

    [<Emit("$0 === null || $0 === undefined")>]
    let private isNullOrUndefined (_value: obj) : bool = jsNative

    [<Emit("!new URLSearchParams(window.location.search).has('rbCanvasMeasure') && window.localStorage.getItem('rb.word.canvasMeasure') !== '1'")>]
    let private shapedTextEnabled () : bool = jsNative

    [<Emit("window.RedBatWordTextStack && window.RedBatWordTextStack.install($0).then(function () { $1('Shaped text engine ready') }).catch(function (e) { $1('Shaped text engine failed: ' + ((e && e.message) || e)) })")>]
    let private installShapedText (_editor: WordApi.Editor) (_setStatus: string -> unit) : unit = jsNative

    type Host = {
        mutable Editor: WordApi.Editor option
        mutable Canvas: HTMLCanvasElement option
        ParagraphStyle: WordDialogs.ParagraphStyleState
        FindReplace: WordDialogs.FindReplaceState
        TableCellStyle: WordDialogs.TableCellStyleState
        TableMargins: WordDialogs.TableMarginsState
        InsertTable: WordDialogs.InsertTableState
        ImageSize: WordDialogs.ImageSizeState
        TextBoxStyle: WordDialogs.TextBoxStyleState
        Caption: WordDialogs.CaptionState
        CrossReference: WordDialogs.CrossReferenceState
        Bookmark: WordDialogs.BookmarkState
        ContentControl: WordDialogs.ContentControlState
        Field: WordDialogs.FieldState
        Equation: WordDialogs.EquationState
        Revision: WordDialogs.RevisionState
        PageSetup: WordDialogs.PageSetupState
        HeaderFooter: WordDialogs.HeaderFooterState
        Annotation: WordDialogs.AnnotationState
        mutable FindMatches: obj array
        mutable FindIndex: int
        Status: Signal<string>
        TrackChanges: Signal<bool>
        StyleCatalogVersion: Signal<int>
        CommentAnchors: Signal<obj array>
        RevisionInfos: Signal<obj array>
        mutable EditingBookmarkId: string option
        mutable EditingContentControlId: string option
        InitialText: string
        TestPrefix: string
    }

    let create initialText testPrefix = {
        Editor = None
        Canvas = None
        ParagraphStyle = WordDialogs.paragraphStyleState ()
        FindReplace = WordDialogs.findReplaceState ()
        TableCellStyle = WordDialogs.tableCellStyleState ()
        TableMargins = WordDialogs.tableMarginsState ()
        InsertTable = WordDialogs.insertTableState ()
        ImageSize = WordDialogs.imageSizeState ()
        TextBoxStyle = WordDialogs.textBoxStyleState ()
        Caption = WordDialogs.captionState ()
        CrossReference = WordDialogs.crossReferenceState ()
        Bookmark = WordDialogs.bookmarkState ()
        ContentControl = WordDialogs.contentControlState ()
        Field = WordDialogs.fieldState ()
        Equation = WordDialogs.equationState ()
        Revision = WordDialogs.revisionState ()
        PageSetup = WordDialogs.pageSetupState ()
        HeaderFooter = WordDialogs.headerFooterState ()
        Annotation = WordDialogs.annotationState ()
        FindMatches = [||]
        FindIndex = -1
        Status = signal "Ready"
        TrackChanges = signal false
        StyleCatalogVersion = signal 0
        CommentAnchors = signal [||]
        RevisionInfos = signal [||]
        EditingBookmarkId = None
        EditingContentControlId = None
        InitialText = initialText
        TestPrefix = testPrefix
    }

    let refreshCommentAnchors host =
        host.CommentAnchors.Value <-
            match host.Editor with
            | Some editor -> WordApi.getCommentAnchors editor
            | None -> [||]

    let refreshRevisionInfos host =
        host.RevisionInfos.Value <-
            match host.Editor with
            | Some editor -> WordApi.getRevisionInfos editor
            | None -> [||]

    let refreshReviewState host =
        refreshCommentAnchors host
        refreshRevisionInfos host
        match host.Editor with
        | Some editor -> host.TrackChanges.Value <- WordApi.getTrackChanges editor
        | None -> host.TrackChanges.Value <- false

    let private alignmentLabel value =
        match value with
        | "center" -> "Center"
        | "right" -> "Right"
        | "justify" -> "Justify"
        | _ -> "Left"

    let private parseFloat fallback (value: string) =
        match System.Double.TryParse(value) with
        | true, parsed -> parsed
        | false, _ -> fallback

    let private specialIndentLabel = function
        | "firstLine" -> "first line"
        | "hanging" -> "hanging"
        | _ -> "none"

    let private specialIndentValue (state: WordDialogs.ParagraphStyleState) =
        let amount = abs (parseFloat 0.0 state.SpecialIndentBy.Value)
        match state.SpecialIndent.Value with
        | "firstLine" when amount > 0.0 -> Some amount
        | "hanging" when amount > 0.0 -> Some -amount
        | _ -> None

    let runCommand host command =
        match host.Editor with
        | Some editor -> command editor; refreshReviewState host
        | None -> ()

    let exportDocument host : obj array =
        match host.Editor with
        | Some editor -> WordApi.exportToDocx editor
        | None -> [||]

    let exportDocumentPackage host : obj option =
        host.Editor |> Option.map WordApi.exportDocumentPackage

    let exportPageSetup host : obj option =
        host.Editor |> Option.map WordApi.getPageSetup

    let exportHeaderFooter host : obj option =
        host.Editor |> Option.map WordApi.getHeaderFooterText

    let exportCommentAnchors host : obj array =
        host.CommentAnchors.Value

    let exportRevisionInfos host : obj array =
        host.RevisionInfos.Value

    let private selectRevisionInfo host revision =
        runCommand host (fun editor -> WordApi.selectRevisionInfo editor revision)
        host.Status.Value <- "Revision selected"

    [<Emit("URL.createObjectURL($0)")>]
    let private createObjectUrl (_: obj) : string = jsNative

    [<Emit("window.open($0, '_blank')")>]
    let private openWindow (_: string) : unit = jsNative

    // The PDF exporter needs the Word.Pdf build and font data, which the host
    // application bundles; it registers itself as window.RedBatWordPdfExport,
    // the same way the shaped text engine registers RedBatWordTextStack.
    [<Emit("(window.RedBatWordPdfExport ? window.RedBatWordPdfExport.exportPdfFromEditor($0) : Promise.reject(new Error('PDF export is not installed by the host application')))")>]
    let private exportPdfFromEditor (_: WordApi.Editor) : JS.Promise<obj> = jsNative

    let printPdf host =
        match host.Editor with
        | None -> host.Status.Value <- "Editor not ready"
        | Some editor ->
            host.Status.Value <- "Rendering PDF..."
            async {
                try
                    let! blob = exportPdfFromEditor editor |> Async.AwaitPromise
                    let url = createObjectUrl blob
                    openWindow url
                    host.Status.Value <- "PDF opened"
                with ex ->
                    host.Status.Value <- sprintf "PDF render failed: %s" ex.Message
            }
            |> Async.StartImmediate

    let openParagraphStyle host =
        host.ParagraphStyle.IsOpen.Value <- true

    let openFindReplace host =
        host.FindReplace.IsOpen.Value <- true

    let openTableCellStyle host =
        host.TableCellStyle.IsOpen.Value <- true

    let openTableMargins host =
        host.TableMargins.IsOpen.Value <- true

    let openInsertTable host =
        host.InsertTable.IsOpen.Value <- true

    let openImageSize host =
        host.ImageSize.IsOpen.Value <- true

    let openTextBoxStyle host =
        host.TextBoxStyle.IsOpen.Value <- true

    let openCaption host =
        host.Caption.IsOpen.Value <- true

    let openCrossReference host =
        host.CrossReference.IsOpen.Value <- true

    let openBookmark host =
        host.EditingBookmarkId <- None
        host.Bookmark.IsOpen.Value <- true

    let openContentControl host =
        host.EditingContentControlId <- None
        host.ContentControl.IsOpen.Value <- true

    let openBookmarkEdit host bookmarkId bookmarkName =
        host.EditingBookmarkId <- Some bookmarkId
        host.Bookmark.Name.Value <- bookmarkName
        host.Bookmark.IsOpen.Value <- true

    let openContentControlEdit host controlId tag alias =
        host.EditingContentControlId <- Some controlId
        host.ContentControl.Id.Value <- controlId
        host.ContentControl.Tag.Value <- tag
        host.ContentControl.Alias.Value <- alias
        host.ContentControl.IsOpen.Value <- true

    let openField host =
        host.Field.IsOpen.Value <- true

    let openEquation host =
        host.Equation.IsOpen.Value <- true

    let openRevision host =
        host.Revision.IsOpen.Value <- true

    let toggleTrackChanges host =
        match host.Editor with
        | Some editor ->
            WordApi.setTrackChangesOptions editor host.Revision.Author.Value host.Revision.Date.Value
            let enabled = WordApi.toggleTrackChanges editor
            host.TrackChanges.Value <- enabled
            host.Status.Value <- if enabled then "Track changes on" else "Track changes off"
        | None -> host.Status.Value <- "Editor not ready"

    let openPageSetup host =
        host.PageSetup.IsOpen.Value <- true

    let openHeaderFooter host =
        host.HeaderFooter.IsOpen.Value <- true

    let openAnnotation kind host =
        host.Annotation.Kind.Value <- kind
        host.Annotation.IsOpen.Value <- true

    let insertTable host =
        runCommand host (fun editor -> WordApi.insertTable editor 2 3)
        host.Status.Value <- "Table inserted"

    let insertCustomTable host =
        let state = host.InsertTable
        let rows = parseFloat 2.0 state.Rows.Value |> int |> max 1
        let columns = parseFloat 3.0 state.Columns.Value |> int |> max 1
        runCommand host (fun editor -> WordApi.insertTable editor rows columns)
        host.Status.Value <- sprintf "Table inserted: %d x %d" rows columns
        state.IsOpen.Value <- false

    let applyImageSize host =
        let state = host.ImageSize
        let width = parseFloat 240.0 state.Width.Value |> max 1.0
        let height = parseFloat 160.0 state.Height.Value |> max 1.0
        runCommand host (fun editor -> WordApi.resizeCurrentImage editor width height)
        host.Status.Value <- sprintf "Image size: %.0f x %.0f" width height
        state.IsOpen.Value <- false

    let applyTextBoxStyle host =
        let state = host.TextBoxStyle
        let x = parseFloat 96.0 state.X.Value
        let y = parseFloat 96.0 state.Y.Value
        let width = parseFloat 240.0 state.Width.Value |> max 1.0
        let height = parseFloat 96.0 state.Height.Value |> max 1.0
        let borderWidth = parseFloat 1.0 state.BorderWidth.Value
        runCommand host (fun editor -> WordApi.setCurrentTextBoxStyle editor x y width height state.FillColor.Value state.BorderColor.Value borderWidth)
        host.Status.Value <- sprintf "Text box style: %.0f x %.0f at %.0f, %.0f" width height x y
        state.IsOpen.Value <- false

    let insertCaption host =
        let state = host.Caption
        runCommand host (fun editor -> WordApi.insertCaption editor state.Label.Value state.Text.Value)
        host.Status.Value <- sprintf "%s caption inserted" state.Label.Value
        state.IsOpen.Value <- false

    let insertCrossReference host =
        let state = host.CrossReference
        let mutable inserted = false
        runCommand host (fun editor -> inserted <- WordApi.insertCrossReferenceToLatestCaption editor state.Label.Value)
        host.Status.Value <-
            if inserted then sprintf "%s cross-reference inserted" state.Label.Value
            else sprintf "No %s captions to reference" state.Label.Value
        state.IsOpen.Value <- false

    let applyBookmark host =
        let state = host.Bookmark
        match host.EditingBookmarkId with
        | Some bookmarkId ->
            runCommand host (fun editor -> WordApi.renameBookmark editor bookmarkId state.Name.Value)
            host.Status.Value <- sprintf "Bookmark renamed: %s" state.Name.Value
            host.EditingBookmarkId <- None
        | None ->
            runCommand host (fun editor -> WordApi.setBookmarkFor editor state.Name.Value)
            host.Status.Value <- sprintf "Bookmark set: %s" state.Name.Value
        state.IsOpen.Value <- false

    let applyContentControl host =
        let state = host.ContentControl
        match host.EditingContentControlId with
        | Some controlId ->
            runCommand host (fun editor -> WordApi.updateContentControl editor controlId state.Id.Value state.Tag.Value state.Alias.Value)
            host.Status.Value <- sprintf "Content control updated: %s" state.Tag.Value
            host.EditingContentControlId <- None
        | None ->
            runCommand host (fun editor -> WordApi.setContentControlFor editor state.Id.Value state.Tag.Value state.Alias.Value)
            host.Status.Value <- sprintf "Content control set: %s" state.Tag.Value
        state.IsOpen.Value <- false

    let insertField host =
        let state = host.Field
        runCommand host (fun editor -> WordApi.insertFieldAt editor state.Code.Value state.Result.Value)
        host.Status.Value <- sprintf "Field inserted: %s" state.Code.Value
        state.IsOpen.Value <- false

    let insertEquation host =
        let state = host.Equation
        runCommand host (fun editor -> WordApi.insertEquation editor state.Text.Value)
        host.Status.Value <- "Equation inserted"
        state.IsOpen.Value <- false

    let applyRevision host =
        let state = host.Revision
        runCommand host (fun editor -> WordApi.setRevisionFor editor state.RevisionType.Value state.RevisionId.Value state.Author.Value state.Date.Value)
        host.Status.Value <- sprintf "Revision marked: %s" state.RevisionType.Value
        state.IsOpen.Value <- false

    let acceptSelectedRevision host =
        runCommand host WordApi.acceptSelectedRevision
        host.Status.Value <- "Selected revision accepted"

    let rejectSelectedRevision host =
        runCommand host WordApi.rejectSelectedRevision
        host.Status.Value <- "Selected revision rejected"

    let acceptRevisionInfo host revision =
        runCommand host (fun editor ->
            WordApi.selectRevisionInfo editor revision
            WordApi.acceptSelectedRevision editor)
        host.Status.Value <- "Revision accepted"

    let rejectRevisionInfo host revision =
        runCommand host (fun editor ->
            WordApi.selectRevisionInfo editor revision
            WordApi.rejectSelectedRevision editor)
        host.Status.Value <- "Revision rejected"

    let acceptAllRevisions host =
        runCommand host WordApi.acceptAllRevisions
        host.Status.Value <- "All revisions accepted"

    let rejectAllRevisions host =
        runCommand host WordApi.rejectAllRevisions
        host.Status.Value <- "All revisions rejected"

    let selectNextRevision host =
        runCommand host WordApi.selectNextRevision
        host.Status.Value <- "Next revision selected"

    let selectPreviousRevision host =
        runCommand host WordApi.selectPreviousRevision
        host.Status.Value <- "Previous revision selected"

    let applyPageSetup host =
        let state = host.PageSetup
        let inch value fallback = parseFloat fallback value * 96.0
        let pageWidth = inch state.PageWidth.Value 8.5
        let pageHeight = inch state.PageHeight.Value 11.0
        let marginLeft = inch state.MarginLeft.Value 1.0
        let marginRight = inch state.MarginRight.Value 1.0
        let marginTop = inch state.MarginTop.Value 1.0
        let marginBottom = inch state.MarginBottom.Value 1.0
        let headerDistance = inch state.HeaderDistance.Value 0.5
        let footerDistance = inch state.FooterDistance.Value 0.5
        runCommand host (fun editor -> WordApi.setPageSetupOptions editor pageWidth pageHeight marginLeft marginRight marginTop marginBottom headerDistance footerDistance state.DifferentFirstPage.Value state.EvenAndOddHeaders.Value)
        host.Status.Value <- sprintf "Page setup: %s x %s in, margins %s/%s/%s/%s in" state.PageWidth.Value state.PageHeight.Value state.MarginTop.Value state.MarginRight.Value state.MarginBottom.Value state.MarginLeft.Value
        state.IsOpen.Value <- false

    let applyHeaderFooter host =
        let state = host.HeaderFooter
        runCommand host (fun editor ->
            WordApi.setHeaderFooterText editor state.DefaultHeader.Value state.DefaultFooter.Value state.FirstHeader.Value state.FirstFooter.Value state.EvenHeader.Value state.EvenFooter.Value)
        runCommand host (fun editor ->
            WordApi.setHeaderFooterTableText editor state.DefaultHeaderTableLeft.Value state.DefaultHeaderTableRight.Value state.DefaultFooterTableLeft.Value state.DefaultFooterTableRight.Value)
        runCommand host (fun editor ->
            WordApi.setHeaderFooterImageText editor state.DefaultHeaderImageSrc.Value state.DefaultHeaderImageAlt.Value state.DefaultHeaderImageWidth.Value state.DefaultHeaderImageHeight.Value state.DefaultFooterImageSrc.Value state.DefaultFooterImageAlt.Value state.DefaultFooterImageWidth.Value state.DefaultFooterImageHeight.Value)
        if state.LinkSectionToPrevious.Value then
            runCommand host (fun editor -> WordApi.setCurrentSectionHeaderFooterText editor "" "")
            state.OverrideSection.Value <- false
        elif state.OverrideSection.Value then
            runCommand host (fun editor ->
                WordApi.setCurrentSectionHeaderFooterText editor state.SectionHeader.Value state.SectionFooter.Value)
        host.Status.Value <- "Headers and footers updated"
        state.IsOpen.Value <- false

    let findNext host =
        let state = host.FindReplace
        let query = state.Query.Value
        if System.String.IsNullOrWhiteSpace query then
            host.FindMatches <- [||]
            host.FindIndex <- -1
            state.Result.Value <- "Enter text to find"
            host.Status.Value <- "Enter text to find"
        else
            match host.Editor with
            | None -> ()
            | Some editor ->
                let matches = WordApi.findText editor query state.MatchCase.Value
                host.FindMatches <- matches
                if matches.Length = 0 then
                    host.FindIndex <- -1
                    state.Result.Value <- "No matches"
                    host.Status.Value <- "No matches"
                else
                    host.FindIndex <- (host.FindIndex + 1) % matches.Length
                    WordApi.selectFindMatch editor matches.[host.FindIndex]
                    state.Result.Value <- sprintf "Match %d of %d" (host.FindIndex + 1) matches.Length
                    host.Status.Value <- state.Result.Value

    let replaceAll host =
        let state = host.FindReplace
        let query = state.Query.Value
        if System.String.IsNullOrWhiteSpace query then
            state.Result.Value <- "Enter text to find"
            host.Status.Value <- "Enter text to find"
        else
            match host.Editor with
            | None -> ()
            | Some editor ->
                let count = WordApi.replaceAllText editor query state.Replacement.Value state.MatchCase.Value
                host.FindMatches <- [||]
                host.FindIndex <- -1
                state.Result.Value <- sprintf "Replaced %d matches" count
                host.Status.Value <- state.Result.Value

    let applyParagraphStyle host =
        let state = host.ParagraphStyle
        runCommand host (fun editor ->
            let indentLeft = parseFloat 0.0 state.IndentLeft.Value
            let indentRight = parseFloat 0.0 state.IndentRight.Value
            let spacingBefore = parseFloat 0.0 state.SpacingBefore.Value
            let spacingAfter = parseFloat 10.0 state.SpacingAfter.Value
            let lineSpacing = parseFloat 1.5 state.LineSpacing.Value
            let size = parseFloat 11.0 state.FontSize.Value
            WordApi.applyParagraphStyle editor state.Alignment.Value indentLeft indentRight spacingBefore spacingAfter lineSpacing size (specialIndentValue state) state.BackgroundColor.Value)
        host.Status.Value <- sprintf "Paragraph style: %s, %spt before, %spt after, %.1fx line, %spt, %s indent %spt, bg %s" (alignmentLabel state.Alignment.Value) state.SpacingBefore.Value state.SpacingAfter.Value (parseFloat 1.5 state.LineSpacing.Value) state.FontSize.Value (specialIndentLabel state.SpecialIndent.Value) state.SpecialIndentBy.Value state.BackgroundColor.Value
        state.IsOpen.Value <- false

    let applyTableCellStyle host =
        let state = host.TableCellStyle
        let cellWidth = parseFloat 0.0 state.Width.Value
        let borderWidth = parseFloat 1.0 state.BorderWidth.Value
        let paddingTop = parseFloat 10.0 state.PaddingTop.Value
        let paddingRight = parseFloat 18.0 state.PaddingRight.Value
        let paddingBottom = parseFloat 10.0 state.PaddingBottom.Value
        let paddingLeft = parseFloat 18.0 state.PaddingLeft.Value
        runCommand host (fun editor -> WordApi.setSelectedTableCellStyleWithShading editor state.BackgroundColor.Value state.ShadingPattern.Value state.ShadingForegroundColor.Value cellWidth state.BorderColor.Value borderWidth state.VerticalAlignment.Value paddingTop paddingRight paddingBottom paddingLeft)
        host.Status.Value <- sprintf "Table cell style: %s shading, %s border, %s align, %s width, %s/%s/%s/%s padding" state.ShadingPattern.Value state.BorderColor.Value state.VerticalAlignment.Value state.Width.Value state.PaddingTop.Value state.PaddingRight.Value state.PaddingBottom.Value state.PaddingLeft.Value
        state.IsOpen.Value <- false

    let applyTableMargins host =
        let state = host.TableMargins
        let paddingTop = parseFloat 0.0 state.PaddingTop.Value
        let paddingRight = parseFloat 7.2 state.PaddingRight.Value
        let paddingBottom = parseFloat 0.0 state.PaddingBottom.Value
        let paddingLeft = parseFloat 7.2 state.PaddingLeft.Value
        let cellSpacing = parseFloat 0.0 state.CellSpacing.Value
        runCommand host (fun editor -> WordApi.setCurrentTableCellMarginsAndSpacing editor paddingTop paddingRight paddingBottom paddingLeft cellSpacing)
        host.Status.Value <- sprintf "Table cell margins: %s/%s/%s/%s, spacing %s" state.PaddingTop.Value state.PaddingRight.Value state.PaddingBottom.Value state.PaddingLeft.Value state.CellSpacing.Value
        state.IsOpen.Value <- false

    let applyAnnotation host =
        let state = host.Annotation
        runCommand host (fun editor ->
            match state.Kind.Value with
            | "comment" -> WordApi.setCommentFor editor state.AnnotationId.Value state.Text.Value state.Author.Value state.Date.Value
            | "note" -> WordApi.setNoteFor editor state.AnnotationId.Value state.NoteType.Value state.Text.Value
            | _ -> WordApi.setHyperlinkFor editor state.Target.Value)
        host.Status.Value <-
            match state.Kind.Value with
            | "comment" -> sprintf "Comment applied: %s" state.AnnotationId.Value
            | "note" -> sprintf "%s applied: %s" state.NoteType.Value state.AnnotationId.Value
            | _ -> sprintf "Hyperlink applied: %s" state.Target.Value
        refreshCommentAnchors host
        state.IsOpen.Value <- false

    let private numberProp name (value: obj) = value?(name) |> unbox<float>
    let private boolProp name (value: obj) = value?(name) |> unbox<bool>
    let private stringProp name (value: obj) = value?(name) |> string
    let private optionalStringProp name (value: obj) =
        let prop = value?(name)
        if isNullOrUndefined prop then "" else string prop

    let selectCommentAnchor host (anchor: obj) =
        match host.Editor with
        | None -> ()
        | Some editor ->
            WordApi.selectCommentAnchor editor anchor
            refreshCommentAnchors host
            host.Canvas |> Option.iter (fun canvas -> canvas.focus())
            host.Status.Value <- sprintf "Comment selected: %s" (stringProp "Id" anchor)

    let private commentCallout host (anchor: obj) =
        let id = stringProp "Id" anchor
        let textValue = stringProp "Text" anchor
        let author = optionalStringProp "Author" anchor
        let date = optionalStringProp "Date" anchor
        let excerpt = stringProp "Excerpt" anchor
        let isSelected () = boolProp "IsSelected" anchor
        let left () =
            let canvasLeft = host.Canvas |> Option.map (fun canvas -> canvas.offsetLeft) |> Option.defaultValue 0.0
            sprintf "%.0fpx" (canvasLeft + numberProp "CalloutX" anchor)
        let top () =
            let canvasTop = host.Canvas |> Option.map (fun canvas -> canvas.offsetTop) |> Option.defaultValue 0.0
            sprintf "%.0fpx" (canvasTop + numberProp "CalloutY" anchor)
        div [ cls "word-comment-callout"
              attr "data-testid" "word-comment-callout"
              attr "data-comment-id" id
              attrR "data-active" (fun () -> if isSelected () then "true" else "false")
              attr "role" "button"
              attrR "aria-pressed" (fun () -> if isSelected () then "true" else "false")
              attr "tabindex" "0"
              onClick (fun _ -> selectCommentAnchor host anchor)
              onKeyDown (fun event ->
                  if event.key = "Enter" || event.key = " " then
                      event.preventDefault()
                      selectCommentAnchor host anchor)
              style "position" "absolute"
              styleR "left" left
              styleR "top" top
              style "width" "138px"
              style "min-height" "44px"
              style "box-sizing" "border-box"
              style "padding" "7px 8px"
              styleR "border" (fun () -> if isSelected () then "1px solid #2159AA" else "1px solid #8FA2BC")
              style "border-left" "3px solid #2159AA"
              styleR "background" (fun () -> if isSelected () then "#E8F2FF" else "#FFFFF5")
              style "box-shadow" "0 2px 6px rgb(0 0 0 / 0.16)"
              style "font" "11px Arial, Helvetica, sans-serif"
              style "color" "#111111"
              style "pointer-events" "auto"
              style "cursor" "default" ] [
            div [ cls "word-comment-callout-id"; attr "data-testid" "word-comment-callout-id"; style "font-weight" "700"; style "margin-bottom" "3px" ] [ text id ]
            div [ cls "word-comment-callout-meta"; attr "data-testid" "word-comment-callout-meta"; style "font-size" "10px"; style "color" "#475569"; style "margin-bottom" "3px" ] [ text (if author = "" then date elif date = "" then author else sprintf "%s / %s" author date) ]
            div [ cls "word-comment-callout-text"; attr "data-testid" "word-comment-callout-text" ] [ text textValue ]
            div [ cls "word-comment-callout-excerpt"; attr "data-testid" "word-comment-callout-excerpt"; style "margin-top" "4px"; style "color" "#475569" ] [ text excerpt ]
        ]

    let private commentConnector host (anchor: obj) =
        let id = stringProp "Id" anchor
        let canvasOffset getValue = host.Canvas |> Option.map getValue |> Option.defaultValue 0.0
        let x1 () = sprintf "%.0f" (canvasOffset (fun canvas -> canvas.offsetLeft) + numberProp "X" anchor + numberProp "Width" anchor)
        let y1 () = sprintf "%.0f" (canvasOffset (fun canvas -> canvas.offsetTop) + numberProp "Y" anchor + numberProp "Height" anchor / 2.0)
        let x2 () = sprintf "%.0f" (canvasOffset (fun canvas -> canvas.offsetLeft) + numberProp "CalloutX" anchor)
        let y2 () = sprintf "%.0f" (canvasOffset (fun canvas -> canvas.offsetTop) + numberProp "CalloutY" anchor + 22.0)
        el "line" [ attr "data-testid" "word-comment-connector"
                    attr "data-comment-id" id
                    attrR "x1" x1
                    attrR "y1" y1
                    attrR "x2" x2
                    attrR "y2" y2
                    attr "stroke" "#C79B00"
                    attr "stroke-width" "1"
                    attr "stroke-dasharray" "2 2" ] []

    let private commentLayer host =
        div [ cls "word-comment-layer"
              attr "data-testid" "word-comment-layer"
              style "position" "absolute"
              style "inset" "0"
              style "pointer-events" "none" ] [
            el "svg" [ cls "word-comment-connectors"
                       attr "data-testid" "word-comment-connectors"
                       style "position" "absolute"
                       style "inset" "0"
                       style "width" "100%"
                       style "height" "100%"
                       style "overflow" "visible"
                       style "pointer-events" "none" ] [
                children (fun () -> host.CommentAnchors.Value |> Array.toList |> List.map (commentConnector host))
            ]
            children (fun () -> host.CommentAnchors.Value |> Array.toList |> List.map (commentCallout host))
        ]

    let private revisionValue (revision: obj) name : string =
        revision?(name) |> unbox<string>

    let private revisionItem host index revision =
        let revisionType = revisionValue revision "RevisionType"
        let author = revisionValue revision "Author"
        let date = revisionValue revision "Date"
        let excerpt = revisionValue revision "Text"
        let label =
            match revisionType with
            | "delete" -> "Deleted"
            | "moveFrom" -> "Moved from"
            | "moveTo" -> "Moved to"
            | _ -> "Inserted"
        div [ cls "word-revision-item"; attr "data-testid" "word-revision-item" ] [
            button [ cls "word-revision-summary"; attr "data-testid" "word-revision-summary"; onClick (fun _ -> selectRevisionInfo host revision) ] [
                span [ cls "word-revision-kind"; attr "data-testid" "word-revision-kind" ] [ text label ]
                span [ cls "word-revision-index" ] [ text (sprintf " #%d" (index + 1)) ]
                if author <> "" then span [ cls "word-revision-author"; attr "data-testid" "word-revision-author" ] [ text (" by " + author) ]
                if date <> "" then span [ cls "word-revision-date"; attr "data-testid" "word-revision-date" ] [ text (" @ " + date) ]
                div [ cls "word-revision-excerpt"; attr "data-testid" "word-revision-excerpt" ] [ text excerpt ]
            ]
            div [ cls "word-revision-actions" ] [
                button [ cls "word-revision-accept"; attr "data-testid" "word-revision-pane-accept"; onClick (fun _ -> acceptRevisionInfo host revision) ] [ text "Accept" ]
                button [ cls "word-revision-reject"; attr "data-testid" "word-revision-pane-reject"; onClick (fun _ -> rejectRevisionInfo host revision) ] [ text "Reject" ]
            ]
        ]

    let private revisionPane host =
        div [ cls "word-revision-pane"; attr "data-testid" "word-revision-pane" ] [
            div [ cls "word-revision-pane-header" ] [
                span [] [ textR (fun () -> sprintf "Revisions (%d)" host.RevisionInfos.Value.Length) ]
                button [ attr "data-testid" "word-revision-pane-previous"; onClick (fun _ -> selectPreviousRevision host) ] [ text "Previous" ]
                button [ attr "data-testid" "word-revision-pane-next"; onClick (fun _ -> selectNextRevision host) ] [ text "Next" ]
                button [ attr "data-testid" "word-revision-pane-accept-all"; onClick (fun _ -> acceptAllRevisions host) ] [ text "Accept all" ]
                button [ attr "data-testid" "word-revision-pane-reject-all"; onClick (fun _ -> rejectAllRevisions host) ] [ text "Reject all" ]
            ]
            children (fun () -> host.RevisionInfos.Value |> Array.toList |> List.mapi (revisionItem host))
        ]

    let view host =
        let hostAttrs =
            [ cls "word-host"; attr "data-testid" (host.TestPrefix + "-host") ]
            @ (Ui.chromeCssVariables |> List.map (fun (name, value) -> style name value))

        div hostAttrs [
            div [ cls "word-workspace"; style "position" "relative" ] [
                canvas [ cls "word-canvas"
                         attr "data-testid" (host.TestPrefix + "-canvas")
                         attr "width" "900"
                         attr "height" "560"
                         refEffect (fun element ->
                             let canvas = element :?> HTMLCanvasElement
                             host.Canvas <- Some canvas
                             let editor = WordApi.initializeCanvasWithChange canvas (fun () -> refreshReviewState host)
                             host.Editor <- Some editor
                             WordApi.setContent editor host.InitialText
                             refreshReviewState host
                             host.Status.Value <- "Editor ready"
                             if shapedTextEnabled () then
                                 host.Status.Value <- "Loading shaped text engine"
                                 installShapedText editor (fun status -> host.Status.Value <- status)
                             fun () ->
                                 WordApi.dispose editor
                                 host.Editor <- None
                                 host.Canvas <- None) ] []
                commentLayer host
            ]

            showIf (fun () -> host.RevisionInfos.Value.Length > 0) (revisionPane host)

            WordDialogs.paragraphStyleDialog "word-paragraph-dialog-overlay" "word-paragraph-dialog" host.ParagraphStyle (fun () -> applyParagraphStyle host)
            WordDialogs.findReplaceDialog "word-find-dialog-overlay" "word-find-dialog" host.FindReplace (fun () -> findNext host) (fun () -> replaceAll host)
            WordDialogs.tableCellStyleDialog "word-table-cell-style-dialog-overlay" "word-table-cell-style-dialog" host.TableCellStyle (fun () -> applyTableCellStyle host)
            WordDialogs.tableMarginsDialog "word-table-margins-dialog-overlay" "word-table-margins-dialog" host.TableMargins (fun () -> applyTableMargins host)
            WordDialogs.insertTableDialog "word-insert-table-dialog-overlay" "word-insert-table-dialog" host.InsertTable (fun () -> insertCustomTable host)
            WordDialogs.imageSizeDialog "word-image-size-dialog-overlay" "word-image-size-dialog" host.ImageSize (fun () -> applyImageSize host)
            WordDialogs.textBoxStyleDialog "word-textbox-style-dialog-overlay" "word-textbox-style-dialog" host.TextBoxStyle (fun () -> applyTextBoxStyle host)
            WordDialogs.captionDialog "word-caption-dialog-overlay" "word-caption-dialog" host.Caption (fun () -> insertCaption host)
            WordDialogs.crossReferenceDialog "word-cross-reference-dialog-overlay" "word-cross-reference-dialog" host.CrossReference (fun () -> insertCrossReference host)
            WordDialogs.bookmarkDialog "word-bookmark-dialog-overlay" "word-bookmark-dialog" host.Bookmark (fun () -> applyBookmark host)
            WordDialogs.contentControlDialog "word-content-control-dialog-overlay" "word-content-control-dialog" host.ContentControl (fun () -> applyContentControl host)
            WordDialogs.fieldDialog "word-field-dialog-overlay" "word-field-dialog" host.Field (fun () -> insertField host)
            WordDialogs.equationDialog "word-equation-dialog-overlay" "word-equation-dialog" host.Equation (fun () -> insertEquation host)
            WordDialogs.revisionDialog "word-revision-dialog-overlay" "word-revision-dialog" host.Revision (fun () -> applyRevision host)
            WordDialogs.pageSetupDialog "word-page-setup-dialog-overlay" "word-page-setup-dialog" host.PageSetup (fun () -> applyPageSetup host)
            WordDialogs.headerFooterDialog "word-header-footer-dialog-overlay" "word-header-footer-dialog" host.HeaderFooter (fun () -> applyHeaderFooter host)
            WordDialogs.annotationDialog "word-annotation-dialog-overlay" "word-annotation-dialog" host.Annotation (fun () -> applyAnnotation host)
        ]

    let mutable private mountedHost: Host option = None

    let private defaultText = "FableUI Word host proof. The host owns the canvas and dialogs without toolbar chrome."

    let mount selector =
        let host = create defaultText "word-host"
        mountedHost <- Some host
        mountApp selector (view host)

    let openMountedParagraphStyle () =
        mountedHost |> Option.iter openParagraphStyle

    let openMountedFindReplace () =
        mountedHost |> Option.iter openFindReplace

    let insertMountedTable () =
        mountedHost |> Option.iter insertTable

    let openMountedInsertTable () =
        mountedHost |> Option.iter openInsertTable

    let openMountedImageSize () =
        mountedHost |> Option.iter openImageSize

    let openMountedTextBoxStyle () =
        mountedHost |> Option.iter openTextBoxStyle

    let openMountedCaption () =
        mountedHost |> Option.iter openCaption

    let openMountedCrossReference () =
        mountedHost |> Option.iter openCrossReference

    let openMountedBookmark () =
        mountedHost |> Option.iter openBookmark

    let openMountedContentControl () =
        mountedHost |> Option.iter openContentControl

    let openMountedField () =
        mountedHost |> Option.iter openField

    let openMountedEquation () =
        mountedHost |> Option.iter openEquation

    let openMountedRevision () =
        mountedHost |> Option.iter openRevision

    let openMountedPageSetup () =
        mountedHost |> Option.iter openPageSetup

    let openMountedHeaderFooter () =
        mountedHost |> Option.iter openHeaderFooter

    let openMountedTableCellStyle () =
        mountedHost |> Option.iter openTableCellStyle

    let openMountedTableMargins () =
        mountedHost |> Option.iter openTableMargins

    let openMountedAnnotation kind =
        mountedHost |> Option.iter (openAnnotation kind)

    let exportMountedDocument () =
        mountedHost |> Option.map exportDocument |> Option.defaultValue [||]

    let exportMountedDocumentPackage () =
        mountedHost |> Option.bind exportDocumentPackage |> Option.defaultValue null

    let exportMountedPageSetup () =
        mountedHost |> Option.bind exportPageSetup |> Option.defaultValue null

    let exportMountedHeaderFooter () =
        mountedHost |> Option.bind exportHeaderFooter |> Option.defaultValue null

    let exportMountedCommentAnchors () =
        mountedHost |> Option.map exportCommentAnchors |> Option.defaultValue [||]

    let exportMountedRevisionInfos () =
        mountedHost |> Option.map exportRevisionInfos |> Option.defaultValue [||]

    let printMountedPdf () =
        mountedHost |> Option.iter printPdf

#endif
