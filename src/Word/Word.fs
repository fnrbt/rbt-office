namespace RedBat.Word

#if FABLE_COMPILER

module API =
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types

    module Doc = Document
    module Rope = Rope
    module TextBuf = TextBuffer
    module Ren = Renderer
    module Inp = Input

    let private optionalText value = if isNull (box value) || value = "" then None else Some value

    /// Editor instance
    type Editor = {
        Handler: Inp.InputHandler
        mutable Doc: Doc.Document
    }

    let initializeCanvasWithChange (canvas: HTMLCanvasElement) (afterChange: unit -> unit) =
        let doc = Doc.empty
        let mutable editor = Unchecked.defaultof<Editor>

        let onChange newDoc =
            editor.Doc <- newDoc
            afterChange ()

        let handler = Inp.create canvas doc onChange

        editor <- {
            Handler = handler
            Doc = doc
        }

        editor

    /// Initialize editor on an existing canvas element.
    let initializeCanvas (canvas: HTMLCanvasElement) =
        initializeCanvasWithChange canvas ignore

    /// Initialize editor on canvas element id.
    let initialize (canvasId: string) =
        let canvas = Browser.Dom.document.getElementById(canvasId) :?> HTMLCanvasElement
        if isNull canvas then
            failwith (sprintf "Canvas element not found: %s" canvasId)

        initializeCanvas canvas

    /// Create a read-only viewer on an existing canvas element.
    let createViewer (canvas: HTMLCanvasElement) =
        let doc = Doc.empty
        let ctx = Ren.createContext canvas
        Ren.render ctx doc
        createObj [
            "canvas" ==> canvas
            "document" ==> doc
            "render" ==> fun (nextDoc: Doc.Document) -> Ren.render ctx nextDoc
        ]
    
    /// Get document text
    let getText (editor: Editor) =
        Doc.getText editor.Doc

    let getTrackChanges (editor: Editor) =
        editor.Handler.TrackChanges

    let setTrackChanges (editor: Editor) enabled =
        editor.Handler.TrackChanges <- enabled

    let setTrackChangesOptions (editor: Editor) author date =
        editor.Handler.TrackChangesAuthor <- if isNull (box author) || author = "" then "RedBat Word" else author
        editor.Handler.TrackChangesDate <- if isNull (box date) then "" else date

    let setTextMeasurer (editor: Editor) (measureText: Layout.TextMeasurer) =
        editor.Handler.RenderContext.TextMeasurer <- Some measureText
        Ren.render editor.Handler.RenderContext editor.Doc

    let setTextMeasurers (editor: Editor) (measureText: Layout.TextMeasurer) (measureRange: Layout.TextRangeMeasurer) =
        editor.Handler.RenderContext.TextMeasurer <- Some measureText
        editor.Handler.RenderContext.TextRangeMeasurer <- Some measureRange
        Ren.render editor.Handler.RenderContext editor.Doc

    let setTextRangeMeasurer (editor: Editor) (measureRange: Layout.TextRangeMeasurer) =
        editor.Handler.RenderContext.TextRangeMeasurer <- Some measureRange
        Ren.render editor.Handler.RenderContext editor.Doc

    let setTextPathPainter (editor: Editor) (paintText: string -> Paint.TextPaintStyle -> float -> float -> float -> bool) =
        editor.Handler.RenderContext.TextPathPainter <- Some paintText
        Ren.render editor.Handler.RenderContext editor.Doc

    let clearTextPathPainter (editor: Editor) =
        editor.Handler.RenderContext.TextPathPainter <- None
        Ren.render editor.Handler.RenderContext editor.Doc

    let clearTextMeasurer (editor: Editor) =
        editor.Handler.RenderContext.TextMeasurer <- None
        editor.Handler.RenderContext.TextRangeMeasurer <- None
        editor.Handler.RenderContext.TextPathPainter <- None
        Ren.render editor.Handler.RenderContext editor.Doc

    let getCachedLayouts (editor: Editor) =
        Ren.cachedLayouts editor.Handler.RenderContext editor.Doc

    let toggleTrackChanges (editor: Editor) =
        editor.Handler.TrackChanges <- not editor.Handler.TrackChanges
        editor.Handler.TrackChanges

    let getCommentAnchors (editor: Editor) =
        Ren.commentAnchors editor.Handler.RenderContext editor.Doc
        |> List.map (fun anchor ->
            createObj [
                "Id" ==> anchor.Id
                "Text" ==> anchor.Text
                "Author" ==> anchor.Author
                "Date" ==> anchor.Date
                "Excerpt" ==> anchor.Excerpt
                "BlockId" ==> anchor.BlockId
                "StartOffset" ==> anchor.StartOffset
                "EndOffset" ==> anchor.EndOffset
                "X" ==> anchor.X
                "Y" ==> anchor.Y
                "Width" ==> anchor.Width
                "Height" ==> anchor.Height
                "CalloutX" ==> anchor.CalloutX
                "CalloutY" ==> anchor.CalloutY
                "IsSelected" ==> anchor.IsSelected
            ])
        |> List.toArray

    let selectCommentAnchor (editor: Editor) (anchor: obj) =
        let blockId = anchor?BlockId |> unbox<string>
        let startOffset = anchor?StartOffset |> unbox<int>
        let endOffset = anchor?EndOffset |> unbox<int>
        let newDoc = Doc.setSelection { BlockId = blockId; Offset = startOffset } { BlockId = blockId; Offset = endOffset } editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getSelection (editor: Editor) =
        match editor.Doc.Selection with
        | None -> null
        | Some selection ->
            createObj [
                "AnchorBlockId" ==> selection.Anchor.BlockId
                "AnchorOffset" ==> selection.Anchor.Offset
                "FocusBlockId" ==> selection.Focus.BlockId
                "FocusOffset" ==> selection.Focus.Offset
                "IsCollapsed" ==> selection.IsCollapsed
            ]

    let clearSelection (editor: Editor) =
        let newDoc = { editor.Doc with Selection = None }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getImageHitTest (editor: Editor) (x: float) (y: float) =
        match Ren.hitTestForImage editor.Handler.RenderContext editor.Doc x y with
        | None -> null
        | Some (blockId, offset, width, height, imageX, imageY) ->
            createObj [
                "BlockId" ==> blockId
                "Offset" ==> offset
                "Width" ==> width
                "Height" ==> height
                "X" ==> imageX
                "Y" ==> imageY
            ]

    let getLayoutSegments (editor: Editor) =
        Ren.layoutDocument editor.Handler.RenderContext editor.Doc
        |> List.collect (fun layout ->
            layout.Lines
            |> List.collect (fun line ->
                let alignOffset = Layout.alignmentOffset layout line
                line.Segments
                |> List.map (function
                    | Layout.TextSegment (text, _, x, width, startOffset) ->
                        createObj [
                            "Kind" ==> "text"
                            "Text" ==> text
                            "BlockId" ==> layout.BlockId
                            "StartOffset" ==> startOffset
                            "X" ==> (x + alignOffset)
                            "Y" ==> line.Y
                            "Width" ==> width
                            "Height" ==> line.Height
                        ]
                    | Layout.ImageSegment (_, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, _, _, _, _) ->
                        createObj [
                            "Kind" ==> "image"
                            "BlockId" ==> layout.BlockId
                            "StartOffset" ==> startOffset
                            "X" ==> (x + alignOffset)
                            "Y" ==> (line.Y + yOffset)
                            "Width" ==> width
                            "Height" ==> height
                            "Floating" ==> floating
                            "Wrap" ==> wrap
                            "RelativeHeight" ==> relativeHeight
                        ])) )
        |> List.toArray

    /// Find literal text matches across editable paragraphs.
    let findText (editor: Editor) (query: string) (matchCase: bool) =
        Doc.findText query matchCase editor.Doc
        |> List.map (fun matchInfo ->
            createObj [
                "BlockId" ==> matchInfo.BlockId
                "Start" ==> matchInfo.Start
                "End" ==> matchInfo.End
                "Text" ==> matchInfo.Text
            ])
        |> List.toArray

    let getRevisionInfos (editor: Editor) =
        Doc.revisionInfos editor.Doc
        |> List.map (fun revision ->
            createObj [
                "BlockId" ==> revision.BlockId
                "Start" ==> revision.Start
                "End" ==> revision.End
                "RevisionType" ==> revision.RevisionType
                "Author" ==> (revision.Author |> Option.defaultValue "")
                "Date" ==> (revision.Date |> Option.defaultValue "")
                "Text" ==> revision.Text
            ])
        |> List.toArray

    /// Select a previously returned find match.
    let selectFindMatch (editor: Editor) (matchObj: obj) =
        let matchInfo: Doc.FindMatch = {
            BlockId = matchObj?BlockId |> unbox<string>
            Start = matchObj?Start |> unbox<int>
            End = matchObj?End |> unbox<int>
            Text = matchObj?Text |> unbox<string>
        }
        let newDoc = Doc.selectFindMatch matchInfo editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let selectRevisionInfo (editor: Editor) (revisionObj: obj) =
        let blockId = revisionObj?BlockId |> unbox<string>
        let startOffset = revisionObj?Start |> unbox<int>
        let endOffset = revisionObj?End |> unbox<int>
        let newDoc = Doc.setSelection { BlockId = blockId; Offset = startOffset } { BlockId = blockId; Offset = endOffset } editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Replace all literal text matches across editable paragraphs.
    let replaceAllText (editor: Editor) (query: string) (replacement: string) (matchCase: bool) =
        let (newDoc, count) = Doc.replaceAllText query replacement matchCase editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
        count
    
    /// Insert text at cursor
    let insertTextAt (editor: Editor) (text: string) =
        let newDoc =
            if editor.Handler.TrackChanges then Doc.insertTrackedText editor.Handler.TrackChangesAuthor editor.Handler.TrackChangesDate text editor.Doc
            else Doc.insertText text editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertFieldAt (editor: Editor) (fieldCode: string) (result: string) =
        let newDoc = Doc.insertField fieldCode result editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let private pageNumberForY (doc: Doc.Document) y =
        let stride = doc.PageSetup.PageHeight + WordLayout.Page.gap
        int (floor (max 0.0 y / stride)) + 1

    let private displayPageNumberForPhysicalPage (restartPoints: (int * int) list) physicalPage =
        restartPoints
        |> List.filter (fun (restartPhysicalPage, _) -> restartPhysicalPage <= physicalPage)
        |> List.tryLast
        |> Option.map (fun (restartPhysicalPage, restartDisplayPage) -> restartDisplayPage + physicalPage - restartPhysicalPage)
        |> Option.defaultValue physicalPage

    let private pageMapFromLayout (editor: Editor) =
        let layouts = Ren.layoutDocument editor.Handler.RenderContext editor.Doc
        let restartPoints =
            layouts
            |> List.choose (fun layout ->
                Doc.findParagraph editor.Doc layout.BlockId
                |> Option.bind (fun para ->
                    para.Metadata.SectionPageNumberStart
                    |> Option.map (fun start -> pageNumberForY editor.Doc layout.Y, max 1 start)))
            |> List.sortBy fst

        let displayPageNumberForY y =
            y
            |> pageNumberForY editor.Doc
            |> displayPageNumberForPhysicalPage restartPoints

        let fieldAnchors =
            layouts
            |> List.collect (fun layout ->
                layout.Lines
                |> List.collect (fun line ->
                    line.Segments
                    |> List.choose (function
                        | Layout.TextSegment (_, style, _, _, startOffset) ->
                            style.FieldCode
                            |> Option.map (fun code ->
                                ({
                                    PageFieldBlockId = layout.BlockId
                                    PageFieldStartOffset = startOffset
                                    PageFieldCode = code
                                    PageNumber = displayPageNumberForY line.Y
                                }: Doc.PageFieldAnchor))
                        | Layout.ImageSegment _ -> None)))
        let bookmarkAnchors =
            layouts
            |> List.collect (fun layout ->
                layout.Lines
                |> List.collect (fun line ->
                    line.Segments
                    |> List.choose (function
                        | Layout.TextSegment (_, style, _, _, startOffset) ->
                            style.BookmarkName
                            |> Option.map (fun name ->
                                ({
                                    PageBookmarkName = name
                                    PageBookmarkBlockId = layout.BlockId
                                    PageBookmarkStartOffset = startOffset
                                    PageNumber = displayPageNumberForY line.Y
                                }: Doc.PageBookmarkAnchor))
                        | Layout.ImageSegment _ -> None)))
            |> List.distinctBy (fun anchor -> anchor.PageBookmarkName)
        let blockAnchors =
            layouts
            |> List.map (fun layout ->
                ({
                    PageBlockId = layout.BlockId
                    PageNumber = displayPageNumberForY layout.Y
                }: Doc.PageBlockAnchor))
        let pageCount =
            layouts
            |> List.map (fun layout -> pageNumberForY editor.Doc (layout.Y + layout.Height))
            |> List.fold max 1
        let pageMap: Doc.PageMap = {
            PageCount = pageCount
            BlockAnchors = blockAnchors
            Fields = fieldAnchors
            Bookmarks = bookmarkAnchors
        }
        pageMap

    let refreshFields (editor: Editor) =
        let newDoc = Doc.refreshFieldsWithPageMap (pageMapFromLayout editor) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertEquation (editor: Editor) (equationText: string) =
        let newDoc = Doc.insertEquation equationText editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Set document content
    let setContent (editor: Editor) (text: string) =
        let doc: Doc.Document = {
            Blocks = [
                Doc.Paragraph {
                    Id = "block-0"
                    Buffer = TextBuf.ofString text
                    Metadata = Doc.defaultMetadata
                }
            ]
            Selection = Some {
                Anchor = { BlockId = "block-0"; Offset = text.Length }
                Focus = { BlockId = "block-0"; Offset = text.Length }
            }
            PageSetup = Doc.defaultPageSetup
            Styles = []
            Headers = []
            Footers = []
            HeaderParts = []
            FooterParts = []
        }
        Inp.updateDocument editor.Handler doc
        editor.Doc <- doc
    
    /// Toggle bold on selection
    let toggleBold (editor: Editor) =
        let newDoc = Doc.toggleBold editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Toggle italic on selection
    let toggleItalic (editor: Editor) =
        let newDoc = Doc.toggleItalic editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Toggle underline on selection
    let toggleUnderline (editor: Editor) =
        let newDoc = Doc.toggleUnderline editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Toggle all caps on selection
    let toggleAllCaps (editor: Editor) =
        let newDoc = Doc.toggleAllCaps editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Toggle small caps on selection
    let toggleSmallCaps (editor: Editor) =
        let newDoc = Doc.toggleSmallCaps editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set numbered-list marker format on current block
    let setNumberFormat (editor: Editor) (format: string) =
        let newDoc = Doc.setNumberFormat format editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Toggle strikethrough on selection
    let toggleStrike (editor: Editor) =
        let newDoc = Doc.toggleStrike editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Clear direct character formatting on selection
    let clearFormatting (editor: Editor) =
        let newDoc = Doc.clearFormatting editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set vertical alignment on selection
    let setVertAlignFor (editor: Editor) (align: string) =
        let newDoc = Doc.setVertAlign (if isNull (box align) || align = "" || align = "baseline" then None else Some align) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set text color on selection
    let setColorFor (editor: Editor) (color: string) =
        let newDoc = Doc.setColor (if isNull (box color) || color = "" then None else Some color) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set highlight color on selection
    let setHighlightFor (editor: Editor) (color: string) =
        let newDoc = Doc.setHighlight (if isNull (box color) || color = "" || color = "none" then None else Some color) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set hyperlink on selection
    let setHyperlinkFor (editor: Editor) (target: string) =
        let newDoc = Doc.setHyperlink (if isNull (box target) || target = "" then None else Some target) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set comment metadata on selection
    let setCommentFor (editor: Editor) (commentId: string) (commentText: string) (author: string) (date: string) =
        let commentIdOpt = if isNull (box commentId) || commentId = "" then None else Some commentId
        let commentTextOpt = if isNull (box commentText) || commentText = "" then None else Some commentText
        let authorOpt = if isNull (box author) || author = "" then None else Some author
        let dateOpt = if isNull (box date) || date = "" then None else Some date
        let newDoc = Doc.setComment commentIdOpt commentTextOpt authorOpt dateOpt editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Set footnote/endnote metadata on selection
    let setNoteFor (editor: Editor) (noteId: string) (noteType: string) (noteText: string) =
        let noteIdOpt = if isNull (box noteId) || noteId = "" then None else Some noteId
        let noteTypeOpt = if isNull (box noteType) || noteType = "" then None else Some noteType
        let noteTextOpt = if isNull (box noteText) || noteText = "" then None else Some noteText
        let newDoc = Doc.setNote noteIdOpt noteTypeOpt noteTextOpt editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setBookmarkFor (editor: Editor) (bookmarkName: string) =
        let newDoc = Doc.setBookmark bookmarkName editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setContentControlFor (editor: Editor) (controlId: string) (tag: string) (alias: string) =
        let newDoc = Doc.setContentControl controlId tag alias editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getBookmarkInfos (editor: Editor) =
        Doc.bookmarkInfos editor.Doc
        |> List.map (fun info ->
            createObj [
                "Id" ==> info.Id
                "Name" ==> info.Name
                "BlockId" ==> info.BlockId
                "StartOffset" ==> info.StartOffset
                "EndOffset" ==> info.EndOffset
            ])
        |> List.toArray

    let selectBookmark (editor: Editor) (bookmarkId: string) =
        let newDoc = Doc.selectBookmark bookmarkId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let removeBookmark (editor: Editor) (bookmarkId: string) =
        let newDoc = Doc.removeBookmark bookmarkId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let renameBookmark (editor: Editor) (bookmarkId: string) (bookmarkName: string) =
        let newDoc = Doc.renameBookmark bookmarkId bookmarkName editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getContentControlInfos (editor: Editor) =
        Doc.contentControlInfos editor.Doc
        |> List.map (fun info ->
            createObj [
                "Id" ==> info.Id
                "Tag" ==> info.Tag
                "Alias" ==> info.Alias
                "BlockId" ==> info.BlockId
            ])
        |> List.toArray

    let selectContentControl (editor: Editor) (controlId: string) =
        let newDoc = Doc.selectContentControl controlId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let removeContentControl (editor: Editor) (controlId: string) =
        let newDoc = Doc.removeContentControl controlId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let updateContentControl (editor: Editor) (controlId: string) (newControlId: string) (tag: string) (alias: string) =
        let newDoc = Doc.updateContentControl controlId newControlId tag alias editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setParagraphStyleFor (editor: Editor) (styleId: string) =
        let newDoc = Doc.setParagraphStyle styleId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setCharacterStyleFor (editor: Editor) (styleId: string) =
        let newDoc = Doc.setCharacterStyle styleId editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getStyleCatalog (editor: Editor) =
        Docx.exportStylesToEditor editor.Doc.Styles

    let setRevisionFor (editor: Editor) (revisionType: string) (revisionId: string) (author: string) (date: string) =
        let newDoc = Doc.setRevision revisionType revisionId author date editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let acceptSelectedRevision (editor: Editor) =
        let newDoc = Doc.acceptSelectedRevision editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let rejectSelectedRevision (editor: Editor) =
        let newDoc = Doc.rejectSelectedRevision editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let acceptAllRevisions (editor: Editor) =
        let newDoc = Doc.acceptAllRevisions editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let rejectAllRevisions (editor: Editor) =
        let newDoc = Doc.rejectAllRevisions editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let selectNextRevision (editor: Editor) =
        let newDoc = Doc.selectNextRevision editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let selectPreviousRevision (editor: Editor) =
        let newDoc = Doc.selectPreviousRevision editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Set font size on selection
    let setFontSizeFor (editor: Editor) (size: float) =
        let newDoc = Doc.setFontSize size editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Set font family on selection
    let setFontFamilyFor (editor: Editor) (family: string) =
        let newDoc = Doc.setFontFamily family editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Apply Heading 1-3 to the current block
    let setHeadingLevel (editor: Editor) (level: int) =
        let newDoc = Doc.setHeadingLevel level editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Convert the current heading block back to a normal paragraph
    let clearHeading (editor: Editor) =
        let newDoc = Doc.clearHeading editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
    
    /// Select all text
    let selectAll (editor: Editor) =
        let newDoc = Doc.selectAll editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Align left
    let alignLeft (editor: Editor) =
        let newDoc = Doc.alignLeft editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Align center
    let alignCenter (editor: Editor) =
        let newDoc = Doc.alignCenter editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Align right
    let alignRight (editor: Editor) =
        let newDoc = Doc.alignRight editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Insert image at cursor
    let insertImage (editor: Editor) (src: string) (width: float) (height: float) (alt: string) =
        let widthOpt = if width > 0.0 then Some width else None
        let heightOpt = if height > 0.0 then Some height else None
        let altOpt = if isNull (box alt) || alt = "" then None else Some alt
        let newDoc = Doc.insertImageAtCursor src widthOpt heightOpt altOpt editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let resizeCurrentImage (editor: Editor) (width: float) (height: float) =
        let newDoc = Doc.resizeImageNearCursor width height editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Insert a table row after the row containing the current selection
    let insertTableRowAfter (editor: Editor) =
        let newDoc = Doc.insertTableRowAfterSelection editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Insert a table column after the column containing the current selection
    let insertTableColumnAfter (editor: Editor) =
        let newDoc = Doc.insertTableColumnAfterSelection editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Delete the table row containing the current selection
    let deleteTableRow (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableRow editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Delete the table column containing the current selection
    let deleteTableColumn (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableColumn editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertTable (editor: Editor) (rows: int) (columns: int) =
        let newDoc = Doc.insertTable rows columns editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertTextBox (editor: Editor) (text: string) =
        let newDoc = Doc.insertTextBox text editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertCaption (editor: Editor) (label: string) (captionText: string) =
        let newDoc = Doc.insertCaption label captionText editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let insertCrossReferenceToLatestCaption (editor: Editor) (label: string) =
        let beforeDoc = editor.Doc
        let newDoc = Doc.insertCrossReferenceToLatestCaption label editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc
        newDoc <> beforeDoc

    let setCurrentTextBoxStyle (editor: Editor) (x: float) (y: float) (width: float) (height: float) (fillColor: string) (borderColor: string) (borderWidth: float) =
        let newDoc = Doc.setCurrentTextBoxStyle x y width height fillColor borderColor borderWidth editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setSelectedTableCellStyle (editor: Editor) (backgroundColor: string) (cellWidth: float) (borderColor: string) (borderWidth: float) (verticalAlignment: string) (paddingTop: float) (paddingRight: float) (paddingBottom: float) (paddingLeft: float) =
        let optionalCellWidth = if cellWidth > 0.0 then Some cellWidth else None
        let optionalWidth = if borderWidth > 0.0 then Some borderWidth else None
        let optionalPadding value = if value >= 0.0 then Some value else None
        let newDoc = Doc.setSelectedTableCellStyle (optionalText backgroundColor) optionalCellWidth (optionalText borderColor) optionalWidth (optionalText verticalAlignment) (optionalPadding paddingTop) (optionalPadding paddingRight) (optionalPadding paddingBottom) (optionalPadding paddingLeft) editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Align justify
    let alignJustify (editor: Editor) =
        let newDoc = Doc.alignJustify editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let toggleBulletList (editor: Editor) =
        let newDoc = Doc.toggleBulletList editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let toggleNumberedList (editor: Editor) =
        let newDoc = Doc.toggleNumberedList editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let increaseIndent (editor: Editor) =
        let newDoc = Doc.increaseIndent editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let decreaseIndent (editor: Editor) =
        let newDoc = Doc.decreaseIndent editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let applyParagraphStyle (editor: Editor) (alignment: string) (indentLeft: float) (indentRight: float) (spacingBefore: float) (spacingAfter: float) (lineSpacing: float) (fontSize: float) (firstLineIndent: float option) (backgroundColor: string) =
        let alignmentValue =
            match alignment with
            | "center" -> Doc.Center
            | "right" -> Doc.Right
            | "justify" -> Doc.Justify
            | _ -> Doc.Left
        let newDoc =
            editor.Doc
            |> Doc.setCurrentParagraphStyle alignmentValue (Some indentLeft) (Some indentRight) (Some spacingBefore) (Some spacingAfter) (Some lineSpacing) firstLineIndent (optionalText backgroundColor)
            |> Doc.selectAll
            |> Doc.setFontSize fontSize
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setSectionBreak (editor: Editor) (kind: string) =
        let sectionBreak =
            match kind with
            | "continuous" -> Some "continuous"
            | "none" -> None
            | _ -> Some "nextPage"
        let newDoc = Doc.setCurrentSectionBreak sectionBreak editor.Doc
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Import document from EditorParagraphs (Office.Formats format)
    /// Use with: Office.Formats processors.forEditor() output
    let importFromDocx (editor: Editor) (editorParagraphs: obj array) =
        let newDoc = Docx.importFromEditor editorParagraphs
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let importFromDocxWithHeaderFooters (editor: Editor) (editorParagraphs: obj array) (headerParagraphs: obj array) (footerParagraphs: obj array) =
        let newDoc = Docx.importFromEditorWithHeaderFooters editorParagraphs headerParagraphs footerParagraphs
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let importFromDocxWithHeaderFooterParts (editor: Editor) (editorParagraphs: obj array) (headerParts: obj array) (footerParts: obj array) =
        let newDoc = Docx.importFromEditorWithHeaderFooterParts editorParagraphs headerParts footerParts
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let importFromDocxWithPageSetupAndHeaderFooterParts (editor: Editor) (editorParagraphs: obj array) (pageSetup: obj) (headerParts: obj array) (footerParts: obj array) =
        let newDoc = Docx.importFromEditorWithPageSetupAndHeaderFooterParts editorParagraphs pageSetup headerParts footerParts
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let importFromDocxPackage (editor: Editor) (package: obj) =
        let newDoc = Docx.importFromEditorPackage package
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setPageSetup (editor: Editor) (pageWidth: float) (pageHeight: float) (marginLeft: float) (marginRight: float) (marginTop: float) (marginBottom: float) (headerDistance: float) (footerDistance: float) =
        let setup: Doc.PageSetup = {
            PageWidth = max 120.0 pageWidth
            PageHeight = max 120.0 pageHeight
            MarginLeft = max 0.0 marginLeft
            MarginRight = max 0.0 marginRight
            MarginTop = max 0.0 marginTop
            MarginBottom = max 0.0 marginBottom
            HeaderDistance = max 0.0 headerDistance
            FooterDistance = max 0.0 footerDistance
            DifferentFirstPage = editor.Doc.PageSetup.DifferentFirstPage
            EvenAndOddHeaders = editor.Doc.PageSetup.EvenAndOddHeaders
        }
        let newDoc = { editor.Doc with PageSetup = setup }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setPageSetupOptions (editor: Editor) (pageWidth: float) (pageHeight: float) (marginLeft: float) (marginRight: float) (marginTop: float) (marginBottom: float) (headerDistance: float) (footerDistance: float) (differentFirstPage: bool) (evenAndOddHeaders: bool) =
        let setup: Doc.PageSetup = {
            PageWidth = max 120.0 pageWidth
            PageHeight = max 120.0 pageHeight
            MarginLeft = max 0.0 marginLeft
            MarginRight = max 0.0 marginRight
            MarginTop = max 0.0 marginTop
            MarginBottom = max 0.0 marginBottom
            HeaderDistance = max 0.0 headerDistance
            FooterDistance = max 0.0 footerDistance
            DifferentFirstPage = differentFirstPage
            EvenAndOddHeaders = evenAndOddHeaders
        }
        let newDoc = { editor.Doc with PageSetup = setup }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getPageSetup (editor: Editor) =
        let setup = editor.Doc.PageSetup
        createObj [
            "PageWidth" ==> setup.PageWidth
            "PageHeight" ==> setup.PageHeight
            "MarginLeft" ==> setup.MarginLeft
            "MarginRight" ==> setup.MarginRight
            "MarginTop" ==> setup.MarginTop
            "MarginBottom" ==> setup.MarginBottom
            "HeaderDistance" ==> setup.HeaderDistance
            "FooterDistance" ==> setup.FooterDistance
            "DifferentFirstPage" ==> setup.DifferentFirstPage
            "EvenAndOddHeaders" ==> setup.EvenAndOddHeaders
        ]

    let private paragraphFromText id text : Doc.ParagraphBlock =
        { Id = id; Buffer = TextBuf.ofString text; Metadata = Doc.defaultMetadata }

    let private textFromParagraph (para: Doc.ParagraphBlock) = Rope.toString para.Buffer.Rope

    let private textFromPart kind (parts: Doc.HeaderFooterPart list) =
        parts
        |> List.tryFind (fun part -> part.Kind = kind)
        |> Option.bind (fun part ->
            part.Blocks
            |> List.tryPick (function
                | Doc.Paragraph para
                | Doc.Heading (_, para) -> Some (textFromParagraph para)
                | Doc.Table _
                | Doc.TextBox _ -> None))
        |> Option.defaultValue ""

    let private upsertPart prefix kind text parts =
        let remaining = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> part.Kind <> kind || part.RelationshipId.IsSome)
        if System.String.IsNullOrWhiteSpace text then remaining
        else
            { Kind = kind; RelationshipId = None; Blocks = [Doc.Paragraph (paragraphFromText (sprintf "%s-%s" prefix kind) text)] } :: remaining

    let private upsertReferencedPart prefix relationshipId text parts =
        let remaining = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> part.RelationshipId <> Some relationshipId)
        if System.String.IsNullOrWhiteSpace text then remaining
        else
            { Kind = "default"; RelationshipId = Some relationshipId; Blocks = [Doc.Paragraph (paragraphFromText (sprintf "%s-%s" prefix relationshipId) text)] } :: remaining

    let private tableCell id text : Doc.TableCell =
        { Content = [paragraphFromText id text]
          ColSpan = 1
          BackgroundColor = None
          Width = None
          PreferredWidth = None
          BorderColor = None
          BorderWidth = None
          Borders = None
          VerticalAlignment = None
          VerticalMerge = None
          PaddingTop = None
          PaddingRight = None
          PaddingBottom = None
          PaddingLeft = None }

    let private twoCellTable id left right : Doc.TableBlock =
        { Id = id
          Rows = [{ Cells = [tableCell (id + "-left") left; tableCell (id + "-right") right]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None }]
          PreferredWidth = Some { Kind = Doc.PercentWidth; Value = Some 100.0 }
          LayoutMode = Some Doc.FixedTableLayout
          GridColumns = [{ Kind = Doc.PercentWidth; Value = Some 50.0 }; { Kind = Doc.PercentWidth; Value = Some 50.0 }]
          Borders = None
          BackgroundColor = None
          CellSpacing = None
          PaddingTop = None
          PaddingRight = None
          PaddingBottom = None
          PaddingLeft = None
          Floating = false }

    let private imageParagraph id src alt width height : Doc.ParagraphBlock =
        { Id = id
          Buffer = TextBuf.empty |> TextBuf.insertImage 0 src (Some (max 1.0 width)) (Some (max 1.0 height)) (optionalText alt)
          Metadata = Doc.defaultMetadata }

    let private isTableId id = function
        | Doc.Table table -> table.Id = id
        | _ -> false

    let private upsertDefaultTable tableId left right parts =
        let hasTable = not (System.String.IsNullOrWhiteSpace left) || not (System.String.IsNullOrWhiteSpace right)
        let defaultPart = parts |> List.tryFind (fun (part: Doc.HeaderFooterPart) -> part.Kind = "default" && part.RelationshipId.IsNone)
        let otherParts = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> not (part.Kind = "default" && part.RelationshipId.IsNone))
        let existingBlocks = defaultPart |> Option.map _.Blocks |> Option.defaultValue [] |> List.filter (isTableId tableId >> not)
        let blocks = if hasTable then existingBlocks @ [Doc.Table (twoCellTable tableId left right)] else existingBlocks
        if blocks.IsEmpty then otherParts else { Kind = "default"; RelationshipId = None; Blocks = blocks } :: otherParts

    let private isImageParagraphId id = function
        | Doc.Paragraph para -> para.Id = id
        | _ -> false

    let private upsertDefaultImage paragraphId src alt width height parts =
        let hasImage = not (System.String.IsNullOrWhiteSpace src)
        let defaultPart = parts |> List.tryFind (fun (part: Doc.HeaderFooterPart) -> part.Kind = "default" && part.RelationshipId.IsNone)
        let otherParts = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> not (part.Kind = "default" && part.RelationshipId.IsNone))
        let existingBlocks = defaultPart |> Option.map _.Blocks |> Option.defaultValue [] |> List.filter (isImageParagraphId paragraphId >> not)
        let blocks = if hasImage then existingBlocks @ [Doc.Paragraph (imageParagraph paragraphId src alt width height)] else existingBlocks
        if blocks.IsEmpty then otherParts else { Kind = "default"; RelationshipId = None; Blocks = blocks } :: otherParts

    let private tableText tableId cellIndex parts =
        parts
        |> List.tryPick (fun (part: Doc.HeaderFooterPart) ->
            part.Blocks
            |> List.tryPick (function
                | Doc.Table table when table.Id = tableId ->
                    table.Rows
                    |> List.tryHead
                    |> Option.bind (fun row -> row.Cells |> List.tryItem cellIndex)
                    |> Option.bind (fun cell -> cell.Content |> List.tryHead)
                    |> Option.map textFromParagraph
                | _ -> None))
        |> Option.defaultValue ""

    let private imageValue paragraphId value parts =
        parts
        |> List.tryPick (fun (part: Doc.HeaderFooterPart) ->
            part.Blocks
            |> List.tryPick (function
                | Doc.Paragraph para when para.Id = paragraphId ->
                    para.Buffer.Runs
                    |> List.tryPick (function
                        | TextBuf.Image image ->
                            match value with
                            | "alt" -> image.Alt
                            | "width" -> image.Width |> Option.map string
                            | "height" -> image.Height |> Option.map string
                            | _ -> Some image.Src
                        | TextBuf.Text _ -> None)
                | _ -> None))
        |> Option.defaultValue ""

    let setHeaderFooterText (editor: Editor) defaultHeader defaultFooter firstHeader firstFooter evenHeader evenFooter =
        let headerParts =
            editor.Doc.HeaderParts
            |> upsertPart "header" "default" defaultHeader
            |> upsertPart "header" "first" firstHeader
            |> upsertPart "header" "even" evenHeader
        let footerParts =
            editor.Doc.FooterParts
            |> upsertPart "footer" "default" defaultFooter
            |> upsertPart "footer" "first" firstFooter
            |> upsertPart "footer" "even" evenFooter
        let newDoc = {
            editor.Doc with
                HeaderParts = headerParts
                FooterParts = footerParts
                Headers = if System.String.IsNullOrWhiteSpace defaultHeader then [] else [paragraphFromText "header-default" defaultHeader]
                Footers = if System.String.IsNullOrWhiteSpace defaultFooter then [] else [paragraphFromText "footer-default" defaultFooter]
        }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let getHeaderFooterText (editor: Editor) =
        createObj [
            "DefaultHeader" ==> textFromPart "default" editor.Doc.HeaderParts
            "DefaultFooter" ==> textFromPart "default" editor.Doc.FooterParts
            "DefaultHeaderTableLeft" ==> tableText "default-header-table" 0 editor.Doc.HeaderParts
            "DefaultHeaderTableRight" ==> tableText "default-header-table" 1 editor.Doc.HeaderParts
            "DefaultFooterTableLeft" ==> tableText "default-footer-table" 0 editor.Doc.FooterParts
            "DefaultFooterTableRight" ==> tableText "default-footer-table" 1 editor.Doc.FooterParts
            "DefaultHeaderImageSrc" ==> imageValue "default-header-image" "src" editor.Doc.HeaderParts
            "DefaultHeaderImageAlt" ==> imageValue "default-header-image" "alt" editor.Doc.HeaderParts
            "DefaultHeaderImageWidth" ==> imageValue "default-header-image" "width" editor.Doc.HeaderParts
            "DefaultHeaderImageHeight" ==> imageValue "default-header-image" "height" editor.Doc.HeaderParts
            "DefaultFooterImageSrc" ==> imageValue "default-footer-image" "src" editor.Doc.FooterParts
            "DefaultFooterImageAlt" ==> imageValue "default-footer-image" "alt" editor.Doc.FooterParts
            "DefaultFooterImageWidth" ==> imageValue "default-footer-image" "width" editor.Doc.FooterParts
            "DefaultFooterImageHeight" ==> imageValue "default-footer-image" "height" editor.Doc.FooterParts
            "FirstHeader" ==> textFromPart "first" editor.Doc.HeaderParts
            "FirstFooter" ==> textFromPart "first" editor.Doc.FooterParts
            "EvenHeader" ==> textFromPart "even" editor.Doc.HeaderParts
            "EvenFooter" ==> textFromPart "even" editor.Doc.FooterParts
        ]

    let setHeaderFooterTableText (editor: Editor) defaultHeaderLeft defaultHeaderRight defaultFooterLeft defaultFooterRight =
        let newDoc = {
            editor.Doc with
                HeaderParts = editor.Doc.HeaderParts |> upsertDefaultTable "default-header-table" defaultHeaderLeft defaultHeaderRight
                FooterParts = editor.Doc.FooterParts |> upsertDefaultTable "default-footer-table" defaultFooterLeft defaultFooterRight
        }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let private parseFloat fallback (value: string) =
        match System.Double.TryParse(value: string) with
        | true, parsed -> parsed
        | false, _ -> fallback

    let setHeaderFooterImageText (editor: Editor) defaultHeaderSrc defaultHeaderAlt defaultHeaderWidth defaultHeaderHeight defaultFooterSrc defaultFooterAlt defaultFooterWidth defaultFooterHeight =
        let newDoc = {
            editor.Doc with
                HeaderParts = editor.Doc.HeaderParts |> upsertDefaultImage "default-header-image" defaultHeaderSrc defaultHeaderAlt (parseFloat 96.0 defaultHeaderWidth) (parseFloat 48.0 defaultHeaderHeight)
                FooterParts = editor.Doc.FooterParts |> upsertDefaultImage "default-footer-image" defaultFooterSrc defaultFooterAlt (parseFloat 96.0 defaultFooterWidth) (parseFloat 48.0 defaultFooterHeight)
        }
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    let setCurrentSectionHeaderFooterText (editor: Editor) headerText footerText =
        let headerRelationshipId = "section-header-current"
        let footerRelationshipId = "section-footer-current"
        let headerParts = upsertReferencedPart "header" headerRelationshipId headerText editor.Doc.HeaderParts
        let footerParts = upsertReferencedPart "footer" footerRelationshipId footerText editor.Doc.FooterParts
        let headerRefs = if System.String.IsNullOrWhiteSpace headerText then [] else ["default", headerRelationshipId]
        let footerRefs = if System.String.IsNullOrWhiteSpace footerText then [] else ["default", footerRelationshipId]
        let newDoc =
            { editor.Doc with HeaderParts = headerParts; FooterParts = footerParts }
            |> Doc.setCurrentHeaderFooterReferences headerRefs footerRefs
        Inp.updateDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Export document to EditorParagraphs (Office.Formats format)
    /// Use with: Office.Formats WordWriter.writeFromEditor()
    let exportToDocx (editor: Editor) : obj array =
        Docx.exportToEditor editor.Doc

    let exportDocumentPackage (editor: Editor) =
        createObj [
            "Paragraphs" ==> Docx.exportToEditor editor.Doc
            "PageSetup" ==> getPageSetup editor
            "Styles" ==> Docx.exportStylesToEditor editor.Doc.Styles
            "HeaderParts" ==> Docx.exportHeaderFooterPartsToEditor editor.Doc.HeaderParts
            "FooterParts" ==> Docx.exportHeaderFooterPartsToEditor editor.Doc.FooterParts
        ]

    /// Test rope operations
    let testRope () =
        // Test basic rope operations
        let rope = Rope.ofString "Hello, World!"
        let len = Rope.length rope
        let text = Rope.toString rope

        // Test insert
        let rope2 = Rope.insert 7 "Beautiful " rope
        let text2 = Rope.toString rope2

        // Test delete
        let rope3 = Rope.delete 0 7 rope2
        let text3 = Rope.toString rope3

        // Test slice
        let slice = Rope.slice 0 9 rope3

        createObj [
            "original" ==> text
            "afterInsert" ==> text2
            "afterDelete" ==> text3
            "slice" ==> slice
            "length" ==> len
        ]
    
    /// Export API
    [<ExportDefault>]
    let exports =
        createObj [
            "initialize" ==> initialize
            "createViewer" ==> createViewer
            "getText" ==> getText
            "getTrackChanges" ==> getTrackChanges
            "setTrackChanges" ==> setTrackChanges
            "setTrackChangesOptions" ==> setTrackChangesOptions
            "setTextMeasurer" ==> setTextMeasurer
            "setTextMeasurers" ==> setTextMeasurers
            "setTextRangeMeasurer" ==> setTextRangeMeasurer
            "setTextPathPainter" ==> setTextPathPainter
            "clearTextPathPainter" ==> clearTextPathPainter
            "clearTextMeasurer" ==> clearTextMeasurer
            "toggleTrackChanges" ==> toggleTrackChanges
            "getCommentAnchors" ==> getCommentAnchors
            "getRevisionInfos" ==> getRevisionInfos
            "selectCommentAnchor" ==> selectCommentAnchor
            "getSelection" ==> getSelection
            "clearSelection" ==> clearSelection
            "getImageHitTest" ==> getImageHitTest
            "getLayoutSegments" ==> getLayoutSegments
            "findText" ==> findText
            "selectFindMatch" ==> selectFindMatch
            "selectRevisionInfo" ==> selectRevisionInfo
            "replaceAllText" ==> replaceAllText
            "insertText" ==> insertTextAt
            "setContent" ==> setContent
            "testRope" ==> testRope
            "toggleBold" ==> toggleBold
            "toggleItalic" ==> toggleItalic
            "toggleUnderline" ==> toggleUnderline
            "toggleAllCaps" ==> toggleAllCaps
            "toggleSmallCaps" ==> toggleSmallCaps
            "setNumberFormat" ==> setNumberFormat
            "toggleStrike" ==> toggleStrike
            "clearFormatting" ==> clearFormatting
            "setVertAlign" ==> setVertAlignFor
            "setColor" ==> setColorFor
            "setHighlight" ==> setHighlightFor
            "setHyperlink" ==> setHyperlinkFor
            "setComment" ==> setCommentFor
            "setNote" ==> setNoteFor
            "refreshFields" ==> refreshFields
            "setBookmark" ==> setBookmarkFor
            "setContentControl" ==> setContentControlFor
            "getBookmarkInfos" ==> getBookmarkInfos
            "selectBookmark" ==> selectBookmark
            "removeBookmark" ==> removeBookmark
            "renameBookmark" ==> renameBookmark
            "getContentControlInfos" ==> getContentControlInfos
            "selectContentControl" ==> selectContentControl
            "removeContentControl" ==> removeContentControl
            "updateContentControl" ==> updateContentControl
            "setParagraphStyle" ==> setParagraphStyleFor
            "setCharacterStyle" ==> setCharacterStyleFor
            "getStyleCatalog" ==> getStyleCatalog
            "setRevision" ==> setRevisionFor
            "acceptSelectedRevision" ==> acceptSelectedRevision
            "rejectSelectedRevision" ==> rejectSelectedRevision
            "acceptAllRevisions" ==> acceptAllRevisions
            "rejectAllRevisions" ==> rejectAllRevisions
            "selectNextRevision" ==> selectNextRevision
            "selectPreviousRevision" ==> selectPreviousRevision
            "setFontSize" ==> setFontSizeFor
            "setFontFamily" ==> setFontFamilyFor
            "setHeadingLevel" ==> setHeadingLevel
            "clearHeading" ==> clearHeading
            "selectAll" ==> selectAll
            "alignLeft" ==> alignLeft
            "alignCenter" ==> alignCenter
            "alignRight" ==> alignRight
            "alignJustify" ==> alignJustify
            "toggleBulletList" ==> toggleBulletList
            "toggleNumberedList" ==> toggleNumberedList
            "increaseIndent" ==> increaseIndent
            "decreaseIndent" ==> decreaseIndent
            "insertImage" ==> insertImage
            "resizeCurrentImage" ==> resizeCurrentImage
            "insertTableRowAfter" ==> insertTableRowAfter
            "insertTableColumnAfter" ==> insertTableColumnAfter
            "deleteTableRow" ==> deleteTableRow
            "deleteTableColumn" ==> deleteTableColumn
            "insertTable" ==> insertTable
            "setSelectedTableCellStyle" ==> setSelectedTableCellStyle
            "setSectionBreak" ==> setSectionBreak
            "importFromDocx" ==> importFromDocx
            "importFromDocxWithHeaderFooters" ==> importFromDocxWithHeaderFooters
            "importFromDocxWithHeaderFooterParts" ==> importFromDocxWithHeaderFooterParts
            "importFromDocxWithPageSetupAndHeaderFooterParts" ==> importFromDocxWithPageSetupAndHeaderFooterParts
            "importFromDocxPackage" ==> importFromDocxPackage
            "setPageSetup" ==> setPageSetup
            "setPageSetupOptions" ==> setPageSetupOptions
            "getPageSetup" ==> getPageSetup
            "setHeaderFooterText" ==> setHeaderFooterText
            "getHeaderFooterText" ==> getHeaderFooterText
            "setHeaderFooterTableText" ==> setHeaderFooterTableText
            "setHeaderFooterImageText" ==> setHeaderFooterImageText
            "setCurrentSectionHeaderFooterText" ==> setCurrentSectionHeaderFooterText
            "exportToDocx" ==> exportToDocx
            "exportDocumentPackage" ==> exportDocumentPackage
            "insertTextBox" ==> insertTextBox
            "insertCaption" ==> insertCaption
            "insertCrossReferenceToLatestCaption" ==> insertCrossReferenceToLatestCaption
            "insertEquation" ==> insertEquation
            "setCurrentTextBoxStyle" ==> setCurrentTextBoxStyle
        ]

#else

module API =
    // Server-side stubs
    type Editor = unit
    let initialize _ = ()
    let getText _ = ""
    let insertTextAt _ _ = ()
    let setContent _ _ = ()
    let testRope () = obj()

#endif
