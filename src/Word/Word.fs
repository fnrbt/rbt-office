namespace Rbt.Word

#if FABLE_COMPILER

module API =
    open System.Collections.Generic
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types

    module Doc = Document
    module Rope = Rope
    module TextBuf = TextBuffer
    module Ren = Renderer
    module Inp = Input

    let private optionalText value = if isNull (box value) || value = "" then None else Some value

    let private clearArray (items: 'T array) =
        for i = 0 to items.Length - 1 do
            items.[i] <- Unchecked.defaultof<'T>

    let private referenceHash (value: obj) =
        System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(value)

    let private referenceSlot (value: obj) slotCount =
        referenceHash value &&& (slotCount - 1)

    /// Editor instance
    type Editor = {
        Handler: Inp.InputHandler
        mutable Doc: Doc.Document
    }

    let private layoutsForEditor (editor: Editor) =
        Ren.cachedLayouts editor.Handler.RenderContext editor.Doc
        |> Option.defaultWith (fun () -> Ren.layoutDocument editor.Handler.RenderContext editor.Doc)

    /// Apply a change made through the API (toolbar, dialogs, scripts). It
    /// becomes one step in the same undo history as keystrokes; a change that
    /// only moves the selection is shown without being recorded.
    let private commitDocumentUpdate (editor: Editor) newDoc =
        if not (obj.ReferenceEquals(newDoc, editor.Doc)) then
            Inp.applyEdit editor.Handler "Edit" newDoc
            editor.Doc <- newDoc

    /// Show a newly loaded document (import, setContent). It replaces the
    /// previous document wholesale, so the undo history is cleared rather
    /// than letting Ctrl+Z swap the old document back in.
    let private loadDocument (editor: Editor) newDoc =
        Inp.loadDocument editor.Handler newDoc
        editor.Doc <- newDoc

    /// Undo the last change, whether typed or made through the API.
    /// Returns false when there was nothing to undo.
    let undo (editor: Editor) =
        Inp.undo editor.Handler

    /// Redo the last undone change. Returns false when there was nothing to redo.
    let redo (editor: Editor) =
        Inp.redo editor.Handler

    let canUndo (editor: Editor) =
        Inp.canUndo editor.Handler

    let canRedo (editor: Editor) =
        Inp.canRedo editor.Handler

    /// Detach the editor's window-level listeners. Call when the canvas is removed.
    let dispose (editor: Editor) =
        Inp.dispose editor.Handler

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
        Ren.clearLayoutCache editor.Handler.RenderContext
        Ren.render editor.Handler.RenderContext editor.Doc

    let setTextMeasurers (editor: Editor) (measureText: Layout.TextMeasurer) (measureRange: Layout.TextRangeMeasurer) =
        editor.Handler.RenderContext.TextMeasurer <- Some measureText
        editor.Handler.RenderContext.TextRangeMeasurer <- Some measureRange
        Ren.clearLayoutCache editor.Handler.RenderContext
        Ren.render editor.Handler.RenderContext editor.Doc

    let setTextRangeMeasurer (editor: Editor) (measureRange: Layout.TextRangeMeasurer) =
        editor.Handler.RenderContext.TextRangeMeasurer <- Some measureRange
        Ren.clearLayoutCache editor.Handler.RenderContext
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
        Ren.clearLayoutCache editor.Handler.RenderContext
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
        commitDocumentUpdate editor newDoc

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
        match editor.Doc.Selection with
        | None -> ()
        | Some _ ->
            let newDoc = { editor.Doc with Selection = None }
            commitDocumentUpdate editor newDoc

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
        let result = ResizeArray<obj>()
        for layout in layoutsForEditor editor do
            for line in layout.Lines do
                let alignOffset = Layout.alignmentOffset layout line
                for segment in line.Segments do
                    match segment with
                    | Layout.TextSegment (text, style, x, width, startOffset) ->
                        result.Add(createObj [
                            "Kind" ==> "text"
                            "Text" ==> text
                            "BlockId" ==> layout.BlockId
                            "StartOffset" ==> startOffset
                            "X" ==> (x + alignOffset)
                            "Y" ==> line.Y
                            "Width" ==> width
                            "Height" ==> line.Height
                            "FontFamily" ==> style.FontFamily
                            "FontSize" ==> style.FontSize
                            "Bold" ==> style.Bold
                            "Italic" ==> style.Italic
                            "Underline" ==> style.Underline
                            "Color" ==> style.Color
                        ])
                    | Layout.ImageSegment (_, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, _, _, _, _) ->
                        result.Add(createObj [
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
                        ])
                    | Layout.MathSegment (_, style, x, width, ascent, descent, startOffset, _) ->
                        result.Add(createObj [
                            "Kind" ==> "math"
                            "BlockId" ==> layout.BlockId
                            "StartOffset" ==> startOffset
                            "X" ==> (x + alignOffset)
                            "Y" ==> line.Y
                            "Width" ==> width
                            "Height" ==> (ascent + descent)
                            "FontFamily" ==> style.FontFamily
                            "FontSize" ==> style.FontSize
                        ])
        result.ToArray()

    type private ParagraphLayoutContext = {
        ParagraphIndex: int
        TableIndex: int option
        RowIndex: int option
        ColumnIndex: int option
        EndColumnIndex: int option
        MaxRows: int option
        MaxColumns: int option
        InTextBox: bool
    }

    let private jsOpt value =
        match value with
        | Some actual -> box actual
        | None -> null

    let private pageStride (doc: Doc.Document) =
        doc.PageSetup.PageHeight + WordLayout.Page.gap

    let private pageForYInDoc (doc: Doc.Document) y =
        int (floor (max 0.0 y / pageStride doc)) + 1

    let private pageCountFromLines (doc: Doc.Document) (layouts: Layout.BlockLayout list) =
        let mutable pageCount = 1
        for layout in layouts do
            for line in layout.Lines do
                pageCount <- max pageCount (pageForYInDoc doc (line.Y + line.Height))
        pageCount

    let private pageTopForYInDoc (doc: Doc.Document) y =
        floor (max 0.0 y / pageStride doc) * pageStride doc

    let private paragraphMap (doc: Doc.Document) =
        let paragraphs = Dictionary<string, Doc.ParagraphBlock>()
        for block in doc.Blocks do
            match block with
            | Doc.Paragraph para
            | Doc.Heading (_, para) ->
                paragraphs.[para.Id] <- para
            | Doc.Table table ->
                for row in table.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            paragraphs.[para.Id] <- para
            | Doc.TextBox textBox ->
                for para in textBox.Content do
                    paragraphs.[para.Id] <- para
        paragraphs


    let private tryFindParagraph paragraphMap paragraphId =
        match (paragraphMap: Dictionary<string, Doc.ParagraphBlock>).TryGetValue paragraphId with
        | true, paragraph -> Some paragraph
        | false, _ -> None

    let private tryFindDictionary key (dictionary: Dictionary<'K, 'V>) =
        match dictionary.TryGetValue key with
        | true, value -> Some value
        | false, _ -> None

    let private maxVisualColumns (table: Doc.TableBlock) =
        let mutable maxColumns = 0
        for row in table.Rows do
            let mutable columns = 0
            for cell in row.Cells do
                columns <- columns + max 1 cell.ColSpan
            if columns > maxColumns then
                maxColumns <- columns
        maxColumns

    let private paragraphContexts (doc: Doc.Document) =
        let mutable paragraphIndex = 0
        let mutable tableIndex = 0
        let contexts = Dictionary<string, ParagraphLayoutContext>()

        let addContext (para: Doc.ParagraphBlock) (context: ParagraphLayoutContext) =
            paragraphIndex <- paragraphIndex + 1
            contexts.[para.Id] <- { context with ParagraphIndex = paragraphIndex }

        for block in doc.Blocks do
            match block with
            | Doc.Paragraph para
            | Doc.Heading (_, para) ->
                addContext para {
                    ParagraphIndex = 0
                    TableIndex = None
                    RowIndex = None
                    ColumnIndex = None
                    EndColumnIndex = None
                    MaxRows = None
                    MaxColumns = None
                    InTextBox = false
                }
            | Doc.Table table ->
                tableIndex <- tableIndex + 1
                let maxRows = table.Rows.Length
                let maxColumns = maxVisualColumns table
                table.Rows
                |> List.iteri (fun rowIndex row ->
                    let mutable columnIndex = 1
                    for cell in row.Cells do
                        let span = max 1 cell.ColSpan
                        for para in cell.Content do
                            addContext para {
                                ParagraphIndex = 0
                                TableIndex = Some tableIndex
                                RowIndex = Some (rowIndex + 1)
                                ColumnIndex = Some columnIndex
                                EndColumnIndex = Some (columnIndex + span - 1)
                                MaxRows = Some maxRows
                                MaxColumns = Some maxColumns
                                InTextBox = false
                            }
                        columnIndex <- columnIndex + span)
            | Doc.TextBox textBox ->
                for para in textBox.Content do
                    addContext para {
                        ParagraphIndex = 0
                        TableIndex = None
                        RowIndex = None
                        ColumnIndex = None
                        EndColumnIndex = None
                        MaxRows = None
                        MaxColumns = None
                        InTextBox = true
                    }

        contexts


    let private canonicalTableYFromLayouts settings measureText measureRange doc (layouts: Layout.BlockLayout list) (table: Doc.TableBlock) =
        let layoutById = Dictionary<string, Layout.BlockLayout>()
        for layout in layouts do
            if not (layoutById.ContainsKey layout.BlockId) then
                layoutById.[layout.BlockId] <- layout

        let mutable firstActualContentY = System.Double.PositiveInfinity
        for row in table.Rows do
            for cell in row.Cells do
                for para in cell.Content do
                    match layoutById.TryGetValue para.Id with
                    | true, layout -> firstActualContentY <- min firstActualContentY layout.Y
                    | false, _ -> ()

        if System.Double.IsPositiveInfinity firstActualContentY then
            None
        else
            let planAtZero = Layout.buildTableLayoutPlanForDocumentAt settings measureText measureRange doc table 0.0
            let resolveAlignmentHeight = Layout.tableCellVerticalAlignmentHeightResolverForPlan planAtZero
            let mutable firstPlannedContentY = System.Double.PositiveInfinity
            for row in planAtZero.Rows do
                let mutable groupIndex = 0
                for group in row.Measurement.CellGroups do
                    let alignmentHeight = resolveAlignmentHeight row.RowIndex groupIndex row.Height
                    let offset = WordLayout.Table.verticalContentOffset group.VerticalAlignment alignmentHeight group.PaddingTop group.PaddingBottom group.ContentHeight
                    for layout in group.Layouts do
                        firstPlannedContentY <- min firstPlannedContentY (layout.Y + offset)
                    groupIndex <- groupIndex + 1
            if System.Double.IsPositiveInfinity firstPlannedContentY then
                None
            else
                Some (firstActualContentY - firstPlannedContentY)

    let private strokeGeometry (stroke: WordLayout.Table.PrintStroke option) =
        match stroke with
        | None -> null
        | Some value ->
            createObj [
                "Color" ==> value.Color
                "Width" ==> value.Width
                "Style" ==> value.Style
            ]

    let getTableGeometry (editor: Editor) =
        let ctx = editor.Handler.RenderContext
        let settings = Ren.layoutSettings ctx editor.Doc
        let measureText = Ren.measureWithContext ctx
        let measureRange = Ren.measureRangeWithContext ctx
        let layouts = layoutsForEditor editor
        let layoutsById = Dictionary<string, Layout.BlockLayout>()
        for layout in layouts do
            if not (layoutsById.ContainsKey layout.BlockId) then
                layoutsById.[layout.BlockId] <- layout
        let mutable currentY = editor.Doc.PageSetup.MarginTop
        let mutable tableIndex = 0
        let rows = ResizeArray<obj>()
        let cells = ResizeArray<obj>()
        let segments = ResizeArray<obj>()

        let paragraphEndY (para: Doc.ParagraphBlock) =
            match layoutsById.TryGetValue para.Id with
            | true, layout -> Some (layout.Y + layout.Height)
            | false, _ -> None

        for block in editor.Doc.Blocks do
            match block with
            | Doc.Paragraph para
            | Doc.Heading (_, para) ->
                paragraphEndY para |> Option.iter (fun y -> currentY <- y)
                if para.Metadata.PageBreakAfter || para.Metadata.SectionBreak = Some "nextPage" then
                    currentY <- WordLayout.Page.advanceToNextPage settings.PageHeight settings.PaddingTop currentY
            | Doc.Table table ->
                tableIndex <- tableIndex + 1
                let tableY =
                    canonicalTableYFromLayouts settings measureText measureRange editor.Doc layouts table
                    |> Option.defaultValue currentY
                let plan = Layout.buildTableLayoutPlanForDocumentAt settings measureText measureRange editor.Doc table tableY

                for row in plan.Rows do
                    rows.Add(createObj [
                        "Kind" ==> "table-row"
                        "Table" ==> tableIndex
                        "Row" ==> (row.RowIndex + 1)
                        "Y" ==> row.Y
                        "PageY" ==> (row.Y - pageTopForYInDoc editor.Doc row.Y)
                        "Height" ==> row.Height
                        "FragmentCount" ==> row.Fragments.Length
                        "IsHeader" ==> row.IsHeader
                        "CanSplit" ==> row.CanSplit
                    ])

                let cellFragments = Layout.tableCellFragmentPlans plan
                for cell in cellFragments do
                    cells.Add(createObj [
                        "Kind" ==> "table-cell"
                        "Table" ==> tableIndex
                        "Row" ==> (cell.RowIndex + 1)
                        "Cell" ==> (cell.CellIndex + 1)
                        "StartColumn" ==> (cell.StartColumn + 1)
                        "ColumnSpan" ==> cell.ColumnSpan
                        "RowSpan" ==> cell.RowSpan
                        "Page" ==> cell.Page
                        "X" ==> cell.CellX
                        "Y" ==> cell.CellY
                        "PageY" ==> (cell.CellY - pageTopForYInDoc editor.Doc cell.CellY)
                        "Width" ==> cell.CellWidth
                        "Height" ==> cell.CellHeight
                        "BorderTop" ==> strokeGeometry cell.BorderTop
                        "BorderRight" ==> strokeGeometry cell.BorderRight
                        "BorderBottom" ==> strokeGeometry cell.BorderBottom
                        "BorderLeft" ==> strokeGeometry cell.BorderLeft
                    ])

                for segment in cellFragments |> Layout.tableCellFragmentBorderCells |> RedBat.BorderGrid.resolve do
                    segments.Add(createObj [
                        "Kind" ==> "table-border-segment"
                        "Table" ==> tableIndex
                        "X1" ==> segment.X1
                        "Y1" ==> segment.Y1
                        "X2" ==> segment.X2
                        "Y2" ==> segment.Y2
                        "PageY1" ==> (segment.Y1 - pageTopForYInDoc editor.Doc segment.Y1)
                        "PageY2" ==> (segment.Y2 - pageTopForYInDoc editor.Doc segment.Y2)
                        "Color" ==> segment.Stroke.Color
                        "Width" ==> segment.Stroke.Width
                        "Style" ==> segment.Stroke.Style
                    ])

                currentY <- plan.NextY
            | Doc.TextBox _ -> ()

        let result = ResizeArray<obj>(rows.Count + cells.Count + segments.Count)
        for item in rows do result.Add item
        for item in cells do result.Add item
        for item in segments do result.Add item
        result.ToArray()

    let private sectionNumbersByParagraphId (doc: Doc.Document) =
        let mutable section = 1
        let sections = Dictionary<string, int>()

        let addParagraph (para: Doc.ParagraphBlock) =
            sections.[para.Id] <- section

        for block in doc.Blocks do
            match block with
            | Doc.Paragraph para
            | Doc.Heading (_, para) ->
                addParagraph para
                if para.Metadata.SectionBreak.IsSome then section <- section + 1
            | Doc.Table table ->
                for row in table.Rows do
                    for cell in row.Cells do
                        for para in cell.Content do
                            addParagraph para
            | Doc.TextBox textBox ->
                for para in textBox.Content do
                    addParagraph para

        sections


    let private measureRangeForEditor (editor: Editor) =
        match editor.Handler.RenderContext.TextRangeMeasurer with
        | Some measureRange -> measureRange
        | None -> Layout.rangeMeasurerFromTextMeasurer (Ren.measureTextStyled editor.Handler.RenderContext)

    let private lineHasFieldCode (line: Layout.MeasuredLine) =
        line.Segments
        |> List.exists (function
            | Layout.TextSegment (_, style, _, _, _) -> style.FieldCode.IsSome
            | Layout.ImageSegment _ -> false
            | Layout.MathSegment _ -> false)

    let private lineHasFieldResult (line: Layout.MeasuredLine) =
        line.Segments
        |> List.exists (function
            | Layout.TextSegment (_, style, _, _, _) -> style.FieldResult.IsSome
            | Layout.ImageSegment _ -> false
            | Layout.MathSegment _ -> false)

    let private lineLeft pageLeft (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) =
        let alignOffset = Layout.alignmentOffset layout line
        let mutable hasBounds = false
        let mutable left = System.Double.PositiveInfinity
        let mutable right = System.Double.NegativeInfinity
        let addBounds x width =
            let x0 = pageLeft + x + alignOffset
            let x1 = x0 + width
            if x0 < left then left <- x0
            if x1 > right then right <- x1
            hasBounds <- true
        for segment in line.Segments do
            match segment with
            | Layout.TextSegment (_, _, x, width, _) -> addBounds x width
            | Layout.ImageSegment (_, x, _, width, _, _, _, _, _, _, _, _, _) -> addBounds x width
            | Layout.MathSegment (_, _, x, width, _, _, _, _) -> addBounds x width
        if hasBounds then left, max 0.0 (right - left)
        else pageLeft + alignOffset, line.Width

    let private isTableContext context =
        context |> Option.bind _.TableIndex |> Option.isSome

    let private textBoundaryOrigin (doc: Doc.Document) pageLeft (layout: Layout.BlockLayout) (line: Layout.MeasuredLine) context =
        if isTableContext context then
            match layout.Lines |> List.tryHead with
            | Some firstLine ->
                let firstLeft, _ = lineLeft pageLeft layout firstLine
                firstLeft, firstLine.Y
            | None -> pageLeft + layout.Indent, layout.Y
        elif context |> Option.map _.InTextBox |> Option.defaultValue false then
            pageLeft + layout.Indent, layout.Y
        else
            pageLeft, pageTopForYInDoc doc line.Y + doc.PageSetup.MarginTop

    let private rangeBounds (doc: Doc.Document) (tryFindLayout: string -> Layout.BlockLayout option) (measureRange: Layout.TextRangeMeasurer) (blockId: string) (startOffset: int) (endOffset: int) =
        let pageLeft = doc.PageSetup.MarginLeft
        match tryFindLayout blockId with
        | None -> None
        | Some layout ->
            let maxOffset = Layout.maxOffsetInLayout layout
            let startOffset = max 0 (min startOffset maxOffset)
            let endOffset = max startOffset (min endOffset maxOffset)
            let mutable found = false
            let mutable left = 0.0
            let mutable top = 0.0
            let mutable right = 0.0
            let mutable bottom = 0.0
            let includeFragment x1 y1 x2 y2 =
                if found then
                    if x1 < left then left <- x1
                    if y1 < top then top <- y1
                    if x2 > right then right <- x2
                    if y2 > bottom then bottom <- y2
                else
                    found <- true
                    left <- x1
                    top <- y1
                    right <- x2
                    bottom <- y2

            if startOffset = endOffset then
                let line = Layout.lineForOffset layout startOffset
                let alignOffset = Layout.alignmentOffset layout line
                let x = pageLeft + alignOffset + Layout.xForOffsetWithRange measureRange line startOffset
                includeFragment x line.Y x (line.Y + line.Height)
            else
                for line in layout.Lines do
                    let lineStart = max startOffset line.StartOffset
                    let lineEnd = min endOffset line.EndOffset
                    if not (lineEnd < lineStart || lineEnd <= line.StartOffset || lineStart >= line.EndOffset) then
                        let alignOffset = Layout.alignmentOffset layout line
                        let x1 = pageLeft + alignOffset + Layout.xForOffsetWithRange measureRange line lineStart
                        let x2 = pageLeft + alignOffset + Layout.xForOffsetWithRange measureRange line lineEnd
                        includeFragment (min x1 x2) line.Y (max x1 x2) (line.Y + line.Height)

            if found then Some (left, top, right - left, bottom - top) else None

    let getRangeBounds (editor: Editor) (blockId: string) (startOffset: int) (endOffset: int) =
        let tryFindLayout = Ren.layoutLookupFor editor.Handler.RenderContext editor.Doc
        let measureRange = measureRangeForEditor editor
        match rangeBounds editor.Doc tryFindLayout measureRange blockId startOffset endOffset with
        | None -> null
        | Some (x, y, width, height) ->
            let pageTop = pageTopForYInDoc editor.Doc y
            createObj [
                "BlockId" ==> blockId
                "StartOffset" ==> startOffset
                "EndOffset" ==> endOffset
                "X" ==> x
                "Y" ==> y
                "PageX" ==> x
                "PageY" ==> (y - pageTop)
                "Width" ==> width
                "Height" ==> height
                "Page" ==> pageForYInDoc editor.Doc y
            ]

    let private isWordSeparator c =
        c = ' ' || c = '\t' || c = '\r' || c = '\n'

    let private countWordsInRange (text: string) start length =
        let endIndex = min text.Length (start + length)
        let mutable count = 0
        let mutable inWord = false
        let mutable i = max 0 start
        while i < endIndex do
            if isWordSeparator text.[i] then
                inWord <- false
            elif not inWord then
                count <- count + 1
                inWord <- true
            i <- i + 1
        count

    let private markerColumnOffset (metadata: Doc.BlockMetadata) =
        let markerLengthWithSeparator (marker: string) =
            if System.String.IsNullOrEmpty marker then 0
            elif marker.EndsWith(" ") || marker.EndsWith("\t") then marker.Length
            else marker.Length + 1

        match metadata.ListStyle with
        | Doc.NoList -> 0
        | Doc.Bullet ->
            metadata.ListLevelText
            |> Option.defaultValue "•"
            |> markerLengthWithSeparator
        | Doc.Numbered ->
            let marker =
                match metadata.ListLevelText with
                | Some template when template.Contains("%") ->
                    let mutable text = template
                    for level in 1 .. 9 do
                        text <- text.Replace("%" + string level, "1")
                    text
                | Some template -> template
                | None -> "1."
            markerLengthWithSeparator marker

    let private firstCharacterColumn (paragraph: Doc.ParagraphBlock option) (line: Layout.MeasuredLine) offset =
        let listOffset =
            if line.StartOffset = 0 then
                paragraph
                |> Option.map (fun para -> markerColumnOffset para.Metadata)
                |> Option.defaultValue 0
            else 0
        max 1 (offset - line.StartOffset + 1 + listOffset)

    let computeRangeStatistics (editor: Editor) (blockId: string) (startOffset: int) (endOffset: int) =
        match Doc.findParagraph editor.Doc blockId with
        | None -> null
        | Some para ->
            let text = TextBuf.toString para.Buffer
            let boundedStart = max 0 (min startOffset text.Length)
            let boundedEnd = max boundedStart (min endOffset text.Length)
            let rangeLength = boundedEnd - boundedStart
            let layouts = layoutsForEditor editor
            let context = paragraphContexts editor.Doc |> tryFindDictionary blockId
            let isInTable = context |> Option.bind _.TableIndex |> Option.isSome
            let mutable rangeLineCount = 0
            let pages = HashSet<int>()
            if not isInTable then
                for layout in layouts do
                    if layout.BlockId = blockId then
                        for line in layout.Lines do
                            let overlaps =
                                if boundedStart = boundedEnd then
                                    boundedStart >= line.StartOffset && boundedStart <= line.EndOffset
                                else
                                    boundedStart < line.EndOffset && boundedEnd > line.StartOffset
                            if overlaps then
                                rangeLineCount <- rangeLineCount + 1
                                pages.Add(pageForYInDoc editor.Doc line.Y) |> ignore

            createObj [
                "Lines" ==> rangeLineCount
                "Pages" ==> pages.Count
                "Paragraphs" ==> 1
                "Words" ==> countWordsInRange text boundedStart rangeLength
                "Characters" ==> rangeLength
                "CharactersWithSpaces" ==> rangeLength
            ]

    let private pageLineNumber (doc: Doc.Document) (layouts: Layout.BlockLayout list) (targetLayout: Layout.BlockLayout) (targetLine: Layout.MeasuredLine) =
        let page = pageForYInDoc doc targetLine.Y
        let targetLeft, _ = lineLeft doc.PageSetup.MarginLeft targetLayout targetLine
        let mutable beforeCount = 0
        let mutable equalBeforeTarget = 0
        let mutable foundTarget = false
        for layout in layouts do
            for line in layout.Lines do
                if pageForYInDoc doc line.Y = page then
                    let left, _ = lineLeft doc.PageSetup.MarginLeft layout line
                    if line.Y < targetLine.Y || (line.Y = targetLine.Y && left < targetLeft) then
                        beforeCount <- beforeCount + 1
                    elif line.Y = targetLine.Y && left = targetLeft && not foundTarget then
                        if layout.BlockId = targetLayout.BlockId &&
                           line.StartOffset = targetLine.StartOffset &&
                           line.EndOffset = targetLine.EndOffset then
                            foundTarget <- true
                        else
                            equalBeforeTarget <- equalBeforeTarget + 1
        if foundTarget then beforeCount + equalBeforeTarget + 1 else 0

    // Word/PDF range positions use a text-top coordinate for large default-spaced
    // runs; our layout line keeps the line-box top for rendering and pagination.
    let private wordVerticalTextTopOffsetFactor = 0.21
    let private wordVerticalTextTopOffsetMinFontSize = 24.0

    let private usesDefaultWordRangeTopOffset (metadata: Doc.BlockMetadata) =
        metadata.SpacingBefore.IsNone
        && metadata.SpacingAfter.IsNone
        && metadata.LineSpacing.IsNone
        && metadata.LineSpacingRule.IsNone
        && metadata.ContextualSpacing.IsNone

    let private wordVerticalPositionTop applyTextTopOffset (line: Layout.MeasuredLine) =
        let mutable fontSize = 0.0
        for segment in line.Segments do
            match segment with
            | Layout.TextSegment (_, style, _, _, _)
            | Layout.MathSegment (_, style, _, _, _, _, _, _) ->
                fontSize <- max fontSize (WordLayout.Text.printFontSize style)
            | Layout.ImageSegment _ -> ()
        if applyTextTopOffset && fontSize >= wordVerticalTextTopOffsetMinFontSize then
            line.Y + fontSize * wordVerticalTextTopOffsetFactor
        else line.Y

    let getRangeInformation (editor: Editor) (blockId: string) (startOffset: int) (endOffset: int) =
        let layouts = layoutsForEditor editor
        let tryFindLayout = Ren.layoutLookupFor editor.Handler.RenderContext editor.Doc
        let contexts = paragraphContexts editor.Doc
        let sections = sectionNumbersByParagraphId editor.Doc
        let paragraphs = paragraphMap editor.Doc
        let measureRange = measureRangeForEditor editor
        let pageCount = pageCountFromLines editor.Doc layouts

        match tryFindLayout blockId with
        | None -> null
        | Some layout ->
            let offset = startOffset
            let line = Layout.lineForOffset layout offset
            let alignOffset = Layout.alignmentOffset layout line
            let x = editor.Doc.PageSetup.MarginLeft + alignOffset + Layout.xForOffsetWithRange measureRange line offset
            let page = pageForYInDoc editor.Doc line.Y
            let pageTop = pageTopForYInDoc editor.Doc line.Y
            let context = contexts |> tryFindDictionary blockId
            let isInTable = context.IsSome && context.Value.TableIndex.IsSome
            // Word's VBA Information API reports the unaligned cell content origin,
            // ignoring center/bottom vertical alignment; subtract the applied offset to match.
            let wordTop = wordVerticalPositionTop (not isInTable && usesDefaultWordRangeTopOffset layout.Metadata) line - layout.VAlignOffset
            let pageY = wordTop - pageTop
            let boundaryLeft, boundaryTop =
                textBoundaryOrigin editor.Doc editor.Doc.PageSetup.MarginLeft layout line context
            let bounds = rangeBounds editor.Doc tryFindLayout measureRange blockId startOffset endOffset

            createObj [
                "BlockId" ==> blockId
                "StartOffset" ==> startOffset
                "EndOffset" ==> endOffset
                "wdActiveEndAdjustedPageNumber" ==> page
                "wdActiveEndPageNumber" ==> page
                "wdActiveEndSectionNumber" ==> (sections |> tryFindDictionary blockId |> Option.defaultValue 1)
                "wdNumberOfPagesInDocument" ==> pageCount
                "wdHorizontalPositionRelativeToPage" ==> x
                "wdVerticalPositionRelativeToPage" ==> pageY
                "wdHorizontalPositionRelativeToTextBoundary" ==> (x - boundaryLeft)
                "wdVerticalPositionRelativeToTextBoundary" ==> (max 0.0 (wordTop - boundaryTop))
                "wdFirstCharacterColumnNumber" ==> firstCharacterColumn (tryFindParagraph paragraphs blockId) line offset
                "wdFirstCharacterLineNumber" ==> pageLineNumber editor.Doc layouts layout line
                "wdWithInTable" ==> isInTable
                "wdStartOfRangeRowNumber" ==> (context |> Option.bind _.RowIndex |> jsOpt)
                "wdEndOfRangeRowNumber" ==> (context |> Option.bind _.RowIndex |> jsOpt)
                "wdMaximumNumberOfRows" ==> (context |> Option.bind _.MaxRows |> jsOpt)
                "wdStartOfRangeColumnNumber" ==> (context |> Option.bind _.ColumnIndex |> jsOpt)
                "wdEndOfRangeColumnNumber" ==> (context |> Option.bind _.EndColumnIndex |> jsOpt)
                "wdMaximumNumberOfColumns" ==> (context |> Option.bind _.MaxColumns |> jsOpt)
                "wdAtEndOfRowMarker" ==> false
                "wdInContentControl" ==> layout.Metadata.ContentControlId.IsSome
                "wdInFieldCode" ==> lineHasFieldCode line
                "wdInFieldResult" ==> lineHasFieldResult line
                "wdFrameIsSelected" ==> (context |> Option.map _.InTextBox |> Option.defaultValue false)
                "wdRevisionMarking" ==> editor.Handler.TrackChanges
                "Bounds" ==> (
                    match bounds with
                    | None -> null
                    | Some (bx, by, bw, bh) ->
                        createObj [
                            "X" ==> bx
                            "Y" ==> by
                            "PageY" ==> (by - pageTopForYInDoc editor.Doc by)
                            "Width" ==> bw
                            "Height" ==> bh
                        ])
            ]

    let getSelectionInformation (editor: Editor) =
        match editor.Doc.Selection with
        | None -> null
        | Some selection ->
            getRangeInformation editor selection.Focus.BlockId selection.Focus.Offset selection.Focus.Offset

    let getLayoutLines (editor: Editor) =
        let pageLeft = editor.Doc.PageSetup.MarginLeft
        let layouts = layoutsForEditor editor
        let contexts = paragraphContexts editor.Doc
        let sections = sectionNumbersByParagraphId editor.Doc
        let paragraphs = paragraphMap editor.Doc
        let pageCount = pageCountFromLines editor.Doc layouts
        let pageLineCounts = Dictionary<int, int>()

        let nextPageLine page =
            let value =
                match pageLineCounts.TryGetValue page with
                | true, count -> count + 1
                | false, _ -> 1
            pageLineCounts.[page] <- value
            value

        let result = ResizeArray<obj>()
        let mutable layoutIndex = 0
        for layout in layouts do
            let mutable lineIndex = 0
            for line in layout.Lines do
                let alignOffset = Layout.alignmentOffset layout line
                let segments = ResizeArray<obj>()
                let text = System.Text.StringBuilder()
                for segment in line.Segments do
                    match segment with
                    | Layout.TextSegment (segmentText, _, x, width, startOffset) ->
                        text.Append segmentText |> ignore
                        segments.Add(createObj [
                            "Kind" ==> "text"
                            "Text" ==> segmentText
                            "StartOffset" ==> startOffset
                            "X" ==> (pageLeft + x + alignOffset)
                            "Width" ==> width
                        ])
                    | Layout.ImageSegment (_, x, yOffset, width, height, startOffset, floating, wrap, relativeHeight, _, _, _, _) ->
                        text.Append "\uFFFC" |> ignore
                        segments.Add(createObj [
                            "Kind" ==> "image"
                            "Text" ==> "\uFFFC"
                            "StartOffset" ==> startOffset
                            "X" ==> (pageLeft + x + alignOffset)
                            "Y" ==> (line.Y + yOffset)
                            "Width" ==> width
                            "Height" ==> height
                            "Floating" ==> floating
                            "Wrap" ==> wrap
                            "RelativeHeight" ==> relativeHeight
                        ])
                    | Layout.MathSegment (_, _, x, width, _, _, startOffset, _) ->
                        text.Append "\uFFFC" |> ignore
                        segments.Add(createObj [
                            "Kind" ==> "math"
                            "Text" ==> "\uFFFC"
                            "StartOffset" ==> startOffset
                            "X" ==> (pageLeft + x + alignOffset)
                            "Width" ==> width
                        ])

                let x, width =
                    lineLeft pageLeft layout line

                let page = pageForYInDoc editor.Doc line.Y
                let pageTop = pageTopForYInDoc editor.Doc line.Y
                let pageLine = nextPageLine page
                let context = contexts |> tryFindDictionary layout.BlockId
                let isInTable = context |> Option.bind _.TableIndex |> Option.isSome
                let boundaryLeft, boundaryTop =
                    textBoundaryOrigin editor.Doc pageLeft layout line context
                // Match Word's VBA Information API, which reports the unaligned cell
                // content origin regardless of center/bottom vertical alignment.
                let wordTop = wordVerticalPositionTop (not isInTable && usesDefaultWordRangeTopOffset layout.Metadata) line - layout.VAlignOffset
                result.Add(createObj [
                    "Kind" ==> "line"
                    "BlockId" ==> layout.BlockId
                    "BlockIndex" ==> layoutIndex
                    "LineIndex" ==> lineIndex
                    "Page" ==> page
                    "AdjustedPage" ==> page
                    "Section" ==> (sections |> tryFindDictionary layout.BlockId |> Option.defaultValue 1)
                    "PageCount" ==> pageCount
                    "X" ==> x
                    "Y" ==> wordTop
                    "PageY" ==> (wordTop - pageTop)
                    "Width" ==> width
                    "Height" ==> line.Height
                    "LineWidth" ==> line.Width
                    "HorizontalPositionRelativeToPage" ==> x
                    "VerticalPositionRelativeToPage" ==> (wordTop - pageTop)
                    "HorizontalPositionRelativeToTextBoundary" ==> (x - boundaryLeft)
                    "VerticalPositionRelativeToTextBoundary" ==> (max 0.0 (wordTop - boundaryTop))
                    "FirstCharacterLineNumber" ==> pageLine
                    "FirstCharacterColumnNumber" ==> firstCharacterColumn (tryFindParagraph paragraphs layout.BlockId) line line.StartOffset
                    "StartOffset" ==> line.StartOffset
                    "EndOffset" ==> line.EndOffset
                    "BlockY" ==> layout.Y
                    "BlockHeight" ==> layout.Height
                    "AvailableWidth" ==> layout.AvailableWidth
                    "Indent" ==> layout.Indent
                    "InTable" ==> isInTable
                    "Table" ==> (context |> Option.bind _.TableIndex |> jsOpt)
                    "Row" ==> (context |> Option.bind _.RowIndex |> jsOpt)
                    "Column" ==> (context |> Option.bind _.ColumnIndex |> jsOpt)
                    "StartRow" ==> (context |> Option.bind _.RowIndex |> jsOpt)
                    "EndRow" ==> (context |> Option.bind _.RowIndex |> jsOpt)
                    "MaximumRows" ==> (context |> Option.bind _.MaxRows |> jsOpt)
                    "StartColumn" ==> (context |> Option.bind _.ColumnIndex |> jsOpt)
                    "EndColumn" ==> (context |> Option.bind _.EndColumnIndex |> jsOpt)
                    "MaximumColumns" ==> (context |> Option.bind _.MaxColumns |> jsOpt)
                    "AtEndOfRowMarker" ==> false
                    "InContentControl" ==> layout.Metadata.ContentControlId.IsSome
                    "InFieldCode" ==> lineHasFieldCode line
                    "InFieldResult" ==> lineHasFieldResult line
                    "FrameIsSelected" ==> (context |> Option.map _.InTextBox |> Option.defaultValue false)
                    "RevisionMarking" ==> editor.Handler.TrackChanges
                    "RangeLines" ==> (if isInTable then 0 else 1)
                    "RangePages" ==> (if isInTable then 0 else 1)
                    "Text" ==> (text.ToString())
                    "Segments" ==> (segments.ToArray())
                ])
                lineIndex <- lineIndex + 1
            layoutIndex <- layoutIndex + 1
        result.ToArray()

    let getVbaLayoutLines (editor: Editor) =
        getLayoutLines editor

    /// Find literal text matches across editable paragraphs.
    let findText (editor: Editor) (query: string) (matchCase: bool) =
        let matches = Doc.findTextArray query matchCase editor.Doc
        let result = Array.zeroCreate<obj> matches.Length
        for i = 0 to matches.Length - 1 do
            let matchInfo = matches.[i]
            result.[i] <-
                createObj [
                    "BlockId" ==> matchInfo.BlockId
                    "Start" ==> matchInfo.Start
                    "End" ==> matchInfo.End
                    "Text" ==> matchInfo.Text
                ]
        result

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
        commitDocumentUpdate editor newDoc

    let selectRevisionInfo (editor: Editor) (revisionObj: obj) =
        let blockId = revisionObj?BlockId |> unbox<string>
        let startOffset = revisionObj?Start |> unbox<int>
        let endOffset = revisionObj?End |> unbox<int>
        let newDoc = Doc.setSelection { BlockId = blockId; Offset = startOffset } { BlockId = blockId; Offset = endOffset } editor.Doc
        commitDocumentUpdate editor newDoc

    /// Replace all literal text matches across editable paragraphs.
    let replaceAllText (editor: Editor) (query: string) (replacement: string) (matchCase: bool) =
        let (newDoc, count) = Doc.replaceAllText query replacement matchCase editor.Doc
        if count > 0 then
            commitDocumentUpdate editor newDoc
        count
    
    /// Insert text at cursor
    let insertTextAt (editor: Editor) (text: string) =
        let newDoc =
            if editor.Handler.TrackChanges then Doc.insertTrackedText editor.Handler.TrackChangesAuthor editor.Handler.TrackChangesDate text editor.Doc
            else Doc.insertText text editor.Doc
        commitDocumentUpdate editor newDoc

    let insertFieldAt (editor: Editor) (fieldCode: string) (result: string) =
        let newDoc = Doc.insertField fieldCode result editor.Doc
        commitDocumentUpdate editor newDoc

    let private pageNumberForY (doc: Doc.Document) y =
        pageForYInDoc doc y

    let private displayPageNumberForPhysicalPage (restartPoints: (int * int) list) physicalPage =
        let mutable displayPage = physicalPage
        for restartPhysicalPage, restartDisplayPage in restartPoints do
            if restartPhysicalPage <= physicalPage then
                displayPage <- restartDisplayPage + physicalPage - restartPhysicalPage
        displayPage

    let private pageMapFromLayout (editor: Editor) =
        let layouts = layoutsForEditor editor
        let paragraphs = paragraphMap editor.Doc
        let restartPoints =
            layouts
            |> List.choose (fun layout ->
                tryFindParagraph paragraphs layout.BlockId
                |> Option.bind (fun para ->
                    para.Metadata.SectionPageNumberStart
                    |> Option.map (fun start -> pageNumberForY editor.Doc layout.Y, max 1 start)))
            |> List.sortBy fst

        let displayPageNumberForY y =
            y
            |> pageNumberForY editor.Doc
            |> displayPageNumberForPhysicalPage restartPoints

        let fieldAnchors = ResizeArray<Doc.PageFieldAnchor>()
        let bookmarkAnchors = ResizeArray<Doc.PageBookmarkAnchor>()
        let blockAnchors = ResizeArray<Doc.PageBlockAnchor>()
        let seenBookmarks = HashSet<string>()
        for layout in layouts do
            blockAnchors.Add {
                PageBlockId = layout.BlockId
                PageNumber = displayPageNumberForY layout.Y
            }
            for line in layout.Lines do
                let pageNumber = displayPageNumberForY line.Y
                for segment in line.Segments do
                    match segment with
                    | Layout.TextSegment (_, style, _, _, startOffset) ->
                        match style.FieldCode with
                        | Some code ->
                            fieldAnchors.Add {
                                PageFieldBlockId = layout.BlockId
                                PageFieldStartOffset = startOffset
                                PageFieldCode = code
                                PageNumber = pageNumber
                            }
                        | None -> ()
                        match style.BookmarkName with
                        | Some name when seenBookmarks.Add name ->
                            bookmarkAnchors.Add {
                                PageBookmarkName = name
                                PageBookmarkBlockId = layout.BlockId
                                PageBookmarkStartOffset = startOffset
                                PageNumber = pageNumber
                            }
                        | _ -> ()
                    | Layout.ImageSegment _
                    | Layout.MathSegment _ -> ()
        let mutable pageCount = 1
        for layout in layouts do
            pageCount <- max pageCount (pageNumberForY editor.Doc (layout.Y + layout.Height))
        let pageMap: Doc.PageMap = {
            PageCount = pageCount
            BlockAnchors = List.ofSeq blockAnchors
            Fields = List.ofSeq fieldAnchors
            Bookmarks = List.ofSeq bookmarkAnchors
        }
        pageMap

    let refreshFields (editor: Editor) =
        let newDoc = Doc.refreshFieldsWithPageMap (pageMapFromLayout editor) editor.Doc
        commitDocumentUpdate editor newDoc

    let insertEquation (editor: Editor) (equationText: string) =
        let newDoc = Doc.insertEquation equationText editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Set document content
    let setContent (editor: Editor) (text: string) =
        let doc = Doc.setPlainTextContent text editor.Doc
        if not (obj.ReferenceEquals(doc, editor.Doc)) then loadDocument editor doc

    /// Set document content as one paragraph per input string (unlike
    /// setContent, which puts the whole text in a single paragraph with
    /// hard breaks)
    let setParagraphs (editor: Editor) (texts: string array) =
        let doc = Doc.setParagraphContents (Array.toList texts) editor.Doc
        if not (obj.ReferenceEquals(doc, editor.Doc)) then loadDocument editor doc
    
    /// Toggle bold on selection
    let toggleBold (editor: Editor) =
        let newDoc = WordStyles.toggleBold editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Toggle italic on selection
    let toggleItalic (editor: Editor) =
        let newDoc = WordStyles.toggleItalic editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Toggle underline on selection
    let toggleUnderline (editor: Editor) =
        let newDoc = WordStyles.toggleUnderline editor.Doc
        commitDocumentUpdate editor newDoc

    /// Toggle all caps on selection
    let toggleAllCaps (editor: Editor) =
        let newDoc = WordStyles.toggleAllCaps editor.Doc
        commitDocumentUpdate editor newDoc

    /// Toggle small caps on selection
    let toggleSmallCaps (editor: Editor) =
        let newDoc = WordStyles.toggleSmallCaps editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set numbered-list marker format on current block
    let setNumberFormat (editor: Editor) (format: string) =
        let newDoc = Doc.setNumberFormat format editor.Doc
        commitDocumentUpdate editor newDoc

    /// Toggle strikethrough on selection
    let toggleStrike (editor: Editor) =
        let newDoc = WordStyles.toggleStrike editor.Doc
        commitDocumentUpdate editor newDoc

    /// Clear direct character formatting on selection
    let clearFormatting (editor: Editor) =
        let newDoc = Doc.clearFormatting editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set vertical alignment on selection
    let setVertAlignFor (editor: Editor) (align: string) =
        let newDoc = Doc.setVertAlign (if isNull (box align) || align = "" || align = "baseline" then None else Some align) editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set text color on selection
    let setColorFor (editor: Editor) (color: string) =
        let newDoc = Doc.setColor (if isNull (box color) || color = "" then None else Some color) editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set highlight color on selection
    let setHighlightFor (editor: Editor) (color: string) =
        let newDoc = Doc.setHighlight (if isNull (box color) || color = "" || color = "none" then None else Some color) editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set hyperlink on selection
    let setHyperlinkFor (editor: Editor) (target: string) =
        let newDoc = Doc.setHyperlink (if isNull (box target) || target = "" then None else Some target) editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set comment metadata on selection
    let setCommentFor (editor: Editor) (commentId: string) (commentText: string) (author: string) (date: string) =
        let commentIdOpt = if isNull (box commentId) || commentId = "" then None else Some commentId
        let commentTextOpt = if isNull (box commentText) || commentText = "" then None else Some commentText
        let authorOpt = if isNull (box author) || author = "" then None else Some author
        let dateOpt = if isNull (box date) || date = "" then None else Some date
        let newDoc = Doc.setComment commentIdOpt commentTextOpt authorOpt dateOpt editor.Doc
        commitDocumentUpdate editor newDoc

    /// Set footnote/endnote metadata on selection
    let setNoteFor (editor: Editor) (noteId: string) (noteType: string) (noteText: string) =
        let noteIdOpt = if isNull (box noteId) || noteId = "" then None else Some noteId
        let noteTypeOpt = if isNull (box noteType) || noteType = "" then None else Some noteType
        let noteTextOpt = if isNull (box noteText) || noteText = "" then None else Some noteText
        let newDoc = Doc.setNote noteIdOpt noteTypeOpt noteTextOpt editor.Doc
        commitDocumentUpdate editor newDoc

    let setBookmarkFor (editor: Editor) (bookmarkName: string) =
        let newDoc = Doc.setBookmark bookmarkName editor.Doc
        commitDocumentUpdate editor newDoc

    let setContentControlFor (editor: Editor) (controlId: string) (tag: string) (alias: string) =
        let newDoc = Doc.setContentControl controlId tag alias editor.Doc
        commitDocumentUpdate editor newDoc

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
        commitDocumentUpdate editor newDoc

    let removeBookmark (editor: Editor) (bookmarkId: string) =
        let newDoc = Doc.removeBookmark bookmarkId editor.Doc
        commitDocumentUpdate editor newDoc

    let renameBookmark (editor: Editor) (bookmarkId: string) (bookmarkName: string) =
        let newDoc = Doc.renameBookmark bookmarkId bookmarkName editor.Doc
        commitDocumentUpdate editor newDoc

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
        commitDocumentUpdate editor newDoc

    let removeContentControl (editor: Editor) (controlId: string) =
        let newDoc = Doc.removeContentControl controlId editor.Doc
        commitDocumentUpdate editor newDoc

    let updateContentControl (editor: Editor) (controlId: string) (newControlId: string) (tag: string) (alias: string) =
        let newDoc = Doc.updateContentControl controlId newControlId tag alias editor.Doc
        commitDocumentUpdate editor newDoc

    let setParagraphStyleFor (editor: Editor) (styleId: string) =
        let newDoc = Doc.setParagraphStyle styleId editor.Doc
        commitDocumentUpdate editor newDoc

    let setCharacterStyleFor (editor: Editor) (styleId: string) =
        let newDoc = Doc.setCharacterStyle styleId editor.Doc
        commitDocumentUpdate editor newDoc

    let getStyleCatalog (editor: Editor) =
        Docx.exportStylesToEditor editor.Doc.Styles

    let setRevisionFor (editor: Editor) (revisionType: string) (revisionId: string) (author: string) (date: string) =
        let newDoc = Doc.setRevision revisionType revisionId author date editor.Doc
        commitDocumentUpdate editor newDoc

    let acceptSelectedRevision (editor: Editor) =
        let newDoc = Doc.acceptSelectedRevision editor.Doc
        commitDocumentUpdate editor newDoc

    let rejectSelectedRevision (editor: Editor) =
        let newDoc = Doc.rejectSelectedRevision editor.Doc
        commitDocumentUpdate editor newDoc

    let acceptAllRevisions (editor: Editor) =
        let newDoc = Doc.acceptAllRevisions editor.Doc
        commitDocumentUpdate editor newDoc

    let rejectAllRevisions (editor: Editor) =
        let newDoc = Doc.rejectAllRevisions editor.Doc
        commitDocumentUpdate editor newDoc

    let selectNextRevision (editor: Editor) =
        let newDoc = Doc.selectNextRevision editor.Doc
        commitDocumentUpdate editor newDoc

    let selectPreviousRevision (editor: Editor) =
        let newDoc = Doc.selectPreviousRevision editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Set font size on selection
    let setFontSizeFor (editor: Editor) (size: float) =
        let newDoc = Doc.setFontSize size editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Set font family on selection
    let setFontFamilyFor (editor: Editor) (family: string) =
        let newDoc = Doc.setFontFamily family editor.Doc
        commitDocumentUpdate editor newDoc

    /// Apply Heading 1-3 to the current block
    let setHeadingLevel (editor: Editor) (level: int) =
        let newDoc = Doc.setHeadingLevel level editor.Doc
        commitDocumentUpdate editor newDoc

    /// Convert the current heading block back to a normal paragraph
    let clearHeading (editor: Editor) =
        let newDoc = Doc.clearHeading editor.Doc
        commitDocumentUpdate editor newDoc
    
    /// Select all text
    let selectAll (editor: Editor) =
        let newDoc = Doc.selectAll editor.Doc
        commitDocumentUpdate editor newDoc

    /// Align left
    let alignLeft (editor: Editor) =
        let newDoc = Doc.alignLeft editor.Doc
        commitDocumentUpdate editor newDoc

    /// Align center
    let alignCenter (editor: Editor) =
        let newDoc = Doc.alignCenter editor.Doc
        commitDocumentUpdate editor newDoc

    /// Align right
    let alignRight (editor: Editor) =
        let newDoc = Doc.alignRight editor.Doc
        commitDocumentUpdate editor newDoc

    /// Insert image at cursor
    let insertImage (editor: Editor) (src: string) (width: float) (height: float) (alt: string) =
        let widthOpt = if width > 0.0 then Some width else None
        let heightOpt = if height > 0.0 then Some height else None
        let altOpt = if isNull (box alt) || alt = "" then None else Some alt
        let newDoc = Doc.insertImageAtCursor src widthOpt heightOpt altOpt editor.Doc
        commitDocumentUpdate editor newDoc

    let resizeCurrentImage (editor: Editor) (width: float) (height: float) =
        let newDoc = Doc.resizeImageNearCursor width height editor.Doc
        commitDocumentUpdate editor newDoc

    /// Insert a table row after the row containing the current selection
    let insertTableRowAfter (editor: Editor) =
        let newDoc = Doc.insertTableRowAfterSelection editor.Doc
        commitDocumentUpdate editor newDoc

    /// Insert a table column after the column containing the current selection
    let insertTableColumnAfter (editor: Editor) =
        let newDoc = Doc.insertTableColumnAfterSelection editor.Doc
        commitDocumentUpdate editor newDoc

    /// Delete the table row containing the current selection
    let deleteTableRow (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableRow editor.Doc
        commitDocumentUpdate editor newDoc

    /// Delete the table column containing the current selection
    let deleteTableColumn (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableColumn editor.Doc
        commitDocumentUpdate editor newDoc

    let insertTable (editor: Editor) (rows: int) (columns: int) =
        let newDoc = Doc.insertTable rows columns editor.Doc
        commitDocumentUpdate editor newDoc

    let insertTextBox (editor: Editor) (text: string) =
        let newDoc = Doc.insertTextBox text editor.Doc
        commitDocumentUpdate editor newDoc

    let insertCaption (editor: Editor) (label: string) (captionText: string) =
        let newDoc = Doc.insertCaption label captionText editor.Doc
        commitDocumentUpdate editor newDoc

    let insertCrossReferenceToLatestCaption (editor: Editor) (label: string) =
        let beforeDoc = editor.Doc
        let newDoc = Doc.insertCrossReferenceToLatestCaption label editor.Doc
        commitDocumentUpdate editor newDoc
        not (obj.ReferenceEquals(newDoc, beforeDoc))

    let setCurrentTextBoxStyle (editor: Editor) (x: float) (y: float) (width: float) (height: float) (fillColor: string) (borderColor: string) (borderWidth: float) =
        let newDoc = Doc.setCurrentTextBoxStyle x y width height fillColor borderColor borderWidth editor.Doc
        commitDocumentUpdate editor newDoc

    let setSelectedTableCellStyle (editor: Editor) (backgroundColor: string) (cellWidth: float) (borderColor: string) (borderWidth: float) (verticalAlignment: string) (paddingTop: float) (paddingRight: float) (paddingBottom: float) (paddingLeft: float) =
        let optionalCellWidth = if cellWidth > 0.0 then Some cellWidth else None
        let optionalWidth = if borderWidth > 0.0 then Some borderWidth else None
        let optionalPadding value = if value >= 0.0 then Some value else None
        let newDoc = Doc.setSelectedTableCellStyle (optionalText backgroundColor) optionalCellWidth (optionalText borderColor) optionalWidth (optionalText verticalAlignment) (optionalPadding paddingTop) (optionalPadding paddingRight) (optionalPadding paddingBottom) (optionalPadding paddingLeft) editor.Doc
        commitDocumentUpdate editor newDoc

    let setSelectedTableCellStyleWithShading (editor: Editor) (backgroundColor: string) (shadingPattern: string) (shadingForegroundColor: string) (cellWidth: float) (borderColor: string) (borderWidth: float) (verticalAlignment: string) (paddingTop: float) (paddingRight: float) (paddingBottom: float) (paddingLeft: float) =
        let optionalCellWidth = if cellWidth > 0.0 then Some cellWidth else None
        let optionalWidth = if borderWidth > 0.0 then Some borderWidth else None
        let optionalPadding value = if value >= 0.0 then Some value else None
        let normalizedPattern = if isNull (box shadingPattern) then "" else shadingPattern.Trim()
        let background = optionalText backgroundColor
        let shading: Doc.Shading option =
            match normalizedPattern with
            | ""
            | "clear"
            | "none" -> None
            | pattern ->
                Some {
                    Pattern = Some pattern
                    ForegroundColor = optionalText shadingForegroundColor
                    BackgroundColor = background
                }
        let fillColor =
            if normalizedPattern = "none" then None
            elif shading.IsSome then None
            else background
        let newDoc = Doc.setSelectedTableCellStyleWithShading fillColor shading optionalCellWidth (optionalText borderColor) optionalWidth (optionalText verticalAlignment) (optionalPadding paddingTop) (optionalPadding paddingRight) (optionalPadding paddingBottom) (optionalPadding paddingLeft) editor.Doc
        commitDocumentUpdate editor newDoc

    let setCurrentTableCellMargins (editor: Editor) (paddingTop: float) (paddingRight: float) (paddingBottom: float) (paddingLeft: float) =
        let optionalPadding value = if value >= 0.0 then Some value else None
        let newDoc = Doc.setCurrentTableCellMargins (optionalPadding paddingTop) (optionalPadding paddingRight) (optionalPadding paddingBottom) (optionalPadding paddingLeft) editor.Doc
        commitDocumentUpdate editor newDoc

    let setCurrentTableCellMarginsAndSpacing (editor: Editor) (paddingTop: float) (paddingRight: float) (paddingBottom: float) (paddingLeft: float) (cellSpacing: float) =
        let optionalPadding value = if value >= 0.0 then Some value else None
        let optionalSpacing value = if value >= 0.0 then Some value else None
        let newDoc = Doc.setCurrentTableCellMarginsAndSpacing (optionalPadding paddingTop) (optionalPadding paddingRight) (optionalPadding paddingBottom) (optionalPadding paddingLeft) (optionalSpacing cellSpacing) editor.Doc
        commitDocumentUpdate editor newDoc

    /// Align justify
    let alignJustify (editor: Editor) =
        let newDoc = Doc.alignJustify editor.Doc
        commitDocumentUpdate editor newDoc

    let toggleBulletList (editor: Editor) =
        let newDoc = Doc.toggleBulletList editor.Doc
        commitDocumentUpdate editor newDoc

    let toggleNumberedList (editor: Editor) =
        let newDoc = Doc.toggleNumberedList editor.Doc
        commitDocumentUpdate editor newDoc

    let increaseIndent (editor: Editor) =
        let newDoc = Doc.increaseIndent editor.Doc
        commitDocumentUpdate editor newDoc

    let decreaseIndent (editor: Editor) =
        let newDoc = Doc.decreaseIndent editor.Doc
        commitDocumentUpdate editor newDoc

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
        commitDocumentUpdate editor newDoc

    let setSectionBreak (editor: Editor) (kind: string) =
        let sectionBreak =
            match kind with
            | "continuous" -> Some "continuous"
            | "none" -> None
            | _ -> Some "nextPage"
        let newDoc = Doc.setCurrentSectionBreak sectionBreak editor.Doc
        commitDocumentUpdate editor newDoc

    /// Import document from EditorParagraphs (Office.Formats format)
    /// Use with: Office.Formats processors.forEditorPackage() Paragraphs output
    let importFromDocx (editor: Editor) (editorParagraphs: obj array) =
        let newDoc = Docx.importFromEditor editorParagraphs
        loadDocument editor newDoc

    let importFromDocxWithHeaderFooters (editor: Editor) (editorParagraphs: obj array) (headerParagraphs: obj array) (footerParagraphs: obj array) =
        let newDoc = Docx.importFromEditorWithHeaderFooters editorParagraphs headerParagraphs footerParagraphs
        loadDocument editor newDoc

    let importFromDocxWithHeaderFooterParts (editor: Editor) (editorParagraphs: obj array) (headerParts: obj array) (footerParts: obj array) =
        let newDoc = Docx.importFromEditorWithHeaderFooterParts editorParagraphs headerParts footerParts
        loadDocument editor newDoc

    let importFromDocxWithPageSetupAndHeaderFooterParts (editor: Editor) (editorParagraphs: obj array) (pageSetup: obj) (headerParts: obj array) (footerParts: obj array) =
        let newDoc = Docx.importFromEditorWithPageSetupAndHeaderFooterParts editorParagraphs pageSetup headerParts footerParts
        loadDocument editor newDoc

    let importFromDocxPackage (editor: Editor) (package: obj) =
        let newDoc = Docx.importFromEditorPackage package
        loadDocument editor newDoc

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
            Columns = editor.Doc.PageSetup.Columns
            SectionStart = editor.Doc.PageSetup.SectionStart
        }
        let newDoc = Doc.setPageSetup setup editor.Doc
        commitDocumentUpdate editor newDoc

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
            Columns = editor.Doc.PageSetup.Columns
            SectionStart = editor.Doc.PageSetup.SectionStart
        }
        let newDoc = Doc.setPageSetup setup editor.Doc
        commitDocumentUpdate editor newDoc

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

    /// The page frames as shown on the canvas, top to bottom: each page's
    /// position and size (sections can differ in page size and orientation)
    /// and margins. Hosts size the canvas from the last frame's Top + Height
    /// and the widest frame.
    let getPageFrames (editor: Editor) =
        Ren.pageFrames editor.Handler.RenderContext editor.Doc
        |> List.map (fun page ->
            createObj [
                "Top" ==> page.Top
                "Width" ==> page.Width
                "Height" ==> page.Height
                "MarginTop" ==> page.MarginTop
                "MarginBottom" ==> page.MarginBottom
                "MarginLeft" ==> page.MarginLeft
                "MarginRight" ==> page.MarginRight
            ])
        |> List.toArray

    let getPageCount (editor: Editor) =
        Ren.pageCount editor.Handler.RenderContext editor.Doc

    /// Document position under a canvas point, from the layout that is on screen.
    let hitTestPosition (editor: Editor) (x: float) (y: float) =
        match Ren.hitTest editor.Handler.RenderContext editor.Doc x y with
        | Some (blockId, offset) -> createObj [ "BlockId" ==> blockId; "Offset" ==> offset ]
        | None -> null

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
        let nextParts =
            if System.String.IsNullOrWhiteSpace text then remaining
            else
                { Kind = kind; RelationshipId = None; Blocks = [Doc.Paragraph (paragraphFromText (sprintf "%s-%s" prefix kind) text)] } :: remaining
        if nextParts = parts then parts else nextParts

    let private upsertReferencedPart prefix relationshipId text parts =
        let remaining = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> part.RelationshipId <> Some relationshipId)
        let nextParts =
            if System.String.IsNullOrWhiteSpace text then remaining
            else
                { Kind = "default"; RelationshipId = Some relationshipId; Blocks = [Doc.Paragraph (paragraphFromText (sprintf "%s-%s" prefix relationshipId) text)] } :: remaining
        if nextParts = parts then parts else nextParts

    let private tableCell id text : Doc.TableCell =
        { Content = [paragraphFromText id text]
          NestedTables = []
          NestedTableAnchors = []
          ColSpan = 1
          BackgroundColor = None
          Shading = None
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
          Rows = [{ Cells = [tableCell (id + "-left") left; tableCell (id + "-right") right]; IsHeader = false; Height = None; HeightRule = None; CanSplit = None; CellSpacing = None; PaddingTop = None; PaddingRight = None; PaddingBottom = None; PaddingLeft = None; GridBefore = 0; GridAfter = 0; WidthBefore = None; WidthAfter = None; BordersOverride = None }]
          PreferredWidth = Some { Kind = Doc.PercentWidth; Value = Some 100.0 }
          Alignment = None
          LayoutMode = Some Doc.FixedTableLayout
          GridColumns = [{ Kind = Doc.PercentWidth; Value = Some 50.0 }; { Kind = Doc.PercentWidth; Value = Some 50.0 }]
          Borders = None
          BackgroundColor = None
          Shading = None
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
        let nextParts = if blocks.IsEmpty then otherParts else { Kind = "default"; RelationshipId = None; Blocks = blocks } :: otherParts
        if nextParts = parts then parts else nextParts

    let private isImageParagraphId id = function
        | Doc.Paragraph para -> para.Id = id
        | _ -> false

    let private upsertDefaultImage paragraphId src alt width height parts =
        let hasImage = not (System.String.IsNullOrWhiteSpace src)
        let defaultPart = parts |> List.tryFind (fun (part: Doc.HeaderFooterPart) -> part.Kind = "default" && part.RelationshipId.IsNone)
        let otherParts = parts |> List.filter (fun (part: Doc.HeaderFooterPart) -> not (part.Kind = "default" && part.RelationshipId.IsNone))
        let existingBlocks = defaultPart |> Option.map _.Blocks |> Option.defaultValue [] |> List.filter (isImageParagraphId paragraphId >> not)
        let blocks = if hasImage then existingBlocks @ [Doc.Paragraph (imageParagraph paragraphId src alt width height)] else existingBlocks
        let nextParts = if blocks.IsEmpty then otherParts else { Kind = "default"; RelationshipId = None; Blocks = blocks } :: otherParts
        if nextParts = parts then parts else nextParts

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
        let defaultHeaderParagraphs =
            let next = if System.String.IsNullOrWhiteSpace defaultHeader then [] else [paragraphFromText "header-default" defaultHeader]
            if next = editor.Doc.Headers then editor.Doc.Headers else next
        let defaultFooterParagraphs =
            let next = if System.String.IsNullOrWhiteSpace defaultFooter then [] else [paragraphFromText "footer-default" defaultFooter]
            if next = editor.Doc.Footers then editor.Doc.Footers else next
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
        let newDoc =
            if obj.ReferenceEquals(headerParts, editor.Doc.HeaderParts)
                && obj.ReferenceEquals(footerParts, editor.Doc.FooterParts)
                && obj.ReferenceEquals(defaultHeaderParagraphs, editor.Doc.Headers)
                && obj.ReferenceEquals(defaultFooterParagraphs, editor.Doc.Footers) then
                editor.Doc
            else {
                editor.Doc with
                    HeaderParts = headerParts
                    FooterParts = footerParts
                    Headers = defaultHeaderParagraphs
                    Footers = defaultFooterParagraphs
            }
        commitDocumentUpdate editor newDoc

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
        let headerParts = editor.Doc.HeaderParts |> upsertDefaultTable "default-header-table" defaultHeaderLeft defaultHeaderRight
        let footerParts = editor.Doc.FooterParts |> upsertDefaultTable "default-footer-table" defaultFooterLeft defaultFooterRight
        let newDoc =
            if obj.ReferenceEquals(headerParts, editor.Doc.HeaderParts)
                && obj.ReferenceEquals(footerParts, editor.Doc.FooterParts) then
                editor.Doc
            else {
                editor.Doc with
                    HeaderParts = headerParts
                    FooterParts = footerParts
            }
        commitDocumentUpdate editor newDoc

    let private parseFloat fallback (value: string) =
        match System.Double.TryParse(value: string) with
        | true, parsed -> parsed
        | false, _ -> fallback

    let setHeaderFooterImageText (editor: Editor) defaultHeaderSrc defaultHeaderAlt defaultHeaderWidth defaultHeaderHeight defaultFooterSrc defaultFooterAlt defaultFooterWidth defaultFooterHeight =
        let headerParts = editor.Doc.HeaderParts |> upsertDefaultImage "default-header-image" defaultHeaderSrc defaultHeaderAlt (parseFloat 96.0 defaultHeaderWidth) (parseFloat 48.0 defaultHeaderHeight)
        let footerParts = editor.Doc.FooterParts |> upsertDefaultImage "default-footer-image" defaultFooterSrc defaultFooterAlt (parseFloat 96.0 defaultFooterWidth) (parseFloat 48.0 defaultFooterHeight)
        let newDoc =
            if obj.ReferenceEquals(headerParts, editor.Doc.HeaderParts)
                && obj.ReferenceEquals(footerParts, editor.Doc.FooterParts) then
                editor.Doc
            else {
                editor.Doc with
                    HeaderParts = headerParts
                    FooterParts = footerParts
            }
        commitDocumentUpdate editor newDoc

    let setCurrentSectionHeaderFooterText (editor: Editor) headerText footerText =
        let headerRelationshipId = "section-header-current"
        let footerRelationshipId = "section-footer-current"
        let headerParts = upsertReferencedPart "header" headerRelationshipId headerText editor.Doc.HeaderParts
        let footerParts = upsertReferencedPart "footer" footerRelationshipId footerText editor.Doc.FooterParts
        let headerRefs = if System.String.IsNullOrWhiteSpace headerText then [] else ["default", headerRelationshipId]
        let footerRefs = if System.String.IsNullOrWhiteSpace footerText then [] else ["default", footerRelationshipId]
        let newDoc =
            if obj.ReferenceEquals(headerParts, editor.Doc.HeaderParts)
                && obj.ReferenceEquals(footerParts, editor.Doc.FooterParts) then
                editor.Doc
            else
                { editor.Doc with HeaderParts = headerParts; FooterParts = footerParts }
            |> Doc.setCurrentHeaderFooterReferences headerRefs footerRefs
        commitDocumentUpdate editor newDoc

    /// Export document to EditorParagraphs (Office.Formats format)
    /// Use with: Office.Formats WordWriter.writeFromEditor()
    let exportToDocx (editor: Editor) : obj array =
        Docx.exportToEditor editor.Doc

    let exportDocumentPackage (editor: Editor) =
        // One export for save and for callers: Docx.exportDocumentPackage knows the
        // document's styles and so leaves inherited formatting out of header/footer runs too.
        Docx.exportDocumentPackage editor.Doc

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
            "undo" ==> undo
            "redo" ==> redo
            "canUndo" ==> canUndo
            "canRedo" ==> canRedo
            "dispose" ==> dispose
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
            "getLayoutLines" ==> getLayoutLines
            "getVbaLayoutLines" ==> getVbaLayoutLines
            "getRangeInformation" ==> getRangeInformation
            "getSelectionInformation" ==> getSelectionInformation
            "getRangeBounds" ==> getRangeBounds
            "computeRangeStatistics" ==> computeRangeStatistics
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
            "setSelectedTableCellStyleWithShading" ==> setSelectedTableCellStyleWithShading
            "setCurrentTableCellMargins" ==> setCurrentTableCellMargins
            "setCurrentTableCellMarginsAndSpacing" ==> setCurrentTableCellMarginsAndSpacing
            "setSectionBreak" ==> setSectionBreak
            "importFromDocx" ==> importFromDocx
            "importFromDocxWithHeaderFooters" ==> importFromDocxWithHeaderFooters
            "importFromDocxWithHeaderFooterParts" ==> importFromDocxWithHeaderFooterParts
            "importFromDocxWithPageSetupAndHeaderFooterParts" ==> importFromDocxWithPageSetupAndHeaderFooterParts
            "importFromDocxPackage" ==> importFromDocxPackage
            "setPageSetup" ==> setPageSetup
            "setPageSetupOptions" ==> setPageSetupOptions
            "getPageSetup" ==> getPageSetup
            "getPageFrames" ==> getPageFrames
            "getPageCount" ==> getPageCount
            "hitTestPosition" ==> hitTestPosition
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
