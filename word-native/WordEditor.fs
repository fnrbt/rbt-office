module Composit.Ui.WordEditor

open Renderer
open Composit.Ui.Reactive
open Composit.Ui.Widget

open RedBat.Word.Document

module WordLayout = RedBat.Word.Layout
module WordPaint = RedBat.Word.Paint
module WordTextBuffer = RedBat.Word.TextBuffer

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

let private rendererFont () = Composit.Ui.Layout.font()

let private pointsToPixels points = points * 96.0 / 72.0

let private measureText (text: string) (style: RedBat.Word.TextBuffer.TextStyle) =
    let font = rendererFont()
    Renderer.Text.measureWidth font (pointsToPixels style.FontSize) text

let private parseHexColor (value: string) =
    let value =
        match value.Trim().ToLowerInvariant() with
        | "black" -> "#000000"
        | "blue" -> "#0000FF"
        | "cyan" -> "#00FFFF"
        | "green" -> "#00FF00"
        | "magenta" -> "#FF00FF"
        | "red" -> "#FF0000"
        | "yellow" -> "#FFFF00"
        | "white" -> "#FFFFFF"
        | "darkblue" -> "#000080"
        | "darkcyan" -> "#008080"
        | "darkgreen" -> "#008000"
        | "darkmagenta" -> "#800080"
        | "darkred" -> "#800000"
        | "darkyellow" -> "#808000"
        | "darkgray" -> "#808080"
        | "lightgray" -> "#C0C0C0"
        | _ -> value
    let parseByte (s: string) =
        match System.Byte.TryParse(s, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture) with
        | true, value -> Some value
        | false, _ -> None
    let colorFromParts r g b =
        match parseByte r, parseByte g, parseByte b with
        | Some r, Some g, Some b -> Color.rgb r g b
        | _ -> Color.black
    let colorFromPartsWithAlpha r g b a =
        match parseByte r, parseByte g, parseByte b, parseByte a with
        | Some r, Some g, Some b, Some a -> Color.rgba r g b a
        | _ -> Color.black
    let s = if value.StartsWith("#") then value.Substring(1) else value

    match s.Length with
    | 3 ->
        let expand (c: char) = string c + string c
        colorFromParts (expand s.[0]) (expand s.[1]) (expand s.[2])
    | 6 ->
        colorFromParts s.[0..1] s.[2..3] s.[4..5]
    | 8 ->
        colorFromPartsWithAlpha s.[0..1] s.[2..3] s.[4..5] s.[6..7]
    | _ when value.StartsWith("rgba(") ->
        let parts =
            value.Substring(5, value.Length - 6).Split(',')
            |> Array.map (fun p -> p.Trim())

        if parts.Length = 4 then
            let alpha = float parts.[3]
            Color.rgba (byte parts.[0]) (byte parts.[1]) (byte parts.[2]) (byte (max 0.0 (min 255.0 (alpha * 255.0))))
        else
            Color.black
    | _ -> Color.black

let private textFont (_style: WordPaint.TextPaintStyle) = rendererFont()

type WordEditorState = {
    Document: Signal<Document>
    HoveredImage: Signal<(BlockId * int) option>
    ScrollY: Signal<float>
    DragAnchor: Signal<Position option>
    WordDragAnchor: Signal<(Position * Position) option>
    ParagraphDragAnchor: Signal<(Position * Position) option>
    ResizingImage: Signal<(BlockId * int * float * float * float * float * Document) option>
    DraggingImage: Signal<(BlockId * int * float * float * Document) option>
    LastClick: Signal<(float * float * float * int) option>
    LastCursorRect: Signal<Rect option>
    PendingStyle: Signal<WordTextBuffer.TextStyle option>
    UndoStack: Signal<Document list>
    RedoStack: Signal<Document list>
    FocusId: int
}

let createState doc = {
    Document = signal doc
    HoveredImage = signal None
    ScrollY = signal 0.0
    DragAnchor = signal None
    WordDragAnchor = signal None
    ParagraphDragAnchor = signal None
    ResizingImage = signal None
    DraggingImage = signal None
    LastClick = signal None
    LastCursorRect = signal None
    PendingStyle = signal None
    UndoStack = signal []
    RedoStack = signal []
    FocusId = Focus.allocId()
}

let private layoutSettings width (doc: Document) : WordLayout.LayoutSettings =
    let sourceWidth = max 1.0 doc.PageSetup.PageWidth
    let scale = width / sourceWidth
    let lineHeight = 24.0
    let hasFooter = not doc.Footers.IsEmpty || not doc.FooterParts.IsEmpty
    let footerReserve =
        if hasFooter then doc.PageSetup.FooterDistance * scale + lineHeight * 1.5
        else doc.PageSetup.MarginBottom * scale
    {
        Width = width
        PageHeight = max 1.0 (doc.PageSetup.PageHeight * scale)
        FontSize = 16.0
        LineHeight = lineHeight
        PaddingLeft = doc.PageSetup.MarginLeft * scale
        PaddingTop = RedBat.Word.WordLayout.Page.bodyTopWithRunningHeader doc * scale
        MarginRight = doc.PageSetup.MarginRight * scale
        MarginBottom = max (doc.PageSetup.MarginBottom * scale) footerReserve
        InTableCell = false
        TableContinuationHeaderHeight = 0.0
    }

let hitTest width doc x y =
    WordLayout.hitTest (layoutSettings width doc) measureText doc x y

let documentHeight width doc =
    let settings = layoutSettings width doc
    let layouts = WordLayout.layoutDocument settings measureText doc
    match layouts with
    | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
    | _ ->
        layouts
        |> List.map (fun layout -> layout.Y + layout.Height)
        |> List.max
        |> fun bottom -> bottom + settings.PaddingTop

let setCursorAt width x y state =
    match hitTest width state.Document.Value x y with
    | None -> false
    | Some (blockId, offset) ->
        state.Document.Value <- setCursor blockId offset state.Document.Value
        true

let private positionAt width x y state =
    hitTest width state.Document.Value x y
    |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })

let private imageAt width x y state =
    WordLayout.hitTestForImage (layoutSettings width state.Document.Value) measureText state.Document.Value x y

let private isResizeHandle x y imageX imageY imageWidth imageHeight =
    let handleSize = 8.0
    let half = handleSize / 2.0
    let corners = [
        (imageX, imageY)
        (imageX + imageWidth, imageY)
        (imageX, imageY + imageHeight)
        (imageX + imageWidth, imageY + imageHeight)
    ]

    corners
    |> List.exists (fun (cx, cy) -> x >= cx - half && x <= cx + half && y >= cy - half && y <= cy + half)

let private isWordChar (c: char) = System.Char.IsLetterOrDigit(c) || c = '_'

let private wordRangeAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> None
    | Some para ->
        let text = WordTextBuffer.toString para.Buffer
        if text = "" then Some (position, position)
        else
            let index = max 0 (min (text.Length - 1) position.Offset)
            let isSelectedClass = isWordChar text.[index]
            let mutable startOffset = index
            let mutable endOffset = index + 1

            while startOffset > 0 && isWordChar text.[startOffset - 1] = isSelectedClass do
                startOffset <- startOffset - 1

            while endOffset < text.Length && isWordChar text.[endOffset] = isSelectedClass do
                endOffset <- endOffset + 1

            Some ({ position with Offset = startOffset }, { position with Offset = endOffset })

let private selectWordAt (position: Position) doc =
    match wordRangeAt position doc with
    | Some (startPosition, endPosition) -> setSelection startPosition endPosition doc
    | None -> doc

let private selectParagraphAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> doc
    | Some para ->
        let length = WordTextBuffer.length para.Buffer
        setSelection { position with Offset = 0 } { position with Offset = length } doc

let private paragraphRangeAt (position: Position) doc =
    match findParagraph doc position.BlockId with
    | None -> None
    | Some para ->
        let length = WordTextBuffer.length para.Buffer
        Some ({ position with Offset = 0 }, { position with Offset = length })

let private endOfParagraph blockId doc =
    match findParagraph doc blockId with
    | Some para -> { BlockId = blockId; Offset = WordTextBuffer.length para.Buffer }
    | None -> { BlockId = blockId; Offset = 0 }

let private imageMoveTarget width x y blockId sourceOffset doc =
    let target =
        match hitTest width doc x y with
        | Some (targetBlockId, targetOffset) -> { BlockId = targetBlockId; Offset = targetOffset }
        | None -> endOfParagraph blockId doc

    if target.BlockId = blockId && target.Offset <= sourceOffset then endOfParagraph blockId doc
    else target

let private editableParagraphs doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph p
        | Heading (_, p) -> [p]
        | Table table ->
            table.Rows
            |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content)))

let private documentStart doc : Position option =
    editableParagraphs doc
    |> List.tryHead
    |> Option.map (fun para -> { BlockId = para.Id; Offset = 0 })

let private documentEnd doc : Position option =
    editableParagraphs doc
    |> List.tryLast
    |> Option.map (fun para -> { BlockId = para.Id; Offset = WordTextBuffer.length para.Buffer })

let private moveToDocumentBoundary (boundary: Document -> Position option) doc =
    match boundary doc with
    | Some position -> setCursor position.BlockId position.Offset doc
    | None -> doc

let private extendToDocumentBoundary (boundary: Document -> Position option) doc =
    match doc.Selection, boundary doc with
    | Some selection, Some position -> { doc with Selection = Some { selection with Focus = position } }
    | _ -> doc

let private previousWordOffset (text: string) offset =
    let mutable index = max 0 (min text.Length offset)
    while index > 0 && not (isWordChar text.[index - 1]) do
        index <- index - 1
    while index > 0 && isWordChar text.[index - 1] do
        index <- index - 1
    index

let private nextWordOffset (text: string) offset =
    let mutable index = max 0 (min text.Length offset)
    while index < text.Length && not (isWordChar text.[index]) do
        index <- index + 1
    while index < text.Length && isWordChar text.[index] do
        index <- index + 1
    index

let private moveCursorByWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset
            setCursor selection.Focus.BlockId offset doc

let private extendSelectionByWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset
            { doc with Selection = Some { selection with Focus = { selection.Focus with Offset = offset } } }

let private deleteWord direction doc =
    match doc.Selection with
    | None -> doc
    | Some selection when not selection.IsCollapsed -> deleteSelection doc
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | None -> doc
        | Some para ->
            let text = WordTextBuffer.toString para.Buffer
            let offset =
                if direction < 0 then previousWordOffset text selection.Focus.Offset
                else nextWordOffset text selection.Focus.Offset

            if offset = selection.Focus.Offset then doc
            else
                let rangeStart = min offset selection.Focus.Offset
                let rangeEnd = max offset selection.Focus.Offset
                { doc with
                    Selection = Some {
                        Anchor = { selection.Focus with Offset = rangeStart }
                        Focus = { selection.Focus with Offset = rangeEnd }
                    } }
                |> deleteSelection

let private styleAtCursor doc =
    match doc.Selection with
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para ->
            let length = WordTextBuffer.length para.Buffer
            let offset = if selection.Focus.Offset > 0 && selection.Focus.Offset >= length then selection.Focus.Offset - 1 else selection.Focus.Offset
            WordTextBuffer.getStyleAt offset para.Buffer
        | None -> WordTextBuffer.defaultStyle
    | None -> WordTextBuffer.defaultStyle

let private togglePendingStyle transform state =
    let before = state.PendingStyle.Value |> Option.defaultValue (styleAtCursor state.Document.Value)
    state.PendingStyle.Value <- Some (transform before)
    Platform.requestRedraw()
    true

let private hasCollapsedSelection doc =
    match doc.Selection with
    | Some selection -> selection.IsCollapsed
    | None -> false

let private clickCount x y state =
    let now = Platform.now()
    let count =
        match state.LastClick.Value with
        | Some (lastTime, lastX, lastY, lastCount) ->
            let dx = x - lastX
            let dy = y - lastY
            if now > lastTime && now - lastTime <= 0.45 && dx * dx + dy * dy <= 36.0 then min 3 (lastCount + 1)
            else 1
        | None -> 1

    state.LastClick.Value <- Some (now, x, y, count)
    count

let private caretRect width state doc =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        let settings = layoutSettings width doc
        let layouts = WordLayout.layoutDocument settings measureText doc
        match layouts |> List.tryFind (fun layout -> layout.BlockId = selection.Focus.BlockId) with
        | None -> None
        | Some layout ->
            let line = WordLayout.lineForOffset layout selection.Focus.Offset
            let x = WordLayout.xForOffset measureText line selection.Focus.Offset
            let lineAlignmentOffset = WordLayout.alignmentOffset layout line
            Some {
                X = settings.PaddingLeft + x + lineAlignmentOffset
                Y = line.Y - state.ScrollY.Value
                Width = 2.0
                Height = line.Height
            }
    | _ -> None

let private updateCursorRect width state doc =
    let nextRect = if Focus.isFocused state.FocusId then caretRect width state doc else None
    if nextRect <> state.LastCursorRect.Value then
        match nextRect with
        | Some rect ->
            Platform.onTextInputFocus rect
            Platform.onCursorRectChange rect
        | None -> Platform.onTextInputBlur()

        state.LastCursorRect.Value <- nextRect

let private applyDocumentChange state transform =
    let before = state.Document.Value
    let after = transform before
    if after <> before then
        state.UndoStack.Value <- before :: state.UndoStack.Value
        state.RedoStack.Value <- []
        state.Document.Value <- after
        Platform.requestRedraw()
        true
    else
        false

let private applyDocumentChangeNoHistory state transform =
    let before = state.Document.Value
    let after = transform before
    if after <> before then
        state.Document.Value <- after
        Platform.requestRedraw()
        true
    else
        false

let private insertPlainText text doc =
    if System.String.IsNullOrEmpty text then doc
    else
        let normalized = text.Replace("\r\n", "\n").Replace("\r", "\n")
        let lines = normalized.Split([|'\n'|], System.StringSplitOptions.None)
        let mutable nextDoc = doc

        for i = 0 to lines.Length - 1 do
            nextDoc <- insertText lines.[i] nextDoc
            if i < lines.Length - 1 then
                nextDoc <- splitBlock nextDoc

        nextDoc

let private insertStyledText text style doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let docBeforeInsert = if selection.IsCollapsed then doc else deleteSelection doc
        match docBeforeInsert.Selection with
        | Some cursor ->
            match findParagraph docBeforeInsert cursor.Focus.BlockId with
            | Some para ->
                let startOffset = cursor.Focus.Offset
                let insertedBuffer = WordTextBuffer.insert startOffset text para.Buffer
                let styledBuffer = WordTextBuffer.applyStyle startOffset text.Length (fun _ -> style) insertedBuffer
                let nextDoc = replaceParagraph docBeforeInsert para.Id { para with Buffer = styledBuffer }
                { nextDoc with
                    Selection = Some {
                        Anchor = { cursor.Focus with Offset = startOffset + text.Length }
                        Focus = { cursor.Focus with Offset = startOffset + text.Length }
                    } }
            | None -> docBeforeInsert
        | None -> docBeforeInsert

let private insertPlainTextWithPendingStyle state text doc =
    match state.PendingStyle.Value with
    | None -> insertPlainText text doc
    | Some style when not (text.Contains("\n") || text.Contains("\r")) -> insertStyledText text style doc
    | Some _ -> insertPlainText text doc

let private autoListAtCursor doc =
    match doc.Selection with
    | Some selection when selection.IsCollapsed ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para ->
            let beforeCursor = WordTextBuffer.slice 0 selection.Focus.Offset para.Buffer
            let marker =
                match beforeCursor with
                | "-" | "*" -> Some (1, Bullet)
                | "1." -> Some (2, Numbered)
                | _ -> None

            match marker with
            | Some (markerLength, listStyle) ->
                let nextBuffer = WordTextBuffer.delete 0 markerLength para.Buffer
                let nextPara = {
                    para with
                        Buffer = nextBuffer
                        Metadata = { para.Metadata with ListStyle = listStyle; ListLevel = max 1 para.Metadata.ListLevel }
                }
                let nextDoc = replaceParagraph doc para.Id nextPara
                { nextDoc with
                    Selection = Some {
                        Anchor = { selection.Focus with Offset = 0 }
                        Focus = { selection.Focus with Offset = 0 }
                    } }
            | None -> insertPlainText " " doc
        | None -> insertPlainText " " doc
    | _ -> insertPlainText " " doc

let private clipRuns startOffset endOffset shift (buffer: WordTextBuffer.TextBuffer) =
    buffer.Runs
    |> List.choose (function
        | WordTextBuffer.Text run ->
            let runStart = run.Start
            let runEnd = run.Start + run.Length
            let clippedStart = max runStart startOffset
            let clippedEnd = min runEnd endOffset
            if clippedStart < clippedEnd then
                Some (WordTextBuffer.Text {
                    run with
                        Start = clippedStart + shift
                        Length = clippedEnd - clippedStart
                })
            else None
        | WordTextBuffer.Image image ->
            if image.Start >= startOffset && image.Start < endOffset then
                Some (WordTextBuffer.Image { image with Start = image.Start + shift })
            else None)

let private splitBufferPreservingRuns offset buffer =
    let length = WordTextBuffer.length buffer
    let beforeText = WordTextBuffer.slice 0 offset buffer
    let afterText = WordTextBuffer.slice offset (length - offset) buffer
    let beforeBuffer = WordTextBuffer.ofString beforeText
    let afterBuffer = WordTextBuffer.ofString afterText
    ({ beforeBuffer with Runs = clipRuns 0 offset 0 buffer },
     { afterBuffer with Runs = clipRuns offset length -offset buffer })

let private splitBlockPreservingStyles doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        match findBlock doc selection.Focus.BlockId, findBlockIndex doc selection.Focus.BlockId with
        | Some block, Some index ->
            let para = getParagraph block
            let beforeBuffer, afterBuffer = splitBufferPreservingRuns selection.Focus.Offset para.Buffer
            let newBlockId = "editor-block-" + System.Guid.NewGuid().ToString("N")
            let beforePara = { para with Buffer = beforeBuffer }
            let afterPara = { para with Id = newBlockId; Buffer = afterBuffer }
            let beforeBlock, afterBlock =
                match block with
                | Paragraph _ -> Paragraph beforePara, Paragraph afterPara
                | Heading (level, _) -> Heading (level, beforePara), Heading (level, afterPara)
                | Table _ -> block, block
            let blocks =
                doc.Blocks
                |> List.mapi (fun i current -> if i = index then [beforeBlock; afterBlock] else [current])
                |> List.concat
            { doc with
                Blocks = blocks
                Selection = Some {
                    Anchor = { BlockId = newBlockId; Offset = 0 }
                    Focus = { BlockId = newBlockId; Offset = 0 }
                } }
        | _ -> splitBlock doc

let private autoScrollForPointer viewportHeight contentHeight y state =
    let margin = 32.0
    let step = 48.0
    let top = state.ScrollY.Value
    let bottom = top + viewportHeight
    let maxScrollY = max 0.0 (contentHeight - viewportHeight)
    let nextScrollY =
        if y < top + margin then max 0.0 (top - step)
        elif y > bottom - margin then min maxScrollY (top + step)
        else top

    if nextScrollY <> top then
        state.ScrollY.Value <- nextScrollY
        Platform.requestRedraw()

let private enterAtCursor doc =
    match doc.Selection with
    | Some selection ->
        match findParagraph doc selection.Focus.BlockId with
        | Some para when WordTextBuffer.length para.Buffer = 0 && para.Metadata.ListStyle <> NoList -> decreaseIndent doc
        | _ -> splitBlockPreservingStyles doc
    | None -> doc

let private undo state =
    match state.UndoStack.Value with
    | previous :: rest ->
        state.RedoStack.Value <- state.Document.Value :: state.RedoStack.Value
        state.UndoStack.Value <- rest
        state.Document.Value <- previous
        Platform.requestRedraw()
        true
    | [] -> true

let private redo state =
    match state.RedoStack.Value with
    | next :: rest ->
        state.UndoStack.Value <- state.Document.Value :: state.UndoStack.Value
        state.RedoStack.Value <- rest
        state.Document.Value <- next
        Platform.requestRedraw()
        true
    | [] -> true

let private cancelSelection state =
    state.DragAnchor.Value <- None
    state.WordDragAnchor.Value <- None
    state.ParagraphDragAnchor.Value <- None
    state.ResizingImage.Value <- None
    state.DraggingImage.Value <- None

    match state.Document.Value.Selection with
    | Some selection when not selection.IsCollapsed ->
        state.Document.Value <- setCursor selection.Focus.BlockId selection.Focus.Offset state.Document.Value
        Platform.requestRedraw()
        true
    | _ -> true

let private extendSelectionBy move doc =
    match doc.Selection with
    | None -> doc
    | Some selection ->
        let moved = move doc
        match moved.Selection with
        | Some movedSelection -> { moved with Selection = Some { selection with Focus = movedSelection.Focus } }
        | None -> doc

let private caretDocumentPoint width doc =
    match doc.Selection with
    | Some selection ->
        let settings = layoutSettings width doc
        let layouts = WordLayout.layoutDocument settings measureText doc
        match layouts |> List.tryFind (fun layout -> layout.BlockId = selection.Focus.BlockId) with
        | Some layout ->
            let line = WordLayout.lineForOffset layout selection.Focus.Offset
            let x = WordLayout.xForOffset measureText line selection.Focus.Offset
            let lineAlignmentOffset = WordLayout.alignmentOffset layout line
            Some (settings.PaddingLeft + x + lineAlignmentOffset, line.Y + line.Height / 2.0)
        | None -> None
    | None -> None

let private movePage width height direction extendSelection state =
    let doc = state.Document.Value
    let amount = max 24.0 (height * 0.8)
    let targetPosition =
        match caretDocumentPoint width doc with
        | Some (x, y) ->
            let targetY = y + amount * float direction
            hitTest width doc x targetY
            |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })
        | None -> None

    let nextDoc =
        match doc.Selection, targetPosition with
        | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
        | _, Some position -> setCursor position.BlockId position.Offset doc
        | Some selection, None when direction < 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentStart doc |> Option.defaultValue selection.Focus } }
        | Some selection, None when direction > 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentEnd doc |> Option.defaultValue selection.Focus } }
        | _, None when direction < 0 -> moveToDocumentBoundary documentStart doc
        | _, None -> moveToDocumentBoundary documentEnd doc

    state.ScrollY.Value <- max 0.0 (state.ScrollY.Value + amount * float direction)
    if nextDoc <> doc then state.Document.Value <- nextDoc
    Platform.requestRedraw()
    true

let private moveVisualLine width direction extendSelection doc =
    let targetPosition =
        match caretDocumentPoint width doc with
        | Some (x, y) ->
            let settings = layoutSettings width doc
            let targetY = y + settings.LineHeight * float direction
            hitTest width doc x targetY
            |> Option.map (fun (blockId, offset) -> { BlockId = blockId; Offset = offset })
        | None -> None

    match doc.Selection, targetPosition with
    | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
    | _, Some position -> setCursor position.BlockId position.Offset doc
    | Some selection, None when direction < 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentStart doc |> Option.defaultValue selection.Focus } }
    | Some selection, None when direction > 0 && extendSelection -> { doc with Selection = Some { selection with Focus = documentEnd doc |> Option.defaultValue selection.Focus } }
    | _, None when direction < 0 -> moveToDocumentBoundary documentStart doc
    | _, None -> moveToDocumentBoundary documentEnd doc

let private moveToVisualLineBoundary width toEnd extendSelection doc =
    let visualLineForBoundary toEnd offset (layout: WordLayout.BlockLayout) =
        if toEnd then
            layout.Lines
            |> List.tryFind (fun line -> offset >= line.StartOffset && offset < line.EndOffset)
            |> Option.orElseWith (fun () -> layout.Lines |> List.tryFind (fun line -> offset = line.EndOffset))
            |> Option.defaultValue (WordLayout.lineForOffset layout offset)
        else
            layout.Lines
            |> List.tryFind (fun line -> offset = line.StartOffset)
            |> Option.orElseWith (fun () -> layout.Lines |> List.tryFind (fun line -> offset > line.StartOffset && offset <= line.EndOffset))
            |> Option.defaultValue (WordLayout.lineForOffset layout offset)

    let targetPosition =
        match doc.Selection with
        | Some selection ->
            let settings = layoutSettings width doc
            let layouts = WordLayout.layoutDocument settings measureText doc
            match layouts |> List.tryFind (fun layout -> layout.BlockId = selection.Focus.BlockId) with
            | Some layout ->
                let line = visualLineForBoundary toEnd selection.Focus.Offset layout
                Some { BlockId = selection.Focus.BlockId; Offset = if toEnd then line.EndOffset else line.StartOffset }
            | None -> None
        | None -> None

    match doc.Selection, targetPosition with
    | Some selection, Some position when extendSelection -> { doc with Selection = Some { selection with Focus = position } }
    | _, Some position -> setCursor position.BlockId position.Offset doc
    | _ -> doc

let private handleKey width height state (ev: KeyEvent) =
    match ev with
    | { Key = "Escape" } -> cancelSelection state
    | { Ctrl = true; Key = key } when key = "a" || key = "A" -> applyDocumentChangeNoHistory state selectAll
    | { Ctrl = true; Shift = false; Key = key } when key = "z" || key = "Z" -> undo state
    | { Ctrl = true; Key = key } when key = "y" || key = "Y" -> redo state
    | { Ctrl = true; Shift = true; Key = key } when key = "z" || key = "Z" -> redo state
    | { Ctrl = true; Key = key } when key = "c" || key = "C" ->
        let text = getSelectionText state.Document.Value
        if text <> "" then Platform.clipboard.SetText text
        true
    | { Ctrl = true; Key = key } when key = "x" || key = "X" ->
        let text = getSelectionText state.Document.Value
        if text <> "" then
            Platform.clipboard.SetText text
            applyDocumentChange state deleteSelection
        else
            true
    | { Ctrl = true; Key = key } when key = "v" || key = "V" ->
        match Platform.clipboard.GetText() with
        | Some text when text <> "" -> applyDocumentChange state (insertPlainTextWithPendingStyle state text)
        | _ -> true
    | { Ctrl = true; Key = key } when key = "b" || key = "B" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Bold = not style.Bold }) state
        else applyDocumentChange state toggleBold
    | { Ctrl = true; Key = key } when key = "i" || key = "I" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Italic = not style.Italic }) state
        else applyDocumentChange state toggleItalic
    | { Ctrl = true; Key = key } when key = "u" || key = "U" ->
        if hasCollapsedSelection state.Document.Value then togglePendingStyle (fun style -> { style with Underline = not style.Underline }) state
        else applyDocumentChange state toggleUnderline
    | { Ctrl = true; Shift = true; Key = "8" } -> applyDocumentChange state toggleBulletList
    | { Ctrl = true; Shift = true; Key = "7" } -> applyDocumentChange state toggleNumberedList
    | { Ctrl = true; Shift = true; Key = key } when key = "l" || key = "L" -> applyDocumentChange state alignLeft
    | { Ctrl = true; Shift = true; Key = key } when key = "e" || key = "E" -> applyDocumentChange state alignCenter
    | { Ctrl = true; Shift = true; Key = key } when key = "r" || key = "R" -> applyDocumentChange state alignRight
    | { Ctrl = true; Shift = true; Key = key } when key = "j" || key = "J" -> applyDocumentChange state alignJustify
    | { Ctrl = true; Shift = true; Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state (extendSelectionByWord -1)
    | { Ctrl = true; Shift = true; Key = "ArrowRight" } -> applyDocumentChangeNoHistory state (extendSelectionByWord 1)
    | { Ctrl = true; Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state (moveCursorByWord -1)
    | { Ctrl = true; Key = "ArrowRight" } -> applyDocumentChangeNoHistory state (moveCursorByWord 1)
    | { Ctrl = true; Key = "Backspace" } -> applyDocumentChange state (deleteWord -1)
    | { Ctrl = true; Key = "Delete" } -> applyDocumentChange state (deleteWord 1)
    | { Ctrl = true; Shift = true; Key = "Home" } -> applyDocumentChangeNoHistory state (extendToDocumentBoundary documentStart)
    | { Ctrl = true; Shift = true; Key = "End" } -> applyDocumentChangeNoHistory state (extendToDocumentBoundary documentEnd)
    | { Ctrl = true; Key = "Home" } -> applyDocumentChangeNoHistory state (moveToDocumentBoundary documentStart)
    | { Ctrl = true; Key = "End" } -> applyDocumentChangeNoHistory state (moveToDocumentBoundary documentEnd)
    | { Key = "Backspace" } -> applyDocumentChange state deleteBackward
    | { Key = "Delete" } -> applyDocumentChange state deleteForward
    | { Key = "ArrowLeft"; Shift = true } -> applyDocumentChangeNoHistory state extendSelectionLeft
    | { Key = "ArrowRight"; Shift = true } -> applyDocumentChangeNoHistory state extendSelectionRight
    | { Key = "ArrowUp"; Shift = true } -> applyDocumentChangeNoHistory state (moveVisualLine width -1 true)
    | { Key = "ArrowDown"; Shift = true } -> applyDocumentChangeNoHistory state (moveVisualLine width 1 true)
    | { Key = "Home"; Shift = true } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width false true)
    | { Key = "End"; Shift = true } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width true true)
    | { Key = "PageUp"; Shift = true } -> movePage width height -1 true state
    | { Key = "PageDown"; Shift = true } -> movePage width height 1 true state
    | { Key = "PageUp" } -> movePage width height -1 false state
    | { Key = "PageDown" } -> movePage width height 1 false state
    | { Key = "ArrowLeft" } -> applyDocumentChangeNoHistory state moveCursorLeft
    | { Key = "ArrowRight" } -> applyDocumentChangeNoHistory state moveCursorRight
    | { Key = "ArrowUp" } -> applyDocumentChangeNoHistory state (moveVisualLine width -1 false)
    | { Key = "ArrowDown" } -> applyDocumentChangeNoHistory state (moveVisualLine width 1 false)
    | { Key = "Home" } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width false false)
    | { Key = "End" } -> applyDocumentChangeNoHistory state (moveToVisualLineBoundary width true false)
    | { Key = "Enter" } -> applyDocumentChange state enterAtCursor
    | { Key = "Tab"; Shift = true } -> applyDocumentChange state decreaseIndent
    | { Key = "Tab" } -> applyDocumentChange state increaseIndent
    | { Ctrl = false; Alt = false; Key = " " } -> applyDocumentChange state autoListAtCursor
    | { Ctrl = false; Alt = false; Key = key } when key.Length = 1 -> applyDocumentChange state (insertPlainTextWithPendingStyle state key)
    | _ -> false

let private toRendererCommand (resolveImage: string -> Framebuffer option) = function
    | WordPaint.DrawText (text, style, x, baseline, _) ->
        let font = textFont style
        [ Cmd.DrawGlyphs(vec2 x baseline, Renderer.Text.layout font (pointsToPixels style.FontSize) text, Solid (parseHexColor style.Color)) ]

    | WordPaint.DrawUnderline (color, x, y, width) ->
        [ Cmd.Line(vec2 x y, vec2 (x + width) y, parseHexColor color, px 1.0) ]

    | WordPaint.DrawImage (src, x, y, width, height, _, _, _, _) ->
        match resolveImage src with
        | Some fb ->
            [ Cmd.DrawImage(fb, rect 0.0 0.0 (float fb.W) (float fb.H), rect x y width height, true) ]
        | None ->
            // The browser path can draw data URLs directly. Widgets need a decoded framebuffer.
            [ Cmd.StrokeRect(rect x y width height, Color.rgb 180uy 180uy 180uy, px 1.0) ]

    | WordPaint.FillRect (color, x, y, width, height) ->
        [ Cmd.FillRect(rect x y width height, parseHexColor color) ]

    | WordPaint.StrokeRect (color, lineWidth, x, y, width, height) ->
        [ Cmd.StrokeRect(rect x y width height, parseHexColor color, px lineWidth) ]

    | WordPaint.StrokeLine (color, lineWidth, dashStyle, dashOffset, x1, y1, x2, y2) ->
        let dashArray =
            match dashStyle with
            | 1 -> [px lineWidth; px lineWidth]
            | 2 -> [px (lineWidth * 3.0); px (lineWidth * 1.5)]
            | 3 -> [px (lineWidth * 3.0); px (lineWidth * 1.5); px lineWidth; px (lineWidth * 1.5)]
            | _ -> []
        let style = {
            Width = px lineWidth
            Cap = Butt
            Join = Miter
            Paint = Solid (parseHexColor color)
            DashArray = dashArray
            DashOffset = px dashOffset
            MiterLimit = 4.0
        }
        [ Cmd.StrokePath([MoveTo (vec2 x1 y1); LineTo (vec2 x2 y2)], style) ]

let commands (width: float) (height: float) (resolveImage: string -> Framebuffer option) (hoveredImage: (BlockId * int) option) (doc: Document) =
    let settings = layoutSettings width doc
    let layouts, contentCommands, _ = WordPaint.collectDocumentCommands settings "Calibri" measureText (WordLayout.rangeMeasurerFromTextMeasurer measureText) doc

    let overlayCommands =
        WordPaint.selectionCommands settings measureText doc layouts @
        WordPaint.hoverCommands settings layouts hoveredImage

    contentCommands @ overlayCommands
    |> List.collect (toRendererCommand resolveImage)

let widget width height resolveImage hoveredImage doc =
    canvas width height (fun w h -> commands w h resolveImage hoveredImage doc)

let editor width height doc =
    widget width height (fun _ -> None) None doc

let private interactiveWidgetInViewport width height viewportHeight resolveImage state =
    dynamic (fun () ->
        let doc = state.Document.Value
        updateCursorRect width state doc
        let hoveredImage = state.HoveredImage.Value
        let handleText = function
            | TextCommit text ->
                applyDocumentChange state (insertPlainTextWithPendingStyle state text)
            | CompositionStart | CompositionUpdate _ | CompositionEnd ->
                // Hosts own native IME/preedit UI; committed text arrives as TextCommit.
                true
            | TextKeyPress ev -> handleKey width height state ev

        let handlePointer = function
            | PointerDown (x, y) ->
                Focus.request state.FocusId
                match imageAt width x y state with
                | Some (blockId, offset, imageWidth, imageHeight, imageX, imageY) when isResizeHandle x y imageX imageY imageWidth imageHeight ->
                    state.ResizingImage.Value <- Some (blockId, offset, x, y, imageWidth, imageHeight, state.Document.Value)
                    state.HoveredImage.Value <- Some (blockId, offset)
                    Platform.requestRedraw()
                | Some (blockId, offset, _, _, _, _) ->
                    state.DraggingImage.Value <- Some (blockId, offset, x, y, state.Document.Value)
                    state.HoveredImage.Value <- Some (blockId, offset)
                    Platform.requestRedraw()
                | _ ->
                    let count = clickCount x y state
                    match positionAt width x y state with
                    | Some pos ->
                        if count >= 3 then
                            state.ParagraphDragAnchor.Value <- paragraphRangeAt pos state.Document.Value
                            state.WordDragAnchor.Value <- None
                            applyDocumentChangeNoHistory state (selectParagraphAt pos) |> ignore
                        elif count = 2 then
                            state.WordDragAnchor.Value <- wordRangeAt pos state.Document.Value
                            state.ParagraphDragAnchor.Value <- None
                            applyDocumentChangeNoHistory state (selectWordAt pos) |> ignore
                        else
                            state.WordDragAnchor.Value <- None
                            state.ParagraphDragAnchor.Value <- None
                            state.DragAnchor.Value <- Some pos
                            applyDocumentChangeNoHistory state (setCursor pos.BlockId pos.Offset) |> ignore
                    | None -> ()
            | PointerMove (x, y) ->
                if state.DragAnchor.Value.IsSome || state.WordDragAnchor.Value.IsSome || state.ParagraphDragAnchor.Value.IsSome || state.DraggingImage.Value.IsSome then
                    autoScrollForPointer viewportHeight height y state

                match state.ResizingImage.Value with
                | Some (blockId, offset, startX, startY, startWidth, startHeight, _) ->
                    let deltaX = x - startX
                    let deltaY = y - startY
                    let delta = if abs deltaX > abs deltaY then deltaX else deltaY
                    let newWidth = max 20.0 (startWidth + delta)
                    let newHeight = max 20.0 (startHeight + delta)
                    applyDocumentChangeNoHistory state (resizeImage blockId offset newWidth newHeight) |> ignore
                | None ->
                    match state.DraggingImage.Value with
                    | Some _ -> ()
                    | None ->
                        match state.ParagraphDragAnchor.Value, positionAt width x y state with
                        | Some (anchorStart, anchorEnd), Some focus ->
                            match paragraphRangeAt focus state.Document.Value with
                            | Some (focusStart, focusEnd) ->
                                let nextDoc =
                                    let normalized = getNormalizedRange { Anchor = anchorStart; Focus = focusEnd } state.Document.Value
                                    let startPosition, _, _ = normalized
                                    if startPosition.BlockId = anchorStart.BlockId && startPosition.Offset = anchorStart.Offset then
                                        setSelection anchorStart focusEnd state.Document.Value
                                    else
                                        setSelection anchorEnd focusStart state.Document.Value
                                applyDocumentChangeNoHistory state (fun _ -> nextDoc) |> ignore
                            | None -> ()
                        | None, _ ->
                            match state.WordDragAnchor.Value, positionAt width x y state with
                            | Some (anchorStart, anchorEnd), Some focus ->
                                match wordRangeAt focus state.Document.Value with
                                | Some (focusStart, focusEnd) ->
                                    let nextDoc =
                                        let normalized = getNormalizedRange { Anchor = anchorStart; Focus = focusEnd } state.Document.Value
                                        let startPosition, _, _ = normalized
                                        if startPosition.BlockId = anchorStart.BlockId && startPosition.Offset = anchorStart.Offset then
                                            setSelection anchorStart focusEnd state.Document.Value
                                        else
                                            setSelection anchorEnd focusStart state.Document.Value
                                    applyDocumentChangeNoHistory state (fun _ -> nextDoc) |> ignore
                                | None -> ()
                            | None, _ ->
                                match state.DragAnchor.Value, positionAt width x y state with
                                | Some anchor, Some focus ->
                                    applyDocumentChangeNoHistory state (setSelection anchor focus) |> ignore
                                | _ -> ()
                            | _, None -> ()
                        | _, None -> ()
                        let nextHover =
                            imageAt width x y state
                            |> Option.map (fun (blockId, offset, _, _, _, _) -> (blockId, offset))
                        if nextHover <> state.HoveredImage.Value then
                            state.HoveredImage.Value <- nextHover
                            Platform.requestRedraw()
            | PointerUp (x, y) ->
                match state.DraggingImage.Value with
                | Some (blockId, offset, startX, startY, originalDocument) ->
                    let dx = x - startX
                    let dy = y - startY
                    if abs dx > 3.0 || abs dy > 3.0 then
                        let target = imageMoveTarget width x y blockId offset state.Document.Value
                        let nextDocument = moveImage blockId offset target.BlockId target.Offset state.Document.Value
                        if nextDocument <> state.Document.Value then
                            state.UndoStack.Value <- originalDocument :: state.UndoStack.Value
                            state.RedoStack.Value <- []
                            state.Document.Value <- nextDocument
                            Platform.requestRedraw()
                | None -> ()

                match state.ResizingImage.Value with
                | Some (_, _, _, _, _, _, originalDocument) when state.Document.Value <> originalDocument ->
                    state.UndoStack.Value <- originalDocument :: state.UndoStack.Value
                    state.RedoStack.Value <- []
                | _ -> ()
                state.DragAnchor.Value <- None
                state.WordDragAnchor.Value <- None
                state.ParagraphDragAnchor.Value <- None
                state.ResizingImage.Value <- None
                state.DraggingImage.Value <- None

        Focusable(state.FocusId, onKey (handleKey width height state) (onText handleText (pointerRegion handlePointer (widget width height resolveImage hoveredImage doc)))))

let interactiveWidget width height resolveImage state =
    interactiveWidgetInViewport width height height resolveImage state

let interactiveEditor width height doc =
    interactiveWidget width height (fun _ -> None) (createState doc)

let scrollWidget width height resolveImage state =
    dynamic (fun () ->
        let contentHeight = max height (documentHeight width state.Document.Value)
        scrollWith false true (signal 0.0) state.ScrollY (interactiveWidgetInViewport width contentHeight height resolveImage state))

let scrollEditor width height doc =
    scrollWidget width height (fun _ -> None) (createState doc)
