module RedBat.Word.Input

let private sameListOrEqual left right =
    obj.ReferenceEquals(left, right) || left = right

let private documentsEqualForChange (left: RedBat.Word.Document.Document) (right: RedBat.Word.Document.Document) =
    obj.ReferenceEquals(left, right)
    || (left.Selection = right.Selection
        && left.PageSetup = right.PageSetup
        && sameListOrEqual left.Blocks right.Blocks
        && sameListOrEqual left.Styles right.Styles
        && sameListOrEqual left.Headers right.Headers
        && sameListOrEqual left.Footers right.Footers
        && sameListOrEqual left.HeaderParts right.HeaderParts
        && sameListOrEqual left.FooterParts right.FooterParts)

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
open Browser.Dom
open RedBat.Word.Renderer
open RedBat.Word.Commands
open RedBat.Word.Document

module Doc = RedBat.Word.Document
module Cmd = RedBat.Word.Commands
module Clipboard = RedBat.Word.Clipboard

[<Emit("$0.style.cursor = $1")>]
let setCursor (elem: HTMLCanvasElement) (cursor: string) : unit = jsNative

let private keyEquals (expected: char) (key: string) =
    if key.Length <> 1 then
        false
    else
        let actual = key.[0]
        actual = expected || System.Char.ToUpperInvariant actual = System.Char.ToUpperInvariant expected

/// Resize state for images
type ResizeState = {
    BlockId: string
    Offset: int
    InitialX: float
    InitialY: float
    InitialWidth: float
    InitialHeight: float
}

type DragImageState = {
    BlockId: string
    Offset: int
    InitialX: float
    InitialY: float
}

/// Input handler state
type InputHandler = {
    Canvas: HTMLCanvasElement
    RenderContext: RenderContext
    mutable Document: Doc.Document
    OnChange: Doc.Document -> unit
    mutable IsMouseDown: bool
    mutable SelectionAnchor: Doc.Position option
    mutable History: CommandHistory
    mutable ResizingImage: ResizeState option
    mutable DraggingImage: DragImageState option
    mutable HoveredImage: (string * int) option  // (blockId, offset) of hovered image
    mutable TrackChanges: bool
    mutable TrackChangesAuthor: string
    mutable TrackChangesDate: string
}

/// Create input handler
let create (canvas: HTMLCanvasElement) (initialDoc: Doc.Document) (onChange: Doc.Document -> unit) =
    let ctx = createContext canvas
    let handler = {
        Canvas = canvas
        RenderContext = ctx
        Document = initialDoc
        OnChange = onChange
        IsMouseDown = false
        SelectionAnchor = None
        History = Cmd.empty
        ResizingImage = None
        DraggingImage = None
        HoveredImage = None
        TrackChanges = false
        TrackChangesAuthor = "RedBat Word"
        TrackChangesDate = ""
    }

    let insertTextForMode text doc =
        if handler.TrackChanges then Doc.insertTrackedText handler.TrackChangesAuthor handler.TrackChangesDate text doc else Doc.insertText text doc

    let deleteSelectionForMode doc =
        if handler.TrackChanges then Doc.trackDeleteSelection handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteSelection doc

    let deleteBackwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteBackward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteBackward doc

    let deleteForwardForMode doc =
        if handler.TrackChanges then Doc.trackDeleteForward handler.TrackChangesAuthor handler.TrackChangesDate doc else Doc.deleteForward doc

    let isResizeHandle x y imageX imageY width height =
        let handle = 10.0
        let nearLeft = abs (x - imageX) <= handle
        let nearRight = abs (x - (imageX + width)) <= handle
        let nearTop = abs (y - imageY) <= handle
        let nearBottom = abs (y - (imageY + height)) <= handle
        (nearLeft || nearRight) && (nearTop || nearBottom)

    let applyPointerDocumentChange newDoc =
        if not (documentsEqualForChange newDoc handler.Document) then
            handler.Document <- newDoc
            renderWithHover handler.RenderContext newDoc handler.HoveredImage
            handler.OnChange newDoc

    // Make canvas focusable
    canvas.tabIndex <- 0

    // Handle keyboard input
    canvas.addEventListener("keydown", fun (e: Event) ->
        let ke = e :?> KeyboardEvent
        let mutable newDoc = handler.Document
        let mutable newHistory = handler.History
        let mutable useCommand = false

        let commandKey = ke.ctrlKey || ke.metaKey

        match ke.key with
        | key when commandKey && keyEquals 'z' key ->
            // Undo/Redo
            if ke.shiftKey then
                // Redo (Ctrl+Shift+Z)
                let (doc', hist') = Cmd.redo handler.Document handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            else
                // Undo (Ctrl+Z)
                let (doc', hist') = Cmd.undo handler.Document handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'a' key ->
            newDoc <- Doc.selectAll handler.Document
            e.preventDefault()
        | key when commandKey && keyEquals 'b' key ->
            let beforeDoc = handler.Document
            let afterDoc = Doc.toggleBold handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "bold") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'i' key ->
            let beforeDoc = handler.Document
            let afterDoc = Doc.toggleItalic handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "italic") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'u' key ->
            let beforeDoc = handler.Document
            let afterDoc = Doc.toggleUnderline handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = FormatTextCommand(beforeDoc, afterDoc, "underline") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "1" when commandKey ->
            newDoc <- Doc.setHeadingLevel 1 handler.Document
            e.preventDefault()
        | "2" when commandKey ->
            newDoc <- Doc.setHeadingLevel 2 handler.Document
            e.preventDefault()
        | "3" when commandKey ->
            newDoc <- Doc.setHeadingLevel 3 handler.Document
            e.preventDefault()
        | key when commandKey && keyEquals 'c' key ->
            // Copy (Ctrl+C)
            let text = Doc.getSelectionText handler.Document
            if text <> "" then
                Clipboard.copyText text |> ignore
            e.preventDefault()
        | key when commandKey && keyEquals 'x' key ->
            // Cut (Ctrl+X)
            let text = Doc.getSelectionText handler.Document
            if text <> "" then
                Clipboard.copyText text |> ignore
                // Delete selection
                let beforeDoc = handler.Document
                let afterDoc = deleteSelectionForMode handler.Document
                if not (documentsEqualForChange afterDoc beforeDoc) then
                    let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Cut") :> ICommand
                    let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                    newDoc <- doc'
                    newHistory <- hist'
                    useCommand <- true
            e.preventDefault()
        | key when commandKey && keyEquals 'v' key ->
            // Paste (Ctrl+V)
            e.preventDefault()
            Clipboard.pasteText().``then``(fun text ->
                if text <> "" then
                    let beforeDoc = handler.Document
                    let mutable doc' = beforeDoc

                    // Delete selection first if any
                    let hasSelection =
                        match doc'.Selection with
                        | Some sel when sel.Anchor <> sel.Focus -> true
                        | _ -> false

                    if hasSelection then
                        doc' <- deleteSelectionForMode doc'

                    let mutable segmentStart = 0
                    let mutable index = 0
                    let mutable sawLineBreak = false
                    let insertSegment segmentEnd =
                        doc' <- insertTextForMode (text.Substring(segmentStart, segmentEnd - segmentStart)) doc'

                    while index < text.Length do
                        let ch = text.[index]
                        if ch = '\r' || ch = '\n' then
                            sawLineBreak <- true
                            insertSegment index
                            doc' <- Doc.splitBlock doc'
                            if ch = '\r' && index + 1 < text.Length && text.[index + 1] = '\n' then
                                index <- index + 1
                            segmentStart <- index + 1
                        index <- index + 1

                    if sawLineBreak then
                        insertSegment text.Length
                    else
                        doc' <- insertTextForMode text doc'

                    if not (documentsEqualForChange doc' beforeDoc) then
                        let cmd = DeleteTextCommand(beforeDoc, doc', "Paste") :> ICommand
                        let doc'', hist' = Cmd.execute cmd beforeDoc handler.History
                        handler.Document <- doc''
                        handler.History <- hist'
                        renderWithHover handler.RenderContext doc'' handler.HoveredImage
                        handler.OnChange doc''
            ) |> ignore
        | "Backspace" ->
            let beforeDoc = handler.Document
            let afterDoc = deleteBackwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete backward") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "Delete" ->
            let beforeDoc = handler.Document
            let afterDoc = deleteForwardForMode handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, "Delete forward") :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | "ArrowLeft" ->
            if ke.shiftKey then
                newDoc <- Doc.extendSelectionLeft handler.Document
            else
                newDoc <- Doc.moveCursorLeft handler.Document
            e.preventDefault()
        | "ArrowRight" ->
            if ke.shiftKey then
                newDoc <- Doc.extendSelectionRight handler.Document
            else
                newDoc <- Doc.moveCursorRight handler.Document
            e.preventDefault()
        | "ArrowUp" ->
            newDoc <- Doc.moveCursorUp handler.Document
            e.preventDefault()
        | "ArrowDown" ->
            newDoc <- Doc.moveCursorDown handler.Document
            e.preventDefault()
        | "Home" ->
            newDoc <- Doc.moveCursorToLineStart handler.Document
            e.preventDefault()
        | "End" ->
            newDoc <- Doc.moveCursorToLineEnd handler.Document
            e.preventDefault()
        | "Tab" ->
            if ke.shiftKey then
                // Shift+Tab: Decrease indent
                newDoc <- Doc.decreaseIndent handler.Document
            else
                // Tab: Increase indent
                newDoc <- Doc.increaseIndent handler.Document
            e.preventDefault()
        | "Enter" ->
            // Check if we're on an empty list item - if so, exit list
            match handler.Document.Selection with
            | Some sel ->
                match Doc.findBlock handler.Document sel.Focus.BlockId with
                | Some block ->
                    let para = Doc.getParagraph block
                    let isEmpty = TextBuffer.length para.Buffer = 0
                    let inList = para.Metadata.ListStyle <> NoList

                    if isEmpty && inList then
                        // Exit list instead of splitting
                        newDoc <- Doc.decreaseIndent handler.Document
                        e.preventDefault()
                    else
                        // Normal split
                        let beforeDoc = handler.Document
                        let afterDoc = Doc.splitBlock handler.Document
                        if not (documentsEqualForChange afterDoc beforeDoc) then
                            let cmd = SplitBlockCommand(beforeDoc, afterDoc) :> ICommand
                            let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                            newDoc <- doc'
                            newHistory <- hist'
                            useCommand <- true
                        e.preventDefault()
                | None ->
                    e.preventDefault()
            | None ->
                e.preventDefault()
        | key when key.Length = 1 && not ke.ctrlKey && not ke.metaKey ->
            // Regular character - use command
            let beforeDoc = handler.Document
            let afterDoc = insertTextForMode key handler.Document
            if not (documentsEqualForChange afterDoc beforeDoc) then
                let cmd = DeleteTextCommand(beforeDoc, afterDoc, sprintf "Insert '%s'" key) :> ICommand
                let (doc', hist') = Cmd.execute cmd beforeDoc handler.History
                newDoc <- doc'
                newHistory <- hist'
                useCommand <- true
            e.preventDefault()
        | _ -> ()

        if not (documentsEqualForChange newDoc handler.Document) then
            handler.Document <- newDoc
            if useCommand || newHistory <> handler.History then
                handler.History <- newHistory
            renderWithHover handler.RenderContext newDoc handler.HoveredImage
            handler.OnChange newDoc
    )

    // Handle mouse down - start selection or image resize
    canvas.addEventListener("mousedown", fun (e: Event) ->
        let me = e :?> MouseEvent
        let rect = canvas.getBoundingClientRect()
        let x = me.clientX - rect.left
        let y = me.clientY - rect.top

        // Check if clicking on an image first
        match hitTestForImage handler.RenderContext handler.Document x y with
        | Some (blockId, offset, width, height, imageX, imageY) ->
            if isResizeHandle x y imageX imageY width height then
                handler.ResizingImage <- Some {
                    BlockId = blockId
                    Offset = offset
                    InitialX = x
                    InitialY = y
                    InitialWidth = width
                    InitialHeight = height
                }
            else
                handler.DraggingImage <- Some { BlockId = blockId; Offset = offset; InitialX = x; InitialY = y }
            handler.IsMouseDown <- true
            canvas.focus()
        | None ->
            // Regular text selection
            match hitTest handler.RenderContext handler.Document x y with
            | None -> ()
            | Some (blockId, offset) ->
                handler.IsMouseDown <- true
                handler.SelectionAnchor <- Some { BlockId = blockId; Offset = offset }
                let newDoc = Doc.setCursor blockId offset handler.Document
                applyPointerDocumentChange newDoc
                canvas.focus()
    )

    // Handle mouse move - extend selection, resize image, or update hover
    canvas.addEventListener("mousemove", fun (e: Event) ->
        let me = e :?> MouseEvent
        let rect = canvas.getBoundingClientRect()
        let x = me.clientX - rect.left
        let y = me.clientY - rect.top

        if handler.IsMouseDown then
            match handler.ResizingImage with
            | Some resizeState ->
                // Resize image based on mouse delta
                let deltaX = x - resizeState.InitialX
                let deltaY = y - resizeState.InitialY

                // Maintain aspect ratio by using the larger delta
                let delta = if abs deltaX > abs deltaY then deltaX else deltaY
                let newWidth = max 20.0 (resizeState.InitialWidth + delta)
                let newHeight = max 20.0 (resizeState.InitialHeight + delta)

                let newDoc = Doc.resizeImage resizeState.BlockId resizeState.Offset newWidth newHeight handler.Document
                applyPointerDocumentChange newDoc
            | None ->
                match handler.DraggingImage with
                | Some _ -> setCursor canvas "move"
                | None ->
                    // Regular text selection
                    match hitTest handler.RenderContext handler.Document x y, handler.SelectionAnchor with
                    | Some (blockId, offset), Some anchor ->
                        let newDoc = Doc.setSelection anchor { BlockId = blockId; Offset = offset } handler.Document
                        applyPointerDocumentChange newDoc
                    | _ -> ()
        else
            // Update hover state and cursor
            match hitTestForImage handler.RenderContext handler.Document x y with
            | Some (blockId, offset, width, height, imageX, imageY) ->
                setCursor canvas (if isResizeHandle x y imageX imageY width height then "nwse-resize" else "move")
                let newHover = Some (blockId, offset)
                if handler.HoveredImage <> newHover then
                    handler.HoveredImage <- newHover
                    renderWithHover handler.RenderContext handler.Document handler.HoveredImage
            | None ->
                setCursor canvas "text"
                if handler.HoveredImage.IsSome then
                    handler.HoveredImage <- None
                    renderWithHover handler.RenderContext handler.Document None
    )

    // Handle mouse up - end selection or resize
    canvas.addEventListener("mouseup", fun (e: Event) ->
        let me = e :?> MouseEvent
        let rect = canvas.getBoundingClientRect()
        let x = me.clientX - rect.left
        let y = me.clientY - rect.top

        match handler.DraggingImage with
        | Some dragState when abs (x - dragState.InitialX) > 3.0 || abs (y - dragState.InitialY) > 3.0 ->
            let target =
                match hitTest handler.RenderContext handler.Document x y with
                | Some target -> target
                | None ->
                    match Doc.findParagraph handler.Document dragState.BlockId with
                    | Some para -> (dragState.BlockId, TextBuffer.length para.Buffer)
                    | None -> (dragState.BlockId, dragState.Offset)
            let targetBlockId, targetOffset = target
            let effectiveOffset =
                match Doc.findParagraph handler.Document targetBlockId with
                | Some para when targetBlockId = dragState.BlockId && targetOffset <= dragState.Offset -> TextBuffer.length para.Buffer
                | Some para when targetBlockId = dragState.BlockId && targetOffset = dragState.Offset -> TextBuffer.length para.Buffer
                | _ -> targetOffset
            let newDoc = Doc.moveImage dragState.BlockId dragState.Offset targetBlockId effectiveOffset handler.Document
            applyPointerDocumentChange newDoc
        | _ -> ()

        handler.IsMouseDown <- false
        handler.ResizingImage <- None
        handler.DraggingImage <- None
    )

    // Initial render
    renderWithHover ctx initialDoc None

    handler

/// Update document
let updateDocument (handler: InputHandler) (doc: Doc.Document) =
    handler.Document <- doc
    renderWithHover handler.RenderContext doc None
    handler.OnChange doc

#else
// Server-side stub
type InputHandler = unit
let create _ _ _ = ()
let updateDocument _ _ = ()
#endif
