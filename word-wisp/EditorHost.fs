module Rbt.WordWisp.EditorHost

// The Wisp editor host for wisp-word — the per-backend glue the Office UI abstraction never covered.
// All the actual editing is shared, toolkit-free Rbt.Word.Document ops (setCursor, insertText,
// deleteBackward, toggleBold, moveCursorLeft, …); this just maps Wisp key/pointer events to them,
// ports the two layout-based helpers (visual-line nav, caret rect), and renders selection + caret.

open Rbt.Renderer
open System.Collections.Generic
open Wisp.Reactive
open Wisp.Layout
open Rbt.Word.Document

module WordLayout = Rbt.Word.Layout
module WordPaint = Rbt.Word.Paint
module WordTextBuffer = Rbt.Word.TextBuffer
module WR = Rbt.WordWisp.WordRender

type KeyEvent = { Key: string; Shift: bool; Ctrl: bool; Alt: bool }

/// In-process clipboard (the shared Platform has no X11 selection).
module Clip = let mutable text = ""

type EditorState =
    { Doc: RefSignal<Document>
      mutable Undo: Document list
      mutable Redo: Document list
      mutable DragAnchor: Position option
      LayoutCache: LayoutCache }

and LayoutCache =
    { mutable Width: float
      mutable Doc: Document option
      mutable LayoutDoc: Document option
      mutable Settings: WordLayout.LayoutSettings option
      mutable Layouts: WordLayout.BlockLayout list option }

let private maxUndoDepth = 50
let private trimHistory stack = stack |> List.truncate maxUndoDepth

let private createLayoutCache () =
    { Width = nan; Doc = None; LayoutDoc = None; Settings = None; Layouts = None }

let create (doc: Document) : EditorState =
    { Doc = refSignal doc; Undo = []; Redo = []; DragAnchor = None; LayoutCache = createLayoutCache () }

let private cachedLayouts (st: EditorState) width doc =
    let cache = st.LayoutCache
    let layoutDoc = { doc with Selection = None }
    let matches =
        cache.Width = width &&
        match cache.Doc with
        | Some cachedDoc when System.Object.ReferenceEquals(cachedDoc, doc) -> true
        | _ ->
            match cache.LayoutDoc with
            | Some cachedLayoutDoc -> cachedLayoutDoc = layoutDoc
            | None -> false
    match matches, cache.Settings, cache.Layouts with
    | true, Some settings, Some layouts ->
        cache.Doc <- Some doc
        settings, layouts
    | _ ->
        let settings = WR.layoutSettings width doc
        let layouts = WordLayout.layoutDocument settings WR.measureText doc
        cache.Width <- width
        cache.Doc <- Some doc
        cache.LayoutDoc <- Some layoutDoc
        cache.Settings <- Some settings
        cache.Layouts <- Some layouts
        settings, layouts

// ── document boundaries (ported from the native host) ──
let private editableParagraphs doc =
    doc.Blocks
    |> List.collect (function
        | Paragraph p | Heading (_, p) -> [ p ]
        | Table table -> table.Rows |> List.collect (fun row -> row.Cells |> List.collect (fun cell -> cell.Content)))

let private documentStart doc : Position option =
    editableParagraphs doc |> List.tryHead |> Option.map (fun para -> { BlockId = para.Id; Offset = 0 })
let private documentEnd doc : Position option =
    editableParagraphs doc |> List.tryLast |> Option.map (fun para -> { BlockId = para.Id; Offset = WordTextBuffer.length para.Buffer })
let private moveToDocumentBoundary boundary doc =
    match boundary doc with Some (p: Position) -> setCursor p.BlockId p.Offset doc | None -> doc

let private hitTest st width doc x y =
    let settings, layouts = cachedLayouts st width doc
    WordLayout.hitTestLayouts WR.measureText settings.PaddingLeft layouts x y

/// Caret position in document space (x, top-y, height) for a collapsed selection.
let private caretGeom st width doc : (float * float * float) option =
    match doc.Selection with
    | Some sel ->
        let settings, layouts = cachedLayouts st width doc
        match layouts |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId) with
        | Some layout ->
            let line = WordLayout.lineForOffset layout sel.Focus.Offset
            let x = WordLayout.xForOffset WR.measureText line sel.Focus.Offset
            Some (settings.PaddingLeft + x + WordLayout.alignmentOffset layout line, line.Y, line.Height)
        | None -> None
    | None -> None

let private moveVisualLine st width direction extend doc =
    let target =
        match caretGeom st width doc with
        | Some (x, y, _) ->
            let settings = WR.layoutSettings width doc
            hitTest st width doc x (y + settings.LineHeight * float direction)
            |> Option.map (fun (b, o) -> { BlockId = b; Offset = o })
        | None -> None
    match doc.Selection, target with
    | Some sel, Some pos when extend -> { doc with Selection = Some { sel with Focus = pos } }
    | _, Some pos -> setCursor pos.BlockId pos.Offset doc
    | _, None when direction < 0 -> moveToDocumentBoundary documentStart doc
    | _, None -> moveToDocumentBoundary documentEnd doc

let private moveToLineBoundary st width toEnd extend doc =
    let target =
        match doc.Selection with
        | Some sel ->
            let _, layouts = cachedLayouts st width doc
            layouts
            |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId)
            |> Option.map (fun layout ->
                let line = WordLayout.lineForOffset layout sel.Focus.Offset
                { BlockId = sel.Focus.BlockId; Offset = (if toEnd then line.EndOffset else line.StartOffset) })
        | None -> None
    match doc.Selection, target with
    | Some sel, Some pos when extend -> { doc with Selection = Some { sel with Focus = pos } }
    | _, Some pos -> setCursor pos.BlockId pos.Offset doc
    | _ -> doc

// ── apply + undo ──
let applyChange (st: EditorState) (transform: Document -> Document) : bool =
    let before = st.Doc.Value
    let after = transform before
    if after <> before then
        st.Undo <- (before :: st.Undo) |> trimHistory
        st.Redo <- []
        st.Doc.Value <- after
    true
let applyNoHistory (st: EditorState) (transform: Document -> Document) : bool =
    let after = transform st.Doc.Value
    if after <> st.Doc.Value then st.Doc.Value <- after
    true
let undo (st: EditorState) : bool =
    match st.Undo with d :: rest -> st.Redo <- (st.Doc.Value :: st.Redo) |> trimHistory; st.Undo <- rest; st.Doc.Value <- d; true | [] -> true
let redo (st: EditorState) : bool =
    match st.Redo with d :: rest -> st.Undo <- (st.Doc.Value :: st.Undo) |> trimHistory; st.Redo <- rest; st.Doc.Value <- d; true | [] -> true

// convenience for the chrome actions
let edit (st: EditorState) (f: Document -> Document) = applyChange st f |> ignore

// ── keyboard ──
let handleKey (st: EditorState) (width: float) (ev: KeyEvent) : bool =
    let ch f = applyChange st f
    let nh f = applyNoHistory st f
    let isCtrl k = ev.Ctrl && (ev.Key = k || ev.Key = k.ToUpperInvariant())
    if ev.Ctrl then
        if isCtrl "a" then nh selectAll
        elif ev.Shift && isCtrl "z" then redo st
        elif isCtrl "z" then undo st
        elif isCtrl "y" then redo st
        elif isCtrl "c" then Clip.text <- getSelectionText st.Doc.Value; true
        elif isCtrl "x" then (let t = getSelectionText st.Doc.Value in if t <> "" then (Clip.text <- t; ch deleteSelection |> ignore)); true
        elif isCtrl "v" then (if Clip.text <> "" then ch (insertText Clip.text) |> ignore); true
        elif isCtrl "b" then ch Rbt.Word.WordStyles.toggleBold
        elif isCtrl "i" then ch Rbt.Word.WordStyles.toggleItalic
        elif isCtrl "u" then ch Rbt.Word.WordStyles.toggleUnderline
        elif ev.Key = "ArrowLeft" then nh moveCursorLeft
        elif ev.Key = "ArrowRight" then nh moveCursorRight
        elif ev.Key = "Home" then nh (moveToDocumentBoundary documentStart)
        elif ev.Key = "End" then nh (moveToDocumentBoundary documentEnd)
        else false
    else
        match ev.Key with
        | "Backspace" -> ch deleteBackward
        | "Delete" -> ch deleteForward
        | "Enter" -> ch splitBlock
        | "Tab" when ev.Shift -> ch decreaseIndent
        | "Tab" -> ch increaseIndent
        | "ArrowLeft" when ev.Shift -> nh extendSelectionLeft
        | "ArrowRight" when ev.Shift -> nh extendSelectionRight
        | "ArrowUp" -> nh (moveVisualLine st width -1 ev.Shift)
        | "ArrowDown" -> nh (moveVisualLine st width 1 ev.Shift)
        | "ArrowLeft" -> nh moveCursorLeft
        | "ArrowRight" -> nh moveCursorRight
        | "Home" -> nh (moveToLineBoundary st width false ev.Shift)
        | "End" -> nh (moveToLineBoundary st width true ev.Shift)
        | k when k.Length = 1 && not ev.Alt -> ch (insertText k)
        | _ -> false

// ── mouse (coords in document space) ──
let handlePointer (st: EditorState) (width: float) (kind: PointerKind) (x: float) (y: float) =
    match kind with
    | Down ->
        match hitTest st width st.Doc.Value x y with
        | Some (b, o) -> st.DragAnchor <- Some { BlockId = b; Offset = o }; st.Doc.Value <- setCursor b o st.Doc.Value
        | None -> ()
    | Move ->
        match st.DragAnchor, hitTest st width st.Doc.Value x y with
        | Some anchor, Some (b, o) -> st.Doc.Value <- setSelection anchor { BlockId = b; Offset = o } st.Doc.Value
        | _ -> ()
    | Up -> st.DragAnchor <- None

// ── render: content + selection + caret, all in document space ──
let documentHeight (st: EditorState) (width: float) (doc: Document) : float =
    let settings, layouts = cachedLayouts st width doc
    match layouts with
    | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
    | _ ->
        let mutable bottom = 0.0
        for layout in layouts do
            bottom <- max bottom (layout.Y + layout.Height)
        bottom + settings.MarginBottom

let renderDoc (st: EditorState) (width: float) (doc: Document) : Cmd list =
    let settings, cached = cachedLayouts st width doc
    let layouts, content, _ =
        WordPaint.collectDocumentCommandsWithLayouts settings "Calibri" WR.measureText (WordLayout.rangeMeasurerFromTextMeasurer WR.measureText) doc (Some cached)
    let body = ResizeArray<Cmd>()
    for paintCommand in content do
        body.AddRange(WR.toRendererCommand paintCommand)
    for paintCommand in WordPaint.selectionCommands settings WR.measureText doc layouts do
        body.AddRange(WR.toRendererCommand paintCommand)
    let caret =
        match doc.Selection with
        | Some s when s.IsCollapsed ->
            match caretGeom st width doc with
            | Some (cx, cy, ch) -> [ FillRect(rect cx cy 1.6 ch, Color.rgb 32uy 64uy 140uy) ]
            | None -> []
        | _ -> []
    body.AddRange caret
    List.ofSeq body
