module Rbt.WordWisp.EditorHost

// Wisp maps key and pointer events to the shared Word editing session. Only
// visual-line target calculation, caret geometry and painting remain here.

open Rbt.Renderer
open System.Collections.Generic
open Wisp.Reactive
open Wisp.Layout
open Rbt.Word.Document

module WordLayout = Rbt.Word.Layout
module WordPaint = Rbt.Word.Paint
module WordTextBuffer = Rbt.Word.TextBuffer
module WR = Rbt.WordNative.Render

type KeyEvent = { Key: string; Shift: bool; Ctrl: bool; Alt: bool }

/// Clipboard data stays rich only while the OS transport proves ownership of this write.
module private Clip =
    let mutable capture: (string * Rbt.Word.Session.ClipboardCapture) option = None

type ClipboardTransport =
    { WriteText: string -> Result<string, string>
      ReadText: unit -> Result<Wisp.OfficeClipboard.Contents, string> }

let private systemClipboard =
    { WriteText = Wisp.OfficeClipboard.writeText
      ReadText = Wisp.OfficeClipboard.readText }

type EditorState =
    { Session: Rbt.Word.Session.EditingSession
      Doc: RefSignal<Document>
      Clipboard: ClipboardTransport
      mutable ReportError: string -> unit
      mutable DragAnchor: Position option
      ResolveImage: string -> Framebuffer option
      LayoutCache: LayoutCache }

and LayoutCache =
    { mutable Doc: Document option
      mutable LayoutDoc: Document option
      mutable Settings: WordLayout.LayoutSettings option
      mutable Layouts: WordLayout.BlockLayout list option }

let private historyDepth = 50

let private createLayoutCache () =
    { Doc = None; LayoutDoc = None; Settings = None; Layouts = None }

let createWithClipboard (clipboard: ClipboardTransport) (doc: Document) : EditorState =
    { Session = Rbt.Word.Session.EditingSession(doc, maxHistory = historyDepth)
      Doc = refSignal doc; Clipboard = clipboard; ReportError = (fun error -> eprintfn "[wisp-word] %s" error)
      DragAnchor = None
      ResolveImage = WR.createImageResolver ()
      LayoutCache = createLayoutCache () }

let create (doc: Document) : EditorState = createWithClipboard systemClipboard doc

let private insertWithPending (st: EditorState) text doc =
    let inserted = st.Session.WithPendingFormat(doc, insertText text doc)
    if not (obj.ReferenceEquals(inserted, doc)) then st.Session.CommitPendingInsertion(inserted)
    inserted

let private updateDocument (st: EditorState) doc =
    st.Session.Document <- doc
    st.Doc.Value <- st.Session.Document

/// Run a structural command in the shared session, then publish its projection.
let executeCommand (st: EditorState) (command: Rbt.Word.Session.WordCommand) tracking =
    let result = st.Session.Execute(command, tracking)
    match result with
    | Ok Rbt.Word.Session.NoChange -> ()
    | Ok _ -> st.Doc.Value <- st.Session.Document
    | Error _ -> ()
    result

let private handleCommand st command =
    match executeCommand st command None with
    | Ok _ -> true
    | Error _ -> false



/// Release this host's rich transfer without altering the OS clipboard.
let dispose (st: EditorState) =
    match Clip.capture with
    | Some (_, capture) when obj.ReferenceEquals(capture.Epoch, st.Session.Epoch) ->
        Clip.capture <- None
    | _ -> ()

let replaceDocument (st: EditorState) (doc: Document) =
    dispose st
    st.Session.Replace doc
    st.DragAnchor <- None
    st.Doc.Value <- st.Session.Document

/// Enter a portable story without adding native story UI.
let enterStory (st: EditorState) (story: Rbt.Word.Session.StorySession) selection =
    st.Session.EnterStory(story, selection)
    st.DragAnchor <- None
    st.Doc.Value <- st.Session.Document

let exitStory (st: EditorState) =
    let exited = st.Session.ExitStory()
    if exited then
        st.DragAnchor <- None
        st.Doc.Value <- st.Session.Document
    exited

let private cachedLayouts (st: EditorState) doc =
    let cache = st.LayoutCache
    match cache.Doc, cache.Settings, cache.Layouts with
    | Some cachedDoc, Some settings, Some layouts when System.Object.ReferenceEquals(cachedDoc, doc) ->
        settings, layouts
    | _ ->
        let layoutDoc = { doc with Selection = None }
        match cache.LayoutDoc, cache.Settings, cache.Layouts with
        | Some cachedLayoutDoc, Some settings, Some layouts when cachedLayoutDoc = layoutDoc ->
            cache.Doc <- Some doc
            settings, layouts
        | _ ->
            let settings = WR.layoutSettings doc
            let layouts = WordLayout.layoutDocument settings (WR.measureTextFor doc) doc
            cache.Doc <- Some doc
            cache.LayoutDoc <- Some layoutDoc
            cache.Settings <- Some settings
            cache.Layouts <- Some layouts
            settings, layouts


let private hitTest st width doc x y =
    let settings, layouts = cachedLayouts st doc
    WordLayout.hitTestLayouts (WR.measureTextFor doc) settings.PaddingLeft layouts x y

/// Caret position in document space (x, top-y, height) for a collapsed selection.
let private caretGeom st width doc : (float * float * float) option =
    match doc.Selection with
    | Some sel ->
        let settings, layouts = cachedLayouts st doc
        match layouts |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId) with
        | Some layout ->
            let line = WordLayout.lineForOffset layout sel.Focus.Offset
            let x = WordLayout.xForOffset (WR.measureTextFor doc) line sel.Focus.Offset
            Some (settings.PaddingLeft + x + WordLayout.alignmentOffset layout line, line.Y, line.Height)
        | None -> None
    | None -> None

let private moveVisualLine st width direction extend doc =
    let target =
        match caretGeom st width doc with
        | Some (x, y, _) ->
            let settings = WR.layoutSettings doc
            hitTest st width doc x (y + settings.LineHeight * float direction)
            |> Option.map (fun (b, o) -> { BlockId = b; Offset = o })
        | None -> None
    Rbt.Word.Session.navigate
        (Rbt.Word.Session.Navigation.VisualTarget(target, direction > 0, extend)) doc

let private moveToLineBoundary st width toEnd extend doc =
    let target =
        match doc.Selection with
        | Some sel ->
            let _, layouts = cachedLayouts st doc
            layouts
            |> List.tryFind (fun l -> l.BlockId = sel.Focus.BlockId)
            |> Option.map (fun layout ->
                let line = WordLayout.lineForOffset layout sel.Focus.Offset
                { BlockId = sel.Focus.BlockId; Offset = (if toEnd then line.EndOffset else line.StartOffset) })
        | None -> None
    match target with
    | Some position ->
        Rbt.Word.Session.navigate (Rbt.Word.Session.Navigation.VisualTarget(Some position, toEnd, extend)) doc
    | None -> doc

// ── apply + undo ──
let applyChange (st: EditorState) (transform: Document -> Document) : bool =
    st.Session.StopTyping()
    let before = st.Session.Document
    let after = transform before
    if after <> before then
        st.Session.RecordChange("Edit document", before, after, id)
        updateDocument st after
    true

let private applyTyping (st: EditorState) kind (transform: Document -> Document) : bool =
    let before = st.Session.Document
    let after = transform before
    if after <> before then
        st.Session.RecordTyping(kind, "Type " + kind, before, after, id)
        updateDocument st after
    true

let applyNoHistory (st: EditorState) (transform: Document -> Document) : bool =
    st.Session.StopTyping()
    let before = st.Session.Document
    let after = transform before
    if after <> before then updateDocument st after
    true

let undo (st: EditorState) : bool =
    if st.Session.Undo() then st.Doc.Value <- st.Session.Document
    true

let redo (st: EditorState) : bool =
    if st.Session.Redo() then st.Doc.Value <- st.Session.Document
    true

// convenience for the chrome actions
/// The session owns both the caret policy and the undo transaction.
let formatCharacter (st: EditorState) format =
    handleCommand st (Rbt.Word.Session.WordCommand.Character format) |> ignore
let edit (st: EditorState) (f: Document -> Document) = applyChange st f |> ignore
let private navigate (st: EditorState) command =
    let outcome = st.Session.Navigate(command)
    if outcome <> Rbt.Word.Session.NoChange then st.Doc.Value <- st.Session.Document
    true

let private visualNavigate st extend move =
    let after = move st.Session.Document
    match after.Selection with
    | Some selection -> navigate st (Rbt.Word.Session.Navigation.VisualTarget(Some selection.Focus, false, extend))
    | None -> true

// ── keyboard ──
let handleKey (st: EditorState) (width: float) (ev: KeyEvent) : bool =
    let ch f = applyChange st f
    let publishClipboard outcome =
        if outcome <> Rbt.Word.Session.NoChange then st.Doc.Value <- st.Session.Document
        true
    let typed kind f =
        match st.Session.Document.Selection with
        | Some selection when selection.IsCollapsed -> applyTyping st kind f
        | _ -> ch f
    let insert text = typed "insert" (insertWithPending st text)
    let isCtrl k = ev.Ctrl && (ev.Key = k || ev.Key = k.ToUpperInvariant())
    if ev.Ctrl then
        if isCtrl "a" then navigate st Rbt.Word.Session.Navigation.SelectAll
        elif ev.Shift && isCtrl "z" then redo st
        elif isCtrl "z" then undo st
        elif isCtrl "y" then redo st
        elif isCtrl "c" || isCtrl "x" then
            // Private provenance marker also identifies rich clips with no plain text.
            match st.Session.CaptureClipboard(fun _ -> System.Guid.NewGuid().ToString("N")) with
            | Some capture ->
                match st.Clipboard.WriteText capture.Text with
                | Ok token ->
                    Clip.capture <- Some (token, capture)
                    if isCtrl "x" then st.Session.CommitCut(capture, None) |> Option.iter (publishClipboard >> ignore)
                | Error error ->
                    Clip.capture <- None
                    st.ReportError("Clipboard write failed: " + error)
            | None -> ()
            true
        elif isCtrl "v" then
            match st.Clipboard.ReadText() with
            | Error error ->
                Clip.capture <- None
                st.ReportError("Clipboard read failed: " + error)
                true
            | Ok contents ->
                let html, decode =
                    match Clip.capture, contents.Token with
                    | Some (token, capture), Some current when token = current ->
                        if obj.ReferenceEquals(capture.Epoch, st.Session.Epoch) then
                            capture.Html, Rbt.Word.Session.noExternalClip
                        else capture.Html, (fun _ -> Some capture.Clip)
                    | _ ->
                        Clip.capture <- None
                        "", Rbt.Word.Session.noExternalClip
                publishClipboard (st.Session.PasteClipboard(contents.Text, html, decode, None))
        elif isCtrl "b" then formatCharacter st Rbt.Word.Session.CharacterCommand.Bold; true
        elif isCtrl "i" then formatCharacter st Rbt.Word.Session.CharacterCommand.Italic; true
        elif isCtrl "u" then formatCharacter st Rbt.Word.Session.CharacterCommand.Underline; true
        elif ev.Key = "Tab" && not ev.Alt then handleCommand st (Rbt.Word.Session.Tab(ev.Shift, true))
        elif ev.Key = "Enter" && not ev.Alt then
            handleCommand st (if ev.Shift then Rbt.Word.Session.ColumnBreak else Rbt.Word.Session.PageBreak)
        elif ev.Key = "ArrowLeft" then navigate st (Rbt.Word.Session.Navigation.Word(false, ev.Shift))
        elif ev.Key = "ArrowRight" then navigate st (Rbt.Word.Session.Navigation.Word(true, ev.Shift))
        elif ev.Key = "Home" then navigate st (Rbt.Word.Session.Navigation.DocumentEdge(false, ev.Shift))
        elif ev.Key = "End" then navigate st (Rbt.Word.Session.Navigation.DocumentEdge(true, ev.Shift))
        else st.Session.StopTyping(); false
    else
        match ev.Key with
        | "Backspace" -> typed "backspace" deleteBackward
        | "Delete" -> typed "delete" deleteForward
        | "Enter" when ev.Shift && not ev.Alt -> handleCommand st Rbt.Word.Session.LineBreak
        | "Enter" when not ev.Alt -> handleCommand st Rbt.Word.Session.Enter
        | "Tab" when not ev.Alt -> handleCommand st (Rbt.Word.Session.Tab(ev.Shift, ev.Ctrl))
        | "ArrowLeft" -> navigate st (Rbt.Word.Session.Navigation.Character(false, ev.Shift))
        | "ArrowRight" -> navigate st (Rbt.Word.Session.Navigation.Character(true, ev.Shift))
        | "ArrowUp" -> visualNavigate st ev.Shift (moveVisualLine st width -1 ev.Shift)
        | "ArrowDown" -> visualNavigate st ev.Shift (moveVisualLine st width 1 ev.Shift)
        | "Home" -> visualNavigate st ev.Shift (moveToLineBoundary st width false ev.Shift)
        | "End" -> visualNavigate st ev.Shift (moveToLineBoundary st width true ev.Shift)
        | k when k.Length = 1 && not ev.Alt -> insert k
        | _ -> st.Session.StopTyping(); false

// ── mouse (coords in document space) ──
let handlePointer (st: EditorState) (width: float) (kind: PointerKind) (x: float) (y: float) =
    let scale = WR.scale width st.Doc.Value
    let x, y = x / scale, y / scale
    match kind with
    | Down ->
        st.Session.StopTyping()
        match hitTest st width st.Doc.Value x y with
        | Some (b, o) ->
            st.DragAnchor <- Some { BlockId = b; Offset = o }
            updateDocument st (setCursor b o st.Session.Document)
        | None -> ()
    | Move ->
        match st.DragAnchor, hitTest st width st.Doc.Value x y with
        | Some anchor, Some (b, o) -> updateDocument st (setSelection anchor { BlockId = b; Offset = o } st.Session.Document)
        | _ -> ()
    | Up -> st.DragAnchor <- None

// ── render: content + selection + caret, all in document space ──
let documentHeight (st: EditorState) (width: float) (doc: Document) : float =
    let settings, layouts = cachedLayouts st doc
    let height =
        match layouts with
        | [] -> settings.PaddingTop * 2.0 + settings.LineHeight
        | _ ->
            let mutable bottom = 0.0
            for layout in layouts do
                bottom <- max bottom (layout.Y + layout.Height)
            bottom + settings.MarginBottom
    max settings.PageHeight height * WR.scale width doc

let renderDoc (st: EditorState) (width: float) (doc: Document) : Cmd list =
    let settings, cached = cachedLayouts st doc
    let layouts, content, _ =
        WordPaint.collectDocumentCommandsWithLayouts settings "Calibri" (WR.measureTextFor doc) (WordLayout.rangeMeasurerFromTextMeasurer (WR.measureTextFor doc)) doc (Some cached)
    let body = ResizeArray<Cmd>()
    for paintCommand in content do
        body.AddRange(WR.toRendererCommandFor doc st.ResolveImage paintCommand)
    for paintCommand in WordPaint.selectionCommands settings (WR.measureTextFor doc) doc layouts do
        body.AddRange(WR.toRendererCommandFor doc st.ResolveImage paintCommand)
    let caret =
        match doc.Selection with
        | Some s when s.IsCollapsed ->
            match caretGeom st width doc with
            | Some (cx, cy, ch) -> [ FillRect(rect cx cy 1.6 ch, Color.rgb 32uy 64uy 140uy) ]
            | None -> []
        | _ -> []
    body.AddRange caret
    List.ofSeq body |> WR.viewTransform width doc
