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
        /// Callbacks registered with onChange.
        Listeners: ResizeArray<obj -> unit>
        /// Undo / redo availability as last reported to the listeners.
        mutable NotifiedCanUndo: bool
        mutable NotifiedCanRedo: bool
        /// The revision of the document's content (see getRevision): a
        /// number given to each state of the content when it first appears.
        /// Undo and redo go back to states that have their number already.
        mutable Revision: int
        mutable LastRevision: int
        /// The revision the host last saved (see markSaved).
        mutable SavedRevision: int
        /// The document the current revision was given to, and the states
        /// undo / redo can come back to, newest first.
        mutable RevisionDocument: Doc.Document
        KnownRevisions: ResizeArray<Doc.Document * int>
        /// Font families the host's text stack can draw (see setFontFamilies).
        mutable FontFamilies: string array
        /// The whole document as the listeners last heard of it. (While a
        /// header or footer is edited in place `Doc` is the story's own
        /// document; this is the body with that story in its part.)
        mutable Whole: Doc.Document
        /// The package the document on display is of, while it is still being
        /// read (see importFromDocxPackage): the document grows as it is.
        mutable Loading: Docx.ProgressiveImport option
        /// Ends the editor's interest in that package.
        mutable StopLoading: unit -> unit
        /// When the listeners were last told that more was read, and who waits for the end.
        mutable LoadNotified: float
        LoadWaiters: ResizeArray<unit -> unit>
        /// Why the document on display could not be read to its end, when it
        /// could not (see isLoadFailed): it is what was read, and is not saved.
        mutable LoadFailure: string option
    }

    /// Read the document on display to its end now, when its package is
    /// still being read: what goes by the whole document (saving, the text,
    /// finding, the page count, selecting everything, the layout queries)
    /// starts with this. The editor's document is the whole one afterwards.
    let private ensureLoaded (editor: Editor) =
        match editor.Loading with
        | Some state -> Docx.completeProgressiveImport state
        | None -> ()

    /// The whole document: what is saved, and what a revision is a state of.
    let private wholeDocument (editor: Editor) : Doc.Document =
        Inp.wholeDocument editor.Handler editor.Doc

    /// What a double click does (into a header or footer, back to the body):
    /// set further down, where the commands it needs are defined.
    let private doubleClick : (Editor -> float -> float -> bool) ref = ref (fun _ _ _ -> false)

    /// Puts the caret into a page's header or footer (enterHeaderFooter), and
    /// brings the fields of the story being edited up to the layout
    /// (refreshStoryFields): set further down as well.
    let private enterStoryOfPage : (Editor -> int -> bool -> bool) ref = ref (fun _ _ _ -> false)
    let private storyFieldsRefresh : (Editor -> unit) ref = ref ignore

    /// Bring the revision in line with the document that is now shown: the
    /// same content keeps its number (a moved caret is not a new revision),
    /// content seen before (undo, redo) gets its number back, anything else
    /// a new one.
    let private observeRevision (editor: Editor) (doc: Doc.Document) =
        // (By identity, in constant time: a caret move keeps every part of the
        // document but the selection, and undo / redo bring back the very
        // documents that were left. Comparing contents walked the blocks of
        // up to 400 remembered states at every keystroke.)
        if not (Inp.sameContentObjects doc editor.RevisionDocument) then
            // What is being left can be come back to.
            editor.KnownRevisions.Insert(0, (editor.RevisionDocument, editor.Revision))
            if editor.KnownRevisions.Count > 400 then editor.KnownRevisions.RemoveAt (editor.KnownRevisions.Count - 1)
            let mutable known = -1
            let mutable i = 0
            while known < 0 && i < editor.KnownRevisions.Count do
                let state, revision = editor.KnownRevisions.[i]
                if Inp.sameContentObjects state doc then known <- revision
                i <- i + 1
            if known >= 0 then editor.Revision <- known
            else
                editor.LastRevision <- editor.LastRevision + 1
                editor.Revision <- editor.LastRevision
            editor.RevisionDocument <- doc

    /// The frames of the pages of the layout the layout queries read (set by
    /// `notePages` when a query begins): which page a y is on, and where that
    /// page starts, follow the pages as they are - sections differ in page
    /// height, so a page is not found by dividing by one page's height.
    let mutable private pagesInUse : Layout.PageInfo array = [||]

    let private notePages (editor: Editor) =
        ensureLoaded editor
        Ren.completeLayout editor.Handler.RenderContext
        pagesInUse <- Ren.pageFrames editor.Handler.RenderContext editor.Doc |> List.toArray

    /// The page (0-based) a layout y is on: the last one that starts at or above it.
    let private pageIndexAt (y: float) =
        let mutable lo = 0
        let mutable hi = pagesInUse.Length - 1
        let mutable found = 0
        while lo <= hi do
            let mid = lo + (hi - lo) / 2
            if pagesInUse.[mid].Top <= y then
                found <- mid
                lo <- mid + 1
            else hi <- mid - 1
        found

    /// The layouts of the whole document (what the commands that read the
    /// layout want: any block may be asked about).
    let private layoutsForEditor (editor: Editor) =
        notePages editor
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
    let private stopLoading (editor: Editor) =
        editor.StopLoading ()
        editor.StopLoading <- ignore
        editor.Loading <- None
        editor.Handler.RenderContext.Progress.LoadPending <- false
        if editor.LoadWaiters.Count > 0 then
            let waiters = editor.LoadWaiters.ToArray()
            editor.LoadWaiters.Clear()
            for waiter in waiters do waiter ()

    /// What follows a step of a package that is being read (set further down).
    let private loadStep : (Editor -> Docx.ImportStep -> unit) ref = ref (fun _ _ -> ())

    /// `loading`: the package the document is of, when it is still being read.
    let private loadDocumentOf (editor: Editor) (newDoc: Doc.Document) (loading: Docx.ProgressiveImport option) =
        // (Held before the document that was on display is let go of: it may be of the same package.)
        let reading = loading |> Option.filter (fun state -> state.Failed.IsNone && not state.Complete)
        reading |> Option.iter Docx.holdProgressiveImport
        // (The document that was on display may have been one that was being read: the editor
        // is done with that package - which ends its reading, unless a viewer shows it too.)
        stopLoading editor
        editor.LoadFailure <- loading |> Option.bind (fun state -> state.Failed)
        match reading with
        | Some state ->
            let listener = fun (step: Docx.ImportStep) -> loadStep.Value editor step
            state.Listeners.Add listener
            editor.Loading <- Some state
            editor.StopLoading <- fun () ->
                state.Listeners.Remove listener |> ignore
                Docx.releaseProgressiveImport state
            editor.Handler.RenderContext.Progress.LoadPending <- true
        | None -> ()
        // Through a viewport the start of the document is laid out and shown
        // first; the rest follows in idle slices (see isLayoutComplete).
        Ren.beginProgressiveLayout editor.Handler.RenderContext
        Inp.loadDocument editor.Handler newDoc
        editor.Doc <- newDoc
        editor.Whole <- newDoc
        // A loaded document is a new document: nothing of the old one can be
        // come back to, and it is as it was saved.
        editor.KnownRevisions.Clear()
        observeRevision editor newDoc
        editor.SavedRevision <- editor.Revision

    let private loadDocument (editor: Editor) newDoc = loadDocumentOf editor newDoc None

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

    /// Detach the editor's window-level listeners and remove its hidden text
    /// input element. Call when the canvas is removed.
    let dispose (editor: Editor) =
        editor.StopLoading ()
        editor.StopLoading <- ignore
        editor.LoadWaiters.Clear()
        editor.Listeners.Clear()
        Ren.cancelScheduledLayout editor.Handler.RenderContext
        Inp.dispose editor.Handler

    /// The hidden text input that holds the keyboard focus while the editor is
    /// in use (IME, dead keys and virtual keyboards type into it); null once
    /// the editor is disposed. Hosts that test for "the editor has focus"
    /// should compare document.activeElement with this element.
    let getInputElement (editor: Editor) : obj =
        match editor.Handler.InputProxy with
        | Some element -> box element
        | None -> null

    /// Tell the onChange listeners what an engine change was. Nothing is
    /// reported when neither the document, the selection nor the undo / redo
    /// availability changed.
    let private notifyListeners (editor: Editor) (previous: Doc.Document) (current: Doc.Document) =
        if editor.Listeners.Count > 0 then
            let documentChanged = not (Inp.sameContent previous current)
            let selectionChanged = previous.Selection <> current.Selection
            let canUndoNow = Inp.canUndo editor.Handler
            let canRedoNow = Inp.canRedo editor.Handler
            if documentChanged || selectionChanged || canUndoNow <> editor.NotifiedCanUndo || canRedoNow <> editor.NotifiedCanRedo then
                editor.NotifiedCanUndo <- canUndoNow
                editor.NotifiedCanRedo <- canRedoNow
                let change =
                    createObj [
                        "DocumentChanged" ==> documentChanged
                        "SelectionChanged" ==> selectionChanged
                        "CanUndo" ==> canUndoNow
                        "CanRedo" ==> canRedoNow
                        "Revision" ==> editor.Revision
                        "Modified" ==> (editor.Revision <> editor.SavedRevision)
                        "LayoutChanged" ==> false
                        "LayoutComplete" ==> Ren.layoutIsComplete editor.Handler.RenderContext
                        "LoadChanged" ==> false
                        "LoadComplete" ==> (editor.Loading.IsNone && editor.LoadFailure.IsNone)
                        "LoadFailed" ==> editor.LoadFailure.IsSome
                        "LoadError" ==> Option.toObj editor.LoadFailure
                    ]
                // A listener may unsubscribe (or subscribe) while being called.
                for listener in editor.Listeners.ToArray() do
                    if editor.Listeners.Contains listener then listener change

    /// Tell the onChange listeners that more of a long document is laid out
    /// (more pages, a larger document size), or all of it.
    let private notifyProgress (editor: Editor) (layoutChanged: bool) (loadChanged: bool) =
        if editor.Listeners.Count > 0 then
            let change =
                createObj [
                    "DocumentChanged" ==> false
                    "SelectionChanged" ==> false
                    "CanUndo" ==> Inp.canUndo editor.Handler
                    "CanRedo" ==> Inp.canRedo editor.Handler
                    "Revision" ==> editor.Revision
                    "Modified" ==> (editor.Revision <> editor.SavedRevision)
                    "LayoutChanged" ==> layoutChanged
                    "LayoutComplete" ==> Ren.layoutIsComplete editor.Handler.RenderContext
                    "LoadChanged" ==> loadChanged
                    "LoadComplete" ==> (editor.Loading.IsNone && editor.LoadFailure.IsNone)
                    "LoadFailed" ==> editor.LoadFailure.IsSome
                    "LoadError" ==> Option.toObj editor.LoadFailure
                ]
            for listener in editor.Listeners.ToArray() do
                if editor.Listeners.Contains listener then listener change

    let private notifyLayoutProgress (editor: Editor) = notifyProgress editor true false

    /// Lay out the rest of a long document in idle slices (nothing to do, and
    /// nothing scheduled, when the whole document is laid out).
    let private continueLayout (editor: Editor) =
        let ctx = editor.Handler.RenderContext
        // (A document that is still being read is read first: its slices and the layout's would
        // take turns, and each layout pass would be of another document.)
        if editor.Loading.IsNone && not (Ren.layoutIsComplete ctx) then
            Ren.scheduleLayout ctx (fun () -> Ren.bodyDocument ctx editor.Doc) (fun complete tell ->
                // (What is on screen was laid out before it was shown: the frame is brought up to the layout at the end.)
                if complete then
                    // (The page count is known now: the fields of a header being edited show it.)
                    storyFieldsRefresh.Value editor
                    Inp.repaint editor.Handler
                if tell then notifyLayoutProgress editor)

    /// Whether the whole document is laid out. A document loaded through a
    /// viewport (setViewport before the import) shows its first screens at
    /// once; the rest is laid out in idle slices, during which
    /// getDocumentSize / getPageFrames report the pages there are so far and
    /// the onChange listeners are told as they grow ({ LayoutChanged: true,
    /// LayoutComplete }). Commands that need the whole layout (getPageCount,
    /// the layout queries, find and go to, renderPage of a later page, a
    /// selection further down) lay the rest out at once.
    let isLayoutComplete (editor: Editor) : bool =
        Ren.layoutIsComplete editor.Handler.RenderContext

    /// A promise that resolves (with the document size, see getDocumentSize)
    /// once the whole document is laid out; at once when it is.
    let whenLayoutComplete (editor: Editor) : JS.Promise<obj> =
        let ctx = editor.Handler.RenderContext
        JS.Constructors.Promise.Create(fun resolve _ ->
            Ren.whenLayoutComplete ctx (fun () -> resolve (Viewer.documentSizeToJs ctx editor.Doc))
            continueLayout editor)

    /// Lay out the rest of the document now (what getPageCount does, without the count).
    let completeLayout (editor: Editor) : unit =
        ensureLoaded editor
        let ctx = editor.Handler.RenderContext
        if not (Ren.layoutIsComplete ctx && ctx.LayoutState.LastWasWhole) then
            Ren.completeLayout ctx
            Inp.repaint editor.Handler
            notifyLayoutProgress editor

    [<Emit("performance.now()")>]
    let private nowMs () : float = jsNative

    // More of the document was read: its blocks come after the blocks the
    // editor has - which may have been edited meanwhile; the caret, the
    // selection, the revision and the undo history are what they were (the
    // documents the history comes back to get the blocks too, see
    // Input.extendDocuments). Nothing of it is laid out before it is shown
    // or its turn comes: the rest of the document is laid out in slices once
    // it is read.
    loadStep.Value <- fun (editor: Editor) (step: Docx.ImportStep) ->
        let ctx = editor.Handler.RenderContext
        // (A header or footer is not edited in place while the document is read - entering one
        // reads the rest first; should one be, it is left: the blocks belong to the body.)
        if editor.Handler.Story.IsSome then Inp.exitStory editor.Handler |> ignore
        let blocksBefore = List.length editor.Doc.Blocks
        let sameWhole = obj.ReferenceEquals(editor.Whole, editor.Doc)
        let doc = Inp.extendDocuments editor.Handler step.Extend
        editor.Doc <- doc
        editor.Whole <- (if sameWhole then doc else step.Extend editor.Whole)
        // (The same content as before, with the same number: reading is not a change of the user's.)
        editor.RevisionDocument <- step.Extend editor.RevisionDocument
        for index in 0 .. editor.KnownRevisions.Count - 1 do
            let known, revision = editor.KnownRevisions.[index]
            editor.KnownRevisions.[index] <- (step.Extend known, revision)
        let wasWhole = Ren.documentExtended ctx blocksBefore
        // (A package that cannot be read on: the document is what was read, from here on. It is
        // laid out, found in and counted as any document; it is not saved.)
        let ended = step.Complete || step.Failed.IsSome
        if step.Failed.IsSome then editor.LoadFailure <- step.Failed
        if ended then stopLoading editor
        // (What was laid out was all there was: the view may reach into what is there now.)
        if wasWhole then Inp.repaint editor.Handler
        let now = nowMs ()
        if ended || now - editor.LoadNotified > 150.0 then
            editor.LoadNotified <- now
            notifyProgress editor false true
        if ended then
            // (The page count is known when the rest is laid out: in slices, from here on.)
            continueLayout editor

    /// Whether the document on display is all there. A package opened front
    /// first (Office.Formats `word.openProgressive`) is shown as soon as its
    /// first blocks are read; the rest is read in slices, during which the
    /// document grows at its end: isLoadComplete is false, the onChange
    /// listeners are told ({ LoadChanged: true, LoadComplete }) and
    /// isLayoutComplete is false as well. The caret moves and the document can
    /// be edited meanwhile, with undo and redo; what goes by the whole
    /// document (saving, getText, find and replace, getPageCount, select all,
    /// Ctrl+End, the layout queries, entering a header or footer) reads the
    /// rest at once.
    ///
    /// A package whose body cannot be read to its end (text that is not well
    /// formed somewhere in a long document) ends the load in a failed state:
    /// the document is what was read before that place and stays on display.
    /// isLoadComplete is false for good, isLoadFailed true and getLoadError
    /// says why; the onChange listeners are told once ({ LoadChanged: true,
    /// LoadComplete: false, LoadFailed: true, LoadError } - every notice from
    /// then on carries LoadFailed and LoadError, as state), whenLoadComplete
    /// resolves to false (it never rejects), and nothing more is read - what
    /// goes by the whole document goes by what was read: the layout completes
    /// over it, getText, find and getPageCount answer for it. It cannot be
    /// saved: exportDocumentPackage and exportToDocx throw.
    let isLoadComplete (editor: Editor) : bool =
        editor.Loading.IsNone && editor.LoadFailure.IsNone

    /// Whether the document on display could not be read to its end (see isLoadComplete).
    let isLoadFailed (editor: Editor) : bool =
        editor.LoadFailure.IsSome

    /// Why the document on display could not be read to its end; null when it was, or is being, read.
    let getLoadError (editor: Editor) : string =
        Option.toObj editor.LoadFailure

    /// A promise that resolves once the document is all there - to true; at
    /// once when it is - or cannot be read on: to false (see isLoadFailed).
    let whenLoadComplete (editor: Editor) : JS.Promise<bool> =
        JS.Constructors.Promise.Create(fun resolve _ ->
            match editor.Loading with
            | None -> resolve editor.LoadFailure.IsNone
            | Some state ->
                editor.LoadWaiters.Add (fun () -> resolve editor.LoadFailure.IsNone)
                // (A package that was opened not to read on by itself is read on from here.)
                state.Package?Progressive?start () |> ignore)

    /// Read the rest of the document now.
    let completeLoad (editor: Editor) : unit =
        ensureLoaded editor

    /// Stop reading the document on display: it stays what was read so far,
    /// as a document that could not be read to its end (isLoadFailed is true,
    /// getLoadError says that the reading was cancelled, whenLoadComplete
    /// resolves to false, the listeners are told, it is not saved). The
    /// package reads no further slice and lets go of the file. False when
    /// there was nothing to cancel (the document is read, or failed).
    ///
    /// The editor cancels by itself the reading of a package it is done with:
    /// when another document is loaded (importFromDocxPackage, setContent ...)
    /// and when it is disposed of - unless a viewer or another editor still
    /// shows that package. A host that wants such a package read to its end
    /// regardless reads it first (completeLoad, or the package's own
    /// Progressive.complete()).
    let cancelLoad (editor: Editor) : bool =
        match editor.Loading with
        | Some state -> Docx.cancelProgressiveImport state
        | None -> false

    /// Call `callback` after every change of the document or the selection
    /// (typing, pointer, toolbar / API commands, undo, redo, loading a
    /// document). It receives { DocumentChanged, SelectionChanged, CanUndo,
    /// CanRedo, Revision, Modified, LayoutChanged, LayoutComplete, LoadChanged,
    /// LoadComplete } (see getRevision / isModified / isLayoutComplete /
    /// isLoadComplete; LayoutChanged is set when the call says that more of a
    /// long document is laid out, LoadChanged when it says that more of it
    /// was read, and nothing else changed). Returns a function that removes
    /// the subscription.
    let onChange (editor: Editor) (callback: obj -> unit) : unit -> unit =
        editor.Listeners.Add callback
        editor.NotifiedCanUndo <- Inp.canUndo editor.Handler
        editor.NotifiedCanRedo <- Inp.canRedo editor.Handler
        fun () -> editor.Listeners.Remove callback |> ignore

    /// The revision of the document's content: a number that changes with
    /// every change of the content (not with the caret or the selection),
    /// never repeats for different content, and comes back when undo or redo
    /// brings the content back. Undoing to the state that was saved gives the
    /// revision that was saved.
    let getRevision (editor: Editor) : int = editor.Revision

    /// Remember the current content as the saved one: isModified is false
    /// until the content differs from it (and false again after undoing or
    /// redoing back to it).
    ///
    /// With a revision it is that revision that counts as saved - the one a
    /// save wrote, read before the save began:
    ///   const r = Word.getRevision(editor);
    ///   const bytes = await exportDocx(editor); await upload(bytes);
    ///   Word.markSaved(editor, r);
    /// What the user typed while the upload ran then stays unsaved
    /// (isModified is true), and undoing it gives the saved state. Without a
    /// revision (or with null) the current content is the saved one.
    let markSaved (editor: Editor) (revision: obj) : unit =
        editor.SavedRevision <-
            if isNull revision || jsTypeof revision <> "number" || System.Double.IsNaN (unbox<float> revision) then editor.Revision
            else int (unbox<float> revision)

    /// Whether the content differs from what was last loaded or marked saved.
    let isModified (editor: Editor) : bool =
        editor.Revision <> editor.SavedRevision

    let private formatValueToJs (convert: 'T -> obj) (value: WordStyles.FormatValue<'T>) : obj =
        match value with
        | WordStyles.NoValue -> null
        | WordStyles.Uniform item -> convert item
        | WordStyles.Mixed -> box "mixed"

    let private optionalStringToJs (value: string option) : obj =
        match value with
        | Some text -> box text
        | None -> null

    /// The selection format last worked out, and the document it is of.
    let mutable private lastSelectionFormat : (Doc.Document * WordStyles.SelectionFormat) option = None

    /// The formatting of the current selection, for toolbar state. Each value
    /// is the effective one (styles resolved), "mixed" when the selection
    /// holds different values, or null when there is nothing to report.
    ///   Bold, Italic, Underline, Strike: boolean
    ///   FontFamily: string (null: the default face)
    ///   FontSize: number, in layout pixels (CSS pixels at 100%; the unit of setFontSize)
    ///   FontSizePt: the same size in points (what a document states and a
    ///     size box shows: 11, 12, 14 ...; the unit of setFontSizePt)
    ///   Hyperlink: the target of the link the caret / selection is in (null: none)
    ///   Color: "#RRGGBB" (null: automatic), Highlight: colour name / value (null: none)
    ///   Paragraph: { StyleId (null: default style; never an internal entry such as "__docDefaults"),
    ///     HeadingLevel (0: body text),
    ///     Alignment "left" | "center" | "right" | "justify",
    ///     ListKind "none" | "bullet" | "numbered", ListLevel, IndentLeft, IndentRight, FirstLineIndent,
    ///     LineSpacing (a multiple of single spacing - what setLineSpacing takes - for the rule
    ///       "auto"; layout pixels for "exact" / "atLeast"; null: not stated, single spacing),
    ///     LineSpacingRule "auto" | "exact" | "atLeast" | null,
    ///     SpaceBefore, SpaceAfter (points, what setSpacingBefore / setSpacingAfter take) }
    ///   HasSelection, IsCollapsed, CanUndo, CanRedo: boolean
    let getSelectionFormat (editor: Editor) =
        // (Kept for the document it was read off - content and selection: a host
        // with several listeners asks more than once after a change.)
        let format =
            match lastSelectionFormat with
            | Some (doc, format) when obj.ReferenceEquals(doc, editor.Doc) -> format
            | _ ->
                let format = WordStyles.selectionFormat editor.Doc
                lastSelectionFormat <- Some (editor.Doc, format)
                format
        let boolean = formatValueToJs (fun (value: bool) -> box value)
        let number = formatValueToJs (fun (value: float) -> box value)
        let text = formatValueToJs optionalStringToJs
        createObj [
            "Bold" ==> boolean format.Bold
            "Italic" ==> boolean format.Italic
            "Underline" ==> boolean format.Underline
            "Strike" ==> boolean format.Strike
            "FontFamily" ==> text format.FontFamily
            "FontSize" ==> number format.FontSize
            "FontSizePt" ==> formatValueToJs (fun (pixels: float) -> box (System.Math.Round (pixels * 72.0 / 96.0 * 100.0) / 100.0)) format.FontSize
            "Hyperlink" ==>
                (match editor.Doc.Selection with
                 | Some selection ->
                    match Doc.findParagraph editor.Doc selection.Focus.BlockId |> Option.bind (fun para -> Doc.hyperlinkAt para (if selection.Anchor.BlockId = selection.Focus.BlockId then min selection.Anchor.Offset selection.Focus.Offset else selection.Focus.Offset)) with
                    | Some (target, _, _) -> box target
                    | None -> null
                 | None -> null)
            "Color" ==> text format.Color
            "Highlight" ==> text format.Highlight
            "Paragraph" ==> createObj [
                "StyleId" ==> text format.StyleId
                "HeadingLevel" ==> formatValueToJs (fun (level: int) -> box level) format.HeadingLevel
                "Alignment" ==> formatValueToJs (fun alignment ->
                                    match alignment with
                                    | Doc.Left -> box "left"
                                    | Doc.Center -> box "center"
                                    | Doc.Right -> box "right"
                                    | Doc.Justify -> box "justify") format.Alignment
                "ListKind" ==> formatValueToJs (fun kind ->
                                    match kind with
                                    | Doc.NoList -> box "none"
                                    | Doc.Bullet -> box "bullet"
                                    | Doc.Numbered -> box "numbered") format.ListKind
                "ListLevel" ==> formatValueToJs (fun (level: int) -> box level) format.ListLevel
                "IndentLeft" ==> number format.IndentLeft
                "IndentRight" ==> number format.IndentRight
                "FirstLineIndent" ==> number format.FirstLineIndent
                "LineSpacing" ==> formatValueToJs (fun (spacing: float option) -> match spacing with Some value -> box value | None -> null) format.ParagraphLineSpacing
                "LineSpacingRule" ==> text format.ParagraphLineSpacingRule
                "SpaceBefore" ==> formatValueToJs (fun (pixels: float) -> box (System.Math.Round (pixels * 72.0 / 96.0 * 100.0) / 100.0)) format.ParagraphSpacingBefore
                "SpaceAfter" ==> formatValueToJs (fun (pixels: float) -> box (System.Math.Round (pixels * 72.0 / 96.0 * 100.0) / 100.0)) format.ParagraphSpacingAfter
            ]
            "HasSelection" ==> editor.Doc.Selection.IsSome
            "IsCollapsed" ==> (match editor.Doc.Selection with Some selection -> selection.IsCollapsed | None -> true)
            "CanUndo" ==> Inp.canUndo editor.Handler
            "CanRedo" ==> Inp.canRedo editor.Handler
        ]

    let initializeCanvasWithChange (canvas: HTMLCanvasElement) (afterChange: unit -> unit) =
        let doc = Doc.empty
        let mutable editor = Unchecked.defaultof<Editor>

        let onChange newDoc =
            // (Revisions and what the listeners hear are about the whole document: entering a
            // header or footer changes the selection, not the content.)
            let previous = editor.Whole
            editor.Doc <- newDoc
            let whole = Inp.wholeDocument editor.Handler newDoc
            editor.Whole <- whole
            observeRevision editor whole
            afterChange ()
            notifyListeners editor previous whole
            continueLayout editor
            if editor.Handler.Story.IsSome then storyFieldsRefresh.Value editor

        let handler = Inp.create canvas doc onChange

        handler.OnDoubleClick <- fun x y -> doubleClick.Value editor x y
        handler.OnHistoryPlace <- fun pageIndex isHeader -> enterStoryOfPage.Value editor pageIndex isHeader
        editor <- {
            Handler = handler
            Doc = doc
            Listeners = ResizeArray()
            NotifiedCanUndo = false
            NotifiedCanRedo = false
            Revision = 0
            LastRevision = 0
            SavedRevision = 0
            RevisionDocument = doc
            KnownRevisions = ResizeArray()
            FontFamilies = [||]
            Whole = doc
            Loading = None
            StopLoading = ignore
            LoadNotified = 0.0
            LoadWaiters = ResizeArray()
            LoadFailure = None
        }
        handler.EnsureWhole <- fun () -> ensureLoaded editor

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

    let private pageFramesToJs (ctx: Ren.RenderContext) (pages: Layout.PageInfo list) = Viewer.pageFramesToJs ctx pages

    let private numberField (value: obj) (name: string) (pascal: string) fallback =
        let pick (key: string) =
            let field = value?(key)
            if isNull field || jsTypeof field <> "number" || System.Double.IsNaN (unbox<float> field) then None else Some (unbox<float> field)
        pick name |> Option.orElse (pick pascal) |> Option.defaultValue fallback

    /// { scale, devicePixelRatio, scrollTop, scrollLeft, width, height, originX, originY } -> Viewport.
    /// null / undefined (or a viewport without a size) means whole-canvas mode.
    let private viewportFromJs (value: obj) : Ren.Viewport option =
        if isNull value || jsTypeof value <> "object" then None
        else
            let width = numberField value "width" "Width" 0.0
            let height = numberField value "height" "Height" 0.0
            if width <= 0.0 || height <= 0.0 then None
            else
                Some {
                    Scale = max 0.01 (numberField value "scale" "Scale" 1.0)
                    DevicePixelRatio = max 0.1 (numberField value "devicePixelRatio" "DevicePixelRatio" 1.0)
                    ScrollLeft = numberField value "scrollLeft" "ScrollLeft" 0.0
                    ScrollTop = numberField value "scrollTop" "ScrollTop" 0.0
                    Width = width
                    Height = height
                    OriginX = numberField value "originX" "OriginX" 0.0
                    OriginY = numberField value "originY" "OriginY" 0.0
                }

    let private viewportToJs (viewport: Ren.Viewport option) : obj =
        match viewport with
        | None -> null
        | Some vp ->
            createObj [
                "scale" ==> vp.Scale
                "devicePixelRatio" ==> vp.DevicePixelRatio
                "scrollLeft" ==> vp.ScrollLeft
                "scrollTop" ==> vp.ScrollTop
                "width" ==> vp.Width
                "height" ==> vp.Height
                "originX" ==> vp.OriginX
                "originY" ==> vp.OriginY
            ]

    let private documentSizeToJs (ctx: Ren.RenderContext) (doc: Doc.Document) = Viewer.documentSizeToJs ctx doc

    /// Show the document through a viewport: the canvas becomes exactly as
    /// large as the visible area (the engine sizes it: backing store
    /// width x height x devicePixelRatio, CSS size width x height), and only
    /// the pages and paint commands that intersect the visible band are drawn,
    /// at scale x devicePixelRatio. Call it on scroll, resize and zoom.
    ///
    /// `viewport` is { scale, devicePixelRatio, scrollTop, scrollLeft, width,
    /// height, originX?, originY? }, all in CSS pixels: scrollTop / scrollLeft
    /// are the scroll position of the host's scrollable content, originX /
    /// originY where the document's top-left corner sits in that content (a
    /// margin, or the offset that centres the pages; default 0). Pointer
    /// events, the caret and the IME proxy follow the transform. Passing null
    /// goes back to whole-canvas mode, where the host sizes the canvas from
    /// getPageFrames. Returns the document size (see getDocumentSize).
    let setViewport (editor: Editor) (viewport: obj) =
        let ctx = editor.Handler.RenderContext
        Ren.setViewport ctx (viewportFromJs viewport)
        Inp.repaint editor.Handler
        continueLayout editor
        documentSizeToJs ctx editor.Doc

    /// The viewport set with setViewport, or null in whole-canvas mode.
    let getViewport (editor: Editor) =
        viewportToJs editor.Handler.RenderContext.Viewport

    /// Total size of the document: { Width, Height } in layout pixels (widest
    /// page, bottom of the last page), { ScaledWidth, ScaledHeight } in CSS
    /// pixels at the viewport's scale (what a scroll spacer needs, plus the
    /// host's own origin / margins), and PageCount.
    let getDocumentSize (editor: Editor) =
        documentSizeToJs editor.Handler.RenderContext editor.Doc

    /// Layout (document) coordinates of a point given in CSS pixels relative
    /// to the canvas' top-left corner. One to one in whole-canvas mode.
    let viewportToDocument (editor: Editor) (x: float) (y: float) =
        let layoutX, layoutY = Ren.viewportToLayout editor.Handler.RenderContext x y
        createObj [ "X" ==> layoutX; "Y" ==> layoutY ]

    /// CSS pixels relative to the canvas' top-left corner of a point in layout coordinates.
    let documentToViewport (editor: Editor) (x: float) (y: float) =
        let viewX, viewY = Ren.layoutToViewport editor.Handler.RenderContext x y
        createObj [ "X" ==> viewX; "Y" ==> viewY ]

    /// Where the caret is: { X, Y, Height } in layout coordinates and
    /// { ViewportX, ViewportY, ViewportHeight } in CSS pixels relative to the
    /// canvas. null without a selection.
    let getCaretRect (editor: Editor) : obj =
        let ctx = editor.Handler.RenderContext
        match Ren.caretRect ctx editor.Doc with
        | None -> null
        | Some (x, top, height) ->
            let viewX, viewY = Ren.layoutToViewport ctx x top
            createObj [
                "X" ==> x
                "Y" ==> top
                "Height" ==> height
                "ViewportX" ==> viewX
                "ViewportY" ==> viewY
                "ViewportHeight" ==> height * Ren.viewportScale ctx
            ]

    /// Give the editor the decoded images of the document: a Map of image
    /// source (as in the editor package) -> HTMLImageElement. Images are drawn
    /// synchronously from it. Repaints.
    let setImageCache (editor: Editor) (cache: obj) =
        editor.Handler.RenderContext.ImageCache <- (if isNull cache then None else Some cache)
        editor.Handler.RenderContext.LastPaint <- None
        Inp.repaint editor.Handler

    /// Paint everything again (the host resized or cleared the canvas, or
    /// images finished loading).
    let repaint (editor: Editor) =
        editor.Handler.RenderContext.LastPaint <- None
        Inp.repaint editor.Handler
        continueLayout editor

    /// Paint one page of the editor's document onto another canvas (thumbnails,
    /// print preview). The canvas is resized to the page at `scale` backing
    /// pixels per layout pixel. Read-only: no selection or caret. Returns false
    /// when there is no such page.
    let renderPage (editor: Editor) (canvas: HTMLCanvasElement) (pageIndex: int) (scale: float) =
        // (A page there is not yet: of the whole document.)
        if editor.Loading.IsSome && pageIndex >= List.length (Ren.pageFrames editor.Handler.RenderContext editor.Doc) then ensureLoaded editor
        Ren.renderPage editor.Handler.RenderContext editor.Doc canvas pageIndex scale

    /// Create a read-only viewer (see `Viewer.createViewer`, which this is:
    /// a page that only shows documents imports `WordViewer.js` instead of
    /// this module and leaves the editor out of its bundle).
    let createViewer (canvas: HTMLCanvasElement) = Viewer.createViewer canvas

    /// Get document text
    let getText (editor: Editor) =
        ensureLoaded editor
        Doc.getText editor.Doc

    let getTrackChanges (editor: Editor) =
        editor.Handler.TrackChanges

    let setTrackChanges (editor: Editor) enabled =
        editor.Handler.TrackChanges <- enabled

    let setTrackChangesOptions (editor: Editor) author date =
        editor.Handler.TrackChangesAuthor <- if isNull (box author) || author = "" then "RedBat Word" else author
        editor.Handler.TrackChangesDate <- if isNull (box date) then "" else date

    /// Measure text with `measureText(text, style)` instead of canvas text.
    /// A plain JavaScript function. The measurers are part of what a layout
    /// and a painted frame were made with: installing or clearing them lays
    /// the document out again; nothing about the function itself is compared
    /// afterwards.
    let setTextMeasurer (editor: Editor) (measureText: Layout.TextMeasurer) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextMeasurers ctx (if isNull (box measureText) then None else Some measureText) ctx.TextRangeMeasurer
        Ren.render ctx editor.Doc
        continueLayout editor

    /// As setTextMeasurer, with `measureRange(text, style, start, length)`
    /// for the width of a part of a shaped run.
    let setTextMeasurers (editor: Editor) (measureText: Layout.TextMeasurer) (measureRange: Layout.TextRangeMeasurer) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextMeasurers ctx (if isNull (box measureText) then None else Some measureText) (if isNull (box measureRange) then None else Some measureRange)
        Ren.render ctx editor.Doc
        continueLayout editor

    let setTextRangeMeasurer (editor: Editor) (measureRange: Layout.TextRangeMeasurer) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextMeasurers ctx ctx.TextMeasurer (if isNull (box measureRange) then None else Some measureRange)
        Ren.render ctx editor.Doc
        continueLayout editor

    /// Paint text with `paintText(text, style, x, baseline, width)` - a plain
    /// JavaScript function returning true when it painted the text - instead
    /// of canvas text, on the editor's own canvas. The painted frame is built
    /// again when a painter is installed or cleared, and at no other time:
    /// scrolling and caret moves reuse it whatever function object the
    /// painter is (see getRenderStats).
    let setTextPathPainter (editor: Editor) (paintText: string -> Paint.TextPaintStyle -> float -> float -> float -> bool) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextPathPainter ctx (if isNull (box paintText) then None else Some paintText) None
        Ren.render ctx editor.Doc
        continueLayout editor

    /// As setTextPathPainter, for every canvas the editor paints on:
    /// `painterFor(context2d)` returns the painter for that canvas context, so
    /// renderPage paints its pages with it too (a painter draws on one context).
    let setTextPathPainterFactory (editor: Editor) (painterFor: System.Func<CanvasRenderingContext2D, System.Func<string, Paint.TextPaintStyle, float, float, float, bool>>) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextPathPainterFactory ctx painterFor
        Ren.render ctx editor.Doc
        continueLayout editor

    let clearTextPathPainter (editor: Editor) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextPathPainter ctx None None
        Ren.render ctx editor.Doc
        continueLayout editor

    let clearTextMeasurer (editor: Editor) =
        let ctx = editor.Handler.RenderContext
        Ren.setTextPathPainter ctx None None
        Ren.setTextMeasurers ctx None None
        Ren.render ctx editor.Doc
        continueLayout editor

    /// Counters of the editor's renderer, for hosts and checks:
    /// { FrameBuilds (times a paint frame was built: the document was laid
    ///   out and its paint commands collected - a caret move builds none),
    ///   LaidOutBlocks / CollectedBlocks (of the last frame built: the
    ///   top-level blocks that were laid out / whose paint commands were
    ///   collected again; the others kept what they had), CollectedOnPaint
    ///   (paragraphs whose commands were collected when they were first
    ///   painted, so far), Blocks (top-level blocks of the document),
    ///   PainterVersion, MeasurerVersion (times a text path painter / the
    ///   measurers were installed or cleared) }.
    let getRenderStats (editor: Editor) =
        Ren.renderStats editor.Handler.RenderContext editor.Doc

    /// The font faces a document's text is set in: [{ family, bold, italic }],
    /// after style resolution, with the faces the engine itself introduces (a
    /// list marker in Symbol, Wingdings or Courier New; note text; headers
    /// and footers). `target` is an editor (its current document) or an
    /// editor package (Office.Formats `forEditorPackage`) that is not shown
    /// yet - so a host fetches exactly these faces before it installs its
    /// measurers and loads the document. Of a package that is still being
    /// read (`word.openProgressive`) it is the faces of the styles and of
    /// what is read so far - see onFontsDiscovered for the rest; of an editor
    /// whose document is still being read, the rest is read first.
    let getRequiredFonts (target: obj) : obj array =
        if isNull target then [||]
        elif not (isNull (target?Handler)) then
            let editor = unbox<Editor> target
            ensureLoaded editor
            Ren.requiredFonts editor.Handler.RenderContext editor.Doc
        else
            Ren.requiredFontsOfDocument (Docx.importFromEditorPackage target)

    /// The faces a package that is still being read turns out to need as it
    /// is read on: see Viewer.onFontsDiscovered.
    let onFontsDiscovered (package: obj) (callback: obj array -> unit) : unit -> unit =
        Viewer.onFontsDiscovered package callback

    let getCachedLayouts (editor: Editor) =
        ensureLoaded editor
        Ren.completeLayout editor.Handler.RenderContext
        Ren.cachedLayouts editor.Handler.RenderContext editor.Doc

    let toggleTrackChanges (editor: Editor) =
        editor.Handler.TrackChanges <- not editor.Handler.TrackChanges
        editor.Handler.TrackChanges

    let getCommentAnchors (editor: Editor) =
        ensureLoaded editor
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
        ensureLoaded editor
        let blockId = anchor?BlockId |> unbox<string>
        let startOffset = anchor?StartOffset |> unbox<int>
        let endOffset = anchor?EndOffset |> unbox<int>
        let newDoc = Doc.setSelection { BlockId = blockId; Offset = startOffset } { BlockId = blockId; Offset = endOffset } editor.Doc
        commitDocumentUpdate editor newDoc

    /// Select from one document position to another (the same position twice
    /// places the caret). Positions are (block id, offset) as getSelection,
    /// hitTestPosition and getLayoutLines report them.
    let setSelection (editor: Editor) (anchorBlockId: string) (anchorOffset: int) (focusBlockId: string) (focusOffset: int) =
        let newDoc = Doc.setSelection { BlockId = anchorBlockId; Offset = anchorOffset } { BlockId = focusBlockId; Offset = focusOffset } editor.Doc
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
        ensureLoaded editor
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
        if pagesInUse.Length = 0 then int (floor (max 0.0 y / pageStride doc)) + 1
        else pageIndexAt y + 1

    let private pageCountFromLines (doc: Doc.Document) (layouts: Layout.BlockLayout list) =
        let mutable pageCount = 1
        for layout in layouts do
            for line in layout.Lines do
                pageCount <- max pageCount (pageForYInDoc doc (line.Y + line.Height))
        pageCount

    let private pageTopForYInDoc (doc: Doc.Document) y =
        if pagesInUse.Length = 0 then floor (max 0.0 y / pageStride doc) * pageStride doc
        else pagesInUse.[pageIndexAt y].Top

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
                // (A table in a cell: its paragraphs come where the table is
                // among the cell's, with the row and column they have in it;
                // the table's number is that of the top-level table.)
                let rec addTable (table: Doc.TableBlock) =
                    let maxRows = table.Rows.Length
                    let maxColumns = maxVisualColumns table
                    table.Rows
                    |> List.iteri (fun rowIndex row ->
                        let mutable columnIndex = 1
                        for cell in row.Cells do
                            let span = max 1 cell.ColSpan
                            let context = {
                                ParagraphIndex = 0
                                TableIndex = Some tableIndex
                                RowIndex = Some (rowIndex + 1)
                                ColumnIndex = Some columnIndex
                                EndColumnIndex = Some (columnIndex + span - 1)
                                MaxRows = Some maxRows
                                MaxColumns = Some maxColumns
                                InTextBox = false
                            }
                            if cell.NestedTables.IsEmpty then
                                for para in cell.Content do addContext para context
                            else
                                let count = cell.Content.Length
                                let anchors = cell.NestedTableAnchors |> List.toArray
                                let nested =
                                    cell.NestedTables
                                    |> List.mapi (fun i inner -> inner, (if i < anchors.Length then anchors.[i] else count) |> max 0 |> min count)
                                    |> List.sortBy snd
                                cell.Content
                                |> List.iteri (fun index para ->
                                    for (inner, anchor) in nested do
                                        if anchor = index then addTable inner
                                    addContext para context)
                                for (inner, anchor) in nested do
                                    if anchor = count then addTable inner
                            columnIndex <- columnIndex + span)
                addTable table
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
        ensureLoaded editor
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
        ensureLoaded editor
        notePages editor
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
        ensureLoaded editor
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
        ensureLoaded editor
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
        ensureLoaded editor
        match editor.Doc.Selection with
        | None -> null
        | Some selection ->
            getRangeInformation editor selection.Focus.BlockId selection.Focus.Offset selection.Focus.Offset

    let getLayoutLines (editor: Editor) =
        ensureLoaded editor
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

    /// The laid-out document block by block, for comparing two layouts (see
    /// LayoutFingerprint): [{ Page, Y, Height, LineCount, LineHeights, Text,
    /// Paragraph, Runs }] with the resolved paragraph and run formatting as
    /// strings. Lays the whole document out.
    let getLayoutFingerprint (editor: Editor) =
        ensureLoaded editor
        let layouts = layoutsForEditor editor
        LayoutFingerprint.ofLayouts (pageForYInDoc editor.Doc) layouts
        |> List.map (fun block ->
            createObj [
                "Page" ==> block.Page
                "Y" ==> block.Y
                "Height" ==> block.Height
                "LineCount" ==> block.LineCount
                "LineHeights" ==> block.LineHeights
                "Text" ==> block.Text
                "Paragraph" ==> block.Paragraph
                "Runs" ==> block.Runs
            ])
        |> List.toArray

    let getVbaLayoutLines (editor: Editor) =
        ensureLoaded editor
        getLayoutLines editor

    /// Find literal text matches across editable paragraphs.
    let findText (editor: Editor) (query: string) (matchCase: bool) =
        ensureLoaded editor
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
        ensureLoaded editor
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
        ensureLoaded editor
        let matchInfo: Doc.FindMatch = {
            BlockId = matchObj?BlockId |> unbox<string>
            Start = matchObj?Start |> unbox<int>
            End = matchObj?End |> unbox<int>
            Text = matchObj?Text |> unbox<string>
        }
        let newDoc = Doc.selectFindMatch matchInfo editor.Doc
        commitDocumentUpdate editor newDoc

    let selectRevisionInfo (editor: Editor) (revisionObj: obj) =
        ensureLoaded editor
        let blockId = revisionObj?BlockId |> unbox<string>
        let startOffset = revisionObj?Start |> unbox<int>
        let endOffset = revisionObj?End |> unbox<int>
        let newDoc = Doc.setSelection { BlockId = blockId; Offset = startOffset } { BlockId = blockId; Offset = endOffset } editor.Doc
        commitDocumentUpdate editor newDoc

    /// Replace all literal text matches across editable paragraphs.
    let replaceAllText (editor: Editor) (query: string) (replacement: string) (matchCase: bool) =
        ensureLoaded editor
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
        // Numbering restarts on the first page of the section that asks for it.
        let restarts = Doc.pageNumberRestarts editor.Doc
        let restartPoints =
            layouts
            |> List.choose (fun layout ->
                restarts
                |> List.tryFind (fun (blockId, paragraphId, _) -> blockId = layout.BlockId || paragraphId = Some layout.BlockId)
                |> Option.map (fun (_, _, start) -> pageNumberForY editor.Doc layout.Y, max 1 start))
            |> List.distinct
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
        // (The pages of the notes after the body are pages of the document: what getPageCount says.)
        let mutable pageCount = max 1 (Ren.pageCount editor.Handler.RenderContext editor.Doc)
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
        ensureLoaded editor
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
        ensureLoaded editor
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
        ensureLoaded editor
        let newDoc = Doc.selectBookmark bookmarkId editor.Doc
        commitDocumentUpdate editor newDoc

    let removeBookmark (editor: Editor) (bookmarkId: string) =
        ensureLoaded editor
        let newDoc = Doc.removeBookmark bookmarkId editor.Doc
        commitDocumentUpdate editor newDoc

    let renameBookmark (editor: Editor) (bookmarkId: string) (bookmarkName: string) =
        ensureLoaded editor
        let newDoc = Doc.renameBookmark bookmarkId bookmarkName editor.Doc
        commitDocumentUpdate editor newDoc

    let getContentControlInfos (editor: Editor) =
        ensureLoaded editor
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
        ensureLoaded editor
        let newDoc = Doc.selectContentControl controlId editor.Doc
        commitDocumentUpdate editor newDoc

    let removeContentControl (editor: Editor) (controlId: string) =
        ensureLoaded editor
        let newDoc = Doc.removeContentControl controlId editor.Doc
        commitDocumentUpdate editor newDoc

    let updateContentControl (editor: Editor) (controlId: string) (newControlId: string) (tag: string) (alias: string) =
        ensureLoaded editor
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
        ensureLoaded editor
        let newDoc = Doc.acceptAllRevisions editor.Doc
        commitDocumentUpdate editor newDoc

    let rejectAllRevisions (editor: Editor) =
        ensureLoaded editor
        let newDoc = Doc.rejectAllRevisions editor.Doc
        commitDocumentUpdate editor newDoc

    let selectNextRevision (editor: Editor) =
        ensureLoaded editor
        let newDoc = Doc.selectNextRevision editor.Doc
        commitDocumentUpdate editor newDoc

    let selectPreviousRevision (editor: Editor) =
        ensureLoaded editor
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

    /// Set the font size of the selection in points (11, 12, 14 ...: the unit
    /// documents state sizes in and a size box shows). setFontSize takes
    /// layout pixels (points x 96 / 72).
    let setFontSizePt (editor: Editor) (points: float) =
        setFontSizeFor editor (points * 96.0 / 72.0)

    /// Tell the editor which font families the host's text stack can draw
    /// (getFontFamilies then offers them).
    let setFontFamilies (editor: Editor) (families: string array) =
        editor.FontFamilies <- (if isNull (box families) then [||] else families |> Array.filter (fun name -> not (System.String.IsNullOrWhiteSpace name)))

    /// The font families to offer in a font box: the ones the text stack
    /// installed (setFontFamilies) and the ones the document uses (its text
    /// and its styles), sorted, each once. Without a text stack and with an
    /// empty document: the editor's default face.
    let getFontFamilies (editor: Editor) : string array =
        let used = System.Collections.Generic.HashSet<string>()
        let names = ResizeArray<string>()
        let add (name: string) =
            if not (System.String.IsNullOrWhiteSpace name) && used.Add (name.Trim().ToLowerInvariant()) then names.Add (name.Trim())
        for name in editor.FontFamilies do add name
        for style in editor.Doc.Styles do
            match style.Text |> Option.bind (fun text -> text.FontFamily) with
            | Some name -> add name
            | None -> ()
        for para in Doc.allParagraphs editor.Doc do
            for run in para.Buffer.Runs do
                match run with
                | TextBuf.Text text ->
                    match text.Style.FontFamily with
                    | Some name -> add name
                    | None -> ()
                | _ -> ()
        if names.Count = 0 then add editor.Handler.RenderContext.FontFamily
        names.ToArray() |> Array.sortBy (fun name -> name.ToLowerInvariant())

    /// The document's styles, for a style box: [{ Id, Name, Type ("paragraph" |
    /// "character" | "table" | "numbering"), IsDefault, IsQuickStyle, Hidden,
    /// HeadingLevel (1-9, or 0), BasedOn (id or null) }]. setParagraphStyle /
    /// setCharacterStyle take the Id.
    let getStyles (editor: Editor) : obj array =
        // "heading 3" / "Heading3" is a level-3 heading whatever its outline level says.
        let headingLevel (style: Doc.NamedStyle) =
            let name = (style.Name |> Option.defaultValue style.Id).Trim().ToLowerInvariant().Replace(" ", "")
            if name.StartsWith "heading" && name.Length = 8 && name.[7] >= '1' && name.[7] <= '9' then int name.[7] - int '0'
            else
                match style.OutlineLevel with
                // The file's outline level counts from 0.
                | Some level when level >= 0 && level <= 8 && style.StyleType = "paragraph" -> level + 1
                | _ -> 0
        editor.Doc.Styles
        // (The document defaults are kept among the styles - as "__docDefaults",
        // whatever type the file's reader gave the entry; they are not a style.)
        |> List.filter (fun style -> style.StyleType <> "defaults" && not (WordStyles.isInternalStyleId style.Id))
        |> List.map (fun style ->
            createObj [
                "Id" ==> style.Id
                "Name" ==> (style.Name |> Option.defaultValue style.Id)
                "Type" ==> style.StyleType
                "IsDefault" ==> style.IsDefault
                "IsQuickStyle" ==> style.IsQuickStyle
                "Hidden" ==> (style.SemiHidden && not style.UnhideWhenUsed)
                "HeadingLevel" ==> headingLevel style
                "BasedOn" ==> optionalStringToJs style.BasedOn
            ])
        |> List.toArray

    /// Line spacing of the selected paragraphs as a multiple of single
    /// spacing (1, 1.15, 1.5, 2 ...); 0 or less goes back to the style's.
    let setLineSpacing (editor: Editor) (multiple: float) =
        commitDocumentUpdate editor (Doc.setLineSpacing (if multiple > 0.0 then Some multiple else None) editor.Doc)

    /// Space before / after the selected paragraphs, in points; a negative
    /// value goes back to the style's.
    let setSpacingBefore (editor: Editor) (points: float) =
        commitDocumentUpdate editor (Doc.setSpacingBefore (if points >= 0.0 then Some (points * 96.0 / 72.0) else None) editor.Doc)

    let setSpacingAfter (editor: Editor) (points: float) =
        commitDocumentUpdate editor (Doc.setSpacingAfter (if points >= 0.0 then Some (points * 96.0 / 72.0) else None) editor.Doc)

    /// Insert a page break at the caret (Ctrl+Enter in Word).
    let insertPageBreak (editor: Editor) =
        commitDocumentUpdate editor (Doc.insertPageBreak editor.Doc)

    /// Insert a hyperlink to `target`: the selected text becomes the link;
    /// with a caret, `text` (or the target, when it is empty) is inserted as
    /// the link.
    let insertHyperlink (editor: Editor) (target: string) (text: string) =
        if not (isNull (box target)) && target <> "" then
            commitDocumentUpdate editor (Doc.insertHyperlink target (if isNull (box text) then "" else text) editor.Doc)

    /// The hyperlink the caret is in: { Target, BlockId, StartOffset,
    /// EndOffset, Text }, or null.
    let getHyperlink (editor: Editor) : obj =
        match editor.Doc.Selection with
        | None -> null
        | Some selection ->
            match Doc.findParagraph editor.Doc selection.Focus.BlockId with
            | None -> null
            | Some para ->
                let offset = if selection.Anchor.BlockId = selection.Focus.BlockId then min selection.Anchor.Offset selection.Focus.Offset else selection.Focus.Offset
                match Doc.hyperlinkAt para offset with
                | Some (target, first, last) ->
                    createObj [
                        "Target" ==> target
                        "BlockId" ==> para.Id
                        "StartOffset" ==> first
                        "EndOffset" ==> last
                        "Text" ==> (TextBuf.toString para.Buffer).Substring(first, last - first)
                    ]
                | None -> null

    /// Change the target of the hyperlink the caret is in (the whole link);
    /// null or "" removes the link and keeps its text.
    let setHyperlinkTarget (editor: Editor) (target: string) =
        commitDocumentUpdate editor (Doc.setHyperlinkTarget (if isNull (box target) || target = "" then None else Some target) editor.Doc)

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
        ensureLoaded editor
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

    /// Insert a table row above the row containing the current selection
    let insertTableRowBefore (editor: Editor) =
        commitDocumentUpdate editor (Doc.insertTableRowBeforeSelection editor.Doc)

    /// Insert a table column to the left of the column containing the current selection
    let insertTableColumnBefore (editor: Editor) =
        commitDocumentUpdate editor (Doc.insertTableColumnBeforeSelection editor.Doc)

    /// Whether the selection spans several cells of one table (mergeTableCells would merge them).
    let canMergeTableCells (editor: Editor) : bool =
        Doc.canMergeSelectedTableCells editor.Doc

    /// Merge the selected table cells into one.
    let mergeTableCells (editor: Editor) =
        commitDocumentUpdate editor (Doc.mergeSelectedTableCells editor.Doc)

    /// Split the merged cell the caret is in back into single cells.
    let splitTableCell (editor: Editor) =
        commitDocumentUpdate editor (Doc.splitSelectedTableCell editor.Doc)

    /// Delete the table row containing the current selection
    let deleteTableRow (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableRow editor.Doc
        commitDocumentUpdate editor newDoc

    /// Delete the table column containing the current selection
    let deleteTableColumn (editor: Editor) =
        let newDoc = Doc.deleteSelectedTableColumn editor.Doc
        commitDocumentUpdate editor newDoc

    /// Insert a table of `rows` x `columns` after the block the caret is in.
    /// It has single thin borders (and refers to the document's "Table Grid"
    /// or default table style when it has one), and the caret is in its first
    /// cell: getCaretRect, getSelectionInformation and getCaretContext answer
    /// at once, and Tab / Shift+Tab go from cell to cell.
    let insertTable (editor: Editor) (rows: int) (columns: int) =
        let newDoc = Doc.insertTable rows columns editor.Doc
        commitDocumentUpdate editor newDoc

    // ---- headers and footers edited in place ----------------------------------

    let private storyLabel (isHeader: bool) (kind: string) =
        (match kind with
         | "first" -> "First Page "
         | "even" -> "Even Page "
         | _ -> "")
        + (if isHeader then "Header" else "Footer")

    /// The paragraphs that end the sections before the last one, in order.
    let private sectionEnds (doc: Doc.Document) : Doc.ParagraphBlock list =
        doc.Blocks
        |> List.choose (function
            | Doc.Paragraph para
            | Doc.Heading (_, para) when para.Metadata.SectionBreak.IsSome -> Some para
            | _ -> None)

    /// The header (or footer) references a section states itself: (kind, relationship id).
    let private sectionReferences (doc: Doc.Document) (section: int) (isHeader: bool) : (string * string) list =
        match List.tryItem section (sectionEnds doc) with
        | Some para -> if isHeader then para.Metadata.HeaderReferences else para.Metadata.FooterReferences
        | None -> if isHeader then doc.PageSetup.HeaderReferences else doc.PageSetup.FooterReferences

    let private withSectionReferences (doc: Doc.Document) (section: int) (isHeader: bool) (references: (string * string) list) : Doc.Document =
        match List.tryItem section (sectionEnds doc) with
        | Some para ->
            let metadata = if isHeader then { para.Metadata with HeaderReferences = references } else { para.Metadata with FooterReferences = references }
            Doc.replaceParagraph doc para.Id { para with Metadata = metadata }
        | None ->
            { doc with PageSetup = (if isHeader then { doc.PageSetup with HeaderReferences = references } else { doc.PageSetup with FooterReferences = references }) }

    let private freshId (prefix: string) = prefix + System.Guid.NewGuid().ToString().Replace("-", "")

    /// A relationship id no header or footer of the document has.
    let private freshPartId (doc: Doc.Document) =
        let used = System.Collections.Generic.HashSet<string>()
        for part in doc.HeaderParts @ doc.FooterParts do
            match part.RelationshipId with
            | Some id -> used.Add id |> ignore
            | None -> ()
        let mutable number = 1
        while used.Contains (sprintf "rIdRbtHf%d" number) do
            number <- number + 1
        sprintf "rIdRbtHf%d" number

    /// The document with a new header (or footer) of `kind` for a section,
    /// holding `blocks` (none: one empty paragraph), and the part's place.
    let private withNewPart (doc: Doc.Document) (section: int) (isHeader: bool) (kind: string) (blocks: Doc.Block list) : Doc.Document * int =
        let blocks =
            if blocks.IsEmpty then
                let para : Doc.ParagraphBlock =
                    { Id = freshId (if isHeader then "header-" else "footer-")
                      Buffer = TextBuf.empty
                      Metadata = { Doc.defaultMetadata with Style = Some (if isHeader then "Header" else "Footer") } }
                [ Doc.Paragraph para ]
            else blocks
        let id = freshPartId doc
        let part : Doc.HeaderFooterPart = { Kind = kind; RelationshipId = Some id; Blocks = blocks }
        let references = (sectionReferences doc section isHeader |> List.filter (fun (referenceKind, _) -> referenceKind <> kind)) @ [ kind, id ]
        let doc = withSectionReferences doc section isHeader references
        if isHeader then { doc with HeaderParts = doc.HeaderParts @ [ part ]; Headers = [] }, doc.HeaderParts.Length
        else { doc with FooterParts = doc.FooterParts @ [ part ]; Footers = [] }, doc.FooterParts.Length

    /// The blocks of a header or footer as a part of their own: every block - paragraphs,
    /// tables (and the tables in their cells), text boxes - with an id of its own.
    let private copiedBlocks (isHeader: bool) (blocks: Doc.Block list) : Doc.Block list =
        let prefix = if isHeader then "header-" else "footer-"
        blocks |> Doc.copyBlocksWithNewIds (fun kind -> freshId (if kind = "paragraph" then prefix else prefix + kind + "-"))

    /// Back to the body from a header or footer that is edited in place
    /// (Escape and a double click on the body do the same). The body's
    /// selection is what it was. False when none was being edited.
    let exitHeaderFooter (editor: Editor) : bool =
        Inp.exitStory editor.Handler

    /// Edit the header or the footer (`kind`: "header" | "footer") of a page
    /// in place: the caret goes into it, the body is dimmed and a labelled
    /// line marks where it ends. Which header that is follows the page - its
    /// section, the first page of a section with a title page, an even page
    /// of a document that tells even and odd pages apart. Entering changes
    /// nothing in the document: where the page shows no header (none at all,
    /// or a title / even page whose section - and those before it - has none
    /// of that kind, which shows a blank one as in Word), the caret stands in
    /// an empty one that becomes a header of the document with the first
    /// edit made in it (that edit's undo step takes it away again). Typing,
    /// selection, formatting, tables, pictures, fields
    /// (insertPageNumberField / insertPageCountField /
    /// insertSectionPageCountField), undo and redo then work there: the
    /// document the commands act on is the header's own until
    /// exitHeaderFooter. getCaretContext says so (InHeaderFooter, Story,
    /// HeaderFooter). False when there is no such page.
    let enterHeaderFooter (editor: Editor) (pageIndex: int) (kind: string) : bool =
        ensureLoaded editor
        let isHeader = kind <> "footer"
        Inp.exitStory editor.Handler |> ignore
        let ctx = editor.Handler.RenderContext
        match Ren.headerFooterTarget ctx editor.Doc pageIndex isHeader with
        | None -> false
        | Some target ->
            let doc = editor.Doc
            let parts = if isHeader then doc.HeaderParts else doc.FooterParts
            let existing =
                match target.TargetPart with
                | Some part -> parts |> List.tryFindIndex (fun candidate -> obj.ReferenceEquals(candidate, part))
                | None -> None
            // The document the story's blocks belong in once they are edited, and the part's place there.
            let edited, partIndex =
                match existing with
                | Some index ->
                    match List.item index parts with
                    | part when part.Blocks.IsEmpty ->
                        // (A part without a paragraph: it gets one with the first edit.)
                        let para : Doc.ParagraphBlock =
                            { Id = freshId (if isHeader then "header-" else "footer-"); Buffer = TextBuf.empty
                              Metadata = { Doc.defaultMetadata with Style = Some (if isHeader then "Header" else "Footer") } }
                        let filled = parts |> List.mapi (fun at candidate -> if at = index then { candidate with Blocks = [ Doc.Paragraph para ] } else candidate)
                        (if isHeader then { doc with HeaderParts = filled } else { doc with FooterParts = filled }), index
                    | _ -> doc, index
                | None ->
                    // No header here: one is made by the first edit (the document's plain header paragraphs become its blocks).
                    let seed =
                        if target.TargetKind <> "default" || not parts.IsEmpty then []
                        else (if isHeader then doc.Headers else doc.Footers) |> List.map Doc.Paragraph
                    withNewPart doc target.TargetSection isHeader target.TargetKind seed
            match List.tryItem partIndex (if isHeader then edited.HeaderParts else edited.FooterParts) with
            | None -> false
            | Some part ->
                let pageCount = Ren.pageCount ctx doc
                let session : Inp.StorySession =
                    { IsHeader = isHeader
                      PageIndex = min pageIndex (max 0 (pageCount - 1))
                      Kind = target.TargetKind
                      Section = target.TargetSection
                      PartIndex = partIndex
                      Label = storyLabel isHeader target.TargetKind
                      Main = doc
                      Edited = edited
                      PartBlocks = part.Blocks
                      EntryBlocks = Paint.blocksWithRunningFields target.TargetFields part.Blocks
                      Folded = ResizeArray() }
                Inp.enterStory editor.Handler session None
                true

    do enterStoryOfPage.Value <- fun editor pageIndex isHeader -> enterHeaderFooter editor pageIndex (if isHeader then "header" else "footer")

    /// Where an offset of a text is in the text it became (a field's result
    /// changed in it): before the change as it was, after it moved by the
    /// difference of the lengths.
    let private offsetInChangedText (before: string) (after: string) (offset: int) : int =
        if before = after then offset
        else
            let limit = min before.Length after.Length
            let mutable prefix = 0
            while prefix < limit && before.[prefix] = after.[prefix] do prefix <- prefix + 1
            let mutable suffix = 0
            while suffix < limit - prefix && before.[before.Length - 1 - suffix] = after.[after.Length - 1 - suffix] do suffix <- suffix + 1
            if offset <= prefix then offset
            elif offset >= before.Length - suffix then offset + after.Length - before.Length
            else prefix

    let mutable private refreshingStoryFields = false

    /// The fields of the header or footer being edited show what its page
    /// shows as the layout stands (the page count grows while a long document
    /// is laid out, and when the header itself grows): the story's text is
    /// brought up to it, the caret staying where it is in the text. No undo
    /// step; a story nothing was changed in stays unchanged.
    let private refreshStoryFields (editor: Editor) =
        match editor.Handler.Story with
        | Some session when not refreshingStoryFields ->
            let ctx = editor.Handler.RenderContext
            match Ren.runningFieldsOfPage ctx editor.Doc session.PageIndex with
            | None -> ()
            | Some values ->
                let doc = editor.Doc
                let blocks = Paint.blocksWithRunningFields values doc.Blocks
                if not (obj.ReferenceEquals(blocks, doc.Blocks)) then
                    let next = { doc with Blocks = blocks }
                    let moved (position: Doc.Position) : Doc.Position =
                        match Doc.findParagraph doc position.BlockId, Doc.findParagraph next position.BlockId with
                        | Some before, Some after ->
                            { position with Offset = offsetInChangedText (TextBuf.toString before.Buffer) (TextBuf.toString after.Buffer) position.Offset }
                        | _ -> position
                    let selection = doc.Selection |> Option.map (fun selection -> ({ Anchor = moved selection.Anchor; Focus = moved selection.Focus } : Doc.Selection))
                    refreshingStoryFields <- true
                    try
                        if obj.ReferenceEquals(doc.Blocks, session.EntryBlocks) then
                            // (Nothing was changed in the story: it is the story as entered, with the values of now.)
                            Inp.enterStory editor.Handler { session with EntryBlocks = blocks; Folded = ResizeArray() } selection
                        else
                            Inp.updateDocument editor.Handler { next with Selection = selection }
                    finally
                        refreshingStoryFields <- false
        | _ -> ()

    do storyFieldsRefresh.Value <- refreshStoryFields

    /// What getCaretContext reports of the header or footer being edited:
    /// { Kind "header" | "footer", Variant "default" | "first" | "even",
    ///   PageIndex, Section (0-based), LinkedToPrevious (the section shows
    ///   the header of the section before it: it states none of its own),
    ///   CanLinkToPrevious (false in the first section) }; null in the body.
    let private headerFooterInfo (editor: Editor) : obj =
        match editor.Handler.Story with
        | None -> null
        | Some session ->
            let whole = wholeDocument editor
            let own = sectionReferences whole session.Section session.IsHeader |> List.exists (fun (kind, _) -> kind = session.Kind)
            createObj [
                "Kind" ==> (if session.IsHeader then "header" else "footer")
                "Variant" ==> session.Kind
                "PageIndex" ==> session.PageIndex
                "Section" ==> session.Section
                "LinkedToPrevious" ==> (session.Section > 0 && not own)
                "CanLinkToPrevious" ==> (session.Section > 0)
            ]

    /// Whether the header or footer being edited is the one of the section
    /// before it ("Link to Previous"). False in the body and in the first section.
    let getHeaderFooterLinkToPrevious (editor: Editor) : bool =
        match editor.Handler.Story with
        | Some session when session.Section > 0 ->
            not (sectionReferences (wholeDocument editor) session.Section session.IsHeader |> List.exists (fun (kind, _) -> kind = session.Kind))
        | _ -> false

    /// "Link to Previous" for the header or footer being edited. Off: the
    /// section gets a header of its own, a copy of the one it showed (every
    /// block of the copy with an id of its own). On: it shows the header of
    /// the section before it again (its own is no longer referred to). One
    /// undo step; the caret stays in the page's header, and is there again
    /// after undo and redo of the step.
    /// False when nothing changed (the body, the first section, already so).
    let setHeaderFooterLinkToPrevious (editor: Editor) (linked: bool) : bool =
        match editor.Handler.Story with
        | Some session when session.Section > 0 && getHeaderFooterLinkToPrevious editor <> linked ->
            let pageIndex, kind, isHeader = session.PageIndex, session.Kind, session.IsHeader
            Inp.exitStory editor.Handler |> ignore
            let doc = editor.Doc
            let next =
                if linked then
                    withSectionReferences doc session.Section isHeader (sectionReferences doc session.Section isHeader |> List.filter (fun (referenceKind, _) -> referenceKind <> kind))
                else
                    let shown =
                        Ren.headerFooterTarget editor.Handler.RenderContext doc pageIndex isHeader
                        |> Option.bind (fun target -> target.TargetPart)
                        |> Option.map (fun part -> part.Blocks)
                        |> Option.defaultValue []
                    withNewPart doc session.Section isHeader kind (copiedBlocks isHeader shown) |> fst
            // (A step of the whole document that belongs to this page's header: undo and redo come back here.)
            Inp.applyEditAtStoryPlace editor.Handler (if linked then "Link to previous" else "Unlink from previous") pageIndex isHeader next
            if not (enterHeaderFooter editor pageIndex (if isHeader then "header" else "footer")) then Inp.repaint editor.Handler
            true
        | _ -> false

    /// The page the caret is on (0-based): the page of the header or footer
    /// being edited, or the page of the caret's line.
    let private caretPageIndex (editor: Editor) : int =
        match editor.Handler.Story with
        | Some session -> session.PageIndex
        | None ->
            let ctx = editor.Handler.RenderContext
            match Ren.caretRect ctx editor.Doc with
            | Some (x, top, height) ->
                match Ren.pageRegionAt ctx editor.Doc x (top + height / 2.0) with
                | Some (pageIndex, _) -> pageIndex
                | None -> 0
            | None -> 0

    /// What the running fields show on the caret's page.
    let private caretFieldValues (editor: Editor) : Paint.RunningFieldValues =
        let ctx = editor.Handler.RenderContext
        // (The page count is that of the whole document.)
        let pageCount = Ren.pageCount ctx editor.Doc
        let pageIndex = caretPageIndex editor
        Ren.runningFieldsOfPage ctx editor.Doc pageIndex |> Option.defaultValue (Paint.runningFieldValues (pageIndex + 1) pageCount)

    /// Insert a page number field (PAGE) at the caret: every page shows its
    /// own number where the field stands in a header or footer - as its
    /// section counts and writes it (a section's first page number and number
    /// format, w:pgNumType). The field has the formatting of the text at the caret.
    let insertPageNumberField (editor: Editor) =
        let values = caretFieldValues editor
        commitDocumentUpdate editor (Doc.insertField "PAGE" (Paint.formatPageNumber values.FieldPageFormat values.FieldPage) editor.Doc)

    /// Insert a page count field (NUMPAGES) at the caret: the number of pages
    /// of the document (what getPageCount says).
    let insertPageCountField (editor: Editor) =
        commitDocumentUpdate editor (Doc.insertField "NUMPAGES" (string (caretFieldValues editor).FieldPageCount) editor.Doc)

    /// Insert a field with the number of pages of the caret's section (SECTIONPAGES).
    let insertSectionPageCountField (editor: Editor) =
        commitDocumentUpdate editor (Doc.insertField "SECTIONPAGES" (string (caretFieldValues editor).FieldSectionPages) editor.Doc)

    /// A double click (layout coordinates): on a page's header or footer area
    /// it is edited in place, with the caret where the click was; on the body
    /// while one is edited, back to the body. True when the click was taken.
    let private onDoubleClick (editor: Editor) (x: float) (y: float) : bool =
        let ctx = editor.Handler.RenderContext
        match Ren.pageRegionAt ctx editor.Doc x y, editor.Handler.Story with
        | Some (_, "body"), Some _ -> Inp.exitStory editor.Handler
        | Some (pageIndex, region), story when region = "header" || region = "footer" ->
            let here =
                match story with
                | Some session -> session.PageIndex = pageIndex && session.IsHeader = (region = "header")
                | None -> false
            // (In the story that is being edited a double click selects a word, as anywhere.)
            if here then false
            elif enterHeaderFooter editor pageIndex region then
                match Ren.hitTest ctx editor.Doc x y with
                | Some (blockId, offset) -> Inp.updateDocument editor.Handler (Doc.setCursor blockId offset editor.Doc)
                | None -> ()
                true
            else false
        | _ -> false

    do doubleClick.Value <- onDoubleClick

    /// What the caret is in, read off the document alone (no layout pass: a
    /// few microseconds whatever the document's length, where
    /// getSelectionInformation lays the caret's surroundings out):
    /// { InTable, Row, Column (0-based; -1 outside a table), RowCount,
    ///   ColumnCount, CanMergeCells (the selection spans several cells),
    ///   CanSplitCell (the caret's cell is a merged one), InHyperlink, InList,
    ///   InHeaderFooter (a header or footer is being edited in place, see enterHeaderFooter),
    ///   Story "body" | "header" | "footer", HeaderFooter (null in the body, else
    ///   { Kind, Variant, PageIndex, Section, LinkedToPrevious, CanLinkToPrevious }),
    ///   (before: InHeaderFooter was always false: headers and footers were not edited in
    ///   place), ImageSelected (the selection is one picture) }.
    let getCaretContext (editor: Editor) : obj =
        let context = Doc.caretContext editor.Doc
        createObj [
            "InTable" ==> context.InTable
            "Row" ==> context.Row
            "Column" ==> context.Column
            "RowCount" ==> context.RowCount
            "ColumnCount" ==> context.ColumnCount
            "CanMergeCells" ==> context.CanMergeCells
            "CanSplitCell" ==> context.CanSplitCell
            "InHyperlink" ==> context.InHyperlink
            "InList" ==> context.InList
            "InHeaderFooter" ==> editor.Handler.Story.IsSome
            "Story" ==> (match editor.Handler.Story with
                         | Some session -> if session.IsHeader then "header" else "footer"
                         | None -> "body")
            "HeaderFooter" ==> headerFooterInfo editor
            "ImageSelected" ==> context.ImageSelected
        ]

    /// Tab (forward) / Shift+Tab in a table, for a host's own keys or
    /// buttons: true when the caret moved to another cell.
    let moveToAdjacentTableCell (editor: Editor) (forward: bool) : bool =
        match Doc.moveToAdjacentTableCell forward editor.Doc with
        | Some moved ->
            commitDocumentUpdate editor moved
            true
        | None -> false

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

    /// Whether linked (not embedded) pictures may be loaded from the address the
    /// document gives. Off by default: a placeholder frame is shown instead.
    let setAllowLinkedImages (editor: Editor) (allow: bool) =
        if TextBuf.allowLinkedImages <> allow then
            TextBuf.allowLinkedImages <- allow
            Ren.clearLayoutCache editor.Handler.RenderContext
            Ren.render editor.Handler.RenderContext editor.Doc

    let getAllowLinkedImages (_: Editor) = TextBuf.allowLinkedImages

    /// Show an editor package: Office.Formats' `forEditorPackage(bytes)`, or -
    /// for a document that is shown before it is read to its end -
    /// `word.openProgressive(bytes)`. With the latter (and a viewport, which
    /// shows a part of the document) the import returns with what is read so
    /// far on display; see isLoadComplete. Without a viewport the canvas
    /// shows the whole document: it is read to its end first.
    let importFromDocxPackage (editor: Editor) (package: obj) =
        match Docx.progressiveImportOf package with
        | Some state when state.Failed.IsSome ->
            // (It could not be read to its end: what was read, see isLoadFailed.)
            loadDocumentOf editor state.Document (Some state)
        | Some state when not state.Complete ->
            if editor.Handler.RenderContext.Viewport.IsNone then
                Docx.completeProgressiveImport state
                loadDocumentOf editor state.Document (if state.Failed.IsSome then Some state else None)
            else
                loadDocumentOf editor state.Document (Some state)
                state.Package?Progressive?start () |> ignore
        | _ -> loadDocument editor (Docx.importFromEditorPackage package)

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
            HeaderReferences = editor.Doc.PageSetup.HeaderReferences
            FooterReferences = editor.Doc.PageSetup.FooterReferences
            PageNumberStart = editor.Doc.PageSetup.PageNumberStart
            PageNumberFormat = editor.Doc.PageSetup.PageNumberFormat
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
            HeaderReferences = editor.Doc.PageSetup.HeaderReferences
            FooterReferences = editor.Doc.PageSetup.FooterReferences
            PageNumberStart = editor.Doc.PageSetup.PageNumberStart
            PageNumberFormat = editor.Doc.PageSetup.PageNumberFormat
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
        pageFramesToJs editor.Handler.RenderContext (Ren.pageFrames editor.Handler.RenderContext editor.Doc)

    let getPageCount (editor: Editor) =
        ensureLoaded editor
        Ren.pageCount editor.Handler.RenderContext editor.Doc

    /// Document position under a point in layout coordinates (canvas pixels in
    /// whole-canvas mode; with a viewport, convert with viewportToDocument or
    /// use hitTestViewport), from the layout that is on screen.
    let hitTestViewport (editor: Editor) (x: float) (y: float) =
        let layoutX, layoutY = Ren.viewportToLayout editor.Handler.RenderContext x y
        match Ren.hitTest editor.Handler.RenderContext editor.Doc layoutX layoutY with
        | Some (blockId, offset) -> createObj [ "BlockId" ==> blockId; "Offset" ==> offset ]
        | None -> null

    /// Document position under a point in layout coordinates.
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
          Floating = false
          FloatingXml = None
          StyleId = None
          Look = None }

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
        Inp.exitStory editor.Handler |> ignore
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
        Inp.exitStory editor.Handler |> ignore
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
        Inp.exitStory editor.Handler |> ignore
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
        Inp.exitStory editor.Handler |> ignore
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

    /// A document that could not be read to its end is not saved: the file
    /// written would be the start of the one that was opened.
    let private refuseFailedLoad (editor: Editor) =
        match editor.LoadFailure with
        | Some error -> failwith ("The document could not be read to its end and cannot be saved (" + error + ")")
        | None -> ()

    /// Export document to EditorParagraphs (Office.Formats format)
    /// Use with: Office.Formats WordWriter.writeFromEditor()
    /// (Both exports throw for a document that could not be read to its end, see isLoadFailed.)
    let exportToDocx (editor: Editor) : obj array =
        ensureLoaded editor
        refuseFailedLoad editor
        Docx.exportToEditor (wholeDocument editor)

    let exportDocumentPackage (editor: Editor) =
        ensureLoaded editor
        refuseFailedLoad editor
        // One export for save and for callers: Docx.exportDocumentPackage knows the
        // document's styles and so leaves inherited formatting out of header/footer runs too.
        Docx.exportDocumentPackage (wholeDocument editor)

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
            "initializeCanvas" ==> initializeCanvas
            "initializeCanvasWithChange" ==> initializeCanvasWithChange
            "createViewer" ==> createViewer
            "getText" ==> getText
            "undo" ==> undo
            "redo" ==> redo
            "canUndo" ==> canUndo
            "canRedo" ==> canRedo
            "onChange" ==> onChange
            "isLoadComplete" ==> isLoadComplete
            "whenLoadComplete" ==> whenLoadComplete
            "completeLoad" ==> completeLoad
            "isLoadFailed" ==> isLoadFailed
            "cancelLoad" ==> cancelLoad
            "getLoadError" ==> getLoadError
            "onFontsDiscovered" ==> onFontsDiscovered
            "isLayoutComplete" ==> isLayoutComplete
            "whenLayoutComplete" ==> whenLayoutComplete
            "completeLayout" ==> completeLayout
            "getSelectionFormat" ==> getSelectionFormat
            "dispose" ==> dispose
            "getInputElement" ==> getInputElement
            "getTrackChanges" ==> getTrackChanges
            "setTrackChanges" ==> setTrackChanges
            "setAllowLinkedImages" ==> setAllowLinkedImages
            "getAllowLinkedImages" ==> getAllowLinkedImages
            "setTrackChangesOptions" ==> setTrackChangesOptions
            "setTextMeasurer" ==> setTextMeasurer
            "setTextMeasurers" ==> setTextMeasurers
            "setTextRangeMeasurer" ==> setTextRangeMeasurer
            "setTextPathPainter" ==> setTextPathPainter
            "clearTextPathPainter" ==> clearTextPathPainter
            "clearTextMeasurer" ==> clearTextMeasurer
            "getRenderStats" ==> getRenderStats
            "getRequiredFonts" ==> getRequiredFonts
            "getCaretContext" ==> getCaretContext
            "enterHeaderFooter" ==> enterHeaderFooter
            "exitHeaderFooter" ==> exitHeaderFooter
            "getHeaderFooterLinkToPrevious" ==> getHeaderFooterLinkToPrevious
            "setHeaderFooterLinkToPrevious" ==> setHeaderFooterLinkToPrevious
            "insertPageNumberField" ==> insertPageNumberField
            "insertPageCountField" ==> insertPageCountField
            "insertSectionPageCountField" ==> insertSectionPageCountField
            "moveToAdjacentTableCell" ==> moveToAdjacentTableCell
            "toggleTrackChanges" ==> toggleTrackChanges
            "getCommentAnchors" ==> getCommentAnchors
            "getRevisionInfos" ==> getRevisionInfos
            "selectCommentAnchor" ==> selectCommentAnchor
            "getSelection" ==> getSelection
            "setSelection" ==> setSelection
            "clearSelection" ==> clearSelection
            "getImageHitTest" ==> getImageHitTest
            "getLayoutSegments" ==> getLayoutSegments
            "getLayoutLines" ==> getLayoutLines
            "getLayoutFingerprint" ==> getLayoutFingerprint
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
            "setParagraphs" ==> setParagraphs
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
            "setFontSizePt" ==> setFontSizePt
            "setFontFamily" ==> setFontFamilyFor
            "setFontFamilies" ==> setFontFamilies
            "getFontFamilies" ==> getFontFamilies
            "getStyles" ==> getStyles
            "setLineSpacing" ==> setLineSpacing
            "setSpacingBefore" ==> setSpacingBefore
            "setSpacingAfter" ==> setSpacingAfter
            "applyParagraphStyle" ==> applyParagraphStyle
            "insertPageBreak" ==> insertPageBreak
            "insertHyperlink" ==> insertHyperlink
            "getHyperlink" ==> getHyperlink
            "setHyperlinkTarget" ==> setHyperlinkTarget
            "insertTableRowBefore" ==> insertTableRowBefore
            "insertTableColumnBefore" ==> insertTableColumnBefore
            "canMergeTableCells" ==> canMergeTableCells
            "mergeTableCells" ==> mergeTableCells
            "splitTableCell" ==> splitTableCell
            "getRevision" ==> getRevision
            "markSaved" ==> markSaved
            "isModified" ==> isModified
            "setTextPathPainterFactory" ==> setTextPathPainterFactory
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
            "hitTestViewport" ==> hitTestViewport
            "setViewport" ==> setViewport
            "getViewport" ==> getViewport
            "getDocumentSize" ==> getDocumentSize
            "viewportToDocument" ==> viewportToDocument
            "documentToViewport" ==> documentToViewport
            "getCaretRect" ==> getCaretRect
            "repaint" ==> repaint
            "setImageCache" ==> setImageCache
            "renderPage" ==> renderPage
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
