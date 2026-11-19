namespace Rbt.Word

#if FABLE_COMPILER

module Host =
    open Browser.Types
    open Fable.Core
    open Fable.Core.JsInterop
    open Rbt.FableUI.Reactive
    open Browser.Dom
    open Rbt.FableUI.Dom

    module WordApi = Rbt.Word.API
    module WordDialogs = Rbt.Word.Dialogs

    [<Emit("$0 === null || $0 === undefined")>]
    let private isNullOrUndefined (_value: obj) : bool = jsNative
    [<Emit("$0 === 'mixed'")>]
    let private isMixedValue (_value: obj) : bool = jsNative
    [<Emit("!!($0 && ($0.SelectionChanged || $0.FormatChanged || $0.DocumentChanged || $0.LoadChanged))")>]
    let private isSelectionNotice (_change: obj) : bool = jsNative

    // (Shaped text is the host's: a page that puts no window.RbtWordTextStack measures with the canvas, and
    // says nothing of an engine that would never load.)
    [<Emit("!!(window.RbtWordTextStack && typeof window.RbtWordTextStack.install === 'function') && !new URLSearchParams(window.location.search).has('rbCanvasMeasure') && window.localStorage.getItem('rb.word.canvasMeasure') !== '1'")>]
    let private shapedTextEnabled () : bool = jsNative

    [<Emit("Promise.resolve().then(function () { return window.RbtWordTextStack.install($0) }).then(function () { $1('Shaped text engine ready') }).catch(function (e) { $1('Shaped text engine failed: ' + ((e && e.message) || e)) })")>]
    let private installShapedText (_editor: WordApi.Editor) (_setStatus: string -> unit) : unit = jsNative

    type Host = {
        mutable Editor: WordApi.Editor option
        mutable Canvas: HTMLCanvasElement option
        mutable Saving: bool
        mutable Printing: bool
        mutable PdfFonts: obj array
        mutable OpenGeneration: int
        FileName: Signal<string>
        Dirty: Signal<bool>
        mutable ParagraphStyleInitial: string array
        mutable TableCellInitial: string array
        mutable TableMarginsInitial: string array
        mutable FontInitial: string array
        mutable BordersInitial: string array
        mutable ListsInitial: string array
        mutable PageSetupInitial: string array
        mutable PageSetupGeometry: Rbt.Word.Document.SectionPageSetup option
        mutable DialogLastRevision: int
        mutable DialogSelection: Rbt.Word.Document.Selection option
        mutable DialogEditor: WordApi.Editor option
        mutable DialogGeneration: int
        mutable DialogRevision: int
        mutable LinkInitial: string
        ParagraphStyle: WordDialogs.ParagraphStyleState
        FontDialog: WordDialogs.FormatDialogState
        BordersDialog: WordDialogs.FormatDialogState
        ListsDialog: WordDialogs.FormatDialogState
        FindReplace: WordDialogs.FindReplaceState
        TableCellStyle: WordDialogs.TableCellStyleState
        TableMargins: WordDialogs.TableMarginsState
        InsertTable: WordDialogs.InsertTableState
        ImageSize: WordDialogs.ImageSizeState
        TextBoxStyle: WordDialogs.TextBoxStyleState
        Caption: WordDialogs.CaptionState
        CrossReference: WordDialogs.CrossReferenceState
        Bookmark: WordDialogs.BookmarkState
        VbaEditor: WordDialogs.VbaEditorState
        ContentControl: WordDialogs.ContentControlState
        Field: WordDialogs.FieldState
        Equation: WordDialogs.EquationState
        Revision: WordDialogs.RevisionState
        PageSetup: WordDialogs.PageSetupState
        HeaderFooter: WordDialogs.HeaderFooterState
        Annotation: WordDialogs.AnnotationState
        Status: Signal<string>
        TrackChanges: Signal<bool>
        SelectionVersion: Signal<int>
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
        Saving = false
        Printing = false
        PdfFonts = [||]
        OpenGeneration = 0
        FileName = signal "Document1.docx"
        Dirty = signal false
        ParagraphStyleInitial = [||]
        TableCellInitial = [||]
        TableMarginsInitial = [||]
        FontInitial = [||]
        BordersInitial = [||]
        ListsInitial = [||]
        PageSetupInitial = [||]
        PageSetupGeometry = None
        DialogLastRevision = 0
        DialogSelection = None
        DialogEditor = None
        DialogGeneration = 0
        DialogRevision = 0
        LinkInitial = ""
        ParagraphStyle = WordDialogs.paragraphStyleState ()
        FontDialog = WordDialogs.formatDialogState 12
        BordersDialog = WordDialogs.formatDialogState 19
        ListsDialog = WordDialogs.formatDialogState 6
        FindReplace = WordDialogs.findReplaceState ()
        TableCellStyle = WordDialogs.tableCellStyleState ()
        TableMargins = WordDialogs.tableMarginsState ()
        InsertTable = WordDialogs.insertTableState ()
        ImageSize = WordDialogs.imageSizeState ()
        TextBoxStyle = WordDialogs.textBoxStyleState ()
        Caption = WordDialogs.captionState ()
        CrossReference = WordDialogs.crossReferenceState ()
        Bookmark = WordDialogs.bookmarkState ()
        VbaEditor = WordDialogs.vbaEditorState ()
        ContentControl = WordDialogs.contentControlState ()
        Field = WordDialogs.fieldState ()
        Equation = WordDialogs.equationState ()
        Revision = WordDialogs.revisionState ()
        PageSetup = WordDialogs.pageSetupState ()
        HeaderFooter = WordDialogs.headerFooterState ()
        Annotation = WordDialogs.annotationState ()
        Status = signal "Ready"
        TrackChanges = signal false
        SelectionVersion = signal 0
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

    let refreshDirty host =
        host.Dirty.Value <- host.Editor |> Option.exists WordApi.isModified

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

    [<Emit("window.open('about:blank', '_blank')")>]
    let private reservePdfWindow () : obj = jsNative

    [<Emit("$0.close()")>]
    let private closePdfWindow (_: obj) : unit = jsNative

    [<Emit("$0.closed")>]
    let private pdfWindowClosed (_: obj) : bool = jsNative

    [<Emit("((popup, blob) => { const url = URL.createObjectURL(blob); try { popup.location.replace(url); } finally { setTimeout(() => URL.revokeObjectURL(url), 60000); } })($0, $1)")>]
    let private showPdf (_: obj) (_: obj) : unit = jsNative
    // Load the DOCX writer only when a document is saved, and the DOCX reader when one is opened: the
    // modules of the packaged formats engine (office-formats' Main) that do it, not Main, which loads the
    // XLSX reader and writer too. (Main's word.writeFromEditorPackage and processors.forEditorPackage
    // are these functions.)
    [<Emit("import('../office-formats/src/WordPackageWriterAssembly.js').then(m => m.writeFromEditorPackage($0))")>]
    let private writeDocumentPackage (_: obj) : JS.Promise<obj> = jsNative

    [<Emit("$0.arrayBuffer().then(bytes => import('../office-formats/src/WordPackageOpen.js').then(m => m.forEditorPackage(bytes)))")>]
    let private readDocumentPackage (_: File) : JS.Promise<obj> = jsNative

    [<Emit("$0 != null && $0.Preserved != null")>]
    let private isDocumentPackage (_: obj) : bool = jsNative

    [<Emit("window.confirm('Discard unsaved changes and open another document?')")>]
    let private confirmReplacement () : bool = jsNative

    [<Emit("window.addEventListener('beforeunload', $0)")>]
    let private addBeforeUnload (_: Event -> unit) : unit = jsNative

    [<Emit("window.removeEventListener('beforeunload', $0)")>]
    let private removeBeforeUnload (_: Event -> unit) : unit = jsNative

    [<Emit("$0.preventDefault(); $0.returnValue = ''")>]
    let private warnBeforeUnload (_: Event) : unit = jsNative

    [<Emit("$0 === $1")>]
    let private sameObject (_: obj) (_: obj) : bool = jsNative

    [<Emit("$0.isConnected")>]
    let private isConnected (_: HTMLCanvasElement) : bool = jsNative

    // Revocation is deferred so the browser can consume the URL after click returns.
    [<Emit("""((blob) => {
        const url = URL.createObjectURL(blob);
        let anchor;
        try {
            anchor = document.createElement('a');
            anchor.href = url;
            anchor.download = $1;
            anchor.style.display = 'none';
            document.body.appendChild(anchor);
            anchor.click();
        } finally {
            if (anchor) anchor.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
        }
    })($0)""")>]
    let private requestDownload (_: obj) (_: string) : unit = jsNative

    let private isCurrentEditor host editor revisionIndex =
        host.Editor |> Option.exists (fun current ->
            sameObject (box current) (box editor)
            && sameObject current.Handler.Session.Epoch revisionIndex
            && not (isNullOrUndefined (WordApi.getInputElement current))
            && (host.Canvas |> Option.exists isConnected))

    // Invalidate older opens even if this picker is cancelled.
    let beginOpen host =
        host.OpenGeneration <- host.OpenGeneration + 1

    /// The pictures of a package (its `data:image/...` sources), decoded; then
    /// `onDecoded` with a Map of source to image element. The renderer draws
    /// pictures synchronously from the cache it is given, and without one an
    /// image is drawn only if it happens to be decoded already: in this host,
    /// which gave none, pictures never showed.
    [<Emit("""(function (pkg, onDecoded) {
        const found = new Set();
        const walk = (v) => {
            if (typeof v === 'string') { if (v.length > 11 && v.startsWith('data:image/')) found.add(v); }
            else if (Array.isArray(v)) { for (let i = 0; i < v.length; i++) walk(v[i]); }
            else if (v !== null && typeof v === 'object') {
                if (ArrayBuffer.isView(v)) return;
                for (const k in v) if (k !== 'Progressive') walk(v[k]);
            }
        };
        walk(pkg);
        if (found.size === 0) return;
        const cache = new Map();
        Promise.all(Array.from(found, (src) => { const img = new Image(); img.src = src; return img.decode().then(() => { cache.set(src, img); }, () => {}); }))
            .then(() => onDecoded(cache));
    })($0, $1)""")>]
    let private decodePackageImages (package: obj) (onDecoded: obj -> unit) : unit = jsNative

    /// Drive the editor's viewport from its scrolling workspace: the canvas is
    /// pinned to the workspace's visible area, a spacer gives the workspace the
    /// document's size, and every scroll, resize and change tells the engine
    /// what is in view (setViewport). The engine's own scroll moves (a change
    /// with ScrollTop) move the workspace; the caret is kept in view after an
    /// edit or a keyboard move. Returns { notify(change), dispose() }: the host
    /// passes it the editor's change notices and disposes it with the editor.
    [<Emit("""(function (canvas, docSize, setViewport, caretRect) {
        // (The canvas is mounted before its workspace and the spacer are in the
        // document: they are found on the next frame.)
        let scroll = null, spacer = null;
        const margin = 18;
        let applied = '', queued = false, follow = false, pointerDown = false, disposed = false;
        const apply = () => {
            queued = false;
            if (disposed || !scroll) return;
            const w = scroll.clientWidth, h = scroll.clientHeight;
            if (w <= 0 || h <= 0) return;
            const size = docSize();
            const cw = Math.max(w, Math.ceil(size.Width + 2 * margin));
            let ch = Math.ceil(size.Height + 2 * margin);
            if (size.LayoutComplete === false) ch = Math.max(ch, Math.ceil(scroll.scrollTop + h));
            spacer.style.width = cw + 'px';
            spacer.style.height = ch + 'px';
            const vp = { scale: 1, devicePixelRatio: window.devicePixelRatio || 1, scrollTop: scroll.scrollTop, scrollLeft: scroll.scrollLeft,
                         width: w, height: h, originX: (cw - size.Width) / 2, originY: margin };
            const key = JSON.stringify(vp);
            if (key !== applied) { applied = key; setViewport(vp); }
            if (follow && !pointerDown) {
                follow = false;
                const caret = caretRect();
                if (caret && typeof caret.ViewportY === 'number') {
                    const top = caret.ViewportY, bottom = caret.ViewportY + (caret.ViewportHeight || 0);
                    if (top < 0) scroll.scrollTop += top - 24;
                    else if (bottom > h) scroll.scrollTop += bottom - h + 24;
                }
            }
            follow = false;
        };
        const queue = () => { if (!queued) { queued = true; requestAnimationFrame(apply); } };
        const onScroll = () => apply();
        const down = () => { pointerDown = true; };
        const up = () => { pointerDown = false; };
        const observer = new ResizeObserver(queue);
        const attach = () => {
            if (disposed) return;
            scroll = canvas.closest('.rb-word-workspace');
            spacer = scroll && scroll.querySelector('.rb-word-spacer');
            if (!scroll || !spacer) { requestAnimationFrame(attach); return; }
            scroll.addEventListener('scroll', onScroll, { passive: true });
            scroll.addEventListener('pointerdown', down);
            window.addEventListener('pointerup', up);
            observer.observe(scroll);
            apply();
        };
        const notify = (change) => {
            if (!scroll) return;
            if (change && typeof change.ScrollTop === 'number') {
                const size = docSize();
                const need = Math.ceil(Math.max(size.Height + 2 * margin, change.ScrollTop + scroll.clientHeight));
                if (need > spacer.offsetHeight) spacer.style.height = need + 'px';
                scroll.scrollTop = change.ScrollTop;
            }
            if (change && (change.DocumentChanged || change.SelectionChanged)) follow = true;
            queue();
        };
        requestAnimationFrame(attach);
        const dispose = function () {
            disposed = true;
            observer.disconnect();
            if (!scroll) return;
            scroll.removeEventListener('scroll', onScroll);
            scroll.removeEventListener('pointerdown', down);
            window.removeEventListener('pointerup', up);
        };
        return { notify, dispose };
    })($0, $1, $2, $3)""")>]
    let private attachViewport (canvas: HTMLCanvasElement) (docSize: unit -> obj) (setViewport: obj -> unit) (caretRect: unit -> obj) : obj = jsNative

    let private replacePackage host editor package fileName =
        WordApi.importFromDocxPackage editor package
        decodePackageImages package (fun cache ->
            match host.Editor with
            | Some current when obj.ReferenceEquals(current, editor) -> WordApi.setImageCache editor cache
            | _ -> ())
        host.FileName.Value <- fileName
        host.SelectionVersion.Value <- host.SelectionVersion.Value + 1
        host.StyleCatalogVersion.Value <- host.StyleCatalogVersion.Value + 1
        refreshReviewState host
        refreshDirty host
        host.Status.Value <- sprintf "Opened %s" fileName
    
    let openParagraphStyle host =
        match host.Editor with
        | Some editor ->
            let format = WordApi.getSelectionFormat editor
            let paragraph = format?Paragraph
            let setNumber (signal: Signal<string>) (value: obj) =
                if isNullOrUndefined value || isMixedValue value then signal.Value <- ""
                else signal.Value <- string (unbox<float> value)
            let setString (signal: Signal<string>) (value: obj) =
                if isNullOrUndefined value || isMixedValue value then signal.Value <- ""
                else signal.Value <- unbox<string> value
            setString host.ParagraphStyle.Alignment paragraph?Alignment
            setNumber host.ParagraphStyle.IndentLeft paragraph?IndentLeft
            setNumber host.ParagraphStyle.IndentRight paragraph?IndentRight
            setNumber host.ParagraphStyle.SpacingBefore paragraph?SpaceBefore
            setNumber host.ParagraphStyle.SpacingAfter paragraph?SpaceAfter
            setNumber host.ParagraphStyle.LineSpacing paragraph?LineSpacing
            setString host.ParagraphStyle.BackgroundColor paragraph?BackgroundColor
            setNumber host.ParagraphStyle.FontSize format?FontSizePt
            let firstLineIndent = paragraph?FirstLineIndent
            if isNullOrUndefined firstLineIndent || isMixedValue firstLineIndent then
                host.ParagraphStyle.SpecialIndent.Value <- ""
                host.ParagraphStyle.SpecialIndentBy.Value <- ""
            else
                let indent = unbox<float> firstLineIndent
                host.ParagraphStyle.SpecialIndent.Value <- if indent > 0.0 then "firstLine" elif indent < 0.0 then "hanging" else "none"
                host.ParagraphStyle.SpecialIndentBy.Value <- string (abs indent)
        | None -> ()
        host.ParagraphStyleInitial <- [|
            host.ParagraphStyle.Alignment.Value
            host.ParagraphStyle.IndentLeft.Value
            host.ParagraphStyle.IndentRight.Value
            host.ParagraphStyle.SpacingBefore.Value
            host.ParagraphStyle.SpacingAfter.Value
            host.ParagraphStyle.LineSpacing.Value
            host.ParagraphStyle.BackgroundColor.Value
            host.ParagraphStyle.FontSize.Value
            host.ParagraphStyle.SpecialIndent.Value
            host.ParagraphStyle.SpecialIndentBy.Value
        |]
        host.ParagraphStyle.IsOpen.Value <- true

    let importDocumentPackage host package fileName =
        beginOpen host
        match host.Editor with
        | Some editor when not (WordApi.isModified editor) || confirmReplacement () ->
            try replacePackage host editor package fileName
            with ex -> host.Status.Value <- sprintf "DOCX open failed: %s" ex.Message
        | _ -> ()

    let openFile host (file: File) =
        match host.Editor with
        | None -> host.Status.Value <- "Editor not ready"
        | Some editor when WordApi.isModified editor && not (confirmReplacement ()) -> ()
        | Some editor ->
            let generation = host.OpenGeneration
            let revisionIndex = editor.Handler.Session.Epoch
            let lastRevision = editor.Handler.Session.LastRevision
            let isCurrent () = generation = host.OpenGeneration && isCurrentEditor host editor revisionIndex
            host.Status.Value <- sprintf "Opening %s..." file.name
            async {
                try
                    let! package = readDocumentPackage file |> Async.AwaitPromise
                    if not (isDocumentPackage package) then failwith "The file does not contain a Word document"
                    if isCurrent () then
                        // LastRevision is monotonic even when an edit is undone.
                        if editor.Handler.Session.LastRevision <> lastRevision then
                            host.Status.Value <- "Open cancelled because the document changed"
                        else
                            replacePackage host editor package file.name
                with ex ->
                    if isCurrent () then host.Status.Value <- sprintf "DOCX open failed: %s" ex.Message
            }
            |> Async.StartImmediate

    let saveFile host =
        match host.Editor with
        | None -> host.Status.Value <- "Editor not ready"
        | Some _ when host.Saving -> ()
        | Some editor ->
            host.Saving <- true
            let revisionIndex = editor.Handler.Session.Epoch
            let fileName = host.FileName.Value
            let isCurrent () = isCurrentEditor host editor revisionIndex
            try
                let revision = WordApi.getRevision editor
                let package = WordApi.exportDocumentPackage editor
                host.Status.Value <- "Preparing DOCX download..."
                async {
                    try
                        let! blob = writeDocumentPackage package |> Async.AwaitPromise
                        if isCurrent () then
                            requestDownload blob fileName
                            WordApi.markSaved editor (box revision)
                            refreshDirty host
                            host.Status.Value <- "DOCX download requested"
                    with ex ->
                        if isCurrent () then host.Status.Value <- sprintf "DOCX save failed: %s" ex.Message
                    host.Saving <- false
                }
                |> Async.StartImmediate
            with ex ->
                host.Saving <- false
                if isCurrent () then host.Status.Value <- sprintf "DOCX save failed: %s" ex.Message


    [<Emit("import('../word-pdf/Browser.js').then(api => api.exportPdf($0, $1, $2))")>]
    let private exportPdfSnapshot (_: Rbt.Word.Document.Document) (_: obj array) (_: bool) : JS.Promise<obj> = jsNative

    let setPdfFonts host (fonts: obj array) =
        host.PdfFonts <- if isNull fonts then [||] else fonts

    let printPdf host =
        match host.Editor with
        | None -> host.Status.Value <- "Editor not ready"
        | Some _ when host.Printing -> ()
        | Some editor ->
            let revisionIndex = editor.Handler.Session.Epoch
            let isCurrent () = isCurrentEditor host editor revisionIndex
            host.Printing <- true
            try
                WordApi.completeLoad editor
                if WordApi.isLoadFailed editor then failwith (WordApi.getLoadError editor)
                let snapshot = Rbt.Word.Input.wholeDocument editor.Handler editor.Doc
                let fonts = host.PdfFonts
                let allowLinked = WordApi.getAllowLinkedImages editor
                // Reserve during the user's gesture; opening after import/render
                // would be blocked by normal browser popup protection.
                let popup = reservePdfWindow ()
                if isNullOrUndefined popup then failwith "The browser blocked the PDF window"
                popup?opener <- null
                host.Status.Value <- "Rendering PDF..."
                async {
                    try
                        let! blob = exportPdfSnapshot snapshot fonts allowLinked |> Async.AwaitPromise
                        if isCurrent () && not (pdfWindowClosed popup) then
                            showPdf popup blob
                            host.Status.Value <- "PDF opened"
                        else closePdfWindow popup
                    with ex ->
                        closePdfWindow popup
                        if isCurrent () then host.Status.Value <- sprintf "PDF export failed: %s" ex.Message
                    host.Printing <- false
                } |> Async.StartImmediate
            with ex ->
                host.Printing <- false
                if isCurrent () then host.Status.Value <- sprintf "PDF export failed: %s" ex.Message


    let openFindReplace host =
        host.FindReplace.IsOpen.Value <- true

    let private currentTable (editor: WordApi.Editor) =
        let doc = editor.Doc
        let rec find (table: Rbt.Word.Document.TableBlock) blockId =
            table.Rows |> List.tryPick (fun row ->
                row.Cells |> List.tryPick (fun cell ->
                    if cell.Content |> List.exists (fun paragraph -> paragraph.Id = blockId) then Some (table, cell)
                    else cell.NestedTables |> List.tryPick (fun nested -> find nested blockId)))
        doc.Selection |> Option.bind (fun selection ->
            match Rbt.Word.Document.locateBlock doc selection.Focus.BlockId with
            | ValueSome (struct (_, Rbt.Word.Document.Table table)) -> find table selection.Focus.BlockId
            | _ -> None)
    let private captureDialogSelection host =
        host.DialogEditor <- host.Editor
        host.DialogSelection <- host.Editor |> Option.bind (fun editor -> editor.Doc.Selection)
        host.DialogGeneration <- host.OpenGeneration
        host.DialogRevision <- host.Editor |> Option.map WordApi.getRevision |> Option.defaultValue -1
        host.DialogLastRevision <- host.Editor |> Option.map (fun editor -> editor.Handler.Session.LastRevision) |> Option.defaultValue -1
    [<Emit("!!($0 && !$0.readOnly && !$0.disabled)")>]
    let private inputEditable (_: obj) : bool = jsNative


    let private restoreDialogSelection host =
        match host.Editor, host.DialogEditor, host.DialogSelection with
        | Some editor, Some captured, Some selection when obj.ReferenceEquals(editor, captured) && host.DialogGeneration = host.OpenGeneration && host.DialogRevision = WordApi.getRevision editor && inputEditable (WordApi.getInputElement editor) ->
            WordApi.setSelection editor selection.Anchor.BlockId selection.Anchor.Offset selection.Focus.BlockId selection.Focus.Offset
        | _ -> false
    let private dialogSelectionCurrent host =
        match host.Editor, host.DialogEditor, host.DialogSelection with
        | Some editor, Some captured, Some selection ->
            obj.ReferenceEquals(editor, captured)
            && host.DialogGeneration = host.OpenGeneration
            && host.DialogRevision = WordApi.getRevision editor
            && host.DialogLastRevision = editor.Handler.Session.LastRevision
            && editor.Doc.Selection = Some selection
            && inputEditable (WordApi.getInputElement editor)
        | _ -> false

    let private formatValues (state: WordDialogs.FormatDialogState) =
        state.Fields |> Array.map (fun field -> field.Value)
    [<Emit("(function(p,ids){const used=new Set(); const add=x=>{if(x!=null && /^\\s*\\+?\\d+\\s*$/.test(String(x))){const n=Number(x); if(Number.isInteger(n) && n>0 && n<=2147483647)used.add(String(n))}}; ids.forEach(add); const entry=p && p.SourceXml && p.SourceXml.find(x=>x.Part==='numbering'); if(entry){const xml=new DOMParser().parseFromString(entry.Xml,'application/xml'); if(xml.getElementsByTagName('parsererror').length) throw new Error('Cannot safely allocate a list ID from invalid numbering XML'); for(const e of xml.getElementsByTagName('*')){if(e.localName==='num')add(e.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','numId')); if(e.localName==='numId')add(e.getAttributeNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main','val'))}} let n=1; return function(){while(used.has(String(n)) && n<=2147483647)n++; if(n>2147483647)throw new Error('No numbering IDs available'); const id=String(n++); used.add(id); return id}})($0,$1)")>]
    let private listIdAllocator (_preserved: obj) (_ids: string array) : (unit -> string) = jsNative

    let openListsDialog host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            let targets = Rbt.Word.Document.paragraphFormattingTargets editor.Doc
            if not targets.IsEmpty then
                let uniform get =
                    match targets |> List.map (fun p -> get p.Metadata) |> List.distinct with
                    | [ value ] -> value
                    | _ -> ""
                let kind (m: Rbt.Word.Document.BlockMetadata) =
                    match m.ListStyle with Rbt.Word.Document.NoList -> "none" | Rbt.Word.Document.Bullet -> "bullet" | Rbt.Word.Document.Numbered -> "numbered"
                let values =
                    [| uniform kind
                       uniform (fun m -> if m.ListStyle = Rbt.Word.Document.Bullet then m.ListLevelText |> Option.defaultValue "•" else "")
                       uniform (fun m -> if m.ListStyle = Rbt.Word.Document.Numbered then m.ListNumberFormat |> Option.defaultValue "decimal" else "")
                       uniform (fun m -> if m.ListStyle = Rbt.Word.Document.NoList then "" else string m.ListLevel)
                       "keep"; "1" |]
                for i in 0 .. values.Length - 1 do host.ListsDialog.Fields.[i].Value <- values.[i]
                host.ListsInitial <- values
                captureDialogSelection host
                host.ListsDialog.Error.Value <- ""
                host.ListsDialog.IsOpen.Value <- true

    let applyListsDialog host =
        let state = host.ListsDialog
        let values = formatValues state
        let changed i = values.[i] <> host.ListsInitial.[i]
        let start = match System.Int32.TryParse values.[5] with true, n when n > 0 -> Some n | _ -> None
        let level = match System.Int32.TryParse values.[3] with true, n when n >= 1 && n <= 9 -> Some n | _ -> None
        if values = host.ListsInitial then state.IsOpen.Value <- false
        elif changed 3 && level.IsNone then state.Error.Value <- "Choose a list level from 1 to 9."
        elif values.[4] = "restart" && start.IsNone then state.Error.Value <- "Enter a positive whole starting number."
        elif changed 1 && changed 2 then state.Error.Value <- "Choose either a bullet marker or a number format."
        elif changed 0 && ((changed 1 && values.[0] <> "bullet") || (changed 2 && values.[0] <> "numbered")) then state.Error.Value <- "Choose a marker matching the list kind."
        elif not (dialogSelectionCurrent host) then state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            try
                let kind =
                    if changed 1 then Some Rbt.Word.Document.Bullet
                    elif changed 2 then Some Rbt.Word.Document.Numbered
                    elif changed 0 then
                        match values.[0] with
                        | "none" -> Some Rbt.Word.Document.NoList
                        | "bullet" -> Some Rbt.Word.Document.Bullet
                        | "numbered" -> Some Rbt.Word.Document.Numbered
                        | _ -> None
                    else None
                let marker = if changed 1 then Some values.[1] elif changed 2 then Some values.[2] else None
                let change: Rbt.Word.Document.ListDialogChange =
                    { Kind = kind; Marker = marker; Level = if changed 3 then level else None
                      Sequence = values.[4]; Start = start |> Option.defaultValue 1 }
                let mutable result = Ok ()
                runCommand host (fun editor ->
                    let doc = editor.Doc
                    let paras =
                        Rbt.Word.Document.allParagraphs doc
                        @ (doc.Package.Notes |> List.collect (fun note -> Rbt.Word.Document.paragraphsOfBlocks note.Blocks))
                        @ (doc.Package.Comments |> List.collect (fun comment -> Rbt.Word.Document.paragraphsOfBlocks comment.Blocks))
                    let ids = paras |> List.choose (fun p -> p.Metadata.ListId) |> List.toArray
                    let allocate = listIdAllocator (WordEditorPackage.preservedToEditorObject doc.Package.Preserved) ids
                    result <- WordApi.applyListDialog editor allocate change)
                match result with
                | Ok () -> state.IsOpen.Value <- false
                | Error message -> state.Error.Value <- message
            with ex -> state.Error.Value <- ex.Message

    let openFontDialog host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            let format = WordApi.getSelectionFormat editor
            let names = [| "FontFamily"; "FontSizePt"; "Bold"; "Italic"; "UnderlineMode"; "Strike"; "AllCaps"; "SmallCaps"; "VertAlign"; "Color"; "Highlight"; "Hidden" |]
            let defaults = [| "default"; ""; ""; ""; "none"; ""; ""; ""; "baseline"; "auto"; "none"; "" |]
            for i in 0 .. names.Length - 1 do
                let value: obj = format?(names.[i])
                host.FontDialog.Fields.[i].Value <-
                    if isMixedValue value then ""
                    elif isNullOrUndefined value then defaults.[i]
                    elif i = 2 || i = 3 || i = 5 || i = 6 || i = 7 || i = 11 then if unbox<bool> value then "true" else "false"
                    else string value
            host.FontInitial <- formatValues host.FontDialog
            captureDialogSelection host
            host.FontDialog.Error.Value <- ""
            host.FontDialog.IsOpen.Value <- true

    let openParagraphBorders host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            let targets = Rbt.Word.Document.paragraphFormattingTargets editor.Doc
            let side index (metadata: Rbt.Word.Document.BlockMetadata) =
                metadata.ParagraphBorders |> Option.bind (fun borders ->
                    match index with 0 -> borders.Top | 1 -> borders.Right | 2 -> borders.Bottom | _ -> borders.Left)
            let valueFor i (metadata: Rbt.Word.Document.BlockMetadata) =
                if i = 16 then metadata.BackgroundColor |> Option.defaultValue "none"
                elif i = 17 then metadata.ParagraphShadingPattern |> Option.defaultValue "clear"
                elif i = 18 then metadata.ParagraphShadingForeground |> Option.defaultValue "auto"
                else
                    let border = side (i / 4) metadata
                    match i % 4 with
                    | 0 -> border |> Option.map (fun b -> b.Style |> Option.defaultValue "single") |> Option.defaultValue "none"
                    | 1 -> border |> Option.bind (fun b -> b.Color) |> Option.defaultValue "auto"
                    | 2 -> border |> Option.bind (fun b -> b.Width) |> Option.map string |> Option.defaultValue "unset"
                    | _ -> border |> Option.bind (fun b -> b.Space) |> Option.map string |> Option.defaultValue "unset"
            if not targets.IsEmpty then
                for i in 0 .. 18 do
                    let values = targets |> List.map (fun p -> valueFor i p.Metadata) |> List.distinct
                    host.BordersDialog.Fields.[i].Value <- match values with [ value ] -> value | _ -> ""
                host.BordersInitial <- formatValues host.BordersDialog
                captureDialogSelection host
                host.BordersDialog.Error.Value <- ""
                host.BordersDialog.IsOpen.Value <- true

    [<Emit("/^#[0-9a-fA-F]{6}$/.test($0)")>]
    let private validHexColor (_: string) : bool = jsNative

    let private validColor special value =
        value = special || validHexColor value

    let applyFontDialog host =
        let state = host.FontDialog
        let values = formatValues state
        let changed i = values.[i] <> host.FontInitial.[i]
        let validSize =
            if not (changed 1) then true
            else match System.Double.TryParse values.[1] with true, n -> n > 0.0 && n <= 1638.0 && not (System.Double.IsNaN n || System.Double.IsInfinity n) | _ -> false
        if values = host.FontInitial then state.IsOpen.Value <- false
        elif not validSize || (changed 0 && System.String.IsNullOrWhiteSpace values.[0])
             || (changed 9 && not (validColor "auto" values.[9]))
             || (changed 10 && not (validColor "none" values.[10])) then
            state.Error.Value <- "Enter a valid font, size, and #RRGGBB colors (or auto/none)."
        elif not (dialogSelectionCurrent host) then state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            let change (style: Rbt.Word.TextBuffer.TextStyle) =
                let mutable result = style
                if changed 0 then
                    let family = if values.[0] = "default" then None else Some (values.[0].Trim())
                    let flag = Rbt.Word.TextBuffer.ExplicitFormat.FontFamily
                    result <- { result with FontFamily = family; Explicit = if family.IsSome then result.Explicit ||| flag else result.Explicit &&& (~~~flag) }
                if changed 1 then result <- Rbt.Word.TextBuffer.withExplicitFontSize (float values.[1] * 96.0 / 72.0) result
                if changed 2 then result <- Rbt.Word.TextBuffer.withExplicitBold (values.[2] = "true") result
                if changed 3 then result <- Rbt.Word.TextBuffer.withExplicitItalic (values.[3] = "true") result
                if changed 4 then
                    result <- Rbt.Word.TextBuffer.withExplicitUnderline (values.[4] <> "none") result
                    result <- { result with UnderlineStyle = if values.[4] = "none" then None else Some values.[4] }
                if changed 5 then
                    result <- Rbt.Word.TextBuffer.withExplicitStrike (values.[5] = "true") result
                    if values.[5] = "false" then result <- Rbt.Word.TextBuffer.withExplicitDoubleStrike false result
                if changed 6 then result <- { result with AllCaps = values.[6] = "true"; Explicit = result.Explicit ||| Rbt.Word.TextBuffer.ExplicitFormat.AllCaps }
                if changed 7 then result <- { result with SmallCaps = values.[7] = "true"; Explicit = result.Explicit ||| Rbt.Word.TextBuffer.ExplicitFormat.SmallCaps }
                if changed 8 then result <- { result with VertAlign = if values.[8] = "baseline" then None else Some values.[8] }
                if changed 9 then
                    result <- { result with
                                    Color = if values.[9] = "auto" then None else Some values.[9]
                                    Explicit = if values.[9] = "auto" then result.Explicit ||| Rbt.Word.TextBuffer.ExplicitFormat.ColorAuto else result.Explicit &&& (~~~Rbt.Word.TextBuffer.ExplicitFormat.ColorAuto) }
                if changed 10 then result <- { result with Highlight = if values.[10] = "none" then None else Some values.[10] }
                if changed 11 then result <- Rbt.Word.TextBuffer.withExplicitHidden (values.[11] = "true") result
                result
            try
                runCommand host (fun editor -> WordApi.applyFontDialogFormat editor change)
                state.IsOpen.Value <- false
            with ex -> state.Error.Value <- ex.Message

    let applyParagraphBorders host =
        let state = host.BordersDialog
        let values = formatValues state
        let changed i = values.[i] <> host.BordersInitial.[i]
        let numberValid i =
            if not (changed i) || values.[i] = "" || values.[i] = "unset" then true
            else match System.Double.TryParse values.[i] with true, n -> n >= 0.0 && n <= 1000.0 && not (System.Double.IsNaN n || System.Double.IsInfinity n) | _ -> false
        let colorsValid = [ 1; 5; 9; 13; 18 ] |> List.forall (fun i -> not (changed i) || validColor "auto" values.[i])
        let numbersValid = [ 2; 3; 6; 7; 10; 11; 14; 15 ] |> List.forall numberValid
        if values = host.BordersInitial then state.IsOpen.Value <- false
        elif not colorsValid || not numbersValid || (changed 16 && not (validColor "none" values.[16])) then
            state.Error.Value <- "Enter non-negative widths and spacing, and #RRGGBB colors (or auto/none)."
        elif not (dialogSelectionCurrent host) then state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            let number i (old: float option) =
                if not (changed i) then old
                elif values.[i] = "" || values.[i] = "unset" then None
                else Some (float values.[i])
            let text i special (old: string option) =
                if not (changed i) then old
                elif values.[i] = special then None
                else Some values.[i]
            let updateSide offset (old: Rbt.Word.Document.ParagraphBorder option) =
                if not ([ offset .. offset + 3 ] |> List.exists changed) then old
                elif changed offset && values.[offset] = "none" then None
                else
                    let border: Rbt.Word.Document.ParagraphBorder = old |> Option.defaultValue { Style = Some "single"; Color = None; Width = None; Space = None }
                    Some { Style = text offset "none" border.Style
                           Color = text (offset + 1) "auto" border.Color
                           Width = number (offset + 2) border.Width
                           Space = number (offset + 3) border.Space }
            let update (metadata: Rbt.Word.Document.BlockMetadata) =
                let old: Rbt.Word.Document.ParagraphBorders = metadata.ParagraphBorders |> Option.defaultValue { Top = None; Right = None; Bottom = None; Left = None }
                let borders = { old with Top = updateSide 0 old.Top; Right = updateSide 4 old.Right; Bottom = updateSide 8 old.Bottom; Left = updateSide 12 old.Left }
                let shadingChanged = changed 16 || changed 17 || changed 18
                let pattern =
                    if changed 17 then Some (if values.[17] = "none" then "nil" else values.[17])
                    elif changed 16 && values.[16] = "none" && metadata.ParagraphShadingPattern.IsNone then Some "nil"
                    else metadata.ParagraphShadingPattern
                let foreground = text 18 "auto" metadata.ParagraphShadingForeground
                { metadata with
                    ParagraphBorders = if borders.Top.IsNone && borders.Right.IsNone && borders.Bottom.IsNone && borders.Left.IsNone then None else Some borders
                    ParagraphBorderBefore = if [ 0 .. 3 ] |> List.exists changed then None else metadata.ParagraphBorderBefore
                    ParagraphBorderAfter = if [ 8 .. 11 ] |> List.exists changed then None else metadata.ParagraphBorderAfter
                    BackgroundColor = text 16 "none" metadata.BackgroundColor
                    ParagraphShadingPattern = pattern
                    ParagraphShadingForeground = foreground
                    ExplicitParagraphShadingPattern = if shadingChanged then pattern |> Option.orElse (Some "clear") else metadata.ExplicitParagraphShadingPattern
                    ExplicitParagraphShadingForeground = if shadingChanged then foreground else metadata.ExplicitParagraphShadingForeground
                    Explicit = if shadingChanged && (text 16 "none" metadata.BackgroundColor).IsSome then metadata.Explicit ||| Rbt.Word.Document.ParagraphExplicit.BackgroundColor else metadata.Explicit }
            try
                runCommand host (fun editor -> WordApi.applyParagraphDialogFormat editor update)
                state.IsOpen.Value <- false
            with ex -> state.Error.Value <- ex.Message



    let private numberText value = value |> Option.map string |> Option.defaultValue ""
    let private cellValues (state: WordDialogs.TableCellStyleState) =
        [| state.BackgroundColor.Value; state.ShadingPattern.Value; state.ShadingForegroundColor.Value
           state.Width.Value; state.BorderColor.Value; state.BorderWidth.Value; state.VerticalAlignment.Value
           state.PaddingTop.Value; state.PaddingRight.Value; state.PaddingBottom.Value; state.PaddingLeft.Value |]
    let private marginValues (state: WordDialogs.TableMarginsState) =
        Array.append [| state.PaddingTop.Value; state.PaddingRight.Value; state.PaddingBottom.Value; state.PaddingLeft.Value; state.CellSpacing.Value |] (state.Fields |> Array.map (fun field -> field.Value))

    let openTableCellStyle host =
        host.Editor |> Option.bind currentTable |> Option.iter (fun (_, cell) ->
            let state = host.TableCellStyle
            state.BackgroundColor.Value <- cell.Shading |> Option.bind (fun s -> s.BackgroundColor) |> Option.orElseWith (fun () -> cell.BackgroundColor) |> Option.defaultValue "#FFFFFF"
            state.ShadingPattern.Value <- cell.Shading |> Option.bind (fun s -> s.Pattern) |> Option.defaultValue (if cell.BackgroundColor.IsSome then "clear" else "none")
            state.ShadingForegroundColor.Value <- cell.Shading |> Option.bind (fun s -> s.ForegroundColor) |> Option.defaultValue "#000000"
            state.Width.Value <- numberText cell.Width
            state.BorderColor.Value <- cell.BorderColor |> Option.defaultValue ""
            state.BorderWidth.Value <- numberText cell.BorderWidth
            state.VerticalAlignment.Value <- cell.VerticalAlignment |> Option.defaultValue "top"
            state.PaddingTop.Value <- numberText cell.PaddingTop
            state.PaddingRight.Value <- numberText cell.PaddingRight
            state.PaddingBottom.Value <- numberText cell.PaddingBottom
            state.PaddingLeft.Value <- numberText cell.PaddingLeft
            host.TableCellInitial <- cellValues state
            captureDialogSelection host
            state.Error.Value <- ""
            state.IsOpen.Value <- true)

    let openTableMargins host =
        host.Editor |> Option.bind currentTable |> Option.iter (fun (table, cell) ->
            let row = table.Rows |> List.find (fun row -> row.Cells |> List.exists (fun candidate -> obj.ReferenceEquals(candidate, cell)))
            let state = host.TableMargins
            state.PaddingTop.Value <- numberText table.PaddingTop
            state.PaddingRight.Value <- numberText table.PaddingRight
            state.PaddingBottom.Value <- numberText table.PaddingBottom
            state.PaddingLeft.Value <- numberText table.PaddingLeft
            state.CellSpacing.Value <- numberText table.CellSpacing
            let widthKind =
                match table.PreferredWidth |> Option.map (fun width -> width.Kind) with
                | Some Rbt.Word.Document.AutoWidth -> "auto"
                | Some Rbt.Word.Document.DxaWidth -> "dxa"
                | Some Rbt.Word.Document.PercentWidth -> "pct"
                | Some Rbt.Word.Document.PointWidth -> "pt"
                | None -> ""
            let firstBorder = table.Borders |> Option.bind (fun borders -> borders.Top |> Option.orElse borders.Left |> Option.orElse borders.Right |> Option.orElse borders.Bottom)
            let values =
                [| widthKind; table.PreferredWidth |> Option.bind (fun width -> width.Value) |> numberText
                   table.Alignment |> Option.defaultValue ""; numberText table.Indent; ""
                   numberText row.Height
                   row.CanSplit |> Option.map (fun value -> if value then "true" else "false") |> Option.defaultValue ""
                   if row.IsHeader then "true" else "false"
                   table.Shading |> Option.bind (fun shading -> shading.BackgroundColor) |> Option.orElse table.BackgroundColor |> Option.defaultValue ""
                   "keep"; firstBorder |> Option.bind (fun border -> border.Color) |> Option.defaultValue "#000000"
                   firstBorder |> Option.bind (fun border -> border.Width) |> Option.defaultValue 0.75 |> string
                   table.Shading |> Option.bind (fun shading -> shading.Pattern) |> Option.defaultValue (if table.BackgroundColor.IsSome then "clear" else "none")
                   table.Shading |> Option.bind (fun shading -> shading.ForegroundColor) |> Option.defaultValue "#000000" |]
            for i in 0 .. values.Length - 1 do state.Fields.[i].Value <- values.[i]
            host.TableMarginsInitial <- marginValues state
            captureDialogSelection host
            state.Error.Value <- ""
            state.IsOpen.Value <- true)

    let openInsertTable host =
        host.InsertTable.Error.Value <- ""
        captureDialogSelection host
        host.InsertTable.IsOpen.Value <- true

    let openImageSize host =
        host.ImageSize.IsOpen.Value <- true

    let openTextBoxStyle host =
        host.TextBoxStyle.IsOpen.Value <- true

    let openCaption host =
        host.Caption.IsOpen.Value <- true

    let openCrossReference host =
        host.CrossReference.IsOpen.Value <- true

    // ---- the VBA editor ----------------------------------------------------------------------------------

    /// The project as the editor shows it: its modules (a blank Module1 when it has none), its macros,
    /// what macros printed and how runs ended.
    let private refreshVba host (editor: WordApi.Editor) (shown: string option) =
        let state = host.VbaEditor
        let text (value: obj) = if isNull value then "" else string value
        let modules =
            match WordApi.getMacroModules editor |> Array.map (fun m -> text m?name, text m?kind, text m?code) |> List.ofArray with
            | [] -> [ "Module1", "standard", "" ]
            | found -> found
        state.Modules <- modules
        state.Macros <- WordApi.getMacros editor |> List.ofArray
        let current = match shown with Some name -> name | None -> state.Module.Value
        let current = if modules |> List.exists (fun (n, _, _) -> n = current) then current else (let n, _, _ = modules.Head in n)
        state.Module.Value <- current
        state.Code.Value <- modules |> List.pick (fun (n, _, c) -> if n = current then Some c else None)
        if not (state.Macros |> List.contains state.Macro.Value) then state.Macro.Value <- (state.Macros |> List.tryHead |> Option.defaultValue "")
        let lines = WordApi.getMacroOutput editor
        state.Output.Value <- String.concat "\n" (if lines.Length > 40 then lines.[lines.Length - 40 ..] else lines)

    /// Show the editor again (its module and macro lists follow the project).
    let private reopenVba host =
        host.VbaEditor.IsOpen.Value <- false
        host.VbaEditor.IsOpen.Value <- true

    /// Open the VBA editor on the document's macros (they are loaded first: the VBA engine is fetched then).
    let openVbaEditor host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            (WordApi.loadMacros editor).``then``(fun () ->
                refreshVba host editor None
                host.VbaEditor.IsOpen.Value <- true) |> ignore

    /// Run an action of the editor with the document's macros enabled (as Word's "Enable Content").
    let private withMacrosEnabled host (action: WordApi.Editor -> Fable.Core.JS.Promise<obj>) =
        match host.Editor with
        | None -> ()
        | Some editor ->
            let run () = (action editor).``then``(fun _ -> refreshVba host editor None; reopenVba host) |> ignore
            if WordApi.getMacrosEnabled editor then run ()
            else (WordApi.enableMacros editor).``then``(fun _ -> run ()) |> ignore

    let private applyVbaModule host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            let name = host.VbaEditor.Module.Value
            (WordApi.setMacroModule editor name host.VbaEditor.Code.Value).``then``(fun _ ->
                refreshVba host editor (Some name)
                reopenVba host) |> ignore

    let private runVbaMacro host =
        let name = host.VbaEditor.Macro.Value
        if name <> "" then withMacrosEnabled host (fun editor -> WordApi.runMacro editor name)

    let private runVbaLine host =
        let line = host.VbaEditor.Immediate.Value
        if line.Trim() <> "" then withMacrosEnabled host (fun editor -> WordApi.runMacroLine editor line)

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

    let private pageSetupValues (state: WordDialogs.PageSetupState) =
        [| state.PageWidth.Value; state.PageHeight.Value; state.MarginLeft.Value; state.MarginRight.Value
           state.MarginTop.Value; state.MarginBottom.Value; state.HeaderDistance.Value; state.FooterDistance.Value |]

    let openPageSetup host =
        match host.Editor with
        | None -> ()
        | Some editor ->
            match WordApi.currentPageSetupForDialog editor with
            | None -> host.Status.Value <- "Select a section first"
            | Some setup ->
                let state = host.PageSetup
                let inch value = string (value / 96.0)
                state.PageWidth.Value <- inch setup.PageWidth
                state.PageHeight.Value <- inch setup.PageHeight
                state.MarginLeft.Value <- inch setup.MarginLeft
                state.MarginRight.Value <- inch setup.MarginRight
                state.MarginTop.Value <- inch setup.MarginTop
                state.MarginBottom.Value <- inch setup.MarginBottom
                state.HeaderDistance.Value <- inch setup.HeaderDistance
                state.FooterDistance.Value <- inch setup.FooterDistance
                state.DifferentFirstPage.Value <- if setup.DifferentFirstPage then "true" else "false"
                state.EvenAndOddHeaders.Value <- if setup.EvenAndOddHeaders then "true" else "false"
                state.Preset.Value <- "custom"
                state.Orientation.Value <- if setup.PageWidth > setup.PageHeight then "landscape" else "portrait"
                state.Error.Value <- ""
                host.PageSetupInitial <- pageSetupValues state
                host.PageSetupGeometry <- Some setup
                captureDialogSelection host
                state.IsOpen.Value <- true

    let openHeaderFooter host =
        host.HeaderFooter.IsOpen.Value <- true

    let openAnnotation kind host =
        let state = host.Annotation
        state.Kind.Value <- kind
        state.Error.Value <- ""
        if kind = "hyperlink" then
            let link = host.Editor |> Option.map WordApi.getHyperlink |> Option.defaultValue null
            state.HasLink.Value <- if isNullOrUndefined link then "false" else "true"
            state.Target.Value <- if state.HasLink.Value = "true" then unbox<string> link?Target else ""
            state.LinkText.Value <- ""
            host.LinkInitial <- state.Target.Value
            captureDialogSelection host
        state.IsOpen.Value <- true

    let insertTable host =
        runCommand host (fun editor -> WordApi.insertTable editor 2 3)
        host.Status.Value <- "Table inserted"

    let insertCustomTable host =
        let state = host.InsertTable
        match System.Int32.TryParse state.Rows.Value, System.Int32.TryParse state.Columns.Value with
        | (true, rows), (true, columns) when rows >= 1 && columns >= 1 && restoreDialogSelection host ->
            runCommand host (fun editor -> WordApi.insertTable editor rows columns)
            host.Status.Value <- sprintf "Table inserted: %d x %d" rows columns
            state.IsOpen.Value <- false
        | (true, rows), (true, columns) when rows >= 1 && columns >= 1 ->
            state.Error.Value <- "The document or selection changed; reopen this dialog."
        | _ -> state.Error.Value <- "Enter positive whole numbers for rows and columns."

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
        let values = pageSetupValues state
        match host.PageSetupGeometry with
        | None -> state.Error.Value <- "Reopen page setup."
        | Some original ->
            let parsed =
                values |> Array.map (fun value ->
                    match System.Double.TryParse value with
                    | true, number when not (System.Double.IsNaN number || System.Double.IsInfinity number) -> Some (number * 96.0)
                    | _ -> None)
            let unchanged =
                values = host.PageSetupInitial
                && ((state.DifferentFirstPage.Value = "true") = original.DifferentFirstPage)
                && ((state.EvenAndOddHeaders.Value = "true") = original.EvenAndOddHeaders)
            if not (dialogSelectionCurrent host) then
                state.Error.Value <- "The document or selection changed, or editing is unavailable; reopen this dialog."
            elif unchanged then state.IsOpen.Value <- false
            elif parsed |> Array.exists Option.isNone then
                state.Error.Value <- "Enter finite numbers for all dimensions."
            else
                let old = [| original.PageWidth; original.PageHeight; original.MarginLeft; original.MarginRight; original.MarginTop; original.MarginBottom; original.HeaderDistance; original.FooterDistance |]
                let dimensions = parsed |> Array.mapi (fun i number -> if values.[i] = host.PageSetupInitial.[i] then old.[i] else Option.get number)
                if dimensions |> Array.exists (fun n -> System.Double.IsNaN n || System.Double.IsInfinity n || n < 0.0)
                   || dimensions.[0] <= 0.0 || dimensions.[1] <= 0.0
                   || dimensions.[2] + dimensions.[3] >= dimensions.[0]
                   || dimensions.[4] + dimensions.[5] >= dimensions.[1] then
                    state.Error.Value <- "Page dimensions must be positive; margins and distances must be non-negative and leave a positive body."
                else
                    let setup =
                        { original with PageWidth = dimensions.[0]; PageHeight = dimensions.[1]
                                        MarginLeft = dimensions.[2]; MarginRight = dimensions.[3]
                                        MarginTop = dimensions.[4]; MarginBottom = dimensions.[5]
                                        HeaderDistance = dimensions.[6]; FooterDistance = dimensions.[7]
                                        DifferentFirstPage = state.DifferentFirstPage.Value = "true"
                                        EvenAndOddHeaders = state.EvenAndOddHeaders.Value = "true" }
                    runCommand host (fun editor -> WordApi.applyCurrentPageSetupDialog editor setup)
                    host.Status.Value <- "Page setup applied to current section"
                    state.IsOpen.Value <- false

    let applyHeaderFooter host =
        let state = host.HeaderFooter
        runCommand host (fun editor ->
            WordApi.setHeaderFooterText editor state.DefaultHeader.Value state.DefaultFooter.Value state.FirstHeader.Value state.FirstFooter.Value state.EvenHeader.Value state.EvenFooter.Value)
        runCommand host (fun editor ->
            WordApi.setHeaderFooterTableText editor state.DefaultHeaderTableLeft.Value state.DefaultHeaderTableRight.Value state.DefaultFooterTableLeft.Value state.DefaultFooterTableRight.Value)
        runCommand host (fun editor ->
            WordApi.setHeaderFooterImageText editor state.DefaultHeaderImageSrc.Value state.DefaultHeaderImageAlt.Value state.DefaultHeaderImageWidth.Value state.DefaultHeaderImageHeight.Value state.DefaultFooterImageSrc.Value state.DefaultFooterImageAlt.Value state.DefaultFooterImageWidth.Value state.DefaultFooterImageHeight.Value)
        if state.LinkSectionToPrevious.Value = "true" then
            runCommand host (fun editor -> WordApi.setCurrentSectionHeaderFooterText editor "" "")
            state.OverrideSection.Value <- "false"
        elif state.OverrideSection.Value = "true" then
            runCommand host (fun editor ->
                WordApi.setCurrentSectionHeaderFooterText editor state.SectionHeader.Value state.SectionFooter.Value)
        host.Status.Value <- "Headers and footers updated"
        state.IsOpen.Value <- false

    let findNext host =
        let state = host.FindReplace
        let query = state.Query.Value
        if System.String.IsNullOrWhiteSpace query then
            state.Result.Value <- "Enter text to find"
            host.Status.Value <- "Enter text to find"
        else
            match host.Editor with
            | None -> ()
            | Some editor ->
                let matches = WordApi.findText editor query (state.MatchCase.Value = "true")
                if matches.Length = 0 then
                    state.Result.Value <- "No matches"
                    host.Status.Value <- "No matches"
                else
                    let selection = WordApi.getSelection editor
                    let paragraphOrder = WordApi.getParagraphOrder editor
                    let orderByBlock = System.Collections.Generic.Dictionary<string, int>()
                    paragraphOrder |> Array.iteri (fun index blockId -> orderByBlock.[blockId] <- index)
                    let blockId, endOffset =
                        if isNullOrUndefined selection then "", -1
                        else
                            let focusBlock = unbox<string> selection?FocusBlockId
                            let anchorBlock = unbox<string> selection?AnchorBlockId
                            let focusOffset = unbox<int> selection?FocusOffset
                            let anchorOffset = unbox<int> selection?AnchorOffset
                            if focusBlock = anchorBlock then focusBlock, max focusOffset anchorOffset
                            else
                                let focusOrder = match orderByBlock.TryGetValue focusBlock with true, value -> value | _ -> -1
                                let anchorOrder = match orderByBlock.TryGetValue anchorBlock with true, value -> value | _ -> -1
                                if focusOrder >= anchorOrder then focusBlock, focusOffset else anchorBlock, anchorOffset
                    let currentOrder =
                        match Portable.Dict.tryFind blockId orderByBlock with
                        | ValueSome value -> value
                        | _ -> -1
                    let currentIndex =
                        if currentOrder < 0 then None
                        else
                            matches
                            |> Array.tryFindIndex (fun item ->
                                let matchOrder =
                                    match Portable.Dict.tryFind (unbox<string> item?BlockId) orderByBlock with
                                    | ValueSome value -> value
                                    | _ -> -1
                                matchOrder > currentOrder
                                || (matchOrder = currentOrder && unbox<int> item?Start >= endOffset))
                    let startIndex = currentIndex |> Option.defaultValue 0
                    WordApi.selectFindMatch editor matches.[startIndex]
                    state.Result.Value <- sprintf "Match %d of %d" (startIndex + 1) matches.Length
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
                let count = WordApi.replaceAllText editor query state.Replacement.Value (state.MatchCase.Value = "true")
                state.Result.Value <- sprintf "Replaced %d matches" count
                host.Status.Value <- state.Result.Value

    let applyParagraphStyle host =
        let state = host.ParagraphStyle
        let changed index value = host.ParagraphStyleInitial.Length <= index || host.ParagraphStyleInitial.[index] <> value
        let changes =
            createObj [
                if changed 0 state.Alignment.Value then "Alignment" ==> state.Alignment.Value
                if changed 1 state.IndentLeft.Value then "IndentLeft" ==> parseFloat 0.0 state.IndentLeft.Value
                if changed 2 state.IndentRight.Value then "IndentRight" ==> parseFloat 0.0 state.IndentRight.Value
                if changed 3 state.SpacingBefore.Value then "SpacingBefore" ==> parseFloat 0.0 state.SpacingBefore.Value
                if changed 4 state.SpacingAfter.Value then "SpacingAfter" ==> parseFloat 10.0 state.SpacingAfter.Value
                if changed 5 state.LineSpacing.Value then "LineSpacing" ==> parseFloat 1.5 state.LineSpacing.Value
                if changed 6 state.BackgroundColor.Value then "BackgroundColor" ==> state.BackgroundColor.Value
                if changed 7 state.FontSize.Value then "FontSizePt" ==> parseFloat 11.0 state.FontSize.Value
                if changed 8 state.SpecialIndent.Value || changed 9 state.SpecialIndentBy.Value then
                    "FirstLineIndent" ==> (specialIndentValue state |> Option.defaultValue 0.0)
            ]
        runCommand host (fun editor -> WordApi.applyParagraphStyleChanges editor changes)
        host.Status.Value <- sprintf "Paragraph style: %s, %spt before, %spt after, %.1fx line, %spt, %s indent %spt, bg %s" (alignmentLabel state.Alignment.Value) state.SpacingBefore.Value state.SpacingAfter.Value (parseFloat 1.5 state.LineSpacing.Value) state.FontSize.Value (specialIndentLabel state.SpecialIndent.Value) state.SpecialIndentBy.Value state.BackgroundColor.Value
        state.IsOpen.Value <- false

    let private validNumber allowZero (value: string) =
        if value = "" then true
        else
            match System.Double.TryParse value with
            | true, number -> not (System.Double.IsNaN number || System.Double.IsInfinity number) && (if allowZero then number >= 0.0 else number > 0.0)
            | _ -> false

    let private optionalNumber value = if value = "" then -1.0 else parseFloat -1.0 value

    let applyTableCellStyle host =
        let state = host.TableCellStyle
        let values = cellValues state
        let numeric = [ 3; 5; 7; 8; 9; 10 ]
        if values = host.TableCellInitial then state.IsOpen.Value <- false
        elif numeric |> List.exists (fun i -> values.[i] <> host.TableCellInitial.[i] && not (validNumber (i >= 7) values.[i])) then
            state.Error.Value <- "Enter positive widths and non-negative padding, or leave them blank."
        elif not (restoreDialogSelection host) then state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            let changed i = values.[i] <> host.TableCellInitial.[i]
            let number i old = if changed i then (if values.[i] = "" then None else Some (parseFloat 0.0 values.[i])) else old
            let text i old = if changed i then (if values.[i] = "" then None else Some values.[i]) else old
            runCommand host (fun editor ->
                WordApi.updateSelectedTableCell editor (fun cell ->
                    let cell =
                        if changed 1 then
                            match values.[1] with
                            | "none" -> { cell with BackgroundColor = None; Shading = None }
                            | "clear" -> { cell with BackgroundColor = Some values.[0]; Shading = None }
                            | pattern ->
                                { cell with BackgroundColor = None
                                            Shading = Some { Pattern = Some pattern; ForegroundColor = Some values.[2]; BackgroundColor = Some values.[0] } }
                        elif changed 0 || changed 2 then
                            match cell.Shading with
                            | Some shading ->
                                { cell with Shading = Some { shading with
                                                                BackgroundColor = text 0 shading.BackgroundColor
                                                                ForegroundColor = text 2 shading.ForegroundColor } }
                            | None -> { cell with BackgroundColor = text 0 cell.BackgroundColor }
                        else cell
                    let borders =
                        if changed 4 || changed 5 then
                            let patch (border: Rbt.Word.Document.TableBorder option) =
                                border |> Option.map (fun border ->
                                    { border with Color = text 4 border.Color; Width = number 5 border.Width })
                            cell.Borders |> Option.map (fun borders ->
                                { borders with Top = patch borders.Top; Right = patch borders.Right
                                               Bottom = patch borders.Bottom; Left = patch borders.Left
                                               InsideH = patch borders.InsideH; InsideV = patch borders.InsideV })
                        else cell.Borders
                    { cell with
                        Width = number 3 cell.Width
                        PreferredWidth =
                            if changed 3 then
                                number 3 cell.Width |> Option.map (fun width ->
                                    { Kind = Rbt.Word.Document.DxaWidth; Value = Some width })
                            else cell.PreferredWidth
                        BorderColor = text 4 cell.BorderColor
                        BorderWidth = number 5 cell.BorderWidth
                        Borders = borders
                        VerticalAlignment = text 6 cell.VerticalAlignment
                        PaddingTop = number 7 cell.PaddingTop
                        PaddingRight = number 8 cell.PaddingRight
                        PaddingBottom = number 9 cell.PaddingBottom
                        PaddingLeft = number 10 cell.PaddingLeft }))
            host.Status.Value <- "Table cell style applied"
            state.IsOpen.Value <- false

    let applyTableMargins host =
        let state = host.TableMargins
        let values = marginValues state
        let initial = host.TableMarginsInitial
        let changed i = values.[i] <> initial.[i]
        let validSigned value =
            value = "" || (match System.Double.TryParse value with | true, number -> not (System.Double.IsNaN number || System.Double.IsInfinity number) | _ -> false)
        let widthChanged = changed 5 || changed 6
        let widthValid =
            not widthChanged ||
            (match values.[5] with
             | "" | "auto" -> values.[6] = "" || values.[6] = "0"
             | "dxa" | "pct" | "pt" -> values.[6] <> "" && validNumber false values.[6]
             | _ -> false)
        let colorValid =
            not (changed 13) || values.[13] = "" ||
            System.Text.RegularExpressions.Regex.IsMatch(values.[13], "^#[0-9a-fA-F]{6}$")
        let borderValid =
            values.[14] <> "grid" ||
            (System.Text.RegularExpressions.Regex.IsMatch(values.[15], "^#[0-9a-fA-F]{6}$") && values.[16] <> "" && validNumber false values.[16])
        if values.[14] = "keep" && values.[0..13] = initial.[0..13] && values.[17..18] = initial.[17..18] then state.IsOpen.Value <- false
        elif not widthValid || (changed 8 && not (validSigned values.[8])) ||
             (changed 10 && not (validNumber true values.[10])) ||
             ([0..4] |> List.exists (fun i -> changed i && not (validNumber true values.[i]))) || not colorValid || not borderValid ||
             (changed 18 && not (System.Text.RegularExpressions.Regex.IsMatch(values.[18], "^#[0-9a-fA-F]{6}$"))) then
            state.Error.Value <- "Enter valid width, height, indent, margins, and border and background colors."
        elif not (dialogSelectionCurrent host) || not (restoreDialogSelection host) then state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            let number i old = if changed i then (if values.[i] = "" then None else Some (parseFloat 0.0 values.[i])) else old
            runCommand host (fun editor ->
                WordApi.updateSelectedTableProperties editor (fun table rowIndex ->
                    let width =
                        if not widthChanged then table.PreferredWidth
                        else
                            let kind =
                                match values.[5] with
                                | "auto" -> Some Rbt.Word.Document.AutoWidth
                                | "dxa" -> Some Rbt.Word.Document.DxaWidth
                                | "pct" -> Some Rbt.Word.Document.PercentWidth
                                | "pt" -> Some Rbt.Word.Document.PointWidth
                                | _ -> None
                            kind |> Option.map (fun kind -> ({ Kind = kind; Value = if values.[6] = "" then None else Some (parseFloat 0.0 values.[6]) } : Rbt.Word.Document.TableWidth))
                    let background = if values.[13] = "" then None else Some values.[13]
                    let shading =
                        if changed 17 then
                            match values.[17] with
                            | "none" | "clear" -> None
                            | pattern -> Some ({ Pattern = Some pattern; ForegroundColor = Some values.[18]; BackgroundColor = background } : Rbt.Word.Document.Shading)
                        elif changed 13 || changed 18 then
                            table.Shading |> Option.map (fun shading ->
                                { shading with
                                    BackgroundColor = if changed 13 then background else shading.BackgroundColor
                                    ForegroundColor = if changed 18 then Some values.[18] else shading.ForegroundColor })
                        else table.Shading
                    let borders =
                        if values.[14] = "keep" then table.Borders
                        else
                            let border : Rbt.Word.Document.TableBorder =
                                if values.[14] = "none" then { Style = Some "nil"; Color = None; Width = None }
                                else { Style = Some "single"; Color = Some values.[15]; Width = Some (parseFloat 0.0 values.[16]) }
                            Some { Top = Some border; Right = Some border; Bottom = Some border; Left = Some border; InsideH = Some border; InsideV = Some border }
                    { table with
                        PreferredWidth = width
                        Alignment = if changed 7 then (if values.[7] = "" then None else Some values.[7]) else table.Alignment
                        Indent = number 8 table.Indent
                        Borders = borders
                        BackgroundColor = if changed 17 && values.[17] = "none" then None elif changed 13 || changed 17 then background else table.BackgroundColor
                        Shading = shading
                        PaddingTop = number 0 table.PaddingTop
                        PaddingRight = number 1 table.PaddingRight
                        PaddingBottom = number 2 table.PaddingBottom
                        PaddingLeft = number 3 table.PaddingLeft
                        CellSpacing = number 4 table.CellSpacing
                        Rows = table.Rows |> List.mapi (fun i row ->
                            if i <> rowIndex then row else
                            { row with
                                Height = number 10 row.Height
                                HeightRule = if changed 10 then (if values.[10] = "" then None else Some "atLeast") else row.HeightRule
                                CanSplit = if changed 11 then (match values.[11] with "true" -> Some true | "false" -> Some false | _ -> None) else row.CanSplit
                                IsHeader = if changed 12 then values.[12] = "true" else row.IsHeader }) }))
            host.Status.Value <- "Table properties applied"
            state.IsOpen.Value <- false

    let applyAnnotation host =
        let state = host.Annotation
        let target = state.Target.Value.Trim()
        if state.Kind.Value = "hyperlink" && state.HasLink.Value = "true" && target = host.LinkInitial then state.IsOpen.Value <- false
        elif state.Kind.Value = "hyperlink" && target = "" && state.HasLink.Value <> "true" then
            state.Error.Value <- "Enter a link target."
        elif state.Kind.Value = "hyperlink" && target <> "" && not (Rbt.Word.TextBuffer.isActiveHyperlinkTarget target) then
            state.Error.Value <- "Use an http, https, mailto, tel, or #bookmark target."
        elif state.Kind.Value = "hyperlink" && not (restoreDialogSelection host) then
            state.Error.Value <- "The document or selection changed; reopen this dialog."
        else
            runCommand host (fun editor ->
                match state.Kind.Value with
                | "comment" -> WordApi.setCommentFor editor state.AnnotationId.Value state.Text.Value state.Author.Value state.Date.Value
                | "note" -> WordApi.setNoteFor editor state.AnnotationId.Value state.NoteType.Value state.Text.Value
                | _ when state.HasLink.Value = "true" -> WordApi.setHyperlinkTarget editor target
                | _ -> WordApi.insertHyperlink editor target state.LinkText.Value)
            host.Status.Value <-
                match state.Kind.Value with
                | "comment" -> sprintf "Comment applied: %s" state.AnnotationId.Value
                | "note" -> sprintf "%s applied: %s" state.NoteType.Value state.AnnotationId.Value
                | _ when target = "" -> "Hyperlink removed"
                | _ -> sprintf "Hyperlink applied: %s" target
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
                button [ cls "word-revision-accept"; attr "aria-label" (sprintf "Accept revision %d" (index + 1)); attr "data-testid" "word-revision-pane-accept"; onClick (fun _ -> acceptRevisionInfo host revision) ] [ text "Accept" ]
                button [ cls "word-revision-reject"; attr "aria-label" (sprintf "Reject revision %d" (index + 1)); attr "data-testid" "word-revision-pane-reject"; onClick (fun _ -> rejectRevisionInfo host revision) ] [ text "Reject" ]
            ]
        ]

    let private revisionPane host =
        div [ cls "word-revision-pane"; attr "role" "region"; attr "aria-label" "Revisions"; attr "data-testid" "word-revision-pane" ] [
            div [ cls "word-revision-pane-header" ] [
                span [] [ textR (fun () -> sprintf "Revisions (%d)" host.RevisionInfos.Value.Length) ]
                button [ attr "aria-label" "Previous revision"; attr "data-testid" "word-revision-pane-previous"; onClick (fun _ -> selectPreviousRevision host) ] [ text "Previous" ]
                button [ attr "aria-label" "Next revision"; attr "data-testid" "word-revision-pane-next"; onClick (fun _ -> selectNextRevision host) ] [ text "Next" ]
                button [ attr "aria-label" "Accept all revisions"; attr "data-testid" "word-revision-pane-accept-all"; onClick (fun _ -> acceptAllRevisions host) ] [ text "Accept all" ]
                button [ attr "aria-label" "Reject all revisions"; attr "data-testid" "word-revision-pane-reject-all"; onClick (fun _ -> rejectAllRevisions host) ] [ text "Reject all" ]
            ]
            children (fun () -> host.RevisionInfos.Value |> Array.toList |> List.mapi (revisionItem host))
        ]

    let view host =
        let dialogs: (Signal<bool> * (unit -> WordDialogs.Form)) list = [
            host.ParagraphStyle.IsOpen, (fun () -> WordDialogs.paragraphStyleDialog "word-paragraph-dialog" host.ParagraphStyle (fun () -> applyParagraphStyle host))
            host.FontDialog.IsOpen, (fun () -> WordDialogs.fontDialog host.FontDialog (fun () -> applyFontDialog host))
            host.BordersDialog.IsOpen, (fun () -> WordDialogs.paragraphBordersDialog host.BordersDialog (fun () -> applyParagraphBorders host))
            host.ListsDialog.IsOpen, (fun () -> WordDialogs.listsDialog host.ListsDialog (fun () -> applyListsDialog host))
            host.FindReplace.IsOpen, (fun () -> WordDialogs.findReplaceDialog "word-find-dialog" host.FindReplace (fun () -> findNext host) (fun () -> replaceAll host))
            host.TableCellStyle.IsOpen, (fun () -> WordDialogs.tableCellStyleDialog "word-table-cell-style-dialog" host.TableCellStyle (fun () -> applyTableCellStyle host))
            host.TableMargins.IsOpen, (fun () -> WordDialogs.tableMarginsDialog "word-table-margins-dialog" host.TableMargins (fun () -> applyTableMargins host))
            host.InsertTable.IsOpen, (fun () -> WordDialogs.insertTableDialog "word-insert-table-dialog" host.InsertTable (fun () -> insertCustomTable host))
            host.ImageSize.IsOpen, (fun () -> WordDialogs.imageSizeDialog "word-image-size-dialog" host.ImageSize (fun () -> applyImageSize host))
            host.TextBoxStyle.IsOpen, (fun () -> WordDialogs.textBoxStyleDialog "word-textbox-style-dialog" host.TextBoxStyle (fun () -> applyTextBoxStyle host))
            host.Caption.IsOpen, (fun () -> WordDialogs.captionDialog "word-caption-dialog" host.Caption (fun () -> insertCaption host))
            host.CrossReference.IsOpen, (fun () -> WordDialogs.crossReferenceDialog "word-cross-reference-dialog" host.CrossReference (fun () -> insertCrossReference host))
            host.Bookmark.IsOpen, (fun () -> WordDialogs.bookmarkDialog "word-bookmark-dialog" host.Bookmark (fun () -> applyBookmark host))
            host.ContentControl.IsOpen, (fun () -> WordDialogs.contentControlDialog "word-content-control-dialog" host.ContentControl (fun () -> applyContentControl host))
            host.Field.IsOpen, (fun () -> WordDialogs.fieldDialog "word-field-dialog" host.Field (fun () -> insertField host))
            host.Equation.IsOpen, (fun () -> WordDialogs.equationDialog "word-equation-dialog" host.Equation (fun () -> insertEquation host))
            host.Revision.IsOpen, (fun () -> WordDialogs.revisionDialog "word-revision-dialog" host.Revision (fun () -> applyRevision host))
            host.PageSetup.IsOpen, (fun () -> WordDialogs.pageSetupDialog "word-page-setup-dialog" host.PageSetup (fun () -> applyPageSetup host))
            host.HeaderFooter.IsOpen, (fun () -> WordDialogs.headerFooterDialog "word-header-footer-dialog" host.HeaderFooter (fun () -> applyHeaderFooter host))
            host.Annotation.IsOpen, (fun () -> WordDialogs.annotationDialog "word-annotation-dialog" host.Annotation (fun () -> applyAnnotation host))
            host.VbaEditor.IsOpen, (fun () -> WordDialogs.vbaEditorDialog "word-vba-dialog" host.VbaEditor (fun () -> applyVbaModule host) (fun () -> runVbaMacro host) (fun () -> runVbaLine host))
        ]
        let uiDialog = Rbt.FableUI.OfficeUiDialog.createState ()
        let mutable openedFor: Signal<bool> option = None
        let hostAttrs =
            [ cls "word-host"; attr "role" "region"; attr "aria-label" "Document workspace"; attr "tabindex" "0"; attr "data-testid" (host.TestPrefix + "-host")
              refEffect (fun _ ->
                  let subscription = effect (fun () ->
                      let openDialogs = dialogs |> List.map (fun (isOpen, makeForm) -> isOpen.Value, isOpen, makeForm) |> List.filter (fun (openNow, _, _) -> openNow)
                      let active =
                          openDialogs
                          |> List.tryFind (fun (_, isOpen, _) -> openedFor |> Option.exists (fun previous -> not (obj.ReferenceEquals(previous, isOpen))))
                          |> Option.orElseWith (fun () -> openDialogs |> List.tryHead)
                      untrack (fun () ->
                          match active with
                          | None ->
                              openedFor <- None
                              Rbt.FableUI.OfficeUiDialog.closeDialog uiDialog
                          | Some (_, isOpen, makeForm) when not (openedFor |> Option.exists (fun previous -> obj.ReferenceEquals(previous, isOpen))) ->
                              for other, _ in dialogs do
                                  if not (obj.ReferenceEquals(other, isOpen)) then other.Value <- false
                              openedFor <- Some isOpen
                              let spec, values = makeForm ()
                              Rbt.FableUI.OfficeUiDialog.openDialogWithValues uiDialog spec values
                          | _ -> ()))
                  fun () -> subscription.Dispose (); Rbt.FableUI.OfficeUiDialog.closeDialog uiDialog) ]
            @ (Ui.chromeCssVariables |> List.map (fun (name, value) -> style name value))

        div hostAttrs [
            div [ cls "word-workspace"; attr "data-testid" (host.TestPrefix + "-workspace") ] [
              div [ cls "word-pin" ] [
                canvas [ cls "word-canvas"
                         attr "data-testid" (host.TestPrefix + "-canvas")
                         attr "width" "900"
                         attr "height" "560"
                         refEffect (fun element ->
                             let canvas = element :?> HTMLCanvasElement
                             host.Canvas <- Some canvas
                             let editor = WordApi.initializeCanvasWithChange canvas (fun () -> host.SelectionVersion.Value <- host.SelectionVersion.Value + 1; refreshReviewState host; refreshDirty host)
                             host.Editor <- Some editor
                             let unsubscribeSelection =
                                 WordApi.onChange editor (fun change ->
                                     if isSelectionNotice change then
                                         host.SelectionVersion.Value <- host.SelectionVersion.Value + 1)
                             WordApi.setContent editor host.InitialText
                             host.SelectionVersion.Value <- host.SelectionVersion.Value + 1
                             refreshReviewState host
                             refreshDirty host
                             let beforeUnload event =
                                 if WordApi.isModified editor then warnBeforeUnload event
                             addBeforeUnload beforeUnload
                             let viewport =
                                 attachViewport canvas
                                     (fun () -> WordApi.getDocumentSize editor)
                                     (fun viewport -> WordApi.setViewport editor viewport |> ignore)
                                     (fun () -> WordApi.getCaretRect editor)
                             let unsubscribeViewport = WordApi.onChange editor (fun change -> viewport?notify(change) |> ignore)
                             host.Status.Value <- "Editor ready"
                             if shapedTextEnabled () then
                                 host.Status.Value <- "Loading shaped text engine"
                                 installShapedText editor (fun status -> host.Status.Value <- status)
                             fun () ->
                                 unsubscribeViewport ()
                                 viewport?dispose() |> ignore
                                 host.OpenGeneration <- host.OpenGeneration + 1
                                 removeBeforeUnload beforeUnload
                                 unsubscribeSelection ()
                                 WordApi.dispose editor
                                 host.Editor <- None
                                 host.Canvas <- None) ] []
                commentLayer host
              ]
              div [ cls "word-spacer" ] []
            ]

            showIf (fun () -> host.RevisionInfos.Value.Length > 0) (revisionPane host)

            Rbt.FableUI.OfficeUiDialog.renderDialog uiDialog
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
