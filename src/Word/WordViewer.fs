namespace Rbt.Word

#if FABLE_COMPILER

/// The read-only viewer: shows the pages of a document, edits nothing. This
/// module depends on the document import, the layout, the paint commands and
/// the renderer only - not on the editor (input, commands, clipboard, the
/// editor API) - so a page that only previews documents imports
/// `word/WordViewer.js` and its bundle carries none of that.
///   import createViewer from '@rbt/office/word/WordViewer.js'
///   const viewer = createViewer(measureCanvas); viewer.load(editorPackage); ...
/// `Word.createViewer` is this function.
module Viewer =
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types

    module Doc = Document
    module Ren = Renderer

    let pageFramesToJs (ctx: Ren.RenderContext) (pages: Layout.PageInfo list) =
        List.zip pages (Ren.pageOffsets ctx pages)
        |> List.map (fun (page, x) ->
            createObj [
                // Where the page's left edge is in the document's width, in
                // layout pixels: 0 in whole-canvas mode; through a viewport a
                // page narrower than the widest one is centred.
                "X" ==> x
                "Top" ==> page.Top
                "Width" ==> page.Width
                "Height" ==> page.Height
                "MarginTop" ==> page.MarginTop
                "MarginBottom" ==> page.MarginBottom
                "MarginLeft" ==> page.MarginLeft
                "MarginRight" ==> page.MarginRight
            ])
        |> List.toArray

    let documentSizeToJs (ctx: Ren.RenderContext) (doc: Doc.Document) =
        let width, height, pageCount = Ren.documentSize ctx doc
        let scale = Ren.viewportScale ctx
        createObj [
            "Width" ==> width
            "Height" ==> height
            "ScaledWidth" ==> width * scale
            "ScaledHeight" ==> height * scale
            "PageCount" ==> pageCount
            // False while a long document is still being laid out from its
            // start: the size and the page count are of what is laid out so far.
            "LayoutComplete" ==> Ren.layoutIsComplete ctx
        ]

    /// Create a read-only viewer. `canvas` is used for text measurement (and by
    /// `render`, which paints the whole document onto it); pages are painted
    /// onto canvases of the host's with `renderPage`.
    ///   load(editorPackage, options?)  show a package from Office.Formats (forEditorPackage).
    ///                                  options { progressive: true }: the start of the document is laid out
    ///                                  first and the rest in idle slices - getPageFrames / getDocumentSize
    ///                                  then give the pages there are so far (see onLayoutProgress), and
    ///                                  getPageCount, or renderPage of a later page, lays the rest out at once.
    ///                                  Without it everything is laid out by the first call that asks.
    ///   isLayoutComplete() / whenLayoutComplete() (a promise of the page frames) /
    ///   onLayoutProgress(callback)     callback({ PageCount, LayoutComplete }) as pages are added; returns
    ///                                  the function that ends the subscription
    ///   setImageCache(map)             Map of image source -> decoded HTMLImageElement
    ///   getPageCount() / getPageFrames() / getDocumentSize()
    ///   renderPage(canvas, pageIndex, scale)   paint one page (see API.renderPage)
    ///   setTextMeasurers(measureText, measureRange) / setTextMeasurer / setTextRangeMeasurer /
    ///   clearTextMeasurer()            measure text as an editor with the same measurers does:
    ///                                  the viewer then has that editor's lines and pages
    ///   setTextPathPainter(painterFor) painterFor(context2d) gives the text painter for a canvas
    ///   getRequiredFonts()             [{ family, bold, italic }]: the faces the loaded document is set in
    ///   getRenderStats()               { FrameBuilds, LaidOutBlocks, CollectedBlocks, CollectedOnPaint, Blocks,
    ///                                    PainterVersion, MeasurerVersion } - see API.getRenderStats
    ///   render(document)               paint a Document on the viewer's canvas (as before)
    let createViewer (canvas: HTMLCanvasElement) =
        let ctx = Ren.createContext canvas
        let state = createObj [ "canvas" ==> canvas; "document" ==> Doc.empty ]
        let current () : Doc.Document = unbox state?document
        Ren.render ctx (current ())
        state?render <- fun (nextDoc: Doc.Document) ->
            state?document <- nextDoc
            Ren.render ctx nextDoc
        let progressListeners = ResizeArray<obj -> unit>()
        // The package the document is of while it is still being read, and what ends the viewer's interest in it.
        let mutable loading : Docx.ProgressiveImport option = None
        let mutable stopLoading : unit -> unit = ignore
        let loadWaiters = ResizeArray<unit -> unit>()
        let tellProgress (complete: bool) =
            if progressListeners.Count > 0 then
                let progress = createObj [ "PageCount" ==> List.length (Ren.pageFrames ctx (current ())); "LayoutComplete" ==> complete; "LoadComplete" ==> loading.IsNone ]
                for listener in progressListeners.ToArray() do
                    if progressListeners.Contains listener then listener progress
        // The rest of a document that is laid out from its start, in idle slices
        // (one that is still being read is read first).
        let continueLayout () =
            if loading.IsNone && not (Ren.layoutIsComplete ctx) then
                Ren.scheduleLayout ctx current (fun complete tell -> if tell then tellProgress complete)
        let endLoading () =
            stopLoading ()
            stopLoading <- ignore
            loading <- None
            ctx.Progress.LoadPending <- false
            if loadWaiters.Count > 0 then
                let waiters = loadWaiters.ToArray()
                loadWaiters.Clear()
                for waiter in waiters do waiter ()
        /// The whole document, when it is still being read: read to its end now.
        let ensureLoaded () =
            match loading with
            | Some importing -> Docx.completeProgressiveImport importing
            | None -> ()
        let mutable loadTold = 0.0
        state?load <- fun (package: obj) (options: obj) ->
            endLoading ()
            let progressive = not (isNull options) && not (isNull (options?progressive)) && unbox<bool> (options?progressive)
            let importing =
                match Docx.progressiveImportOf package with
                | Some importing when not importing.Complete ->
                    // (Not asked to show the document before it is all there: read to its end first.)
                    if progressive then Some importing
                    else
                        Docx.completeProgressiveImport importing
                        None
                | _ -> None
            // A viewer shows no caret.
            let doc = { Docx.importFromEditorPackage package with Selection = None }
            state?document <- doc
            Ren.clearLayoutCache ctx
            Ren.cancelScheduledLayout ctx
            match importing with
            | Some importing ->
                // More of the document was read: its blocks after the blocks there are,
                // laid out in their turn.
                let listener = fun (step: Docx.ImportStep) ->
                    let before = current ()
                    state?document <- { importing.Document with Selection = None }
                    let wasWhole = Ren.documentExtended ctx (List.length before.Blocks)
                    if step.Complete then endLoading ()
                    let now = JS.Constructors.Date.now ()
                    if step.Complete || wasWhole || now - loadTold > 150.0 then
                        loadTold <- now
                        // (What was laid out was all there was: the pages of what is there now, as far as the first slices go.)
                        tellProgress (Ren.layoutIsComplete ctx)
                    if step.Complete then continueLayout ()
                importing.Listeners.Add listener
                loading <- Some importing
                stopLoading <- fun () -> importing.Listeners.Remove listener |> ignore
                ctx.Progress.LoadPending <- true
                importing.Package?Progressive?start () |> ignore
            | None -> ()
            if progressive then
                ctx.LayoutState.BlockLimit <- 48
                ctx.Progress.Chunk <- 48
                continueLayout ()
            else Ren.completeLayout ctx
        // Whether the document is all there: false while a package opened front
        // first (Office.Formats `word.openProgressive`, loaded with { progressive: true })
        // is still being read - the pages are then those of what is read so far.
        state?isLoadComplete <- fun () -> loading.IsNone
        state?whenLoadComplete <- fun () ->
            JS.Constructors.Promise.Create(fun (resolve: obj -> unit) _ ->
                match loading with
                | None -> resolve (box true)
                | Some importing ->
                    loadWaiters.Add (fun () -> resolve (box true))
                    importing.Package?Progressive?start () |> ignore)
        state?completeLoad <- fun () -> ensureLoaded ()
        state?isLayoutComplete <- fun () -> Ren.layoutIsComplete ctx
        state?whenLayoutComplete <- fun () ->
            JS.Constructors.Promise.Create(fun (resolve: obj -> unit) _ ->
                Ren.whenLayoutComplete ctx (fun () -> resolve (box (pageFramesToJs ctx (Ren.pageFrames ctx (current ())))))
                continueLayout ())
        state?onLayoutProgress <- fun (callback: obj -> unit) ->
            progressListeners.Add callback
            continueLayout ()
            fun () -> progressListeners.Remove callback |> ignore
        state?setImageCache <- fun (cache: obj) ->
            ctx.ImageCache <- (if isNull cache then None else Some cache)
            ctx.LastPaint <- None
        // The same measurers (and text painter) as an editor: with them the
        // viewer breaks lines and pages exactly as that editor does.
        state?setTextMeasurers <- fun (measureText: Layout.TextMeasurer) (measureRange: Layout.TextRangeMeasurer) ->
            Ren.setTextMeasurers ctx (if isNull (box measureText) then None else Some measureText) (if isNull (box measureRange) then None else Some measureRange)
        state?setTextMeasurer <- fun (measureText: Layout.TextMeasurer) ->
            Ren.setTextMeasurers ctx (if isNull (box measureText) then None else Some measureText) ctx.TextRangeMeasurer
        state?setTextRangeMeasurer <- fun (measureRange: Layout.TextRangeMeasurer) ->
            Ren.setTextMeasurers ctx ctx.TextMeasurer (if isNull (box measureRange) then None else Some measureRange)
        // `painterFor(context2d)` returns the painter for that canvas context
        // (text, style, x, baseline, width) -> true when it painted the text.
        state?setTextPathPainter <- fun (painterFor: System.Func<CanvasRenderingContext2D, System.Func<string, Paint.TextPaintStyle, float, float, float, bool>>) ->
            Ren.setTextPathPainterFactory ctx painterFor
        state?clearTextMeasurer <- fun () ->
            Ren.setTextPathPainter ctx None None
            Ren.setTextMeasurers ctx None None
        state?getPageFrames <- fun () ->
            continueLayout ()
            pageFramesToJs ctx (Ren.pageFrames ctx (current ()))
        state?getPageCount <- fun () ->
            ensureLoaded ()
            Ren.pageCount ctx (current ())
        state?getDocumentSize <- fun () -> documentSizeToJs ctx (current ())
        state?renderPage <- fun (target: HTMLCanvasElement) (pageIndex: int) (scale: float) ->
            // (A page there is not yet: of the whole document.)
            if loading.IsSome && pageIndex >= List.length (Ren.pageFrames ctx (current ())) then ensureLoaded ()
            Ren.renderPage ctx (current ()) target pageIndex scale
        // The font faces the loaded document is set in: [{ family, bold, italic }].
        state?getRequiredFonts <- fun () ->
            ensureLoaded ()
            Ren.requiredFonts ctx (current ())
        // Counters of the viewer's renderer (as the editor's getRenderStats).
        state?getRenderStats <- fun () -> Ren.renderStats ctx (current ())
        state

    /// The font faces an editor package's document is set in, before anything
    /// is shown: [{ family, bold, italic }], after style resolution, with the
    /// faces the engine itself uses (list markers in Symbol / Wingdings /
    /// Courier New, notes). `import { getRequiredFonts } from '.../word/WordViewer.js'`
    /// (the editor API has the same function, which also takes an editor).
    /// Of a package that is still being read (Office.Formats `word.openProgressive`)
    /// it is the faces of the styles and of what is read so far: see
    /// onFontsDiscovered for those the rest of the document turns out to need.
    let getRequiredFonts (package: obj) : obj array =
        Ren.requiredFontsOfDocument (Docx.importFromEditorPackage package)

    [<Emit("($0.family + '|' + ($0.bold ? 'b' : '') + ($0.italic ? 'i' : ''))")>]
    let private faceKey (_face: obj) : string = jsNative

    /// Be told the font faces a package turns out to need as the rest of it
    /// is read: `callback(faces)` with the faces ([{ family, bold, italic }])
    /// that getRequiredFonts(package) did not name when this was called and
    /// no earlier call of the callback named. A host fetches them while the
    /// text that needs them is still to be laid out. Returns the function that
    /// ends it; nothing is ever told of a package that is read.
    let onFontsDiscovered (package: obj) (callback: obj array -> unit) : unit -> unit =
        match Docx.progressiveImportOf package with
        | Some importing when not importing.Complete ->
            let known = System.Collections.Generic.HashSet<string>()
            for face in Ren.requiredFontsOfDocument importing.Document do known.Add (faceKey face) |> ignore
            let listener = fun (step: Docx.ImportStep) ->
                if not step.Appended.IsEmpty then
                    // (The faces of the blocks that were read, in the document's styles.)
                    let found =
                        Ren.requiredFontsOfDocument { importing.Document with Blocks = step.Appended; Selection = None }
                        |> Array.filter (fun face -> known.Add (faceKey face))
                    if found.Length > 0 then callback found
            importing.Listeners.Add listener
            fun () -> importing.Listeners.Remove listener |> ignore
        | _ -> ignore

    /// `import createViewer from '.../word/WordViewer.js'`.
    [<ExportDefault>]
    let viewer : HTMLCanvasElement -> obj = createViewer

#endif
