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
        ]

    /// Create a read-only viewer. `canvas` is used for text measurement (and by
    /// `render`, which paints the whole document onto it); pages are painted
    /// onto canvases of the host's with `renderPage`.
    ///   load(editorPackage)            show a package from Office.Formats (forEditorPackage)
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
        state?load <- fun (package: obj) ->
            // A viewer shows no caret.
            let doc = { Docx.importFromEditorPackage package with Selection = None }
            state?document <- doc
            Ren.clearLayoutCache ctx
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
        state?getPageFrames <- fun () -> pageFramesToJs ctx (Ren.pageFrames ctx (current ()))
        state?getPageCount <- fun () -> Ren.pageCount ctx (current ())
        state?getDocumentSize <- fun () -> documentSizeToJs ctx (current ())
        state?renderPage <- fun (target: HTMLCanvasElement) (pageIndex: int) (scale: float) -> Ren.renderPage ctx (current ()) target pageIndex scale
        // The font faces the loaded document is set in: [{ family, bold, italic }].
        state?getRequiredFonts <- fun () -> Ren.requiredFonts ctx (current ())
        // Counters of the viewer's renderer (as the editor's getRenderStats).
        state?getRenderStats <- fun () -> Ren.renderStats ctx (current ())
        state

    /// The font faces an editor package's document is set in, before anything
    /// is shown: [{ family, bold, italic }], after style resolution, with the
    /// faces the engine itself uses (list markers in Symbol / Wingdings /
    /// Courier New, notes). `import { getRequiredFonts } from '.../word/WordViewer.js'`
    /// (the editor API has the same function, which also takes an editor).
    let getRequiredFonts (package: obj) : obj array =
        Ren.requiredFontsOfDocument (Docx.importFromEditorPackage package)

    /// `import createViewer from '.../word/WordViewer.js'`.
    [<ExportDefault>]
    let viewer : HTMLCanvasElement -> obj = createViewer

#endif
