module Rbt.Spreadsheet.Spreadsheet

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Grid.Dom
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.XlsxImporter
module SpreadsheetCalc = Rbt.Spreadsheet.Calc
module SpreadsheetLayout = Rbt.Spreadsheet.Layout
module SpreadsheetPaint = Rbt.Spreadsheet.Paint
module SpreadsheetCellText = Rbt.Spreadsheet.CellText
module SpreadsheetInput = Rbt.Spreadsheet.Input

// Additional helpers
[<Emit("$0.scrollLeft")>]
let private getScrollLeft (el: Element): float = jsNative

[<Emit("$0.clientWidth")>]
let private getClientWidth (el: Element): float = jsNative

[<Emit("$0.clientX")>]
let private getClientX (e: obj): float = jsNative

[<Emit("$0.clientY")>]
let private getClientY (e: obj): float = jsNative

[<Emit("$0.getBoundingClientRect().left")>]
let private getBoundingLeft (el: Element): float = jsNative

[<Emit("$0.getBoundingClientRect().top")>]
let private getBoundingTop (el: Element): float = jsNative

[<Emit("$0.width = Math.ceil($1); $0.height = Math.ceil($2)")>]
let private setCanvasSize (canvas: Element) (width: float) (height: float): unit = jsNative

[<Emit("$0.getContext('2d')")>]
let private getCanvasContext (canvas: Element): obj = jsNative

[<Emit("$0.save()")>]
let private canvasSave (ctx: obj): unit = jsNative

[<Emit("$0.restore()")>]
let private canvasRestore (ctx: obj): unit = jsNative

[<Emit("$0.setSelectionRange($0.value.length, $0.value.length)")>]
let private caretToEnd (input: Element): unit = jsNative

[<Emit("setTimeout($0, $1)")>]
let private setTimeoutMs (callback: unit -> unit) (delay: int): int = jsNative

[<Emit("clearTimeout($0)")>]
let private clearTimeoutId (id: int): unit = jsNative

[<Emit("($0 === document.activeElement || $0.contains(document.activeElement))")>]
let private holdsFocus (el: Element): bool = jsNative

[<Emit("$0.focus({ preventScroll: true })")>]
let private focusQuietly (el: Element): unit = jsNative

[<Emit("$0.stopPropagation()")>]
let private stopPropagation (e: obj): unit = jsNative

[<Emit("new Date().getFullYear()")>]
let private currentYear (): int = jsNative

[<Emit("$0.insertBefore($1, $0.firstChild)")>]
let private prependChild (parent: Element) (child: Element): unit = jsNative

// Vector-path + transform emits (for VecPath / VecGroup — the chart scene).
[<Emit("$0.beginPath()")>]
let private canvasBeginPath (ctx: obj): unit = jsNative
[<Emit("$0.moveTo($1, $2)")>]
let private canvasMoveTo (ctx: obj) (x: float) (y: float): unit = jsNative
[<Emit("$0.lineTo($1, $2)")>]
let private canvasLineTo (ctx: obj) (x: float) (y: float): unit = jsNative
[<Emit("$0.bezierCurveTo($1, $2, $3, $4, $5, $6)")>]
let private canvasBezierTo (ctx: obj) (x1: float) (y1: float) (x2: float) (y2: float) (x: float) (y: float): unit = jsNative
[<Emit("$0.closePath()")>]
let private canvasClosePath (ctx: obj): unit = jsNative
[<Emit("$0.fillStyle = $1; $0.fill()")>]
let private canvasFillPath (ctx: obj) (color: string): unit = jsNative
[<Emit("$0.strokeStyle = $1; $0.lineWidth = $2; $0.stroke()")>]
let private canvasStrokePath (ctx: obj) (color: string) (width: float): unit = jsNative
[<Emit("$0.transform($1, $2, $3, $4, $5, $6)")>]
let private canvasTransform (ctx: obj) (a: float) (b: float) (c: float) (d: float) (e: float) (f: float): unit = jsNative

[<Emit("$0.beginPath(); $0.rect($1, $2, $3, $4); $0.clip()")>]
let private canvasClipRect (ctx: obj) (x: float) (y: float) (width: float) (height: float): unit = jsNative

[<Emit("$0.fillStyle = $5; $0.fillRect($1, $2, $3, $4)")>]
let private canvasFillRect (ctx: obj) (x: float) (y: float) (width: float) (height: float) (color: string): unit = jsNative

[<Emit("$0.strokeStyle = $5; $0.lineWidth = $6; $0.strokeRect($1, $2, $3, $4)")>]
let private canvasStrokeRect (ctx: obj) (x: float) (y: float) (width: float) (height: float) (color: string) (lineWidth: float): unit = jsNative

[<Emit("$0.strokeStyle = $5; $0.lineWidth = $6; $0.beginPath(); $0.moveTo($1, $2); $0.lineTo($3, $4); $0.stroke()")>]
let private canvasLine (ctx: obj) (x1: float) (y1: float) (x2: float) (y2: float) (color: string) (lineWidth: float): unit = jsNative

[<Emit("(function(text,fontSize,fontName,bold,italic){const c=globalThis.__rbSpreadsheetMeasureCanvas||(globalThis.__rbSpreadsheetMeasureCanvas=document.createElement('canvas'));const ctx=c.getContext('2d');const family=fontName||'Arial';ctx.font=(bold?'bold ':'')+(italic?'italic ':'')+fontSize+'px '+family+', sans-serif';return ctx.measureText(text).width;})($0,$1,$2,$3,$4)")>]
let private browserTextWidth (text: string) (fontSize: float) (fontName: string) (bold: bool) (italic: bool): float = jsNative

[<Emit("$0.fillStyle = $5; $0.font = ($6 ? 'bold ' : '') + ($7 ? 'italic ' : '') + $4 + 'px ' + $9 + ', sans-serif'; $0.textBaseline = 'alphabetic'; $0.textAlign = $8; $0.fillText($1, $2, $3)")>]
let private canvasText (ctx: obj) (text: string) (x: float) (y: float) (fontSize: float) (color: string) (bold: bool) (italic: bool) (align: string) (fontName: string): unit = jsNative

// Draw text rotated `angle` radians about an anchor (tx,ty): translate to the
// anchor, rotate (canvas rotate is clockwise for +angle, matching screen
// y-down), draw left-aligned at the origin, restore. Used for cell-text
// rotation; the caller positions the anchor so the rotated bbox aligns.
[<Emit("$0.save(); $0.translate($3, $4); $0.rotate($5); $0.fillStyle = $6; $0.font = ($7 ? 'bold ' : '') + ($8 ? 'italic ' : '') + $2 + 'px ' + $9 + ', sans-serif'; $0.textBaseline = 'alphabetic'; $0.textAlign = 'left'; $0.fillText($1, 0, 0); $0.restore()")>]
let private canvasTextRotated (ctx: obj) (text: string) (fontSize: float) (tx: float) (ty: float) (angle: float) (color: string) (bold: bool) (italic: bool) (fontName: string): unit = jsNative

// Pictures: one object URL and one decoded <img> per picture's data, shared
// by the DOM renderer and the canvas painter.
[<Emit("(($0.__rbtPictureUrl) || ($0.__rbtPictureUrl = URL.createObjectURL(new Blob([$0], { type: $1 }))))")>]
let private pictureUrl (data: byte[]) (contentType: string): string = jsNative

[<Emit("(($0.__rbtPictureElement) || ($0.__rbtPictureElement = (() => { const img = new Image(); img.src = ($0.__rbtPictureUrl) || ($0.__rbtPictureUrl = URL.createObjectURL(new Blob([$0], { type: $1 }))); return img; })()))")>]
let private pictureElement (data: byte[]) (contentType: string): Element = jsNative

[<Emit("($0.complete && $0.naturalWidth > 0)")>]
let private pictureReady (image: Element): bool = jsNative

[<Emit("$0.drawImage($1, $2, $3, $4, $5)")>]
let private canvasDrawImage (ctx: obj) (image: Element) (x: float) (y: float) (width: float) (height: float): unit = jsNative

[<Emit("(($0.__rbtPictureUrl) ? (URL.revokeObjectURL($0.__rbtPictureUrl), delete $0.__rbtPictureUrl, delete $0.__rbtPictureElement) : undefined)")>]
let private releasePicture (data: byte[]): unit = jsNative

[<Emit("(typeof ResizeObserver === 'function' ? (() => { const o = new ResizeObserver(() => $1()); o.observe($0); return () => o.disconnect(); })() : (() => {}))")>]
let private observeResize (el: Element) (callback: unit -> unit): unit -> unit = jsNative

[<Emit("($0.parentNode ? $0.parentNode.removeChild($0) : undefined)")>]
let private detach (el: Element): unit = jsNative

let private sameRef (a: 'T) (b: 'T) = obj.ReferenceEquals(a, b)

/// What the package writer resolves to (a Blob in the browser) as bytes.
[<Emit("$0.then((written) => (written instanceof Uint8Array ? written : (written && typeof written.arrayBuffer === 'function' ? written.arrayBuffer().then((buffer) => new Uint8Array(buffer)) : new Uint8Array(written))))")>]
let private packageBytes (written: JS.Promise<obj>): JS.Promise<byte[]> = jsNative

/// Whether two models of a sheet hold the same document content: everything
/// that is part of the workbook, as opposed to view and editing state
/// (selection, edit buffer, clipboard, search, undo history).
let private sameSheetContent (a: SpreadsheetModel) (b: SpreadsheetModel) =
    sameRef a b
    || (sameRef a.Cells b.Cells
        && sameRef a.RichRuns b.RichRuns
        && sameRef a.RowHeights b.RowHeights
        && sameRef a.ColumnWidths b.ColumnWidths
        && sameRef a.Validations b.Validations
        && (sameRef a.MergedCells b.MergedCells || a.MergedCells = b.MergedCells)
        && sameRef a.ConditionalFormats b.ConditionalFormats
        && (sameRef a.Tables b.Tables || a.Tables = b.Tables)
        && sameRef a.Filters b.Filters
        && (sameRef a.PageSetup b.PageSetup || a.PageSetup = b.PageSetup)
        && a.RowCount = b.RowCount
        && a.ColCount = b.ColCount
        && a.FrozenRows = b.FrozenRows
        && a.FrozenCols = b.FrozenCols
        && a.ShowGridLines = b.ShowGridLines
        && sameRef a.Drawings b.Drawings
        && sameRef a.Threads b.Threads
        && a.Macros = b.Macros)

/// A comment thread as the text of a tooltip: the comment, then its replies.
let private threadText (thread: CommentThread) =
    let line (comment: CommentThread) =
        match comment.Author with
        | Some author -> sprintf "%s:\n%s" author comment.Text
        | None -> comment.Text
    line thread :: (thread.Replies |> List.map (fun reply -> "\u21B3 " + line reply))
    |> String.concat "\n\n"

/// Spreadsheet component — interactively edits a Workbook (a collection
/// of sheets sharing strings, styles, theme, defined names). The class
/// renders one sheet at a time; chrome layers a tab strip on top via
/// SheetNames / ActiveSheetIndex / SwitchSheet + the OnSheetChanged event.
///
/// Internally `model` stays the mutable "active sheet" for minimal diff
/// across the existing 240+ references; `sheets` is the parallel array
/// holding every sheet's state. On switch we write `model` back to
/// `sheets.[activeSheetIndex]`, advance the index, then read the new
/// sheet into `model`. Selection/scroll/edit state lives in
/// SpreadsheetModel so each sheet remembers where the user was, the way
/// Excel does.
[<AttachMembers>]
type Spreadsheet(container: Element) =
    let mutable sheets: SpreadsheetModel array = [| createEmpty 1000 26 |]
    let mutable sheetNames: string array = [| "Sheet1" |]
    let mutable activeSheetIndex = 0
    let mutable model = sheets.[0]
    let sheetChangedListeners = System.Collections.Generic.List<int -> unit>()
    /// OnChange / OnSelectionChange subscribers.
    let changeListeners = ResizeArray<obj -> unit>()
    let selectionListeners = ResizeArray<obj -> unit>()
    /// What the listeners were last told about: every sheet's model and name,
    /// and the selection of the active sheet.
    let mutable notifiedSheets: SpreadsheetModel array = [||]
    let mutable notifiedNames: string array = [||]
    let mutable notifiedSelection: (int * Selection) option = None
    /// Undo what Mount set up outside the grid's own nodes (document-level
    /// listeners, observers, overlays appended to <body>).
    let cleanups = ResizeArray<unit -> unit>()
    let mutable disposed = false
    let mutable workbook: Workbook option = None  // Optional workbook for table references
    /// The XLSX file the current workbook was imported from, for ExportXlsx:
    /// what the editor does not model is carried over from it.
    let mutable xlsxSource: Rbt.Spreadsheet.XlsxBridge.XlsxSource option = None
    let mutable calcEngine: SpreadsheetCalc.CalcEngine option = None
    let mutable readOnly = false
    let mutable scrollTop = 0.0
    let mutable scrollLeft = 0.0
    let mutable containerHeight = 600.0
    let mutable containerWidth = 800.0
    let mutable useSharedLayoutGeometry = true

    let rowHeight = 25.0
    let colWidth = 100.0
    let headerHeight = 30.0
    let rowHeaderWidth = 50.0

    let mutable gridContainer: Element option = None
    let mutable contentElement: Element option = None
    let cellPool = System.Collections.Generic.Dictionary<string, Element>()
    let chromeNodes = ResizeArray<Element>()
    let mutable lastRenderedModel: SpreadsheetModel option = None

    let currentViewport (): SpreadsheetLayout.Viewport = {
        ScrollX = scrollLeft
        ScrollY = scrollTop
        Width = containerWidth
        Height = containerHeight
        HeaderHeight = headerHeight
        RowHeaderWidth = rowHeaderWidth
        Measurer = Some (fun text fontSize fontName bold italic ->
            browserTextWidth text fontSize (fontName |> Option.defaultValue "Arial") bold italic)
    }

    let restoreCellOptions (changes: seq<CellAddress * CellData option>) (m: SpreadsheetModel) =
        let mutable result = m
        for addr, value in changes do
            match value with
            | Some cellData -> result <- setCell addr cellData result
            | None -> result <- clearCell addr result
        result

    let restoreCellMap (cells: Map<CellAddress, CellData>) (m: SpreadsheetModel) =
        { m with Cells = cells } |> invalidateFilterProjection

    let invalidateCalcEngine () =
        calcEngine <- None

    let sheetIdForIndex (index: int) =
        match workbook with
        | Some wb when index >= 0 && index < List.length wb.SheetOrder ->
            List.item index wb.SheetOrder
        | _ -> sprintf "sheet%d" (index + 1)

    let currentWorkbookSnapshot () =
        let sheetEntries =
            [ for i in 0 .. sheets.Length - 1 do
                let sheetId = sheetIdForIndex i
                let sheetModel = if i = activeSheetIndex then model else sheets.[i]
                let name =
                    if i < sheetNames.Length then sheetNames.[i]
                    else sprintf "Sheet%d" (i + 1)
                let visible =
                    match workbook |> Option.bind (fun wb -> Map.tryFind sheetId wb.Sheets) with
                    | Some sheet -> sheet.Visible
                    | None -> true
                sheetId, { Id = sheetId; Name = name; Model = sheetModel; Visible = visible } ]
        let activeSheet = sheetIdForIndex activeSheetIndex
        match workbook with
        | Some wb ->
            { wb with
                Sheets = Map.ofList sheetEntries
                SheetOrder = sheetEntries |> List.map fst
                ActiveSheet = activeSheet }
        | None ->
            { Sheets = Map.ofList sheetEntries
              SheetOrder = sheetEntries |> List.map fst
              ActiveSheet = activeSheet
              Tables = Map.empty
              Names = Map.empty }

    let installWorkbookSnapshot (wb: Workbook) =
        for i in 0 .. sheets.Length - 1 do
            let sheetId = sheetIdForIndex i
            match Map.tryFind sheetId wb.Sheets with
            | Some sheet -> sheets.[i] <- sheet.Model
            | None -> ()
        model <- sheets.[activeSheetIndex]
        workbook <- workbook |> Option.map (fun existing ->
            { existing with
                Sheets = wb.Sheets
                SheetOrder = wb.SheetOrder
                ActiveSheet = wb.ActiveSheet
                Tables = wb.Tables
                Names = wb.Names })

    /// The workbook that conditional-format and validation formulas are
    /// evaluated against; refreshed at each render.
    let mutable ruleWorkbook: Workbook option = None

    /// Evaluates a rule's formula at a cell of the active sheet, with the
    /// other sheets and defined names in reach.
    let ruleEvaluator : Rbt.Spreadsheet.ConditionalFormatting.FormulaEvaluator =
        fun formula rows cols sheetModel cell ->
            evaluateFormulaShifted ruleWorkbook (Some (sheetIdForIndex activeSheetIndex)) formula rows cols sheetModel cell

    let refreshRuleWorkbook () =
        ruleWorkbook <- Some (currentWorkbookSnapshot())
        Rbt.Spreadsheet.ConditionalFormatting.setFormulaEvaluator ruleEvaluator

    let renderPortablePaintToCanvas (canvas: Element) =
        if not model.ConditionalFormats.IsEmpty then refreshRuleWorkbook()
        let viewport = currentViewport()
        let ctx = getCanvasContext canvas
        setCanvasSize canvas viewport.Width viewport.Height

        // Recursive renderer for vector paint items (VecPath / VecGroup) — a
        // group concatenates its 2×3 transform and draws its children.
        let rec drawVec (item: SpreadsheetPaint.PaintItem) =
            match item with
            | SpreadsheetPaint.VecPath (segs, fill, stroke) ->
                canvasBeginPath ctx
                for s in segs do
                    match s with
                    | SpreadsheetPaint.VMove (x, y) -> canvasMoveTo ctx x y
                    | SpreadsheetPaint.VLine (x, y) -> canvasLineTo ctx x y
                    | SpreadsheetPaint.VCubic (x1, y1, x2, y2, x, y) -> canvasBezierTo ctx x1 y1 x2 y2 x y
                    | SpreadsheetPaint.VClose -> canvasClosePath ctx
                (match fill with Some c -> canvasFillPath ctx c | None -> ())
                (match stroke with Some (c, w) -> canvasStrokePath ctx c w | None -> ())
            | SpreadsheetPaint.VecGroup ((a, b, c, d, e, f), children) ->
                canvasSave ctx
                canvasTransform ctx a b c d e f
                children |> List.iter drawVec
                canvasRestore ctx
            | SpreadsheetPaint.FillRect (r, col) -> canvasFillRect ctx r.X r.Y r.Width r.Height col
            | SpreadsheetPaint.StrokeRect (r, col, w) -> canvasStrokeRect ctx r.X r.Y r.Width r.Height col w
            | _ -> ()
        for painted in SpreadsheetPaint.buildViewportPaintItems model viewport do
            // viewportClipFor narrows BodyZone to the area past the frozen
            // rows/cols so the body paint can't bleed into the pinned bands.
            let clip = SpreadsheetLayout.viewportClipFor model viewport painted.Zone
            canvasSave ctx
            canvasClipRect ctx clip.X clip.Y clip.Width clip.Height
            match painted.Item with
            | SpreadsheetPaint.FillRect (rect, color) ->
                canvasFillRect ctx rect.X rect.Y rect.Width rect.Height color
            | SpreadsheetPaint.StrokeRect (rect, color, width) ->
                canvasStrokeRect ctx rect.X rect.Y rect.Width rect.Height color width
            | SpreadsheetPaint.GridLine (x1, y1, x2, y2, color, width) ->
                // Pixel-snap axis-aligned strokes (mirrors the native
                // SpreadsheetGrid path): an odd-width line centred on an
                // integer straddles the canvas pixel grid and anti-aliases to
                // ~50% grey edges. Round the width and centre odd widths on
                // n+0.5, even widths on n, so the stroke covers whole pixels.
                let w = max 1.0 (System.Math.Round width)
                let snap c = if int w % 2 = 1 then floor c + 0.5 else System.Math.Round c
                if abs (y1 - y2) < 0.01 then canvasLine ctx x1 (snap y1) x2 (snap y1) color w
                elif abs (x1 - x2) < 0.01 then canvasLine ctx (snap x1) y1 (snap x1) y2 color w
                else canvasLine ctx x1 y1 x2 y2 color w
            | SpreadsheetPaint.VecPath _ | SpreadsheetPaint.VecGroup _ -> drawVec painted.Item
            | SpreadsheetPaint.DrawImage (rect, image) ->
                // A picture that is not decoded yet shows its frame; the canvas
                // is painted again by the caller once it has loaded.
                let element = pictureElement image.Data image.ContentType
                if pictureReady element then canvasDrawImage ctx element rect.X rect.Y rect.Width rect.Height
                else
                    canvasFillRect ctx rect.X rect.Y rect.Width rect.Height "#f6f8fa"
                    canvasStrokeRect ctx rect.X rect.Y rect.Width rect.Height "#8a94a6" 1.0
            | SpreadsheetPaint.DrawText text ->
                let align =
                    match text.Align with
                    | SpreadsheetPaint.AlignCenter -> "center"
                    | SpreadsheetPaint.AlignRight -> "right"
                    | SpreadsheetPaint.AlignLeft -> "left"
                let x =
                    match text.Align with
                    | SpreadsheetPaint.AlignCenter -> text.Rect.X + text.Rect.Width / 2.0
                    | SpreadsheetPaint.AlignRight -> text.Rect.X + text.Rect.Width
                    | SpreadsheetPaint.AlignLeft -> text.Rect.X
                // Approximate Calibri's em-relative metrics (asc 0.952, desc
                // -0.266, lineHeight 1.218). Browser canvas doesn't expose
                // font metrics portably, so we accept a small drift for
                // non-Calibri fonts — Excel ships Calibri as the default.
                let ascent = text.FontSize * 0.952
                let lineHeight = text.FontSize * 1.218
                let blockHeight = lineHeight * float text.LineCount
                let blockTopY =
                    match text.VerticalAlign with
                    | SpreadsheetPaint.VAlignTop -> text.Rect.Y
                    | SpreadsheetPaint.VAlignMiddle -> text.Rect.Y + (text.Rect.Height - blockHeight) / 2.0
                    | SpreadsheetPaint.VAlignBottom -> text.Rect.Y + text.Rect.Height - blockHeight
                let baseline = blockTopY + float text.LineIndex * lineHeight + ascent + text.BaselineShift
                let fontName = text.FontName |> Option.defaultValue "Arial"
                let isRotated = text.Rotation > 0 && text.Rotation <> 255
                if isRotated then
                    // Same rule as the native + PDF paths (derived from
                    // build-rotation-sweep.vbs): rotate the run about its anchor,
                    // then place the rotated bbox right-aligned when
                    // 90 ≤ textRotation < 180 else left-aligned, always
                    // bottom-aligned. Screen is y-down so a negative angle is CCW.
                    let theta =
                        if text.Rotation <= 90 then -(float text.Rotation) * System.Math.PI / 180.0
                        else float (text.Rotation - 90) * System.Math.PI / 180.0
                    let totalWidth = browserTextWidth text.Text text.FontSize fontName text.Bold text.Italic
                    let descent = text.FontSize * 0.266
                    let ax0 = text.Rect.X + 4.0
                    let ay0 = text.Rect.Y + text.Rect.Height - 2.0
                    let cosS, sinS = cos theta, sin theta
                    let rotPt px py =
                        let dx, dy = px - ax0, py - ay0
                        ax0 + dx * cosS - dy * sinS, ay0 + dx * sinS + dy * cosS
                    let corners =
                        [ rotPt ax0 (ay0 - ascent); rotPt (ax0 + totalWidth) (ay0 - ascent)
                          rotPt ax0 (ay0 + descent); rotPt (ax0 + totalWidth) (ay0 + descent) ]
                    let minX = corners |> List.map fst |> List.min
                    let maxX = corners |> List.map fst |> List.max
                    let maxY = corners |> List.map snd |> List.max
                    let rightAlign = text.Rotation >= 90 && text.Rotation < 180
                    let shiftX =
                        if rightAlign then (text.Rect.X + text.Rect.Width - 2.0) - maxX
                        else (text.Rect.X + 2.0) - minX
                    let shiftY = (text.Rect.Y + text.Rect.Height - 2.0) - maxY
                    canvasTextRotated ctx text.Text text.FontSize (ax0 + shiftX) (ay0 + shiftY) theta text.Color text.Bold text.Italic fontName
                else
                canvasText ctx text.Text x baseline text.FontSize text.Color text.Bold text.Italic align fontName
                if not isRotated && (text.Underline || text.UnderlineDouble || text.Strikethrough) then
                    let textWidth = browserTextWidth text.Text text.FontSize fontName text.Bold text.Italic
                    let startX =
                        match text.Align with
                        | SpreadsheetPaint.AlignCenter -> x - textWidth / 2.0
                        | SpreadsheetPaint.AlignRight -> x - textWidth
                        | SpreadsheetPaint.AlignLeft -> x
                    let endX = startX + textWidth
                    if text.Underline || text.UnderlineDouble then
                        let y = baseline + max 1.0 (text.FontSize * 0.12)
                        canvasLine ctx startX y endX y text.Color 1.0
                        if text.UnderlineDouble then
                            canvasLine ctx startX (y + 2.0) endX (y + 2.0) text.Color 1.0
                    if text.Strikethrough then
                        let y = baseline - text.FontSize * 0.30
                        canvasLine ctx startX y endX y text.Color 1.0
            canvasRestore ctx
        // Selection overlay: drawn after all cell content so the mint range
        // fill and green outline sit on top of borders + text. Selection is
        // intentionally NOT emitted by `buildCellPaintItems` so the native
        // grid widget can cache the cell items across cursor moves.
        // `selectionPaintItems` emits rects in absolute (scroll-free) coords
        // so they need the body-zone translation manually here — the native
        // path uses an outer ScrollView that translates the whole canvas,
        // but on web we paint into a fixed canvas and shift each item.
        let selectionItems = SpreadsheetPaint.selectionPaintItems model viewport
        if not (List.isEmpty selectionItems) then
            let clip = SpreadsheetLayout.viewportClipFor model viewport SpreadsheetLayout.BodyZone
            canvasSave ctx
            canvasClipRect ctx clip.X clip.Y clip.Width clip.Height
            let shift (r: SpreadsheetLayout.Rect) : SpreadsheetLayout.Rect =
                { r with X = r.X - viewport.ScrollX; Y = r.Y - viewport.ScrollY }
            for item in selectionItems do
                match item with
                | SpreadsheetPaint.FillRect (rect, color) ->
                    let r = shift rect
                    canvasFillRect ctx r.X r.Y r.Width r.Height color
                | SpreadsheetPaint.StrokeRect (rect, color, width) ->
                    let r = shift rect
                    canvasStrokeRect ctx r.X r.Y r.Width r.Height color width
                | _ -> ()
            canvasRestore ctx

    /// Full recalculation of the whole workbook. A formula on any sheet can
    /// read the sheet that was just edited, so every sheet is recalculated
    /// (formulas, then dynamic-array spills, then formulas again so readers of
    /// spilled cells settle), not only the visible one.
    let recalculateFormulas () =
        invalidateCalcEngine()
        currentWorkbookSnapshot()
        |> recalculateWorkbook
        |> installWorkbookSnapshot

    /// Insert or delete rows/columns on the active sheet as one undoable
    /// step. The edit and its undo both go through the workbook so formulas on
    /// other sheets, defined names and structured-reference tables follow.
    let applyStructureEdit (edit: StructureEdit) (description: string) =
        let sheetId = sheetIdForIndex activeSheetIndex
        let edited, undoInfo = Rbt.Spreadsheet.Workbook.applyStructureEdit sheetId edit (currentWorkbookSnapshot())
        installWorkbookSnapshot edited
        recalculateFormulas()
        // Undo()/Redo() recalculate after running the command.
        let command = {
            Undo = fun _ ->
                currentWorkbookSnapshot()
                |> Rbt.Spreadsheet.Workbook.undoStructureEdit undoInfo
                |> installWorkbookSnapshot
                model
            Redo = fun _ ->
                currentWorkbookSnapshot()
                |> Rbt.Spreadsheet.Workbook.applyStructureEdit sheetId edit
                |> fst
                |> installWorkbookSnapshot
                model
            Description = description
        }
        model <- pushCommand command model

    /// Sort the active sheet by a column as one undoable step, then
    /// recalculate: rows moved, so every formula reading them is stale.
    let sortActiveSheet (col: int) (ascending: bool) (description: string) =
        let oldCells = model.Cells
        let oldRichRuns = model.RichRuns
        let oldSortState = model.SortState
        model <- sortByColumn col ascending model
        recalculateFormulas()
        let command = {
            Undo = fun m -> { restoreCellMap oldCells m with RichRuns = oldRichRuns; SortState = oldSortState }
            Redo = fun m -> sortByColumn col ascending m
            Description = description
        }
        model <- pushCommand command model

    /// Cut the selection: copy it to the clipboard, clear the cells and
    /// recalculate, as one undoable step.
    let cutSelection () =
        let selection = model.Selection
        let clipboardData = serializeSelection selection model true
        let oldCells = model.Cells
        model <- { model with Clipboard = Some clipboardData }
        model <- clearSelectionCells selection model
        recalculateFormulas()
        if not (obj.ReferenceEquals(oldCells, model.Cells)) then
            let command = {
                Undo = fun m -> restoreCellMap oldCells m
                Redo = fun m -> clearSelectionCells selection m
                Description = "Cut"
            }
            model <- pushCommand command model

    /// Paste the clipboard at the selection as one undoable step. A pasted
    /// cut is a move: every formula and defined name in the workbook that
    /// referred to the moved cells is re-pointed at their new place.
    let pasteClipboardAtSelection () =
        match model.Clipboard with
        | None -> ()
        | Some clipboardData ->
            let anchors = pasteAnchors model.Selection clipboardData
            let touched =
                [ for anchor in anchors do
                    for (relRow, relCol, _) in clipboardData.Cells do
                        yield { Row = anchor.Row + relRow; Col = anchor.Col + relCol } ]
                |> List.distinct
            let oldCells = touched |> List.map (fun addr -> addr, getCell addr model)
            for anchor in anchors do
                model <- pasteClipboard anchor clipboardData model
            let moveUndo =
                match anchors with
                | [ anchor ] when clipboardData.IsCut ->
                    let sheetId = sheetIdForIndex activeSheetIndex
                    let source =
                        clipboardData.SourceStartRow, clipboardData.SourceStartCol,
                        clipboardData.SourceEndRow, clipboardData.SourceEndCol
                    let delta = anchor.Row - clipboardData.SourceStartRow, anchor.Col - clipboardData.SourceStartCol
                    let repointed, undoInfo =
                        Rbt.Spreadsheet.Workbook.repointMovedCells sheetId source sheetId delta (currentWorkbookSnapshot())
                    installWorkbookSnapshot repointed
                    Some undoInfo
                | _ -> None
            recalculateFormulas()
            let newCells = touched |> List.map (fun addr -> addr, getCell addr model)
            let afterSelection = pasteSelection anchors clipboardData model.Selection
            model <- { model with Selection = afterSelection }
            let command = {
                Undo = fun m ->
                    model <- restoreCellOptions oldCells m
                    match moveUndo with
                    | Some undoInfo ->
                        currentWorkbookSnapshot()
                        |> Rbt.Spreadsheet.Workbook.undoRepointMovedCells undoInfo
                        |> installWorkbookSnapshot
                    | None -> ()
                    model
                Redo = fun m ->
                    model <- restoreCellOptions newCells m
                    match moveUndo with
                    | Some _ ->
                        // Re-run the move against the workbook as it is now.
                        match anchors with
                        | [ anchor ] ->
                            let sheetId = sheetIdForIndex activeSheetIndex
                            let source =
                                clipboardData.SourceStartRow, clipboardData.SourceStartCol,
                                clipboardData.SourceEndRow, clipboardData.SourceEndCol
                            let delta = anchor.Row - clipboardData.SourceStartRow, anchor.Col - clipboardData.SourceStartCol
                            currentWorkbookSnapshot()
                            |> Rbt.Spreadsheet.Workbook.repointMovedCells sheetId source sheetId delta
                            |> fst
                            |> installWorkbookSnapshot
                        | _ -> ()
                    | None -> ()
                    model
                Description = "Paste"
            }
            model <- pushCommand command model
            // A cut can be pasted once.
            if clipboardData.IsCut then
                model <- { model with Clipboard = None }

    /// Replace in the selected cell as one undoable step. True when it matched.
    let replaceInSelection (findText: string) (replaceText: string) (caseSensitive: bool) =
        let changedAddr =
            match model.Selection with
            | SingleCell (row, col) -> Some { Row = row; Col = col }
            | _ -> None
        let oldCells =
            match changedAddr with
            | Some addr -> [ addr, getCell addr model ]
            | None -> []
        match replaceAtSelection findText replaceText caseSensitive model with
        | Some newModel ->
            let newCells =
                match changedAddr with
                | Some addr -> [ addr, getCell addr newModel ]
                | None -> []
            let command = {
                Undo = restoreCellOptions oldCells
                Redo = restoreCellOptions newCells
                Description = "Replace"
            }
            model <- newModel
            // Formulas reading the replaced cell must see the new text.
            recalculateFormulas()
            model <- pushCommand command model
            true
        | None -> false

    /// Replace every match on the sheet as one undoable step; returns the count.
    let replaceEverywhere (findText: string) (replaceText: string) (caseSensitive: bool) =
        let newModel, count, oldCells, newCells = replaceAllWithChanges findText replaceText caseSensitive model
        if count > 0 then
            model <- newModel
            recalculateFormulas()
            let command = {
                Undo = restoreCellOptions oldCells
                Redo = restoreCellOptions newCells
                Description = "Replace All"
            }
            model <- pushCommand command model
        count

    let recalculateSingleCellEdit (addr: CellAddress) (cellData: CellData) =
        let wb = currentWorkbookSnapshot()
        let sheetCell = { Sheet = wb.ActiveSheet; Cell = addr }
        match SpreadsheetCalc.trySingleCellEdit sheetCell cellData wb calcEngine with
        | SpreadsheetCalc.IncrementalApplied (updatedWorkbook, updatedEngine) ->
            installWorkbookSnapshot updatedWorkbook
            calcEngine <- Some updatedEngine
        | SpreadsheetCalc.IncrementalFallback _ ->
            model <- setCell addr cellData model
            recalculateFormulas()

    let hitTestEvent (e: obj) =
        match gridContainer with
        | None -> SpreadsheetLayout.Empty
        | Some grid ->
            let x = getClientX e - getBoundingLeft grid
            let y = getClientY e - getBoundingTop grid
            SpreadsheetLayout.hitTest model (currentViewport()) x y

    let selectCellAddress (addr: CellAddress) =
        let selection =
            match findMergedCell addr model with
            | Some merge -> Range (merge.StartRow, merge.StartCol, merge.EndRow, merge.EndCol)
            | None -> SingleCell (addr.Row, addr.Col)

        model <- { model with Selection = selection }

    // Drag selection state
    let mutable isDragging = false
    let mutable dragStartCell: CellAddress option = None

    // Resize state
    let mutable isResizingColumn = false
    let mutable isResizingRow = false
    let mutable resizeColumnIndex: int option = None
    let mutable resizeRowIndex: int option = None
    let mutable resizeStartPos = 0.0
    let mutable resizeStartSize = 0.0
    let mutable resizeRenderScheduled = false
    let mutable scrollRenderScheduled = false

    /// Check if cell is in a table and get table info
    let getCellTable (row: int) (col: int): (Rbt.Spreadsheet.Workbook.Table * bool * bool) option =
        match workbook with
        | None -> None
        | Some wb ->
            wb.Tables
            |> Map.tryPick (fun _ table ->
                if row >= table.StartRow && row <= table.EndRow &&
                   col >= table.StartCol && col <= table.EndCol then
                    let isHeader = table.HasHeaders && row = table.StartRow
                    let isEvenRow = (row - table.StartRow - (if table.HasHeaders then 1 else 0)) % 2 = 0
                    Some (table, isHeader, isEvenRow)
                else
                    None
            )

    /// Get column X position (cumulative width of all columns before it)
    let getColumnX (col: int): float =
        if useSharedLayoutGeometry then
            SpreadsheetLayout.columnX model (currentViewport()) col
        else
            let mutable x = rowHeaderWidth
            for c = 0 to col - 1 do
                x <- x + getColumnWidth c model
            x

    /// Get row Y position (cumulative height of all rows before it)
    let getRowY (row: int): float =
        if useSharedLayoutGeometry then
            SpreadsheetLayout.rowY model (currentViewport()) row
        else
            let mutable y = headerHeight
            for r = 0 to row - 1 do
                y <- y + getRowHeight r model
            y

    /// Calculate visible range
    let calculateVisibleRange (): (int * int * int * int) =
        if useSharedLayoutGeometry then
            let range =
                SpreadsheetLayout.visibleRange model (currentViewport())
            (range.StartRow, range.EndRowExclusive, range.StartCol, range.EndColExclusive)
        else
            // For columns, find which ones are visible based on cumulative widths
            let mutable startCol = 0
            let mutable x = rowHeaderWidth
            while startCol < model.ColCount && x < scrollLeft do
                x <- x + getColumnWidth startCol model
                if x < scrollLeft then
                    startCol <- startCol + 1

            let mutable endCol = startCol
            let mutable xEnd = x
            while endCol < model.ColCount && xEnd < (scrollLeft + containerWidth) do
                xEnd <- xEnd + getColumnWidth endCol model
                endCol <- endCol + 1

            // For rows, find which ones are visible based on cumulative heights
            let mutable startRow = 0
            let mutable y = headerHeight
            while startRow < model.RowCount && y < scrollTop do
                y <- y + getRowHeight startRow model
                if y < scrollTop then
                    startRow <- startRow + 1

            let mutable endRow = startRow
            let mutable yEnd = y
            while endRow < model.RowCount && yEnd < (scrollTop + containerHeight) do
                yEnd <- yEnd + getRowHeight endRow model
                endRow <- endRow + 1

            (startRow, endRow, startCol, endCol)

    let cellKey (row: int) (col: int) = sprintf "%d:%d" row col

    let isFrozenBodyCell (row: int) (col: int) =
        (model.FrozenRows > 0 && row < model.FrozenRows) ||
        (model.FrozenCols > 0 && col < model.FrozenCols)

    let isRenderableCell (row: int) (col: int) =
        let addr = { Row = row; Col = col }
        match findMergedCell addr model with
        | Some merge when addr.Row <> merge.StartRow || addr.Col <> merge.StartCol -> false
        | _ -> true

    /// Check if row is visible (not filtered out)
    let isRowVisible (row: int): bool =
        if Map.isEmpty model.Filters then
            true
        else
            let visibleRows = getVisibleRows model
            Set.contains row visibleRows

    // --- Validation notices ---------------------------------------------------
    // Excel answers an invalid entry with a modal alert. Here the answer never
    // blocks: a "stop" rule refuses the entry (the cell keeps its old content
    // and editing ends), "warning" and "information" rules accept it, and in
    // every case a note with the rule's title and message appears in the grid
    // and goes away by itself. Hosts can also listen for it.
    let mutable validationNotice: (string * string * string * CellAddress) option = None   // style, title, message, cell
    let mutable noticeTimer: int option = None
    let validationListeners = System.Collections.Generic.List<string -> string -> string -> unit>()

    /// The edit box of the cell being edited, as rendered last.
    let mutable editInput: Element option = None

    /// When true the edit box places the caret after the seeded text instead
    /// of selecting it (typing a character over a cell starts the entry).
    let mutable editCaretAtEnd = false

    // Fill-handle drag state: the block being continued and where to.
    let mutable isFilling = false
    let mutable fillSource: (int * int * int * int) option = None
    let mutable fillTarget: (SpreadsheetInput.FillDirection * int) option = None

    let selectionBounds () =
        match model.Selection with
        | SingleCell (row, col) -> row, col, row, col
        | Range (sr, sc, er, ec) -> min sr er, min sc ec, max sr er, max sc ec

    /// The cell an edit applies to: the selected cell, or the top-left of a
    /// selected range.
    let activeCellAddress () =
        let top, left, _, _ = selectionBounds()
        { Row = top; Col = left }

    /// Commit typed text to a cell: parse it the way Excel reads input, check
    /// it against the cell's validation, store it and recalculate, as one
    /// undoable step. False when a validation rule refused the entry.
    let commitInput (addr: CellAddress) (text: string) (validate: bool) (undoable: bool) : bool =
        let oldCellData = getCell addr model
        if System.String.IsNullOrWhiteSpace text then
            // Empty input clears the cell.
            match oldCellData with
            | Some old ->
                model <- clearCell addr model
                recalculateFormulas()
                if undoable then
                    model <- pushCommand { Undo = (fun m -> setCell addr old m)
                                           Redo = (fun m -> clearCell addr m)
                                           Description = sprintf "Clear cell %s" (formatCellRef addr.Row addr.Col) } model
            | None -> ()
            true
        else
            let parsed = SpreadsheetInput.cellFromInput (currentYear()) oldCellData text
            let outcome =
                if validate && not model.Validations.IsEmpty then
                    refreshRuleWorkbook()
                    checkEntryWith ruleEvaluator addr parsed.Value model
                else EntryValid
            let notify style title message =
                let styleName =
                    match style with
                    | StopStyle -> "stop"
                    | WarningStyle -> "warning"
                    | InformationStyle -> "information"
                validationNotice <- Some (styleName, defaultArg title "", message, addr)
                for listener in validationListeners do listener styleName (defaultArg title "") message
            match outcome with
            | EntryInvalid (StopStyle, title, message) ->
                notify StopStyle title message
                false
            | _ ->
                (match outcome with
                 | EntryInvalid (style, title, message) -> notify style title message
                 | EntryValid -> validationNotice <- None)
                let cellData =
                    match parsed.Value with
                    // Store a first result with the formula; the recalculation
                    // that follows settles it and everything that reads it.
                    | CellValue.Formula (formula, _) -> { parsed with Value = CellValue.Formula (formula, evaluateFormula formula model workbook) }
                    | _ -> parsed
                recalculateSingleCellEdit addr cellData
                if undoable then
                    model <- pushCommand { Undo = (fun m -> match oldCellData with
                                                            | Some old -> setCell addr old m
                                                            | None -> clearCell addr m)
                                           Redo = (fun m -> setCell addr cellData m)
                                           Description = sprintf "Edit cell %s" (formatCellRef addr.Row addr.Col) } model
                true

    /// Continue the block `source` in `direction` by `count` rows/columns as
    /// one undoable step. `series` continues series (the fill handle);
    /// without it the block is copied (Ctrl+D / Ctrl+R).
    let fillFrom (source: int * int * int * int) (direction: SpreadsheetInput.FillDirection) (count: int) (series: bool) =
        if count > 0 && not readOnly then
            let targets = SpreadsheetInput.fillTargets source direction count
            let oldCells = targets |> List.map (fun addr -> addr, getCell addr model)
            model <- SpreadsheetInput.fillCells source direction count series model
            recalculateFormulas()
            let newCells = targets |> List.map (fun addr -> addr, getCell addr model)
            let top, left, bottom, right = source
            let selection =
                match direction with
                | SpreadsheetInput.FillDown -> Range (top, left, bottom + count, right)
                | SpreadsheetInput.FillUp -> Range (max 0 (top - count), left, bottom, right)
                | SpreadsheetInput.FillRight -> Range (top, left, bottom, right + count)
                | SpreadsheetInput.FillLeft -> Range (top, max 0 (left - count), bottom, right)
            model <- { model with Selection = selection }
            model <- pushCommand { Undo = restoreCellOptions oldCells
                                   Redo = restoreCellOptions newCells
                                   Description = "Fill" } model

    /// Ctrl+D / Ctrl+R: copy the first row (column) of the selection through
    /// the rest of it; with one cell selected, copy its neighbour into it.
    let fillSelection (down: bool) =
        let top, left, bottom, right = selectionBounds()
        if down then
            if bottom > top then fillFrom (top, left, top, right) SpreadsheetInput.FillDown (bottom - top) false
            elif top > 0 then
                fillFrom (top - 1, left, top - 1, right) SpreadsheetInput.FillDown 1 false
                model <- { model with Selection = (if left = right then SingleCell (top, left) else Range (top, left, top, right)) }
        else
            if right > left then fillFrom (top, left, bottom, left) SpreadsheetInput.FillRight (right - left) false
            elif left > 0 then
                fillFrom (top, left - 1, bottom, left - 1) SpreadsheetInput.FillRight 1 false
                model <- { model with Selection = (if top = bottom then SingleCell (top, left) else Range (top, left, bottom, left)) }

    /// Clear every cell of the selection as one undoable step.
    let clearSelection () =
        let selection = model.Selection
        let oldCells = model.Cells
        model <- clearSelectionCells selection model
        if not (obj.ReferenceEquals(oldCells, model.Cells)) then
            recalculateFormulas()
            model <- pushCommand { Undo = (fun m -> restoreCellMap oldCells m)
                                   Redo = (fun m -> clearSelectionCells selection m)
                                   Description = "Delete" } model

    /// Start editing a cell. `seed` replaces the cell's content in the edit
    /// box — the character typed over the cell — instead of loading it.
    let rec startEditingWith (row: int) (col: int) (seed: string option) =
        if not readOnly then
            let addr = { Row = row; Col = col }
            // Edit the formula itself, not its cached result, and a date as
            // a date: what is shown must give the same cell when committed.
            let currentValue =
                match seed with
                | Some text -> text
                | None -> SpreadsheetInput.editText (getCell addr model)
            editCaretAtEnd <- seed.IsSome
            validationNotice <- None
            (match noticeTimer with Some id -> clearTimeoutId id | None -> ())
            noticeTimer <- None
            model <- {
                model with
                    Selection = SingleCell (row, col)
                    EditingCell = Some addr
                    EditBuffer = currentValue
            }
            render()

    and startEditing (row: int) (col: int) = startEditingWith row col None

    /// Finish editing. With `save` the entry is committed; the result is
    /// false when a validation rule refused it (the cell is unchanged and a
    /// notice says why).
    and finishEditing (save: bool) : bool =
        match model.EditingCell with
        | None -> true
        | Some addr ->
            let buffer = model.EditBuffer
            model <- { model with EditingCell = None; EditBuffer = "" }
            let accepted = if save then commitInput addr buffer true true else true
            scheduleNoticeDismiss()
            render()
            accepted

    /// A validation notice goes away by itself after a few seconds.
    and scheduleNoticeDismiss () =
        (match noticeTimer with Some id -> clearTimeoutId id | None -> ())
        noticeTimer <- None
        if validationNotice.IsSome then
            noticeTimer <-
                Some (setTimeoutMs (fun () ->
                        noticeTimer <- None
                        validationNotice <- None
                        // Never redraw under an entry in progress.
                        if model.EditingCell.IsNone then render()) 6000)

    /// Move selection
    and moveSelection (deltaRow: int) (deltaCol: int) =
        // A range collapses to its active (top-left) cell, then moves.
        let active = activeCellAddress()
        let newRow = (active.Row + deltaRow) |> max 0 |> min (model.RowCount - 1)
        let newCol = (active.Col + deltaCol) |> max 0 |> min (model.ColCount - 1)
        model <- { model with Selection = SingleCell (newRow, newCol) }
        render()

    /// Handle keyboard events
    and handleKeyDown (e: obj) =
        let key = getEventKey e
        let ctrlKey = getProp e "ctrlKey"
        let metaKey = getProp e "metaKey"
        let isModifier = (ctrlKey = box true) || (metaKey = box true)
        let mutatingShortcut = isModifier && (key = "x" || key = "v" || key = "z" || key = "y" || key = "d" || key = "r")
        let mutatingKey = key = "Delete" || key = "Backspace" || key = "Enter" || key = "F2" || (key.Length = 1 && not isModifier)

        if readOnly && (mutatingShortcut || mutatingKey) then
            preventDefault e
        else

          match model.EditingCell with
          | Some addr ->
            // Editing mode
            match key with
            | "Enter" ->
                preventDefault e
                // A refused entry leaves the selection on its cell.
                if finishEditing true then moveSelection 1 0  // Move down
            | "Escape" ->
                preventDefault e
                finishEditing false |> ignore
            | "Tab" ->
                preventDefault e
                if finishEditing true then moveSelection 0 1  // Move right
            | _ -> ()

          | None ->
            // Navigation mode
            if isModifier then
                // Handle Ctrl/Cmd shortcuts
                match key with
                | "c" ->
                    // Copy
                    preventDefault e
                    let clipboardData = serializeSelection model.Selection model false
                    model <- { model with Clipboard = Some clipboardData }

                | "x" ->
                    // Cut
                    preventDefault e
                    cutSelection()
                    render()

                | "v" ->
                    // Paste
                    preventDefault e
                    pasteClipboardAtSelection()
                    render()

                | "z" ->
                    // Undo
                    preventDefault e
                    match undo model with
                    | Some undoneModel ->
                        model <- undoneModel
                        recalculateFormulas()
                        render()
                    | None -> ()

                | "y" ->
                    // Redo
                    preventDefault e
                    match redo model with
                    | Some redoneModel ->
                        model <- redoneModel
                        recalculateFormulas()
                        render()
                    | None -> ()

                | "f" ->
                    // Find/Replace dialog
                    preventDefault e
                    showFindReplaceDialog()

                | "d" ->
                    // Fill down
                    preventDefault e
                    fillSelection true
                    render()

                | "r" ->
                    // Fill right
                    preventDefault e
                    fillSelection false
                    render()

                | _ -> ()
            else
                // Regular navigation keys
                match key with
                | "ArrowUp" ->
                    preventDefault e
                    moveSelection -1 0
                | "ArrowDown" ->
                    preventDefault e
                    moveSelection 1 0
                | "ArrowLeft" ->
                    preventDefault e
                    moveSelection 0 -1
                | "ArrowRight" ->
                    preventDefault e
                    moveSelection 0 1
                | "Enter" | "F2" ->
                    preventDefault e
                    let active = activeCellAddress()
                    startEditing active.Row active.Col
                | "Delete" | "Backspace" ->
                    preventDefault e
                    clearSelection()
                    render()
                | "Tab" ->
                    preventDefault e
                    moveSelection 0 1
                | _ ->
                    // Type-to-edit: a printable character starts an entry
                    // that replaces the cell's content, beginning with it.
                    let altKey = getProp e "altKey"
                    if key.Length = 1 && not (altKey = box true) then
                        preventDefault e
                        let active = activeCellAddress()
                        startEditingWith active.Row active.Col (Some key)

    /// Handle input change (while editing)
    and handleInput (e: obj) =
        let target = getEventTarget e
        let value = getValue target
        model <- { model with EditBuffer = value }

    and renderColumnHeader (col: int): Element =
        let colX = getColumnX col
        let colW = getColumnWidth col model

        let header = createElement "div"
        setClassName header "spreadsheet-col-header"
        setAttribute header "data-col" (string col)
        setStyles header [
            "position", "absolute"
            "top", sprintf "%.0fpx" (if useSharedLayoutGeometry then scrollTop else 0.0)
            "left", sprintf "%.0fpx" colX
            "width", sprintf "%.0fpx" colW
            "height", sprintf "%.0fpx" headerHeight
            "border-right", "1px solid #d0d0d0"
            "border-bottom", "1px solid #a6a6a6"
            "background", "linear-gradient(180deg, #f8f8f8 0%, #e8e8e8 100%)"
            "display", "flex"
            "align-items", "center"
            "justify-content", "space-between"
            "padding", "0 4px"
            "box-sizing", "border-box"
            "font-weight", "normal"
            "font-size", "11px"
            "font-family", "Arial, sans-serif"
            "color", "#444"
            "user-select", "none"
        ]

        // Header text with sort indicator
        let headerTextDiv = createElement "div"
        setStyles headerTextDiv [
            "flex", "1"
            "text-align", "center"
            "cursor", "pointer"
        ]

        let headerText =
            match model.SortState with
            | Some sortState when sortState.Column = col ->
                let arrow = if sortState.Ascending then " ▲" else " ▼"
                (columnToLetter col) + arrow
            | _ ->
                columnToLetter col

        setTextContent headerTextDiv headerText

        // Add click handler for sorting on header text
        addEventListener headerTextDiv "click" (fun _ ->
            invalidateCalcEngine()
            // Toggle sort direction if clicking same column, otherwise ascending
            let ascending =
                match model.SortState with
                | Some sortState when sortState.Column = col -> not sortState.Ascending
                | _ -> true

            sortActiveSheet col ascending (sprintf "Sort by column %s" (columnToLetter col))

            render()
        )

        appendChild header headerTextDiv

        // Filter icon button
        let filterBtn = createElement "div"
        setStyles filterBtn [
            "width", "16px"
            "height", "16px"
            "display", "flex"
            "align-items", "center"
            "justify-content", "center"
            "cursor", "pointer"
            "font-size", "10px"
            "opacity", (if Map.containsKey col model.Filters then "1.0" else "0.5")
            "border-radius", "2px"
        ]
        setTextContent filterBtn "▼"
        setAttribute filterBtn "title" "Filter"

        // Filter button hover effect
        addEventListener filterBtn "mouseenter" (fun _ ->
            setStyles filterBtn [
                "background", "#d0d0d0"
                "opacity", "1.0"
            ]
        )
        addEventListener filterBtn "mouseleave" (fun _ ->
            setStyles filterBtn [
                "background", "transparent"
                "opacity", (if Map.containsKey col model.Filters then "1.0" else "0.5")
            ]
        )

        // Filter button click - show dropdown
        addEventListener filterBtn "click" (fun e ->
            preventDefault e
            stopPropagation e
            showFilterDropdown col (colX + colW) headerHeight
        )

        appendChild header filterBtn

        // Add resize handle to the right edge
        let resizeHandle = createElement "div"
        setClassName resizeHandle "col-resize-handle"
        setStyles resizeHandle [
            "position", "absolute"
            "top", "0"
            "right", "0"
            "width", "4px"
            "height", "100%"
            "cursor", "col-resize"
            "z-index", "10"
        ]

        // Resize handle mousedown
        addEventListener resizeHandle "mousedown" (fun e ->
            preventDefault e
            stopPropagation e  // Prevent sort click from firing
            isResizingColumn <- true
            resizeColumnIndex <- Some col
            resizeStartPos <- getProp e "clientX" |> unbox<float>
            resizeStartSize <- colW
        )

        appendChild header resizeHandle
        header

    /// Render row header
    and renderRowHeader (row: int): Element =
        let rowY = getRowY row
        let rowH = getRowHeight row model

        let header = createElement "div"
        setClassName header "spreadsheet-row-header"
        setAttribute header "data-row" (string row)
        setStyles header [
            "position", "absolute"
            "top", sprintf "%.0fpx" rowY
            "left", sprintf "%.0fpx" (if useSharedLayoutGeometry then scrollLeft else 0.0)
            "width", sprintf "%.0fpx" rowHeaderWidth
            "height", sprintf "%.0fpx" rowH
            "border-right", "1px solid #a6a6a6"
            "border-bottom", "1px solid #d0d0d0"
            "background", "linear-gradient(90deg, #f8f8f8 0%, #e8e8e8 100%)"
            "display", "flex"
            "align-items", "center"
            "justify-content", "center"
            "box-sizing", "border-box"
            "font-weight", "normal"
            "font-size", "11px"
            "font-family", "Arial, sans-serif"
            "color", "#444"
            "user-select", "none"
            "cursor", "default"
        ]
        setTextContent header (string (row + 1))

        // Add resize handle to the bottom edge
        let resizeHandle = createElement "div"
        setClassName resizeHandle "row-resize-handle"
        setStyles resizeHandle [
            "position", "absolute"
            "bottom", "0"
            "left", "0"
            "width", "100%"
            "height", "4px"
            "cursor", "row-resize"
            "z-index", "10"
        ]

        // Resize handle mousedown
        addEventListener resizeHandle "mousedown" (fun e ->
            preventDefault e
            isResizingRow <- true
            resizeRowIndex <- Some row
            resizeStartPos <- getProp e "clientY" |> unbox<float>
            resizeStartSize <- rowH
        )

        appendChild header resizeHandle
        header

    /// Render a cell (returns None for non-primary cells in merged ranges)
    and renderCell (row: int) (col: int): Element option =
        let addr = { Row = row; Col = col }

        // Determine if cell is part of selection and if it's the primary cell
        let isInSelection, isPrimaryCell =
            match model.Selection with
            | SingleCell (r, c) -> (r = row && c = col), true
            | Range (sr, sc, er, ec) ->
                let inRange = row >= sr && row <= er && col >= sc && col <= ec
                let isPrimary = row = sr && col = sc  // Top-left corner is primary
                inRange, isPrimary

        let isEditing = model.EditingCell = Some addr

        // Check if cell is in a merged range
        let mergedCell = findMergedCell addr model
        match mergedCell with
        | Some merge when not (addr.Row = merge.StartRow && addr.Col = merge.StartCol) ->
            // This cell is part of a merge but not the primary cell - don't render it
            None
        | _ ->
            // Either not merged, or this is the primary cell of a merge
            let cell = createElement "div"
            setClassName cell "spreadsheet-cell"
            setAttribute cell "data-row" (string row)
            setAttribute cell "data-col" (string col)

            // Border: green for primary/single cell, light green for range cells
            let borderStyle =
                if isPrimaryCell && isInSelection then "2px solid #217346"
                elif isInSelection then "1px solid #8BC34A"
                else "1px solid #e0e0e0"

            // Background: light green tint for selected range cells
            let bgColor =
                if isInSelection && not isPrimaryCell then "#E8F5E9"
                else "white"

            // Calculate cell position and size (accounting for merges)
            let cellX = getColumnX col
            let cellY = getRowY row
            let cellW, cellH =
                match mergedCell with
                | Some merge ->
                    // Calculate span size
                    let spanW = [merge.StartCol .. merge.EndCol] |> List.sumBy (fun c -> getColumnWidth c model)
                    let spanH = [merge.StartRow .. merge.EndRow] |> List.sumBy (fun r -> getRowHeight r model)
                    spanW, spanH
                | None ->
                    getColumnWidth col model, getRowHeight row model

            // Frozen-pane compensation. The DOM cell sits inside `content`,
            // which scrolls natively under `gridDiv`. To make a frozen cell
            // appear pinned, offset its absolute top/left by the current
            // scroll position so the browser's scroll cancels out. Same
            // trick the corner header at line 1030 uses for itself.
            let inFrozenRow = model.FrozenRows > 0 && row < model.FrozenRows
            let inFrozenCol = model.FrozenCols > 0 && col < model.FrozenCols
            let topOffset = if inFrozenRow then scrollTop else 0.0
            let leftOffset = if inFrozenCol then scrollLeft else 0.0
            // Z-index stacks frozen cells over body cells; frozen corner
            // over both single-axis frozen zones.
            let zIndex =
                match inFrozenRow, inFrozenCol with
                | true, true -> 7    // corner — above both single-axis bands
                | true, false -> 6
                | false, true -> 6
                | false, false -> 1

            setStyles cell [
                "position", "absolute"
                "top", sprintf "%.0fpx" (cellY + topOffset)
                "left", sprintf "%.0fpx" (cellX + leftOffset)
                "width", sprintf "%.0fpx" cellW
                "height", sprintf "%.0fpx" cellH
                "border", borderStyle
                "padding", "2px 4px"
                "overflow", "hidden"
                "white-space", "nowrap"
                "text-overflow", "ellipsis"
                "background", bgColor
                "box-sizing", "border-box"
                "font-family", "Arial, sans-serif"
                "font-size", "11px"
                "line-height", sprintf "%.0fpx" (cellH - 4.0)
                "cursor", "default"
                "z-index", string zIndex
            ]

            if isEditing then
                // Render input
                let input = createElement "input"
                setType input "text"
                setValue input model.EditBuffer
                setStyles input [
                    "width", "100%"
                    "height", "100%"
                    "border", "none"
                    "outline", "none"
                    "padding", "0"
                    "font-family", "Arial, sans-serif"
                    "font-size", "11px"
                    "background", "transparent"
                    "box-sizing", "border-box"
                ]
                addEventListener input "input" handleInput
                appendChild cell input
                // Focused once the cell is in the document (see `render`).
                editInput <- Some input
                Some cell
            else
                // Render value
                let cellData = getCell addr model
                let value = getCellValue addr model
                let displayText =
                    match cellData with
                    // Unbounded width: the cell's CSS clips, so only the
                    // number format is wanted here, not ##### overflow.
                    | Some data -> SpreadsheetCellText.cellDisplayText None 1.0e6 data.Format value
                    | None -> displayValue value
                setTextContent cell displayText

                // Check if cell is in a table
                let tableInfo = getCellTable row col

                // Apply table styling first (base styles)
                match tableInfo with
                | Some (table, isHeader, isEvenRow) ->
                    if isHeader then
                        // Excel-style table header
                        setStyle cell "background" "linear-gradient(180deg, #4472C4 0%, #2E5AA8 100%)"
                        setStyle cell "color" "white"
                        setStyle cell "font-weight" "bold"
                        setStyle cell "text-align" "center"
                        setStyle cell "border" "1px solid #2E5AA8"
                    else
                        // Excel-style alternating rows
                        let bgColor = if isEvenRow then "#FFFFFF" else "#EDF2FA"
                        setStyle cell "background-color" bgColor
                        setStyle cell "border" "1px solid #D0D8E4"
                | None -> ()

                // Apply cell formatting (overrides table styles)
                match cellData with
                | Some data ->
                    if data.Format.Bold then
                        setStyle cell "font-weight" "bold"
                    if data.Format.Italic then
                        setStyle cell "font-style" "italic"
                    if data.Format.Underline then
                        setStyle cell "text-decoration" "underline"
                    match data.Format.TextColor with
                    | Some color -> setStyle cell "color" color
                    | None -> ()
                    match data.Format.BgColor with
                    | Some color -> setStyle cell "background-color" color
                    | None -> ()
                    match data.Format.FontSize with
                    | Some size -> setStyle cell "font-size" (sprintf "%dpx" size)
                    | None -> ()
                    match data.Format.HorizontalAlign with
                    | Some align -> setStyle cell "text-align" align
                    | None -> ()
                    match data.Format.VerticalAlign with
                    | Some "top" -> setStyle cell "vertical-align" "top"
                    | Some "bottom" -> setStyle cell "vertical-align" "bottom"
                    | Some "middle" | Some "center" -> setStyle cell "vertical-align" "middle"
                    | _ -> ()
                    if data.Format.WrapText then
                        setStyle cell "white-space" "normal"
                        setStyle cell "word-wrap" "break-word"
                    // else: default is nowrap (set in base styles)

                    // Apply borders
                    match data.Format.BorderTop with
                    | Some border ->
                        let borderStyle = sprintf "%s %s" border.Style (defaultArg border.Color "#000000")
                        setStyle cell "border-top" borderStyle
                    | None -> ()
                    match data.Format.BorderBottom with
                    | Some border ->
                        let borderStyle = sprintf "%s %s" border.Style (defaultArg border.Color "#000000")
                        setStyle cell "border-bottom" borderStyle
                    | None -> ()
                    match data.Format.BorderLeft with
                    | Some border ->
                        let borderStyle = sprintf "%s %s" border.Style (defaultArg border.Color "#000000")
                        setStyle cell "border-left" borderStyle
                    | None -> ()
                    match data.Format.BorderRight with
                    | Some border ->
                        let borderStyle = sprintf "%s %s" border.Style (defaultArg border.Color "#000000")
                        setStyle cell "border-right" borderStyle
                    | None -> ()
                | None -> ()

                // Apply conditional formatting (highest priority): the merged
                // effect of every matching rule that covers the cell.
                if not model.ConditionalFormats.IsEmpty then
                    match Rbt.Spreadsheet.ConditionalFormatting.effectAt (Rbt.Spreadsheet.ConditionalFormatting.contextFor model) addr with
                    | Some effect ->
                        if effect.Bold = Some true then setStyle cell "font-weight" "bold"
                        if effect.Italic = Some true then setStyle cell "font-style" "italic"
                        let decorations =
                            [ if effect.Underline = Some true then yield "underline"
                              if effect.Strikethrough = Some true then yield "line-through" ]
                        if not decorations.IsEmpty then setStyle cell "text-decoration" (String.concat " " decorations)
                        match effect.TextColor with
                        | Some color -> setStyle cell "color" color
                        | None -> ()
                        match effect.BgColor with
                        | Some color -> setStyle cell "background-color" color
                        | None -> ()
                        match effect.DataBar with
                        | Some (fraction, color) ->
                            // A bar from the left edge, behind the text.
                            let percent = int (System.Math.Round (fraction * 100.0))
                            setStyle cell "background-image" (sprintf "linear-gradient(to right, %s %d%%, transparent %d%%)" color percent percent)
                            setAttribute cell "data-cf-bar" (string percent)
                        | None -> ()
                        match effect.IconSet with
                        | Some (index, count) when count > 1 ->
                            // Same low → high signal the canvas painter draws:
                            // red, amber, green.
                            let t = float index / float (count - 1)
                            let color = if t < 0.34 then "#E53935" elif t < 0.67 then "#F9A825" else "#43A047"
                            let icon = createElement "span"
                            setTextContent icon "\u25CF"
                            setStyles icon [ "color", color; "margin-right", "3px"; "font-size", "10px" ]
                            setAttribute icon "data-cf-icon" (string index)
                            prependChild cell icon
                        | _ -> ()
                    | None -> ()

                // Reapply selection styling (highest priority - overrides all other backgrounds)
                if isInSelection && not isPrimaryCell then
                    setStyle cell "background-color" "#E8F5E9"
                if isInSelection then
                    let borderStyle =
                        if isPrimaryCell then "2px solid #217346"
                        else "1px solid #8BC34A"
                    setStyle cell "border" borderStyle

                // Add comment indicator if cell has a comment or a comment thread
                let thread = if Map.isEmpty model.Threads then None else Map.tryFind { Row = row; Col = col } model.Threads
                match cellData with
                | _ when thread.IsSome ->
                    let indicator = createElement "div"
                    setStyles indicator [
                        "position", "absolute"
                        "top", "0"
                        "right", "0"
                        "width", "0"
                        "height", "0"
                        "border-style", "solid"
                        "border-width", "0 6px 6px 0"
                        "border-color", "transparent #7b61ff transparent transparent"
                        "pointer-events", "none"
                    ]
                    appendChild cell indicator
                    setAttribute cell "title" (threadText thread.Value)
                | Some data when data.Comment.IsSome ->
                    let indicator = createElement "div"
                    setStyles indicator [
                        "position", "absolute"
                        "top", "0"
                        "right", "0"
                        "width", "0"
                        "height", "0"
                        "border-style", "solid"
                        "border-width", "0 6px 6px 0"
                        "border-color", "transparent #ff0000 transparent transparent"
                        "pointer-events", "none"
                    ]
                    appendChild cell indicator

                    // Add tooltip on hover
                    let comment = data.Comment.Value
                    let tooltipText =
                        match comment.Author with
                        | Some author -> sprintf "%s:\n%s" author comment.Text
                        | None -> comment.Text
                    setAttribute cell "title" tooltipText
                | _ -> ()

                // Mouse down to start drag selection
                addEventListener cell "mousedown" (fun e ->
                    preventDefault e  // Prevent text selection
                    match hitTestEvent e with
                    | SpreadsheetLayout.Cell hitAddr ->
                        // Clicking another cell commits the entry in progress.
                        if model.EditingCell.IsSome then finishEditing true |> ignore
                        isDragging <- true
                        dragStartCell <- Some hitAddr
                        selectCellAddress hitAddr
                        render()
                    | _ -> ()
                )

                // Mouse enter to expand selection while dragging
                addEventListener cell "mouseenter" (fun e ->
                    if isFilling then
                        match fillSource, hitTestEvent e with
                        | Some source, SpreadsheetLayout.Cell hitAddr ->
                            let target = SpreadsheetInput.fillTowards source hitAddr.Row hitAddr.Col
                            if target <> fillTarget then
                                fillTarget <- target
                                // Preview: the selection grows over the cells
                                // the release will fill.
                                let top, left, bottom, right = source
                                let preview =
                                    match target with
                                    | Some (SpreadsheetInput.FillDown, n) -> Range (top, left, bottom + n, right)
                                    | Some (SpreadsheetInput.FillUp, n) -> Range (top - n, left, bottom, right)
                                    | Some (SpreadsheetInput.FillRight, n) -> Range (top, left, bottom, right + n)
                                    | Some (SpreadsheetInput.FillLeft, n) -> Range (top, left - n, bottom, right)
                                    | None -> if top = bottom && left = right then SingleCell (top, left) else Range (top, left, bottom, right)
                                model <- { model with Selection = preview }
                                render()
                        | _ -> ()
                    elif isDragging then
                        match dragStartCell with
                        | Some start ->
                            match hitTestEvent e with
                            | SpreadsheetLayout.Cell hitAddr when start.Row <> hitAddr.Row || start.Col <> hitAddr.Col ->
                                let startRow = min start.Row hitAddr.Row
                                let endRow = max start.Row hitAddr.Row
                                let startCol = min start.Col hitAddr.Col
                                let endCol = max start.Col hitAddr.Col
                                model <- { model with Selection = Range (startRow, startCol, endRow, endCol) }
                                render()
                            | _ -> ()
                        | None -> ()
                )

                addEventListener cell "dblclick" (fun e ->
                    match hitTestEvent e with
                    | SpreadsheetLayout.Cell hitAddr -> startEditing hitAddr.Row hitAddr.Col
                    | _ -> ()
                )

                Some cell

    and renderCornerHeader (): Element =
        // Render corner cell (top-left where headers meet)
        let corner = createElement "div"
        setClassName corner "spreadsheet-corner"
        setStyles corner [
            "position", "absolute"
            "top", sprintf "%.0fpx" (if useSharedLayoutGeometry then scrollTop else 0.0)
            "left", sprintf "%.0fpx" (if useSharedLayoutGeometry then scrollLeft else 0.0)
            "width", sprintf "%.0fpx" rowHeaderWidth
            "height", sprintf "%.0fpx" headerHeight
            "border-right", "1px solid #a6a6a6"
            "border-bottom", "1px solid #a6a6a6"
            "background", "linear-gradient(135deg, #f8f8f8 0%, #e0e0e0 100%)"
            "z-index", "10"
        ]
        corner

    /// Render the grid
    and render () =
        match contentElement with
        | None -> ()
        | Some content ->
            // The edit box (and with it the keyboard focus) is rebuilt below;
            // when the grid had the focus it keeps it.
            let hadFocus = holdsFocus content
            // Clear content
            setInnerHTML content ""
            editInput <- None
            cellPool.Clear()
            chromeNodes.Clear()
            if not model.ConditionalFormats.IsEmpty then refreshRuleWorkbook()

            // Get visible range
            let startRow, endRow, startCol, endCol = calculateVisibleRange()

            let appendChrome (node: Element) =
                appendChild content node
                chromeNodes.Add node

            let appendBodyCell (row: int) (col: int) (cell: Element) =
                appendChild content cell
                if isFrozenBodyCell row col then
                    chromeNodes.Add cell
                else
                    cellPool.[cellKey row col] <- cell

            if useSharedLayoutGeometry then
                for headerGeometry in SpreadsheetLayout.headers model (currentViewport()) do
                    match headerGeometry with
                    | SpreadsheetLayout.CornerHeader _ -> appendChrome (renderCornerHeader())
                    | SpreadsheetLayout.ColumnHeader (col, _) -> appendChrome (renderColumnHeader col)
                    | SpreadsheetLayout.RowHeader (row, _) -> appendChrome (renderRowHeader row)
            else
                appendChrome (renderCornerHeader())

                // Render column headers
                for col = startCol to endCol - 1 do
                    appendChrome (renderColumnHeader col)

                // Render row headers
                for row = startRow to endRow - 1 do
                    appendChrome (renderRowHeader row)

            // Render cells from the shared visibility model so partially visible
            // merged cells are emitted even when their primary cell starts offscreen.
            if useSharedLayoutGeometry then
                for cellGeometry in SpreadsheetLayout.visibleCells model (currentViewport()) do
                    let addr = cellGeometry.Address
                    match renderCell addr.Row addr.Col with
                    | Some cell -> appendBodyCell addr.Row addr.Col cell
                    | None -> ()
            else
                for row = startRow to endRow - 1 do
                    for col = startCol to endCol - 1 do
                        match renderCell row col with
                        | Some cell -> appendBodyCell row col cell
                        | None -> ()  // Skip non-primary merged cells

            renderDrawings content
            renderOverlays content
            match editInput with
            | Some input when model.EditingCell.IsSome ->
                // An element takes the focus only once it is attached.
                focus input
                if editCaretAtEnd then caretToEnd input else select input
            | _ -> if hadFocus then focusQuietly content
            lastRenderedModel <- Some model
            notifyChanges ()

    /// Tell OnChange / OnSelectionChange subscribers what changed since they
    /// were last told. Every change to the workbook ends in a render, so this
    /// runs from there; the snapshot is updated before the callbacks run, so a
    /// callback that causes another render is not told twice.
    and notifyChanges () =
        if changeListeners.Count > 0 || selectionListeners.Count > 0 then
            let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
            let structureChanged = current.Length <> notifiedSheets.Length || sheetNames <> notifiedNames
            let changedSheets =
                [| for i in 0 .. current.Length - 1 do
                    if i >= notifiedSheets.Length || not (sameSheetContent current.[i] notifiedSheets.[i]) then yield i |]
            let selection = Some (activeSheetIndex, model.Selection)
            let selectionChanged = selection <> notifiedSelection
            notifiedSheets <- current
            notifiedNames <- Array.copy sheetNames
            notifiedSelection <- selection
            if structureChanged || changedSheets.Length > 0 then
                let change =
                    createObj [
                        "Sheets" ==> changedSheets
                        "StructureChanged" ==> structureChanged
                        "ActiveSheetIndex" ==> activeSheetIndex
                    ]
                for listener in changeListeners.ToArray() do
                    if changeListeners.Contains listener then (try listener change with _ -> ())
            if selectionChanged then
                let change =
                    createObj [
                        "SheetIndex" ==> activeSheetIndex
                        "Selection" ==> selectionRefText ()
                    ]
                for listener in selectionListeners.ToArray() do
                    if selectionListeners.Contains listener then (try listener change with _ -> ())

    /// The selection as an A1 reference ("B2" or "B2:D5").
    and selectionRefText () =
        match model.Selection with
        | SingleCell (row, col) -> sprintf "%s%d" (columnToLetter col) (row + 1)
        | Range (startRow, startCol, endRow, endCol) ->
            sprintf "%s%d:%s%d" (columnToLetter startCol) (startRow + 1) (columnToLetter endCol) (endRow + 1)

    /// Pictures, and labelled frames for charts and other kept objects, over
    /// the cells at their anchors. They are part of the scrolled content, so
    /// they move with the sheet; clicks go through to the cells under them.
    and renderDrawings (content: Element) =
        for drawing in model.Drawings do
            let rect = SpreadsheetLayout.drawingRect model drawing
            if rect.Width >= 1.0 && rect.Height >= 1.0 then
                let label = SpreadsheetLayout.drawingLabel drawing
                let node =
                    match drawing.Content with
                    | Picture (contentType, _, data) ->
                        let image = createElement "img"
                        setAttribute image "src" (pictureUrl data contentType)
                        setAttribute image "alt" (drawing.Description |> Option.defaultValue label)
                        setAttribute image "draggable" "false"
                        setClassName image "spreadsheet-drawing spreadsheet-picture"
                        image
                    | KeptObject (kind, _) ->
                        let frame = createElement "div"
                        setClassName frame ("spreadsheet-drawing spreadsheet-object spreadsheet-object-" + kind)
                        setTextContent frame label
                        setStyles frame [
                            "display", "flex"
                            "align-items", "center"
                            "justify-content", "center"
                            "box-sizing", "border-box"
                            "background", "#f6f8fa"
                            "border", "1px solid #8a94a6"
                            "color", "#4a5568"
                            "font", "12px Arial, sans-serif"
                            "overflow", "hidden"
                        ]
                        frame
                setAttribute node "title" label
                setStyles node [
                    "position", "absolute"
                    "left", sprintf "%.1fpx" (rowHeaderWidth + rect.X)
                    "top", sprintf "%.1fpx" (headerHeight + rect.Y)
                    "width", sprintf "%.1fpx" rect.Width
                    "height", sprintf "%.1fpx" rect.Height
                    "z-index", "3"
                    "pointer-events", "none"
                ]
                appendChild content node
                chromeNodes.Add node

    /// What sits on top of the cells: the selection's fill handle, the input
    /// message of the selected cell's validation, and the validation notice.
    and renderOverlays (content: Element) =
        let top, left, bottom, right = selectionBounds()
        let cornerX = getColumnX right + getColumnWidth right model
        let cornerY = getRowY bottom + getRowHeight bottom model
        if not readOnly && model.EditingCell.IsNone then
            let handle = createElement "div"
            setClassName handle "spreadsheet-fill-handle"
            setStyles handle [
                "position", "absolute"
                "left", sprintf "%.0fpx" (cornerX - 4.0)
                "top", sprintf "%.0fpx" (cornerY - 4.0)
                "width", "7px"
                "height", "7px"
                "background", "#217346"
                "border", "1px solid white"
                "box-sizing", "border-box"
                "cursor", "crosshair"
                "z-index", "9"
            ]
            addEventListener handle "mousedown" (fun e ->
                preventDefault e
                stopPropagation e
                isFilling <- true
                fillSource <- Some (selectionBounds())
                fillTarget <- None)
            appendChild content handle
            chromeNodes.Add handle
        let note (className: string) (title: string) (message: string) (x: float) (y: float) (background: string) (border: string) =
            let box = createElement "div"
            setClassName box className
            setStyles box [
                "position", "absolute"
                "left", sprintf "%.0fpx" x
                "top", sprintf "%.0fpx" y
                "max-width", "260px"
                "padding", "4px 8px"
                "background", background
                "border", "1px solid " + border
                "box-shadow", "0 2px 6px rgba(0,0,0,0.2)"
                "font-family", "Arial, sans-serif"
                "font-size", "11px"
                "white-space", "pre-wrap"
                "pointer-events", "none"
                "z-index", "12"
            ]
            if title <> "" then
                let heading = createElement "div"
                setTextContent heading title
                setStyle heading "font-weight" "bold"
                appendChild box heading
            let body = createElement "div"
            setTextContent body message
            appendChild box body
            appendChild content box
            chromeNodes.Add box
        match validationNotice with
        | Some (style, title, message, cell) ->
            let left, cornerY = cell.Col, getRowY cell.Row + getRowHeight cell.Row model
            let background, border =
                match style with
                | "stop" -> "#FDECEA", "#D93025"
                | "warning" -> "#FFF4E5", "#F29900"
                | _ -> "#E8F0FE", "#1A73E8"
            setAttributeOnNote (fun box -> setAttribute box "data-style" style) (fun () ->
                note "spreadsheet-validation-notice" title message (getColumnX left) (cornerY + 3.0) background border)
        | None ->
            // Excel's input message: shown while the cell is selected.
            match model.Selection with
            | SingleCell (row, col) ->
                match getValidation { Row = row; Col = col } model with
                | Some dv when dv.InputMessage.IsSome || dv.InputTitle.IsSome ->
                    note "spreadsheet-input-message" (defaultArg dv.InputTitle "") (defaultArg dv.InputMessage "")
                         (getColumnX left) (cornerY + 3.0) "#FFFFE1" "#B0B0B0"
                | _ -> ()
            | _ -> ()

    and setAttributeOnNote (tag: Element -> unit) (build: unit -> unit) =
        let before = chromeNodes.Count
        build()
        if chromeNodes.Count > before then tag chromeNodes.[chromeNodes.Count - 1]

    and renderScroll () =
        match contentElement with
        | None -> ()
        | Some content ->
            match lastRenderedModel with
            | Some renderedModel when obj.ReferenceEquals(renderedModel, model) ->
                let desiredKeys = System.Collections.Generic.HashSet<string>()
                let desiredCells = ResizeArray<int * int>()

                let addDesiredCell row col =
                    if isRenderableCell row col then
                        let key = cellKey row col
                        if desiredKeys.Add key then
                            desiredCells.Add(row, col)

                let appendChrome (node: Element) =
                    appendChild content node
                    chromeNodes.Add node

                if useSharedLayoutGeometry then
                    for cellGeometry in SpreadsheetLayout.visibleCells model (currentViewport()) do
                        let addr = cellGeometry.Address
                        addDesiredCell addr.Row addr.Col
                else
                    let startRow, endRow, startCol, endCol = calculateVisibleRange()
                    for row = startRow to endRow - 1 do
                        for col = startCol to endCol - 1 do
                            addDesiredCell row col

                // Non-frozen body cells are scroll-independent absolute nodes.
                let staleKeys = ResizeArray<string>()
                for kvp in cellPool do
                    if not (desiredKeys.Contains kvp.Key) then
                        staleKeys.Add kvp.Key
                for key in staleKeys do
                    removeChild content cellPool.[key]
                    cellPool.Remove key |> ignore

                // Chrome is pinned by scroll offsets, so rebuild it every frame.
                for node in chromeNodes do
                    removeChild content node
                chromeNodes.Clear()

                if useSharedLayoutGeometry then
                    for headerGeometry in SpreadsheetLayout.headers model (currentViewport()) do
                        match headerGeometry with
                        | SpreadsheetLayout.CornerHeader _ -> appendChrome (renderCornerHeader())
                        | SpreadsheetLayout.ColumnHeader (col, _) -> appendChrome (renderColumnHeader col)
                        | SpreadsheetLayout.RowHeader (row, _) -> appendChrome (renderRowHeader row)
                else
                    let startRow, endRow, startCol, endCol = calculateVisibleRange()
                    appendChrome (renderCornerHeader())
                    for col = startCol to endCol - 1 do
                        appendChrome (renderColumnHeader col)
                    for row = startRow to endRow - 1 do
                        appendChrome (renderRowHeader row)

                for row, col in desiredCells do
                    let key = cellKey row col
                    if isFrozenBodyCell row col then
                        match cellPool.TryGetValue key with
                        | true, pooled ->
                            removeChild content pooled
                            cellPool.Remove key |> ignore
                        | false, _ -> ()
                        match renderCell row col with
                        | Some cell ->
                            appendChild content cell
                            chromeNodes.Add cell
                        | None -> ()
                    else
                        if not (cellPool.ContainsKey key) then
                            match renderCell row col with
                            | Some cell ->
                                appendChild content cell
                                cellPool.[key] <- cell
                            | None -> ()

                lastRenderedModel <- Some model
            | _ ->
                render()

    /// Show filter dropdown for column
    and showFilterDropdown (col: int) (x: float) (y: float) =
        // Remove any existing filter dropdown
        let existingDropdown = getElementById "filter-dropdown"
        if not (isNull existingDropdown) then
            removeChild (getBody()) existingDropdown

        // Get unique values in column
        let uniqueValues = getColumnUniqueValues col model

        // Get currently visible values (or all if no filter)
        let visibleValues =
            match Map.tryFind col model.Filters with
            | Some filter -> filter.VisibleValues
            | None -> Set.ofList uniqueValues

        // Create dropdown container
        let dropdown = createElement "div"
        setAttribute dropdown "id" "filter-dropdown"
        setStyles dropdown [
            "position", "absolute"
            "left", sprintf "%.0fpx" x
            "top", sprintf "%.0fpx" y
            "background", "white"
            "border", "1px solid #999"
            "border-radius", "4px"
            "box-shadow", "0 4px 8px rgba(0,0,0,0.2)"
            "padding", "8px"
            "z-index", "1000"
            "min-width", "180px"
            "max-height", "400px"
            "overflow-y", "auto"
            "font-family", "Arial, sans-serif"
            "font-size", "12px"
        ]

        // Title
        let title = createElement "div"
        setStyles title [
            "font-weight", "bold"
            "margin-bottom", "8px"
            "padding-bottom", "4px"
            "border-bottom", "1px solid #ddd"
        ]
        setTextContent title (sprintf "Filter Column %s" (columnToLetter col))
        appendChild dropdown title

        // Select All / Clear All buttons
        let buttonRow = createElement "div"
        setStyles buttonRow [
            "display", "flex"
            "gap", "8px"
            "margin-bottom", "8px"
        ]

        let selectAllBtn = createElement "button"
        setTextContent selectAllBtn "Select All"
        setStyles selectAllBtn [
            "flex", "1"
            "padding", "4px 8px"
            "font-size", "11px"
            "cursor", "pointer"
        ]
        addEventListener selectAllBtn "click" (fun _ ->
            // Check all checkboxes
            let checkboxes = querySelectorAll dropdown "input[type=checkbox]"
            let length: int = getProp checkboxes "length" |> unbox
            for i in 0 .. length - 1 do
                let checkbox: Element = emitJsExpr (checkboxes, i) "$0[$1]"
                setProp checkbox "checked" (box true)
        )

        let clearAllBtn = createElement "button"
        setTextContent clearAllBtn "Clear All"
        setStyles clearAllBtn [
            "flex", "1"
            "padding", "4px 8px"
            "font-size", "11px"
            "cursor", "pointer"
        ]
        addEventListener clearAllBtn "click" (fun _ ->
            // Uncheck all checkboxes
            let checkboxes = querySelectorAll dropdown "input[type=checkbox]"
            let length: int = getProp checkboxes "length" |> unbox
            for i in 0 .. length - 1 do
                let checkbox: Element = emitJsExpr (checkboxes, i) "$0[$1]"
                setProp checkbox "checked" (box false)
        )

        appendChild buttonRow selectAllBtn
        appendChild buttonRow clearAllBtn
        appendChild dropdown buttonRow

        // Value checkboxes
        let valueList = createElement "div"
        setStyles valueList [
            "max-height", "250px"
            "overflow-y", "auto"
            "margin-bottom", "8px"
        ]

        for value in uniqueValues do
            let valueRow = createElement "label"
            setStyles valueRow [
                "display", "flex"
                "align-items", "center"
                "padding", "4px"
                "cursor", "pointer"
                "user-select", "none"
            ]
            addEventListener valueRow "mouseenter" (fun _ ->
                setStyles valueRow ["background", "#f0f0f0"]
            )
            addEventListener valueRow "mouseleave" (fun _ ->
                setStyles valueRow ["background", "transparent"]
            )

            let checkbox = createElement "input"
            setAttribute checkbox "type" "checkbox"
            setProp checkbox "checked" (box (Set.contains value visibleValues))
            setStyles checkbox [
                "margin-right", "6px"
            ]

            let label = createElement "span"
            setTextContent label (if System.String.IsNullOrWhiteSpace(value) then "(Blank)" else value)

            appendChild valueRow checkbox
            appendChild valueRow label
            appendChild valueList valueRow

        appendChild dropdown valueList

        // Action buttons
        let actionRow = createElement "div"
        setStyles actionRow [
            "display", "flex"
            "gap", "8px"
            "padding-top", "8px"
            "border-top", "1px solid #ddd"
        ]

        let applyBtn = createElement "button"
        setTextContent applyBtn "Apply"
        setStyles applyBtn [
            "flex", "1"
            "padding", "6px 12px"
            "background", "#0066cc"
            "color", "white"
            "border", "none"
            "border-radius", "3px"
            "cursor", "pointer"
            "font-size", "12px"
        ]
        addEventListener applyBtn "click" (fun _ ->
            // Collect checked values
            let checkboxes = querySelectorAll dropdown "input[type=checkbox]"
            let length: int = getProp checkboxes "length" |> unbox
            let mutable checkedValues = Set.empty

            for i in 0 .. length - 1 do
                let checkbox: Element = emitJsExpr (checkboxes, i) "$0[$1]"
                let isChecked = getProp checkbox "checked" |> unbox<bool>
                if isChecked then
                    let value = uniqueValues.[i]
                    checkedValues <- Set.add value checkedValues

            // Apply filter
            if checkedValues.Count = uniqueValues.Length then
                // All selected = no filter
                model <- removeFilter col model
            else
                model <- applyFilter col checkedValues model

            // Remove dropdown
            removeChild (getBody()) dropdown
            render()
        )

        let cancelBtn = createElement "button"
        setTextContent cancelBtn "Cancel"
        setStyles cancelBtn [
            "flex", "1"
            "padding", "6px 12px"
            "background", "#666"
            "color", "white"
            "border", "none"
            "border-radius", "3px"
            "cursor", "pointer"
            "font-size", "12px"
        ]
        addEventListener cancelBtn "click" (fun _ ->
            // Remove dropdown
            removeChild (getBody()) dropdown
        )

        appendChild actionRow applyBtn
        appendChild actionRow cancelBtn
        appendChild dropdown actionRow

        // Add to body
        appendChild (getBody()) dropdown

        // Click outside to close
        let rec closeDropdown = fun (e: obj) ->
            let target = getEventTarget e
            if not (contains dropdown target) then
                removeChild (getBody()) dropdown
                removeEventListener (getDocument()) "click" closeDropdown

        // Delay adding the click listener to avoid immediate closure
        JS.setTimeout (fun () ->
            if not disposed then addEventListener (getDocument()) "click" closeDropdown
        ) 100 |> ignore
        cleanups.Add(fun () ->
            removeEventListener (getDocument()) "click" closeDropdown
            detach dropdown)


    /// Show find/replace dialog
    and showFindReplaceDialog () =
        // Remove any existing dialog
        let existingDialog = getElementById "find-replace-dialog"
        if not (isNull existingDialog) then
            removeChild (getBody()) existingDialog

        // Create modal overlay
        let overlay = createElement "div"
        setAttribute overlay "id" "find-replace-dialog"
        setStyles overlay [
            "position", "fixed"
            "top", "0"
            "left", "0"
            "width", "100%"
            "height", "100%"
            "background", "rgba(0,0,0,0.5)"
            "display", "flex"
            "align-items", "center"
            "justify-content", "center"
            "z-index", "2000"
        ]

        // Create dialog box
        let dialog = createElement "div"
        setStyles dialog [
            "background", "white"
            "border-radius", "8px"
            "padding", "20px"
            "box-shadow", "0 4px 16px rgba(0,0,0,0.3)"
            "width", "400px"
            "font-family", "Arial, sans-serif"
        ]

        // Title
        let title = createElement "h3"
        setTextContent title "Find and Replace"
        setStyles title [
            "margin", "0 0 16px 0"
            "font-size", "16px"
            "font-weight", "bold"
        ]
        appendChild dialog title

        // Find input
        let findLabel = createElement "label"
        setTextContent findLabel "Find:"
        setStyles findLabel [
            "display", "block"
            "margin-bottom", "4px"
            "font-size", "12px"
        ]
        appendChild dialog findLabel

        let findInput = createElement "input"
        setAttribute findInput "type" "text"
        setAttribute findInput "placeholder" "Enter text to find..."
        setStyles findInput [
            "width", "100%"
            "padding", "8px"
            "margin-bottom", "12px"
            "border", "1px solid #ccc"
            "border-radius", "4px"
            "font-size", "14px"
            "box-sizing", "border-box"
        ]
        appendChild dialog findInput

        // Replace input
        let replaceLabel = createElement "label"
        setTextContent replaceLabel "Replace with:"
        setStyles replaceLabel [
            "display", "block"
            "margin-bottom", "4px"
            "font-size", "12px"
        ]
        appendChild dialog replaceLabel

        let replaceInput = createElement "input"
        setAttribute replaceInput "type" "text"
        setAttribute replaceInput "placeholder" "Enter replacement text..."
        setStyles replaceInput [
            "width", "100%"
            "padding", "8px"
            "margin-bottom", "12px"
            "border", "1px solid #ccc"
            "border-radius", "4px"
            "font-size", "14px"
            "box-sizing", "border-box"
        ]
        appendChild dialog replaceInput

        // Options
        let optionsRow = createElement "div"
        setStyles optionsRow [
            "margin-bottom", "16px"
        ]

        let caseSensitiveCheckbox = createElement "input"
        setAttribute caseSensitiveCheckbox "type" "checkbox"
        setAttribute caseSensitiveCheckbox "id" "case-sensitive"
        setStyles caseSensitiveCheckbox [
            "margin-right", "6px"
        ]

        let caseSensitiveLabel = createElement "label"
        setAttribute caseSensitiveLabel "for" "case-sensitive"
        setTextContent caseSensitiveLabel "Match case"
        setStyles caseSensitiveLabel [
            "font-size", "12px"
            "cursor", "pointer"
        ]

        appendChild optionsRow caseSensitiveCheckbox
        appendChild optionsRow caseSensitiveLabel
        appendChild dialog optionsRow

        // Results display
        let resultsDiv = createElement "div"
        setAttribute resultsDiv "id" "search-results"
        setStyles resultsDiv [
            "margin-bottom", "16px"
            "font-size", "12px"
            "color", "#666"
            "min-height", "20px"
        ]
        appendChild dialog resultsDiv

        // Button row
        let buttonRow = createElement "div"
        setStyles buttonRow [
            "display", "flex"
            "gap", "8px"
        ]

        // Find Next button
        let findNextBtn = createElement "button"
        setTextContent findNextBtn "Find Next"
        setStyles findNextBtn [
            "flex", "1"
            "padding", "8px 12px"
            "background", "#0066cc"
            "color", "white"
            "border", "none"
            "border-radius", "4px"
            "cursor", "pointer"
            "font-size", "13px"
        ]
        addEventListener findNextBtn "click" (fun _ ->
            let query = getValue findInput
            if not (System.String.IsNullOrWhiteSpace(query)) then
                let caseSensitive = getChecked caseSensitiveCheckbox

                // Start search if not already searching
                if model.SearchQuery <> Some query || model.SearchCaseSensitive <> Some caseSensitive then
                    model <- startSearch query caseSensitive model

                // Find next
                match nextSearchResult model with
                | Some newModel ->
                    model <- newModel
                    let current = model.SearchIndex + 1
                    let total = model.SearchResults.Length
                    setTextContent resultsDiv (sprintf "Found %d of %d" current total)
                    render()
                | None ->
                    setTextContent resultsDiv "No results found"
        )
        appendChild buttonRow findNextBtn

        // Replace button
        let replaceBtn = createElement "button"
        setTextContent replaceBtn "Replace"
        setStyles replaceBtn [
            "flex", "1"
            "padding", "8px 12px"
            "background", "#28a745"
            "color", "white"
            "border", "none"
            "border-radius", "4px"
            "cursor", "pointer"
            "font-size", "13px"
        ]
        addEventListener replaceBtn "click" (fun _ ->
            let findText = getValue findInput
            let replaceText = getValue replaceInput
            let caseSensitive = getChecked caseSensitiveCheckbox

            match replaceInSelection findText replaceText caseSensitive with
            | true ->
                setTextContent resultsDiv "Replaced current match"

                // Find next automatically
                match nextSearchResult model with
                | Some nextModel ->
                    model <- nextModel
                    render()
                | None ->
                    render()
                    setTextContent resultsDiv "No more matches"
            | false ->
                setTextContent resultsDiv "No match at current selection"
        )
        appendChild buttonRow replaceBtn

        // Replace All button
        let replaceAllBtn = createElement "button"
        setTextContent replaceAllBtn "Replace All"
        setStyles replaceAllBtn [
            "flex", "1"
            "padding", "8px 12px"
            "background", "#ffc107"
            "color", "#000"
            "border", "none"
            "border-radius", "4px"
            "cursor", "pointer"
            "font-size", "13px"
        ]
        addEventListener replaceAllBtn "click" (fun _ ->
            let findText = getValue findInput
            let replaceText = getValue replaceInput
            let caseSensitive = getChecked caseSensitiveCheckbox

            if not (System.String.IsNullOrWhiteSpace(findText)) then
                let count = replaceEverywhere findText replaceText caseSensitive

                if count > 0 then
                    setTextContent resultsDiv (sprintf "Replaced %d matches" count)
                    render()
                else
                    setTextContent resultsDiv "No matches found"
        )
        appendChild buttonRow replaceAllBtn

        // Close button
        let closeBtn = createElement "button"
        setTextContent closeBtn "Close"
        setStyles closeBtn [
            "flex", "1"
            "padding", "8px 12px"
            "background", "#666"
            "color", "white"
            "border", "none"
            "border-radius", "4px"
            "cursor", "pointer"
            "font-size", "13px"
        ]
        addEventListener closeBtn "click" (fun _ ->
            removeChild (getBody()) overlay
        )
        appendChild buttonRow closeBtn

        appendChild dialog buttonRow
        appendChild overlay dialog
        appendChild (getBody()) overlay

        // Focus find input
        JS.setTimeout (fun () ->
            let _ : obj = emitJsExpr findInput "$0.focus()"
            ()
        ) 100 |> ignore

        // Close on Escape key
        let rec handleEscape = fun (e: obj) ->
            let key = getEventKey e
            if key = "Escape" then
                removeChild (getBody()) overlay
                removeEventListener (getDocument()) "keydown" handleEscape

        addEventListener (getDocument()) "keydown" handleEscape
        cleanups.Add(fun () ->
            removeEventListener (getDocument()) "keydown" handleEscape
            detach overlay)

        ()


    /// Show conditional formatting dialog for current selection
    and showConditionalFormattingDialog () =
        // Remove any existing dialog
        let existingDialog = getElementById "conditional-format-dialog"
        if not (isNull existingDialog) then
            removeChild (getBody()) existingDialog

        // Get current selection
        let targetAddr =
            match model.Selection with
            | SingleCell (row, col) -> Some { Row = row; Col = col }
            | _ -> None

        match targetAddr with
        | None ->
            // No single cell selected, show error
            let _ : obj = emitJsExpr "Please select a single cell to apply conditional formatting" "alert($0)"
            ()
        | Some addr ->
            // Create modal overlay
            let overlay = createElement "div"
            setAttribute overlay "id" "conditional-format-dialog"
            setStyles overlay [
                "position", "fixed"
                "top", "0"
                "left", "0"
                "width", "100%"
                "height", "100%"
                "background", "rgba(0,0,0,0.5)"
                "display", "flex"
                "align-items", "center"
                "justify-content", "center"
                "z-index", "2000"
            ]

            // Create dialog box
            let dialog = createElement "div"
            setStyles dialog [
                "background", "white"
                "border-radius", "8px"
                "padding", "20px"
                "box-shadow", "0 4px 16px rgba(0,0,0,0.3)"
                "width", "450px"
                "font-family", "Arial, sans-serif"
            ]

            // Title
            let title = createElement "h3"
            setTextContent title (sprintf "Conditional Formatting for %s" (formatCellRef addr.Row addr.Col))
            setStyles title [
                "margin", "0 0 16px 0"
                "font-size", "16px"
                "font-weight", "bold"
            ]
            appendChild dialog title

            // Rule type selector
            let ruleLabel = createElement "label"
            setTextContent ruleLabel "Rule Type:"
            setStyles ruleLabel [
                "display", "block"
                "margin-bottom", "4px"
                "font-size", "12px"
                "font-weight", "bold"
            ]
            appendChild dialog ruleLabel

            let ruleSelect = createElement "select"
            setStyles ruleSelect [
                "width", "100%"
                "padding", "8px"
                "margin-bottom", "12px"
                "border", "1px solid #ccc"
                "border-radius", "4px"
                "font-size", "14px"
            ]

            let ruleTypes = [
                ("greater", "Greater Than")
                ("less", "Less Than")
                ("between", "Between")
                ("equal", "Equal To")
                ("contains", "Contains Text")
            ]

            for (value, label) in ruleTypes do
                let option = createElement "option"
                setAttribute option "value" value
                setTextContent option label
                appendChild ruleSelect option

            appendChild dialog ruleSelect

            // Value inputs (will be shown/hidden based on rule type)
            let value1Label = createElement "label"
            setTextContent value1Label "Value:"
            setStyles value1Label [
                "display", "block"
                "margin-bottom", "4px"
                "font-size", "12px"
            ]
            appendChild dialog value1Label

            let value1Input = createElement "input"
            setAttribute value1Input "type" "text"
            setAttribute value1Input "placeholder" "Enter value..."
            setStyles value1Input [
                "width", "100%"
                "padding", "8px"
                "margin-bottom", "12px"
                "border", "1px solid #ccc"
                "border-radius", "4px"
                "font-size", "14px"
                "box-sizing", "border-box"
            ]
            appendChild dialog value1Input

            // Second value input (for "between" rule)
            let value2Label = createElement "label"
            setTextContent value2Label "And:"
            setStyles value2Label [
                "display", "none"
                "margin-bottom", "4px"
                "font-size", "12px"
            ]
            appendChild dialog value2Label

            let value2Input = createElement "input"
            setAttribute value2Input "type" "text"
            setAttribute value2Input "placeholder" "Enter second value..."
            setStyles value2Input [
                "display", "none"
                "width", "100%"
                "padding", "8px"
                "margin-bottom", "12px"
                "border", "1px solid #ccc"
                "border-radius", "4px"
                "font-size", "14px"
                "box-sizing", "border-box"
            ]
            appendChild dialog value2Input

            // Update input visibility based on rule type
            let updateInputVisibility () =
                let ruleType = getValue ruleSelect
                if ruleType = "between" then
                    setStyles value2Label ["display", "block"]
                    setStyles value2Input ["display", "block"]
                else
                    setStyles value2Label ["display", "none"]
                    setStyles value2Input ["display", "none"]

                // Update label text based on type
                if ruleType = "contains" then
                    setTextContent value1Label "Text:"
                else
                    setTextContent value1Label "Value:"

            addEventListener ruleSelect "change" (fun _ -> updateInputVisibility())

            // Format section
            let formatTitle = createElement "h4"
            setTextContent formatTitle "Format to Apply:"
            setStyles formatTitle [
                "margin", "16px 0 8px 0"
                "font-size", "14px"
                "font-weight", "bold"
            ]
            appendChild dialog formatTitle

            // Background color
            let bgColorLabel = createElement "label"
            setTextContent bgColorLabel "Background Color:"
            setStyles bgColorLabel [
                "display", "block"
                "margin-bottom", "4px"
                "font-size", "12px"
            ]
            appendChild dialog bgColorLabel

            let bgColorInput = createElement "input"
            setAttribute bgColorInput "type" "color"
            setAttribute bgColorInput "value" "#ff0000"
            setStyles bgColorInput [
                "width", "100%"
                "height", "40px"
                "padding", "4px"
                "margin-bottom", "12px"
                "border", "1px solid #ccc"
                "border-radius", "4px"
                "cursor", "pointer"
            ]
            appendChild dialog bgColorInput

            // Text color
            let textColorLabel = createElement "label"
            setTextContent textColorLabel "Text Color:"
            setStyles textColorLabel [
                "display", "block"
                "margin-bottom", "4px"
                "font-size", "12px"
            ]
            appendChild dialog textColorLabel

            let textColorInput = createElement "input"
            setAttribute textColorInput "type" "color"
            setAttribute textColorInput "value" "#ffffff"
            setStyles textColorInput [
                "width", "100%"
                "height", "40px"
                "padding", "4px"
                "margin-bottom", "12px"
                "border", "1px solid #ccc"
                "border-radius", "4px"
                "cursor", "pointer"
            ]
            appendChild dialog textColorInput

            // Bold checkbox
            let boldRow = createElement "div"
            setStyles boldRow [
                "margin-bottom", "12px"
            ]

            let boldCheckbox = createElement "input"
            setAttribute boldCheckbox "type" "checkbox"
            setAttribute boldCheckbox "id" "bold-check"
            setStyles boldCheckbox [
                "margin-right", "6px"
            ]

            let boldLabel = createElement "label"
            setAttribute boldLabel "for" "bold-check"
            setTextContent boldLabel "Bold"
            setStyles boldLabel [
                "font-size", "12px"
                "cursor", "pointer"
            ]

            appendChild boldRow boldCheckbox
            appendChild boldRow boldLabel
            appendChild dialog boldRow

            // Button row
            let buttonRow = createElement "div"
            setStyles buttonRow [
                "display", "flex"
                "gap", "8px"
                "margin-top", "16px"
            ]

            // Apply button
            let applyBtn = createElement "button"
            setTextContent applyBtn "Apply"
            setStyles applyBtn [
                "flex", "1"
                "padding", "8px 12px"
                "background", "#0066cc"
                "color", "white"
                "border", "none"
                "border-radius", "4px"
                "cursor", "pointer"
                "font-size", "13px"
            ]
            addEventListener applyBtn "click" (fun _ ->
                let ruleType = getValue ruleSelect
                let value1 = getValue value1Input

                if System.String.IsNullOrWhiteSpace(value1) then
                    let _ : obj = emitJsExpr "Please enter a value" "alert($0)"
                    ()
                else
                    // Parse rule
                    let rule =
                        match ruleType with
                        | "greater" ->
                            match System.Double.TryParse(value1) with
                            | true, n -> Some (GreaterThan n)
                            | _ -> None
                        | "less" ->
                            match System.Double.TryParse(value1) with
                            | true, n -> Some (LessThan n)
                            | _ -> None
                        | "between" ->
                            let value2 = getValue value2Input
                            match System.Double.TryParse(value1), System.Double.TryParse(value2) with
                            | (true, n1), (true, n2) -> Some (Between (n1, n2))
                            | _ -> None
                        | "equal" ->
                            match System.Double.TryParse(value1) with
                            | true, n -> Some (Equal n)
                            | _ -> None
                        | "contains" ->
                            Some (ContainsText value1)
                        | _ -> None

                    match rule with
                    | Some ruleValue ->
                        // Create format
                        let format = {
                            defaultFormat with
                                BgColor = Some (getValue bgColorInput)
                                TextColor = Some (getValue textColorInput)
                                Bold = getChecked boldCheckbox
                        }

                        // Create conditional format
                        let conditionalFormat = {
                            Rule = ruleValue
                            Style = format
                        }

                        // Apply to cell
                        model <- addConditionalFormat addr conditionalFormat model

                        // Close dialog
                        removeChild (getBody()) overlay
                        render()
                    | None ->
                        let _ : obj = emitJsExpr "Invalid value(s) for selected rule type" "alert($0)"
                        ()
            )
            appendChild buttonRow applyBtn

            // Clear button
            let clearBtn = createElement "button"
            setTextContent clearBtn "Clear Rules"
            setStyles clearBtn [
                "flex", "1"
                "padding", "8px 12px"
                "background", "#ffc107"
                "color", "#000"
                "border", "none"
                "border-radius", "4px"
                "cursor", "pointer"
                "font-size", "13px"
            ]
            addEventListener clearBtn "click" (fun _ ->
                model <- removeConditionalFormats addr model
                removeChild (getBody()) overlay
                render()
            )
            appendChild buttonRow clearBtn

            // Close button
            let closeBtn = createElement "button"
            setTextContent closeBtn "Cancel"
            setStyles closeBtn [
                "flex", "1"
                "padding", "8px 12px"
                "background", "#666"
                "color", "white"
                "border", "none"
                "border-radius", "4px"
                "cursor", "pointer"
                "font-size", "13px"
            ]
            addEventListener closeBtn "click" (fun _ ->
                removeChild (getBody()) overlay
            )
            appendChild buttonRow closeBtn

            appendChild dialog buttonRow
            appendChild overlay dialog
            appendChild (getBody()) overlay

            // Close on Escape key
            let rec handleEscape = fun (e: obj) ->
                let key = getEventKey e
                if key = "Escape" then
                    removeChild (getBody()) overlay
                    removeEventListener (getDocument()) "keydown" handleEscape

            addEventListener (getDocument()) "keydown" handleEscape
            cleanups.Add(fun () ->
                removeEventListener (getDocument()) "keydown" handleEscape
                detach overlay)

            ()

    /// Handle scroll
    let handleScroll (e: obj) =
        let target = getCurrentTarget e
        scrollTop <- getScrollTop target
        scrollLeft <- getScrollLeft target
        if not scrollRenderScheduled then
            scrollRenderScheduled <- true
            requestAnimationFrame (fun () ->
                scrollRenderScheduled <- false
                renderScroll()
            )

    /// Mount the spreadsheet
    member _.Mount() =
        // Create container
        let gridDiv = createElement "div"
        setClassName gridDiv "spreadsheet-container"
        // The grid fills its container. A container with no height of its own
        // (height: auto) would grow with the grid's content instead of
        // scrolling it, so the grid then takes a fixed height.
        let fillsContainer = getClientHeight container > 0.0
        setStyles gridDiv [
            "width", "100%"
            "height", (if fillsContainer then "100%" else "600px")
            "min-height", "200px"
            "box-sizing", "border-box"
            "overflow", "auto"
            "position", "relative"
            "border", "1px solid #c0c0c0"
            "background", "white"
        ]

        // Create content (sized for all rows/cols)
        let content = createElement "div"
        setClassName content "spreadsheet-content"

        // Calculate total dimensions based on custom row heights and column widths
        let totalHeight = headerHeight + SpreadsheetLayout.rowTop model model.RowCount
        let totalWidth = rowHeaderWidth + SpreadsheetLayout.columnLeft model model.ColCount

        setStyles content [
            "width", sprintf "%.0fpx" totalWidth
            "height", sprintf "%.0fpx" totalHeight
            "position", "relative"
            "background", "white"
        ]

        // Add event listeners
        addEventListener gridDiv "scroll" handleScroll
        addEventListener content "keydown" handleKeyDown

        // Drag/resize handlers - attach to document to catch events anywhere
        let doc = getDocument()

        // Mouse move for resizing (throttled with requestAnimationFrame)
        let onDocumentMouseMove = (fun (e: obj) ->
            if isResizingColumn then
                match resizeColumnIndex with
                | Some col ->
                    let currentX = getProp e "clientX" |> unbox<float>
                    let delta = currentX - resizeStartPos
                    let newWidth = max 30.0 (resizeStartSize + delta)
                    model <- setColumnWidth col newWidth model

                    // Throttle render with requestAnimationFrame
                    if not resizeRenderScheduled then
                        resizeRenderScheduled <- true
                        requestAnimationFrame (fun () ->
                            render()
                            resizeRenderScheduled <- false
                        )
                | None -> ()
            elif isResizingRow then
                match resizeRowIndex with
                | Some row ->
                    let currentY = getProp e "clientY" |> unbox<float>
                    let delta = currentY - resizeStartPos
                    let newHeight = max 20.0 (resizeStartSize + delta)
                    model <- setRowHeight row newHeight model

                    // Throttle render with requestAnimationFrame
                    if not resizeRenderScheduled then
                        resizeRenderScheduled <- true
                        requestAnimationFrame (fun () ->
                            render()
                            resizeRenderScheduled <- false
                        )
                | None -> ()
        )

        // Mouse up for both drag selection and resize
        let onDocumentMouseUp = (fun (_: obj) ->
            if isFilling then
                isFilling <- false
                match fillSource, fillTarget with
                | Some source, Some (direction, count) ->
                    fillFrom source direction count true
                | _ -> ()
                fillSource <- None
                fillTarget <- None
                render()
            elif isDragging then
                isDragging <- false
                dragStartCell <- None
            elif isResizingColumn then
                isResizingColumn <- false
                resizeColumnIndex <- None
                resizeRenderScheduled <- false
                render()  // Final render to ensure clean state
            elif isResizingRow then
                isResizingRow <- false
                resizeRowIndex <- None
                resizeRenderScheduled <- false
                render()  // Final render to ensure clean state
        )

        addEventListener doc "mousemove" onDocumentMouseMove
        addEventListener doc "mouseup" onDocumentMouseUp
        cleanups.Add(fun () ->
            removeEventListener doc "mousemove" onDocumentMouseMove
            removeEventListener doc "mouseup" onDocumentMouseUp)

        // Focus content to receive keyboard events
        setAttribute content "tabindex" "0"

        appendChild gridDiv content
        appendChild container gridDiv

        gridContainer <- Some gridDiv
        contentElement <- Some content
        disposed <- false

        // Initial render
        containerHeight <- getClientHeight gridDiv
        containerWidth <- getClientWidth gridDiv
        render()

        // Follow the container: the visible range depends on the grid's size.
        let stopObserving =
            observeResize gridDiv (fun () ->
                if not disposed then
                    let height = getClientHeight gridDiv
                    let width = getClientWidth gridDiv
                    if height <> containerHeight || width <> containerWidth then
                        containerHeight <- height
                        containerWidth <- width
                        render())
        cleanups.Add stopObserving

        // Focus for keyboard
        focus content

    /// Take the spreadsheet out of the page: removes every listener it
    /// registered outside its own nodes (document-level mouse handlers, the
    /// resize observer, open dialogs and dropdowns) and every DOM node it
    /// created, and drops all subscriptions. The instance can be mounted again.
    member _.Dispose() =
        disposed <- true
        for cleanup in cleanups.ToArray() do
            try cleanup () with _ -> ()
        cleanups.Clear()
        match noticeTimer with
        | Some id ->
            clearTimeoutId id
            noticeTimer <- None
        | None -> ()
        // Object URLs made for the pictures on display.
        for sheetModel in Array.append [| model |] sheets do
            for drawing in sheetModel.Drawings do
                match drawing.Content with
                | Picture (_, _, data) -> releasePicture data
                | KeptObject _ -> ()
        match gridContainer with
        | Some gridDiv -> detach gridDiv
        | None -> ()
        gridContainer <- None
        contentElement <- None
        editInput <- None
        cellPool.Clear()
        chromeNodes.Clear()
        lastRenderedModel <- None
        changeListeners.Clear()
        selectionListeners.Clear()
        sheetChangedListeners.Clear()

    /// Call `callback` after every change to the workbook's content: cell
    /// values and formats, row / column sizes, merges, inserted or deleted
    /// rows and columns, tables, validations, conditional formats, sheets
    /// added, removed or renamed, undo and redo. It receives
    /// { Sheets: number[] (indices of the sheets that changed),
    ///   StructureChanged: boolean (sheet list changed), ActiveSheetIndex }.
    /// Selection moves and in-cell typing before it is committed are not
    /// document changes. Returns a function that removes the subscription.
    member _.OnChange(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        if changeListeners.Count = 0 && selectionListeners.Count = 0 then
            // Changes are reported from here on, not since the last subscriber left.
            notifiedSheets <- Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
            notifiedNames <- Array.copy sheetNames
            notifiedSelection <- Some (activeSheetIndex, model.Selection)
        changeListeners.Add listener
        System.Action(fun () -> changeListeners.Remove listener |> ignore)

    /// Call `callback` after the selection changed (another cell or range, or
    /// another sheet). It receives { SheetIndex, Selection: "B2" | "B2:D5" }.
    /// Returns a function that removes the subscription.
    member _.OnSelectionChange(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        if changeListeners.Count = 0 && selectionListeners.Count = 0 then
            notifiedSheets <- Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
            notifiedNames <- Array.copy sheetNames
            notifiedSelection <- Some (activeSheetIndex, model.Selection)
        selectionListeners.Add listener
        System.Action(fun () -> selectionListeners.Remove listener |> ignore)

    /// Every sheet's current model, in sheet order (the active sheet's
    /// included, with its latest edits).
    member _.GetSheets() : SpreadsheetModel array =
        Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])

    /// Names of every sheet, in sheet order.
    member _.GetSheetNames() : string array =
        Array.copy sheetNames

    /// 0-based index of the sheet on display.
    member _.GetActiveSheetIndex() : int = activeSheetIndex

    /// Get model (for debugging)
    member _.GetModel() = model

    /// Enable or disable viewer mode. Read-only mode preserves navigation and copy.
    member _.SetReadOnly(value: bool) =
        readOnly <- value
        if readOnly && model.EditingCell.IsSome then
            model <- { model with EditingCell = None; EditBuffer = "" }
            render()

    member _.Render() = render()

    /// Names of every sheet in the workbook in source order.
    member _.SheetNames() : string array =
        Array.copy sheetNames

    /// Whether each sheet (parallel to `SheetNames()`) should get a tab.
    /// False for sheets an imported workbook marks hidden / veryHidden: they
    /// stay loaded because formulas refer to them, but Excel shows no tab.
    member _.SheetVisibility() : bool array =
        Array.init sheets.Length (fun i ->
            match workbook |> Option.bind (fun wb -> Map.tryFind (sheetIdForIndex i) wb.Sheets) with
            | Some sheet -> sheet.Visible
            | None -> true)

    /// 0-based index of the currently displayed sheet.
    member _.ActiveSheetIndex
        with get () : int = activeSheetIndex

    /// Switch the visible sheet. Per-sheet selection / scroll / edit /
    /// filter state persists — when the user comes back to a sheet it's
    /// exactly where they left it, the way Excel does.
    member _.SwitchSheet(index: int) =
        if index >= 0 && index < sheets.Length && index <> activeSheetIndex then
            invalidateCalcEngine()
            // Persist current sheet's mutable state into the array.
            sheets.[activeSheetIndex] <- model
            activeSheetIndex <- index
            model <- sheets.[index]
            // Scroll lives outside the SpreadsheetModel; reset to that
            // sheet's saved scroll origin (we don't persist scroll yet —
            // reset-to-zero matches Excel's behavior on workbook reload).
            scrollTop <- 0.0
            scrollLeft <- 0.0
            render()
            for listener in sheetChangedListeners do
                try listener index with _ -> ()

    /// Register a callback invoked after every successful sheet switch.
    /// Chrome uses this to update its tab-strip selection highlight when
    /// the switch is triggered from inside (e.g. via keyboard shortcut)
    /// rather than from the chrome's own click handler.
    member _.OnSheetChanged(listener: System.Action<int>) =
        sheetChangedListeners.Add(listener.Invoke)

    /// Replace the entire workbook content. Used after loading an XLSX
    /// (or other multi-sheet source) to install all sheets in one shot.
    /// `activeIndex` selects which sheet shows first; out-of-range values
    /// clamp to 0.
    member this.SetWorkbook(newSheets: SpreadsheetModel array, names: string array, activeIndex: int) =
        this.InstallWorkbook(newSheets, names, activeIndex, None)

    /// Replace the workbook content together with its context. `context`
    /// carries what the sheets alone cannot: defined names, declared table
    /// columns and sheet visibility. Without one (plain `SetWorkbook`) a
    /// context is built from the sheets, registering each sheet's tables.
    /// Either way the previous workbook's context is discarded — keeping it
    /// would resolve sheet ids, tables and names against sheets that are gone.
    member _.InstallWorkbook(newSheets: SpreadsheetModel array, names: string array, activeIndex: int, context: Workbook option) =
        if newSheets.Length = 0 then () else
        invalidateCalcEngine()
        // A different workbook: nothing of the file the previous one came from
        // belongs to it (InstallImportedWorkbook / ImportXlsx set the new source).
        xlsxSource <- None
        sheets <- newSheets
        sheetNames <-
            if names.Length = newSheets.Length then Array.copy names
            else Array.init newSheets.Length (fun i -> sprintf "Sheet%d" (i + 1))
        let safeIndex = max 0 (min (sheets.Length - 1) activeIndex)
        activeSheetIndex <- safeIndex
        model <- sheets.[safeIndex]
        workbook <-
            match context with
            | Some wb when List.length wb.SheetOrder = sheets.Length ->
                Some { wb with ActiveSheet = List.item safeIndex wb.SheetOrder }
            | _ ->
                Some (buildWorkbookContext sheets sheetNames (Array.create sheets.Length true) Map.empty Map.empty safeIndex)
        scrollTop <- 0.0
        scrollLeft <- 0.0
        render()
        for listener in sheetChangedListeners do
            try listener safeIndex with _ -> ()

    /// Install a workbook imported from XLSX (`XlsxImporter.importWorkbook`
    /// / `importWorkbookPromise`), remembering the file it came from so that
    /// `ExportXlsx` can carry over what the editor does not model. A negative
    /// `activeIndex` shows the first visible sheet.
    member this.InstallImportedWorkbook(imported: ImportedWorkbook, activeIndex: int) =
        let names = imported.Sheets |> Array.map (fun sheet -> sheet.Name)
        let models = imported.Sheets |> Array.map (fun sheet -> sheet.Model)
        let active =
            if activeIndex >= 0 && activeIndex < models.Length then activeIndex
            else imported.Sheets |> Array.tryFindIndex (fun sheet -> sheet.Visible) |> Option.defaultValue 0
        this.InstallWorkbook(models, names, active, Some imported.Context)
        if models.Length > 0 then xlsxSource <- Some imported.Source

    /// Open an XLSX file: every sheet is loaded and the first visible one
    /// shown. `data` is an ArrayBuffer, Uint8Array or Blob. The promise
    /// rejects when the file cannot be read; the current workbook is then
    /// left untouched.
    member this.ImportXlsx(data: obj): JS.Promise<unit> =
        (loadReaderWorkbook data None).``then``(fun readerWorkbook ->
            this.InstallImportedWorkbook(Rbt.Spreadsheet.XlsxBridge.Import.importWorkbook readerWorkbook, -1))

    /// Whether the workbook came from an XLSX file whose unmodelled content
    /// (drawings, named styles, theme, kept parts) ExportXlsx will carry over.
    member _.HasXlsxSource() : bool = xlsxSource.IsSome

    /// The editor's workbook in the XLSX writer's shape (what ExportXlsx writes).
    member _.ToXlsxWorkbook() : XlsxWriter.XlsxWorkbookOut =
        let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
        Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut
            { Sheets = current
              Names = Array.copy sheetNames
              Visible =
                Array.init sheets.Length (fun i ->
                    match workbook |> Option.bind (fun wb -> Map.tryFind (sheetIdForIndex i) wb.Sheets) with
                    | Some sheet -> sheet.Visible
                    | None -> true)
              ActiveSheetIndex = activeSheetIndex
              DefinedNames = (match workbook with Some wb -> wb.Names | None -> Map.empty)
              Source = xlsxSource
              // A VBA project of the opened file is one of its kept parts.
              Macros = None }

    /// Write the editor's current workbook as XLSX: every sheet with its
    /// values, formulas and their cached values, formats, column widths and
    /// row heights (hidden ones too), merges, freeze panes, sheet visibility,
    /// defined names, tables, conditional formats, validations, comments and
    /// page setup. For a workbook that was opened from a file, everything the
    /// editor does not model - hyperlinks, pictures, charts and other drawing
    /// objects, threaded comments, named cell styles, the theme and other
    /// kept parts - is carried over from that file, and cells, tables,
    /// validations and conditional formats the user did not touch are written
    /// exactly as the file had them. Resolves to the file's bytes (Uint8Array).
    member this.ExportXlsx() : JS.Promise<byte[]> =
        packageBytes (XlsxWriter.writeWorkbookPackage (this.ToXlsxWorkbook()))

    /// Render the current model through the shared portable paint command path.
    member _.RenderPortableCanvas(canvas: Element) = renderPortablePaintToCanvas canvas

    /// How many of the portable paint items are drawing objects (pictures;
    /// the frame fill of a chart or other kept object counts once each).
    member _.PortableDrawingItemCount() =
        SpreadsheetPaint.buildDrawingPaintItems model (currentViewport())
        |> List.filter (function
            | SpreadsheetPaint.DrawImage _ -> true
            | SpreadsheetPaint.FillRect (_, "#f6f8fa") -> true
            | _ -> false)
        |> List.length

    /// The active sheet's pictures and other drawing objects, back to front:
    /// { kind: "picture" | "chart" | "shape" | "other", name, description,
    ///   row, col (anchor cell), x, y, width, height (sheet pixels) }.
    member _.GetDrawings() : obj array =
        model.Drawings
        |> List.map (fun drawing ->
            let rect = SpreadsheetLayout.drawingRect model drawing
            createObj [
                "kind" ==> (match drawing.Content with Picture _ -> "picture" | KeptObject (kind, _) -> kind)
                "name" ==> drawing.Name
                "description" ==> drawing.Description
                "row" ==> drawing.Anchor.FromRow
                "col" ==> drawing.Anchor.FromCol
                "x" ==> rect.X
                "y" ==> rect.Y
                "width" ==> rect.Width
                "height" ==> rect.Height
            ])
        |> List.toArray

    /// The comment thread on a cell of the active sheet: { author, date,
    /// text, replies: [{ author, date, text }] }, or null.
    member _.GetCommentThread(row: int, col: int) : obj =
        match Map.tryFind { Row = row; Col = col } model.Threads with
        | None -> null
        | Some thread ->
            let entry (comment: CommentThread) = [ "author" ==> comment.Author; "date" ==> comment.Date; "text" ==> comment.Text ]
            createObj (entry thread @ [ "replies" ==> (thread.Replies |> List.map (fun reply -> createObj (entry reply)) |> List.toArray) ])

    member _.PortablePaintItemCount() =
        SpreadsheetPaint.buildViewportPaintItems model (currentViewport()) |> List.length

    /// VBA-style per-cell information snapshot for parity comparison
    /// against Excel COM. Mirrors what `extract-excel-layout.vbs` dumps
    /// (one row per cell with address / geometry / format / value / text).
    /// Returned as a flat JS array of objects, keyed by the same column
    /// names as the Excel TSV so the comparison harness can join by
    /// `address` and diff field-for-field.
    member _.GetVbaLayoutCells() =
        // Build a measurer using the loaded text-measurer if any (matches
        // what the Paint pass would use at render time).
        let viewport = currentViewport()
        let cellRows = SpreadsheetLayout.visibleCells model viewport
        // Cell rect's X/Y are in viewport-local CSS px; convert to
        // viewport coordinates and to PDF points (1 pt = 96/72 px) so
        // the comparison can target Excel's pt-based geometry.
        let pxToPt (v: float) = v * 72.0 / 96.0
        cellRows
        |> List.map (fun cell ->
            let addr = cell.Address
            let rect = cell.Rect
            let cellData = getCell addr model
            let value = getCellValue addr model
            let format = cellData |> Option.map (fun c -> c.Format) |> Option.defaultValue defaultFormat
            let fontPt = format.FontSize |> Option.map float |> Option.defaultValue 11.0
            let measure =
                viewport.Measurer
                |> Option.map (fun m -> fun (s: string) -> m s fontPt format.FontName format.Bold format.Italic)
            let displayText = SpreadsheetCellText.cellDisplayText measure rect.Width format value
            let valueRepr =
                match value with
                | CellValue.Empty -> ""
                | CellValue.Number n -> string n
                | CellValue.String s -> s
                | CellValue.Boolean b -> if b then "TRUE" else "FALSE"
                | CellValue.Error e -> "#" + e
                | CellValue.Formula (expr, _) -> "=" + expr
            let formula =
                match value with
                | CellValue.Formula (expr, _) -> "=" + expr
                | _ -> ""
            let addrA1 = columnToLetter addr.Col + string (addr.Row + 1)
            let mergeFor =
                model.MergedCells
                |> List.tryFind (fun m ->
                    addr.Row >= m.StartRow && addr.Row <= m.EndRow &&
                    addr.Col >= m.StartCol && addr.Col <= m.EndCol)
            let mergedRows, mergedCols =
                match mergeFor with
                | Some m -> m.EndRow - m.StartRow + 1, m.EndCol - m.StartCol + 1
                | None -> 1, 1
            createObj [
                "row" ==> (addr.Row + 1)
                "column" ==> (addr.Col + 1)
                "address" ==> addrA1
                "x" ==> pxToPt rect.X
                "y" ==> pxToPt rect.Y
                "width" ==> pxToPt rect.Width
                "height" ==> pxToPt rect.Height
                "mergedRows" ==> mergedRows
                "mergedCols" ==> mergedCols
                "rowHeight" ==> pxToPt (Map.tryFind addr.Row model.RowHeights |> Option.defaultValue 20.0)
                "columnWidth" ==> pxToPt (Map.tryFind addr.Col model.ColumnWidths |> Option.defaultValue 64.0)
                "wrapText" ==> format.WrapText
                "hAlign" ==> (format.HorizontalAlign |> Option.defaultValue "general")
                "vAlign" ==> (format.VerticalAlign |> Option.defaultValue "bottom")
                "fontName" ==> (format.FontName |> Option.defaultValue "Calibri")
                "fontSize" ==> fontPt
                "bold" ==> format.Bold
                "italic" ==> format.Italic
                "numberFormat" ==> (format.NumberFormat |> Option.defaultValue "General")
                "formula" ==> formula
                "value" ==> valueRepr
                "text" ==> displayText
            ])
        |> List.toArray

    /// Snapshot portable visible cell geometry for browser parity checks.
    member _.PortableCellGeometry() =
        SpreadsheetLayout.visibleCells model (currentViewport())
        |> List.map (fun cell ->
            let rect = SpreadsheetLayout.toViewportRect (currentViewport()) SpreadsheetLayout.BodyZone cell.Rect
            createObj [
                "row" ==> cell.Address.Row
                "col" ==> cell.Address.Col
                "x" ==> rect.X
                "y" ==> rect.Y
                "width" ==> rect.Width
                "height" ==> rect.Height
            ])
        |> List.toArray

    /// Snapshot portable header geometry for browser parity checks.
    member _.PortableHeaderGeometry() =
        SpreadsheetLayout.headers model (currentViewport())
        |> List.map (fun header ->
            match header with
            | SpreadsheetLayout.CornerHeader rect ->
                createObj [
                    "type" ==> "corner"
                    "x" ==> rect.X
                    "y" ==> rect.Y
                    "width" ==> rect.Width
                    "height" ==> rect.Height
                ]
            | SpreadsheetLayout.ColumnHeader (col, rect) ->
                let rect = SpreadsheetLayout.toViewportRect (currentViewport()) SpreadsheetLayout.ColumnHeaderZone rect
                createObj [
                    "type" ==> "column"
                    "index" ==> col
                    "x" ==> rect.X
                    "y" ==> rect.Y
                    "width" ==> rect.Width
                    "height" ==> rect.Height
                ]
            | SpreadsheetLayout.RowHeader (row, rect) ->
                let rect = SpreadsheetLayout.toViewportRect (currentViewport()) SpreadsheetLayout.RowHeaderZone rect
                createObj [
                    "type" ==> "row"
                    "index" ==> row
                    "x" ==> rect.X
                    "y" ==> rect.Y
                    "width" ==> rect.Width
                    "height" ==> rect.Height
                ])
        |> List.toArray

    /// Toggle shared layout math for DOM geometry comparisons.
    member _.UseSharedLayoutGeometry(enabled: bool) =
        useSharedLayoutGeometry <- enabled
        render()

    /// Check whether the DOM renderer is using the shared layout math.
    member _.IsUsingSharedLayoutGeometry() =
        useSharedLayoutGeometry

    member _.SelectionDebug() =
        match model.Selection with
        | SingleCell (row, col) ->
            createObj [
                "type" ==> "single"
                "row" ==> row
                "col" ==> col
            ]
        | Range (startRow, startCol, endRow, endCol) ->
            createObj [
                "type" ==> "range"
                "startRow" ==> startRow
                "startCol" ==> startCol
                "endRow" ==> endRow
                "endCol" ==> endCol
            ]

    /// Set a cell from text, read the way typed input is (numbers, dates,
    /// percentages, formulas, `'text`…). A programmatic write: it is not
    /// checked against data validation and adds no undo step.
    member _.SetCell(row: int, col: int, value: string) =
        commitInput { Row = row; Col = col } value false false |> ignore
        render()

    /// Enter text into a cell as the user would: validated, undoable. False
    /// when a validation rule refused the entry.
    member _.EnterText(row: int, col: int, text: string) : bool =
        let accepted = commitInput { Row = row; Col = col } text true true
        scheduleNoticeDismiss()
        render()
        accepted

    /// The validation notice currently shown — `{ style, title, message }`
    /// with style "stop" (entry refused), "warning" or "information" (entry
    /// accepted) — or null.
    member _.GetValidationNotice() : obj =
        match validationNotice with
        | Some (style, title, message, _) -> createObj [ "style" ==> style; "title" ==> title; "message" ==> message ]
        | None -> null

    /// Called with (style, title, message) whenever an entry fails validation.
    member _.OnValidationNotice(listener: System.Action<string, string, string>) =
        validationListeners.Add(fun style title message -> listener.Invoke(style, title, message))

    /// Fill the selection down / right from its first row / column (Ctrl+D /
    /// Ctrl+R).
    member _.FillDown() =
        fillSelection true
        render()

    member _.FillRight() =
        fillSelection false
        render()

    /// Continue the selected block up to the cell (row, col), as dragging its
    /// fill handle there does: series are detected and formulas shifted.
    member _.AutoFill(row: int, col: int) =
        let source = selectionBounds()
        match SpreadsheetInput.fillTowards source row col with
        | Some (direction, count) -> fillFrom source direction count true
        | None -> ()
        render()

    /// Select a range programmatically.
    member _.SelectRange(startRow: int, startCol: int, endRow: int, endCol: int) =
        model <- { model with Selection = (if startRow = endRow && startCol = endCol then SingleCell (startRow, startCol) else Range (min startRow endRow, min startCol endCol, max startRow endRow, max startCol endCol)) }
        render()

    /// Select a specific cell programmatically
    member _.SelectCell(row: int, col: int) =
        model <- { model with Selection = SingleCell (row, col) }
        render()

    /// Apply format to selected cell
    member _.ApplyFormat(formatType: string) =
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            match getCell addr model with
            | Some cellData ->
                let newFormat =
                    match formatType with
                    | "bold" -> { cellData.Format with Bold = not cellData.Format.Bold }
                    | "italic" -> { cellData.Format with Italic = not cellData.Format.Italic }
                    | "underline" -> { cellData.Format with Underline = not cellData.Format.Underline }
                    | "fontIncrease" ->
                        let currentSize = defaultArg cellData.Format.FontSize 11
                        { cellData.Format with FontSize = Some (currentSize + 2) }
                    | "fontDecrease" ->
                        let currentSize = defaultArg cellData.Format.FontSize 11
                        { cellData.Format with FontSize = Some (max 8 (currentSize - 2)) }
                    | "alignLeft" -> { cellData.Format with HorizontalAlign = Some "left" }
                    | "alignCenter" -> { cellData.Format with HorizontalAlign = Some "center" }
                    | "alignRight" -> { cellData.Format with HorizontalAlign = Some "right" }
                    | _ -> cellData.Format

                let newCellData = { cellData with Format = newFormat }
                model <- setCell addr newCellData model
                render()
            | None ->
                // Create cell with format
                let newFormat =
                    match formatType with
                    | "bold" -> { defaultFormat with Bold = true }
                    | "italic" -> { defaultFormat with Italic = true }
                    | "underline" -> { defaultFormat with Underline = true }
                    | "alignLeft" -> { defaultFormat with HorizontalAlign = Some "left" }
                    | "alignCenter" -> { defaultFormat with HorizontalAlign = Some "center" }
                    | "alignRight" -> { defaultFormat with HorizontalAlign = Some "right" }
                    | _ -> defaultFormat

                let cellData = {
                    Value = CellValue.Empty
                    Format = newFormat
                    Comment = None
                }
                model <- setCell addr cellData model
                render()
        | _ -> ()

    /// Apply text color to selected cell
    member _.ApplyTextColor(color: string) =
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            let cellData = getCell addr model |> Option.defaultValue {
                Value = CellValue.Empty
                Format = defaultFormat
                Comment = None
            }
            let newFormat = { cellData.Format with TextColor = Some color }
            let newCellData = { cellData with Format = newFormat }
            model <- setCell addr newCellData model
            render()
        | _ -> ()

    /// Apply background color to selected cell
    member _.ApplyBgColor(color: string) =
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            let cellData = getCell addr model |> Option.defaultValue {
                Value = CellValue.Empty
                Format = defaultFormat
                Comment = None
            }
            let newFormat = { cellData.Format with BgColor = Some color }
            let newCellData = { cellData with Format = newFormat }
            model <- setCell addr newCellData model
            render()
        | _ -> ()

    /// Apply number format to selected cell
    member _.ApplyNumberFormat(format: string) =
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            let cellData = getCell addr model |> Option.defaultValue {
                Value = CellValue.Empty
                Format = defaultFormat
                Comment = None
            }
            let newFormat = { cellData.Format with NumberFormat = Some format }
            let newCellData = { cellData with Format = newFormat }
            model <- setCell addr newCellData model
            render()
        | _ -> ()

    /// Apply border to selected cell
    member _.ApplyBorder(side: string, style: string, color: string) =
        match model.Selection with
        | SingleCell (row, col) ->
            let addr = { Row = row; Col = col }
            let cellData = getCell addr model |> Option.defaultValue {
                Value = CellValue.Empty
                Format = defaultFormat
                Comment = None
            }
            let border = { Style = style; Color = Some color }
            let newFormat =
                match side with
                | "top" -> { cellData.Format with BorderTop = Some border }
                | "bottom" -> { cellData.Format with BorderBottom = Some border }
                | "left" -> { cellData.Format with BorderLeft = Some border }
                | "right" -> { cellData.Format with BorderRight = Some border }
                | "all" ->
                    { cellData.Format with
                        BorderTop = Some border
                        BorderBottom = Some border
                        BorderLeft = Some border
                        BorderRight = Some border }
                | _ -> cellData.Format
            let newCellData = { cellData with Format = newFormat }
            model <- setCell addr newCellData model
            render()
        | _ -> ()

    /// Merge cells in a range
    member _.MergeCells(startRow: int, startCol: int, endRow: int, endCol: int) =
        model <- addMergedCell startRow startCol endRow endCol model
        render()

    /// Unmerge cells at address
    member _.UnmergeCells(row: int, col: int) =
        let addr = { Row = row; Col = col }
        model <- removeMergedCell addr model
        render()

    /// Add or update comment on a cell
    member _.SetComment(row: int, col: int, text: string, author: string option) =
        let addr = { Row = row; Col = col }
        let cellData = getCell addr model |> Option.defaultValue {
            Value = CellValue.Empty
            Format = defaultFormat
            Comment = None
        }
        let comment = { Text = text; Author = author }
        let newCellData = { cellData with Comment = Some comment }
        model <- setCell addr newCellData model
        render()

    /// Remove comment from a cell
    member _.RemoveComment(row: int, col: int) =
        let addr = { Row = row; Col = col }
        match getCell addr model with
        | Some cellData ->
            let newCellData = { cellData with Comment = None }
            model <- setCell addr newCellData model
            render()
        | None -> ()

    /// Set workbook context for table reference support.
    member _.SetWorkbookContext(wb: Workbook) =
        invalidateCalcEngine()
        workbook <- Some wb
        render()  // Re-render to apply table styles

    /// Get workbook
    member _.GetWorkbook() =
        workbook

    /// Define a table in the current workbook
    member _.DefineTable(name: string, startRow: int, startCol: int, endRow: int, endCol: int, columns: string list, hasHeaders: bool) =
        match workbook with
        | Some wb ->
            let table: Rbt.Spreadsheet.Workbook.Table = {
                Name = name
                SheetId = sheetIdForIndex activeSheetIndex  // the sheet being edited
                StartRow = startRow
                StartCol = startCol
                EndRow = endRow
                EndCol = endCol
                Columns = columns
                HasHeaders = hasHeaders
                HasTotals = false
            }
            let newWb = Rbt.Spreadsheet.Workbook.defineTable table wb
            workbook <- Some newWb
            // Structured references to the table resolve from now on.
            recalculateFormulas()
            render()  // Re-render to apply table styles
        | None -> ()

    /// Delete a table
    member _.DeleteTable(name: string) =
        match workbook with
        | Some wb ->
            let newWb = Rbt.Spreadsheet.Workbook.deleteTable name wb
            workbook <- Some newWb
            recalculateFormulas()
            render()
        | None -> ()

    /// Get table names
    member _.GetTableNames() =
        match workbook with
        | Some wb -> Rbt.Spreadsheet.Workbook.getTableNames wb
        | None -> []

    /// Insert a row at the specified index
    member _.InsertRow(atRow: int) =
        applyStructureEdit (InsertRows (atRow, 1)) (sprintf "Insert row at %d" atRow)
        render()

    /// Delete a row at the specified index
    member _.DeleteRow(atRow: int) =
        applyStructureEdit (DeleteRows (atRow, 1)) (sprintf "Delete row %d" atRow)
        render()

    /// Insert a column at the specified index
    member _.InsertColumn(atCol: int) =
        applyStructureEdit (InsertCols (atCol, 1)) (sprintf "Insert column at %s" (columnToLetter atCol))
        render()

    /// Delete a column at the specified index
    member _.DeleteColumn(atCol: int) =
        applyStructureEdit (DeleteCols (atCol, 1)) (sprintf "Delete column %s" (columnToLetter atCol))
        render()

    /// Start a search
    member _.StartSearch(query: string, caseSensitive: bool) =
        model <- startSearch query caseSensitive model
        render()

    /// Find next occurrence
    member _.FindNext() =
        match nextSearchResult model with
        | Some newModel ->
            model <- newModel
            render()
            true
        | None -> false

    /// Replace current selection
    member _.Replace(findText: string, replaceText: string, caseSensitive: bool) =
        let replaced = replaceInSelection findText replaceText caseSensitive
        if replaced then render()
        replaced

    /// Replace all occurrences
    member _.ReplaceAll(findText: string, replaceText: string, caseSensitive: bool) =
        replaceEverywhere findText replaceText caseSensitive |> ignore
        render()

    /// Get search results count
    member _.GetSearchResultsCount() =
        model.SearchResults.Length

    /// Get current search index
    member _.GetSearchIndex() =
        model.SearchIndex

    /// Add a conditional format rule to a cell or range
    member _.AddConditionalFormat(row: int, col: int, ruleType: string, threshold: float, bgColor: string) =
        let addr = { Row = row; Col = col }
        let rule =
            match ruleType with
            | "greaterThan" -> GreaterThan threshold
            | "lessThan" -> LessThan threshold
            | "equal" -> Equal threshold
            | _ -> GreaterThan threshold

        let format = {
            Rule = rule
            Style = { defaultFormat with BgColor = Some bgColor }
        }

        model <- addConditionalFormat addr format model
        render()

    /// Remove conditional formatting from a cell
    member _.RemoveConditionalFormats(row: int, col: int) =
        let addr = { Row = row; Col = col }
        model <- removeConditionalFormats addr model
        render()

    /// Show conditional formatting dialog for current selection
    member _.ShowConditionalFormattingDialog() =
        showConditionalFormattingDialog()

    /// Show find/replace dialog
    member _.ShowFindReplaceDialog() =
        showFindReplaceDialog()

    /// Import CSV data
    member _.ImportCSV(csvText: string, startRow: int, startCol: int) =
        let oldCells = model.Cells
        model <- importCSV csvText startRow startCol model
        recalculateFormulas()

        // Undo restores the full pre-import cell map (an O(1) reference to the
        // persistent map) so dependent formulas recalculated by the import are
        // rolled back too, not just the imported rectangle.
        let command = {
            Undo = fun m -> restoreCellMap oldCells m
            Redo = fun m ->
                let sheet = workbook |> Option.map (fun w -> w.ActiveSheet)
                importCSV csvText startRow startCol m
                |> recalculateAllFormulas workbook sheet
                |> spillArrays workbook sheet
                |> recalculateAllFormulas workbook sheet
            Description = "Import CSV"
        }
        model <- pushCommand command model
        render()

    /// Export spreadsheet to CSV
    member _.ExportCSV(): string =
        exportCSV model

    /// Import XLSX file from ArrayBuffer (returns Promise). Loads EVERY
    /// sheet — pass `sheetName` to pick which one shows first; pass None
    /// to land on sheet 0. After this resolves, SheetNames() returns the
    /// full list and the chrome can drive SwitchSheet for the rest.
    /// Use: await spreadsheet.ImportXlsxFromBuffer(buffer, sheetName)
    /// To import from File: const buffer = await file.arrayBuffer(); await spreadsheet.ImportXlsxFromBuffer(buffer, null);
    member this.ImportXlsxFromBuffer(buffer: JS.ArrayBuffer, sheetName: string option): JS.Promise<unit> =
        // The promise rejects when the file cannot be read; the current
        // workbook is left untouched in that case.
        (importWorkbookPromise buffer).``then``(fun imported ->
            // Land on the requested sheet, else the first one Excel shows.
            let active =
                match sheetName with
                | Some n -> imported.Sheets |> Array.tryFindIndex (fun sheet -> sheet.Name = n) |> Option.defaultValue -1
                | None -> -1
            // The imported context (defined names, tables, hidden sheets) is
            // what lets later recalculation resolve cross-sheet references.
            this.InstallImportedWorkbook(imported, active))

    /// Get list of sheet names from XLSX buffer (returns Promise). Hidden
    /// sheets are included; see GetXlsxSheetInfos.
    /// Use: const sheets = await spreadsheet.GetXlsxSheetNames(buffer)
    member _.GetXlsxSheetNames(buffer: JS.ArrayBuffer): JS.Promise<string[]> =
        getSheetNamesPromise buffer

    /// `{ name, state, visible }` for every sheet in an XLSX buffer, so a
    /// sheet picker can leave out hidden sheets.
    member _.GetXlsxSheetInfos(buffer: JS.ArrayBuffer): JS.Promise<obj[]> =
        getSheetInfosPromise buffer

    // ===== Formula Bar Integration =====

    /// Get the current selection as a cell reference string (e.g., "A1" or "A1:B10")
    member _.GetSelectionRef(): string =
        match model.Selection with
        | SingleCell (row, col) ->
            sprintf "%s%d" (columnToLetter col) (row + 1)
        | Range (startRow, startCol, endRow, endCol) ->
            sprintf "%s%d:%s%d"
                (columnToLetter startCol) (startRow + 1)
                (columnToLetter endCol) (endRow + 1)

    /// Get the formula or value of the currently selected cell
    member _.GetCurrentCellFormula(): string =
        match model.Selection with
        | SingleCell (row, col) -> SpreadsheetInput.editText (getCell { Row = row; Col = col } model)
        | Range _ -> ""  // Can't edit range

    /// Set the value of the currently selected cell from formula bar. The
    /// text is read like any typed entry and checked against the cell's
    /// validation; false when the entry was refused.
    member _.SetCurrentCellFromFormula(input: string) : bool =
        match model.Selection with
        | SingleCell (row, col) ->
            let accepted = commitInput { Row = row; Col = col } input true true
            scheduleNoticeDismiss()
            render()
            accepted
        | Range _ -> false  // Can't edit range

    /// Undo last operation
    member _.Undo() =
        match undo model with
        | Some newModel ->
            model <- newModel
            recalculateFormulas()
            render()
        | None -> ()

    /// Redo last undone operation
    member _.Redo() =
        match redo model with
        | Some newModel ->
            model <- newModel
            recalculateFormulas()
            render()
        | None -> ()

    /// Sort current column ascending
    member _.SortCurrentColumnAsc() =
        let col = match model.Selection with SingleCell (_, col) -> col | Range (_, startCol, _, _) -> startCol
        sortActiveSheet col true (sprintf "Sort %s ascending" (columnToLetter col))
        render()

    /// Sort current column descending
    member _.SortCurrentColumnDesc() =
        let col = match model.Selection with SingleCell (_, col) -> col | Range (_, startCol, _, _) -> startCol
        sortActiveSheet col false (sprintf "Sort %s descending" (columnToLetter col))
        render()

    /// Cut selection to clipboard
    member _.Cut() =
        cutSelection()
        render()

    /// Copy selection to clipboard
    member _.Copy() =
        let clipboardData = serializeSelection model.Selection model false
        model <- { model with Clipboard = Some clipboardData }

    /// Paste from clipboard
    member _.Paste() =
        pasteClipboardAtSelection()
        render()

/// Export for JavaScript
[<ExportDefault>]
let createSpreadsheet (container: Element): Spreadsheet =
    new Spreadsheet(container)
