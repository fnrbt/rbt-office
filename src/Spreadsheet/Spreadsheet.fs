module Rbt.Spreadsheet.Spreadsheet

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Grid.Dom
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.XlsxImporter
module SpreadsheetRecalc = Rbt.Spreadsheet.Recalc
module SpreadsheetLayout = Rbt.Spreadsheet.Layout
module SpreadsheetPaint = Rbt.Spreadsheet.Paint
module SpreadsheetCellText = Rbt.Spreadsheet.CellText
module SpreadsheetInput = Rbt.Spreadsheet.Input
module SheetOps = Rbt.Spreadsheet.SheetOps
module Host = Rbt.Spreadsheet.HostShapes

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

/// The colour of the grid lines of the DOM grid.
let private defaultGridColor = "#e0e0e0"

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

// The width of a cell's text in the grid's own font (Arial at the cell's
// size). Widths are remembered: the cells of a sheet repeat their texts, and
// setting a canvas font costs more than measuring with it.
[<Emit("(function(text,fontSize,bold,italic){const g=globalThis.__rbSpreadsheetCellMeasure||(globalThis.__rbSpreadsheetCellMeasure={ctx:document.createElement('canvas').getContext('2d'),font:'',widths:new Map()});const font=(bold?'bold ':'')+(italic?'italic ':'')+fontSize+'px Arial, sans-serif';const key=font+'|'+text;let width=g.widths.get(key);if(width===undefined){if(g.font!==font){g.ctx.font=font;g.font=font;}width=g.ctx.measureText(text).width;if(g.widths.size>20000)g.widths.clear();g.widths.set(key,width);}return width;})($0,$1,$2,$3)")>]
let private cellTextWidth (text: string) (fontSize: float) (bold: bool) (italic: bool): float = jsNative

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

[<Emit("$0.addEventListener($1, $2)")>]
let private addListenerRaw (target: Element) (eventName: string) (handler: obj): unit = jsNative

[<Emit("$0.removeEventListener($1, $2)")>]
let private removeListenerRaw (target: Element) (eventName: string) (handler: obj): unit = jsNative

[<Emit("(!!$1 && typeof $0.contains === 'function' && typeof $1.nodeType === 'number' && $0.contains($1))")>]
let private nodeContains (root: Element) (node: obj): bool = jsNative

let private sameRef (a: 'T) (b: 'T) = obj.ReferenceEquals(a, b)

/// What the package writer resolves to (a Blob in the browser) as bytes.
[<Emit("$0.then((written) => (written instanceof Uint8Array ? written : (written && typeof written.arrayBuffer === 'function' ? written.arrayBuffer().then((buffer) => new Uint8Array(buffer)) : new Uint8Array(written))))")>]
let private packageBytes (written: JS.Promise<obj>): JS.Promise<byte[]> = jsNative

/// The document content a workbook holds outside its sheets (defined names
/// with their spelling and comments). A change of it is a change of the
/// document: the revision and OnChange follow it as they follow the sheets.
type private BookContent = {
    Names: Map<SheetId option * string, string>
    NameInfo: Map<SheetId option * string, DefinedNameInfo>
    /// The sheets without a tab, in tab order (the sheets themselves, their
    /// names and their order are compared sheet by sheet).
    Hidden: SheetId list
}

let private bookContentOf (workbook: Workbook option) : BookContent =
    match workbook with
    | Some wb ->
        { Names = wb.Names
          NameInfo = wb.NameInfo
          Hidden = wb.SheetOrder |> List.filter (fun id -> match Map.tryFind id wb.Sheets with Some sheet -> not sheet.Visible | None -> false) }
    | None -> { Names = Map.empty; NameInfo = Map.empty; Hidden = [] }

let private sameBookContent (a: BookContent) (b: BookContent) =
    (sameRef a.Names b.Names || a.Names = b.Names)
    && (sameRef a.NameInfo b.NameInfo || a.NameInfo = b.NameInfo)
    && a.Hidden = b.Hidden

/// Whether two models of a sheet hold the same document content: everything
/// that is part of the workbook, as opposed to view and editing state
/// (selection, edit buffer, clipboard, search, undo history).
let private sameSheetContent (a: SpreadsheetModel) (b: SpreadsheetModel) =
    sameRef a b
    || (sameRef a.Cells b.Cells
        && sameRef a.RichRuns b.RichRuns
        && sameRef a.RowHeights b.RowHeights
        && sameRef a.ColumnWidths b.ColumnWidths
        && (sameRef a.Axes b.Axes || a.Axes = b.Axes)
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
        && a.ShowHeaders = b.ShowHeaders
        && a.ShowFormulas = b.ShowFormulas
        && a.ShowZeros = b.ShowZeros
        && a.TabColor = b.TabColor
        && a.VeryHidden = b.VeryHidden
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

[<Emit("$0.closest($1)")>]
let private closest (el: Element) (selector: string): Element = jsNative

[<Emit("$0.scrollLeft = $1")>]
let private setScrollLeft (el: Element) (value: float): unit = jsNative

[<Emit("$0.removeAttribute($1)")>]
let private removeAttribute (el: Element) (name: string): unit = jsNative

[<Emit("$0.parentNode === $1")>]
let private isChildOf (el: Element) (parent: Element): bool = jsNative

[<Emit("$0.getBoundingClientRect().right + window.scrollX")>]
let private pageRight (el: Element): float = jsNative

[<Emit("$0.getBoundingClientRect().bottom + window.scrollY")>]
let private pageBottom (el: Element): float = jsNative

[<Emit("window.scrollX")>]
let private windowScrollX (): float = jsNative

[<Emit("window.scrollY")>]
let private windowScrollY (): float = jsNative

[<Emit("($0.button || 0)")>]
let private mouseButton (e: obj): int = jsNative

[<Emit("!!$0[$1]")>]
let private eventFlag (e: obj) (name: string): bool = jsNative

[<Emit("document.activeElement === $0")>]
let private isActiveElement (el: Element): bool = jsNative

// ---- system clipboard ----

[<Emit("($0.clipboardData ? ($0.clipboardData.getData($1) || '') : '')")>]
let private clipboardEventGet (e: obj) (flavour: string): string = jsNative

[<Emit("($0.clipboardData ? ($0.clipboardData.setData($1, $2), true) : false)")>]
let private clipboardEventSet (e: obj) (flavour: string) (data: string): bool = jsNative

/// Put text and an HTML table on the system clipboard (outside a copy event).
/// Resolves to false when the browser does not allow it.
[<Emit("(function (text, html) { try { if (typeof navigator !== 'undefined' && navigator.clipboard && typeof ClipboardItem !== 'undefined') return navigator.clipboard.write([new ClipboardItem({ 'text/plain': new Blob([text], { type: 'text/plain' }), 'text/html': new Blob([html], { type: 'text/html' }) })]).then(function () { return true; }, function () { return false; }); if (typeof navigator !== 'undefined' && navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(text).then(function () { return true; }, function () { return false; }); } catch (e) { } return Promise.resolve(false); })($0, $1)")>]
let private writeSystemClipboard (text: string) (html: string): JS.Promise<bool> = jsNative

/// Read the system clipboard: [| text; html |] ("" for a missing flavour), or
/// null when the browser does not allow it.
[<Emit("(function () { var none = function () { return null; }; try { var c = typeof navigator !== 'undefined' ? navigator.clipboard : null; if (c && c.read) return c.read().then(function (items) { var jobs = []; var out = ['', '']; items.forEach(function (item) { if (item.types.indexOf('text/plain') >= 0) jobs.push(item.getType('text/plain').then(function (b) { return b.text(); }).then(function (t) { out[0] = t; })); if (item.types.indexOf('text/html') >= 0) jobs.push(item.getType('text/html').then(function (b) { return b.text(); }).then(function (t) { out[1] = t; })); }); return Promise.all(jobs).then(function () { return out; }); }).catch(function () { return c.readText ? c.readText().then(function (t) { return [t, '']; }, none) : null; }); if (c && c.readText) return c.readText().then(function (t) { return [t, '']; }, none); } catch (e) { } return Promise.resolve(null); })()")>]
let private readSystemClipboard (): JS.Promise<string array> = jsNative

/// The cells of the first table of an HTML fragment (what Excel and Sheets
/// put on the clipboard), row by row, or null when it has no table.
[<Emit("(function (html) { if (!html || html.indexOf('<t') < 0) return null; var doc = new DOMParser().parseFromString(html, 'text/html'); var table = doc.querySelector('table'); if (!table) return null; return Array.prototype.map.call(table.rows, function (row) { return Array.prototype.map.call(row.cells, function (cell) { var lines = []; var walk = function (node) { if (node.nodeType === 3) lines.push(node.nodeValue.replace(/\\s+/g, ' ')); else if (node.nodeName === 'BR') lines.push('\\n'); else for (var i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]); }; walk(cell); return lines.join('').replace(/\\u00a0/g, ' ').replace(/ *\\n */g, '\\n').trim(); }); }); })($0)")>]
let private htmlTableCells (html: string): string array array = jsNative

[<Emit("$0.then($1)")>]
let private promiseThen (promise: JS.Promise<'T>) (next: 'T -> 'U): JS.Promise<'U> = jsNative

[<Emit("Promise.resolve($0)")>]
let private promiseOf (value: 'T): JS.Promise<'T> = jsNative

/// Selection as text for the clipboard: cells separated by tabs, rows by line
/// feeds; a cell holding a tab, a line break or a quote is quoted, the way
/// Excel and Sheets write and read it.
let private tsvOfRows (rows: string array array) : string =
    let field (text: string) =
        if text.IndexOf '\t' >= 0 || text.IndexOf '\n' >= 0 || text.IndexOf '\r' >= 0 || text.IndexOf '"' >= 0 then
            "\"" + text.Replace("\"", "\"\"") + "\""
        else text
    rows |> Array.map (fun row -> row |> Array.map field |> String.concat "\t") |> String.concat "\n"

/// The inverse of `tsvOfRows`.
let private rowsOfTsv (text: string) : string array array =
    let rows = ResizeArray<string array>()
    let row = ResizeArray<string>()
    let field = System.Text.StringBuilder()
    let mutable quoted = false
    let mutable atFieldStart = true
    let mutable i = 0
    let endField () =
        row.Add (field.ToString())
        field.Clear() |> ignore
        atFieldStart <- true
    let endRow () =
        endField ()
        rows.Add (row.ToArray())
        row.Clear()
    while i < text.Length do
        let ch = text.[i]
        if quoted then
            if ch = '"' then
                if i + 1 < text.Length && text.[i + 1] = '"' then
                    field.Append '"' |> ignore
                    i <- i + 1
                else quoted <- false
            else field.Append ch |> ignore
        elif ch = '"' && atFieldStart then
            quoted <- true
            atFieldStart <- false
        elif ch = '\t' then endField ()
        elif ch = '\r' then
            if i + 1 < text.Length && text.[i + 1] = '\n' then i <- i + 1
            endRow ()
        elif ch = '\n' then endRow ()
        else
            field.Append ch |> ignore
            atFieldStart <- false
        i <- i + 1
    // A trailing line end does not start another row.
    if field.Length > 0 || row.Count > 0 then endRow ()
    rows.ToArray()

let private htmlEscape (text: string) =
    text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")

/// The node of a cell on display, and what it was last drawn from.
type private CellNode = {
    Element: Element
    mutable Key: float
    /// 0 body, 1 frozen rows, 2 frozen columns, 3 frozen corner.
    mutable Pane: int
    /// The cell's data (the CellData record itself, or null).
    mutable Data: obj
    /// The data of the cells whose borders and fills the node's edges also
    /// depend on: the cell above, the cell to the left, and the one above that.
    mutable NearAbove: obj
    mutable NearLeft: obj
    mutable NearCorner: obj
    mutable X: float
    mutable Y: float
    mutable W: float
    mutable H: float
    /// The style epoch the node was drawn in (see `render`).
    mutable Epoch: int
    /// The render pass that last wanted the node.
    mutable Seen: int
    mutable HasTitle: bool
    /// The cell's text is wider than the cell: how it is drawn (run over the
    /// empty cells beside it, or cut off) depends on its neighbours.
    mutable Wide: bool
}

/// The node of a row or column header on display.
type private HeaderNode = {
    Element: Element
    Label: Element
    /// The filter button (column headers) and the resize handle.
    Button: Element option
    Handle: Element
    mutable Index: int
    mutable Signature: string
    mutable Seen: int
}

/// The fixed nodes of the grid (see `Mount`).
type private GridLayers = {
    /// Where cells go, by pane: the scrolled content, and the three sticky
    /// wrappers of the frozen panes.
    Panes: Element array
    ColHeaders: Element
    RowHeaders: Element
    /// Pinned on both axes: the corner, and the headers of frozen rows / columns.
    FrozenHeaders: Element
    /// The selection rectangle and the active-cell box, one per pane.
    Selections: Element array
    Actives: Element array
    FillHandle: Element
    EditInput: Element
    Notice: Element
    InputMessage: Element
    /// The button beside the active cell that opens its validation list.
    ListButton: Element
}

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
    let mutable notifiedBook: BookContent = bookContentOf None
    let mutable notifiedSelection: (int * Selection * CellAddress) option = None
    // --- revisions ----------------------------------------------------------------
    // The revision is a number that names the workbook's content: every
    // change gives a new one, and undo / redo bring back the number the
    // content had before (so a host that remembers the revision it saved
    // knows the workbook is clean again after undoing back to it).
    let mutable revision = 0
    let mutable lastRevision = 0
    let mutable savedRevision = 0
    /// The content the current revision was given to.
    let mutable revisionSheets: SpreadsheetModel array = Array.copy sheets
    let mutable revisionNames: string array = Array.copy sheetNames
    /// ... and the workbook-level content (defined names) it was given to.
    let mutable revisionBook: BookContent = bookContentOf None
    /// The undoable steps in the order they were taken, each with the
    /// revision before and after it, and the ones undone since (for redo).
    let appliedSteps = ResizeArray<Command * int * int>()
    let undoneSteps = ResizeArray<Command * int * int>()
    /// Undo what Mount set up outside the grid's own nodes (document-level
    /// listeners, observers, overlays appended to <body>).
    let cleanups = ResizeArray<unit -> unit>()
    let mutable disposed = false
    /// Every DOM listener the component registered: (target, event, handler).
    /// Dispose removes them all, also those on nodes that are no longer in
    /// the page, so that nothing the component made holds on to it.
    let listeners = ResizeArray<obj * string * obj>()
    /// Register a listener (the same function object is kept for removal).
    let listen (target: Element) (eventName: string) (handler: obj -> unit) =
        let boxed = box handler
        addListenerRaw target eventName boxed
        listeners.Add((box target, eventName, boxed))
    let unlisten (target: Element) (eventName: string) (handler: obj -> unit) =
        // The handler registered under this name on this target (a function
        // value handed over twice need not be the same object).
        for index = listeners.Count - 1 downto 0 do
            let known, name, boxed = listeners.[index]
            if obj.ReferenceEquals(known, target) && name = eventName
               && (obj.ReferenceEquals(boxed, box handler) || listeners |> Seq.filter (fun (t, n, _) -> obj.ReferenceEquals(t, target) && n = eventName) |> Seq.length = 1) then
                removeListenerRaw target eventName boxed
                listeners.RemoveAt index
    /// Take a dialog or dropdown out of the page and drop the listeners on its nodes.
    let closePopup (popup: Element) =
        detach popup
        for index = listeners.Count - 1 downto 0 do
            let known, name, boxed = listeners.[index]
            if nodeContains popup known then
                removeListenerRaw (unbox known) name boxed
                listeners.RemoveAt index
    let mutable workbook: Workbook option = None  // Optional workbook for table references
    /// The XLSX file the current workbook was imported from, for ExportXlsx:
    /// what the editor does not model is carried over from it.
    let mutable xlsxSource: Rbt.Spreadsheet.XlsxBridge.XlsxSource option = None
    /// The dependency graph of the workbook as it was last calculated (see
    /// Recalc): an edit of cells evaluates the formulas that read them, not
    /// the workbook. Built at the first edit (or shortly after a workbook is
    /// installed, when no sheet waits to be converted) from
    /// `calculatedWorkbook`, the workbook as it was last calculated or
    /// installed, and kept up to date by the edits.
    let mutable recalcEngine: SpreadsheetRecalc.Engine option = None
    let mutable calculatedWorkbook: Workbook option = None
    let mutable engineGeneration = 0
    /// What the last recalculation did: "incremental" (with the engine's
    /// counts) or "full".
    /// False: every edit recalculates the whole workbook (SetIncrementalRecalculation).
    let mutable incrementalRecalculation = true
    let mutable lastRecalculation = "none"
    let mutable lastFullReason = ""
    let mutable lastRecalcStats : SpreadsheetRecalc.RecalcStats = { ChangedCells = 0; ReachedFormulas = 0; EvaluatedFormulas = 0 }
    let mutable readOnly = false
    let mutable scrollTop = 0.0
    let mutable scrollLeft = 0.0
    let mutable containerHeight = 600.0
    let mutable containerWidth = 800.0
    let mutable useSharedLayoutGeometry = true

    let rowHeight = 25.0
    let colWidth = 100.0
    /// The room the column letters and row numbers take: none on a sheet
    /// that hides them (see syncView).
    let mutable headerHeight = 30.0
    let mutable rowHeaderWidth = 50.0
    let mutable appliedHeaders = true
    /// The zoom the grid is drawn at (the zoom of the sheet on display, see
    /// syncView). The scrolled content - cells, headers, overlays, pictures,
    /// the edit box - is laid out at 100 % in sheet pixels and scaled as a
    /// whole by the browser (CSS `zoom`, which scales layout and text, not a
    /// bitmap). Everything in the component is in sheet pixels; what is read
    /// from or written to the page (scroll offsets, the size of the grid,
    /// pointer positions) is converted here.
    let mutable appliedZoom = 1.0
    let gridScrollTop (grid: Element) = getScrollTop grid / appliedZoom
    let gridScrollLeft (grid: Element) = getScrollLeft grid / appliedZoom
    let gridClientWidth (grid: Element) = getClientWidth grid / appliedZoom
    let gridClientHeight (grid: Element) = getClientHeight grid / appliedZoom
    let setGridScrollTop (grid: Element) (value: float) = setScrollTop grid (value * appliedZoom)
    let setGridScrollLeft (grid: Element) (value: float) = setScrollLeft grid (value * appliedZoom)
    let pointerX (e: obj) (grid: Element) = (getClientX e - getBoundingLeft grid) / appliedZoom
    let pointerY (e: obj) (grid: Element) = (getClientY e - getBoundingTop grid) / appliedZoom

    let mutable gridContainer: Element option = None
    let mutable contentElement: Element option = None
    // --- what is on display ---------------------------------------------------
    // Cell and header nodes are kept between renders and reused: a render
    // draws the cells that came into view or whose data changed, and leaves
    // the rest alone.
    let mutable layers: GridLayers option = None
    let cellNodes = System.Collections.Generic.Dictionary<float, CellNode>()
    let freeCellNodes = ResizeArray<CellNode>()
    let colHeaderNodes = System.Collections.Generic.Dictionary<int, HeaderNode>()
    let rowHeaderNodes = System.Collections.Generic.Dictionary<int, HeaderNode>()
    let freeColHeaders = ResizeArray<HeaderNode>()
    let freeRowHeaders = ResizeArray<HeaderNode>()
    let drawingNodes = ResizeArray<Element>()
    let mutable renderGeneration = 0
    /// Bumped when everything on display has to be drawn again (another sheet,
    /// conditional formats, tables, read-only mode ...).
    let mutable styleEpoch = 0
    /// What the last render drew from.
    let mutable renderedCells: obj = null
    /// The cells the grid looked up while drawing, by row and column, for the
    /// cell map they were looked up in: a cell's edges are shared with its
    /// neighbours, so each visible cell is asked for several times in a pass.
    let lookedUpCells = System.Collections.Generic.Dictionary<float, CellData option>()
    let mutable lookedUpIn : obj = null
    let cellAt (row: int) (col: int) : CellData option =
        if row < 0 || col < 0 then None
        else
            if not (obj.ReferenceEquals(lookedUpIn, model.Cells)) || lookedUpCells.Count > 40000 then
                lookedUpCells.Clear()
                lookedUpIn <- box model.Cells
            let key = float row * 16384.0 + float col
            match lookedUpCells.TryGetValue key with
            | true, known -> known
            | false, _ ->
                let found = getCell { Row = row; Col = col } model
                lookedUpCells.[key] <- found
                found
    let mutable renderedStyleInputs: obj array = [||]
    let mutable renderedDrawings: obj = null
    let mutable renderedMetrics: obj = null
    let mutable extentWidth = -1.0
    let mutable extentHeight = -1.0
    /// For the rows on display: the nearest filled cell left and right of
    /// the columns in view when it holds text (its text may run into the
    /// view), else -1 / -2. Valid for `overflowCells` and the column range.
    let overflowSources = System.Collections.Generic.Dictionary<int, struct (int * int)>()
    let mutable overflowCells: obj = null
    let mutable overflowFirstCol = -1
    let mutable overflowLastCol = -1
    /// The active cell: where typing goes, and the cell a Shift+arrow or a
    /// Shift+click extends the selection from. The model's Selection is the
    /// rectangle; this is the cell it was started at.
    let mutable anchorCell : CellAddress = { Row = 0; Col = 0 }
    /// The active cell of each sheet (parallel to `sheets`), kept across sheet switches.
    let mutable sheetAnchors : CellAddress array = [| { Row = 0; Col = 0 } |]
    let editStateListeners = ResizeArray<obj -> unit>()
    /// The edit state the listeners were last told (cell and text).
    let mutable notifiedEdit : (CellAddress * string) option = None
    /// What Copy / Cut last put on the system clipboard: a paste of the same
    /// text is a paste of the grid's own cells (formulas, formats).
    let mutable copiedText : string = ""
    /// Counts the clipboard events the grid handled (see handleKeyDown).
    let mutable clipboardEvents = 0

    // --- sheets converted on demand ---------------------------------------------
    // ImportXlsx converts a sheet of the file to a model when it is first
    // needed: when it is shown, when a formula reads it, and - for a sheet
    // with formulas, conditional formats or validations of its own - at the
    // first edit of the workbook (what it shows can depend on the edit, and
    // an edit of rows / columns has to reach its formulas). Until then
    // `sheets` holds a stand-in without cells (see Model.resolveSheet), and
    // ExportXlsx writes the sheet from the file.

    /// Put the sheet at `index` in place of its stand-in (converting it now
    /// unless a formula already made it). Not a change of the workbook: what
    /// the revision and the change listeners were last told is updated too.
    let convertInPlace (index: int) =
        if index >= 0 && index < sheets.Length then
            match sheets.[index].Deferred with
            | Some _ ->
                let standIn = sheets.[index]
                let converted = resolveSheet standIn
                // (What was changed on the sheet while it was not converted.)
                let loaded =
                    if converted.TabColor = standIn.TabColor && converted.VeryHidden = standIn.VeryHidden then converted
                    else { converted with TabColor = standIn.TabColor; VeryHidden = standIn.VeryHidden }
                sheets.[index] <- loaded
                if index < revisionSheets.Length && revisionSheets.[index].Deferred.IsSome then revisionSheets.[index] <- loaded
                if index < notifiedSheets.Length && notifiedSheets.[index].Deferred.IsSome then notifiedSheets.[index] <- loaded
            | None -> ()

    /// Sheets a formula converted while it was evaluated take their place.
    let adoptConvertedSheets () =
        for index in 0 .. sheets.Length - 1 do
            if index <> activeSheetIndex then
                match sheets.[index].Deferred with
                | Some deferred when deferred.Loaded.IsSome -> convertInPlace index
                | _ -> ()

    /// Before the first edit: convert every sheet that has formulas (or
    /// rules) of its own. Sheets of plain data stay as they are.
    let mutable formulaSheetsConverted = true
    let convertFormulaSheets () =
        if not formulaSheetsConverted then
            formulaSheetsConverted <- true
            for index in 0 .. sheets.Length - 1 do
                if index <> activeSheetIndex then
                    match sheets.[index].Deferred with
                    | Some deferred when deferred.HasFormulas -> convertInPlace index
                    | _ -> ()

    /// The undo steps that change no value (formats, merges): undoing or
    /// redoing one needs no recalculation of the workbook.
    let valueNeutralSteps = ResizeArray<Command>()
    let pushValueNeutral (command: Command) =
        valueNeutralSteps.Add command
        if valueNeutralSteps.Count > 200 then valueNeutralSteps.RemoveAt 0
        model <- pushCommand command model
    let isValueNeutral (command: Command option) =
        match command with
        | Some command -> valueNeutralSteps |> Seq.exists (fun known -> obj.ReferenceEquals(known, command))
        | None -> false

    let convertAllSheets () =
        formulaSheetsConverted <- true
        for index in 0 .. sheets.Length - 1 do
            if index <> activeSheetIndex then convertInPlace index

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

    /// Put a sheet's cells back with the spill map they had: the cells a
    /// dynamic array had spilled into are spilled cells again, not ordinary
    /// ones in the formula's way (it showed #SPILL! after the undo of a Delete
    /// or Cut that had removed it).
    let restoreCellMapWithSpill (cells: Map<CellAddress, CellData>) (spill: Map<CellAddress, CellAddress>) (m: SpreadsheetModel) =
        { m with Cells = cells; Spill = spill } |> invalidateFilterProjection

    /// The workbook was replaced or changed in more than its cells: the
    /// dependency graph and the calculated state it stands on are dropped.
    let invalidateCalcEngine () =
        recalcEngine <- None
        calculatedWorkbook <- None
        engineGeneration <- engineGeneration + 1

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
              Names = Map.empty
              NameInfo = Map.empty }

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
                Names = wb.Names
                NameInfo = wb.NameInfo })

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

    /// Reads the source of a list validation of the active sheet: a range on
    /// it or on another sheet (which is converted when it was not needed so
    /// far), or a defined name that is one.
    let listReader : ListSourceReader =
        SheetOps.listSourceReader
            (fun name ->
                sheetNames
                |> Array.tryFindIndex (fun known -> System.String.Equals(known, name, System.StringComparison.OrdinalIgnoreCase))
                |> Option.map (fun index ->
                    if index = activeSheetIndex then model
                    else
                        convertInPlace index
                        sheets.[index]))
            (fun name -> workbook |> Option.bind (fun wb -> Rbt.Spreadsheet.Workbook.getName (Some (sheetIdForIndex activeSheetIndex)) name wb))

    /// The cells CircleInvalid circled on each sheet (parallel to `sheets`
    /// by the sheet's model identity is not possible, so: by sheet index).
    let invalidCircles = System.Collections.Generic.Dictionary<int, CellAddress array>()
    let circleNodes = ResizeArray<Element>()

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
                // Text wider than its cell is cut off at the cells it may run over.
                (match text.Clip with
                 | Some clip -> canvasClipRect ctx clip.X clip.Y clip.Width clip.Height
                 | None -> ())
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
    /// The engine for the workbook as it was last calculated, built now when there is none yet.
    let ensureRecalcEngine () =
        match recalcEngine, calculatedWorkbook with
        | Some engine, _ -> Some engine
        | None, Some calculated ->
            let engine = SpreadsheetRecalc.build calculated
            recalcEngine <- Some engine
            Some engine
        | None, None -> None

    /// Recalculate after an edit. `hint` names the cells the edit touched
    /// when they are known. The formulas that read the changed cells are
    /// evaluated (Recalc); the whole workbook is recalculated - every sheet:
    /// formulas, then dynamic-array spills, then formulas again so readers of
    /// spilled cells settle - when the edit is more than one of cells, when
    /// it reaches a spill or a cycle, and before there is a calculated state
    /// to start from.
    let recalculateAfterEdit (hint: (SheetId * CellAddress) list option) =
        let edited = currentWorkbookSnapshot()
        let incremental =
            match (if incrementalRecalculation then ensureRecalcEngine () else None) with
            | Some engine ->
                match SpreadsheetRecalc.recalculate engine edited hint with
                | Some recalculated ->
                    lastRecalculation <- "incremental"
                    lastRecalcStats <- SpreadsheetRecalc.lastStats engine
                    Some recalculated
                | None ->
                    lastFullReason <- SpreadsheetRecalc.lastDecline engine
                    None
            | None ->
                lastFullReason <- if incrementalRecalculation then "no calculated state to start from" else "incremental recalculation is switched off"
                None
        match incremental with
        | Some recalculated ->
            installWorkbookSnapshot recalculated
            calculatedWorkbook <- Some recalculated
        | None ->
            lastRecalculation <- "full"
            // Every sheet with formulas is recalculated: those not converted yet are now.
            convertFormulaSheets ()
            let edited = currentWorkbookSnapshot()
            let recalculated = recalculateWorkbook edited
            installWorkbookSnapshot recalculated
            // The engine goes on from the result, or is built again at the next edit.
            match recalcEngine with
            | Some engine when not (SpreadsheetRecalc.adopt engine recalculated) -> recalcEngine <- None
            | _ -> ()
            calculatedWorkbook <- Some recalculated
        adoptConvertedSheets ()

    /// Recalculation after an edit whose cells are not named: the sheets'
    /// cells are compared with the calculated state.
    let recalculateFormulas () =
        recalculateAfterEdit None

    /// Full recalculation of the whole workbook, whatever the edit was (rows /
    /// columns moved, sheets, names).
    let recalculateEverything () =
        invalidateCalcEngine()
        convertFormulaSheets ()
        let recalculated = currentWorkbookSnapshot() |> recalculateWorkbook
        installWorkbookSnapshot recalculated
        lastRecalculation <- "full"
        lastFullReason <- "rows, columns, sheets or names changed"
        adoptConvertedSheets ()
        calculatedWorkbook <- Some (currentWorkbookSnapshot())

    /// Insert or delete rows/columns on the active sheet as one undoable
    /// step. The edit and its undo both go through the workbook so formulas on
    /// other sheets, defined names and structured-reference tables follow.
    let applyStructureEdit (edit: StructureEdit) (description: string) =
        // Formulas of every sheet follow the rows / columns that move.
        convertFormulaSheets ()
        let sheetId = sheetIdForIndex activeSheetIndex
        let edited, undoInfo = Rbt.Spreadsheet.Workbook.applyStructureEdit sheetId edit (currentWorkbookSnapshot())
        installWorkbookSnapshot edited
        recalculateEverything()
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
        let oldSpill = model.Spill
        let oldRichRuns = model.RichRuns
        let oldSortState = model.SortState
        model <- sortByColumn col ascending model
        recalculateFormulas()
        let command = {
            Undo = fun m -> { restoreCellMapWithSpill oldCells oldSpill m with RichRuns = oldRichRuns; SortState = oldSortState }
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
        let oldSpill = model.Spill
        model <- { model with Clipboard = Some clipboardData }
        model <- clearSelectionCells selection model
        recalculateFormulas()
        if not (obj.ReferenceEquals(oldCells, model.Cells)) then
            let command = {
                Undo = fun m -> restoreCellMapWithSpill oldCells oldSpill m
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
            // (A pasted cut re-points the formulas of every sheet.)
            convertFormulaSheets ()
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
        // (The engine is built from the workbook as it is before the entry.)
        if calculatedWorkbook.IsNone && recalcEngine.IsNone then calculatedWorkbook <- Some (currentWorkbookSnapshot())
        if incrementalRecalculation then ensureRecalcEngine () |> ignore
        model <- setCell addr cellData model
        recalculateAfterEdit (Some [ sheetIdForIndex activeSheetIndex, addr ])

    let hitTestEvent (e: obj) =
        match gridContainer with
        | None -> SpreadsheetLayout.Empty
        | Some grid ->
            // Where the grid is scrolled to now (the scroll event may not have arrived yet).
            scrollTop <- gridScrollTop grid
            scrollLeft <- gridScrollLeft grid
            let x = pointerX e grid
            let y = pointerY e grid
            SpreadsheetLayout.hitTest model (currentViewport()) x y

    /// The cell under the pointer, or the nearest one when the pointer is over
    /// a header or outside the grid (a drag that runs past the edge).
    let hitCellNear (e: obj) : CellAddress option =
        match gridContainer with
        | None -> None
        | Some grid ->
            scrollTop <- gridScrollTop grid
            scrollLeft <- gridScrollLeft grid
            let width = gridClientWidth grid
            let height = gridClientHeight grid
            let x = pointerX e grid |> max (rowHeaderWidth + 1.0) |> min (width - 1.0)
            let y = pointerY e grid |> max (headerHeight + 1.0) |> min (height - 1.0)
            match SpreadsheetLayout.hitTest model (currentViewport()) x y with
            | SpreadsheetLayout.Cell addr -> Some addr
            | _ ->
                // Past the last row / column: the last one.
                let metrics = SpreadsheetLayout.metricsOf model
                if metrics.RowCount = 0 || metrics.ColCount = 0 then None
                else
                    let frozenW = metrics.ColPrefix.[max 0 (min metrics.ColCount model.FrozenCols)]
                    let frozenH = metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)]
                    let sheetX = if x - rowHeaderWidth < frozenW then x - rowHeaderWidth else x - rowHeaderWidth + scrollLeft
                    let sheetY = if y - headerHeight < frozenH then y - headerHeight else y - headerHeight + scrollTop
                    Some { Row = SpreadsheetLayout.indexAtOffset metrics.RowPrefix metrics.RowCount sheetY
                           Col = SpreadsheetLayout.indexAtOffset metrics.ColPrefix metrics.ColCount sheetX }

    /// Grow a rectangle of cells until it holds every merged cell it touches.
    let expandToMerges (top: int) (left: int) (bottom: int) (right: int) : int * int * int * int =
        if model.MergedCells.IsEmpty then top, left, bottom, right
        else
            let mutable t, l, b, r = top, left, bottom, right
            let mutable grew = true
            while grew do
                grew <- false
                for merge in mergedCellsIntersectingRange t (b + 1) l (r + 1) model do
                    if merge.StartRow < t then (t <- merge.StartRow; grew <- true)
                    if merge.StartCol < l then (l <- merge.StartCol; grew <- true)
                    if merge.EndRow > b then (b <- merge.EndRow; grew <- true)
                    if merge.EndCol > r then (r <- merge.EndCol; grew <- true)
            t, l, b, r

    /// Select the cells from the active cell `anchor` to `extent`. The model
    /// holds the rectangle (top-left to bottom-right); the anchor stays the
    /// active cell.
    let selectFromTo (anchor: CellAddress) (extent: CellAddress) =
        anchorCell <- anchor
        let top, left, bottom, right =
            expandToMerges (min anchor.Row extent.Row) (min anchor.Col extent.Col) (max anchor.Row extent.Row) (max anchor.Col extent.Col)
        model <- { model with Selection = (if top = bottom && left = right then SingleCell (top, left) else Range (top, left, bottom, right)) }

    let selectCellAddress (addr: CellAddress) =
        selectFromTo addr addr

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

    // --- Validation notices ---------------------------------------------------
    // Excel answers an invalid entry with a modal alert. Here the answer never
    // blocks: a "stop" rule refuses the entry (the cell keeps its old content
    // and editing ends), "warning" and "information" rules accept it, and in
    // every case a note with the rule's title and message appears in the grid
    // and goes away by itself. Hosts can also listen for it.
    let mutable validationNotice: (string * string * string * CellAddress) option = None   // style, title, message, cell
    let mutable noticeTimer: int option = None
    let validationListeners = System.Collections.Generic.List<string -> string -> string -> unit>()

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

    /// The cell an edit applies to: the active cell of the selection (the
    /// cell it was started at; of a merged cell, its first cell). When the
    /// selection was set without one - by the host, an undo, a structure
    /// edit - it is the top-left cell.
    let activeCellAddress () =
        let top, left, bottom, right = selectionBounds()
        let inside = anchorCell.Row >= top && anchorCell.Row <= bottom && anchorCell.Col >= left && anchorCell.Col <= right
        let cell = if inside then anchorCell else { Row = top; Col = left }
        match findMergedCell cell model with
        | Some merge -> { Row = merge.StartRow; Col = merge.StartCol }
        | None -> cell

    /// The corner of the selection opposite the active cell: what Shift+arrow moves.
    let selectionExtent () =
        let top, left, bottom, right = selectionBounds()
        let active = activeCellAddress()
        { Row = (if active.Row = top then bottom else top); Col = (if active.Col = left then right else left) }

    /// Whether more than one cell is selected (one merged cell is one cell).
    let selectionIsRange () =
        let top, left, bottom, right = selectionBounds()
        if top = bottom && left = right then false
        else
            match findMergedCell { Row = top; Col = left } model with
            | Some merge -> not (merge.StartRow = top && merge.StartCol = left && merge.EndRow = bottom && merge.EndCol = right)
            | None -> true

    /// Bring the page in line with the view options of the sheet on display:
    /// its zoom (the content is scaled; the cell in the top-left corner stays
    /// there) and whether it shows its headers.
    let syncView () =
        if model.ShowHeaders <> appliedHeaders then
            appliedHeaders <- model.ShowHeaders
            headerHeight <- if appliedHeaders then 30.0 else 0.0
            rowHeaderWidth <- if appliedHeaders then 50.0 else 0.0
            match layers with
            | Some layer ->
                let display = if appliedHeaders then "" else "none"
                setStyle layer.ColHeaders "display" display
                setStyle layer.RowHeaders "display" display
                setStyle layer.FrozenHeaders "display" display
            | None -> ()
        let zoom = if model.Zoom >= 0.1 && model.Zoom <= 4.0 then model.Zoom else 1.0
        if zoom <> appliedZoom then
            match contentElement, gridContainer with
            | Some content, Some grid ->
                let top = gridScrollTop grid
                let left = gridScrollLeft grid
                appliedZoom <- zoom
                setStyle content "zoom" (if zoom = 1.0 then "" else string zoom)
                containerHeight <- gridClientHeight grid
                containerWidth <- gridClientWidth grid
                setGridScrollTop grid top
                setGridScrollLeft grid left
                scrollTop <- gridScrollTop grid
                scrollLeft <- gridScrollLeft grid
            | _ -> appliedZoom <- zoom

    /// Give the scrolled content the extent of the active sheet, so that
    /// every row and column can be scrolled to. Done whenever the sheet's
    /// size or the size of a row / column changes.
    let syncExtent () =
        syncView ()
        match contentElement, gridContainer with
        | Some content, Some grid ->
            let metrics = SpreadsheetLayout.metricsOf model
            let width = rowHeaderWidth + metrics.ColPrefix.[metrics.ColCount]
            let height = headerHeight + metrics.RowPrefix.[metrics.RowCount]
            if width <> extentWidth || height <> extentHeight then
                extentWidth <- width
                extentHeight <- height
                setStyle content "width" (string (System.Math.Round width) + "px")
                setStyle content "height" (string (System.Math.Round height) + "px")
                // A smaller sheet may have pulled the scroll position back.
                scrollTop <- gridScrollTop grid
                scrollLeft <- gridScrollLeft grid
        | _ -> ()

    /// Scroll the grid so that a cell is in view (frozen rows and columns
    /// stay put and cover the start of the scrolled area).
    let revealCell (row: int) (col: int) =
        match gridContainer with
        | None -> ()
        | Some grid ->
            syncExtent()
            let metrics = SpreadsheetLayout.metricsOf model
            if metrics.RowCount > 0 && metrics.ColCount > 0 then
                let row = max 0 (min (metrics.RowCount - 1) row)
                let col = max 0 (min (metrics.ColCount - 1) col)
                if col >= model.FrozenCols then
                    let frozen = metrics.ColPrefix.[max 0 (min metrics.ColCount model.FrozenCols)]
                    let left = metrics.ColPrefix.[col]
                    let right = metrics.ColPrefix.[col + 1]
                    let width = gridClientWidth grid - rowHeaderWidth
                    let current = gridScrollLeft grid
                    let wanted =
                        if left < current + frozen then left - frozen
                        elif right > current + width then min (left - frozen) (right - width)
                        else current
                    if wanted <> current then setGridScrollLeft grid (max 0.0 wanted)
                if row >= model.FrozenRows then
                    let frozen = metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)]
                    let top = metrics.RowPrefix.[row]
                    let bottom = metrics.RowPrefix.[row + 1]
                    let height = gridClientHeight grid - headerHeight
                    let current = gridScrollTop grid
                    let wanted =
                        if top < current + frozen then top - frozen
                        elif bottom > current + height then min (top - frozen) (bottom - height)
                        else current
                    if wanted <> current then setGridScrollTop grid (max 0.0 wanted)
                scrollTop <- gridScrollTop grid
                scrollLeft <- gridScrollLeft grid

    // --- keyboard movement ------------------------------------------------------

    /// The cell `deltaRow` / `deltaCol` away from `from`, leaving a merged
    /// cell by its edge and stepping over hidden rows and columns.
    let stepFrom (from: CellAddress) (deltaRow: int) (deltaCol: int) : CellAddress =
        let metrics = SpreadsheetLayout.metricsOf model
        let merge = findMergedCell from model
        let startRow =
            match merge with
            | Some m when deltaRow > 0 -> m.EndRow
            | Some m when deltaRow < 0 -> m.StartRow
            | _ -> from.Row
        let startCol =
            match merge with
            | Some m when deltaCol > 0 -> m.EndCol
            | Some m when deltaCol < 0 -> m.StartCol
            | _ -> from.Col
        let visibleRow r = r >= metrics.RowCount || metrics.RowPrefix.[r + 1] > metrics.RowPrefix.[r]
        let visibleCol c = c >= metrics.ColCount || metrics.ColPrefix.[c + 1] > metrics.ColPrefix.[c]
        let mutable row = startRow
        let mutable col = startCol
        if deltaRow <> 0 then
            let step = sign deltaRow
            let mutable remaining = abs deltaRow
            let mutable next = row + step
            while remaining > 0 && next >= 0 && next < model.RowCount do
                if visibleRow next then
                    row <- next
                    remaining <- remaining - 1
                next <- next + step
        if deltaCol <> 0 then
            let step = sign deltaCol
            let mutable remaining = abs deltaCol
            let mutable next = col + step
            while remaining > 0 && next >= 0 && next < model.ColCount do
                if visibleCol next then
                    col <- next
                    remaining <- remaining - 1
                next <- next + step
        { Row = max 0 (min (model.RowCount - 1) row); Col = max 0 (min (model.ColCount - 1) col) }

    /// Ctrl+arrow: the edge of the block of data in a direction - the last
    /// filled cell of the run the cell is in, else the next filled cell, else
    /// the edge of the sheet.
    let dataEdge (from: CellAddress) (deltaRow: int) (deltaCol: int) : CellAddress =
        // Filled positions along the line of travel, in order.
        let filled = ResizeArray<int>()
        let hasValue (data: CellData) = match data.Value with CellValue.Empty -> false | CellValue.String "" -> false | _ -> true
        if deltaRow <> 0 then
            for KeyValue(addr, data) in model.Cells do
                if addr.Col = from.Col && hasValue data then filled.Add addr.Row
        else
            for col in 0 .. model.ColCount - 1 do
                match getCell { Row = from.Row; Col = col } model with
                | Some data when hasValue data -> filled.Add col
                | _ -> ()
        let occupied = System.Collections.Generic.HashSet<int>(filled)
        let step = if deltaRow <> 0 then sign deltaRow else sign deltaCol
        let limit = (if deltaRow <> 0 then model.RowCount else model.ColCount) - 1
        let start = if deltaRow <> 0 then from.Row else from.Col
        let target =
            if occupied.Contains start && occupied.Contains (start + step) then
                let mutable at = start
                while occupied.Contains (at + step) do at <- at + step
                at
            else
                let ahead = filled |> Seq.filter (fun index -> if step > 0 then index > start else index < start)
                if Seq.isEmpty ahead then (if step > 0 then limit else 0)
                elif step > 0 then Seq.min ahead
                else Seq.max ahead
        let target = max 0 (min limit target)
        if deltaRow <> 0 then { from with Row = target } else { from with Col = target }

    /// How many rows fit in the scrolled part of the grid (a page).
    let pageRows () =
        let metrics = SpreadsheetLayout.metricsOf model
        let frozen = metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)]
        let visible = containerHeight - headerHeight - frozen
        let active = activeCellAddress()
        let top = metrics.RowPrefix.[max 0 (min metrics.RowCount active.Row)]
        let last = SpreadsheetLayout.indexAtOffset metrics.RowPrefix metrics.RowCount (top + max 25.0 visible - 1.0)
        max 1 (last - active.Row)

    // --- clipboard text -----------------------------------------------------------

    /// The selection clipped to the cells that hold something, as rows of
    /// (address, cell) - a selected column is not a million rows of text.
    let selectionCellRows () : (CellAddress * CellData option) array array =
        let top, left, bottom, right = selectionBounds()
        let bottom, right =
            match usedBounds model with
            | Some (usedRow, usedCol) -> min bottom (max top usedRow), min right (max left usedCol)
            | None -> top, left
        [| for row in top .. bottom ->
            [| for col in left .. right ->
                let addr = { Row = row; Col = col }
                addr, getCell addr model |] |]

    let cellText (data: CellData option) : string =
        match data with
        | Some data -> SpreadsheetCellText.cellDisplayText None 1.0e6 data.Format data.Value
        | None -> ""

    /// The selection as tab-separated text (what the cells show).
    let selectionText () : string =
        selectionCellRows () |> Array.map (Array.map (fun (_, data) -> cellText data)) |> tsvOfRows

    /// The selection as an HTML table, with the formatting other spreadsheets read.
    let selectionHtml () : string =
        let sb = System.Text.StringBuilder()
        sb.Append("<table>") |> ignore
        for row in selectionCellRows () do
            sb.Append("<tr>") |> ignore
            for addr, data in row do
                // The other cells of a merged cell are part of its first cell.
                let merge = findMergedCell addr model
                match merge with
                | Some m when m.StartRow <> addr.Row || m.StartCol <> addr.Col -> ()
                | _ ->
                    sb.Append("<td") |> ignore
                    match merge with
                    | Some m ->
                        if m.EndCol > m.StartCol then sb.Append(" colspan=\"").Append(string (m.EndCol - m.StartCol + 1)).Append('"') |> ignore
                        if m.EndRow > m.StartRow then sb.Append(" rowspan=\"").Append(string (m.EndRow - m.StartRow + 1)).Append('"') |> ignore
                    | None -> ()
                    match data with
                    | Some data ->
                        let format = data.Format
                        let styles =
                            [ if format.Bold then yield "font-weight:bold"
                              if format.Italic then yield "font-style:italic"
                              if format.Underline then yield "text-decoration:underline"
                              match format.TextColor with Some color -> yield "color:" + color | None -> ()
                              match format.BgColor with Some color -> yield "background-color:" + color | None -> ()
                              match format.HorizontalAlign with Some align -> yield "text-align:" + align | None -> ()
                              for name, border in [ "border-top", format.BorderTop; "border-right", format.BorderRight
                                                    "border-bottom", format.BorderBottom; "border-left", format.BorderLeft ] do
                                  match border with
                                  | Some border when (borderCssParts border.Style).IsSome -> yield name + ":" + borderCss border
                                  | _ -> () ]
                        if not styles.IsEmpty then sb.Append(" style=\"").Append(htmlEscape (String.concat ";" styles)).Append('"') |> ignore
                    | None -> ()
                    sb.Append('>').Append((htmlEscape (cellText data)).Replace("\n", "<br>")).Append("</td>") |> ignore
            sb.Append("</tr>") |> ignore
        sb.Append("</table>").ToString()

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
                    checkEntryIn listReader ruleEvaluator addr parsed.Value model
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
        let oldSpill = model.Spill
        model <- clearSelectionCells selection model
        if not (obj.ReferenceEquals(oldCells, model.Cells)) then
            recalculateFormulas()
            model <- pushCommand { Undo = (fun m -> restoreCellMapWithSpill oldCells oldSpill m)
                                   Redo = (fun m -> clearSelectionCells selection m)
                                   Description = "Delete" } model

    /// Paste text from outside the grid (another spreadsheet, a text editor)
    /// at the active cell, as one undoable step: `rows` are the cells, each
    /// read like a typed entry. One copied cell fills the whole selection.
    let pasteExternalRows (rows: string array array) : bool =
        let rows = rows |> Array.filter (fun row -> row.Length > 0)
        if readOnly || rows.Length = 0 then false
        else
            let top, left, bottom, right = selectionBounds()
            let height = rows.Length
            let width = rows |> Array.map Array.length |> Array.max
            // A single value is repeated over the selection, as Excel does.
            let repeat = height = 1 && width = 1 && selectionIsRange ()
            let lastRow = if repeat then bottom else min (model.RowCount - 1) (top + height - 1)
            let lastCol = if repeat then right else min (model.ColCount - 1) (left + width - 1)
            let targets =
                [ for row in top .. lastRow do
                    for col in left .. lastCol do
                        let source = if repeat then rows.[0] else rows.[row - top]
                        let text = if repeat then source.[0] elif col - left < source.Length then source.[col - left] else ""
                        yield { Row = row; Col = col }, text ]
            let year = currentYear()
            let oldCells = targets |> List.map (fun (addr, _) -> addr, getCell addr model)
            for addr, text in targets do
                // Only the first cell of a merged cell takes a value.
                match findMergedCell addr model with
                | Some merge when merge.StartRow <> addr.Row || merge.StartCol <> addr.Col -> ()
                | _ ->
                    let existing = getCell addr model
                    if System.String.IsNullOrEmpty text then
                        match existing with
                        | Some cell when cell.Value <> CellValue.Empty -> model <- setCell addr { cell with Value = CellValue.Empty } model
                        | _ -> ()
                    else
                        model <- setCell addr (SpreadsheetInput.cellFromInput year existing text) model
            recalculateFormulas()
            let newCells = targets |> List.map (fun (addr, _) -> addr, getCell addr model)
            selectFromTo { Row = top; Col = left } { Row = lastRow; Col = lastCol }
            model <- pushCommand { Undo = restoreCellOptions oldCells
                                   Redo = restoreCellOptions newCells
                                   Description = "Paste" } model
            true

    /// Paste what a clipboard holds: the grid's own cells when the text is
    /// what the grid copied last (formulas and formats come along), else the
    /// cells of an HTML table, else tab-separated text.
    let pasteFlavours (text: string) (html: string) : bool =
        if readOnly then false
        else
            let normalised (value: string) = if isNull value then "" else value.Replace("\r\n", "\n").TrimEnd('\n')
            if model.Clipboard.IsSome && copiedText <> "" && normalised text = normalised copiedText then
                pasteClipboardAtSelection()
                true
            else
                let fromHtml = if System.String.IsNullOrEmpty html then null else htmlTableCells html
                if not (isNull (box fromHtml)) && fromHtml.Length > 0 then pasteExternalRows fromHtml
                elif not (System.String.IsNullOrEmpty text) then pasteExternalRows (rowsOfTsv text)
                else false

    /// Apply `update` to the format of every cell of the selection as one
    /// undoable step. A selection too large to give every cell a format of
    /// its own (whole rows or columns) changes the cells that exist.
    let formatBlockAt (description: string) (top: int, left: int, bottom: int, right: int) (updateAt: int * int * int * int -> CellAddress -> CellFormat -> CellFormat) : bool =
        let bottom = min bottom (model.RowCount - 1)
        let right = min right (model.ColCount - 1)
        let block = top, left, bottom, right
        let area = float (bottom - top + 1) * float (right - left + 1)
        let oldCells = model.Cells
        let oldSpill = model.Spill
        let apply (m: SpreadsheetModel) : SpreadsheetModel =
            let mutable cells = m.Cells
            let change (addr: CellAddress) (existing: CellData option) =
                let data = existing |> Option.defaultValue { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
                let format = updateAt block addr data.Format
                if not (obj.ReferenceEquals(format, data.Format)) && format <> data.Format then
                    cells <- Map.add addr { data with Format = format } cells
            if area <= 20000.0 then
                for row in top .. bottom do
                    for col in left .. right do
                        let addr = { Row = row; Col = col }
                        change addr (Map.tryFind addr m.Cells)
            else
                for KeyValue(addr, data) in m.Cells do
                    if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then change addr (Some data)
            if obj.ReferenceEquals(cells, m.Cells) then m else { m with Cells = cells } |> invalidateFilterProjection
        let updated = apply model
        if obj.ReferenceEquals(updated, model) then false
        else
            model <- updated
            pushValueNeutral { Undo = (fun m -> restoreCellMapWithSpill oldCells oldSpill m)
                               Redo = apply
                               Description = description }
            true

    let formatSelectionAt (description: string) (updateAt: int * int * int * int -> CellAddress -> CellFormat -> CellFormat) : bool =
        formatBlockAt description (selectionBounds()) updateAt

    let formatSelection (description: string) (update: CellFormat -> CellFormat) : bool =
        formatSelectionAt description (fun _ _ format -> update format)

    /// An argument a host must give: throws when it is undefined or null.
    let requiredArg (memberName: string) (what: string) (value: obj) : obj =
        if Host.isNullish value then Host.hostFail memberName (what + " is missing") else value

    /// Clear the cells of a block (their contents, their formats, or both and
    /// their comments) as one undoable step; what read the cleared values is
    /// recalculated. False when there was nothing to clear.
    let clearBlock (description: string) (what: SheetOps.ClearWhat) (block: int * int * int * int) : bool =
        let oldCells = model.Cells
        let oldSpill = model.Spill
        let oldRuns = model.RichRuns
        let cleared = SheetOps.clearCells what block model
        if obj.ReferenceEquals(cleared, model) then false
        else
            model <- cleared
            let command =
                { Undo = (fun m -> { restoreCellMapWithSpill oldCells oldSpill m with RichRuns = oldRuns })
                  Redo = SheetOps.clearCells what block
                  Description = description }
            if what = SheetOps.ClearFormats then pushValueNeutral command
            else
                recalculateFormulas()
                model <- pushCommand command model
            true

    /// The highest conditional-format id the component has handed out: a new
    /// rule gets one above it and above every id on its sheet, so that the id
    /// of a removed rule is not given to another.
    let mutable lastRuleId = 0

    /// Change the active sheet's conditional formats as one undoable step
    /// (no value changes: no recalculation). False when nothing changed.
    let changeConditionalFormats (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let before = model.ConditionalFormats
        let changed = change model
        let after = changed.ConditionalFormats
        if obj.ReferenceEquals(before, after) || before = after then false
        else
            model <- { model with ConditionalFormats = after }
            pushValueNeutral { Undo = (fun m -> { m with ConditionalFormats = before })
                               Redo = (fun m -> { m with ConditionalFormats = after })
                               Description = description }
            true

    /// Add a conditional format to the active sheet as one undoable step; its id.
    let addRule (ranges: CellRange list) (stopIfTrue: bool) (format: ConditionalFormat) : int =
        let id = (max lastRuleId (SheetOps.highestConditionalFormatId model)) + 1
        lastRuleId <- id
        changeConditionalFormats "Conditional format" (SheetOps.addConditionalFormatWithId id ranges stopIfTrue format) |> ignore
        id

    /// Change the active sheet's validations as one undoable step (no value
    /// changes: no recalculation). The circles of CircleInvalid go: they
    /// showed what the earlier validations made of the cells. False when
    /// nothing changed.
    let changeValidations (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let before = model.Validations
        let changed = change model
        let after = changed.Validations
        if obj.ReferenceEquals(before, after) || before = after then false
        else
            model <- { model with Validations = after }
            invalidCircles.Remove activeSheetIndex |> ignore
            pushValueNeutral { Undo = (fun m -> { m with Validations = before })
                               Redo = (fun m -> { m with Validations = after })
                               Description = description }
            true

    /// The workbook context, made now for a grid that has none yet.
    let ensureWorkbook () : Workbook =
        match workbook with
        | Some wb -> wb
        | None ->
            let wb = currentWorkbookSnapshot()
            workbook <- Some wb
            wb

    /// The index of the sheet with an id, or -1.
    let sheetIndexOfId (id: SheetId) : int =
        let mutable found = -1
        for index in 0 .. sheets.Length - 1 do
            if found < 0 && sheetIdForIndex index = id then found <- index
        found

    /// The index of the sheet with a name (as formulas match it: whatever the case), or -1.
    let sheetIndexOfName (name: string) : int =
        sheetNames |> Array.tryFindIndex (fun known -> System.String.Equals(known, name, System.StringComparison.OrdinalIgnoreCase)) |> Option.defaultValue -1

    /// Change one defined name as one undoable step: every formula is
    /// recalculated (those that use the name read something else now), and
    /// undo / redo put the name back as it was / as it became. False when
    /// nothing changed.
    let changeDefinedName (description: string) (scope: SheetId option) (name: string) (change: Workbook -> Workbook) : bool =
        let wb = ensureWorkbook ()
        let before = Rbt.Spreadsheet.Workbook.tryFindDefinedName scope name wb
        let changed = change wb
        let after = Rbt.Spreadsheet.Workbook.tryFindDefinedName scope name changed
        if before = after then false
        else
            workbook <- Some changed
            recalculateEverything()
            let restore (state: Rbt.Spreadsheet.Workbook.DefinedName option) (m: SpreadsheetModel) =
                workbook <- workbook |> Option.map (Rbt.Spreadsheet.Workbook.restoreDefinedName scope name state)
                invalidateCalcEngine()
                m
            model <- pushCommand { Undo = restore before; Redo = restore after; Description = description } model
            true

    /// Change the active sheet's frozen panes and what its view shows (grid
    /// lines, headers, formulas, zeros) as one undoable step (no value
    /// changes). False when nothing changed.
    let changeSheetView (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let state (m: SpreadsheetModel) = m.FrozenRows, m.FrozenCols, m.ShowGridLines, m.ShowHeaders, m.ShowFormulas, m.ShowZeros
        let before = state model
        let after = state (change model)
        if before = after then false
        else
            let put (rows, cols, gridLines, headers, formulas, zeros) (m: SpreadsheetModel) =
                { m with FrozenRows = rows; FrozenCols = cols; ShowGridLines = gridLines; ShowHeaders = headers; ShowFormulas = formulas; ShowZeros = zeros }
            model <- put after model
            pushValueNeutral { Undo = put before; Redo = put after; Description = description }
            true

    // --- rows and columns: sizes, hidden ones, outline ---------------------------
    // One undo step holds the sizes of the rows and columns and what the
    // sheet keeps beside them (Model.SheetAxes) before and after.
    let axesState (m: SpreadsheetModel) = m.RowHeights, m.ColumnWidths, m.Axes
    let putAxes (rowHeights: Map<int, float>, columnWidths: Map<int, float>, axes: SheetAxes) (m: SpreadsheetModel) =
        { m with RowHeights = rowHeights; ColumnWidths = columnWidths; Axes = axes }
    /// Make an undo step of the change of sizes since `before` (the grid's own
    /// resize handles change the model while the mouse moves). False when
    /// nothing changed.
    let commitAxes (description: string) (before: Map<int, float> * Map<int, float> * SheetAxes) : bool =
        let rowsBefore, colsBefore, axesBefore = before
        let after = axesState model
        let rowsAfter, colsAfter, axesAfter = after
        if (sameRef rowsBefore rowsAfter || rowsBefore = rowsAfter) && (sameRef colsBefore colsAfter || colsBefore = colsAfter)
           && (sameRef axesBefore axesAfter || axesBefore = axesAfter) then false
        else
            pushValueNeutral { Undo = putAxes before; Redo = putAxes after; Description = description }
            true
    let changeAxes (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let before = axesState model
        model <- change model
        commitAxes description before
    /// Text widths for auto-fit: in the cell's font, and never less than the
    /// grid's own measure of the text (which decides whether a number fits).
    let fitMeasure : SheetOps.TextMeasure =
        fun text fontSize fontName bold italic ->
            match fontName with
            | Some family when family <> "" ->
                max (browserTextWidth text fontSize ("'" + family.Replace("'", "") + "', Arial") bold italic) (cellTextWidth text fontSize bold italic)
            | _ -> cellTextWidth text fontSize bold italic
    /// The rows / columns an argument names: a 0-based number (with `count`,
    /// that many from it on) or an A1 range ("2:5", "B:D", "B2:D9" - its
    /// rows / columns), clipped to the sheet.
    let axisSpanArg (memberName: string) (rows: bool) (target: obj) (count: obj) : int * int =
        let limit = (if rows then model.RowCount else model.ColCount) - 1
        let what = if rows then "row" else "column"
        if Host.jsTypeOf (requiredArg memberName (what + " or range") target) = "number" then
            let first = Host.intArg memberName what target
            if first < 0 || first > limit then Host.hostFail memberName (sprintf "%s %d is not on the sheet (0..%d)" what first limit)
            let count = match Host.optionalArg (Host.intArg memberName "count") count with Some count -> count | None -> 1
            if count < 1 then Host.hostFail memberName (sprintf "count must be at least 1, got %d" count)
            first, min limit (first + count - 1)
        else
            let range = Host.rangeArg memberName target
            let first, last = if rows then range.From.Row, range.To.Row else range.From.Col, range.To.Col
            if first > limit then Host.hostFail memberName (sprintf "%s is not on the sheet" (SheetOps.rangeText range))
            first, min limit last
    /// The rows / columns of the selection.
    let selectedSpan (rows: bool) : int * int =
        let top, left, bottom, right = selectionBounds()
        if rows then top, min bottom (model.RowCount - 1) else left, min right (model.ColCount - 1)
    /// Fit columns / rows to what their cells show, as one undo step.
    let autoFitSpan (rows: bool) (first: int, last: int) : bool =
        let sizes =
            if rows then SheetOps.autoFitRowHeights fitMeasure model (first, last)
            else SheetOps.autoFitColumnWidths fitMeasure model (first, last)
        changeAxes (if rows then "AutoFit rows" else "AutoFit columns") (SheetOps.applyAxisSizes rows sizes)
    /// The sizes as they were when a drag of a resize handle began.
    let mutable resizeBefore : (Map<int, float> * Map<int, float> * SheetAxes) option = None

    /// The highest validation id the component has handed out (see lastRuleId).
    let mutable lastValidationId = 0

    /// What GetSelectionFormat last gave, and what it gave it for (the cells,
    /// the column styles and the selected block): asked again for the same,
    /// it answers without looking at a cell.
    let mutable selectionFormatFor : (obj * obj * (int * int * int * int)) option = None
    let mutable selectionFormatObj : obj = null

    /// The model of the sheet at `index` (converting a sheet of an opened
    /// file that was not needed so far).
    let sheetModelAt (index: int) : SpreadsheetModel =
        if index = activeSheetIndex then model
        else
            convertInPlace index
            sheets.[index]

    // --- the list of sheets ---------------------------------------------------------
    // Adding, deleting, renaming, moving, copying, hiding and colouring a sheet
    // are steps of the workbook, not of one sheet, while undo histories are
    // kept per sheet (Undo acts on the sheet on display). Such a step is put
    // on the history of the sheet that is on display once it is done - the new
    // sheet after AddSheet, the sheet that took a deleted one's place - so
    // that Undo right after it takes it back. Undoing or redoing it can change
    // which sheet is on display (undoing AddSheet removes the sheet on
    // display): the step then moves to the history of the sheet that is on
    // display afterwards, so that Redo / Undo find it there.

    /// The steps that change the list of sheets (see above), by identity.
    let bookSteps = ResizeArray<Command>()
    let isBookStep (command: Command) = bookSteps |> Seq.exists (fun known -> obj.ReferenceEquals(known, command))
    /// Every sheet id the workbook has had: a new sheet gets one that was
    /// never used, since undo brings a deleted sheet back under its id.
    let usedSheetIds = System.Collections.Generic.HashSet<SheetId>()
    let sheetsListeners = ResizeArray<obj -> unit>()
    /// What the OnSheetsChange subscribers were last told (the list as one text).
    let mutable notifiedSheetList : string = ""

    /// Put the grid at the top-left corner of the sheet.
    let scrollHome () =
        match gridContainer with
        | Some grid ->
            syncExtent()
            setGridScrollTop grid 0.0
            setGridScrollLeft grid 0.0
        | None -> ()
        scrollTop <- 0.0
        scrollLeft <- 0.0

    /// The workbook with the list of its sheets as the component holds it now.
    let bookSnapshot () : Workbook =
        ensureWorkbook () |> ignore
        currentWorkbookSnapshot ()

    /// Make `wb` the workbook: its sheets in its order with their names, its
    /// active sheet on display. Each sheet keeps its active cell.
    let installBook (wb: Workbook) =
        let origin = { Row = 0; Col = 0 }
        let anchors = System.Collections.Generic.Dictionary<SheetId, CellAddress>()
        for index in 0 .. sheets.Length - 1 do
            anchors.[sheetIdForIndex index] <-
                (if index = activeSheetIndex then anchorCell elif index < sheetAnchors.Length then sheetAnchors.[index] else origin)
        let activeBefore = sheetIdForIndex activeSheetIndex
        let order = wb.SheetOrder |> List.filter (fun id -> Map.containsKey id wb.Sheets) |> List.toArray
        sheets <- order |> Array.map (fun id -> wb.Sheets.[id].Model)
        sheetNames <- order |> Array.map (fun id -> wb.Sheets.[id].Name)
        sheetAnchors <- order |> Array.map (fun id -> match anchors.TryGetValue id with | true, anchor -> anchor | _ -> origin)
        let active = max 0 (System.Array.IndexOf(order, wb.ActiveSheet))
        activeSheetIndex <- active
        workbook <- Some { wb with SheetOrder = List.ofArray order; ActiveSheet = order.[active] }
        for id in order do usedSheetIds.Add id |> ignore
        convertInPlace active
        model <- sheets.[active]
        anchorCell <- sheetAnchors.[active]
        lookedUpCells.Clear()
        if order.[active] <> activeBefore then
            scrollHome ()
            for listener in sheetChangedListeners do
                try listener active with _ -> ()

    /// The formulas of the workbook stand as they are, but the workbook they
    /// were calculated in is another one (sheets renamed, moved, added): the
    /// dependency graph is built again from the workbook as it is now.
    let rebaseCalcEngine () =
        invalidateCalcEngine ()
        calculatedWorkbook <- Some (currentWorkbookSnapshot ())

    /// Run a change of the list of sheets as one undoable step (see above).
    /// `forward` makes the change and `back` takes it back; each reads the
    /// workbook as it is when it runs, installs the changed one and
    /// recalculates what it has to.
    let runBookStep (description: string) (forward: unit -> unit) (back: unit -> unit) =
        // The sheet on display is in `sheets` as it is now.
        sheets.[activeSheetIndex] <- model
        forward ()
        let step = { Undo = (fun m -> back (); m); Redo = (fun m -> forward (); m); Description = description }
        bookSteps.Add step
        if bookSteps.Count > 400 then bookSteps.RemoveAt 0
        model <- pushCommand step model

    /// The sheet with an id as the workbook has it now, or None.
    let sheetOfId (id: SheetId) (wb: Workbook) = Map.tryFind id wb.Sheets

    /// The sheet list as GetSheetInfos gives it.
    let tabColorHex (color: string option) : string option =
        match color with
        | Some value when value.StartsWith "#" -> Some value
        | Some value when value.StartsWith "theme:" ->
            // The colours of the Office theme, for a colour the file names by
            // its place in the theme (the tint is left out).
            let palette = [| "#FFFFFF"; "#000000"; "#E7E6E6"; "#44546A"; "#4472C4"; "#ED7D31"; "#A5A5A5"; "#FFC000"; "#5B9BD5"; "#70AD47" |]
            (match System.Int32.TryParse (value.Split(':').[1]) with
             | true, index when index >= 0 && index < palette.Length -> Some palette.[index]
             | _ -> None)
        | _ -> None

    let sheetInfos () : obj array =
        Array.init sheets.Length (fun index ->
            let id = sheetIdForIndex index
            let sheetModel = if index = activeSheetIndex then model else sheets.[index]
            let visible =
                match workbook |> Option.bind (fun wb -> Map.tryFind id wb.Sheets) with
                | Some sheet -> sheet.Visible
                | None -> true
            createObj [
                "index" ==> index
                "id" ==> id
                "name" ==> sheetNames.[index]
                "visibility" ==> (if visible then "visible" elif sheetModel.VeryHidden then "veryHidden" else "hidden")
                "tabColor" ==> Host.orNull (tabColorHex sheetModel.TabColor)
                "isLoaded" ==> (index = activeSheetIndex || not (isDeferredSheet sheets.[index]))
                "isActive" ==> (index = activeSheetIndex)
            ])

    /// The list of sheets as one text: which sheets, their names, order, visibility and tab colours.
    let sheetListText () : string =
        Array.init sheets.Length (fun index ->
            let id = sheetIdForIndex index
            let sheetModel = if index = activeSheetIndex then model else sheets.[index]
            let visible =
                match workbook |> Option.bind (fun wb -> Map.tryFind id wb.Sheets) with
                | Some sheet -> sheet.Visible
                | None -> true
            id + "\u0001" + sheetNames.[index] + "\u0001" + (if visible then "v" elif sheetModel.VeryHidden then "x" else "h")
            + "\u0001" + defaultArg sheetModel.TabColor "")
        |> String.concat "\u0002"

    /// Tell the OnSheetsChange subscribers when the list of sheets (which
    /// sheets, their names, order, visibility, tab colours) is not what they
    /// were last told.
    let notifySheetList () =
        if sheetsListeners.Count > 0 then
            let text = sheetListText ()
            if text <> notifiedSheetList then
                notifiedSheetList <- text
                let infos = sheetInfos ()
                for listener in sheetsListeners.ToArray() do
                    if sheetsListeners.Contains listener then (try listener (box infos) with _ -> ())

    /// The format a toggle (bold, italic ...) goes by: the active cell's.
    let activeFormat () : CellFormat =
        match getCell (activeCellAddress()) model with
        | Some data -> data.Format
        | None -> defaultFormat

    /// Tell the OnEditStateChange subscribers when the entry in progress
    /// (which cell, what text) is not what they were last told.
    let notifyEditState () =
        let current = model.EditingCell |> Option.map (fun addr -> addr, model.EditBuffer)
        if current <> notifiedEdit then
            notifiedEdit <- current
            if editStateListeners.Count > 0 then
                let state =
                    match current with
                    | Some (addr, text) -> createObj [ "Editing" ==> true; "Row" ==> addr.Row; "Col" ==> addr.Col; "Text" ==> text; "SheetIndex" ==> activeSheetIndex ]
                    | None -> createObj [ "Editing" ==> false; "Row" ==> -1; "Col" ==> -1; "Text" ==> ""; "SheetIndex" ==> activeSheetIndex ]
                for listener in editStateListeners.ToArray() do
                    if editStateListeners.Contains listener then (try listener state with _ -> ())

    /// Start editing a cell. `seed` replaces the cell's content in the edit
    /// box — the character typed over the cell — instead of loading it.
    let rec startEditingWith (row: int) (col: int) (seed: string option) =
        if not readOnly then
            // The first cell of a merged cell is the one that holds its content.
            let addr =
                match findMergedCell { Row = row; Col = col } model with
                | Some merge -> { Row = merge.StartRow; Col = merge.StartCol }
                | None -> { Row = row; Col = col }
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
            // Typing into a cell of the selected range keeps the range (Enter
            // and Tab then walk through it); any other cell becomes the selection.
            let top, left, bottom, right = selectionBounds()
            if selectionIsRange () && addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right then anchorCell <- addr
            else selectCellAddress addr
            model <- { model with EditingCell = Some addr; EditBuffer = currentValue }
            // The edit box sits on the cell: the cell has to be on screen.
            revealCell addr.Row addr.Col
            render()
            match layers with
            | Some layer ->
                setValue layer.EditInput currentValue
                focus layer.EditInput
                if editCaretAtEnd then caretToEnd layer.EditInput else select layer.EditInput
            | None -> ()
            notifyEditState ()

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
            // The keyboard goes back to the grid.
            match layers, contentElement with
            | Some layer, Some content when isActiveElement layer.EditInput -> focusQuietly content
            | _ -> ()
            render()
            notifyEditState ()
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
                        render()) 6000)

    /// Select one cell and bring it into view.
    and goToCell (addr: CellAddress) =
        selectCellAddress addr
        revealCell addr.Row addr.Col
        render()

    /// Keep the active cell and move the other corner of the selection.
    and extendSelectionTo (extent: CellAddress) =
        selectFromTo (activeCellAddress()) extent
        revealCell extent.Row extent.Col
        render()

    /// Move selection
    and moveSelection (deltaRow: int) (deltaCol: int) =
        // A range collapses to its active cell, then moves.
        goToCell (stepFrom (activeCellAddress()) deltaRow deltaCol)

    /// Enter / Tab: with several cells selected the active cell moves through
    /// them (down the columns for Enter, along the rows for Tab) and the
    /// selection stays; with one cell selected the selection moves.
    and moveWithinSelection (deltaRow: int) (deltaCol: int) =
        if not (selectionIsRange ()) then moveSelection deltaRow deltaCol
        else
            let top, left, bottom, right = selectionBounds()
            let rows = bottom - top + 1
            let cols = right - left + 1
            let mutable cell = activeCellAddress()
            // The next cell that is not swallowed by a merged cell (at most once round).
            let mutable steps = 0
            let mutable found = false
            while not found && steps < rows * cols do
                steps <- steps + 1
                let row, col =
                    if deltaCol <> 0 then
                        let col = cell.Col + sign deltaCol
                        if col > right then (if cell.Row >= bottom then top else cell.Row + 1), left
                        elif col < left then (if cell.Row <= top then bottom else cell.Row - 1), right
                        else cell.Row, col
                    else
                        let row = cell.Row + sign deltaRow
                        if row > bottom then top, (if cell.Col >= right then left else cell.Col + 1)
                        elif row < top then bottom, (if cell.Col <= left then right else cell.Col - 1)
                        else row, cell.Col
                cell <- { Row = row; Col = col }
                found <-
                    (getRowHeight row model > 0.0 && getColumnWidth col model > 0.0)
                    && (match findMergedCell cell model with
                        | Some merge -> merge.StartRow = row && merge.StartCol = col
                        | None -> true)
            anchorCell <- cell
            revealCell cell.Row cell.Col
            render()

    /// Copy the selection: the grid's own clipboard (formulas, formats) and
    /// the text of what is shown.
    and copySelection (isCut: bool) : string * string =
        let text = selectionText ()
        let html = selectionHtml ()
        copiedText <- text
        if isCut && not readOnly then
            cutSelection()
        else
            model <- { model with Clipboard = Some (serializeSelection model.Selection model false) }
        text, html

    /// Handle keyboard events
    and handleKeyDown (e: obj) =
        let key = getEventKey e
        let isModifier = eventFlag e "ctrlKey" || eventFlag e "metaKey"
        let shift = eventFlag e "shiftKey"
        let lower = if key.Length = 1 then key.ToLowerInvariant() else key

        match model.EditingCell with
        | Some _ ->
            // Editing mode: the edit box has the keys, but for these.
            match key with
            | "Enter" ->
                preventDefault e
                // A refused entry leaves the selection on its cell.
                if finishEditing true then moveWithinSelection (if shift then -1 else 1) 0
            | "Escape" ->
                preventDefault e
                finishEditing false |> ignore
            | "Tab" ->
                preventDefault e
                if finishEditing true then moveWithinSelection 0 (if shift then -1 else 1)
            | _ -> ()

        | None ->
            // Navigation mode
            let active = activeCellAddress()
            let extent = selectionExtent()
            // Move the active cell, or with Shift the other corner of the selection.
            let travel (target: CellAddress -> CellAddress) =
                preventDefault e
                if shift then extendSelectionTo (target extent) else goToCell (target active)
            if isModifier && (lower = "c" || lower = "x" || lower = "v") then
                // Copy / cut / paste: the browser raises the clipboard event
                // (see Mount), which is where the system clipboard is. Where
                // a browser raises none for an element that is not editable,
                // the same is done through the asynchronous clipboard.
                let seen = clipboardEvents
                setTimeoutMs (fun () ->
                    if clipboardEvents = seen && not disposed && model.EditingCell.IsNone then
                        match lower with
                        | "v" -> pasteFromSystemClipboard () |> ignore
                        | _ ->
                            let text, html = copySelection (lower = "x")
                            if lower = "x" && not readOnly then render()
                            writeSystemClipboard text html |> ignore) 0 |> ignore
            elif isModifier then
                match lower with
                | "z" when not readOnly ->
                    preventDefault e
                    undoStep ()
                | "y" when not readOnly ->
                    preventDefault e
                    redoStep ()
                | "f" ->
                    // Find/Replace dialog
                    preventDefault e
                    showFindReplaceDialog()
                | "d" when not readOnly ->
                    preventDefault e
                    fillSelection true
                    render()
                | "r" when not readOnly ->
                    preventDefault e
                    fillSelection false
                    render()
                | "a" ->
                    // Everything, with the active cell where it is.
                    preventDefault e
                    anchorCell <- active
                    model <- { model with Selection = Range (0, 0, max 0 (model.RowCount - 1), max 0 (model.ColCount - 1)) }
                    render()
                | "ArrowUp" -> travel (fun from -> dataEdge from -1 0)
                | "ArrowDown" -> travel (fun from -> dataEdge from 1 0)
                | "ArrowLeft" -> travel (fun from -> dataEdge from 0 -1)
                | "ArrowRight" -> travel (fun from -> dataEdge from 0 1)
                | "Home" ->
                    // The first cell of the scrolled area (A1 without frozen panes).
                    travel (fun _ -> { Row = max 0 (min (model.RowCount - 1) model.FrozenRows); Col = max 0 (min (model.ColCount - 1) model.FrozenCols) })
                | "End" ->
                    // The last cell that holds anything.
                    travel (fun from ->
                        match usedBounds model with
                        | Some (row, col) -> { Row = min (model.RowCount - 1) row; Col = min (model.ColCount - 1) col }
                        | None -> from)
                | _ -> ()
            else
                match key with
                // Alt+Down opens the list of a list validation, as in Excel.
                | "ArrowDown" when eventFlag e "altKey" && not readOnly && showsListDropdown active model ->
                    preventDefault e
                    showValidationList () |> ignore
                | "ArrowUp" -> travel (fun from -> stepFrom from -1 0)
                | "ArrowDown" -> travel (fun from -> stepFrom from 1 0)
                | "ArrowLeft" -> travel (fun from -> stepFrom from 0 -1)
                | "ArrowRight" -> travel (fun from -> stepFrom from 0 1)
                | "Home" -> travel (fun from -> { from with Col = 0 })
                | "End" ->
                    travel (fun from ->
                        match lastUsedColumnInRow from.Row model with
                        | Some col -> { from with Col = min (model.ColCount - 1) col }
                        | None -> from)
                | "PageDown" -> travel (fun from -> stepFrom from (pageRows ()) 0)
                | "PageUp" -> travel (fun from -> stepFrom from (-(pageRows ())) 0)
                | "Enter" ->
                    preventDefault e
                    if selectionIsRange () || readOnly then moveWithinSelection (if shift then -1 else 1) 0
                    else startEditing active.Row active.Col
                | "F2" ->
                    preventDefault e
                    startEditing active.Row active.Col
                | "Delete" | "Backspace" ->
                    preventDefault e
                    if not readOnly then
                        clearSelection()
                        render()
                | "Tab" ->
                    preventDefault e
                    moveWithinSelection 0 (if shift then -1 else 1)
                | "Escape" ->
                    // Back to the active cell alone.
                    if selectionIsRange () then
                        preventDefault e
                        goToCell active
                | _ ->
                    // Type-to-edit: a printable character starts an entry
                    // that replaces the cell's content, beginning with it.
                    if key.Length = 1 && not (eventFlag e "altKey") then
                        preventDefault e
                        if not readOnly then startEditingWith active.Row active.Col (Some key)

    /// Paste what the system clipboard holds when the page may read it, else
    /// the grid's own clipboard. Resolves to whether anything was pasted.
    and pasteFromSystemClipboard () : JS.Promise<bool> =
        if readOnly then promiseOf false
        else
            let pasteOwn () =
                if model.Clipboard.IsSome then
                    pasteClipboardAtSelection()
                    render()
                    true
                else false
            promiseThen (readSystemClipboard ()) (fun flavours ->
                if isNull (box flavours) || (flavours.[0] = "" && flavours.[1] = "") then pasteOwn ()
                elif pasteFlavours flavours.[0] flavours.[1] then
                    render()
                    true
                else false)

    and undoStep () =
        observeRevision ()
        let command = List.tryHead model.History.Past
        match command with
        | Some step when isBookStep step ->
            // A step of the sheet list: it leaves this sheet's history, is
            // taken back, and waits for Redo on the sheet then on display.
            model <- { model with History = { model.History with Past = List.tail model.History.Past } }
            sheets.[activeSheetIndex] <- model
            step.Undo model |> ignore
            model <- { model with History = { model.History with Future = step :: model.History.Future } }
            observeRevisionAfter 1 command
            render()
        | _ ->
        match undo model with
        | Some undoneModel ->
            model <- undoneModel
            if not (isValueNeutral command) then recalculateFormulas()
            observeRevisionAfter 1 command
            render()
        | None -> ()

    and redoStep () =
        observeRevision ()
        let command = List.tryHead model.History.Future
        match command with
        | Some step when isBookStep step ->
            model <- { model with History = { model.History with Future = List.tail model.History.Future } }
            sheets.[activeSheetIndex] <- model
            step.Redo model |> ignore
            model <- { model with History = { model.History with Past = step :: model.History.Past } }
            observeRevisionAfter 2 command
            render()
        | _ ->
        match redo model with
        | Some redoneModel ->
            model <- redoneModel
            if not (isValueNeutral command) then recalculateFormulas()
            observeRevisionAfter 2 command
            render()
        | None -> ()

    /// Handle input change (while editing)
    and handleInput (e: obj) =
        let target = getEventTarget e
        let value = getValue target
        model <- { model with EditBuffer = value }
        notifyEditState ()

    /// Draw a cell into its node: position, look and text. Selection is not
    /// drawn here (it is an overlay), so moving it draws no cell.
    and drawCell (layer: GridLayers) (node: CellNode) (row: int) (col: int) (x: float) (y: float) (w: float) (h: float) (pane: int)
                 (cellData: CellData option) (tables: Rbt.Spreadsheet.Workbook.Table list) =
        let cell = node.Element
        // What the cell shows under the sheet's view options (formulas as
        // text, no zeros); `storedData` is the cell itself.
        let storedData = cellData
        let cellData =
            if model.ShowZeros && not model.ShowFormulas then cellData
            else
                cellData |> Option.map (fun data ->
                    let shown = viewedValue model data.Value
                    if obj.ReferenceEquals(shown, data.Value) then data else { data with Value = shown })
        if node.Pane <> pane || not (isChildOf cell layer.Panes.[pane]) then appendChild layer.Panes.[pane] cell
        let addr = { Row = row; Col = col }
        // What every cell has in common is in the grid's style sheet (see
        // Mount); the node's own style is its place and what its format adds.
        let css = System.Text.StringBuilder()
        let px (value: float) = string (int (value + 0.5)) + "px"
        let add (name: string) (value: string) = css.Append(name).Append(':').Append(value).Append(';') |> ignore
        add "left" (px x)
        add "top" (px y)
        add "width" (px w)
        add "height" (px h)
        add "line-height" (px (h - 4.0))
        // The colour of the cell's grid lines (a table has its own).
        let mutable gridColor = defaultGridColor
        // Table styling first (base styles).
        let tableInfo =
            tables
            |> List.tryPick (fun table ->
                if row >= table.StartRow && row <= table.EndRow && col >= table.StartCol && col <= table.EndCol then
                    Some (table.HasHeaders && row = table.StartRow, (row - table.StartRow - (if table.HasHeaders then 1 else 0)) % 2 = 0)
                else None)
        match tableInfo with
        | Some (true, _) ->
            add "background" "linear-gradient(180deg, #4472C4 0%, #2E5AA8 100%)"
            add "color" "white"
            add "font-weight" "bold"
            add "text-align" "center"
            gridColor <- "#2E5AA8"
        | Some (false, isEvenRow) ->
            add "background-color" (if isEvenRow then "#FFFFFF" else "#EDF2FA")
            gridColor <- "#D0D8E4"
        | None -> ()
        // The cell's own format (over the table's).
        match cellData with
        | Some data ->
            let format = data.Format
            if format.Bold then add "font-weight" "bold"
            if format.Italic then add "font-style" "italic"
            if format.Underline || format.UnderlineDouble || format.Strikethrough then
                let lines =
                    (if format.Underline || format.UnderlineDouble then "underline " else "")
                    + (if format.Strikethrough then "line-through " else "")
                    + (if format.UnderlineDouble then "double" else "")
                add "text-decoration" (lines.TrimEnd())
            match format.FontName with
            // (A name is all a family can be here: nothing that ends the declaration.)
            | Some family when family <> "" -> add "font-family" ("'" + family.Replace("'", "").Replace(";", "").Replace("\\", "") + "', Arial, sans-serif")
            | _ -> ()
            match format.TextColor with Some color -> add "color" color | None -> ()
            match format.BgColor with Some color -> add "background-color" color | None -> ()
            match format.FontSize with Some size -> add "font-size" (string size + "px") | None -> ()
            match format.HorizontalAlign with
            | Some "centerContinuous" -> add "text-align" "center"
            | Some "distributed" -> add "text-align" "justify"
            | Some "fill" | Some "general" -> ()
            | Some align -> add "text-align" align
            | None -> ()
            // Where the text sits in the height of the cell. One line of text
            // is centred by the cell's line height (the node's default look);
            // top and bottom give the line its own height and, for bottom,
            // push it down. Wrapped text is a column of lines, placed as a
            // block (at the bottom unless the format says otherwise, as Excel does).
            let lineHeight = System.Math.Round (float (defaultArg format.FontSize 11) * 1.3)
            if format.WrapText then
                add "white-space" "normal"
                add "word-wrap" "break-word"
                add "line-height" (px lineHeight)
                add "display" "flex"
                add "flex-direction" "column"
                add "justify-content"
                    (match format.VerticalAlign with
                     | Some "top" -> "flex-start"
                     | Some "middle" | Some "center" -> "center"
                     | _ -> "flex-end")
            else
                match format.VerticalAlign with
                | Some "top" -> add "line-height" (px lineHeight)
                | Some "bottom" ->
                    add "line-height" (px lineHeight)
                    add "padding-top" (px (max 3.0 (h - 3.0 - lineHeight)))
                | _ -> ()
            // An indent pads the edge the text is aligned to.
            if format.Indent > 0 then
                let indent = float format.Indent * SpreadsheetCellText.mdwPx (float (defaultArg format.FontSize 11)) * 1.4
                if format.HorizontalAlign = Some "right" then add "padding-right" (px (5.0 + indent))
                elif format.HorizontalAlign <> Some "center" then add "padding-left" (px (5.0 + indent))
        | None -> ()
        // Borders and grid lines: one line per edge. A cell draws the edge
        // above it and the edge to its left - each the one border the two
        // cells that share it resolve to (Layout.edgeAbove / edgeLeft), or
        // their grid line - inside its own box; the edges below and to its
        // right are the next cells'. (A cell of the last row / column draws
        // those too.) Grid lines are an inset shadow of the cell; a border is
        // a child element over it, the stronger of the two on top where they
        // meet, so a line runs on through the corners of the cells it passes.
        let formatAt (r: int) (c: int) =
            if r = row && c = col then cellData |> Option.map (fun data -> data.Format)
            else getCell { Row = r; Col = c } model |> Option.map (fun data -> data.Format)
        let endRow, endCol =
            if model.MergedCells.IsEmpty then row, col
            else
                match findMergedCell addr model with
                | Some merge -> merge.EndRow, merge.EndCol
                | None -> row, col
        let edgeChildren = ResizeArray<int * string>()
        let borderWidth (border: BorderStyle) = match borderCssParts border.Style with Some (width, _) -> width | None -> 0
        // The runs of one kind along an edge of the cell: (offset, length, kind).
        let runsAlong (first: int) (last: int) (startOf: int -> float) (kindAt: int -> SpreadsheetLayout.SharedEdge) =
            if first = last then [ 0.0, (startOf (last + 1) - startOf first), kindAt first ]
            else
                let runs = ResizeArray<float * float * SpreadsheetLayout.SharedEdge>()
                let origin = startOf first
                for index in first .. last do
                    let size = startOf (index + 1) - startOf index
                    if size > 0.0 then
                        let kind = kindAt index
                        let offset = startOf index - origin
                        if runs.Count > 0 && (let (_, _, previous) = runs.[runs.Count - 1] in previous = kind) then
                            let (start, length, _) = runs.[runs.Count - 1]
                            runs.[runs.Count - 1] <- (start, length + size, kind)
                        else runs.Add (offset, size, kind)
                List.ofSeq runs
        let shadows = ResizeArray<string>()
        // side: 0 top, 1 left, 2 bottom, 3 right.
        let drawEdge (side: int) (runs: (float * float * SpreadsheetLayout.SharedEdge) list) =
            let gridShadow =
                match side with
                | 0 -> "inset 0 1px 0 0 "
                | 1 -> "inset 1px 0 0 0 "
                | 2 -> "inset 0 -1px 0 0 "
                | _ -> "inset -1px 0 0 0 "
            let place (offset: float) (length: float) (thickness: int) (line: string) =
                let o = string (int (offset + 0.5)) + "px"
                let l = string (int (length + 0.5)) + "px"
                match side with
                | 0 -> "position:absolute;pointer-events:none;left:" + o + ";top:0;width:" + l + ";height:0;border-top:" + line
                | 1 -> "position:absolute;pointer-events:none;left:0;top:" + o + ";width:0;height:" + l + ";border-left:" + line
                | 2 -> "position:absolute;pointer-events:none;left:" + o + ";bottom:0;width:" + l + ";height:0;border-bottom:" + line
                | _ -> "position:absolute;pointer-events:none;right:0;top:" + o + ";width:0;height:" + l + ";border-right:" + line
            match runs with
            | [ (_, _, SpreadsheetLayout.EdgeGrid true) ] -> shadows.Add (gridShadow + gridColor)
            | _ ->
                for offset, length, kind in runs do
                    match kind with
                    | SpreadsheetLayout.EdgeBorder border -> edgeChildren.Add (borderPriority border.Style, place offset length (borderWidth border) (borderCss border))
                    | SpreadsheetLayout.EdgeGrid true -> edgeChildren.Add (0, place offset length 1 ("1px solid " + gridColor))
                    | _ -> ()
        let columnStart (c: int) = SpreadsheetLayout.columnLeft model c
        let rowStart (r: int) = SpreadsheetLayout.rowTop model r
        // The cells the edges are shared with: the one above, the one to the
        // left, and the one above that (for the corner between them).
        let above = SpreadsheetLayout.rowShownBefore model row
        let before = SpreadsheetLayout.columnShownBefore model col
        let aboveData = cellAt above col
        let leftData = cellAt row before
        let cornerData = cellAt above before
        let plainSheet = model.MergedCells.IsEmpty
        // (Without merged cells an edge follows from the two cells' formats: no further look-ups.)
        let edgeBetween (first: CellData option) (second: CellData option) (horizontal: bool) (hasFirst: bool) =
            let border (data: CellData option) (pick: CellFormat -> BorderStyle option) = data |> Option.bind (fun data -> pick data.Format)
            let filled (data: CellData option) = match data with Some data -> data.Format.BgColor.IsSome | None -> false
            let stated =
                resolveSharedBorder
                    (if hasFirst then border first (fun format -> if horizontal then format.BorderBottom else format.BorderRight) else None)
                    (border second (fun format -> if horizontal then format.BorderTop else format.BorderLeft))
            match stated with
            | Some line -> SpreadsheetLayout.EdgeBorder line
            | None -> SpreadsheetLayout.EdgeGrid (model.ShowGridLines && not (hasFirst && filled first) && not (filled second))
        let topRuns =
            if plainSheet then [ 0.0, w, edgeBetween aboveData cellData true (above >= 0) ]
            else runsAlong col endCol columnStart (fun c -> SpreadsheetLayout.edgeAbove model row c)
        let leftRuns =
            if plainSheet then [ 0.0, h, edgeBetween leftData cellData false (before >= 0) ]
            else runsAlong row endRow rowStart (fun r -> SpreadsheetLayout.edgeLeft model r col)
        // The two edges that end at the cell's top-left corner without being its own:
        // the one coming down from above and the one coming in from the left.
        let noEdge = SpreadsheetLayout.EdgeGrid false
        let verticalEdge =
            if above < 0 || before < 0 then noEdge
            elif plainSheet then (if aboveData.IsSome || cornerData.IsSome then edgeBetween cornerData aboveData false true else noEdge)
            else SpreadsheetLayout.edgeLeft model above col
        let horizontalEdge =
            if above < 0 || before < 0 then noEdge
            elif plainSheet then (if leftData.IsSome || cornerData.IsSome then edgeBetween cornerData leftData true true else noEdge)
            else SpreadsheetLayout.edgeAbove model row before
        // A double border turning a corner: the outer rule goes round the outside, the
        // inner one round the inside. Two children crossing (or a square) do not draw
        // that - the outer line of an outline was two pixels off at its right corners -
        // so the corner is one 3 x 3 element with the two sides it joins (CSS mitres a
        // double border on one element), and the cell's own edges start after it.
        let doubleOf (edge: SpreadsheetLayout.SharedEdge) =
            match edge with
            | SpreadsheetLayout.EdgeBorder border when border.Style.ToLowerInvariant() = "double" -> Some border
            | _ -> None
        let isLine (edge: SpreadsheetLayout.SharedEdge) = match edge with SpreadsheetLayout.EdgeBorder _ -> true | _ -> false
        let firstOf (runs: (float * float * SpreadsheetLayout.SharedEdge) list) = match runs with (_, _, edge) :: _ -> edge | [] -> noEdge
        let ownTop, ownLeft = firstOf topRuns, firstOf leftRuns
        let doubleCorner =
            match doubleOf ownTop, doubleOf ownLeft, doubleOf horizontalEdge, doubleOf verticalEdge with
            | Some line, Some _, None, None when not (isLine horizontalEdge) && not (isLine verticalEdge) -> Some (line, "border-top", "border-left", true, true)
            | None, Some line, Some _, None when not (isLine ownTop) && not (isLine verticalEdge) -> Some (line, "border-top", "border-right", false, true)
            | Some line, None, None, Some _ when not (isLine ownLeft) && not (isLine horizontalEdge) -> Some (line, "border-bottom", "border-left", true, false)
            | None, None, Some line, Some _ when not (isLine ownTop) && not (isLine ownLeft) -> Some (line, "border-bottom", "border-right", false, false)
            | _ -> None
        let trim (cut: bool) (runs: (float * float * SpreadsheetLayout.SharedEdge) list) =
            match runs with
            | (offset, length, edge) :: rest when cut && length > 3.0 -> (offset + 3.0, length - 3.0, edge) :: rest
            | _ -> runs
        match doubleCorner with
        | Some (line, sideA, sideB, trimTop, trimLeft) ->
            let css = borderCss line
            edgeChildren.Add (100, "position:absolute;pointer-events:none;left:0;top:0;width:0;height:0;" + sideA + ":" + css + ";" + sideB + ":" + css)
            drawEdge 0 (trim trimTop topRuns)
            drawEdge 1 (trim trimLeft leftRuns)
        | None ->
            drawEdge 0 topRuns
            drawEdge 1 leftRuns
        if endRow >= model.RowCount - 1 then drawEdge 2 (runsAlong col endCol columnStart (fun c -> SpreadsheetLayout.edgeBelow model endRow c))
        if endCol >= model.ColCount - 1 then drawEdge 3 (runsAlong row endRow rowStart (fun r -> SpreadsheetLayout.edgeRight model r endCol))
        // Where a border from above ends at one coming from the left (the
        // far corner of an outline), neither is this cell's own edge: the
        // corner between them is filled, unless an own edge covers it.
        if doubleCorner.IsNone then
            match verticalEdge, horizontalEdge with
            | SpreadsheetLayout.EdgeBorder vertical, SpreadsheetLayout.EdgeBorder horizontal ->
                let cornerWidth = borderWidth vertical
                let cornerHeight = borderWidth horizontal
                let covered =
                    (match topRuns with
                     | (_, _, SpreadsheetLayout.EdgeBorder own) :: _ -> borderWidth own >= cornerHeight
                     | _ -> false)
                    || (match leftRuns with
                        | (_, _, SpreadsheetLayout.EdgeBorder own) :: _ -> borderWidth own >= cornerWidth
                        | _ -> false)
                if not covered && cornerWidth > 0 && cornerHeight > 0 then
                    let stronger = if borderPriority vertical.Style > borderPriority horizontal.Style then vertical else horizontal
                    edgeChildren.Add (100, "position:absolute;pointer-events:none;left:0;top:0;width:" + string cornerWidth + "px;height:" + string cornerHeight
                                           + "px;background:" + (stronger.Color |> Option.defaultValue "#000000"))
            | _ -> ()
        // (The style sheet gives a cell the two grid lines most cells have.)
        if not (shadows.Count = 2 && gridColor = defaultGridColor && shadows.[0].StartsWith "inset 0 1px" && shadows.[1].StartsWith "inset 1px") then
            add "box-shadow" (if shadows.Count = 0 then "none" else String.concat "," shadows)
        // Conditional formatting (highest priority): the merged effect of
        // every matching rule that covers the cell.
        let mutable icon : (int * int) option = None
        let mutable bar = -1
        if not model.ConditionalFormats.IsEmpty then
            match Rbt.Spreadsheet.ConditionalFormatting.effectAt (Rbt.Spreadsheet.ConditionalFormatting.contextFor model) addr with
            | Some effect ->
                if effect.Bold = Some true then add "font-weight" "bold"
                if effect.Italic = Some true then add "font-style" "italic"
                let decorations =
                    [ if effect.Underline = Some true then yield "underline"
                      if effect.Strikethrough = Some true then yield "line-through" ]
                if not decorations.IsEmpty then add "text-decoration" (String.concat " " decorations)
                match effect.TextColor with Some color -> add "color" color | None -> ()
                match effect.BgColor with Some color -> add "background-color" color | None -> ()
                match effect.DataBar with
                | Some (fraction, color) ->
                    // A bar from the left edge, behind the text.
                    let percent = int (System.Math.Round (fraction * 100.0))
                    add "background-image" ("linear-gradient(to right, " + color + " " + string percent + "%, transparent " + string percent + "%)")
                    bar <- percent
                | None -> ()
                match effect.IconSet with
                | Some (index, count) when count > 1 -> icon <- Some (index, count)
                | _ -> ()
            | None -> ()
        // The text. A number that does not fit its cell shows ####, as in
        // Excel (it is never cut off to something that reads as another number).
        let fontPx = match cellData with Some data -> float (defaultArg data.Format.FontSize 11) | None -> 11.0
        let measureText (text: string) =
            match cellData with
            | Some data -> cellTextWidth text fontPx data.Format.Bold data.Format.Italic
            | None -> 0.0
        let displayText =
            match cellData with
            | Some data ->
                match unwrapFormulaValue data.Value with
                | CellValue.Number _ ->
                    let text = SpreadsheetCellText.cellDisplayText None 1.0e6 data.Format data.Value
                    // (Shrunk or turned text is not measured against the cell's width.)
                    if data.Format.ShrinkToFit || data.Format.TextRotation <> 0 then text
                    elif float text.Length * fontPx * 0.7 <= w - 10.0 || measureText text <= w - 9.0 then text
                    else
                        let hash = max 1.0 (measureText "#")
                        System.String('#', max 1 (int ((w - 10.0) / hash)))
                | _ -> SpreadsheetCellText.cellDisplayText None 1.0e6 data.Format data.Value
            | None -> ""
        // Shrink to fit: one line of text wider than its cell is drawn smaller
        // (never larger), down to what fits between the cell's paddings.
        match cellData with
        | Some data when data.Format.ShrinkToFit && not data.Format.WrapText && data.Format.TextRotation = 0 && displayText <> "" ->
            let indent = if data.Format.Indent > 0 then float data.Format.Indent * SpreadsheetCellText.mdwPx fontPx * 1.4 else 0.0
            let room = w - 10.0 - indent
            let width = measureText displayText
            if room > 0.0 && width > room then
                add "font-size" (string (System.Math.Floor (max 1.0 (fontPx * room / width * 0.98) * 100.0) / 100.0) + "px")
        | _ -> ()
        // Turned text: a line through the middle of the cell at the format's
        // angle (stacked text: upright letters one under the other).
        let rotation = match cellData with Some data when displayText <> "" -> data.Format.TextRotation | _ -> 0
        // Text overflow: text wider than its cell runs over the empty cells
        // beside it (to the right when left-aligned, to the left when
        // right-aligned, both ways when centred) up to the first cell that
        // holds something. The node keeps its place; its text is let out of
        // its box and clipped at the run of cells it may cover.
        let mutable wide = false
        let mutable crossed : (float * int) list = []
        match cellData with
        | Some data when displayText <> "" && endRow = row && endCol = col && icon.IsNone ->
            match SpreadsheetLayout.textOverflowSides data.Format data.Value with
            // (Not measured when even the widest glyphs would fit the cell.)
            | Some (toLeft, toRight) when float displayText.Length * fontPx * 1.05 > w - 10.0 ->
                let excess = measureText displayText - (w - 10.0)
                if excess > 0.0 then
                    wide <- true
                    // (Past the cell's edge: its own padding on that side is its to use.)
                    let needLeft = if toLeft then (if toRight then excess / 2.0 else excess) - 3.0 else 0.0
                    let needRight = if toRight then (if toLeft then excess / 2.0 else excess) - 3.0 else 0.0
                    let first, last = SpreadsheetLayout.textOverflowSpan model row col needLeft needRight
                    // Text of a frozen column runs on into the columns beside the frozen
                    // ones only while they are there: scrolled away, it is cut at the frozen edge.
                    let last = if col < model.FrozenCols && scrollLeft > 0.5 then min last (model.FrozenCols - 1) else last
                    let first = if col >= model.FrozenCols then max first model.FrozenCols else first
                    if first <> col || last <> col then
                        let own = SpreadsheetLayout.columnLeft model col
                        let extraLeft = own - SpreadsheetLayout.columnLeft model first
                        let extraRight = SpreadsheetLayout.columnLeft model (last + 1) - SpreadsheetLayout.columnLeft model (col + 1)
                        add "overflow" "visible"
                        add "text-overflow" "clip"
                        add "contain" "layout style"
                        add "z-index" "1"
                        add "clip-path" ("inset(0px " + string (-(int extraRight)) + "px 0px " + string (-(int extraLeft)) + "px)")
                        if toLeft then
                            add "display" "flex"
                            add "justify-content" (if toRight then "center" else "flex-end")
                        // The grid lines the text crosses are not shown: each is the
                        // left grid line of the cell on its right (see the edges above).
                        // Borders stay, and so do the lines above and below the row.
                        crossed <-
                            [ for c in first .. last - 1 do
                                // (A hidden column has no line of its own.)
                                if SpreadsheetLayout.columnLeft model (c + 2) > SpreadsheetLayout.columnLeft model (c + 1) then
                                    match SpreadsheetLayout.edgeLeft model row (c + 1) with
                                    | SpreadsheetLayout.EdgeGrid true ->
                                        let topLine =
                                            match SpreadsheetLayout.edgeAbove model row (c + 1) with
                                            | SpreadsheetLayout.EdgeBorder border -> (match borderCssParts border.Style with Some (width, _) -> width | None -> 0)
                                            | SpreadsheetLayout.EdgeGrid true -> 1
                                            | _ -> 0
                                        yield SpreadsheetLayout.columnLeft model (c + 1) - own, topLine
                                    | _ -> () ]
            | _ -> ()
        | _ -> ()
        node.Wide <- wide
        setCssText cell (css.ToString())
        setAttribute cell "data-row" (string row)
        setAttribute cell "data-col" (string col)
        if bar >= 0 then setAttribute cell "data-cf-bar" (string bar) else removeAttribute cell "data-cf-bar"
        // (This also drops the icon and comment marker of the node's previous cell.)
        if rotation = 0 then setTextContent cell displayText
        else
            setTextContent cell ""
            let turned = createElement "span"
            setClassName turned "spreadsheet-rotated"
            setTextContent turned displayText
            let transform =
                if rotation = 255 then "writing-mode:vertical-rl;text-orientation:upright;letter-spacing:-2px;transform:translate(-50%,-50%)"
                else "transform:translate(-50%,-50%) rotate(" + string (-(Host.rotationDegrees rotation)) + "deg)"
            setCssText turned ("position:absolute;left:50%;top:50%;white-space:nowrap;line-height:normal;text-decoration:inherit;pointer-events:none;" + transform)
            appendChild cell turned
        // The weaker edge first: where two meet, the stronger is on top.
        if edgeChildren.Count > 0 then
            for _, style in edgeChildren |> Seq.sortBy fst do
                let line = createElement "div"
                setClassName line "spreadsheet-edge"
                setCssText line style
                appendChild cell line
        for edge, topLine in crossed do
            // One pixel wide, where that grid line is, from below the line above the row:
            // under the text and over the cell the line belongs to. (A grid line is only
            // there between two cells without a fill: what shows instead is their white.)
            let strip = createElement "div"
            setClassName strip "spreadsheet-overflow-gap"
            setCssText strip ("position:absolute;left:" + string (int (floor (edge + 0.5))) + "px;top:" + string topLine + "px;bottom:0;width:1px;background:white;z-index:-1;pointer-events:none")
            appendChild cell strip
        match icon with
        | Some (index, count) ->
            // Same low → high signal the canvas painter draws: red, amber, green.
            let t = float index / float (count - 1)
            let color = if t < 0.34 then "#E53935" elif t < 0.67 then "#F9A825" else "#43A047"
            let mark = createElement "span"
            setTextContent mark "●"
            setCssText mark ("color:" + color + ";margin-right:3px;font-size:10px")
            setAttribute mark "data-cf-icon" (string index)
            prependChild cell mark
        | None -> ()
        // A marker and a tooltip for a comment or a comment thread.
        let thread = if Map.isEmpty model.Threads then None else Map.tryFind addr model.Threads
        let marker (color: string) (tooltip: string) =
            let indicator = createElement "div"
            setCssText indicator ("position:absolute;top:0;right:0;width:0;height:0;border-style:solid;border-width:0 6px 6px 0;border-color:transparent " + color + " transparent transparent;pointer-events:none")
            appendChild cell indicator
            setAttribute cell "title" tooltip
            node.HasTitle <- true
        match thread, cellData with
        | Some thread, _ -> marker "#7b61ff" (threadText thread)
        | None, Some { Comment = Some comment } ->
            marker "#ff0000" (match comment.Author with
                              | Some author -> author + ":\n" + comment.Text
                              | None -> comment.Text)
        | _ ->
            if node.HasTitle then
                removeAttribute cell "title"
                node.HasTitle <- false
        node.Pane <- pane
        node.Data <- (match storedData with Some data -> box data | None -> null)
        let near (data: CellData option) = match data with Some data -> box data | None -> null
        node.NearAbove <- near aboveData
        node.NearLeft <- near leftData
        node.NearCorner <- near cornerData
        node.X <- x
        node.Y <- y
        node.W <- w
        node.H <- h
        node.Epoch <- styleEpoch

    /// A new column header node: label (click sorts), filter button, resize handle.
    and newColumnHeader () : HeaderNode =
        let header = createElement "div"
        setClassName header "spreadsheet-col-header"
        let label = createElement "div"
        setClassName label "col-header-text"
        setCssText label "cursor:pointer"
        let button = createElement "div"
        setClassName button "col-filter-btn"
        setTextContent button "▼"
        setAttribute button "title" "Filter"
        let handle = createElement "div"
        setClassName handle "col-resize-handle"
        setClassName handle "col-resize-handle"
        appendChild header label
        appendChild header button
        appendChild header handle
        { Element = header; Label = label; Button = Some button; Handle = handle; Index = -1; Signature = ""; Seen = 0 }

    and newRowHeader () : HeaderNode =
        let header = createElement "div"
        setClassName header "spreadsheet-row-header"
        let label = createElement "span"
        let handle = createElement "div"
        setClassName handle "row-resize-handle"
        setClassName handle "row-resize-handle"
        appendChild header label
        appendChild header handle
        { Element = header; Label = label; Button = None; Handle = handle; Index = -1; Signature = ""; Seen = 0 }

    and drawColumnHeader (node: HeaderNode) (col: int) (x: float) (w: float) =
        let sort =
            match model.SortState with
            | Some state when state.Column = col -> if state.Ascending then " ▲" else " ▼"
            | _ -> ""
        let filtered = Map.containsKey col model.Filters
        let signature = string col + "|" + string x + "|" + string w + "|" + sort + "|" + string filtered + "|" + string readOnly
        if node.Signature <> signature then
            node.Signature <- signature
            node.Index <- col
            setAttribute node.Element "data-col" (string col)
            setCssText node.Element ("left:" + string (int (x + 0.5)) + "px;width:" + string (int (w + 0.5)) + "px")
            setTextContent node.Label (columnToLetter col + sort)
            setStyle node.Label "cursor" (if readOnly then "default" else "pointer")
            match node.Button with
            | Some button ->
                setCssText button ("opacity:" + (if filtered then "1.0" else "0.5") + ";display:" + (if readOnly then "none" else "flex"))
            | None -> ()
            setStyle node.Handle "display" (if readOnly then "none" else "block")

    and drawRowHeader (node: HeaderNode) (row: int) (y: float) (h: float) =
        let signature = string row + "|" + string y + "|" + string h + "|" + string readOnly
        if node.Signature <> signature then
            node.Signature <- signature
            node.Index <- row
            setAttribute node.Element "data-row" (string row)
            setCssText node.Element ("top:" + string (int (y + 0.5)) + "px;height:" + string (int (h + 0.5)) + "px")
            setTextContent node.Label (string (row + 1))
            setStyle node.Handle "display" (if readOnly then "none" else "block")

    /// Render the grid: bring what is on display in line with the model and
    /// the scroll position. Nodes of cells that stay in view and whose data
    /// did not change are left alone; nodes of cells that left the view are
    /// reused for the ones that came in.
    and render () =
        match contentElement, layers with
        | Some content, Some layer ->
            if not model.ConditionalFormats.IsEmpty then refreshRuleWorkbook()
            syncExtent()
            let metrics = SpreadsheetLayout.metricsOf model
            let rowPrefix = metrics.RowPrefix
            let colPrefix = metrics.ColPrefix
            let rowCount = metrics.RowCount
            let colCount = metrics.ColCount
            let frozenRows = max 0 (min rowCount model.FrozenRows)
            let frozenCols = max 0 (min colCount model.FrozenCols)
            let frozenH = rowPrefix.[frozenRows]
            let frozenW = colPrefix.[frozenCols]

            // Everything is drawn again when what the look of a cell depends
            // on, besides its own data, changed. With conditional formats a
            // cell's look depends on other cells.
            let activeSheetId = sheetIdForIndex activeSheetIndex
            let tables =
                match workbook with
                | Some wb -> wb.Tables |> Map.toList |> List.map snd |> List.filter (fun table -> table.SheetId = activeSheetId)
                | None -> []
            let styleInputs : obj array =
                [| box model.ConditionalFormats; box model.Threads; box model.MergedCells; box model.ShowGridLines
                   box model.ShowHeaders; box model.ShowFormulas; box model.ShowZeros
                   box (match workbook with Some wb -> box wb.Tables | None -> null); box activeSheetIndex; box readOnly
                   (if model.ConditionalFormats.IsEmpty then null else box model.Cells)
                   // (Text of a frozen column is cut at the frozen edge once the columns beside it are scrolled away.)
                   box (model.FrozenCols > 0 && scrollLeft > 0.5) |]
            // (Compared by identity: each is a value the model replaces when it changes.)
            let sameInputs =
                styleInputs.Length = renderedStyleInputs.Length
                && Array.forall2 (fun (a: obj) (b: obj) -> obj.ReferenceEquals(a, b)) styleInputs renderedStyleInputs
            if not sameInputs then
                styleEpoch <- styleEpoch + 1
                renderedStyleInputs <- styleInputs
            let sameCells = obj.ReferenceEquals(renderedCells, model.Cells)
            renderedCells <- box model.Cells

            renderGeneration <- renderGeneration + 1
            let generation = renderGeneration
            let hasMerges = not model.MergedCells.IsEmpty

            // The rows and columns in view, and one more on each side.
            let viewTop = scrollTop + frozenH
            let viewBottom = scrollTop + containerHeight - headerHeight
            let viewLeft = scrollLeft + frozenW
            let viewRight = scrollLeft + containerWidth - rowHeaderWidth
            let firstRow = if rowCount = 0 then 0 else max frozenRows (SpreadsheetLayout.indexAtOffset rowPrefix rowCount viewTop - 1)
            let lastRow = if rowCount = 0 then -1 else min (rowCount - 1) (SpreadsheetLayout.indexAtOffset rowPrefix rowCount viewBottom + 1)
            let firstCol = if colCount = 0 then 0 else max frozenCols (SpreadsheetLayout.indexAtOffset colPrefix colCount viewLeft - 1)
            let lastCol = if colCount = 0 then -1 else min (colCount - 1) (SpreadsheetLayout.indexAtOffset colPrefix colCount viewRight + 1)

            let visit (row: int) (col: int) =
                let mutable r = row
                let mutable c = col
                let mutable w = colPrefix.[col + 1] - colPrefix.[col]
                let mutable h = rowPrefix.[row + 1] - rowPrefix.[row]
                // A merged cell's look also depends on its other cells (their borders).
                let mutable merged = false
                if hasMerges then
                    match findMergedCell { Row = row; Col = col } model with
                    | Some merge ->
                        merged <- true
                        // A merged cell is drawn once, as its first cell, over its whole span.
                        r <- max 0 (min (rowCount - 1) merge.StartRow)
                        c <- max 0 (min (colCount - 1) merge.StartCol)
                        w <- colPrefix.[min colCount (merge.EndCol + 1)] - colPrefix.[c]
                        h <- rowPrefix.[min rowCount (merge.EndRow + 1)] - rowPrefix.[r]
                    | None -> ()
                if w > 0.0 && h > 0.0 then
                    let key = float r * 16384.0 + float c
                    let x = rowHeaderWidth + colPrefix.[c]
                    let y = headerHeight + rowPrefix.[r]
                    let pane = (if r < frozenRows then 1 else 0) + (if c < frozenCols then 2 else 0)
                    match cellNodes.TryGetValue key with
                    | true, node ->
                        if node.Seen <> generation then
                            node.Seen <- generation
                            let placed = node.Epoch = styleEpoch && node.X = x && node.Y = y && node.W = w && node.H = h && node.Pane = pane
                            if not (placed && sameCells) then
                                let data = cellAt r c
                                let dataRef = match data with Some data -> box data | None -> null
                                // (The edges the node draws are shared with the cells above and to its left.)
                                let nearChanged () =
                                    let mutable above = r - 1
                                    while above >= 0 && rowPrefix.[above + 1] <= rowPrefix.[above] do above <- above - 1
                                    let mutable before = c - 1
                                    while before >= 0 && colPrefix.[before + 1] <= colPrefix.[before] do before <- before - 1
                                    let near (nr: int) (nc: int) = match cellAt nr nc with Some data -> box data | None -> null
                                    not (obj.ReferenceEquals(node.NearAbove, near above c)) || not (obj.ReferenceEquals(node.NearLeft, near r before))
                                    || not (obj.ReferenceEquals(node.NearCorner, near above before))
                                if not placed || merged || node.Wide || not (obj.ReferenceEquals(node.Data, dataRef)) || nearChanged () then
                                    drawCell layer node r c x y w h pane data tables
                    | false, _ ->
                        let node =
                            if freeCellNodes.Count > 0 then
                                let last = freeCellNodes.[freeCellNodes.Count - 1]
                                freeCellNodes.RemoveAt (freeCellNodes.Count - 1)
                                last
                            else
                                let element = createElement "div"
                                setClassName element "spreadsheet-cell"
                                { Element = element; Key = key; Pane = -1; Data = null; NearAbove = null; NearLeft = null; NearCorner = null; X = 0.0; Y = 0.0; W = 0.0; H = 0.0; Epoch = -1; Seen = 0; HasTitle = false; Wide = false }
                        node.Key <- key
                        node.Seen <- generation
                        cellNodes.[key] <- node
                        drawCell layer node r c x y w h pane (cellAt r c) tables
            let visitBlock (r0: int) (r1: int) (c0: int) (c1: int) =
                for row = r0 to r1 do
                    if rowPrefix.[row + 1] > rowPrefix.[row] then
                        for col = c0 to c1 do
                            if colPrefix.[col + 1] > colPrefix.[col] then visit row col
            visitBlock firstRow lastRow firstCol lastCol
            // Text of a cell just outside the view may run into it: the
            // nearest filled cell on each side of a row is drawn too, when
            // it holds text. (Looked up once per row while the cells and the
            // columns in view stay the same.)
            if not (obj.ReferenceEquals(overflowCells, box model.Cells)) || overflowFirstCol <> firstCol || overflowLastCol <> lastCol then
                overflowCells <- box model.Cells
                overflowFirstCol <- firstCol
                overflowLastCol <- lastCol
                overflowSources.Clear()
            let reach = 24
            let visitOutside (r0: int) (r1: int) =
                for row = r0 to r1 do
                    if rowPrefix.[row + 1] > rowPrefix.[row] then
                        let sources =
                            match overflowSources.TryGetValue row with
                            | true, known -> known
                            | false, _ ->
                                let isText (data: CellData) = match unwrapFormulaValue data.Value with CellValue.String text -> text <> "" | _ -> false
                                let find (start: int) (step: int) =
                                    let mutable col = start
                                    let mutable found = -1
                                    let mutable steps = 0
                                    while found = -1 && steps < reach && col >= frozenCols && col < colCount do
                                        match getCell { Row = row; Col = col } model with
                                        | Some data when (match data.Value with CellValue.Empty | CellValue.String "" -> false | _ -> true) ->
                                            found <- if isText data then col else -2
                                        | _ -> ()
                                        col <- col + step
                                        steps <- steps + 1
                                    found
                                let found = struct (find (firstCol - 1) -1, find (lastCol + 1) 1)
                                overflowSources.[row] <- found
                                found
                        let struct (left, right) = sources
                        if left >= 0 && colPrefix.[left + 1] > colPrefix.[left] then visit row left
                        if right >= 0 && colPrefix.[right + 1] > colPrefix.[right] then visit row right
            if not (Map.isEmpty model.Cells) then visitOutside firstRow lastRow
            if frozenRows > 0 then visitBlock 0 (frozenRows - 1) firstCol lastCol
            if frozenCols > 0 then visitBlock firstRow lastRow 0 (frozenCols - 1)
            if frozenRows > 0 && frozenCols > 0 then visitBlock 0 (frozenRows - 1) 0 (frozenCols - 1)

            // Nodes of cells that are no longer wanted go back to the pool.
            let stale = ResizeArray<CellNode>()
            for entry in cellNodes do
                if entry.Value.Seen <> generation then stale.Add entry.Value
            for node in stale do
                cellNodes.Remove node.Key |> ignore
                if freeCellNodes.Count < 600 then
                    setStyle node.Element "display" "none"
                    node.Epoch <- -1
                    freeCellNodes.Add node
                else detach node.Element

            // Headers: the columns and rows in view (frozen ones always).
            // (Frozen ones sit in the wrapper that is pinned on both axes.)
            let placeHeaders (nodes: System.Collections.Generic.Dictionary<int, HeaderNode>) (free: ResizeArray<HeaderNode>) (parent: Element)
                             (create: unit -> HeaderNode) (draw: HeaderNode -> int -> unit) (frozen: int) (first: int) (last: int) (prefix: float array) =
                let place (index: int) =
                    if prefix.[index + 1] > prefix.[index] then
                        let node =
                            match nodes.TryGetValue index with
                            | true, node -> node
                            | false, _ ->
                                let node =
                                    if free.Count > 0 then
                                        let last = free.[free.Count - 1]
                                        free.RemoveAt (free.Count - 1)
                                        last
                                    else create ()
                                nodes.[index] <- node
                                node
                        let wanted = if index < frozen then layer.FrozenHeaders else parent
                        if not (isChildOf node.Element wanted) then appendChild wanted node.Element
                        node.Seen <- generation
                        draw node index
                for index = 0 to frozen - 1 do place index
                for index = first to last do place index
                let stale = ResizeArray<int>()
                for entry in nodes do
                    if entry.Value.Seen <> generation then stale.Add entry.Key
                for index in stale do
                    let node = nodes.[index]
                    nodes.Remove index |> ignore
                    setStyle node.Element "display" "none"
                    node.Signature <- ""
                    free.Add node
            placeHeaders colHeaderNodes freeColHeaders layer.ColHeaders newColumnHeader
                (fun node col -> drawColumnHeader node col (rowHeaderWidth + colPrefix.[col]) (colPrefix.[col + 1] - colPrefix.[col]))
                frozenCols firstCol lastCol colPrefix
            placeHeaders rowHeaderNodes freeRowHeaders layer.RowHeaders newRowHeader
                (fun node row -> drawRowHeader node row (headerHeight + rowPrefix.[row]) (rowPrefix.[row + 1] - rowPrefix.[row]))
                frozenRows firstRow lastRow rowPrefix
            renderDrawings content metrics
            renderOverlays layer metrics frozenRows frozenCols
            notifyChanges ()
        | _ ->
            // Not mounted: nothing to draw, but the revision still follows the content.
            observeRevision ()

    /// Bring the revision in line with the workbook as it is now. `step`
    /// says what brought it there: 1 an undo of `command`, 2 a redo of it,
    /// 0 anything else. Unchanged content keeps its number.
    and observeRevisionAfter (step: int) (command: Command option) =
        let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
        let same =
            current.Length = revisionSheets.Length
            && sheetNames = revisionNames
            && sameBookContent (bookContentOf workbook) revisionBook
            && Array.forall2 sameSheetContent current revisionSheets
        if not same then
            revisionSheets <- current
            revisionNames <- Array.copy sheetNames
            revisionBook <- bookContentOf workbook
            let fresh () =
                lastRevision <- lastRevision + 1
                lastRevision
            let isTop (steps: ResizeArray<Command * int * int>) (command: Command) =
                steps.Count > 0 && (let known, _, _ = steps.[steps.Count - 1] in obj.ReferenceEquals(known, command))
            match step, command with
            | 1, Some command when isTop appliedSteps command ->
                let entry = appliedSteps.[appliedSteps.Count - 1]
                appliedSteps.RemoveAt (appliedSteps.Count - 1)
                undoneSteps.Add entry
                let _, before, _ = entry
                revision <- before
            | 2, Some command when isTop undoneSteps command ->
                let entry = undoneSteps.[undoneSteps.Count - 1]
                undoneSteps.RemoveAt (undoneSteps.Count - 1)
                appliedSteps.Add entry
                let _, _, after = entry
                revision <- after
            | _ ->
                let before = revision
                revision <- fresh ()
                undoneSteps.Clear()
                // An undoable step is the one on top of the active sheet's
                // history; any other change (a host's SetCell, an undo out of
                // order) cannot be walked back to a known revision.
                match (if step = 0 then List.tryHead model.History.Past else None) with
                | Some pushed when not (appliedSteps |> Seq.exists (fun (known, _, _) -> obj.ReferenceEquals(known, pushed))) ->
                    appliedSteps.Add ((pushed, before, revision))
                    if appliedSteps.Count > 400 then appliedSteps.RemoveAt 0
                | _ -> appliedSteps.Clear()

    and observeRevision () = observeRevisionAfter 0 None

    /// Tell OnChange / OnSelectionChange subscribers what changed since they
    /// were last told. Every change to the workbook ends in a render, so this
    /// runs from there; the snapshot is updated before the callbacks run, so a
    /// callback that causes another render is not told twice.
    and notifyChanges () =
        observeRevision ()
        notifySheetList ()
        if changeListeners.Count > 0 || selectionListeners.Count > 0 then
            let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
            let book = bookContentOf workbook
            // (The list of sheets: which there are, their names, their order, which have a tab and its colour.)
            let structureChanged =
                current.Length <> notifiedSheets.Length || sheetNames <> notifiedNames || book.Hidden <> notifiedBook.Hidden
                || Array.exists2 (fun (a: SpreadsheetModel) (b: SpreadsheetModel) -> a.TabColor <> b.TabColor || a.VeryHidden <> b.VeryHidden) current notifiedSheets
            let changedSheets =
                [| for i in 0 .. current.Length - 1 do
                    if i >= notifiedSheets.Length || not (sameSheetContent current.[i] notifiedSheets.[i]) then yield i |]
            let selection = Some (activeSheetIndex, model.Selection, activeCellAddress())
            let selectionChanged = selection <> notifiedSelection
            let bookChanged = not (sameBookContent book notifiedBook)
            notifiedSheets <- current
            notifiedNames <- Array.copy sheetNames
            notifiedBook <- book
            notifiedSelection <- selection
            if structureChanged || bookChanged || changedSheets.Length > 0 then
                let change =
                    createObj [
                        "Sheets" ==> changedSheets
                        "StructureChanged" ==> structureChanged
                        "WorkbookChanged" ==> bookChanged
                        "ActiveSheetIndex" ==> activeSheetIndex
                        "Revision" ==> revision
                        "Modified" ==> (revision <> savedRevision)
                    ]
                for listener in changeListeners.ToArray() do
                    if changeListeners.Contains listener then (try listener change with _ -> ())
            if selectionChanged then
                let active = activeCellAddress()
                let change =
                    createObj [
                        "SheetIndex" ==> activeSheetIndex
                        "Selection" ==> selectionRefText ()
                        "ActiveCell" ==> (columnToLetter active.Col + string (active.Row + 1))
                        "ActiveRow" ==> active.Row
                        "ActiveCol" ==> active.Col
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
    /// Rebuilt only when the drawings or the sizes of rows / columns change.
    and renderDrawings (content: Element) (metrics: SpreadsheetLayout.SheetMetrics) =
        if not (obj.ReferenceEquals(renderedDrawings, model.Drawings)) || not (obj.ReferenceEquals(renderedMetrics, metrics)) then
            renderedDrawings <- box model.Drawings
            renderedMetrics <- box metrics
            for node in drawingNodes do detach node
            drawingNodes.Clear()
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
                    drawingNodes.Add node

    /// What sits on top of the cells: the selection, the active cell, the
    /// edit box, the selection's fill handle, the input message of the
    /// selected cell's validation, and the validation notice.
    and renderOverlays (layer: GridLayers) (metrics: SpreadsheetLayout.SheetMetrics) (frozenRows: int) (frozenCols: int) =
        let rowPrefix = metrics.RowPrefix
        let colPrefix = metrics.ColPrefix
        let rowCount = metrics.RowCount
        let colCount = metrics.ColCount
        let px (value: float) = string (System.Math.Round value) + "px"
        let hide (element: Element) = setStyle element "display" "none"
        // A rectangle of cells, clipped to the rows / columns of a pane.
        let paneRect (pane: int) (top: int) (left: int) (bottom: int) (right: int) : (float * float * float * float) option =
            let rowFrom, rowTo = if pane = 1 || pane = 3 then 0, frozenRows - 1 else frozenRows, rowCount - 1
            let colFrom, colTo = if pane >= 2 then 0, frozenCols - 1 else frozenCols, colCount - 1
            let top, bottom = max top rowFrom, min bottom rowTo
            let left, right = max left colFrom, min right colTo
            if top > bottom || left > right then None
            else
                Some (rowHeaderWidth + colPrefix.[left], headerHeight + rowPrefix.[top],
                      colPrefix.[right + 1] - colPrefix.[left], rowPrefix.[bottom + 1] - rowPrefix.[top])
        let paneOf (row: int) (col: int) = (if row < frozenRows then 1 else 0) + (if col < frozenCols then 2 else 0)
        let place (element: Element) (pane: int) =
            if not (isChildOf element layer.Panes.[pane]) then appendChild layer.Panes.[pane] element
        let top, left, bottom, right = selectionBounds()
        let top, bottom = max 0 top, min (rowCount - 1) bottom
        let left, right = max 0 left, min (colCount - 1) right
        let isRange = selectionIsRange ()
        // The selected range, tinted.
        for pane in 0 .. 3 do
            match (if isRange then paneRect pane top left bottom right else None) with
            | Some (x, y, w, h) when w > 0.0 && h > 0.0 ->
                setCssText layer.Selections.[pane]
                    ("position:absolute;left:" + px x + ";top:" + px y + ";width:" + px w + ";height:" + px h
                     + ";box-sizing:border-box;border:1px solid #217346;background:rgba(33,115,70,0.12);pointer-events:none;z-index:4")
            | _ -> hide layer.Selections.[pane]
        // The active cell, boxed.
        let active = activeCellAddress()
        let activeTop, activeLeft, activeBottom, activeRight =
            match findMergedCell active model with
            | Some merge -> merge.StartRow, merge.StartCol, merge.EndRow, merge.EndCol
            | None -> active.Row, active.Col, active.Row, active.Col
        let activePane = paneOf active.Row active.Col
        for pane in 0 .. 3 do
            match (if pane = activePane && rowCount > 0 && colCount > 0 then paneRect pane activeTop activeLeft (min (rowCount - 1) activeBottom) (min (colCount - 1) activeRight) else None) with
            | Some (x, y, w, h) when w > 0.0 && h > 0.0 ->
                setCssText layer.Actives.[pane]
                    ("position:absolute;left:" + px x + ";top:" + px y + ";width:" + px w + ";height:" + px h
                     + ";box-sizing:border-box;border:2px solid #217346;pointer-events:none;z-index:5")
            | _ -> hide layer.Actives.[pane]
        // The edit box, on the cell being edited.
        match model.EditingCell with
        | Some addr when addr.Row < rowCount && addr.Col < colCount ->
            let editBottom, editRight =
                match findMergedCell addr model with
                | Some merge -> min (rowCount - 1) merge.EndRow, min (colCount - 1) merge.EndCol
                | None -> addr.Row, addr.Col
            let pane = paneOf addr.Row addr.Col
            place layer.EditInput pane
            let x = rowHeaderWidth + colPrefix.[addr.Col]
            let y = headerHeight + rowPrefix.[addr.Row]
            let w = colPrefix.[editRight + 1] - colPrefix.[addr.Col]
            let h = rowPrefix.[editBottom + 1] - rowPrefix.[addr.Row]
            setCssText layer.EditInput
                ("position:absolute;left:" + px x + ";top:" + px y + ";width:" + px (max 40.0 w) + ";height:" + px (max 18.0 h)
                 + ";border:2px solid #217346;outline:none;padding:0 3px;margin:0;font-family:Arial, sans-serif;font-size:11px;background:white;box-sizing:border-box;z-index:11")
            if getValue layer.EditInput <> model.EditBuffer && not (isActiveElement layer.EditInput) then setValue layer.EditInput model.EditBuffer
        | _ -> hide layer.EditInput
        // The fill handle, at the bottom-right corner of the selection.
        if readOnly || model.EditingCell.IsSome || rowCount = 0 || colCount = 0 || bottom < 0 || right < 0 then hide layer.FillHandle
        else
            place layer.FillHandle (paneOf bottom right)
            let cornerX = rowHeaderWidth + colPrefix.[right + 1]
            let cornerY = headerHeight + rowPrefix.[bottom + 1]
            setCssText layer.FillHandle
                ("position:absolute;left:" + px (cornerX - 4.0) + ";top:" + px (cornerY - 4.0)
                 + ";width:7px;height:7px;background:#217346;border:1px solid white;box-sizing:border-box;cursor:crosshair;z-index:9")
        // The button that opens the list of a list validation, beside the active cell.
        if readOnly || model.EditingCell.IsSome || rowCount = 0 || colCount = 0 || active.Row >= rowCount || active.Col >= colCount
           || model.Validations.IsEmpty || not (showsListDropdown active model) then hide layer.ListButton
        else
            place layer.ListButton activePane
            let cellRight = rowHeaderWidth + colPrefix.[min colCount (activeRight + 1)]
            let cellTop = headerHeight + rowPrefix.[activeTop]
            let cellHeight = rowPrefix.[min rowCount (activeBottom + 1)] - rowPrefix.[activeTop]
            setCssText layer.ListButton
                ("position:absolute;left:" + px cellRight + ";top:" + px (cellTop + max 0.0 (cellHeight - 20.0)) + ";width:17px;height:" + px (min 20.0 (max 12.0 cellHeight))
                 + ";box-sizing:border-box;border:1px solid #8a8a8a;background:#f3f3f3;color:#444;font-size:9px;line-height:" + px (min 18.0 (max 10.0 (cellHeight - 2.0)))
                 + ";text-align:center;cursor:pointer;user-select:none;z-index:9")
            if getAttribute layer.ListButton "data-drawn" <> "1" then
                setAttribute layer.ListButton "data-drawn" "1"
                setTextContent layer.ListButton "\u25BC"
        // The circles of CircleInvalid: a red ellipse around each cell, in the cell's pane.
        let circled =
            match invalidCircles.TryGetValue activeSheetIndex with
            | true, cells -> cells
            | false, _ -> [||]
        let mutable used = 0
        for cell in circled do
            if used < 2000 && cell.Row < rowCount && cell.Col < colCount then
                let w = colPrefix.[cell.Col + 1] - colPrefix.[cell.Col]
                let h = rowPrefix.[cell.Row + 1] - rowPrefix.[cell.Row]
                if w > 0.0 && h > 0.0 then
                    let node =
                        if used < circleNodes.Count then circleNodes.[used]
                        else
                            let node = createElement "div"
                            setClassName node "spreadsheet-invalid-circle"
                            circleNodes.Add node
                            node
                    used <- used + 1
                    place node (paneOf cell.Row cell.Col)
                    setAttribute node "data-row" (string cell.Row)
                    setAttribute node "data-col" (string cell.Col)
                    setCssText node
                        ("position:absolute;left:" + px (rowHeaderWidth + colPrefix.[cell.Col] - 3.0) + ";top:" + px (headerHeight + rowPrefix.[cell.Row] - 3.0)
                         + ";width:" + px (w + 6.0) + ";height:" + px (h + 6.0)
                         + ";box-sizing:border-box;border:2px solid #E00000;border-radius:50%;pointer-events:none;z-index:6")
        while circleNodes.Count > used do
            detach circleNodes.[circleNodes.Count - 1]
            circleNodes.RemoveAt (circleNodes.Count - 1)
        // Notes: the validation notice, else the input message of the selected cell.
        let note (target: Element) (title: string) (message: string) (x: float) (y: float) (background: string) (border: string) =
            setCssText target
                ("position:absolute;left:" + px x + ";top:" + px y + ";max-width:260px;padding:4px 8px;background:" + background + ";border:1px solid " + border
                 + ";box-shadow:0 2px 6px rgba(0,0,0,0.2);font-family:Arial, sans-serif;font-size:11px;white-space:pre-wrap;pointer-events:none;z-index:30")
            setInnerHTML target ""
            if title <> "" then
                let heading = createElement "div"
                setTextContent heading title
                setStyle heading "font-weight" "bold"
                appendChild target heading
            let body = createElement "div"
            setTextContent body message
            appendChild target body
        let below (row: int) = headerHeight + rowPrefix.[min rowCount (max 0 row + 1)] + 3.0
        let leftOf (col: int) = rowHeaderWidth + colPrefix.[min colCount (max 0 col)]
        match validationNotice with
        | Some (style, title, message, cell) ->
            let background, border =
                match style with
                | "stop" -> "#FDECEA", "#D93025"
                | "warning" -> "#FFF4E5", "#F29900"
                | _ -> "#E8F0FE", "#1A73E8"
            note layer.Notice title message (leftOf cell.Col) (below cell.Row) background border
            setAttribute layer.Notice "data-style" style
            hide layer.InputMessage
        | None ->
            hide layer.Notice
            // Excel's input message: shown while the cell is selected.
            match model.Selection with
            | SingleCell (row, col) ->
                match getValidation { Row = row; Col = col } model with
                | Some dv when dv.ShowInput && (dv.InputMessage.IsSome || dv.InputTitle.IsSome) ->
                    note layer.InputMessage (defaultArg dv.InputTitle "") (defaultArg dv.InputMessage "") (leftOf left) (below bottom) "#FFFFE1" "#B0B0B0"
                | _ -> hide layer.InputMessage
            | _ -> hide layer.InputMessage

    /// A scroll draws the rows / columns that came into view.
    and renderScroll () = render ()

    /// Open the list of the active cell's list validation under the cell:
    /// a click on a value enters it (validated, one undo step). False when
    /// the cell has no list to offer.
    and showValidationList () : bool =
        let existing = getElementById "validation-list-dropdown"
        if not (isNull existing) then closePopup existing
        let addr = activeCellAddress()
        match (if readOnly then None else SheetOps.listValuesAt listReader addr model), layers with
        | Some values, Some layer when not values.IsEmpty ->
            if model.EditingCell.IsSome then finishEditing false |> ignore
            let dropdown = createElement "div"
            setAttribute dropdown "id" "validation-list-dropdown"
            // Under the cell: the button sits at the cell's bottom-right corner.
            let x, y, width =
                let metrics = SpreadsheetLayout.metricsOf model
                let rightCol =
                    match findMergedCell addr model with
                    | Some merge -> min (metrics.ColCount - 1) merge.EndCol
                    | None -> addr.Col
                let w = (metrics.ColPrefix.[rightCol + 1] - metrics.ColPrefix.[addr.Col]) * appliedZoom
                pageRight layer.ListButton - 17.0 - w, pageBottom layer.ListButton, w + 17.0
            setCssText dropdown
                ("position:absolute;left:" + string (System.Math.Round x) + "px;top:" + string (System.Math.Round y) + "px;min-width:" + string (System.Math.Round (max 60.0 width))
                 + "px;max-height:220px;overflow-y:auto;background:white;border:1px solid #8a8a8a;box-shadow:0 3px 8px rgba(0,0,0,0.25);font-family:Arial, sans-serif;font-size:12px;z-index:1000")
            let current = SpreadsheetCellText.cellDisplayText None 1.0e6 defaultFormat (getCellValue addr model)
            for value in values do
                let item = createElement "div"
                setClassName item "validation-list-item"
                setAttribute item "data-value" value
                setTextContent item value
                setCssText item ("padding:3px 8px;cursor:pointer;white-space:nowrap;" + (if value = current then "background:#d9ead3;" else ""))
                listen item "mousedown" (fun e ->
                    preventDefault e
                    closePopup dropdown
                    commitInput addr value true true |> ignore
                    scheduleNoticeDismiss ()
                    render ())
                appendChild dropdown item
            appendChild (getBody()) dropdown
            let rec closeList = fun (e: obj) ->
                let target = getEventTarget e
                if not (contains dropdown target) then
                    closePopup dropdown
                    unlisten (getDocument()) "mousedown" closeList
            JS.setTimeout (fun () -> if not disposed then listen (getDocument()) "mousedown" closeList) 0 |> ignore
            cleanups.Add(fun () ->
                unlisten (getDocument()) "mousedown" closeList
                closePopup dropdown)
            true
        | _ -> false

    /// Show filter dropdown for column
    and showFilterDropdown (col: int) (x: float) (y: float) =
        // Remove any existing filter dropdown
        let existingDropdown = getElementById "filter-dropdown"
        if not (isNull existingDropdown) then
            closePopup existingDropdown

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
        listen selectAllBtn "click" (fun _ ->
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
        listen clearAllBtn "click" (fun _ ->
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
            listen valueRow "mouseenter" (fun _ ->
                setStyles valueRow ["background", "#f0f0f0"]
            )
            listen valueRow "mouseleave" (fun _ ->
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
        listen applyBtn "click" (fun _ ->
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
            closePopup dropdown
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
        listen cancelBtn "click" (fun _ ->
            // Remove dropdown
            closePopup dropdown
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
                closePopup dropdown
                unlisten (getDocument()) "click" closeDropdown

        // Delay adding the click listener to avoid immediate closure
        JS.setTimeout (fun () ->
            if not disposed then listen (getDocument()) "click" closeDropdown
        ) 100 |> ignore
        cleanups.Add(fun () ->
            unlisten (getDocument()) "click" closeDropdown
            detach dropdown)


    /// Show find/replace dialog
    and showFindReplaceDialog () =
        // Remove any existing dialog
        let existingDialog = getElementById "find-replace-dialog"
        if not (isNull existingDialog) then
            closePopup existingDialog

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
        listen findNextBtn "click" (fun _ ->
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
        listen replaceBtn "click" (fun _ ->
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
        listen replaceAllBtn "click" (fun _ ->
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
        listen closeBtn "click" (fun _ ->
            closePopup overlay
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
                closePopup overlay
                unlisten (getDocument()) "keydown" handleEscape

        listen (getDocument()) "keydown" handleEscape
        cleanups.Add(fun () ->
            unlisten (getDocument()) "keydown" handleEscape
            detach overlay)

        ()


    /// Show conditional formatting dialog for current selection
    and showConditionalFormattingDialog () =
        // Remove any existing dialog
        let existingDialog = getElementById "conditional-format-dialog"
        if not (isNull existingDialog) then
            closePopup existingDialog

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

            listen ruleSelect "change" (fun _ -> updateInputVisibility())

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
            listen applyBtn "click" (fun _ ->
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

                        // Apply to cell (one undo step)
                        addRule [ { From = addr; To = addr } ] false conditionalFormat |> ignore

                        // Close dialog
                        closePopup overlay
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
            listen clearBtn "click" (fun _ ->
                changeConditionalFormats "Clear conditional formats" (removeConditionalFormats addr) |> ignore
                closePopup overlay
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
            listen closeBtn "click" (fun _ ->
                closePopup overlay
            )
            appendChild buttonRow closeBtn

            appendChild dialog buttonRow
            appendChild overlay dialog
            appendChild (getBody()) overlay

            // Close on Escape key
            let rec handleEscape = fun (e: obj) ->
                let key = getEventKey e
                if key = "Escape" then
                    closePopup overlay
                    unlisten (getDocument()) "keydown" handleEscape

            listen (getDocument()) "keydown" handleEscape
            cleanups.Add(fun () ->
                unlisten (getDocument()) "keydown" handleEscape
                detach overlay)

            ()

    /// Put the grid at the top-left corner of the sheet (after another sheet
    /// or workbook became the one on display).
    let resetScroll () =
        match gridContainer with
        | Some grid ->
            syncExtent()
            setGridScrollTop grid 0.0
            setGridScrollLeft grid 0.0
        | None -> ()
        scrollTop <- 0.0
        scrollLeft <- 0.0

    /// Handle scroll
    let handleScroll (e: obj) =
        let target = getCurrentTarget e
        scrollTop <- gridScrollTop target
        scrollLeft <- gridScrollLeft target
        if not scrollRenderScheduled then
            scrollRenderScheduled <- true
            requestAnimationFrame (fun () ->
                scrollRenderScheduled <- false
                if not disposed then renderScroll()
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

        // What all cells, and all headers, have in common. (Only a node's
        // place and what its format adds is set on the node itself.) Cells
        // are boxes of a given size that nothing outside depends on:
        // `contain` lets the browser lay out and paint each on its own.
        let sheetStyle = createElement "style"
        setTextContent sheetStyle
            // (A cell has no border of its own: its grid lines - the edge above it and the edge
            // to its left - are an inset shadow, and borders are child elements; see drawCell.)
            (".spreadsheet-container .spreadsheet-cell{position:absolute;padding:3px 5px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box;"
             + "font-family:Arial, sans-serif;font-size:11px;cursor:default;background:white;border:none;box-shadow:inset 0 1px 0 0 " + defaultGridColor + ",inset 1px 0 0 0 " + defaultGridColor + ";contain:strict}"
             + ".spreadsheet-container .spreadsheet-col-header{position:absolute;top:0;height:" + string headerHeight + "px;border-right:1px solid #d0d0d0;border-bottom:1px solid #a6a6a6;"
             + "background:linear-gradient(180deg, #f8f8f8 0%, #e8e8e8 100%);display:flex;align-items:center;justify-content:space-between;padding:0 4px;box-sizing:border-box;"
             + "font-weight:normal;font-size:11px;font-family:Arial, sans-serif;color:#444;user-select:none}"
             + ".spreadsheet-container .spreadsheet-row-header{position:absolute;left:0;width:" + string rowHeaderWidth + "px;border-right:1px solid #a6a6a6;border-bottom:1px solid #d0d0d0;"
             + "background:linear-gradient(90deg, #f8f8f8 0%, #e8e8e8 100%);display:flex;align-items:center;justify-content:center;box-sizing:border-box;"
             + "font-weight:normal;font-size:11px;font-family:Arial, sans-serif;color:#444;user-select:none;cursor:default}"
             + ".spreadsheet-container .col-header-text{flex:1;text-align:center}"
             + ".spreadsheet-container .col-filter-btn{width:16px;height:16px;align-items:center;justify-content:center;cursor:pointer;font-size:10px;border-radius:2px}"
             + ".spreadsheet-container .col-filter-btn:hover{background:#d0d0d0;opacity:1 !important}"
             + ".spreadsheet-container .col-resize-handle{position:absolute;top:0;right:0;width:4px;height:100%;cursor:col-resize;z-index:10}"
             + ".spreadsheet-container .row-resize-handle{position:absolute;bottom:0;left:0;width:100%;height:4px;cursor:row-resize;z-index:10}")
        appendChild gridDiv sheetStyle

        // The scrolled content, as large as the sheet (see `syncExtent`). Body
        // cells sit in it at their sheet position and scroll with it.
        let content = createElement "div"
        setClassName content "spreadsheet-content"
        setStyles content [
            "position", "relative"
            "background", "white"
            "outline", "none"
        ]

        // What stays in place while the content scrolls sits in zero-size
        // sticky wrappers at the content's origin, stacked above the body
        // cells: frozen rows (pinned vertically), frozen columns (pinned
        // horizontally), the frozen corner, the header strips above those,
        // and the corner with the headers of frozen rows / columns on top.
        let wrapper (className: string) (pinTop: bool) (pinLeft: bool) (zIndex: int) =
            let element = createElement "div"
            setClassName element className
            setCssText element
                ("position:sticky;width:0;height:0;" + (if pinTop then "top:0;" else "") + (if pinLeft then "left:0;" else "") + "z-index:" + string zIndex)
            appendChild content element
            element
        let frozenRowsPane = wrapper "spreadsheet-frozen-rows" true false 20
        let frozenColsPane = wrapper "spreadsheet-frozen-cols" false true 20
        let frozenCornerPane = wrapper "spreadsheet-frozen-corner" true true 21
        let colHeaderStrip = wrapper "spreadsheet-col-headers" true false 22
        let rowHeaderStrip = wrapper "spreadsheet-row-headers" false true 22
        let frozenHeaders = wrapper "spreadsheet-frozen-headers" true true 23
        let corner = createElement "div"
        setClassName corner "spreadsheet-corner"
        setCssText corner
            ("position:absolute;top:0;left:0;width:" + string rowHeaderWidth + "px;height:" + string headerHeight
             + "px;border-right:1px solid #a6a6a6;border-bottom:1px solid #a6a6a6;background:linear-gradient(135deg, #f8f8f8 0%, #e0e0e0 100%);box-sizing:border-box;z-index:3")
        appendChild frozenHeaders corner
        let panes = [| content; frozenRowsPane; frozenColsPane; frozenCornerPane |]
        let overlay (className: string) (pane: Element) =
            let element = createElement "div"
            setClassName element className
            setStyle element "display" "none"
            appendChild pane element
            element
        let editBox = createElement "input"
        setType editBox "text"
        setClassName editBox "spreadsheet-edit-input"
        setAttribute editBox "autocomplete" "off"
        setAttribute editBox "spellcheck" "false"
        setStyle editBox "display" "none"
        appendChild content editBox
        listen editBox "input" handleInput
        layers <-
            Some { Panes = panes
                   ColHeaders = colHeaderStrip
                   RowHeaders = rowHeaderStrip
                   FrozenHeaders = frozenHeaders
                   Selections = panes |> Array.map (overlay "spreadsheet-selection")
                   Actives = panes |> Array.map (overlay "spreadsheet-active-cell")
                   FillHandle = overlay "spreadsheet-fill-handle" content
                   EditInput = editBox
                   Notice = overlay "spreadsheet-validation-notice" content
                   InputMessage = overlay "spreadsheet-input-message" content
                   ListButton = overlay "spreadsheet-list-button" content }

        // Add event listeners
        listen gridDiv "scroll" handleScroll
        listen content "keydown" handleKeyDown

        // One listener each for the whole grid: which cell or header a
        // pointer event is for is worked out from where it happened.
        let headerIndex (target: Element) (selector: string) (attribute: string) : int option =
            let header = closest target selector
            if isNull (box header) then None
            else
                match System.Int32.TryParse (getAttribute header attribute) with
                | true, index -> Some index
                | _ -> None
        let within (target: Element) (selector: string) = not (isNull (box (closest target selector)))
        listen content "mousedown" (fun e ->
            let target = getEventTarget e
            if within target ".spreadsheet-edit-input" then ()   // placing the caret in the edit box
            elif within target ".spreadsheet-list-button" then
                preventDefault e
                stopPropagation e
                showValidationList () |> ignore
            elif within target ".spreadsheet-fill-handle" then
                preventDefault e
                if not readOnly then
                    isFilling <- true
                    fillSource <- Some (selectionBounds())
                    fillTarget <- None
            elif within target ".col-resize-handle" then
                preventDefault e
                match headerIndex target ".spreadsheet-col-header" "data-col" with
                | Some col when not readOnly ->
                    isResizingColumn <- true
                    resizeBefore <- Some (axesState model)
                    resizeColumnIndex <- Some col
                    resizeStartPos <- getClientX e
                    resizeStartSize <- getColumnWidth col model
                | _ -> ()
            elif within target ".row-resize-handle" then
                preventDefault e
                match headerIndex target ".spreadsheet-row-header" "data-row" with
                | Some row when not readOnly ->
                    isResizingRow <- true
                    resizeBefore <- Some (axesState model)
                    resizeRowIndex <- Some row
                    resizeStartPos <- getClientY e
                    resizeStartSize <- getRowHeight row model
                | _ -> ()
            elif within target ".spreadsheet-col-header, .spreadsheet-row-header, .spreadsheet-corner" then
                preventDefault e   // no text selection; a click on a column header sorts (see below)
            elif mouseButton e = 0 then
                preventDefault e  // Prevent text selection
                match hitTestEvent e with
                | SpreadsheetLayout.Cell hitAddr ->
                    // Clicking another cell commits the entry in progress.
                    if model.EditingCell.IsSome then finishEditing true |> ignore
                    if not (holdsFocus content) then focusQuietly content
                    if eventFlag e "shiftKey" then
                        // Shift+click: from the active cell to here.
                        selectFromTo (activeCellAddress()) hitAddr
                    else
                        isDragging <- true
                        dragStartCell <- Some hitAddr
                        selectCellAddress hitAddr
                    render()
                | _ -> ())
        listen content "dblclick" (fun e ->
            let target = getEventTarget e
            // A double click on a resize handle fits the column / row to its
            // content - every selected one when it is among whole selected
            // columns / rows.
            let fitFromHandle (rows: bool) (index: int) =
                if not readOnly then
                    let top, left, bottom, right = selectionBounds()
                    let whole = if rows then left = 0 && right >= model.ColCount - 1 else top = 0 && bottom >= model.RowCount - 1
                    let first, last = selectedSpan rows
                    let span = if whole && index >= first && index <= last then first, last else index, index
                    if autoFitSpan rows span then render()
            if within target ".col-resize-handle" then
                preventDefault e
                match headerIndex target ".spreadsheet-col-header" "data-col" with
                | Some col -> fitFromHandle false col
                | None -> ()
            elif within target ".row-resize-handle" then
                preventDefault e
                match headerIndex target ".spreadsheet-row-header" "data-row" with
                | Some row -> fitFromHandle true row
                | None -> ()
            elif not (within target ".spreadsheet-edit-input, .spreadsheet-col-header, .spreadsheet-row-header, .spreadsheet-corner, .spreadsheet-fill-handle, .spreadsheet-list-button") then
                match hitTestEvent e with
                | SpreadsheetLayout.Cell hitAddr -> startEditing hitAddr.Row hitAddr.Col
                | _ -> ())
        listen content "click" (fun e ->
            let target = getEventTarget e
            if readOnly then ()
            elif within target ".col-filter-btn" then
                preventDefault e
                match headerIndex target ".spreadsheet-col-header" "data-col" with
                | Some col ->
                    let button = closest target ".col-filter-btn"
                    showFilterDropdown col (pageRight button) (pageBottom button)
                | None -> ()
            elif within target ".col-header-text" then
                match headerIndex target ".spreadsheet-col-header" "data-col" with
                | Some col ->
                    // Toggle sort direction if clicking same column, otherwise ascending
                    let ascending =
                        match model.SortState with
                        | Some sortState when sortState.Column = col -> not sortState.Ascending
                        | _ -> true
                    sortActiveSheet col ascending (sprintf "Sort by column %s" (columnToLetter col))
                    render()
                | None -> ())

        // The system clipboard: the browser raises these on Ctrl+C / X / V
        // (and from its Edit menu) while the grid has the focus.
        let onCopy (isCut: bool) (e: obj) =
            if model.EditingCell.IsNone then
                clipboardEvents <- clipboardEvents + 1
                let text, html = copySelection isCut
                clipboardEventSet e "text/plain" text |> ignore
                clipboardEventSet e "text/html" html |> ignore
                preventDefault e
                if isCut && not readOnly then render()
        listen content "copy" (onCopy false)
        listen content "cut" (onCopy true)
        listen content "paste" (fun e ->
            if model.EditingCell.IsNone then
                clipboardEvents <- clipboardEvents + 1
                preventDefault e
                if pasteFlavours (clipboardEventGet e "text/plain") (clipboardEventGet e "text/html") then
                    let active = activeCellAddress()
                    revealCell active.Row active.Col
                    render())

        // Drag/resize handlers - attach to document to catch events anywhere
        let doc = getDocument()

        // Mouse move for dragging and resizing (throttled with requestAnimationFrame)
        let onDocumentMouseMove = (fun (e: obj) ->
            if isResizingColumn then
                match resizeColumnIndex with
                | Some col ->
                    let currentX = getProp e "clientX" |> unbox<float>
                    let delta = (currentX - resizeStartPos) / appliedZoom
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
                    let delta = (currentY - resizeStartPos) / appliedZoom
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
            elif isFilling then
                match fillSource, hitCellNear e with
                | Some source, Some hitAddr ->
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
                match dragStartCell, hitCellNear e with
                | Some start, Some hitAddr ->
                    // The cell the drag started at stays the active cell.
                    let before = model.Selection
                    selectFromTo start hitAddr
                    if model.Selection <> before then
                        revealCell hitAddr.Row hitAddr.Col
                        render()
                | _ -> ()
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
                // The whole drag is one undo step.
                (match resizeBefore with Some before -> commitAxes "Column width" before |> ignore | None -> ())
                resizeBefore <- None
                render()  // Final render to ensure clean state
            elif isResizingRow then
                isResizingRow <- false
                resizeRowIndex <- None
                resizeRenderScheduled <- false
                (match resizeBefore with Some before -> commitAxes "Row height" before |> ignore | None -> ())
                resizeBefore <- None
                render()  // Final render to ensure clean state
        )

        listen doc "mousemove" onDocumentMouseMove
        listen doc "mouseup" onDocumentMouseUp
        cleanups.Add(fun () ->
            unlisten doc "mousemove" onDocumentMouseMove
            unlisten doc "mouseup" onDocumentMouseUp)

        // Focus content to receive keyboard events
        setAttribute content "tabindex" "0"

        appendChild gridDiv content
        appendChild container gridDiv

        gridContainer <- Some gridDiv
        contentElement <- Some content
        disposed <- false
        extentWidth <- -1.0
        extentHeight <- -1.0
        // (The new nodes are at 100 % with their headers shown: see syncView.)
        appliedZoom <- 1.0
        appliedHeaders <- true
        headerHeight <- 30.0
        rowHeaderWidth <- 50.0

        // Initial render
        containerHeight <- gridClientHeight gridDiv
        containerWidth <- gridClientWidth gridDiv
        render()

        // Follow the container: the visible range depends on the grid's size.
        let stopObserving =
            observeResize gridDiv (fun () ->
                if not disposed then
                    let height = gridClientHeight gridDiv
                    let width = gridClientWidth gridDiv
                    if height <> containerHeight || width <> containerWidth then
                        containerHeight <- height
                        containerWidth <- width
                        render())
        cleanups.Add stopObserving

        // Focus for keyboard
        focus content

    /// Take the spreadsheet out of the page: removes every listener it
    /// registered - on the document, on its own nodes and on the nodes of its
    /// dialogs and dropdowns - the resize observer and every DOM node it
    /// created, and drops all subscriptions and what it held outside itself,
    /// so that a disposed instance can be collected. The workbook stays: the
    /// instance can be mounted again.
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
        // Every listener, also on the grid's own nodes: a detached node with
        // a listener keeps the component (and its workbook) reachable from
        // anything that still holds the node.
        for target, eventName, handler in listeners.ToArray() do
            try removeListenerRaw (unbox target) eventName handler with _ -> ()
        listeners.Clear()
        match gridContainer with
        | Some gridDiv -> detach gridDiv
        | None -> ()
        gridContainer <- None
        contentElement <- None
        layers <- None
        // What is held outside the component on its behalf: the evaluator
        // of conditional-format formulas (a closure over this workbook).
        Rbt.Spreadsheet.ConditionalFormatting.resetFormulaEvaluator ()
        ruleWorkbook <- None
        recalcEngine <- None
        calculatedWorkbook <- None
        engineGeneration <- engineGeneration + 1
        overflowSources.Clear()
        overflowCells <- null
        notifiedSheets <- [||]
        notifiedSelection <- None
        validationListeners.Clear()
        isDragging <- false
        dragStartCell <- None
        isFilling <- false
        isResizingColumn <- false
        isResizingRow <- false
        cellNodes.Clear()
        freeCellNodes.Clear()
        colHeaderNodes.Clear()
        rowHeaderNodes.Clear()
        freeColHeaders.Clear()
        freeRowHeaders.Clear()
        drawingNodes.Clear()
        renderedCells <- null
        renderedStyleInputs <- [||]
        renderedDrawings <- null
        renderedMetrics <- null
        editStateListeners.Clear()
        notifiedEdit <- None
        changeListeners.Clear()
        selectionListeners.Clear()
        sheetChangedListeners.Clear()

    /// Call `callback` after every change to the workbook's content: cell
    /// values and formats, row / column sizes, merges, inserted or deleted
    /// rows and columns, tables, validations, conditional formats, sheets
    /// added, removed or renamed, undo and redo. It receives
    /// { Sheets: number[] (indices of the sheets that changed),
    ///   StructureChanged: boolean (sheet list changed), WorkbookChanged:
    ///   boolean (defined names changed), ActiveSheetIndex,
    ///   Revision: number, Modified: boolean } (see GetRevision / IsModified).
    /// Selection moves and in-cell typing before it is committed are not
    /// document changes. Returns a function that removes the subscription.
    member _.OnChange(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        if changeListeners.Count = 0 && selectionListeners.Count = 0 then
            // Changes are reported from here on, not since the last subscriber left.
            notifiedSheets <- Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
            notifiedNames <- Array.copy sheetNames
            notifiedBook <- bookContentOf workbook
            notifiedSelection <- Some (activeSheetIndex, model.Selection, activeCellAddress())
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
            notifiedBook <- bookContentOf workbook
            notifiedSelection <- Some (activeSheetIndex, model.Selection, activeCellAddress())
        selectionListeners.Add listener
        System.Action(fun () -> selectionListeners.Remove listener |> ignore)

    /// The revision of the workbook's content: a number that changes with
    /// every change to it (not with selection moves or typing that is not
    /// committed yet) and that undo / redo bring back - undoing every change
    /// since a save gives the revision that was saved.
    member _.GetRevision() : int =
        observeRevision ()
        revision

    /// Remember a revision as the saved one: IsModified() is false while the
    /// workbook is at it, also after undoing back to it. Without an argument
    /// it is the current revision. A save that takes time passes the revision
    /// it wrote, so that what the user did meanwhile stays unsaved:
    ///   const r = grid.GetRevision(); const bytes = await grid.ExportXlsx();
    ///   await upload(bytes); grid.MarkSaved(r);
    /// ImportXlsx / InstallImportedWorkbook mark the opened workbook as saved;
    /// SetWorkbook / InstallWorkbook count as a change.
    member _.MarkSaved(?savedAt: int) : unit =
        observeRevision ()
        savedRevision <- defaultArg savedAt revision

    /// Whether the workbook's content differs from the saved revision.
    member _.IsModified() : bool =
        observeRevision ()
        revision <> savedRevision

    /// Every sheet's current model, in sheet order (the active sheet's
    /// included, with its latest edits).
    member _.GetSheets() : SpreadsheetModel array =
        // (Converts the sheets of an opened file that were not needed so far:
        // see IsSheetLoaded. GetModel() is the sheet on display alone.)
        convertAllSheets ()
        Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])

    /// Names of every sheet, in sheet order.
    member _.GetSheetNames() : string array =
        Array.copy sheetNames

    /// 0-based index of the sheet on display.
    member _.GetActiveSheetIndex() : int = activeSheetIndex

    /// Get model (for debugging)
    member _.GetModel() = model

    /// Enable or disable viewer mode. A read-only grid is a viewer: cells
    /// cannot be edited, filled, cut, pasted into, sorted or resized, and
    /// nothing that invites it is shown (no fill handle, filter buttons or
    /// resize handles). Selecting with the mouse and the keyboard, copying,
    /// switching sheets and scrolling work as in the editor. An entry in
    /// progress is dropped. (The host's own calls - SetCell, Apply... - are
    /// not blocked.)
    member _.SetReadOnly(value: bool) =
        if readOnly <> value then
            readOnly <- value
            if readOnly && model.EditingCell.IsSome then
                finishEditing false |> ignore
            else render()

    /// Whether the grid is in viewer mode.
    member _.IsReadOnly() : bool = readOnly

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
            // An entry in progress belongs to the sheet being left.
            if model.EditingCell.IsSome then finishEditing true |> ignore
            // Persist current sheet's mutable state into the array.
            sheets.[activeSheetIndex] <- model
            if activeSheetIndex < sheetAnchors.Length then sheetAnchors.[activeSheetIndex] <- anchorCell
            // A sheet of an opened file is converted when it is first shown.
            convertInPlace index
            activeSheetIndex <- index
            model <- sheets.[index]
            anchorCell <- (if index < sheetAnchors.Length then sheetAnchors.[index] else { Row = 0; Col = 0 })
            // Scroll lives outside the SpreadsheetModel; a sheet starts at
            // its top-left corner (we don't persist scroll yet — reset-to-zero
            // matches Excel's behavior on workbook reload). The container is
            // put there too: the content is drawn for where it is scrolled to.
            resetScroll ()
            render()
            for listener in sheetChangedListeners do
                try listener index with _ -> ()

    /// Register a callback invoked after every successful sheet switch.
    /// Chrome uses this to update its tab-strip selection highlight when
    /// the switch is triggered from inside (e.g. via keyboard shortcut)
    /// rather than from the chrome's own click handler.
    member _.OnSheetChanged(listener: System.Action<int>) =
        sheetChangedListeners.Add(listener.Invoke)

    // ===== Sheets =====
    // Every member below that changes the list of sheets is ONE undo step,
    // gives a new revision and is reported by OnChange (StructureChanged) and
    // OnSheetsChange. The step is undone from the sheet that is on display
    // once it is done (the new sheet after AddSheet / DuplicateSheet, the
    // sheet that took a deleted one's place, else the sheet that was on
    // display): CanUndo() / Undo() there take it back. Undo histories stay
    // per sheet: after switching to another sheet, Undo acts on that sheet.

    /// The sheets in tab order: [{ index, id (stays with the sheet when it is
    /// renamed or moved), name, visibility: "visible" | "hidden" |
    /// "veryHidden", tabColor: "#RRGGBB" | null, isLoaded (see
    /// IsSheetLoaded), isActive }].
    member _.GetSheetInfos() : obj array = sheetInfos ()

    /// Call `callback` with the list of sheets (as GetSheetInfos gives it)
    /// after it changed: a sheet added, deleted, renamed, moved, copied,
    /// hidden, shown or given a tab colour - by a call, by undo / redo or by
    /// another workbook being installed. (Which sheet is on display is not
    /// part of it: OnSelectionChange / OnSheetChanged tell that.) Returns a
    /// function that removes the subscription.
    member _.OnSheetsChange(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        // Changes are reported from here on.
        if sheetsListeners.Count = 0 then notifiedSheetList <- sheetListText ()
        sheetsListeners.Add listener
        System.Action(fun () -> sheetsListeners.Remove listener |> ignore)

    /// Add an empty sheet and show it. `name`: its name (omitted or null:
    /// the first of "Sheet1", "Sheet2", ... that is free); `atIndex`: its
    /// position in the tab order, 0-based (omitted or null: after the last
    /// sheet). Returns the new sheet's index. One undo step. Throws on an
    /// invalid or duplicate name (see RenameSheet) or a position outside
    /// 0 .. number of sheets.
    member _.AddSheet(name: obj, atIndex: obj) : int =
        let memberName = "AddSheet"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let wb = bookSnapshot ()
        let name =
            match Host.optionalArg (Host.stringArg memberName "name") name with
            | Some name ->
                (match Rbt.Spreadsheet.Workbook.sheetNameProblem name with
                 | Some problem -> Host.hostFail memberName problem
                 | None -> ())
                if Rbt.Spreadsheet.Workbook.sheetNameTaken None name wb then Host.hostFail memberName (sprintf "there is already a sheet named \"%s\"" name)
                name
            | None -> Rbt.Spreadsheet.Workbook.nextSheetName wb
        let at =
            match Host.optionalArg (Host.intArg memberName "atIndex") atIndex with
            | Some at when at < 0 || at > sheets.Length -> Host.hostFail memberName (sprintf "atIndex %d is outside 0 to %d" at sheets.Length)
            | Some at -> at
            | None -> sheets.Length
        let id = Rbt.Spreadsheet.Workbook.freshSheetId usedSheetIds.Contains wb
        usedSheetIds.Add id |> ignore
        let sheet : Rbt.Spreadsheet.Workbook.Sheet = { Id = id; Name = name; Model = createEmpty 1000 26; Visible = true }
        let activeBefore = wb.ActiveSheet
        // (Once undone, redo puts back the sheet as the undo took it away.)
        let removed : Rbt.Spreadsheet.Workbook.SheetRemoval option ref = ref None
        runBookStep "Add sheet"
            (fun () ->
                let current = bookSnapshot ()
                (match removed.Value with
                 | Some removal -> { Rbt.Spreadsheet.Workbook.restoreSheet removal current with ActiveSheet = id }
                 | None -> Rbt.Spreadsheet.Workbook.insertSheet sheet at current)
                |> installBook
                // (A formula that named a sheet of this name before there was one reads it now.)
                recalculateEverything ())
            (fun () ->
                match Rbt.Spreadsheet.Workbook.removeSheet id (bookSnapshot ()) with
                | Some (without, removal) ->
                    removed.Value <- Some removal
                    installBook (if Map.containsKey activeBefore without.Sheets then { without with ActiveSheet = activeBefore } else without)
                    recalculateEverything ()
                | None -> ())
        render()
        activeSheetIndex

    /// Rename the sheet at `index`. Everything that names the sheet follows:
    /// the formulas of every sheet (cells, conditional formats, validations),
    /// defined names and hyperlinks into the sheet, the new name quoted where
    /// a reference needs it ('My Sheet'!A1). One undo step. Throws when the
    /// name is empty, longer than 31 characters, contains any of : \ / ? * [ ],
    /// begins or ends with an apostrophe, is "History", or is the name of
    /// another sheet (whatever the case); and when there is no such sheet.
    member _.RenameSheet(index: obj, name: obj) : unit =
        let memberName = "RenameSheet"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        let name = Host.stringArg memberName "name" name
        (match Rbt.Spreadsheet.Workbook.sheetNameProblem name with
         | Some problem -> Host.hostFail memberName problem
         | None -> ())
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let wb = bookSnapshot ()
        let id = sheetIdForIndex index
        if Rbt.Spreadsheet.Workbook.sheetNameTaken (Some id) name wb then Host.hostFail memberName (sprintf "there is already a sheet named \"%s\"" name)
        let oldName = sheetNames.[index]
        if oldName <> name then
            let rename (newName: string) () =
                // The formulas of every sheet follow: those not converted yet are now.
                convertFormulaSheets ()
                installBook (Rbt.Spreadsheet.Workbook.renameSheetEverywhere id newName (bookSnapshot ()))
                rebaseCalcEngine ()
            runBookStep "Rename sheet" (rename name) (rename oldName)
            render()

    /// Delete the sheet at `index`. As in Excel, references to it in the
    /// formulas of the other sheets and in defined names become #REF!; the
    /// names of its scope, its tables and what the opened file kept for it
    /// (pictures, charts, pivot tables) go with it. When it was on display
    /// the nearest visible sheet is shown. One undo step (undo brings the
    /// sheet and the references back). Throws when there is no such sheet or
    /// when it is the only visible one.
    member _.DeleteSheet(index: obj) : unit =
        let memberName = "DeleteSheet"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let id = sheetIdForIndex index
        (match Rbt.Spreadsheet.Workbook.lastVisibleSheetProblem id (bookSnapshot ()) with
         | Some problem -> Host.hostFail memberName problem
         | None -> ())
        let removed : Rbt.Spreadsheet.Workbook.SheetRemoval option ref = ref None
        runBookStep "Delete sheet"
            (fun () ->
                convertFormulaSheets ()
                match Rbt.Spreadsheet.Workbook.removeSheet id (bookSnapshot ()) with
                | Some (without, removal) ->
                    removed.Value <- Some removal
                    installBook without
                    recalculateEverything ()
                | None -> ())
            (fun () ->
                match removed.Value with
                | Some removal ->
                    installBook (Rbt.Spreadsheet.Workbook.restoreSheet removal (bookSnapshot ()))
                    recalculateEverything ()
                | None -> ())
        render()

    /// Move the sheet at index `from` so that it is at index `toIndex`
    /// afterwards. The sheet on display stays on display. One undo step.
    /// Throws when either index is not a sheet's.
    member _.MoveSheet(from: obj, toIndex: obj) : unit =
        let memberName = "MoveSheet"
        let from = Host.intArg memberName "from" from
        let target = Host.intArg memberName "to" toIndex
        if from < 0 || from >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" from sheets.Length)
        if target < 0 || target >= sheets.Length then Host.hostFail memberName (sprintf "to %d is outside 0 to %d" target (sheets.Length - 1))
        if from <> target then
            if model.EditingCell.IsSome then finishEditing true |> ignore
            let id = sheetIdForIndex from
            let move (position: int) () =
                installBook (Rbt.Spreadsheet.Workbook.moveSheet id position (bookSnapshot ()))
                rebaseCalcEngine ()
            runBookStep "Move sheet" (move target) (move from)
            render()

    /// Copy the sheet at `index`; the copy is put right after it, named
    /// "Name (2)" (then "(3)", ...) and shown. Returns the copy's index. The
    /// copy's formulas that named the sheet itself name the copy; the names
    /// of the sheet's scope and its tables are copied (a table under a new
    /// name, "Sales_2", which the copy's structured references follow), and
    /// so are its pictures, comments, formats, validations and conditional
    /// formats. What an opened file only kept for the sheet - charts and
    /// other drawing objects the editor does not model, pivot tables, named
    /// cell styles - stays with the original: the copy has none of it. One
    /// undo step. Throws when there is no such sheet.
    member _.DuplicateSheet(index: obj) : int =
        let memberName = "DuplicateSheet"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        if model.EditingCell.IsSome then finishEditing true |> ignore
        // (The copy is one of the sheet as it is: a sheet not converted yet is now.)
        sheetModelAt index |> ignore
        let wb = bookSnapshot ()
        let sourceId = sheetIdForIndex index
        let id = Rbt.Spreadsheet.Workbook.freshSheetId usedSheetIds.Contains wb
        usedSheetIds.Add id |> ignore
        let name = Rbt.Spreadsheet.Workbook.copySheetName sheetNames.[index] wb
        let activeBefore = wb.ActiveSheet
        let removed : Rbt.Spreadsheet.Workbook.SheetRemoval option ref = ref None
        runBookStep "Copy sheet"
            (fun () ->
                let current = bookSnapshot ()
                (match removed.Value with
                 | Some removal -> { Rbt.Spreadsheet.Workbook.restoreSheet removal current with ActiveSheet = id }
                 | None -> Rbt.Spreadsheet.Workbook.duplicateSheet sourceId id name (Rbt.Spreadsheet.Workbook.sheetIndex sourceId current + 1) current)
                |> installBook
                rebaseCalcEngine ())
            (fun () ->
                match Rbt.Spreadsheet.Workbook.removeSheet id (bookSnapshot ()) with
                | Some (without, removal) ->
                    removed.Value <- Some removal
                    installBook (if Map.containsKey activeBefore without.Sheets then { without with ActiveSheet = activeBefore } else without)
                    recalculateEverything ()
                | None -> ())
        render()
        activeSheetIndex

    /// Show or hide the tab of the sheet at `index`: "visible", "hidden" or
    /// "veryHidden" (Excel does not offer to unhide such a sheet; the API
    /// does). A hidden sheet stays part of the workbook: formulas read it.
    /// Hiding the sheet on display shows the nearest visible one. One undo
    /// step. Throws on another word, when there is no such sheet, and when
    /// the sheet is the only visible one.
    member _.SetSheetVisibility(index: obj, visibility: obj) : unit =
        let memberName = "SetSheetVisibility"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        let visible, veryHidden =
            match Host.stringArg memberName "visibility" visibility with
            | "visible" -> true, false
            | "hidden" -> false, false
            | "veryHidden" -> false, true
            | other -> Host.hostFail memberName (sprintf "visibility must be \"visible\", \"hidden\" or \"veryHidden\", got \"%s\"" other)
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let wb = bookSnapshot ()
        let id = sheetIdForIndex index
        match Map.tryFind id wb.Sheets with
        | None -> ()
        | Some sheet ->
            let before = sheet.Visible, (not sheet.Visible && sheet.Model.VeryHidden)
            if before <> (visible, veryHidden) then
                if not visible then
                    match Rbt.Spreadsheet.Workbook.lastVisibleSheetProblem id wb with
                    | Some problem when sheet.Visible -> Host.hostFail memberName problem
                    | _ -> ()
                let activeBefore = wb.ActiveSheet
                runBookStep (if visible then "Unhide sheet" else "Hide sheet")
                    (fun () -> installBook (Rbt.Spreadsheet.Workbook.setSheetVisibility id visible veryHidden (bookSnapshot ())))
                    (fun () ->
                        let restored = Rbt.Spreadsheet.Workbook.setSheetVisibility id (fst before) (snd before) (bookSnapshot ())
                        installBook (if Map.containsKey activeBefore restored.Sheets then { restored with ActiveSheet = activeBefore } else restored))
                render()

    /// Give the tab of the sheet at `index` a colour ("#RRGGBB"), or none
    /// (null). The component draws no tabs itself: the colour is part of the
    /// workbook (GetSheetInfos, the saved file). One undo step. Throws on a
    /// malformed colour or when there is no such sheet.
    member _.SetSheetTabColor(index: obj, color: obj) : unit =
        let memberName = "SetSheetTabColor"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        let color = Host.optionalArg (Host.colorArg memberName "color") color
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let id = sheetIdForIndex index
        let before = (if index = activeSheetIndex then model else sheets.[index]).TabColor
        if before <> color then
            let put (value: string option) () =
                let current = bookSnapshot ()
                match Map.tryFind id current.Sheets with
                | Some sheet -> installBook { current with Sheets = Map.add id { sheet with Model = { sheet.Model with TabColor = value } } current.Sheets }
                | None -> ()
            runBookStep "Tab colour" (put color) (put before)
            render()

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
        invalidCircles.Clear()
        // The file the previous workbook came from stays the source as long as
        // one of the new sheets is a sheet of it (sheets renamed, reordered,
        // added or removed by the host): each model says which file and sheet
        // it came from. A different workbook has no such sheet.
        // (InstallImportedWorkbook / ImportXlsx set the new source.)
        xlsxSource <-
            xlsxSource
            |> Option.filter (fun source ->
                newSheets |> Array.exists (fun sheet -> match sheet.SourceSheet with Some origin -> origin.Book = source.Id | None -> false))
        sheets <- newSheets
        sheetNames <-
            if names.Length = newSheets.Length then Array.copy names
            else Array.init newSheets.Length (fun i -> sprintf "Sheet%d" (i + 1))
        let safeIndex = max 0 (min (sheets.Length - 1) activeIndex)
        // The sheet on display is a converted one; the others may be
        // stand-ins (ImportXlsx), converted when they are needed.
        convertInPlace safeIndex
        formulaSheetsConverted <- not (sheets |> Array.exists isDeferredSheet)
        activeSheetIndex <- safeIndex
        model <- sheets.[safeIndex]
        workbook <-
            match context with
            | Some wb when List.length wb.SheetOrder = sheets.Length ->
                Some { wb with ActiveSheet = List.item safeIndex wb.SheetOrder }
            | _ ->
                Some (buildWorkbookContext sheets sheetNames (Array.create sheets.Length true) Map.empty Map.empty safeIndex)
        sheetAnchors <- Array.create sheets.Length { Row = 0; Col = 0 }
        anchorCell <- { Row = 0; Col = 0 }
        // The steps of the previous workbook cannot be undone into this one.
        appliedSteps.Clear()
        undoneSteps.Clear()
        bookSteps.Clear()
        usedSheetIds.Clear()
        for index in 0 .. sheets.Length - 1 do usedSheetIds.Add (sheetIdForIndex index) |> ignore
        // (The subscribers of OnSheetsChange are told of the new workbook's sheets.)
        notifiedSheetList <- ""
        resetScroll ()
        render()
        // The workbook as it is installed is the calculated state edits start
        // from (a file's formulas come with their values). Its dependency
        // graph is built a moment later (or at the first edit, when that comes
        // first); the formulas of a sheet that is not converted are read off
        // the file, and the sheet stays as it is.
        calculatedWorkbook <- Some (currentWorkbookSnapshot())
        let generation = engineGeneration
        setTimeoutMs (fun () ->
            if not disposed && incrementalRecalculation && generation = engineGeneration && recalcEngine.IsNone then
                ensureRecalcEngine () |> ignore) 250 |> ignore
        for listener in sheetChangedListeners do
            try listener safeIndex with _ -> ()

    /// The four edges of a cell of the active sheet as they are drawn - one
    /// line per edge, shared with the neighbouring cell: { Top, Right,
    /// Bottom, Left }, each "grid" (a grid line), null (nothing: inside a
    /// merged cell, grid lines off, or hidden by a fill) or { Style, Color,
    /// Width } - the border the two cells of the edge resolve to (the
    /// stronger style, the cell below / to the right on a tie).
    member _.GetCellEdges(row: int, col: int) : obj =
        let describe (edge: SpreadsheetLayout.SharedEdge) : obj =
            match edge with
            | SpreadsheetLayout.EdgeBorder border ->
                createObj [
                    "Style" ==> border.Style
                    "Color" ==> (border.Color |> Option.defaultValue "#000000")
                    "Width" ==> (match borderCssParts border.Style with Some (width, _) -> width | None -> 0)
                ]
            | SpreadsheetLayout.EdgeGrid true -> box "grid"
            | _ -> null
        createObj [
            "Top" ==> describe (SpreadsheetLayout.edgeAbove model row col)
            "Right" ==> describe (SpreadsheetLayout.edgeRight model row col)
            "Bottom" ==> describe (SpreadsheetLayout.edgeBelow model row col)
            "Left" ==> describe (SpreadsheetLayout.edgeLeft model row col)
        ]

    /// Switch incremental recalculation off (every edit recalculates the
    /// whole workbook, as before 0.1.7) or on again. It is on by default.
    member _.SetIncrementalRecalculation(enabled: bool) =
        incrementalRecalculation <- enabled
        if not enabled then recalcEngine <- None

    /// What the last recalculation was: { Kind: "incremental" | "full" |
    /// "none", Reason (why it was a full one), ChangedCells, ReachedFormulas,
    /// EvaluatedFormulas (of the last incremental one), Formulas (in the
    /// dependency graph, -1 before it is built) }. For hosts and checks that want to see that an edit
    /// recalculated what it changed.
    member _.GetRecalcStats() : obj =
        createObj [
            "Kind" ==> lastRecalculation
            "Reason" ==> (if lastRecalculation = "full" then lastFullReason else "")
            "ChangedCells" ==> lastRecalcStats.ChangedCells
            "ReachedFormulas" ==> lastRecalcStats.ReachedFormulas
            "EvaluatedFormulas" ==> lastRecalcStats.EvaluatedFormulas
            "Formulas" ==> (match recalcEngine with Some engine -> SpreadsheetRecalc.formulaCount engine | None -> -1)
        ]

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
        // A workbook that was just opened is the saved one.
        this.MarkSaved()

    /// Open an XLSX file: every sheet is read and the first visible one
    /// shown; the other sheets are converted to models as they are needed
    /// (see IsSheetLoaded). `data` is an ArrayBuffer, Uint8Array or Blob. The promise
    /// rejects when the file cannot be read; the current workbook is then
    /// left untouched.
    member this.ImportXlsx(data: obj): JS.Promise<unit> =
        (loadReaderWorkbook data None).``then``(fun readerWorkbook ->
            this.InstallImportedWorkbook(Rbt.Spreadsheet.XlsxBridge.Import.importWorkbookLazy (fun _ -> false) readerWorkbook, -1))

    /// Whether the sheet at `index` has been converted to a model. After
    /// ImportXlsx only the sheet on display is; the others are converted when
    /// they are first shown, when a formula reads them, or - those with
    /// formulas of their own - at the first edit. A sheet that is not
    /// converted is written by ExportXlsx exactly as the file had it.
    member _.IsSheetLoaded(index: int) : bool =
        index >= 0 && index < sheets.Length && (index = activeSheetIndex || not (isDeferredSheet sheets.[index]))

    /// Whether the workbook came from an XLSX file whose unmodelled content
    /// (drawings, named styles, theme, kept parts) ExportXlsx will carry over.
    member _.HasXlsxSource() : bool = xlsxSource.IsSome

    /// The editor's workbook in the XLSX writer's shape (what ExportXlsx writes).
    member _.ToXlsxWorkbook() : XlsxWriter.XlsxWorkbookOut =
        let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
        // The export names the scope of a defined name by the sheet's place
        // ("sheet1" is the first sheet). Sheets keep their ids when they are
        // moved, added or deleted, so the scopes are put by place here.
        let byPlace (names: Map<SheetId option * string, 'T>) : Map<SheetId option * string, 'T> =
            let inPlace = Seq.forall (fun i -> sheetIdForIndex i = Rbt.Spreadsheet.XlsxBridge.sheetIdOfIndex i) (seq { 0 .. sheets.Length - 1 })
            if inPlace then names
            else
                names
                |> Map.toSeq
                |> Seq.choose (fun ((scope, name), value) ->
                    match scope with
                    | None -> Some ((None, name), value)
                    | Some id ->
                        let index = sheetIndexOfId id
                        if index < 0 then None else Some ((Some (Rbt.Spreadsheet.XlsxBridge.sheetIdOfIndex index), name), value))
                |> Map.ofSeq
        Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut
            { Sheets = current
              Names = Array.copy sheetNames
              Visible =
                Array.init sheets.Length (fun i ->
                    match workbook |> Option.bind (fun wb -> Map.tryFind (sheetIdForIndex i) wb.Sheets) with
                    | Some sheet -> sheet.Visible
                    | None -> true)
              ActiveSheetIndex = activeSheetIndex
              DefinedNames = (match workbook with Some wb -> byPlace wb.Names | None -> Map.empty)
              NameInfo = (match workbook with Some wb -> byPlace wb.NameInfo | None -> Map.empty)
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

    // ===== Reading cells (toolbars, host features) =====

    /// One cell of the active sheet (`row`, `col` 0-based):
    /// { row, col, address: "B2",
    ///   value: number | string | boolean | null (empty) - of a formula, its result; an error is its text ("#DIV/0!"),
    ///   valueType: "empty" | "number" | "string" | "boolean" | "error",
    ///   text: string (what the cell shows: the value through its number format),
    ///   formula: string | null (with the leading "="),
    ///   format: { bold, italic, underline, doubleUnderline, strikethrough: boolean,
    ///             fontFamily: string | null, fontSize: number | null (points; null = the default, 11),
    ///             textColor, fill: "#RRGGBB" | null,
    ///             horizontalAlign: "left" | "center" | "right" | "justify" | ... | null (general),
    ///             verticalAlign: "top" | "middle" | "bottom" | null (bottom),
    ///             wrap: boolean, indent: number, shrinkToFit: boolean,
    ///             textRotation: number (degrees, -90..90; 255 = stacked),
    ///             numberFormat: string | null (an Excel format code; null = General),
    ///             borders: { top, bottom, left, right, diagonal: { style, color } | null, diagonalUp, diagonalDown: boolean } }
    ///           - the cell's own format, else its column's default style, else the default (conditional formats are not applied),
    ///   isMerged: boolean, mergeRange: "B2:C3" | null (the merged cell it is part of; its content is in the first cell),
    ///   comment: { text, author: string | null } | null (see GetCommentThread for replies),
    ///   hyperlink: { target, isInternal: boolean, range } | null,
    ///   validation: the rule object of the validation on the cell with its `id` and `range`
    ///               (see GetValidations / GetValidationAt) | null }.
    /// Throws when row / col are not whole numbers inside the sheet.
    member _.GetCell(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetCell" model row; Col = Host.colArg "GetCell" model col }
        Host.cellToObj model addr (cellText (getCell addr model))

    /// The cells of a range as rows of cell objects (see GetCell):
    /// GetRange("B2:D5") -> [[B2, C2, D2], [B3, ...], ...]. `a1` is "B2",
    /// "B2:D5", "A:C" or "3:5", of the active sheet or, with a sheet name in
    /// front ("Data!A1:B2", "'My sheet'!A:A"), of that sheet. Whole columns
    /// and rows are cut down to the used range of the sheet; a block is cut
    /// at the sheet's last row and column. Throws when `a1` is not a range,
    /// names a sheet that does not exist, or asks for more than a million cells.
    member _.GetRange(a1: obj) : obj array array =
        let sheetName, reference = SheetOps.splitSheetPrefix (Host.stringArg "GetRange" "the range" a1)
        let source =
            match sheetName with
            | None -> model
            | Some name ->
                match sheetNames |> Array.tryFindIndex (fun known -> System.String.Equals(known, name, System.StringComparison.OrdinalIgnoreCase)) with
                | Some index -> sheetModelAt index
                | None -> Host.hostFail "GetRange" (sprintf "there is no sheet \"%s\"" name)
        let range = SheetOps.parseRange "GetRange" reference
        let range = if SheetOps.isWholeColumns range || SheetOps.isWholeRows range then SheetOps.clipToUsed source range else range
        match SheetOps.clipRange source.RowCount source.ColCount range with
        | None -> [||]
        | Some range ->
            let rows = range.To.Row - range.From.Row + 1
            let cols = range.To.Col - range.From.Col + 1
            if float rows * float cols > 1.0e6 then
                Host.hostFail "GetRange" (sprintf "%s holds %g cells; ask for at most a million at a time" (SheetOps.rangeText range) (float rows * float cols))
            Array.init rows (fun r ->
                Array.init cols (fun c ->
                    let addr = { Row = range.From.Row + r; Col = range.From.Col + c }
                    Host.cellToObj source addr (cellText (getCell addr source))))

    /// What the cells of the selection have in common, for a toolbar: the
    /// `format` object of GetCell with "mixed" in place of every property in
    /// which the cells differ - { bold, italic, underline, doubleUnderline,
    /// strikethrough, fontFamily, fontSize, textColor, fill, horizontalAlign,
    /// verticalAlign, wrap, indent, shrinkToFit, textRotation, numberFormat,
    /// borders: { top, bottom, left, right, diagonal } } (each border side
    /// { style, color }, null or "mixed"). Made to be called on every
    /// selection change: a selection of whole rows, columns or the sheet
    /// costs one pass over the cells that exist, and asking again for the
    /// same selection of unchanged cells costs nothing.
    member _.GetSelectionFormat() : obj =
        let block = selectionBounds()
        let top, left, bottom, right = block
        let block = top, left, min bottom (model.RowCount - 1), min right (model.ColCount - 1)
        match selectionFormatFor with
        | Some (cells, styles, known) when sameRef cells (box model.Cells) && sameRef styles (box model.ColumnStyles) && known = block -> selectionFormatObj
        | _ ->
            let summary = Host.formatSummaryToObj (SheetOps.summarizeFormats model block)
            selectionFormatFor <- Some (box model.Cells, box model.ColumnStyles, block)
            selectionFormatObj <- summary
            summary

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
    /// Returns a function that removes the subscription.
    member _.OnValidationNotice(listener: System.Action<string, string, string>) : System.Action =
        let wrapped = fun style title message -> listener.Invoke(style, title, message)
        validationListeners.Add wrapped
        System.Action(fun () -> validationListeners.Remove wrapped |> ignore)

    /// Whether a cell is being edited in the grid.
    member _.IsEditing() : bool = model.EditingCell.IsSome

    /// Commit the entry in progress, as Enter does (the selection stays).
    /// False when the cell's validation refused it: the cell is unchanged,
    /// the entry is dropped and a notice says why. True when there was
    /// nothing to commit.
    member _.CommitEdit() : bool = finishEditing true

    /// Drop the entry in progress, as Escape does.
    member _.CancelEdit() : unit = finishEditing false |> ignore

    /// Start editing the active cell (F2), or with `text` start an entry
    /// that replaces its content, as typing does. Not in read-only mode.
    member _.StartEdit(text: string option) : unit =
        let active = activeCellAddress()
        startEditingWith active.Row active.Col text

    /// The text of the entry in progress, set from outside the grid (a
    /// formula bar): the edit box shows it and CommitEdit enters it.
    member _.SetEditText(text: string) : unit =
        if model.EditingCell.IsSome then
            model <- { model with EditBuffer = text }
            match layers with
            | Some layer -> setValue layer.EditInput text
            | None -> ()
            notifyEditState ()

    /// Call `callback` when an entry starts, when its text changes and when
    /// it ends (committed or cancelled), with
    /// { Editing: boolean, Row, Col, Text, SheetIndex } (Row and Col are -1
    /// and Text is "" once it has ended). Returns a function that removes
    /// the subscription.
    member _.OnEditStateChange(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        if editStateListeners.Count = 0 then notifiedEdit <- model.EditingCell |> Option.map (fun addr -> addr, model.EditBuffer)
        editStateListeners.Add listener
        System.Action(fun () -> editStateListeners.Remove listener |> ignore)

    /// Whether Undo() / Redo() have a step to take on the active sheet.
    member _.CanUndo() : bool = not model.History.Past.IsEmpty
    member _.CanRedo() : bool = not model.History.Future.IsEmpty

    /// The active cell { Row, Col } (0-based): the cell of the selection that
    /// typing goes to, where the selection was started.
    member _.GetActiveCell() : obj =
        let active = activeCellAddress()
        createObj [ "Row" ==> active.Row; "Col" ==> active.Col ]

    /// Scroll so that a cell is in view.
    member _.ScrollToCell(row: int, col: int) : unit =
        revealCell row col
        render()

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

    /// Select a range programmatically. (startRow, startCol) becomes the
    /// active cell; the selection is the rectangle between the two cells.
    member _.SelectRange(startRow: int, startCol: int, endRow: int, endCol: int) =
        selectFromTo { Row = startRow; Col = startCol } { Row = endRow; Col = endCol }
        revealCell startRow startCol
        render()

    /// Select a specific cell programmatically (it is scrolled into view).
    member _.SelectCell(row: int, col: int) =
        selectCellAddress { Row = row; Col = col }
        revealCell row col
        render()

    /// Select every cell of the sheet (Ctrl+A); the active cell stays.
    member _.SelectAll() =
        anchorCell <- activeCellAddress()
        model <- { model with Selection = Range (0, 0, max 0 (model.RowCount - 1), max 0 (model.ColCount - 1)) }
        render()

    /// Apply a format command to every cell of the selection, as one change
    /// and one undo step. Toggles go by the active cell (on for all when it
    /// is off there, off for all otherwise): "bold" | "italic" | "underline" |
    /// "doubleUnderline" | "strikethrough" | "wrap" | "shrinkToFit". Steps:
    /// "fontIncrease" | "fontDecrease" (2 points), "indentIncrease" |
    /// "indentDecrease", "decimalIncrease" | "decimalDecrease" (see
    /// IncreaseDecimals). Alignment: "alignLeft" | "alignCenter" |
    /// "alignRight" | "alignJustify" | "alignGeneral" (back to the default),
    /// "alignTop" | "alignMiddle" | "alignBottom". Throws on any other command.
    member this.ApplyFormat(formatType: string) =
        let active = activeFormat ()
        let toggle (description: string) (isOn: CellFormat -> bool) (set: bool -> CellFormat -> CellFormat) =
            let on = not (isOn active)
            formatSelection description (fun format -> if isOn format = on then format else set on format)
        let horizontal (description: string) (align: string option) =
            formatSelection description (fun format -> if format.HorizontalAlign = align then format else { format with HorizontalAlign = align })
        let vertical (description: string) (align: string) =
            formatSelection description (fun format -> if format.VerticalAlign = Some align then format else { format with VerticalAlign = Some align })
        let changed =
            match formatType with
            | "bold" -> toggle "Bold" (fun format -> format.Bold) (fun on format -> { format with Bold = on })
            | "italic" -> toggle "Italic" (fun format -> format.Italic) (fun on format -> { format with Italic = on })
            | "underline" -> toggle "Underline" (fun format -> format.Underline) (fun on format -> if on then { format with Underline = true } else { format with Underline = false; UnderlineDouble = false })
            | "doubleUnderline" -> toggle "Double underline" (fun format -> format.UnderlineDouble) (fun on format -> if on then { format with Underline = true; UnderlineDouble = true } else { format with UnderlineDouble = false })
            | "strikethrough" -> toggle "Strikethrough" (fun format -> format.Strikethrough) (fun on format -> { format with Strikethrough = on })
            | "wrap" -> toggle "Wrap text" (fun format -> format.WrapText) (fun on format -> { format with WrapText = on })
            | "shrinkToFit" -> toggle "Shrink to fit" (fun format -> format.ShrinkToFit) (fun on format -> { format with ShrinkToFit = on })
            | "fontIncrease" -> formatSelection "Font size" (fun format -> { format with FontSize = Some (min 409 (defaultArg format.FontSize 11 + 2)) })
            | "fontDecrease" -> formatSelection "Font size" (fun format -> { format with FontSize = Some (max 8 (defaultArg format.FontSize 11 - 2)) })
            | "indentIncrease" -> formatSelection "Increase indent" (fun format -> if format.Indent >= 250 then format else { format with Indent = format.Indent + 1 })
            | "indentDecrease" -> formatSelection "Decrease indent" (fun format -> if format.Indent <= 0 then format else { format with Indent = format.Indent - 1 })
            | "decimalIncrease" -> this.IncreaseDecimals(); false
            | "decimalDecrease" -> this.DecreaseDecimals(); false
            | "alignLeft" -> horizontal "Align left" (Some "left")
            | "alignCenter" -> horizontal "Align center" (Some "center")
            | "alignRight" -> horizontal "Align right" (Some "right")
            | "alignJustify" -> horizontal "Justify" (Some "justify")
            | "alignGeneral" -> horizontal "General alignment" None
            | "alignTop" -> vertical "Align top" "top"
            | "alignMiddle" -> vertical "Align middle" "middle"
            | "alignBottom" -> vertical "Align bottom" "bottom"
            | other -> Host.hostFail "ApplyFormat" (sprintf "unknown command \"%s\"" (if isNull other then "null" else other))
        if changed then render()

    /// Change the format of every cell of the selection (one undo step):
    /// `format` is a `format` object as GetCell gives it, read as a patch -
    /// the keys that are present are set, the others stay as each cell has
    /// them, and null puts a property back to its default. E.g.
    /// ApplyFormatPatch({ bold: true, fill: "#FFF2CC", borders: { bottom: { style: "thin", color: "#000000" } } }).
    /// Throws on an unknown key or a value of the wrong kind. A selection of
    /// more than 20 000 cells (whole rows / columns) changes the cells that
    /// exist; empty cells of it keep the default format.
    member _.ApplyFormatPatch(format: obj) =
        let patch = Host.formatPatchOfObj "ApplyFormatPatch" "format" format
        if formatSelection "Format" patch then render()

    /// The font of the selection: a family name ("Georgia"), or null for the default font.
    member _.SetFontFamily(family: obj) =
        if formatSelection "Font" (Host.formatPatchOfObj "SetFontFamily" "format" (createObj [ "fontFamily" ==> (if Host.isNullish family then null else family) ])) then render()

    /// The font size of the selection in points (1 to 409; whole points are
    /// kept), or null for the default size (11).
    member _.SetFontSize(points: obj) =
        if formatSelection "Font size" (Host.formatPatchOfObj "SetFontSize" "format" (createObj [ "fontSize" ==> (if Host.isNullish points then null else points) ])) then render()

    /// Strike the text of the selection through, or not.
    member _.SetStrikethrough(on: obj) =
        if formatSelection "Strikethrough" (Host.formatPatchOfObj "SetStrikethrough" "format" (createObj [ "strikethrough" ==> requiredArg "SetStrikethrough" "true or false" on ])) then render()

    /// Horizontal alignment of the selection: "left" | "center" | "right" |
    /// "justify" | "fill" | "centerContinuous" | "distributed", or "general" /
    /// null for the default (text left, numbers right).
    member _.SetHorizontalAlign(align: obj) =
        if formatSelection "Horizontal alignment" (Host.formatPatchOfObj "SetHorizontalAlign" "format" (createObj [ "horizontalAlign" ==> (if Host.isNullish align then null else align) ])) then render()

    /// Vertical alignment of the selection: "top" | "middle" | "bottom" |
    /// "justify" | "distributed", or null for the default.
    member _.SetVerticalAlign(align: obj) =
        if formatSelection "Vertical alignment" (Host.formatPatchOfObj "SetVerticalAlign" "format" (createObj [ "verticalAlign" ==> (if Host.isNullish align then null else align) ])) then render()

    /// Wrap the text of the selection inside its cells, or not.
    member _.SetWrapText(on: obj) =
        if formatSelection "Wrap text" (Host.formatPatchOfObj "SetWrapText" "format" (createObj [ "wrap" ==> requiredArg "SetWrapText" "true or false" on ])) then render()

    /// The indent of the selection (0 to 250 steps from the edge the text is aligned to).
    member _.SetIndent(indent: obj) =
        if formatSelection "Indent" (Host.formatPatchOfObj "SetIndent" "format" (createObj [ "indent" ==> requiredArg "SetIndent" "the indent" indent ])) then render()

    /// One indent step more / fewer for every cell of the selection.
    member this.IncreaseIndent() = this.ApplyFormat "indentIncrease"
    member this.DecreaseIndent() = this.ApplyFormat "indentDecrease"

    /// Rotate the text of the selection: degrees from -90 (reading downwards)
    /// to 90 (reading upwards), 0 for none, 255 for stacked letters.
    member _.SetTextRotation(degrees: obj) =
        if formatSelection "Text rotation" (Host.formatPatchOfObj "SetTextRotation" "format" (createObj [ "textRotation" ==> requiredArg "SetTextRotation" "the angle" degrees ])) then render()

    /// Shrink text that is wider than its cell to fit it, or not.
    member _.SetShrinkToFit(on: obj) =
        if formatSelection "Shrink to fit" (Host.formatPatchOfObj "SetShrinkToFit" "format" (createObj [ "shrinkToFit" ==> requiredArg "SetShrinkToFit" "true or false" on ])) then render()

    /// A number format for the selection: an Excel format code ("0.00",
    /// "#,##0;[Red]-#,##0", "yyyy-mm-dd", "0.0%", "@"), or null / "General"
    /// for none. Throws when the code is not one a format can be (empty, an
    /// open quote or bracket, more than four sections, over 255 characters).
    member _.SetNumberFormat(code: obj) =
        let code =
            Host.optionalArg (Host.stringArg "SetNumberFormat" "the format code") code
            |> Option.filter (fun code -> not (code.Trim().Equals("General", System.StringComparison.OrdinalIgnoreCase)))
        match code |> Option.bind SheetOps.numberFormatProblem with
        | Some problem -> Host.hostFail "SetNumberFormat" problem
        | None -> ()
        if formatSelection "Number format" (fun format -> if format.NumberFormat = code then format else { format with NumberFormat = code }) then render()

    /// One decimal place more for the numbers of the selection (one undo
    /// step): every cell that holds a number gets its own number format with
    /// a decimal more in each section ("0.0" -> "0.00", "0%" -> "0.0%"); a
    /// cell in the General format starts from the decimals its value shows
    /// (12.5 -> "0.00"). Dates, times, fractions and text are left alone.
    member _.IncreaseDecimals() =
        let step (addr: CellAddress) (format: CellFormat) =
            match getCell addr model with
            | Some data ->
                match SheetOps.changeDecimals 1 format.NumberFormat data.Value with
                | Some code -> { format with NumberFormat = Some code }
                | None -> format
            | None -> format
        if formatSelectionAt "Increase decimals" (fun _ addr format -> step addr format) then render()

    /// One decimal place fewer (see IncreaseDecimals); a number without
    /// decimals stays as it is.
    member _.DecreaseDecimals() =
        let step (addr: CellAddress) (format: CellFormat) =
            match getCell addr model with
            | Some data ->
                match SheetOps.changeDecimals -1 format.NumberFormat data.Value with
                | Some code -> { format with NumberFormat = Some code }
                | None -> format
            | None -> format
        if formatSelectionAt "Decrease decimals" (fun _ addr format -> step addr format) then render()

    /// The format of a cell, for a format painter: the `format` object of the
    /// first cell of `a1` ("B2", or a range: its top-left cell), as GetCell
    /// gives it. Hand it to ApplyFormatTo.
    member _.GetFormatOf(a1: obj) : obj =
        let range = Host.rangeArg "GetFormatOf" a1
        if range.From.Row >= model.RowCount || range.From.Col >= model.ColCount then Host.hostFail "GetFormatOf" (sprintf "%s is outside the sheet" (SheetOps.rangeText range))
        Host.formatToObj (SheetOps.effectiveFormat model range.From)

    /// Give every cell of the range `a1` the format `format` (one undo step):
    /// the whole format is replaced - keys left out of `format` are the
    /// defaults, not what the cells had - which is what a format painter
    /// does with GetFormatOf's result. Values, comments and merged cells are
    /// not touched. A range of more than 20 000 cells changes the cells that
    /// exist. Throws when `a1` is not a range of the active sheet or `format`
    /// is not a format object.
    member _.ApplyFormatTo(a1: obj, format: obj) =
        let range = Host.rangeArg "ApplyFormatTo" a1
        let wanted = Host.formatOfObj "ApplyFormatTo" "format" format
        if range.From.Row >= model.RowCount || range.From.Col >= model.ColCount then Host.hostFail "ApplyFormatTo" (sprintf "%s is outside the sheet" (SheetOps.rangeText range))
        let paint (existing: CellFormat) =
            // (What the cell shows because of other things - a filter button, a comment marker - stays.)
            { wanted with FilterButton = existing.FilterButton; HasComment = existing.HasComment; HasListDropdown = existing.HasListDropdown
                          DataBar = existing.DataBar; IconSet = existing.IconSet }
        if formatBlockAt "Paint format" (range.From.Row, range.From.Col, range.To.Row, range.To.Col) (fun _ _ existing -> paint existing) then render()

    /// Clear the formats of the selection, or of the range `a1` when given
    /// (one undo step): font, colours, alignment, number format and borders
    /// go back to the default; values, comments, merged cells and conditional
    /// formats stay.
    member _.ClearFormats(?a1: obj) =
        let block = match a1 with Some a1 when not (Host.isNullish a1) -> (let range = Host.rangeArg "ClearFormats" a1 in range.From.Row, range.From.Col, range.To.Row, range.To.Col) | _ -> selectionBounds()
        if clearBlock "Clear formats" SheetOps.ClearFormats block then render()

    /// Clear the contents of the selection, or of the range `a1` (one undo
    /// step): values and formulas go, formats and comments stay; formulas
    /// that read the cells are recalculated, and a spilled array whose
    /// formula goes is taken back.
    member _.ClearContents(?a1: obj) =
        let block = match a1 with Some a1 when not (Host.isNullish a1) -> (let range = Host.rangeArg "ClearContents" a1 in range.From.Row, range.From.Col, range.To.Row, range.To.Col) | _ -> selectionBounds()
        if clearBlock "Clear contents" SheetOps.ClearContents block then render()

    /// Clear the selection, or the range `a1`, altogether (one undo step):
    /// contents, formats and cell comments. Merged cells, conditional
    /// formats, validations, hyperlinks and comment threads have their own
    /// members and stay.
    member _.ClearAll(?a1: obj) =
        let block = match a1 with Some a1 when not (Host.isNullish a1) -> (let range = Host.rangeArg "ClearAll" a1 in range.From.Row, range.From.Col, range.To.Row, range.To.Col) | _ -> selectionBounds()
        if clearBlock "Clear all" SheetOps.ClearEverything block then render()

    /// Apply text color to every cell of the selection (one undo step).
    member _.ApplyTextColor(color: string) =
        if formatSelection "Text color" (fun format -> { format with TextColor = Some color }) then render()

    /// Apply background color to every cell of the selection (one undo step).
    member _.ApplyBgColor(color: string) =
        if formatSelection "Fill color" (fun format -> { format with BgColor = Some color }) then render()

    /// Apply a number format to every cell of the selection (one undo step).
    member _.ApplyNumberFormat(format: string) =
        if formatSelection "Number format" (fun existing -> { existing with NumberFormat = Some format }) then
            render()

    /// Apply a border to the selection, as one change and one undo step.
    /// `side`: "top" | "bottom" | "left" | "right" | "all" (that side of
    /// every cell), "outside" (the outline of the selection only), "inside"
    /// (the edges between its cells), "none" (clear every border).
    /// `style`: an Excel border style - "thin" | "medium" | "thick" |
    /// "dashed" | "dotted" | "double" | "hair" | "mediumDashed" | "dashDot" |
    /// "mediumDashDot" | "dashDotDot" | "mediumDashDotDot" | "slantDashDot" -
    /// or "none" to clear the sides named. `color`: "#RRGGBB".
    member _.ApplyBorder(side: string, style: string, color: string) =
        let color = if System.String.IsNullOrEmpty color then "#000000" else color
        if formatSelectionAt "Border" (fun block addr format -> borderFormatInBlock block side style color addr.Row addr.Col format) then render()

    /// Merge cells in a range (one undo step).
    member _.MergeCells(startRow: int, startCol: int, endRow: int, endCol: int) =
        let before = model.MergedCells
        model <- addMergedCell startRow startCol endRow endCol model
        let after = model.MergedCells
        if not (obj.ReferenceEquals(before, after)) then
            pushValueNeutral { Undo = (fun m -> { m with MergedCells = before })
                               Redo = (fun m -> { m with MergedCells = after })
                               Description = "Merge cells" }
        render()

    /// Unmerge cells at address (one undo step).
    member _.UnmergeCells(row: int, col: int) =
        let addr = { Row = row; Col = col }
        let before = model.MergedCells
        model <- removeMergedCell addr model
        let after = model.MergedCells
        if before.Length <> after.Length then
            pushValueNeutral { Undo = (fun m -> { m with MergedCells = before })
                               Redo = (fun m -> { m with MergedCells = after })
                               Description = "Unmerge cells" }
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
        calculatedWorkbook <- Some (currentWorkbookSnapshot())
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

    /// Insert `count` rows (default 1) above row `atRow` (0-based): one model
    /// update, one undo step, every formula in the workbook rewritten once.
    member this.InsertRow(atRow: int, ?count: int) = this.InsertRows(atRow, defaultArg count 1)

    /// Delete `count` rows (default 1) from row `atRow` on, as one step.
    member this.DeleteRow(atRow: int, ?count: int) = this.DeleteRows(atRow, defaultArg count 1)

    /// Insert `count` columns (default 1) before column `atCol`, as one step.
    member this.InsertColumn(atCol: int, ?count: int) = this.InsertColumns(atCol, defaultArg count 1)

    /// Delete `count` columns (default 1) from column `atCol` on, as one step.
    member this.DeleteColumn(atCol: int, ?count: int) = this.DeleteColumns(atCol, defaultArg count 1)

    /// Insert `count` rows above row `at`, as one undo step. Called without
    /// arguments it acts on the selection, as Excel's Insert does: as many
    /// rows as are selected, above the first of them. With `at` alone, one row.
    member _.InsertRows(?at: int, ?count: int) =
        let top, _, bottom, _ = selectionBounds()
        let at, count =
            match at with
            | Some at -> at, defaultArg count 1
            | None -> top, defaultArg count (bottom - top + 1)
        let at = max 0 at
        let count = max 0 count
        if count > 0 then
            applyStructureEdit (InsertRows (at, count)) (if count = 1 then sprintf "Insert row at %d" at else sprintf "Insert %d rows at %d" count at)
            render()

    /// Delete `count` rows from row `at` on, as one undo step. Called without
    /// arguments it deletes the rows of the selection.
    member _.DeleteRows(?at: int, ?count: int) =
        let top, _, bottom, _ = selectionBounds()
        let at, count =
            match at with
            | Some at -> at, defaultArg count 1
            | None -> top, defaultArg count (bottom - top + 1)
        let at = max 0 at
        let count = max 0 (min count (model.RowCount - at))
        if count > 0 then
            applyStructureEdit (DeleteRows (at, count)) (if count = 1 then sprintf "Delete row %d" at else sprintf "Delete %d rows from %d" count at)
            render()

    /// Insert `count` columns before column `at`, as one undo step. Called
    /// without arguments: as many columns as are selected, before the first.
    member _.InsertColumns(?at: int, ?count: int) =
        let _, left, _, right = selectionBounds()
        let at, count =
            match at with
            | Some at -> at, defaultArg count 1
            | None -> left, defaultArg count (right - left + 1)
        let at = max 0 at
        let count = max 0 count
        if count > 0 then
            applyStructureEdit (InsertCols (at, count)) (if count = 1 then sprintf "Insert column at %s" (columnToLetter at) else sprintf "Insert %d columns at %s" count (columnToLetter at))
            render()

    /// Delete `count` columns from column `at` on, as one undo step. Called
    /// without arguments it deletes the columns of the selection.
    member _.DeleteColumns(?at: int, ?count: int) =
        let _, left, _, right = selectionBounds()
        let at, count =
            match at with
            | Some at -> at, defaultArg count 1
            | None -> left, defaultArg count (right - left + 1)
        let at = max 0 at
        let count = max 0 (min count (model.ColCount - at))
        if count > 0 then
            applyStructureEdit (DeleteCols (at, count)) (if count = 1 then sprintf "Delete column %s" (columnToLetter at) else sprintf "Delete %d columns from %s" count (columnToLetter at))
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

    /// The conditional formats of a sheet (the active one, or the sheet at
    /// `sheetIndex`), in order of precedence:
    /// [{ id: number, range: "A1:A9" | "A:A" | "A1:A5 C1:C5" (areas separated
    ///    by spaces), priority: number (1 = first: where two matching rules
    ///    set the same property, the first wins), stopIfTrue: boolean (a
    ///    match hides the rules after it), rule }].
    /// `id` names a rule on its sheet while the workbook is open: it stays
    /// with the rule through edits, undo / redo and inserted or deleted rows
    /// and columns. A file holds no ids: after ImportXlsx a sheet's rules are
    /// numbered 1, 2, 3 ... in order of precedence, so read them again.
    /// `rule` (the same object AddConditionalFormat takes; `type` says which keys it has):
    ///   { type: "cellIs", operator: "greaterThan" | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual"
    ///            | "equal" | "notEqual" | "between" | "notBetween", formula1, formula2 (between / notBetween), format }
    ///   { type: "expression", formula, format }                 (formatted where the formula is true)
    ///   { type: "containsText" | "notContainsText" | "beginsWith" | "endsWith", text, format }
    ///   { type: "top", rank: number, percent: boolean, bottom: boolean, format }      (top / bottom N, or N percent)
    ///   { type: "aboveAverage", above: boolean, orEqual: boolean, stdDev: 0..3, format }
    ///   { type: "duplicateValues" | "uniqueValues" | "blanks" | "noBlanks" | "errors" | "noErrors", format }
    ///   { type: "dateOccurring", period: "today" | "yesterday" | "tomorrow" | "last7Days" | "thisWeek" | "lastWeek"
    ///            | "nextWeek" | "thisMonth" | "lastMonth" | "nextMonth", format }      (weeks start on Sunday)
    ///   { type: "colorScale", stops: [{ type, value, color: "#RRGGBB" }] }           (2 or 3 stops, lowest first)
    ///   { type: "dataBar", min: { type, value }, max: { type, value }, color: "#RRGGBB" }
    ///   { type: "iconSet", iconSet: "3Arrows" | "3TrafficLights1" | "4Rating" | "5Quarters" | ..., thresholds: [{ type, value }]
    ///            (one per icon, the first is the floor), reverse: boolean }
    /// A threshold's `type` is "min" | "max" | "number" | "percent" | "percentile" | "formula"; `value` a number (a formula's
    /// text for "formula", null for min and max). Formulas come without the leading "=" and are written for the top-left
    /// cell of the rule's first range: a number ("10"), a quoted text ("\"north\""), a reference ("$C$1") or any formula.
    /// `format` is { textColor, fill: "#RRGGBB" | null, bold, italic, underline, strikethrough: boolean } - what a
    /// conditional format can apply; it is null for colour scales, data bars and icon sets, which draw by themselves.
    /// Rules of a kind the editor does not model (the x14 extensions of a file) are not listed.
    member _.GetConditionalFormats(sheetIndex: obj) : obj array =
        let source =
            match Host.optionalArg (Host.intArg "GetConditionalFormats" "sheetIndex") sheetIndex with
            | None -> model
            | Some index when index >= 0 && index < sheets.Length -> sheetModelAt index
            | Some index -> Host.hostFail "GetConditionalFormats" (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        SheetOps.conditionalFormatsInOrder source
        |> List.mapi (fun index entry -> Host.conditionalFormatToObj (index + 1) entry)
        |> List.toArray

    /// Add a conditional format to cells of the active sheet (one undo step)
    /// and return its id: `AddConditionalFormat(range, rule, options)`.
    /// `range`: "B2:B20", "A:A", or several areas ("A1:A5 C1:C5"). `rule`:
    /// the rule object of GetConditionalFormats; keys with a default may be
    /// left out (percent, bottom: false; above: true; orEqual: false;
    /// stdDev: 0; a data bar's min / max / color; an icon set's thresholds;
    /// `format`: nothing). `options` (optional): { stopIfTrue: boolean }.
    /// The new rule takes precedence over the sheet's other rules, as one
    /// made in Excel does (ReorderConditionalFormats changes that). Throws
    /// on a range that is not one, an unknown rule type or key, a missing
    /// operand, a colour that is not "#RRGGBB".
    /// The older form `AddConditionalFormat(row, col, ruleType, threshold,
    /// bgColor)` - one cell, ruleType "greaterThan" | "lessThan" | "equal", a
    /// fill - still works (and is an undo step too).
    member _.AddConditionalFormat(target: obj, second: obj, third: obj, threshold: obj, bgColor: obj) : int =
        let name = "AddConditionalFormat"
        let id =
            if Host.isJsNumber target then
                let addr = { Row = Host.rowArg name model target; Col = Host.colArg name model second }
                let threshold = Host.numberArg name "threshold" threshold
                let rule =
                    match Host.stringArg name "ruleType" third with
                    | "greaterThan" -> GreaterThan threshold
                    | "lessThan" -> LessThan threshold
                    | "equal" -> Equal threshold
                    | other -> Host.hostFail name (sprintf "ruleType must be \"greaterThan\", \"lessThan\" or \"equal\", got \"%s\" (for other rules: AddConditionalFormat(range, rule))" other)
                addRule [ { From = addr; To = addr } ] false { Rule = rule; Style = { defaultFormat with BgColor = Some (Host.colorArg name "bgColor" bgColor) } }
            else
                let ranges = SheetOps.parseRanges name (Host.stringArg name "the range" target)
                let format = SheetOps.ruleOfDescription name defaultFormat (Host.ruleOfObj name defaultFormat (requiredArg name "the rule" second))
                let stopIfTrue =
                    if Host.isNullish third then false
                    elif not (Host.isJsObject third) then Host.hostFail name "options must be an object { stopIfTrue }"
                    else
                        for key in Host.keysOf third do
                            if key <> "stopIfTrue" then Host.hostFail name (sprintf "options has an unknown key \"%s\" (known: stopIfTrue)" key)
                        Host.optionalArg (Host.boolArg name "options.stopIfTrue") (Host.getKey third "stopIfTrue") |> Option.defaultValue false
                addRule ranges stopIfTrue format
        render()
        id

    /// Change a conditional format of the active sheet (one undo step).
    /// `changes` has one or more of: `range` (as in AddConditionalFormat),
    /// `priority` (its new place in the order of precedence, 1 = first; the
    /// other rules keep their order), `stopIfTrue`, `rule` (a whole rule
    /// object: it replaces the rule; without a `format` key the rule keeps
    /// the format it has) and `format` (a change to the format alone: only
    /// the keys given change). Throws on an unknown id or key and on what
    /// AddConditionalFormat refuses.
    member _.UpdateConditionalFormat(id: obj, changes: obj) : unit =
        let name = "UpdateConditionalFormat"
        let id = Host.intArg name "id" id
        let existing =
            match SheetOps.tryFindConditionalFormat id model with
            | Some entry -> entry
            | None -> Host.hostFail name (sprintf "the sheet has no conditional format with id %d" id)
        if not (Host.isJsObject changes) then Host.hostFail name "changes must be an object { range, priority, stopIfTrue, rule, format }"
        for key in Host.keysOf changes do
            if not (List.contains key [ "range"; "priority"; "stopIfTrue"; "rule"; "format" ]) then
                Host.hostFail name (sprintf "changes has an unknown key \"%s\" (known: range, priority, stopIfTrue, rule, format)" key)
        let given (key: string) = let value = Host.getKey changes key in if Host.isNullish value then None else Some value
        let ranges = given "range" |> Option.map (fun value -> SheetOps.parseRanges name (Host.stringArg name "changes.range" value))
        let stopIfTrue = given "stopIfTrue" |> Option.map (Host.boolArg name "changes.stopIfTrue")
        let priority = given "priority" |> Option.map (Host.intArg name "changes.priority")
        let format =
            match given "rule" with
            | Some rule -> SheetOps.ruleOfDescription name existing.Format.Style (Host.ruleOfObj name existing.Format.Style rule)
            | None -> existing.Format
        let format =
            match given "format" with
            | None -> format
            | Some _ when SheetOps.isVisualRule format.Rule -> Host.hostFail name "a colour scale, data bar or icon set draws by itself and takes no format"
            | Some patch -> { format with Style = SheetOps.ruleStyleOf (Host.ruleFormatPatchOfObj name "changes.format" patch format.Style) }
        let change (sheet: SpreadsheetModel) =
            let updated =
                SheetOps.updateConditionalFormat name id (fun entry ->
                    { entry with
                        Ranges = defaultArg ranges entry.Ranges
                        StopIfTrue = defaultArg stopIfTrue entry.StopIfTrue
                        Format = format }) sheet
            match priority with
            | Some place -> SheetOps.moveConditionalFormat name id place updated
            | None -> updated
        if changeConditionalFormats "Edit conditional format" change then render()

    /// Remove the conditional format with an id from the active sheet (one
    /// undo step). Throws when the sheet has none with that id.
    member _.RemoveConditionalFormat(id: obj) : unit =
        let id = Host.intArg "RemoveConditionalFormat" "id" id
        if changeConditionalFormats "Remove conditional format" (SheetOps.removeConditionalFormat "RemoveConditionalFormat" id) then render()

    /// Put the active sheet's conditional formats in a new order of
    /// precedence (one undo step): `ids` lists every rule of the sheet once,
    /// the one that takes precedence first. Throws on an unknown or repeated
    /// id and when a rule is left out.
    member _.ReorderConditionalFormats(ids: obj) : unit =
        let name = "ReorderConditionalFormats"
        if not (Host.isJsArray ids) then Host.hostFail name "ids must be an array of rule ids"
        let ids = unbox<obj array> ids |> Array.map (Host.intArg name "an id") |> Array.toList
        if changeConditionalFormats "Reorder conditional formats" (SheetOps.reorderConditionalFormats name ids) then render()

    /// Take the conditional formats off the active sheet (no argument), or
    /// off the cells of `range` (areas as in AddConditionalFormat): a rule
    /// that covers more cells keeps them, a rule left with none is removed.
    /// One undo step. Returns how many rules were removed or cut.
    member _.ClearConditionalFormats(range: obj) : int =
        let name = "ClearConditionalFormats"
        let holes = Host.optionalArg (fun value -> SheetOps.parseRanges name (Host.stringArg name "the range" value)) range
        let before = model.ConditionalFormats
        if changeConditionalFormats "Clear conditional formats" (SheetOps.clearConditionalFormats holes) then
            let after = model.ConditionalFormats
            render()
            before |> List.filter (fun entry -> not (after |> List.exists (fun kept -> kept.Id = entry.Id && kept.Ranges = entry.Ranges))) |> List.length
        else 0

    /// Remove the conditional formats from one cell of the active sheet (one
    /// undo step): a rule covering more than the cell keeps the rest of its range.
    member _.RemoveConditionalFormats(row: int, col: int) =
        let addr = { Row = row; Col = col }
        if changeConditionalFormats "Clear conditional formats" (removeConditionalFormats addr) then render()

    // ===== Data validation =====

    /// The data validations of a sheet (the active one, or the one at
    /// `sheetIndex`), in the order they were made:
    /// [{ id, range, rule }]. `id` names the validation on its sheet (stable
    /// through undo / redo and edits of other validations; a file's
    /// validations are numbered 1, 2, ... in file order at each import).
    /// `range` is what it covers, areas separated by spaces ("B2:B9 D2:D9";
    /// a whole column is "A:A"). `rule` (`type` says which keys count; the
    /// others are null):
    ///   { type: "any" }                                  every value (an input message only)
    ///   { type: "list", values: string[], dropdown }      the allowed values written out
    ///   { type: "list", source, dropdown }                read from cells: "$A$1:$A$9", "Lists!$B$2:$B$20" or a defined name
    ///   { type: "whole" | "decimal" | "date" | "time" | "textLength", operator, formula1, formula2 }
    ///       operator: "between" | "notBetween" (formula1 and formula2) | "equal" | "notEqual" | "greaterThan"
    ///                 | "lessThan" | "greaterThanOrEqual" | "lessThanOrEqual" (formula1)
    ///   { type: "custom", formula }                       valid where the formula is true
    /// plus, for every type: allowBlank, showInput, inputTitle, inputMessage,
    /// showError, errorStyle ("stop" refuses the entry, "warning" and
    /// "information" accept it and tell), errorTitle, errorMessage.
    /// Operands and formulas are texts without the leading "=": a number
    /// ("10"), a reference or any formula, written for the top-left cell of
    /// the validation's first range (relative references move with the cell).
    /// Dates and times read back as serial numbers.
    member _.GetValidations(sheetIndex: obj) : obj array =
        let source =
            match Host.optionalArg (Host.intArg "GetValidations" "sheetIndex") sheetIndex with
            | None -> model
            | Some index when index >= 0 && index < sheets.Length -> sheetModelAt index
            | Some index -> Host.hostFail "GetValidations" (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        source.Validations |> List.map Host.validationEntryToObj |> List.toArray

    /// The validation on a cell of the active sheet: its rule object (see
    /// GetValidations) with two more keys, `id` and `range`; null when the
    /// cell has none. The object can be given to SetValidation for other
    /// cells (`id` and `range` are ignored there). GetCell's `validation` is
    /// the same object.
    member _.GetValidationAt(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetValidationAt" model row; Col = Host.colArg "GetValidationAt" model col }
        match validationEntryAt addr model with
        | Some entry -> Host.validationAtToObj entry
        | None -> null

    /// Put a validation on cells of the active sheet, in place of whatever
    /// validation they had (one undo step; a validation that covered more
    /// cells keeps them), and return its id. `range`: "B2:B20", a whole
    /// column "A:A" (no cell is made for it), or several areas
    /// ("A1:A5 C1:C5"). `rule`: the rule object of GetValidations; keys with
    /// a default may be left out (allowBlank true, showInput true, showError
    /// true, errorStyle "stop", dropdown true). Operands may be numbers; a
    /// "date" operand may be "2024-03-31" or a Date, a "time" operand
    /// "14:30"; `value1` / `value2` are accepted for formula1 / formula2.
    /// Cells that already hold a value are not checked (see CircleInvalid).
    /// Throws for an invalid range, an unknown type, operator or key, a
    /// missing operand, formula, list or source, and a list value with a comma.
    member _.SetValidation(range: obj, rule: obj) : int =
        let name = "SetValidation"
        let ranges = SheetOps.parseRanges name (Host.stringArg name "the range" (requiredArg name "the range" range))
        let validation = SheetOps.validationOfDescription name (Host.validationOfObj name (requiredArg name "the rule" rule))
        let id = (max lastValidationId (nextValidationId model - 1)) + 1
        lastValidationId <- id
        if changeValidations "Data validation" (setValidationRangesWithId id ranges validation) then render()
        id

    /// Take the validations off cells of the active sheet (one undo step):
    /// a validation that covers more cells keeps them, one left with no cell
    /// is removed. `range` as in SetValidation. Returns how many validations
    /// were removed or cut.
    member _.ClearValidation(range: obj) : int =
        let name = "ClearValidation"
        let holes = SheetOps.parseRanges name (Host.stringArg name "the range" (requiredArg name "the range" range))
        let before = model.Validations
        if changeValidations "Clear validation" (clearValidationRanges holes) then
            let after = model.Validations
            render()
            before |> List.filter (fun entry -> not (after |> List.exists (fun kept -> kept.Id = entry.Id && kept.Ranges = entry.Ranges))) |> List.length
        else 0

    /// The values the list validation on a cell of the active sheet offers,
    /// in order (a source range is read now), or null when the cell has no
    /// list validation or its source cannot be read.
    member _.GetValidationList(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetValidationList" model row; Col = Host.colArg "GetValidationList" model col }
        match SheetOps.listValuesAt listReader addr model with
        | Some values -> box (List.toArray values)
        | None -> null

    /// Open the built-in dropdown of the active cell's list validation (what
    /// the button beside the cell and Alt+Down do). False when the cell has
    /// no list to offer or the grid is read-only.
    member _.ShowValidationList() : bool = showValidationList ()

    /// The filled cells of the active sheet whose value breaks the
    /// validation on them (Excel's "Circle Invalid Data"):
    /// [{ row, col, address: "B4", value, text, errorStyle, errorTitle, errorMessage }],
    /// in row order. Rules whose error alert is switched off count too;
    /// empty cells never do. The grid draws a red circle around each (up to
    /// 2000) until ClearInvalidCircles(), the next CircleInvalid() or a
    /// change of the sheet's validations; the circles are not part of the
    /// document (no undo step, not saved) and do not follow later edits of
    /// the cells. Pass `false` to get the list without drawing.
    member _.CircleInvalid(draw: obj) : obj array =
        let draw = Host.optionalArg (Host.boolArg "CircleInvalid" "draw") draw |> Option.defaultValue true
        refreshRuleWorkbook ()
        let found = invalidCells listReader ruleEvaluator model
        if draw then
            invalidCircles.[activeSheetIndex] <- found |> List.map fst |> List.toArray
            render()
        found
        |> List.map (fun (addr, outcome) ->
            let data = getCell addr model
            let style, title, message =
                match outcome with
                | EntryInvalid (style, title, message) ->
                    (match style with StopStyle -> "stop" | WarningStyle -> "warning" | InformationStyle -> "information"), title, message
                | EntryValid -> "stop", None, ""
            createObj [
                "row" ==> addr.Row
                "col" ==> addr.Col
                "address" ==> formatCellRef addr.Row addr.Col
                "value" ==> Host.valueToObj (match data with Some data -> data.Value | None -> CellValue.Empty)
                "text" ==> cellText data
                "errorStyle" ==> style
                "errorTitle" ==> Host.orNull title
                "errorMessage" ==> message
            ])
        |> List.toArray

    /// Remove the circles CircleInvalid drew (on every sheet).
    member _.ClearInvalidCircles() : unit =
        if invalidCircles.Count > 0 then
            invalidCircles.Clear()
            render()

    // ===== Freeze panes and view =====

    /// Freeze the rows above `row` and the columns left of `col` (0-based),
    /// as Excel's "Freeze Panes" with that cell selected: FreezeAt(1, 0)
    /// freezes the first row, FreezeAt(0, 1) the first column, FreezeAt(2, 3)
    /// two rows and three columns. FreezeAt(0, 0) unfreezes. The panes are
    /// counted from the first row and column, wherever the grid is scrolled
    /// to. One undo step. Throws when the cell is not on the sheet.
    member _.FreezeAt(row: obj, col: obj) : unit =
        let rows = Host.intArg "FreezeAt" "row" row
        let cols = Host.intArg "FreezeAt" "col" col
        if rows < 0 || rows >= model.RowCount then Host.hostFail "FreezeAt" (sprintf "row %d is not on the sheet (0..%d)" rows (model.RowCount - 1))
        if cols < 0 || cols >= model.ColCount then Host.hostFail "FreezeAt" (sprintf "column %d is not on the sheet (0..%d)" cols (model.ColCount - 1))
        if changeSheetView "Freeze panes" (fun m -> { m with FrozenRows = rows; FrozenCols = cols }) then render()

    /// Unfreeze the panes of the active sheet (one undo step; nothing when none are frozen).
    member _.Unfreeze() : unit =
        if changeSheetView "Unfreeze panes" (fun m -> { m with FrozenRows = 0; FrozenCols = 0 }) then render()

    /// The frozen panes of the active sheet: { rows, cols } (0 = none).
    member _.GetFrozen() : obj =
        createObj [ "rows" ==> model.FrozenRows; "cols" ==> model.FrozenCols ]

    // ===== Rows and columns =====
    // Sizes are sheet pixels (at 100 % zoom). Where a member takes "a row or
    // range" it takes a 0-based number or an A1 string whose rows / columns
    // are meant: "2:5", "B:D", or any range such as "B2:D9".

    /// Set the width of a column (0-based number) or of the columns of an A1
    /// range ("B:D") to `px` pixels (0..1785; 0 hides them, as HideColumns
    /// does). One undo step. Throws on a bad column, range or width.
    member _.SetColumnWidth(target: obj, px: obj) : unit =
        let first, last = axisSpanArg "SetColumnWidth" false target null
        let size = Host.numberArg "SetColumnWidth" "the width" px
        if not (size >= 0.0 && size <= SheetOps.maxColumnWidthPx) then Host.hostFail "SetColumnWidth" (sprintf "the width must be 0 to %g pixels, got %g" SheetOps.maxColumnWidthPx size)
        if changeAxes "Column width" (SheetOps.setAxisSize false first last size) then render()

    /// Set the height of a row (0-based number) or of the rows of an A1 range
    /// ("2:5") to `px` pixels (0..546; 0 hides them). One undo step.
    member _.SetRowHeight(target: obj, px: obj) : unit =
        let first, last = axisSpanArg "SetRowHeight" true target null
        let size = Host.numberArg "SetRowHeight" "the height" px
        if not (size >= 0.0 && size <= SheetOps.maxRowHeightPx) then Host.hostFail "SetRowHeight" (sprintf "the height must be 0 to %g pixels, got %g" SheetOps.maxRowHeightPx size)
        if changeAxes "Row height" (SheetOps.setAxisSize true first last size) then render()

    /// The width of a column in pixels (0 when it is hidden).
    member _.GetColumnWidth(col: obj) : float =
        SheetOps.axisSize false (Host.colArg "GetColumnWidth" model col) model

    /// The height of a row in pixels (0 when it is hidden, by HideRows or by a filter).
    member _.GetRowHeight(row: obj) : float =
        getRowHeight (Host.rowArg "GetRowHeight" model row) model

    /// Fit columns to their content: each column of `range` (a column number
    /// or an A1 range; default: the columns of the selection) gets the width
    /// of its widest cell - what the cell displays, in its font, with its
    /// indent, on one line (wrapped text: its longest line as broken by line
    /// breaks) - at most 1785 pixels. Cells merged across columns do not
    /// count, hidden rows and columns are left out, and a column without
    /// text keeps its width. A double click on the resize handle of a column
    /// header does the same. One undo step.
    member _.AutoFitColumns(?range: obj) : unit =
        let span =
            match range with
            | Some target when not (Host.isNullish target) -> axisSpanArg "AutoFitColumns" false target null
            | _ -> selectedSpan false
        if autoFitSpan false span then render()

    /// Fit rows to their content: each row of `range` (a row number or an A1
    /// range; default: the rows of the selection) gets the height of its
    /// tallest cell - the lines of its text (wrapped text: as many as it
    /// takes at the width of its column or merged columns) at the line height
    /// of its font - at least 20 and at most 546 pixels; a row without text
    /// gets 20. Cells merged across rows do not count; hidden rows stay
    /// hidden. A double click on the resize handle of a row header does the
    /// same. One undo step.
    member _.AutoFitRows(?range: obj) : unit =
        let span =
            match range with
            | Some target when not (Host.isNullish target) -> axisSpanArg "AutoFitRows" true target null
            | _ -> selectedSpan true
        if autoFitSpan true span then render()

    /// Hide rows: `target` is a row number (with `count`, that many rows) or
    /// an A1 range. A hidden row keeps its height for UnhideRows. One undo step.
    member _.HideRows(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "HideRows" true target (defaultArg count null)
        if changeAxes "Hide rows" (SheetOps.hideAxis true first last) then render()

    /// Show the hidden rows among those named (see HideRows) at the height
    /// they had. Rows hidden by a filter are the filter's. One undo step.
    member _.UnhideRows(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "UnhideRows" true target (defaultArg count null)
        if changeAxes "Unhide rows" (SheetOps.unhideAxis true first last) then render()

    /// Hide columns: a column number (with `count`) or an A1 range ("B:D"). One undo step.
    member _.HideColumns(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "HideColumns" false target (defaultArg count null)
        if changeAxes "Hide columns" (SheetOps.hideAxis false first last) then render()

    /// Show the hidden columns among those named at the width they had. One undo step.
    member _.UnhideColumns(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "UnhideColumns" false target (defaultArg count null)
        if changeAxes "Unhide columns" (SheetOps.unhideAxis false first last) then render()

    /// The hidden rows of the active sheet as A1 row ranges, in order:
    /// ["3:5", "9:9"]. Rows a filter hides are not listed.
    member _.GetHiddenRows() : string array =
        SheetOps.hiddenRuns true model |> List.map (SheetOps.axisRunText true) |> List.toArray

    /// The hidden columns of the active sheet as A1 column ranges: ["B:D"].
    member _.GetHiddenColumns() : string array =
        SheetOps.hiddenRuns false model |> List.map (SheetOps.axisRunText false) |> List.toArray

    /// Group rows (Excel's Data > Group): the outline level of the rows named
    /// (a row number with `count`, or an A1 range) goes up by one, to at most
    /// 7. The grid draws no outline gutter: a host draws its own from
    /// GetOutline and collapses with SetGroupCollapsed. One undo step.
    member _.GroupRows(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "GroupRows" true target (defaultArg count null)
        if changeAxes "Group rows" (SheetOps.changeOutlineLevel true first last 1) then render()

    /// Ungroup rows: their outline level goes down by one. One undo step.
    member _.UngroupRows(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "UngroupRows" true target (defaultArg count null)
        if changeAxes "Ungroup rows" (SheetOps.changeOutlineLevel true first last -1) then render()

    /// Group columns (see GroupRows).
    member _.GroupColumns(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "GroupColumns" false target (defaultArg count null)
        if changeAxes "Group columns" (SheetOps.changeOutlineLevel false first last 1) then render()

    /// Ungroup columns (see UngroupRows).
    member _.UngroupColumns(target: obj, ?count: obj) : unit =
        let first, last = axisSpanArg "UngroupColumns" false target (defaultArg count null)
        if changeAxes "Ungroup columns" (SheetOps.changeOutlineLevel false first last -1) then render()

    /// Collapse (true) or expand (false) the group that a row / column is in:
    /// `axis` is "rows" or "columns", `index` any row / column of the group
    /// (0-based; the group at that row's own level). Collapsing hides the
    /// group's rows; expanding shows them again, except those of groups
    /// inside it that are collapsed themselves. One undo step. Throws when
    /// the row / column is in no group.
    member _.SetGroupCollapsed(axis: obj, index: obj, collapsed: obj) : unit =
        let rows =
            match Host.stringArg "SetGroupCollapsed" "the axis" axis with
            | "rows" -> true
            | "columns" -> false
            | other -> Host.hostFail "SetGroupCollapsed" (sprintf "the axis must be \"rows\" or \"columns\", got \"%s\"" other)
        let at = if rows then Host.rowArg "SetGroupCollapsed" model index else Host.colArg "SetGroupCollapsed" model index
        let collapse = Host.boolArg "SetGroupCollapsed" "collapsed" collapsed
        if changeAxes (if collapse then "Collapse group" else "Expand group") (SheetOps.setGroupCollapsed "SetGroupCollapsed" rows at collapse) then render()

    /// The outline of the active sheet: { rows: [group], columns: [group],
    /// rowLevels, columnLevels (the deepest level, 0 = no groups) } with
    /// group = { first, last (0-based), range ("3:7" / "B:D"), level (1..7),
    /// collapsed }; outer groups come before the groups inside them.
    member _.GetOutline() : obj =
        let groups (rows: bool) =
            SheetOps.outlineGroups rows model
            |> List.map (fun (first, last, level, collapsed) ->
                createObj [ "first" ==> first; "last" ==> last; "range" ==> SheetOps.axisRunText rows (first, last)
                            "level" ==> level; "collapsed" ==> collapsed ])
            |> List.toArray
        let deepest (levels: Map<int, int>) = levels |> Map.fold (fun level _ own -> max level own) 0
        createObj [
            "rows" ==> groups true
            "columns" ==> groups false
            "rowLevels" ==> deepest model.Axes.RowLevels
            "columnLevels" ==> deepest model.Axes.ColLevels
        ]

    /// Zoom the active sheet: `factor` 0.1 .. 4 (1 = 100 %; rounded to a whole
    /// percent). Cells, headers, pictures and the edit box scale; the cell in
    /// the top-left corner stays there; pointer positions and scrolling follow.
    /// Each sheet has its own zoom, saved with it (sheetView zoomScale). As in
    /// Excel, zooming is not an edit: no undo step, no new revision, no
    /// OnChange. Throws when `factor` is not a number in that range.
    member _.SetZoom(factor: obj) : unit =
        let factor = Host.numberArg "SetZoom" "factor" factor
        if not (factor >= 0.1 && factor <= 4.0) then Host.hostFail "SetZoom" (sprintf "the factor has to be between 0.1 and 4 (1 = 100%%), not %s" (string factor))
        let zoom = System.Math.Round (factor * 100.0) / 100.0
        if zoom <> model.Zoom then
            model <- { model with Zoom = zoom }
            render()

    /// The zoom of the active sheet (1 = 100 %).
    member _.GetZoom() : float = model.Zoom

    /// Show or hide the active sheet's grid lines (one undo step).
    member _.SetShowGridLines(show: obj) : unit =
        let show = Host.boolArg "SetShowGridLines" "show" show
        if changeSheetView "Grid lines" (fun m -> { m with ShowGridLines = show }) then render()

    /// Show or hide the active sheet's row numbers and column letters (one
    /// undo step). Without them the cells start at the grid's top-left corner.
    member _.SetShowHeaders(show: obj) : unit =
        let show = Host.boolArg "SetShowHeaders" "show" show
        if changeSheetView "Headers" (fun m -> { m with ShowHeaders = show }) then render()

    /// Show formulas instead of their results on the active sheet (Excel's
    /// Ctrl+`): a formula cell shows its text, left-aligned. One undo step.
    /// What cells hold and what GetCell reports is unchanged.
    member _.SetShowFormulas(show: obj) : unit =
        let show = Host.boolArg "SetShowFormulas" "show" show
        if changeSheetView "Show formulas" (fun m -> { m with ShowFormulas = show }) then render()

    /// Show or hide zero values on the active sheet: with `false` a cell
    /// holding the number 0 (typed or calculated) shows nothing. One undo step.
    member _.SetShowZeros(show: obj) : unit =
        let show = Host.boolArg "SetShowZeros" "show" show
        if changeSheetView "Show zeros" (fun m -> { m with ShowZeros = show }) then render()

    member _.GetShowGridLines() : bool = model.ShowGridLines
    member _.GetShowHeaders() : bool = model.ShowHeaders
    member _.GetShowFormulas() : bool = model.ShowFormulas
    member _.GetShowZeros() : bool = model.ShowZeros

    /// The view of the active sheet: { zoom, showGridLines, showHeaders,
    /// showFormulas, showZeros, frozen: { rows, cols } }.
    member _.GetViewOptions() : obj =
        createObj [
            "zoom" ==> model.Zoom
            "showGridLines" ==> model.ShowGridLines
            "showHeaders" ==> model.ShowHeaders
            "showFormulas" ==> model.ShowFormulas
            "showZeros" ==> model.ShowZeros
            "frozen" ==> createObj [ "rows" ==> model.FrozenRows; "cols" ==> model.FrozenCols ] ]

    // ===== Defined names =====

    /// The workbook's defined names: [{ name, refersTo (the formula, without
    /// "="), scope (0-based index of the sheet a sheet-level name belongs
    /// to, null for a workbook-level one), scopeName (that sheet's name or
    /// null), comment (string | null), isBuiltIn (an `_xlnm.` name of the
    /// file, e.g. the range of a filter), hidden (the file hides it from the
    /// name manager) }], workbook-level names first, then by name. Print
    /// areas and print titles are not names here: they belong to the page setup.
    member _.GetDefinedNames() : obj array =
        match workbook with
        | None -> [||]
        | Some wb ->
            Rbt.Spreadsheet.Workbook.definedNames wb
            |> List.filter (fun entry -> not (entry.Name.ToUpperInvariant().StartsWith "_XLNM.PRINT_"))
            |> List.choose (fun entry ->
                let scopeIndex = entry.Scope |> Option.map sheetIndexOfId
                // (A name of a sheet that is gone is not a name any more.)
                if scopeIndex = Some -1 then None
                else
                    Some (createObj [
                            "name" ==> entry.Name
                            "refersTo" ==> entry.RefersTo
                            "scope" ==> Host.orNull scopeIndex
                            "scopeName" ==> Host.orNull (scopeIndex |> Option.map (fun index -> sheetNames.[index]))
                            "comment" ==> Host.orNull entry.Comment
                            "isBuiltIn" ==> entry.Name.ToUpperInvariant().StartsWith "_XLNM."
                            "hidden" ==> entry.Hidden ]))
            |> List.toArray

    /// Define a name, or replace the one of the same scope (names differ by
    /// more than upper / lower case): `name` as Excel accepts it (starts
    /// with a letter, `_` or `\`; letters, digits, `_`, `.`, `\`, `?`; not a
    /// cell reference such as A1 or R1C1; at most 255 characters; not the
    /// name of a table), `refersTo` a reference or formula with or without
    /// "=" ("Sheet1!$A$1:$B$9", "0.19", "OFFSET(Data!$A$1,0,0,5,1)" - a
    /// reference without a sheet is read on the sheet of the formula that
    /// uses the name), `scopeSheetIndex` the sheet of a sheet-level name
    /// (omitted or null: the workbook), `comment` a note (omitted or null:
    /// none). Formulas that use the name are recalculated. One undo step.
    /// Throws on an invalid name, formula or sheet index.
    member _.SetDefinedName(name: obj, refersTo: obj, scopeSheetIndex: obj, comment: obj) : unit =
        let memberName = "SetDefinedName"
        let name = (Host.stringArg memberName "name" name).Trim()
        let refersTo = (Host.stringArg memberName "refersTo" refersTo).Trim()
        let scopeIndex = Host.optionalArg (Host.intArg memberName "scopeSheetIndex") scopeSheetIndex
        let comment = Host.optionalArg (Host.stringArg memberName "comment") comment |> Option.filter (fun text -> text <> "")
        match scopeIndex with
        | Some index when index < 0 || index >= sheets.Length -> Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        | _ -> ()
        (match Rbt.Spreadsheet.Workbook.definedNameProblem name with
         | Some problem -> Host.hostFail memberName problem
         | None -> ())
        if name.ToUpperInvariant().StartsWith "_XLNM." then Host.hostFail memberName (sprintf "\"%s\" is a built-in name" name)
        let wb = ensureWorkbook ()
        if wb.Tables |> Map.exists (fun known _ -> System.String.Equals(known, name, System.StringComparison.OrdinalIgnoreCase)) then
            Host.hostFail memberName (sprintf "\"%s\" is the name of a table" name)
        (match SheetOps.refersToProblem refersTo with
         | Some problem -> Host.hostFail memberName problem
         | None -> ())
        let scope = scopeIndex |> Option.map sheetIdForIndex
        if changeDefinedName "Define name" scope name (Rbt.Spreadsheet.Workbook.setDefinedName scope name refersTo comment) then render()

    /// Delete a defined name: the workbook-level one, or with `scope` (a
    /// sheet index) that sheet's. Formulas that use it show #NAME?. One undo
    /// step. Throws when there is no such name, or for a built-in name.
    member _.DeleteDefinedName(name: obj, scope: obj) : unit =
        let memberName = "DeleteDefinedName"
        let name = (Host.stringArg memberName "name" name).Trim()
        let scopeIndex = Host.optionalArg (Host.intArg memberName "scope") scope
        match scopeIndex with
        | Some index when index < 0 || index >= sheets.Length -> Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        | _ -> ()
        if name.ToUpperInvariant().StartsWith "_XLNM." then Host.hostFail memberName (sprintf "\"%s\" is a built-in name" name)
        let scope = scopeIndex |> Option.map sheetIdForIndex
        let wb = ensureWorkbook ()
        match Rbt.Spreadsheet.Workbook.deleteDefinedName scope name wb with
        | None ->
            Host.hostFail memberName
                (match scopeIndex with
                 | Some index -> sprintf "the sheet \"%s\" has no name \"%s\"" sheetNames.[index] name
                 | None -> sprintf "the workbook has no name \"%s\" (a sheet-level name needs its sheet index)" name)
        | Some _ ->
            if changeDefinedName "Delete name" scope name (fun current -> Rbt.Spreadsheet.Workbook.deleteDefinedName scope name current |> Option.defaultValue current) then render()

    /// Go to a place, as the name box does: a reference ("B3", "B3:C9",
    /// "A:A", "Sheet2!B3:C9", "'My sheet'!A1"), a defined name that refers
    /// to a range (the active sheet's own name before a workbook-level one;
    /// "Sheet2!Name" for another sheet's), or the name of a table. The sheet
    /// is switched to when it is another one, the range selected (clipped to
    /// the sheet) and scrolled into view. Returns { sheetIndex, range }.
    /// Throws when the text names nothing, or a name that is not a range.
    member this.GoTo(nameOrRef: obj) : obj =
        let memberName = "GoTo"
        let text = (Host.stringArg memberName "nameOrRef" nameOrRef).Trim()
        if text = "" then Host.hostFail memberName "nothing to go to"
        let sheetNamed (name: string) =
            let index = sheetIndexOfName name
            if index < 0 then Host.hostFail memberName (sprintf "there is no sheet \"%s\"" name)
            index
        let prefix, rest = SheetOps.splitSheetPrefix text
        let homeIndex = match prefix with Some name -> sheetNamed name | None -> activeSheetIndex
        let sheetIndex, range =
            match SheetOps.tryParseRange rest with
            | Some range -> homeIndex, range
            | None ->
                let named = workbook |> Option.bind (Rbt.Spreadsheet.Workbook.getName (Some (sheetIdForIndex homeIndex)) rest)
                match named with
                | Some refersTo ->
                    // (Of several areas, the first.)
                    let body = if refersTo.StartsWith "=" then refersTo.Substring 1 else refersTo
                    let first = (SheetOps.splitAreas body) |> List.tryHead |> Option.defaultValue body
                    let targetPrefix, targetRest = SheetOps.splitSheetPrefix first
                    match SheetOps.tryParseRange targetRest with
                    | Some range -> (match targetPrefix with Some name -> sheetNamed name | None -> homeIndex), range
                    | None -> Host.hostFail memberName (sprintf "the name \"%s\" does not refer to a range (%s)" rest body)
                | None ->
                    let table =
                        workbook |> Option.bind (fun wb ->
                            wb.Tables |> Map.tryPick (fun known table -> if System.String.Equals(known, rest, System.StringComparison.OrdinalIgnoreCase) then Some table else None))
                    match table with
                    | Some table when sheetIndexOfId table.SheetId >= 0 ->
                        sheetIndexOfId table.SheetId, SheetOps.rangeOf table.StartRow table.StartCol table.EndRow table.EndCol
                    | _ -> Host.hostFail memberName (sprintf "\"%s\" is not a reference, a defined name or a table" text)
        if sheetIndex <> activeSheetIndex then this.SwitchSheet sheetIndex
        let range =
            match SheetOps.clipRange model.RowCount model.ColCount range with
            | Some clipped -> clipped
            | None -> Host.hostFail memberName (sprintf "%s is outside the sheet" (SheetOps.rangeText range))
        if model.EditingCell.IsSome then finishEditing true |> ignore
        selectFromTo range.From range.To
        revealCell range.From.Row range.From.Col
        render()
        createObj [ "sheetIndex" ==> sheetIndex; "range" ==> SheetOps.rangeText range ]

    /// Show the built-in conditional formatting dialog for the active cell.
    /// The dialog offers five kinds of rule (greater than, less than,
    /// between, equal to, text contains) on one cell; hosts that want the
    /// other kinds, ranges, priorities and stop-if-true build their own
    /// dialog on GetConditionalFormats / AddConditionalFormat /
    /// UpdateConditionalFormat / RemoveConditionalFormat.
    member _.ShowConditionalFormattingDialog() =
        showConditionalFormattingDialog()

    /// Show find/replace dialog
    member _.ShowFindReplaceDialog() =
        showFindReplaceDialog()

    /// Import CSV data
    member _.ImportCSV(csvText: string, startRow: int, startCol: int) =
        let oldCells = model.Cells
        let oldSpill = model.Spill
        model <- importCSV csvText startRow startCol model
        recalculateFormulas()

        // Undo restores the full pre-import cell map (an O(1) reference to the
        // persistent map) so dependent formulas recalculated by the import are
        // rolled back too, not just the imported rectangle.
        let command = {
            Undo = fun m -> restoreCellMapWithSpill oldCells oldSpill m
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

    /// The formula or value of the active cell, as it would be typed.
    member _.GetCurrentCellFormula(): string =
        SpreadsheetInput.editText (getCell (activeCellAddress()) model)

    /// Set the value of the active cell from a formula bar. The text is read
    /// like any typed entry and checked against the cell's validation; false
    /// when the entry was refused.
    member _.SetCurrentCellFromFormula(input: string) : bool =
        let accepted = commitInput (activeCellAddress()) input true true
        scheduleNoticeDismiss()
        render()
        accepted

    /// Undo last operation
    member _.Undo() = undoStep ()

    /// Redo last undone operation
    member _.Redo() = redoStep ()

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

    /// The selection as tab-separated text (what the cells show), the way
    /// spreadsheets exchange cells through the clipboard.
    member _.GetSelectionText() : string = selectionText ()

    /// The selection as an HTML table.
    member _.GetSelectionHtml() : string = selectionHtml ()

    /// Cut the selection: as Copy, and the cells are cleared (one undo step).
    /// In read-only mode this copies. Resolves to whether the system
    /// clipboard took the cells.
    member _.Cut() : JS.Promise<bool> =
        let text, html = copySelection true
        render()
        writeSystemClipboard text html

    /// Copy the selection to the grid's own clipboard (formulas and formats
    /// are kept for a paste within the grid) and to the system clipboard as
    /// tab-separated text and an HTML table, for other applications.
    /// Resolves to whether the system clipboard took the cells (a browser
    /// may refuse outside a user gesture; Ctrl+C in the grid always works).
    member _.Copy() : JS.Promise<bool> =
        let text, html = copySelection false
        writeSystemClipboard text html

    /// Paste at the selection: what the system clipboard holds when the
    /// browser lets the page read it (cells copied in Excel or Sheets, or
    /// the grid's own), else the grid's own clipboard. Resolves to whether
    /// anything was pasted. (Ctrl+V in the grid always reads the system
    /// clipboard.)
    member _.Paste() : JS.Promise<bool> = pasteFromSystemClipboard ()

    /// Paste clipboard content the host has in hand (from its own paste
    /// handler): `text` is the text/plain flavour, `html` the text/html one
    /// ("" when missing). One undo step; false when nothing was pasted.
    member _.PasteText(text: string, html: string) : bool =
        let pasted = pasteFlavours text (if isNull html then "" else html)
        if pasted then render()
        pasted

/// Export for JavaScript
[<ExportDefault>]
let createSpreadsheet (container: Element): Spreadsheet =
    new Spreadsheet(container)
