module Rbt.Spreadsheet.Spreadsheet

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Spreadsheet.Dom
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
module SpreadsheetCharts = Rbt.Spreadsheet.Charts
module SpreadsheetChartAuthoring = Rbt.Spreadsheet.ChartAuthoring
module SpreadsheetSlicers = Rbt.Spreadsheet.Slicers
module Host = Rbt.Spreadsheet.HostShapes

// Additional helpers
[<Emit("$0.scrollLeft")>]
let private getScrollLeft (el: Element): float = jsNative

[<Emit("$0.clientWidth")>]
let private getClientWidth (el: Element): float = jsNative

[<Emit("$0.scrollHeight")>]
let private getScrollHeight (el: Element): float = jsNative

[<Emit("$0.scrollWidth")>]
let private getScrollWidth (el: Element): float = jsNative

[<Emit("$0.clientX")>]
let private getClientX (e: obj): float = jsNative

[<Emit("$0.clientY")>]
let private getClientY (e: obj): float = jsNative
[<Emit("$0.deltaX")>]
let private wheelDeltaX (e: obj): float = jsNative

[<Emit("$0.deltaY")>]
let private wheelDeltaY (e: obj): float = jsNative

[<Emit("$0.deltaMode")>]
let private wheelDeltaMode (e: obj): int = jsNative


[<Emit("(function (el) { const r = el.getBoundingClientRect(); return [r.left, r.top, r.width, r.height]; })($0)")>]
let private boundingRect (el: Element): float array = jsNative

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

/// Run `callback` as a task of its own, at once: the page paints and takes input before it, as after a
/// setTimeout, but without the 4 ms a chain of setTimeouts is held back each time (a macro's slices).
[<Emit("(() => { const channel = new MessageChannel(); channel.port1.onmessage = () => { channel.port1.close(); $0(); }; channel.port2.postMessage(null); })()")>]
let private yieldThen (callback: unit -> unit): unit = jsNative

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

[<Emit("$0.querySelector($1)")>]
let private queryOne (root: Element) (selector: string): Element = jsNative
[<Emit("document.activeElement")>]
let private activeElement (): Element = jsNative

/// A picture being moved or sized with the mouse.
type private PictureDrag = {
    Id: int
    /// "move", or the handle: "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w".
    Handle: string
    /// Where the pointer went down (page pixels).
    StartX: float
    StartY: float
    /// The picture's rectangle then, and where the drag has it now (sheet pixels).
    Start: SpreadsheetLayout.Rect
    mutable Current: SpreadsheetLayout.Rect
    mutable Moved: bool
}

[<Emit("new Date().toISOString().slice(0, 22)")>]
let private commentDateNow (): string = jsNative

[<Emit("($0.buttons | 0)")>]
let private pointerButtons (e: obj): int = jsNative

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

/// The CSS font family a cell's text is drawn in: its own family, else the grid's (Arial).
/// (A name is all a family can be here: nothing that ends the declaration.)
let private cellFontFamily (fontName: string option) : string =
    match fontName with
    | Some family when family <> "" -> "'" + (family |> Portable.Str.replace "'" "" |> Portable.Str.replace ";" "" |> Portable.Str.replace "\\" "") + "', Arial, sans-serif"
    | _ -> "Arial, sans-serif"

/// Whether the browser has a font family (its text is not drawn in the next family of the list).
/// Asked once per family.
// (What was measured before a document font finished loading is forgotten when it has: the
// families, and the widths measured in them.)
[<Emit("(function(family){const g=globalThis.__rbSpreadsheetFamilies||(globalThis.__rbSpreadsheetFamilies=new Map());if(!globalThis.__rbSpreadsheetFontsWatched&&typeof document!=='undefined'&&document.fonts&&document.fonts.addEventListener){globalThis.__rbSpreadsheetFontsWatched=true;document.fonts.addEventListener('loadingdone',()=>{const f=globalThis.__rbSpreadsheetFamilies;if(f)f.clear();for(const k of ['__rbSpreadsheetCellMeasure','__rbSpreadsheetPlainMeasure']){const m=globalThis[k];if(m){m.widths.clear();m.font='';}}});}let v=g.get(family);if(v===undefined){const c=document.createElement('canvas').getContext('2d');const t='mmmmmmmmmmlli WQ 0123456789';c.font='72px monospace';const a=c.measureText(t).width;c.font=\"72px '\"+family+\"', monospace\";v=c.measureText(t).width!==a;g.set(family,v);}return v;})($0)")>]
let private fontFamilyAvailable (family: string) : bool = jsNative

/// Ask a host's loader for the bytes of a face (`SetFontLoader`) and make them a document font; once
/// it is loaded, forget the widths measured without it and call `loaded`. A face the loader has not
/// (null, a failure) is left to the families after it.
[<Emit("(function(loader,family,bold,italic,loaded){Promise.resolve().then(()=>loader(family,bold,italic)).then((data)=>{if(!data)return;const face=new FontFace(family,data,{weight:bold?'700':'400',style:italic?'italic':'normal'});document.fonts.add(face);return face.load().then(()=>{const g=globalThis.__rbSpreadsheetCellMeasure;if(g){g.widths.clear();g.font='';}const known=globalThis.__rbSpreadsheetFamilies;if(known)known.clear();loaded();});}).catch(()=>{});})($0,$1,$2,$3,$4)")>]
let private loadFace (loader: obj) (family: string) (bold: bool) (italic: bool) (loaded: unit -> unit) : unit = jsNative

/// Calibri's x-height in ems (Arial's is 0.5186). Calibri, Excel's default font, is drawn
/// where it is not at hand in the next family (Arial) as tall as Calibri's letters are (CSS
/// `font-size-adjust`): 90 % of the size, which gives Calibri's widths within a few percent
/// too, so a number that fits its cell in Excel fits it here.
[<Literal>]
let private calibriAspect = 0.4644

let private isCalibri (fontName: string option) =
    match fontName with
    | Some family -> family.Trim().ToLowerInvariant() = "calibri"
    | None -> false

/// The size a cell's text is drawn at, for one of its size in pixels: smaller for Calibri where
/// the browser has no Calibri (see `calibriAspect`).
let private drawnFontScale (fontName: string option) =
    if isCalibri fontName && not (fontFamilyAvailable "Calibri") then calibriAspect / 0.5186 else 1.0

/// A CSS pixel size, to a hundredth of a pixel ("14.67px").
let private cssPx (size: float) : string = string (System.Math.Round (size * 100.0) / 100.0) + "px"

// The width of a cell's text as the grid draws it: in the cell's font family
// (`cellFontFamily`), at its size in pixels. Widths are remembered: the cells
// of a sheet repeat their texts, and setting a canvas font costs more than
// measuring with it.
[<Emit("(function(text,fontSize,family,bold,italic){const g=globalThis.__rbSpreadsheetCellMeasure||(globalThis.__rbSpreadsheetCellMeasure={ctx:document.createElement('canvas').getContext('2d'),font:'',widths:new Map()});const font=(bold?'bold ':'')+(italic?'italic ':'')+fontSize+'px '+family;const key=font+'|'+text;let width=g.widths.get(key);if(width===undefined){if(g.font!==font){g.ctx.font=font;g.font=font;}width=g.ctx.measureText(text).width;if(g.widths.size>20000)g.widths.clear();g.widths.set(key,width);}return width;})($0,$1,$2,$3,$4)")>]
let private cellTextWidthPx (text: string) (fontPx: float) (family: string) (bold: bool) (italic: bool): float = jsNative

/// The width in pixels of a cell's text of a font size in points (11 pt is 14.67 px, as
/// Excel draws it at 100 %) and family, as the grid draws it.
let private cellTextWidth (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool): float =
    cellTextWidthPx text (SheetOps.pointsToPx fontPt * drawnFontScale fontName) (cellFontFamily fontName) bold italic

// The width of a text as the browser draws it, without kerning (GDI does not kern a cell's text),
// at a size in pixels. Remembered as `cellTextWidthPx`.
[<Emit("(function(text,fontSize,family,bold,italic){const g=globalThis.__rbSpreadsheetPlainMeasure||(globalThis.__rbSpreadsheetPlainMeasure={ctx:document.createElement('canvas').getContext('2d'),font:'',widths:new Map()});const font=(bold?'bold ':'')+(italic?'italic ':'')+fontSize+'px '+family;const key=font+'|'+text;let width=g.widths.get(key);if(width===undefined){if(g.font!==font){g.ctx.font=font;g.font=font;try{g.ctx.fontKerning='none';}catch(e){}}width=g.ctx.measureText(text).width;if(g.widths.size>20000)g.widths.clear();g.widths.set(key,width);}return width;})($0,$1,$2,$3,$4)")>]
let private plainTextWidthPx (text: string) (fontPx: float) (family: string) (bold: bool) (italic: bool): float = jsNative

/// The family the tables of `ExcelText` know a cell's font by (the grid draws a cell without one in Arial).
let private excelFamily (fontName: string option) =
    match fontName with
    | Some family when family.Trim() <> "" -> family.Trim()
    | _ -> "Arial"

/// What text Excel draws in Small Fonts (`ExcelText.isSmallFonts`, under 8 points on screen) is drawn
/// in: a plain, wide sans serif made for small sizes, as Small Fonts is, whatever the cell's font.
let private smallFontsFamily = "Verdana, 'DejaVu Sans', Tahoma, sans-serif"

/// The CSS family and the size factor (`drawnFontScale`) a cell's text is drawn in at a zoom.
let private drawnFace (fontPt: float) (zoom: float) (fontName: string option) : string * float =
    if ExcelText.isSmallFonts fontPt zoom then smallFontsFamily, 1.0
    else cellFontFamily fontName, drawnFontScale fontName

/// The width in screen pixels of a cell's text as Excel draws it at a zoom (`ExcelText`: at a whole
/// pixel size, in GDI's advances; a character the tables do not know as the browser draws it there).
let private excelTextWidth (zoom: float) (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool): float =
    let px = ExcelText.pixelSize fontPt zoom
    let family, scale = drawnFace fontPt zoom fontName
    ExcelText.textWidth (fun s -> plainTextWidthPx s (float px * scale) family bold italic) (excelFamily fontName) bold italic px (ExcelText.isSmallFonts fontPt zoom) text

/// The letter spacing (screen pixels a character) that makes the browser's line of `text`, drawn at
/// Excel's whole pixel size (`drawnFace`), as wide as Excel's (`excelTextWidth`). 0 where the browser
/// draws it so already.
let private excelLetterSpacing (zoom: float) (text: string) (fontPt: float) (fontName: string option) (bold: bool) (italic: bool): float =
    if text.Length = 0 then 0.0
    else
        let px = ExcelText.pixelSize fontPt zoom
        let family, scale = drawnFace fontPt zoom fontName
        let drawn = plainTextWidthPx text (float px * scale) family bold italic
        let spacing = (excelTextWidth zoom text fontPt fontName bold italic - drawn) / float text.Length
        if abs spacing < 0.005 then 0.0 else spacing

/// Call `f` whenever the document's fonts finish loading (a face added to the page and loaded once
/// a text asks for it); gives what stops it.
[<Emit("(function(f){if(typeof document==='undefined'||!document.fonts||!document.fonts.addEventListener)return function(){};const h=function(){f();};document.fonts.addEventListener('loadingdone',h);return function(){document.fonts.removeEventListener('loadingdone',h);};})($0)")>]
let private onFontsLoaded (f: unit -> unit) : (unit -> unit) = jsNative

/// A CSS length to a thousandth of a pixel.
let private cssPxFine (size: float) : string = string (System.Math.Round (size * 1000.0) / 1000.0) + "px"


// ---- macros in the browser -------------------------------------------------------------------------
[<Emit("window.alert($0)")>]
let private alertBox (_: string) : unit = jsNative
[<Emit("window.confirm($0)")>]
let private confirmBox (_: string) : bool = jsNative
[<Emit("window.prompt($0, $1)")>]
let private promptBox (_: string) (_: string) : string = jsNative
[<Emit("console.log($0)")>]
let private consoleLog (_: string) : unit = jsNative
[<Emit("$0($1, $2, $3)")>]
let private callHook3 (_: obj) (_: obj) (_: obj) (_: obj) : obj = jsNative
[<Emit("$0($1)")>]
let private callHook1 (_: obj) (_: obj) : unit = jsNative

/// MsgBox / InputBox / Debug.Print with the browser's own dialogs, until a host gives its own
/// (SetMacroHooks). MsgBox's answer is VBA's: vbOK 1, vbCancel 2, vbYes 6, vbNo 7, vbRetry 4.
let private browserMacroHooks : Rbt.Spreadsheet.MacroApi.MacroHooks =
    { MessageBox = fun prompt buttons title ->
        let text = if title = "" || title = "Microsoft Excel" then prompt else title + "\n\n" + prompt
        match buttons &&& 7 with
        | 0 -> alertBox text; 1
        | 1 -> if confirmBox text then 1 else 2
        | 3 | 4 -> if confirmBox text then 6 else 7
        | 5 -> if confirmBox text then 4 else 2
        | _ -> if confirmBox text then 4 else 3
      InputBox = fun prompt _ defaultText ->
        let answer = promptBox prompt defaultText
        if isNull answer then None else Some answer
      Output = consoleLog
      FileName = "Book1.xlsm"
      OnTime = fun _ _ _ -> ()
      OnKey = fun _ _ -> () }

/// A key as Application.OnKey names it, made comparable: modifiers (^ Ctrl, + Shift, % Alt) in that order,
/// then the key's name in capitals ("{LEFT}" and "{left}" are "LEFT", "~" is "ENTER", "a" is "A").
let private onKeyName (ctrl: bool) (shift: bool) (alt: bool) (key: string) =
    (if ctrl then "^" else "") + (if shift then "+" else "") + (if alt then "%" else "") + key.ToUpperInvariant()

let private parseOnKey (text: string) =
    let mutable i = 0
    let mutable ctrl, shift, alt = false, false, false
    while i < text.Length - 1 && (text.[i] = '^' || text.[i] = '+' || text.[i] = '%') do
        (match text.[i] with '^' -> ctrl <- true | '+' -> shift <- true | _ -> alt <- true)
        i <- i + 1
    let rest = text.Substring i
    let name =
        if rest.Length > 2 && rest.StartsWith "{" && rest.EndsWith "}" then
            match rest.Substring(1, rest.Length - 2).ToUpperInvariant() with
            | "BS" | "BKSP" -> "BACKSPACE" | "DEL" -> "DELETE" | "ESCAPE" -> "ESC" | "RETURN" -> "ENTER"
            | "INS" -> "INSERT" | other -> other
        elif rest = "~" then "ENTER"
        else rest
    onKeyName ctrl shift alt name

/// The OnKey name of a key pressed (KeyboardEvent.key), when it has one.
let private pressedOnKey (key: string) (ctrl: bool) (shift: bool) (alt: bool) =
    let name =
        match key with
        | "ArrowLeft" -> Some "LEFT" | "ArrowRight" -> Some "RIGHT" | "ArrowUp" -> Some "UP" | "ArrowDown" -> Some "DOWN"
        | "Enter" -> Some "ENTER" | "Escape" -> Some "ESC" | "Tab" -> Some "TAB" | "Backspace" -> Some "BACKSPACE"
        | "Delete" -> Some "DELETE" | "Home" -> Some "HOME" | "End" -> Some "END" | "PageUp" -> Some "PGUP"
        | "PageDown" -> Some "PGDN" | "Insert" -> Some "INSERT" | " " -> Some " "
        | k when k.Length = 1 -> Some k
        | k when k.Length <= 3 && k.StartsWith "F" -> Some k
        | _ -> None
    name |> Option.map (onKeyName ctrl shift alt)

/// The VBA engine and Excel's object model (MacroProject.fs): fetched when macros are first enabled, run or
/// edited - the grid never imports it otherwise, so a workbook that runs no macro does not load it.
#if FABLE_COMPILER
[<Emit("import(\"./MacroProject.js\").then(m => m.openProject($0))")>]
let private loadMacroProject (_preserved: XlsxReader.WorkbookPreservation option) : JS.Promise<Rbt.Spreadsheet.MacroApi.IMacroProject> = jsNative
#else
let private loadMacroProject (preserved: XlsxReader.WorkbookPreservation option) : JS.Promise<Rbt.Spreadsheet.MacroApi.IMacroProject> =
    Fable.Core.JS.Constructors.Promise.resolve (Rbt.Spreadsheet.MacroProject.openProject preserved)
#endif

/// The XLSX package writer (XlsxPackageWriter.fs, dist/office-formats/src/XlsxPackageWriter.js): fetched
/// when a workbook is first written (ExportXlsx) - the grid never imports it otherwise, so opening, showing
/// and editing a workbook do not load it. (What ExportXlsx hands it, XlsxWriter's workbook, is made here.)
#if FABLE_COMPILER
type private XlsxPackageWriterModule =
    abstract writeWorkbookPackage: XlsxWriter.XlsxWorkbookOut -> JS.Promise<obj>
    abstract writeWorkbookPackageReusing: obj * (string * byte[]) list * XlsxWriter.XlsxWorkbookOut -> JS.Promise<obj>

[<Emit("import(\"../office-formats/src/XlsxPackageWriter.js\")")>]
let private importXlsxPackageWriter () : JS.Promise<XlsxPackageWriterModule> = jsNative

/// The package of `workbook` (a JSZip archive), written by the package writer once it has loaded;
/// with `source` (the opened file's archive), `reuse` are entries of it written as they were.
let private writeXlsxPackage (source: obj) (reuse: (string * byte[]) list) (workbook: XlsxWriter.XlsxWorkbookOut) : JS.Promise<obj> =
    (importXlsxPackageWriter ()).``then``(fun writer ->
        if isNull source then writer.writeWorkbookPackage workbook
        else writer.writeWorkbookPackageReusing (source, reuse, workbook))
    |> unbox<JS.Promise<obj>>
#else
let private writeXlsxPackage (source: obj) (reuse: (string * byte[]) list) (workbook: XlsxWriter.XlsxWorkbookOut) : JS.Promise<obj> =
    if isNull source then XlsxPackageWriter.writeWorkbookPackage workbook
    else XlsxPackageWriter.writeWorkbookPackageReusing source reuse workbook
#endif

// ---- function providers (Formula.fs, "The function table") ----------------------------------
[<Emit("($0 != null && typeof $0.then === 'function')")>]
let private isThenable (value: obj) : bool = jsNative

[<Emit("$0.then($1, $2)")>]
let private whenSettled (promise: obj) (fulfilled: obj -> unit) (rejected: obj -> unit) : unit = jsNative

[<Emit("$0($1)")>]
let private callHostFunction (f: obj) (argument: obj) : obj = jsNative

[<Emit("new Promise((resolve) => $0(() => resolve()))")>]
let private promiseOfReady (start: (unit -> unit) -> unit) : JS.Promise<unit> = jsNative

/// A value a provider gave: a number, text, a boolean, null (empty) or { error: "#N/A" }.
let private providerValueOf (value: obj) : CellValue =
    if Host.isNullish value then CellValue.Empty
    else
        match Host.jsTypeOf value with
        | "number" ->
            let n = unbox<float> value
            if System.Double.IsNaN n || System.Double.IsInfinity n then CellValue.Error Rbt.Spreadsheet.Formula.Err.num else CellValue.Number n
        | "string" -> CellValue.String (unbox<string> value)
        | "boolean" -> CellValue.Boolean (unbox<bool> value)
        | "object" when not (Host.isNullish (value?error)) && Host.jsTypeOf (value?error) = "string" ->
            CellValue.Error ((unbox<string> (value?error)).TrimStart '#')
        | _ -> CellValue.Error Rbt.Spreadsheet.Formula.Err.value

/// What a provider gave as rows: one value, an array (one row) or an array of rows.
let private providerRowsOf (value: obj) : CellValue[][] =
    if Host.isJsArray value then
        let items = unbox<obj[]> value
        if items.Length > 0 && Host.isJsArray items.[0] then
            let rows = items |> Array.map (fun row -> unbox<obj[]> row |> Array.map providerValueOf)
            let width = rows |> Array.map Array.length |> Array.fold max 0
            rows |> Array.map (fun row -> Array.init width (fun c -> if c < row.Length then row.[c] else CellValue.Error Rbt.Spreadsheet.Formula.Err.na))
        else [| items |> Array.map providerValueOf |]
    else [| [| providerValueOf value |] |]

/// A value handed to a provider: as `providerValueOf` takes it.
let private providerValueJs (value: CellValue) : obj =
    match unwrapFormulaValue value with
    | CellValue.Error code -> createObj [ "error" ==> ("#" + code) ]
    | other -> Host.valueToObj other

/// An argument handed to a provider: one value, or an array of rows.
let private providerArgumentJs (rows: CellValue[][]) : obj =
    if rows.Length = 1 && rows.[0].Length = 1 then providerValueJs rows.[0].[0]
    else box (rows |> Array.map (Array.map providerValueJs))

// ---- PDF in the browser ----------------------------------------------------------------------
/// The spreadsheet's PDF writer (`src/Spreadsheet.Pdf`, the bundle dist/spreadsheet-pdf, which shares this
/// bundle's modules): loaded when a PDF is first asked for, not with the grid.
type private SheetPdfModule =
    abstract exportSheets: string * string * obj array * obj array -> byte[]
    abstract exportSheetsWithDefaultFonts: string * string * obj array * obj array -> JS.Promise<byte[]>
    abstract pdfSheet: SpreadsheetModel * string * (int * int * float * float * float * float * SpreadsheetPaint.PaintItem list) list -> obj
    abstract fontFace: string * string * bool * bool * byte[] -> obj
    abstract pngPictures: obj list -> byte[] array
    abstract withPictures: (byte[] -> (string * byte[]) option) * obj list -> obj list

[<Emit("import(\"../spreadsheet-pdf/Pdf.js\")")>]
let private importSheetPdf () : JS.Promise<SheetPdfModule> = jsNative

/// A PNG as JPEG, drawn by the browser over white (the PDF writer's browser build embeds JPEG as it is but
/// cannot decode PNG); null when the browser cannot read it.
[<Emit("(async (data) => { try { const bitmap = await createImageBitmap(new Blob([data])); const canvas = document.createElement('canvas'); canvas.width = bitmap.width; canvas.height = bitmap.height; const ctx = canvas.getContext('2d'); ctx.fillStyle = '#FFFFFF'; ctx.fillRect(0, 0, canvas.width, canvas.height); ctx.drawImage(bitmap, 0, 0); const jpeg = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', 0.92)); return jpeg ? new Uint8Array(await jpeg.arrayBuffer()) : null; } catch (_) { return null; } })($0)")>]
let private pngAsJpeg (data: byte[]) : JS.Promise<byte[]> = jsNative

[<Emit("Promise.all($0)")>]
let private allPromises (promises: JS.Promise<'T> array) : JS.Promise<'T array> = jsNative

/// Bytes a host hands (a Uint8Array, an ArrayBuffer or an array of numbers) as a Uint8Array; null otherwise.
[<Emit("($0 instanceof Uint8Array ? $0 : ($0 instanceof ArrayBuffer ? new Uint8Array($0) : (Array.isArray($0) ? Uint8Array.from($0) : null)))")>]
let private bytesOfHost (value: obj) : byte[] = jsNative

/// Print a PDF through the browser: in a hidden frame of the page, with the browser's print dialog; resolves
/// once the dialog was opened (the frame is taken away a minute later, or when the next print starts).
[<Emit("(new Promise((resolve, reject) => { const old = document.querySelector('iframe[data-spreadsheet-print]'); if (old) old.remove(); const url = URL.createObjectURL(new Blob([$0], { type: 'application/pdf' })); const frame = document.createElement('iframe'); frame.setAttribute('data-spreadsheet-print', ''); frame.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden'; frame.onload = () => { try { frame.contentWindow.focus(); frame.contentWindow.print(); resolve(); } catch (e) { reject(e); } setTimeout(() => { frame.remove(); URL.revokeObjectURL(url); }, 60000); }; frame.src = url; document.body.appendChild(frame); }))")>]
let private printPdfInFrame (pdf: byte[]) : JS.Promise<unit> = jsNative

[<Emit("$0.fillStyle = $5; $0.font = ($6 ? 'bold ' : '') + ($7 ? 'italic ' : '') + $4 + 'px ' + $9 + ', sans-serif'; $0.textBaseline = 'alphabetic'; $0.textAlign = $8; $0.fillText($1, $2, $3)")>]
let private canvasText (ctx: obj) (text: string) (x: float) (y: float) (fontSize: float) (color: string) (bold: bool) (italic: bool) (align: string) (fontName: string): unit = jsNative

// Draw text rotated `angle` radians about an anchor (tx,ty): translate to the
// anchor, rotate (canvas rotate is clockwise for +angle, matching screen
// y-down), draw left-aligned at the origin, restore. Used for cell-text
// rotation; the caller positions the anchor so the rotated bbox aligns.
[<Emit("$0.save(); $0.translate($3, $4); $0.rotate($5); $0.fillStyle = $6; $0.font = ($7 ? 'bold ' : '') + ($8 ? 'italic ' : '') + $2 + 'px ' + $9 + ', sans-serif'; $0.textBaseline = 'alphabetic'; $0.textAlign = 'left'; $0.fillText($1, 0, 0); $0.restore()")>]
let private canvasTextRotated (ctx: obj) (text: string) (fontSize: float) (tx: float) (ty: float) (angle: float) (color: string) (bold: bool) (italic: bool) (fontName: string): unit = jsNative

[<Emit("$0.innerHTML = $1")>]
let private setInnerHtml (el: Element) (html: string): unit = jsNative

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

[<Emit("(($0.__rbtPictureUrl) ? ($0.__rbtPictureElement?.removeAttribute('src'), URL.revokeObjectURL($0.__rbtPictureUrl), delete $0.__rbtPictureUrl, delete $0.__rbtPictureElement) : undefined)")>]
let private releasePicture (data: byte[]): unit = jsNative

[<Emit("$0.removeAttribute('src'); $0.querySelectorAll('img').forEach(img => img.removeAttribute('src'))")>]
let private clearImageSources (element: Element): unit = jsNative

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

/// A DOM handler that reports what the handler throws and throws it on; it
/// remembers the handler it stands for (see `wrapsHandler`).
[<Emit("(function (handler, report) { const guarded = function (e) { try { return handler(e); } catch (error) { report(error); throw error; } }; guarded.__rbtHandler = handler; return guarded; })($0, $1)")>]
let private guardedHandler (handler: obj) (report: obj -> unit): obj = jsNative

[<Emit("($0 === $1 || $0.__rbtHandler === $1)")>]
let private wrapsHandler (registered: obj) (handler: obj): bool = jsNative

[<Emit("(($0 && typeof $0.message === 'string') ? $0.message : String($0))")>]
let private errorMessageOf (error: obj): string = jsNative

/// A promise that tells `report` why it was rejected, and stays rejected.
[<Emit("$0.catch(function (error) { $1(error); throw error; })")>]
let private reportRejection (promise: JS.Promise<'T>) (report: obj -> unit): JS.Promise<'T> = jsNative

/// A payload property worked out when it is first read.
[<Emit("Object.defineProperty($0, $1, { enumerable: true, configurable: true, get: function () { const value = $2(); Object.defineProperty(this, $1, { enumerable: true, configurable: true, writable: true, value: value }); return value; } })")>]
let private defineLazy (target: obj) (key: string) (compute: unit -> obj): unit = jsNative

[<Emit("Date.now()")>]
let private nowMs (): float = jsNative

/// The middle of an element on the page: [clientX, clientY].
[<Emit("(function (el) { const r = el.getBoundingClientRect(); return [r.left + r.width / 2, r.top + r.height / 2]; })($0)")>]
let private clientCenterOf (el: Element): float array = jsNative

/// What the package writer resolves to (a Blob in the browser) as bytes.
[<Emit("$0.then((written) => (written instanceof Uint8Array ? written : (written && typeof written.arrayBuffer === 'function' ? written.arrayBuffer().then((buffer) => new Uint8Array(buffer)) : new Uint8Array(written))))")>]
let private packageBytes (written: JS.Promise<obj>): JS.Promise<byte[]> = jsNative

/// The bytes of what ImportXlsx was given (an ArrayBuffer, a typed array or
/// a Blob), as a Uint8Array of their own.
[<Emit("(typeof Blob !== 'undefined' && $0 instanceof Blob ? $0.arrayBuffer().then((b) => new Uint8Array(b)) : Promise.resolve($0 instanceof ArrayBuffer ? new Uint8Array($0.slice(0)) : new Uint8Array($0.buffer.slice($0.byteOffset, $0.byteOffset + $0.byteLength))))")>]
let private ownBytesOf (data: obj): JS.Promise<byte[]> = jsNative

[<Emit("Promise.resolve($0.slice())")>]
let private resolvedCopy (bytes: byte[]): JS.Promise<byte[]> = jsNative

[<Emit("new Uint8Array($0)")>]
let private bytesOfArrayBuffer (buffer: obj): byte[] = jsNative
[<Emit("$0.closest($1)")>]
let private gridClosest (element: Element) (selector: string): Element = jsNative


/// The document content a workbook holds outside its sheets (defined names
/// with their spelling and comments). A change of it is a change of the
/// document: the revision and OnChange follow it as they follow the sheets.
type private BookContent = {
    Names: Map<SheetId option * string, string>
    NameInfo: Map<SheetId option * string, DefinedNameInfo>
    /// The sheets without a tab, in tab order (the sheets themselves, their
    /// names and their order are compared sheet by sheet).
    Hidden: SheetId list
    /// The tables as structured references read them (name, place, columns).
    Tables: Map<string, Rbt.Spreadsheet.Workbook.Table>
    /// Iterative calculation of circular references.
    Iteration: IterativeCalculation option
}

let private bookContentOf (workbook: Workbook option) : BookContent =
    match workbook with
    | Some wb ->
        { Names = wb.Names
          NameInfo = wb.NameInfo
          Hidden = wb.SheetOrder |> List.filter (fun id -> match Map.tryFind id wb.Sheets with Some sheet -> not sheet.Visible | None -> false)
          Tables = wb.Tables
          Iteration = wb.Iteration }
    | None -> { Names = Map.empty; NameInfo = Map.empty; Hidden = []; Tables = Map.empty; Iteration = None }

let private sameBookContent (a: BookContent) (b: BookContent) =
    (sameRef a.Names b.Names || a.Names = b.Names)
    && (sameRef a.NameInfo b.NameInfo || a.NameInfo = b.NameInfo)
    && a.Hidden = b.Hidden
    && (sameRef a.Tables b.Tables || a.Tables = b.Tables)
    && a.Iteration = b.Iteration

/// The search a host started with Find: FindNext / FindPrevious /
/// ReplaceMatch go by it.
type private HostSearch = {
    Finder: SheetOps.Finder
    Request: Host.FindRequest
}

/// Compare document content through the same identity policy as the shared
/// session; OnChange must not miss a field that dirties the workbook.
let private sameSheetContent = Rbt.Spreadsheet.Session.sameSheetContent

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

[<Emit("$0.parentNode")>]
let private parentOf (el: Element): Element = jsNative

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
    /// The cell element says what its comment is (`data-comment`).
    mutable HasNote: bool
    mutable HasLink: bool
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
    EditHelp: Element
    Notice: Element
    InputMessage: Element
    /// The box of the comment the pointer is over (a note, or a thread).
    CommentBox: Element
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
type Spreadsheet(container: Element) as this =
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
    /// The picture of the active sheet that is selected (its id), and what
    /// the selection listeners were last told of it.
    let mutable selectedPicture: int option = None
    let mutable notifiedPicture: int option = None
    let mutable pictureDrag: PictureDrag option = None
    let mutable pictureFrame: Element option = None
    let mutable renderedDrawingsReadOnly = false
    /// The frozen panes and header sizes the drawings were last laid out for.
    let mutable renderedDrawingsPanes = ""
    let mutable drawingOriginX = 0.0
    let mutable drawingOriginY = 0.0
    let mutable drawingShiftX = 0.0
    let mutable drawingShiftY = 0.0
    // The portable session owns the workbook, calculation state and revision
    // identities. Sheet models below are projections for the browser grid.
    let initialBook = createWorkbook()
    let session =
        Rbt.Spreadsheet.Session.Session(
            { initialBook with
                Sheets =
                    Map.add initialBook.ActiveSheet
                        { initialBook.Sheets.[initialBook.ActiveSheet] with Model = model } initialBook.Sheets },
            historyBySheet = true)
    let mutable revisionSheets: SpreadsheetModel array = Array.copy sheets
    /// The file ImportXlsx opened; unchanged exports reuse its original bytes.
    let mutable importedPackage: (byte[] * int * int) option = None
    /// Undo what Mount set up outside the grid's own nodes (document-level
    /// listeners, observers, overlays appended to <body>).
    let cleanups = ResizeArray<unit -> unit>()
#if FABLE_COMPILER
    let mutable managerDialog: Rbt.FableUI.OfficeUiDialog.DialogState option = None
#endif
    let mutable disposed = false
    /// Every DOM listener the component registered: (target, event, handler).
    /// Dispose removes them all, also those on nodes that are no longer in
    /// the page, so that nothing the component made holds on to it.
    let listeners = ResizeArray<obj * string * obj>()
    /// OnError subscribers, and whether one of them is being told (an error
    /// in a subscriber is not reported to the subscribers again).
    let errorListeners = ResizeArray<obj -> unit>()
    let mutable reportingError = false
    /// Tell the OnError subscribers of an error the component caught:
    /// `source` names where (a subscription such as "OnChange", a DOM event
    /// such as "event:mousedown", "ImportXlsx" ...).
    let reportError (source: string) (error: obj) =
        if errorListeners.Count > 0 && not reportingError then
            reportingError <- true
            let payload = createObj [ "source" ==> source; "message" ==> errorMessageOf error; "error" ==> error ]
            for listener in errorListeners.ToArray() do
                if errorListeners.Contains listener then (try listener payload with _ -> ())
            reportingError <- false
    /// Register a listener (the same function object is kept for removal).
    /// What the handler throws is reported to OnError and thrown on.
    let listen (target: Element) (eventName: string) (handler: obj -> unit) =
        let boxed = guardedHandler (box handler) (fun error -> reportError ("event:" + eventName) error)
        addListenerRaw target eventName boxed
        listeners.Add((box target, eventName, boxed))
    let unlisten (target: Element) (eventName: string) (handler: obj -> unit) =
        // The handler registered under this name on this target (a function
        // value handed over twice need not be the same object).
        for index = listeners.Count - 1 downto 0 do
            let known, name, boxed = listeners.[index]
            if obj.ReferenceEquals(known, target) && name = eventName
               && (wrapsHandler boxed (box handler) || listeners |> Seq.filter (fun (t, n, _) -> obj.ReferenceEquals(t, target) && n = eventName) |> Seq.length = 1) then
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
    // --- context menus ------------------------------------------------------------
    // The component has no menu of its own. With an OnContextMenu subscriber
    // a right click (or the menu key / Shift+F10) selects what it is on, as
    // Excel does, and tells the subscriber where; the browser's menu is held
    // back. Without one, nothing changes: the browser shows its menu.
    let contextListeners = ResizeArray<obj -> unit>()
    /// When the keyboard last asked for the menu (the browser may follow it
    /// with a `contextmenu` event of its own, which is then not a second request).
    let mutable keyboardMenuAt = 0.0
    /// Tell the OnContextMenu subscribers. False when there is none.
    let raiseContextMenu (target: string) (row: int) (col: int) (picture: int option) (inSelection: bool) (clientX: float) (clientY: float) (sheetIndex: int) : bool =
        if contextListeners.Count = 0 then false
        else
            let payload =
                createObj [ "target" ==> target; "row" ==> row; "col" ==> col; "sheetIndex" ==> sheetIndex
                            "clientX" ==> clientX; "clientY" ==> clientY
                            "pictureId" ==> (match picture with Some id -> box id | None -> null)
                            "inSelection" ==> inSelection ]
            for listener in contextListeners.ToArray() do
                if contextListeners.Contains listener then (try listener payload with error -> reportError "OnContextMenu" (box error))
            true
    // --- AutoFilter buttons -------------------------------------------------------
    // The grid draws a button in every header cell of the sheet's AutoFilter
    // and of a table's header row. With an OnFilterButton subscriber a press
    // of one (or Alt+Down on its cell) tells the subscriber where the button
    // is; without one the grid opens its own checklist.
    let filterButtonListeners = ResizeArray<obj -> unit>()
    /// The button nodes on display (a pool: see `renderOverlays`).
    let filterButtonNodes = ResizeArray<Element>()
    /// The outline of the active cell's dynamic array, one node per pane (made when first needed).
    let spillOutlines : Element array = Array.zeroCreate 4
    let mutable workbook: Workbook option = None  // Optional workbook for table references
    /// The XLSX file the current workbook was imported from, for ExportXlsx:
    /// what the editor does not model is carried over from it.
    let mutable xlsxSource: Rbt.Spreadsheet.XlsxBridge.XlsxSource option = None
    // The VBA macros of the opened workbook (an .xlsm's project, kept as it was), and whether the host
    // let them run: off until EnableMacros, as Excel's "Enable content". The project is loaded (with the
    // VBA engine) on first use; until then only whether the workbook has one is known.
    let mutable macroPreserved: XlsxReader.WorkbookPreservation option = None
    let mutable macroProject: Rbt.Spreadsheet.MacroApi.IMacroProject option = None
    /// The macro running now (RunMacro runs a slice at a time, showing what it does): the grid is read-only meanwhile.
    let mutable macroRunning: Rbt.Spreadsheet.MacroApi.IMacroRun option = None
    let mutable macroProjectLoad: JS.Promise<Rbt.Spreadsheet.MacroApi.IMacroProject> option = None
    // The code names of the sheets (their modules, for Worksheet_Change ...) and of the workbook.
    let mutable sheetCodeNames: Map<SheetId, string> = Map.empty
    let mutable workbookCodeName = "ThisWorkbook"
    let mutable macrosEnabled = false
    let mutable workbookOpened = false
    let mutable macroHooks = browserMacroHooks
    // Whether a run waits at an Application.Wait of so many milliseconds (SetMacroHooks' wait; None: it does).
    let mutable macroWait : (float -> bool) option = None
    // What macros asked for later: procedures to run at a time (Application.OnTime: the timers, by
    // procedure) and keys bound to procedures (Application.OnKey: by OnKey name, see parseOnKey).
    let macroTimers = System.Collections.Generic.Dictionary<string, ResizeArray<int>>()
    let macroKeys = System.Collections.Generic.Dictionary<string, string>()
    /// Run a procedure once no macro is running (Excel runs OnTime and OnKey procedures when it is idle).
    /// The macro assigned to a drawn shape (its `macro` attribute, Excel's OnAction), without the workbook
    /// a name like 'Book1.xlsm'!Module.Proc starts with; None when it has none.
    let assignedMacro (drawing: SheetDrawing) : string option =
        match drawing.Content with
        | KeptObject (_, xml) ->
            let found = System.Text.RegularExpressions.Regex.Match(xml, "<(?:\\w+:)?(?:sp|grpSp|pic|cxnSp|graphicFrame)\\b[^>]*?\\smacro=\"([^\"]+)\"")
            if found.Success then
                let name = found.Groups.[1].Value
                let bang = name.LastIndexOf '!'
                Some (if bang >= 0 then name.Substring(bang + 1) else name)
            else None
        | _ -> None
    let rec runMacroWhenIdle (procedure: string) =
        if macrosEnabled then
            if macroRunning.IsSome then setTimeoutMs (fun () -> runMacroWhenIdle procedure) 20 |> ignore
            else (try this.RunMacro(procedure) |> ignore with _ -> ())
    let scheduleMacro (delayMs: float) (procedure: string) (schedule: bool) =
        let key = procedure.ToLowerInvariant()
        if schedule then
            let timers =
                match Portable.Dict.tryFind key macroTimers with
                | ValueSome found -> found
                | _ -> let created = ResizeArray() in macroTimers.[key] <- created; created
            let id = ref 0
            id.Value <- setTimeoutMs (fun () -> timers.Remove id.Value |> ignore; runMacroWhenIdle procedure) (int (min 2147483647.0 delayMs))
            timers.Add id.Value
        else
            match Portable.Dict.tryFind key macroTimers with
            | ValueSome timers -> (for id in timers do clearTimeoutId id); timers.Clear()
            | _ -> ()
    let bindMacroKey (key: string) (procedure: string option) =
        let name = parseOnKey key
        match procedure with
        | Some p -> macroKeys.[name] <- p
        | None -> macroKeys.Remove name |> ignore
    /// A workbook's macros end: its timers and keys go with them.
    let clearMacroCallbacks () =
        for KeyValue (_, timers) in macroTimers do (for id in timers do clearTimeoutId id)
        macroTimers.Clear()
        macroKeys.Clear()
    // What macros printed and how runs ended (the VBA editor's output), and whether a module was edited
    // here (then a save writes the project from the modules instead of the one opened).
    let macroLog = ResizeArray<string>()
    let mutable macroProjectEdited = false
    // The browser keeps a generation to cancel stale scheduled warmups.
    // Calculation graphs and statistics belong only to the shared session.
    let mutable engineGeneration = 0
    let mutable readOnly = false
    /// What the host asked for (SetReadOnly). `readOnly` is that, or a sheet
    /// on display that has no cells to edit (a chart sheet, a dialog sheet).
    let mutable viewerMode = false
    /// Whether the sheet on display takes no edits of its cells (see `sheetKindAt`).
    let mutable sheetLocked = false
    /// Set by a change that is how a viewer looks at the sheet (a filter
    /// applied in read-only mode): the next look at the revision takes the
    /// sheets as they are for the state already known - no new revision, no
    /// change event.
    let mutable absorbViewState = false
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
    // --- the outline gutter ----------------------------------------------------------
    // A sheet with grouped rows has a strip left of its row numbers, one with
    // grouped columns a strip above its column letters: a bracket and a +/-
    // button per group and the level buttons (1, 2, 3 ...). The strips are part
    // of the header area (`rowHeaderWidth` / `headerHeight` include them).
    let mutable showOutline = true
    /// The width of one level of the strip (a 13 px button and its gap).
    let outlineStep = 15.0
    /// The width of the rows' strip and the height of the columns' one, as drawn (0: none).
    let mutable gutterWidth = 0.0
    let mutable gutterHeight = 0.0
    /// The band of the column letters (30) and of the row numbers (50). A
    /// sheet that hides its headers has none - but for a narrow one beside an
    /// outline strip, where the level buttons of the other axis are.
    let mutable headBand = 30.0
    let mutable sideBand = 50.0
    /// The groups of the sheet on display, per `SheetAxes` value.
    let mutable outlineAxes: obj = null
    let mutable outlineRowGroups: (int * int * int * bool) array = [||]
    let mutable outlineColGroups: (int * int * int * bool) array = [||]
    let mutable outlineRowDepth = 0
    let mutable outlineColDepth = 0
    let outlineNodes = ResizeArray<Element>()
    let mutable outlineSignature: obj array = [||]
    let mutable cornerElement: Element option = None
    /// The zoom the grid is drawn at (the zoom of the sheet on display, see
    /// syncView). The scrolled content - cells, headers, overlays, pictures,
    /// the edit box - is laid out at 100 % in sheet pixels and scaled as a
    /// whole by the browser (CSS `zoom`, which scales layout and text, not a
    /// bitmap). Everything in the component is in sheet pixels; what is read
    /// from or written to the page (scroll offsets, the size of the grid,
    /// pointer positions) is converted here.
    let mutable appliedZoom = 1.0
    // Keep both the unzoomed CSS extent and the zoomed physical extent under
    // eight million pixels, below Chromium and Firefox layout-coordinate limits.
    let physicalExtentLimit = 8000000.0
    let maxLogicalExtent () = physicalExtentLimit / max 1.0 appliedZoom
    let mutable extentWidth = -1.0
    let mutable extentHeight = -1.0
    let mutable originX = 0.0
    let mutable originY = 0.0
    // A reveal can fall between physical pixels; keep its exact logical
    // position until the browser scrollbar is moved independently.
    let mutable anchorPhysicalX = -1.0
    let mutable anchorPhysicalY = -1.0
    let mutable anchorLogicalX = 0.0
    let mutable anchorLogicalY = 0.0
    let logicalScroll (physical: float) (extent: float) (client: float) =
        let physicalMax = max 0.0 (min extent (maxLogicalExtent ()) * appliedZoom - client)
        let logicalMax = max 0.0 (extent - client / appliedZoom)
        if physicalMax <= 0.0 then 0.0 else min logicalMax (max 0.0 (physical * logicalMax / physicalMax))
    let physicalScroll (logical: float) (extent: float) (client: float) =
        let physicalMax = max 0.0 (min extent (maxLogicalExtent ()) * appliedZoom - client)
        let logicalMax = max 0.0 (extent - client / appliedZoom)
        if logicalMax <= 0.0 then 0.0 else min physicalMax (max 0.0 (logical * physicalMax / logicalMax))
    let gridScrollTop (grid: Element) =
        let physical = getScrollTop grid
        if extentHeight > maxLogicalExtent () && physical = anchorPhysicalY then anchorLogicalY
        else
            anchorPhysicalY <- -1.0
            logicalScroll physical extentHeight (getClientHeight grid)
    let gridScrollLeft (grid: Element) =
        let physical = getScrollLeft grid
        if extentWidth > maxLogicalExtent () && physical = anchorPhysicalX then anchorLogicalX
        else
            anchorPhysicalX <- -1.0
            logicalScroll physical extentWidth (getClientWidth grid)
    let gridClientWidth (grid: Element) = getClientWidth grid / appliedZoom
    let gridClientHeight (grid: Element) = getClientHeight grid / appliedZoom
    let setGridScrollTop (grid: Element) (value: float) =
        setScrollTop grid (physicalScroll value extentHeight (getClientHeight grid))
        anchorPhysicalY <- getScrollTop grid
        anchorLogicalY <- max 0.0 (min (max 0.0 (extentHeight - gridClientHeight grid)) value)
    let setGridScrollLeft (grid: Element) (value: float) =
        setScrollLeft grid (physicalScroll value extentWidth (getClientWidth grid))
        anchorPhysicalX <- getScrollLeft grid
        anchorLogicalX <- max 0.0 (min (max 0.0 (extentWidth - gridClientWidth grid)) value)
    let paneX pane x = if pane >= 2 then x else x - originX
    let paneY pane y = if pane = 1 || pane = 3 then y else y - originY
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
    let mutable lookedUpStyles : obj = null
    let mutable lookedUpRules : obj = null
    let cellAt (row: int) (col: int) : CellData option =
        if row < 0 || col < 0 then None
        else
            if not (obj.ReferenceEquals(lookedUpIn, model.Cells)) || not (obj.ReferenceEquals(lookedUpStyles, model.ColumnStyles))
               || not (obj.ReferenceEquals(lookedUpRules, model.ConditionalFormats)) || lookedUpCells.Count > 40000 then
                lookedUpCells.Clear()
                lookedUpIn <- box model.Cells
                lookedUpStyles <- box model.ColumnStyles
                lookedUpRules <- box model.ConditionalFormats
            let key = float row * 16384.0 + float col
            match Portable.Dict.tryFind key lookedUpCells with
            | ValueSome known -> known
            | ValueNone ->
                // (A cell without an entry is drawn in its column's default style.)
                let found = SheetOps.shownCell model { Row = row; Col = col }
                lookedUpCells.[key] <- found
                found
    let mutable renderedStyleInputs: obj array = [||]
    let mutable renderedDrawings: obj = null
    let mutable renderedMetrics: obj = null
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
    let readOnlyListeners = ResizeArray<bool -> unit>()
    /// The edit state the listeners were last told (cell and text).
    let mutable notifiedEdit : (CellAddress * string) option = None
    let mutable clipboardGeneration = 0
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

    /// What kind of sheet the file had the sheet at `index` as: "worksheet"
    /// (also every sheet that is not a file's), "chartsheet" or "dialogsheet".
    /// The last two have no cells: the editor shows them as an empty sheet
    /// (a chart sheet with its chart on it) and writes them back as the parts
    /// they were read from.
    let sheetKindAt (index: int) : string =
        if index < 0 || index >= sheets.Length then "worksheet"
        else
            match (if index = activeSheetIndex then model else sheets.[index]).SourceSheet, xlsxSource with
            | Some origin, Some source when origin.Book = source.Id ->
                match source.Workbook.Sheets |> List.tryItem origin.Sheet |> Option.bind (fun sheet -> sheet.Extras.SourcePart) with
                | Some (_, relType) when relType.EndsWith "/chartsheet" -> "chartsheet"
                | Some (_, relType) when relType.EndsWith "/dialogsheet" -> "dialogsheet"
                | _ -> "worksheet"
            | _ -> "worksheet"
    let refreshSheetLock () =
        sheetLocked <- sheetKindAt activeSheetIndex <> "worksheet"
        readOnly <- viewerMode || sheetLocked || macroRunning.IsSome

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
                // The array has already been reordered by installBook; the
                // committed journal can still contain the previous order.
                match workbook |> Option.bind (fun current -> current.SheetOrder |> List.tryItem index) with
                | Some id -> session.ResolveDeferredSheet(id, loaded)
                | None -> ()
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


    let convertAllSheets () =
        formulaSheetsConverted <- true
        for index in 0 .. sheets.Length - 1 do
            if index <> activeSheetIndex then convertInPlace index

    // --- charts and shapes of the drawing --------------------------------------------
    // A chart, a shape or a group of an opened file is drawn from the parts of the file
    // (`SpreadsheetCharts`): the chart part its anchor names, with the values of the cells its
    // series read as they are now. The parts themselves are never changed.
    /// The parts of the opened file by path and its theme (made when a kept object is first drawn).
    let mutable drawingSource : (int * System.Collections.Generic.Dictionary<string, byte[]> * SpreadsheetCharts.ThemeColors) option = None
    /// The slicers and timelines of the opened file, drawn (`Slicers.drawer`: its parts read when one is first drawn).
    let mutable slicerDrawer : (int * (string -> string -> SpreadsheetCharts.ChartData -> float -> float -> SpreadsheetCharts.Primitive list option)) option = None
    /// The chart parts that were read, by path (None: not a chart that is drawn).
    let chartModels = System.Collections.Generic.Dictionary<string, SpreadsheetCharts.Chart option>()
    /// The frames on display that hold a drawn object: the node, its drawing, its size, and the blocks of
    /// cells its charts read - sheet (in capitals), first row / column, last row / column, the cells as they
    /// were when it was drawn (null: the sheet was not there, or not converted yet).
    let chartFrames = ResizeArray<Element * SheetDrawing * float * float * (string * int * int * int * int * CellData option [] []) array>()
    /// "normal", or "pageBreak": the page break preview - the pages the sheet prints on are outlined over its cells.
    let mutable viewMode = "normal"
    /// The outlines of the pages on display in the page break preview.
    let pageNodes = ResizeArray<Element>()
    /// The pages of the sheet on display as last worked out, and what they were worked out from.
    let mutable pagesKnown : (obj array * float * float * Rbt.Spreadsheet.Pdf.Pagination.Pagination) option = None
    /// How many times a drawn object was laid out again because cells it reads changed (`GetChartRedraws`).
    let mutable chartRedraws = 0
    /// How many times an object was laid out at all (`GetChartLayouts`).
    let mutable chartLayouts = 0
    /// What was drawn for an object, by the sheet of the file it is on and its id: its anchor, its size,
    /// the SVG and the blocks of cells read. A sheet that is shown again has its objects as they were
    /// drawn, while the cells they read are as they were.
    let drawnObjects = System.Collections.Generic.Dictionary<string, string * float * float * string * (string * int * int * int * int * CellData option [] []) array>()
    /// The cells of every sheet as the charts on display were last drawn from them.
    let mutable chartStamp : obj array = [||]
    /// Chart bindings change independently of cell values. Keep the imported parts untouched:
    /// a new binding state parses them again, then lays out the visible frames again.
    let mutable chartBindingState : (int * obj * (string * string * SheetOrigin option * int * obj) array) option = None
    let mutable chartBindingsDifferFromSource = false
    let mutable chartXmlRewrite : (string -> string) option = None
    let mutable chartBindingVersion = 0
    let mutable renderedChartBindingVersion = -1
    let mutable authoredBindingState : Map<string, string> array = [||]
    let ensureChartBindings () =
        let mutable authoredChanged = sheets.Length <> authoredBindingState.Length
        let mutable index = 0
        while not authoredChanged && index < sheets.Length do
            let current = if index = activeSheetIndex then model.AuthoredCharts else sheets.[index].AuthoredCharts
            authoredChanged <- not (obj.ReferenceEquals(current, authoredBindingState.[index]))
            index <- index + 1
        let same =
            match xlsxSource, workbook, chartBindingState with
            | None, _, None | _, None, None -> true
            | Some source, Some wb, Some (id, names, previous) ->
                let mutable matches = source.Id = id && obj.ReferenceEquals(box wb.Names, names) && sheets.Length = previous.Length
                let mutable index = 0
                for sheetId in wb.SheetOrder do
                    if matches then
                        if index >= previous.Length then matches <- false
                        else
                            let sheet = if index = activeSheetIndex then model else sheets.[index]
                            let oldId, oldName, oldOrigin, oldVersion, oldLog = previous.[index]
                            matches <-
                                sheetId = oldId && sheetNames.[index] = oldName && sheet.SourceSheet = oldOrigin
                                && sheet.StructureVersion = oldVersion && obj.ReferenceEquals(box sheet.StructureLog, oldLog)
                    index <- index + 1
                matches && index = previous.Length
            | _ -> false
        if not same || authoredChanged then
            chartBindingState <-
                match xlsxSource, workbook with
                | Some source, Some wb ->
                    let identities =
                        wb.SheetOrder
                        |> List.mapi (fun index sheetId ->
                            let sheet = if index = activeSheetIndex then model else sheets.[index]
                            sheetId, sheetNames.[index], sheet.SourceSheet, sheet.StructureVersion, box sheet.StructureLog)
                        |> List.toArray
                    Some (source.Id, box wb.Names, identities)
                | _ -> None
            authoredBindingState <- Array.init sheets.Length (fun index -> if index = activeSheetIndex then model.AuthoredCharts else sheets.[index].AuthoredCharts)
            chartModels.Clear()
            drawnObjects.Clear()
            chartXmlRewrite <- None
            chartBindingVersion <- chartBindingVersion + 1
            chartBindingsDifferFromSource <-
                match xlsxSource, workbook with
                | Some source, Some wb ->
                    let mutable changed = wb.Names <> source.BaselineNames || sheets.Length <> source.Baseline.Length
                    let mutable index = 0
                    while not changed && index < sheets.Length do
                        let sheet = if index = activeSheetIndex then model else sheets.[index]
                        changed <-
                            match sheet.SourceSheet with
                            | Some origin when origin.Book = source.Id && origin.Sheet >= 0 && origin.Sheet < source.Baseline.Length ->
                                sheet.StructureLog <> source.Baseline.[origin.Sheet].StructureLog
                                || sheetNames.[index] <> (List.item origin.Sheet source.Workbook.Sheets).Name
                                || index <> origin.Sheet
                            | _ -> true
                        index <- index + 1
                    changed
                | _ -> false
    let chartRewriter (source: Rbt.Spreadsheet.XlsxBridge.XlsxSource) =
        match chartXmlRewrite with
        | Some rewrite -> rewrite
        | None ->
            let rewrite =
                Rbt.Spreadsheet.XlsxBridge.Export.chartXmlForWorkbook source
                    (Array.init sheets.Length (fun index -> if index = activeSheetIndex then model else sheets.[index]))
                    sheetNames
            chartXmlRewrite <- Some rewrite
            rewrite
    /// The host's loader of faces (`SetFontLoader`), null without one; the faces asked of it
    /// ("family|bi"), and the families the system had when first drawn (never asked for).
    let mutable fontLoader : obj = null
    let wantedFaces = System.Collections.Generic.HashSet<string>()
    let systemFamilies = System.Collections.Generic.Dictionary<string, bool>()
    let mutable fontRedrawPending = false
    /// The faces of the PDF export (`SetPdfFonts`): (face name, family, bold, italic, TrueType bytes).
    let mutable pdfFonts : (string * string * bool * bool * byte[]) array = [||]
    /// The faces a host hands (`SetPdfFonts`, `ExportPdf`'s fonts).
    let pdfFontsArg (memberName: string) (fonts: obj) : (string * string * bool * bool * byte[]) array =
        if Host.isNullish fonts then [||]
        elif not (Host.isJsArray fonts) then Host.hostFail memberName "the fonts are an array of { faceName, family, bold, italic, data }"
        else
            (unbox<obj array> fonts) |> Array.mapi (fun at face ->
                if not (Host.isJsObject face) then Host.hostFail memberName (sprintf "font %d is not { faceName, family, bold, italic, data }" at)
                let field (upper: string) (lower: string) = if Host.isNullish (Host.getKey face upper) then Host.getKey face lower else Host.getKey face upper
                let family = field "Family" "family"
                if Host.jsTypeOf family <> "string" then Host.hostFail memberName (sprintf "font %d has no family (a face is { faceName, family, bold, italic, data })" at)
                let family = unbox<string> family
                let bold = Host.jsTypeOf (field "Bold" "bold") = "boolean" && unbox<bool> (field "Bold" "bold")
                let italic = Host.jsTypeOf (field "Italic" "italic") = "boolean" && unbox<bool> (field "Italic" "italic")
                let data = bytesOfHost (field "Data" "data")
                if isNull data then Host.hostFail memberName (sprintf "font %d (%s) has no data (its TrueType bytes: a Uint8Array or an ArrayBuffer)" at family)
                let named = field "FaceName" "faceName"
                let faceName =
                    if Host.jsTypeOf named = "string" && unbox<string> named <> "" then unbox<string> named
                    else family + (match bold, italic with | true, true -> "-BoldItalic" | true, false -> "-Bold" | false, true -> "-Italic" | _ -> "")
                faceName.Replace(" ", ""), family, bold, italic, data)
    let drawingPartsOf (sheet: SpreadsheetModel) : SpreadsheetCharts.DrawingParts option =
        ensureChartBindings ()
        let authored theme id =
            sheet.AuthoredCharts |> Map.tryFind id |> Option.bind (SpreadsheetCharts.parseChart theme)
        match xlsxSource, sheet.SourceSheet with
        | Some source, Some origin when origin.Book = source.Id && origin.Sheet >= 0 && origin.Sheet < source.Workbook.Sheets.Length ->
            let parts, theme =
                match drawingSource with
                | Some (id, parts, theme) when id = source.Id -> parts, theme
                | _ ->
                    let parts = System.Collections.Generic.Dictionary<string, byte[]>()
                    for path, data in source.Workbook.Preserved.Parts do parts.[path] <- data
                    let theme =
                        match source.Workbook.Preserved.ThemePath with
                        | Some path when parts.ContainsKey path -> SpreadsheetCharts.parseTheme (System.Text.Encoding.UTF8.GetString parts.[path])
                        | _ -> SpreadsheetCharts.defaultTheme
                    drawingSource <- Some (source.Id, parts, theme)
                    chartModels.Clear()
                    drawnObjects.Clear()
                    parts, theme
            let slicers =
                match slicerDrawer with
                | Some (id, drawer) when id = source.Id -> drawer
                | _ ->
                    let drawer = SpreadsheetSlicers.drawer source.Workbook
                    slicerDrawer <- Some (source.Id, drawer)
                    drawer
            let relationships = (List.item origin.Sheet source.Workbook.Sheets).Extras.DrawingRelationships
            let target (id: string) =
                relationships |> List.tryPick (fun (known, _, target, isExternal) -> if known = id && not isExternal then Some (ZipInterop.normalizePath target) else None)
            Some { Theme = theme
                   Chart =
                    (fun id ->
                        match authored theme id with
                        | Some chart -> Some chart
                        | None ->
                            target id
                            |> Option.bind (fun path ->
                                match Portable.Dict.tryFind path chartModels with
                                | ValueSome known -> known
                                | ValueNone ->
                                    let read =
                                        match Portable.Dict.tryFind path parts with
                                        | ValueSome data ->
                                            // Always parse the original part: structural and sheet edits are
                                            // applied once, from the imported baseline, without changing its bytes.
                                            let sheetIndex =
                                                if obj.ReferenceEquals(sheet, model) then activeSheetIndex
                                                else sheets |> Array.tryFindIndex (fun known -> obj.ReferenceEquals(known, sheet)) |> Option.defaultValue origin.Sheet
                                            let names (name: string) =
                                                workbook |> Option.bind (fun wb ->
                                                    Rbt.Spreadsheet.Workbook.getName (List.tryItem sheetIndex wb.SheetOrder) name wb)
                                            let originalXml = System.Text.Encoding.UTF8.GetString data
                                            let xml = chartRewriter source originalXml
                                            let parsed = SpreadsheetCharts.parseChartWith names theme xml
                                            if xml = originalXml
                                               && (not (SpreadsheetCharts.isChartEx xml)
                                                   || (workbook |> Option.forall (fun wb -> wb.Names = source.BaselineNames))) then parsed
                                            else
                                                let baselineNames (name: string) =
                                                    let key = name.ToUpperInvariant()
                                                    source.BaselineNames
                                                    |> Map.tryFind (Some (Rbt.Spreadsheet.XlsxBridge.sheetIdOfIndex origin.Sheet), key)
                                                    |> Option.orElseWith (fun () -> Map.tryFind (None, key) source.BaselineNames)
                                                match SpreadsheetCharts.parseChartWith baselineNames theme originalXml, parsed with
                                                | Some original, Some current -> Some (SpreadsheetCharts.markChangedReferences original current)
                                                | _ -> parsed
                                        | ValueNone -> None
                                    chartModels.[path] <- read
                                    read))
                   Picture =
                    (fun id ->
                        target id
                        |> Option.bind (fun path ->
                            match Portable.Dict.tryFind path parts with
                            | ValueSome data ->
                                let extension = (match path.LastIndexOf '.' with -1 -> "" | at -> path.Substring(at + 1).ToLowerInvariant())
                                let contentType =
                                    match extension with
                                    | "png" -> "image/png" | "jpg" | "jpeg" -> "image/jpeg" | "gif" -> "image/gif" | "svg" -> "image/svg+xml" | "bmp" -> "image/bmp" | "webp" -> "image/webp"
                                    | _ -> ""
                                if contentType = "" then None else Some (pictureUrl data contentType)
                            | ValueNone -> None))
                   Slicer = slicers }
        | _ ->
            // (A sheet of no file: the charts and shapes a macro or the user made, in the default theme.)
            Some { Theme = SpreadsheetCharts.defaultTheme
                   Chart = authored SpreadsheetCharts.defaultTheme
                   Picture = (fun _ -> None)
                   Slicer = (fun _ _ _ _ _ -> None) }
    /// A block of cells of a sheet (by name) as a chart reads it; None when the sheet is not there.
    let chartBlock (name: string) (r0: int) (c0: int) (r1: int) (c1: int) : CellData option [] [] option =
        match sheetNames |> Array.tryFindIndex (fun known -> known.Equals(name, System.StringComparison.OrdinalIgnoreCase)) with
        | Some index ->
            let sheet =
                if index = activeSheetIndex then model
                else
                    // A changed binding may point into a different block of a deferred
                    // sheet: its imported chart cache describes the *old* reference.
                    if chartBindingsDifferFromSource then convertInPlace index
                    sheets.[index]
            if isDeferredSheet sheet || float (r1 - r0 + 1) * float (c1 - c0 + 1) > 200000.0 then None
            else Some (Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } sheet)))
        | None -> None
    /// What a chart reads of the workbook; `seen` gets the blocks it read, with their cells.
    let chartDataReading (seen: ResizeArray<string * int * int * int * int * CellData option [] []>) : SpreadsheetCharts.ChartData =
        { Cells =
            (fun name r0 c0 r1 c1 ->
                let block = chartBlock name r0 c0 r1 c1
                seen.Add ((name.ToUpperInvariant(), r0, c0, r1, c1, (match block with Some cells -> cells | None -> null)))
                block)
          Display = (fun cell -> SpreadsheetCellText.cellDisplayTextIn model.Date1904 None 1.0e6 cell.Format cell.Value)
          Format = (fun code value -> SpreadsheetCellText.formatNumberWithCodeIn model.Date1904 code value) }
    /// Whether a block a chart read holds the cells it held then.
    let chartBlockStands (name: string, r0: int, c0: int, r1: int, c1: int, cells: CellData option [] []) : bool =
        match chartBlock name r0 c0 r1 c1 with
        | None -> isNull cells
        | Some now ->
            not (isNull cells) && now.Length = cells.Length
            && Array.forall2 (fun (a: CellData option []) (b: CellData option []) ->
                    a.Length = b.Length
                    && Array.forall2 (fun (x: CellData option) (y: CellData option) ->
                            match x, y with
                            | None, None -> true
                            | Some left, Some right -> obj.ReferenceEquals(left, right) || left = right
                            | _ -> false) a b) now cells
    /// The range a linked picture (Excel's camera, `<a14:cameraTool cellRange="'Raw Data'!$B$4:$D$11">`)
    /// shows: its sheet's index and the range.
    let cameraRange (xml: string) : (int * CellRange) option =
        let at = xml.IndexOf "cameraTool"
        if at < 0 then None
        else
            let key = "cellRange=\""
            let start = xml.IndexOf(key, at)
            if start < 0 then None
            else
                let start = start + key.Length
                let finish = xml.IndexOf('"', start)
                if finish < 0 then None
                else
                    let reference = xml.Substring(start, finish - start).Replace("&apos;", "'").Replace("&quot;", "\"").Replace("&amp;", "&")
                    match reference.LastIndexOf '!' with
                    | -1 -> None
                    | bang ->
                        let sheet = reference.Substring(0, bang).Trim()
                        let sheet = if sheet.StartsWith "'" && sheet.EndsWith "'" && sheet.Length >= 2 then sheet.Substring(1, sheet.Length - 2).Replace("''", "'") else sheet
                        match sheetNames |> Array.tryFindIndex (fun known -> known.Equals(sheet, System.StringComparison.OrdinalIgnoreCase)), SheetOps.tryParseRange (reference.Substring(bang + 1)) with
                        | Some index, Some range when (range.To.Row - range.From.Row + 1) * (range.To.Col - range.From.Col + 1) <= 10000 -> Some (index, range)
                        | _ -> None
    /// A linked picture as SVG: the cells of its range as they are now - fills, lines, text in their
    /// fonts and colours, the grid lines of a sheet that shows them -, stretched to the picture, as
    /// Excel draws it (acre-apartment's "Property Return Metrics" box: a picture of 'Raw Data'!B4:D11
    /// on five of its sheets). None when its range cannot be read.
    let cameraSvg (index: int) (range: CellRange) (width: float) (height: float) (seen: ResizeArray<string * int * int * int * int * CellData option [] []>) : string option =
        if index <> activeSheetIndex && isDeferredSheet sheets.[index] then convertInPlace index
        let source = if index = activeSheetIndex then model else sheets.[index]
        let r0, c0, r1, c1 = range.From.Row, range.From.Col, range.To.Row, range.To.Col
        let left (c: int) = SpreadsheetLayout.columnLeft source c - SpreadsheetLayout.columnLeft source c0
        let top (r: int) = SpreadsheetLayout.rowTop source r - SpreadsheetLayout.rowTop source r0
        let totalWidth, totalHeight = left (c1 + 1), top (r1 + 1)
        if totalWidth <= 0.0 || totalHeight <= 0.0 then None
        else
            // (Read as a chart reads its cells: drawn again when they change.)
            seen.Add ((sheetNames.[index].ToUpperInvariant(), r0, c0, r1, c1,
                       Array.init (r1 - r0 + 1) (fun r -> Array.init (c1 - c0 + 1) (fun c -> getCell { Row = r0 + r; Col = c0 + c } source))))
            let number (value: float) = (System.Math.Round (value * 100.0) / 100.0).ToString(System.Globalization.CultureInfo.InvariantCulture)
            let fills = System.Text.StringBuilder()
            let texts = System.Text.StringBuilder()
            let lines = System.Text.StringBuilder()
            let line (x1: float) (y1: float) (x2: float) (y2: float) (edge: SpreadsheetLayout.SharedEdge) =
                match edge with
                | SpreadsheetLayout.EdgeBorder border ->
                    (match borderCssParts border.Style with
                     | Some (weight, kind) ->
                        let dash = match kind with "dotted" -> " stroke-dasharray=\"1 1\"" | "dashed" -> " stroke-dasharray=\"3 1\"" | _ -> ""
                        lines.Append(sprintf "<line x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\" stroke=\"%s\" stroke-width=\"%d\"%s/>" (number x1) (number y1) (number x2) (number y2) (htmlEscape (defaultArg border.Color "#000000")) weight dash) |> ignore
                     | None -> ())
                // (A grid line where the sheet shows one: between two cells without a fill.)
                | SpreadsheetLayout.EdgeGrid true ->
                    lines.Append(sprintf "<line x1=\"%s\" y1=\"%s\" x2=\"%s\" y2=\"%s\" stroke=\"#d4d4d4\" stroke-width=\"1\"/>" (number x1) (number y1) (number x2) (number y2)) |> ignore
                | _ -> ()
            for r in r0 .. r1 do
                for c in c0 .. c1 do
                    let merge = findMergedCell { Row = r; Col = c } source
                    let primary = match merge with Some m -> m.StartRow = r && m.StartCol = c | None -> true
                    let x, y = left c, top r
                    let w, h = left (c + 1) - x, top (r + 1) - y
                    if primary && w > 0.0 && h > 0.0 then
                        let spanW = match merge with Some m -> left (min c1 m.EndCol + 1) - x | None -> w
                        let spanH = match merge with Some m -> top (min r1 m.EndRow + 1) - y | None -> h
                        let look = SpreadsheetPaint.cellLook source r c (SheetOps.shownCell source { Row = r; Col = c }) (Some (fun text fontPt fontName bold italic -> cellTextWidth text fontPt fontName bold italic)) spanW
                        fills.Append(sprintf "<rect x=\"%s\" y=\"%s\" width=\"%s\" height=\"%s\" fill=\"%s\"/>" (number x) (number y) (number spanW) (number spanH) (htmlEscape (defaultArg look.Fill "#ffffff"))) |> ignore
                        if look.Text <> "" then
                            let fontPx = SheetOps.pointsToPx look.FontSize
                            let anchor, tx =
                                match look.Align with
                                | SpreadsheetLayout.AlignRight -> "end", x + spanW - 3.0 - look.IndentRight
                                | SpreadsheetLayout.AlignCenter -> "middle", x + spanW / 2.0
                                | SpreadsheetLayout.AlignLeft -> "start", x + 3.0 + look.IndentLeft
                            let baseline =
                                match look.VerticalAlign with
                                | SpreadsheetPaint.VAlignTop -> y + 2.0 + fontPx * 0.95
                                | SpreadsheetPaint.VAlignMiddle -> y + spanH / 2.0 + fontPx * 0.35
                                | SpreadsheetPaint.VAlignBottom -> y + spanH - 4.0
                            texts.Append(sprintf "<text x=\"%s\" y=\"%s\" text-anchor=\"%s\" font-family=\"%s\" font-size=\"%s\" fill=\"%s\"%s%s>%s</text>"
                                            (number tx) (number baseline) anchor (htmlEscape (cellFontFamily look.FontName)) (number fontPx) (htmlEscape (defaultArg look.TextColor "#000000"))
                                            (if look.Bold then " font-weight=\"bold\"" else "") (if look.Italic then " font-style=\"italic\"" else "")
                                            (htmlEscape look.Text)) |> ignore
                    // The lines of the cell's top and left edges, and of the range's right and bottom ones.
                    if w > 0.0 && h > 0.0 then
                        line x y (x + w) y (SpreadsheetLayout.edgeAbove source r c)
                        line x y x (y + h) (SpreadsheetLayout.edgeLeft source r c)
                        if c = c1 then line (x + w) y (x + w) (y + h) (SpreadsheetLayout.edgeRight source r c)
                        if r = r1 then line x (y + h) (x + w) (y + h) (SpreadsheetLayout.edgeBelow source r c)
            Some (sprintf "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"%s\" height=\"%s\" viewBox=\"0 0 %s %s\" preserveAspectRatio=\"none\" style=\"display:block\">%s%s%s</svg>"
                    (number width) (number height) (number totalWidth) (number totalHeight) (fills.ToString()) (lines.ToString()) (texts.ToString()))
    /// A kept object of the sheet on display as SVG for its rectangle, with the blocks of cells its
    /// charts read; None for what is not drawn (the frame then says what the object is).
    let objectSvg (xml: string) (width: float) (height: float) : (string * (string * int * int * int * int * CellData option [] []) array) option =
        match cameraRange xml with
        | Some (index, range) ->
            let seen = ResizeArray<string * int * int * int * int * CellData option [] []>()
            (try cameraSvg index range width height seen with error -> reportError "drawing" (box error); None)
            |> Option.map (fun svg -> svg, seen.ToArray())
        | None ->
        try
            drawingPartsOf model
            |> Option.bind (fun parts ->
                let seen = ResizeArray<string * int * int * int * int * CellData option [] []>()
                // A chartEx chart (a waterfall, a funnel) keeps no copy of its data in its part: a sheet it
                // reads that is not converted yet is converted, and the chart drawn again with its cells
                // (a classic chart draws its cache until then).
                let mutable uncached = false
                let reading =
                    { parts with
                        Chart = (fun id ->
                            let chart = parts.Chart id
                            (match chart with
                             | Some read when read.Plots |> List.exists (fun plot -> match plot.Kind with SpreadsheetCharts.WaterfallPlot _ | SpreadsheetCharts.FunnelPlot -> true | _ -> false) -> uncached <- true
                             | _ -> ())
                            chart) }
                chartLayouts <- chartLayouts + 1
                let first = SpreadsheetCharts.objectScene reading (chartDataReading seen) xml width height
                let pending =
                    if not uncached then [||]
                    else
                        seen.ToArray()
                        |> Array.choose (fun (name, _, _, _, _, cells) -> if isNull cells then Some name else None)
                        |> Array.distinct
                        |> Array.choose (fun name -> sheetNames |> Array.tryFindIndex (fun known -> known.ToUpperInvariant() = name))
                        |> Array.filter (fun index -> index <> activeSheetIndex && isDeferredSheet sheets.[index])
                if pending.Length = 0 then first |> Option.map (fun scene -> SpreadsheetCharts.sceneSvg scene, seen.ToArray())
                else
                    for index in pending do convertInPlace index
                    seen.Clear()
                    SpreadsheetCharts.objectScene parts (chartDataReading seen) xml width height
                    |> Option.map (fun scene -> SpreadsheetCharts.sceneSvg scene, seen.ToArray()))
        with error ->
            // (A part that cannot be drawn is a frame that says what it is; the error is the host's to see.)
            reportError "drawing" (box error)
            None
    /// `objectSvg` for a drawing of the sheet on display: what was drawn for it before, when its
    /// anchor, its size and the cells it reads are what they were then.
    /// (`layOut` false: only what is known - None when the object would have to be laid out.)
    let drawingSvgWhen (layOut: bool) (drawing: SheetDrawing) (xml: string) (width: float) (height: float) : (string * (string * int * int * int * int * CellData option [] []) array) option option =
        ensureChartBindings ()
        let key = string activeSheetIndex + ":" + string drawing.Id
        // (An object that is not drawn is kept as such - no SVG: a slicer is not looked at again either.)
        let kept =
            match Portable.Dict.tryFind key drawnObjects with
            | ValueSome (knownXml, knownWidth, knownHeight, svg, reads)
                when knownWidth = width && knownHeight = height && (obj.ReferenceEquals(knownXml, xml) || knownXml = xml) && Array.forall chartBlockStands reads ->
                Some (if isNull svg then None else Some (svg, reads))
            | _ -> None
        match kept with
        | Some known -> Some known
        // (An object of no file part is not drawn: there is nothing to lay out later - but a shape a macro
        // drew is all in its XML.)
        | None when Map.isEmpty model.AuthoredCharts && xlsxSource.IsNone && (match drawing.Content with KeptObject ("shape", _) -> false | _ -> true) -> Some None
        | None ->
            if not layOut then None
            else
                let drawn = objectSvg xml width height
                if key <> "" then
                    match drawn with
                    | Some (svg, reads) -> drawnObjects.[key] <- (xml, width, height, svg, reads)
                    | None -> drawnObjects.[key] <- (xml, width, height, null, [||])
                Some drawn
    let drawingSvg (drawing: SheetDrawing) (xml: string) (width: float) (height: float) = drawingSvgWhen true drawing xml width height |> Option.bind id
    /// The objects of the sheet on display that are not laid out yet (by id), and whether their layout is under way.
    let pendingObjects = System.Collections.Generic.HashSet<int>()
    let mutable objectsScheduled = false
    let cellsStamp () : obj array = Array.init sheets.Length (fun index -> box (if index = activeSheetIndex then model.Cells else sheets.[index].Cells))

    let currentViewport (): SpreadsheetLayout.Viewport = {
        ScrollX = scrollLeft
        ScrollY = scrollTop
        Width = containerWidth
        Height = containerHeight
        HeaderHeight = headerHeight
        RowHeaderWidth = rowHeaderWidth
        Measurer = Some (fun text fontPt fontName bold italic ->
            browserTextWidth text (SheetOps.pointsToPx fontPt * drawnFontScale fontName) (fontName |> Option.defaultValue "Arial") bold italic)
    }

    let restoreCellOptions (changes: seq<CellAddress * CellData option>) (m: SpreadsheetModel) =
        let mutable result = m
        for addr, value in changes do
            match value with
            | Some cellData -> result <- setCell addr cellData result
            | None -> result <- clearCell addr result
        result


    /// The workbook was replaced or changed in more than its cells: the
    /// dependency graph and the calculated state it stands on are dropped.
    let invalidateCalcEngine () =
        session.InvalidateCalculation()
        engineGeneration <- engineGeneration + 1

    /// What hosts gave SetFilter, by "sheet id|column", with the criteria it
    /// became: GetFilters reads a column's filter back as it was given.
    let filterGiven = System.Collections.Generic.Dictionary<string, FilterCriteria * obj>()

    let sheetIdForIndex (index: int) =
        match workbook with
        | Some wb when index >= 0 && index < List.length wb.SheetOrder ->
            List.item index wb.SheetOrder
        | _ -> sprintf "sheet%d" (index + 1)

    /// What the grid remembers the room of the sheet on display by.
    let sheetKeyForRoom () = sheetIdForIndex activeSheetIndex
    /// The cell map whose last used row `roomUsedRow` is.
    let mutable roomCells : obj = null
    let mutable roomUsedRow = -1
    /// The size a sheet had before the caret went past what it uses, by sheet id: the grid
    /// gives the rows and columns back once neither the caret, a cell nor the view is there.
    let roomBefore = System.Collections.Generic.Dictionary<string, int * int>()
    /// A sheet always has room beyond what is used and beyond the active cell: with fewer than
    /// 50 rows (10 columns) left there, the grid grows to 100 rows (26 columns) past it - up to
    /// 1 048 576 rows and 16 384 columns. (A sheet of a file has as many rows as the file uses.)
    /// As in Excel that is all there is to scroll through: scrolling to the end adds nothing,
    /// the caret (or the selection) going further does - Ctrl+Down in an empty column goes to
    /// row 1 048 576 -, and when it comes back the sheet is as long as it was (or as its cells,
    /// or the rows in view, need it). Rows have no cost of their own (see
    /// `SpreadsheetLayout.AxisOffsets`). Nothing is done in a frame that changes none of this.
    let giveRoom () =
        let caretRow, caretCol =
            match model.Selection with
            | SingleCell (row, col) -> row, col
            // (Whole rows / columns selected: that is not a caret at the end of the sheet.)
            | Range (r1, c1, r2, c2) ->
                (if min r1 r2 <= 0 && max r1 r2 >= model.RowCount - 1 then -1 else max r1 r2),
                (if min c1 c2 <= 0 && max c1 c2 >= model.ColCount - 1 then -1 else max c1 c2)
        if not (obj.ReferenceEquals(roomCells, model.Cells)) then
            roomCells <- box model.Cells
            roomUsedRow <- if Map.isEmpty model.Cells then -1 else (Map.maxKeyValue model.Cells |> fst).Row
        let lastRow = max caretRow roomUsedRow
        let rows = if model.RowCount - (lastRow + 1) < 50 then min 1048576 (lastRow + 1 + 100) else model.RowCount
        let cols = if model.ColCount - (caretCol + 1) < 10 then min 16384 (caretCol + 1 + 26) else model.ColCount
        // (Narrow columns: as many as fill the view, as Excel's sheet always does.)
        let cols =
            let mutable n = 0
            let mutable right = 0.0
            let reach = scrollLeft + containerWidth - rowHeaderWidth
            while n < 16384 && (n < cols || right < reach) do
                right <- right + (model.ColumnWidths |> Map.tryFind n |> Option.defaultValue model.DefaultColumnWidth)
                n <- n + 1
            max cols n
        if rows > model.RowCount || cols > model.ColCount then
            if (rows > model.RowCount && caretRow > roomUsedRow) || cols > model.ColCount then
                let id = sheetKeyForRoom ()
                if not (roomBefore.ContainsKey id) then roomBefore.[id] <- (model.RowCount, model.ColCount)
            model <- { model with RowCount = max model.RowCount rows; ColCount = max model.ColCount cols }
        elif roomBefore.Count > 0 then
            let id = sheetKeyForRoom ()
            match Portable.Dict.tryFind id roomBefore with
            | ValueSome (rowsBefore, colsBefore) ->
                // What is in view stays (the scroll position is not pulled back).
                let metrics = SpreadsheetLayout.metricsOf model
                let viewRow = SpreadsheetLayout.indexAtOffset metrics.RowPrefix metrics.RowCount (scrollTop + containerHeight - headerHeight)
                let viewCol = SpreadsheetLayout.indexAtOffset metrics.ColPrefix metrics.ColCount (scrollLeft + containerWidth - rowHeaderWidth)
                let usedCol = if model.ColCount > colsBefore then (match usedBounds model with Some (_, col) -> col | None -> -1) else -1
                let lastCol = max caretCol usedCol
                let needRows = max (if rowsBefore - (lastRow + 1) < 50 then lastRow + 1 + 100 else rowsBefore) (viewRow + 2) |> min 1048576
                let needCols = max (if colsBefore - (lastCol + 1) < 10 then lastCol + 1 + 26 else colsBefore) (viewCol + 2) |> min 16384
                if needRows < model.RowCount || needCols < model.ColCount then
                    model <- { model with RowCount = min model.RowCount needRows; ColCount = min model.ColCount needCols }
                if model.RowCount <= rowsBefore && model.ColCount <= colsBefore then roomBefore.Remove id |> ignore
            | _ -> ()
    /// ... and a cell that is written to, or gone to (GoTo), is a cell of the sheet. The size it
    /// had is remembered as for the caret: once nothing is out there any more (the entry undone,
    /// the caret back), the grid is as small again.
    let giveRoomFor (addr: CellAddress) =
        if addr.Row >= 0 && addr.Col >= 0 && addr.Row < 1048576 && addr.Col < 16384 && (addr.Row >= model.RowCount || addr.Col >= model.ColCount) then
            let id = sheetKeyForRoom ()
            if not (roomBefore.ContainsKey id) then roomBefore.[id] <- (model.RowCount, model.ColCount)
            model <- { model with RowCount = max model.RowCount (addr.Row + 1); ColCount = max model.ColCount (addr.Col + 1) }

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
              NameInfo = Map.empty
              Iteration = None
              External = Map.empty }

    let installWorkbookSnapshot (wb: Workbook) =
        for i in 0 .. sheets.Length - 1 do
            let sheetId = sheetIdForIndex i
            match Map.tryFind sheetId wb.Sheets with
            | Some sheet -> sheets.[i] <- sheet.Model
            | None -> ()
        model <- sheets.[activeSheetIndex]
        workbook <- Some wb

    /// Register an already-applied browser operation through the shared
    /// workbook journal; the browser retains only UI projections.
    let pushCommand (description: string) (updated: SpreadsheetModel) =
        let transaction: Rbt.Spreadsheet.Session.AppliedTransaction =
            { Description = description; Workbook = currentWorkbookSnapshot() }
        session.CommitApplied(transaction)
        updated
    let pushValueNeutral (description: string) =
        model <- pushCommand description model

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
        SheetOps.listSourceReaderIn
            // (A source that is a formula is evaluated in the workbook as it is now, every sheet converted.)
            (fun () ->
                if not formulaSheetsConverted then convertAllSheets ()
                Some (currentWorkbookSnapshot ()))
            (fun () -> Some (sheetIdForIndex activeSheetIndex))
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
            | SpreadsheetPaint.Clipped (_, children) -> children |> List.iter drawVec
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
            | SpreadsheetPaint.VecPath _ | SpreadsheetPaint.VecGroup _ | SpreadsheetPaint.Clipped _ -> drawVec painted.Item
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
                    | SpreadsheetLayout.AlignCenter -> "center"
                    | SpreadsheetLayout.AlignRight -> "right"
                    | SpreadsheetLayout.AlignLeft -> "left"
                let x =
                    match text.Align with
                    | SpreadsheetLayout.AlignCenter -> text.Rect.X + text.Rect.Width / 2.0
                    | SpreadsheetLayout.AlignRight -> text.Rect.X + text.Rect.Width
                    | SpreadsheetLayout.AlignLeft -> text.Rect.X
                // Approximate Calibri's em-relative metrics (asc 0.952, desc
                // -0.266, lineHeight 1.218). Browser canvas doesn't expose
                // font metrics portably, so we accept a small drift for
                // non-Calibri fonts — Excel ships Calibri as the default.
                // (The size is in points: 11 pt is drawn 14.67 px tall, as Excel does at 100 %.)
                let fontPx = SheetOps.pointsToPx text.FontSize * drawnFontScale text.FontName
                let ascent = fontPx * 0.952
                let lineHeight = fontPx * 1.218
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
                    let totalWidth = browserTextWidth text.Text fontPx fontName text.Bold text.Italic
                    let descent = fontPx * 0.266
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
                    canvasTextRotated ctx text.Text fontPx (ax0 + shiftX) (ay0 + shiftY) theta text.Color text.Bold text.Italic fontName
                else
                canvasText ctx text.Text x baseline fontPx text.Color text.Bold text.Italic align fontName
                if not isRotated && (text.Underline || text.UnderlineDouble || text.Strikethrough) then
                    let textWidth = browserTextWidth text.Text fontPx fontName text.Bold text.Italic
                    let startX =
                        match text.Align with
                        | SpreadsheetLayout.AlignCenter -> x - textWidth / 2.0
                        | SpreadsheetLayout.AlignRight -> x - textWidth
                        | SpreadsheetLayout.AlignLeft -> x
                    let endX = startX + textWidth
                    if text.Underline || text.UnderlineDouble then
                        let y = baseline + max 1.0 (fontPx * 0.12)
                        canvasLine ctx startX y endX y text.Color 1.0
                        if text.UnderlineDouble then
                            canvasLine ctx startX (y + 2.0) endX (y + 2.0) text.Color 1.0
                    if text.Strikethrough then
                        let y = baseline - fontPx * 0.30
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

    /// Browser projection of the shared session's incremental calculation.
    let ensureRecalcEngine () = session.EnsureRecalcEngine()

    let recalculateAfterEdit (hint: (SheetId * CellAddress) list option) =
        session.UpdateView(currentWorkbookSnapshot())
        let prepareFull () =
            // Import keeps formula sheets deferred until incremental evaluation
            // needs them, or until the engine declines and a full pass is required.
            convertFormulaSheets ()
            currentWorkbookSnapshot()
        session.Recalculate(?hint = hint, beforeFull = prepareFull)
        installWorkbookSnapshot session.Workbook
        adoptConvertedSheets ()

    let recalculateFormulas () = recalculateAfterEdit None
    let recalculateEverything () = recalculateAfterEdit None

    let recalculateFromScratch () =
        invalidateCalcEngine()
        convertFormulaSheets ()
        session.UpdateView(currentWorkbookSnapshot())
        session.RecalculateFromScratch()
        installWorkbookSnapshot session.Workbook
        adoptConvertedSheets ()


    /// Insert or delete rows/columns on the active sheet as one undoable
    /// step. The edit and its undo both go through the workbook so formulas on
    /// other sheets, defined names and structured-reference tables follow.
    let applyStructureEdit (edit: StructureEdit) (description: string) =
        // Formulas of every sheet follow the rows / columns that move.
        convertFormulaSheets ()
        let sheetId = sheetIdForIndex activeSheetIndex
        session.UpdateView(currentWorkbookSnapshot())
        let prepareFull () =
            installWorkbookSnapshot session.Workbook
            convertFormulaSheets ()
            currentWorkbookSnapshot()
        match session.ExecuteStructural(Rbt.Spreadsheet.Session.EditStructure (sheetId, edit), description = description, beforeFull = prepareFull) with
        | Ok true ->
            installWorkbookSnapshot session.Workbook
            adoptConvertedSheets ()
        | Ok false -> ()
        | Result.Error problem -> failwith problem

    let prepareClipboardFull () =
        installWorkbookSnapshot session.Workbook
        convertFormulaSheets ()
        currentWorkbookSnapshot()

    let pasteClipboardAtSelection () =
        match session.Clipboard with
        | None -> false
        | Some captured ->
            if captured.Data.IsCut then convertFormulaSheets ()
            session.UpdateView(currentWorkbookSnapshot())
            match session.PasteClipboard(sheetIdForIndex activeSheetIndex, beforeFull = prepareClipboardFull) with
            | Result.Error problem -> failwith problem
            | Ok _ ->
                installWorkbookSnapshot session.Workbook
                adoptConvertedSheets ()
                true

    /// Replace in the selected cell as one undoable step. True when it matched.
    let replaceInSelection (findText: string) (replaceText: string) (caseSensitive: bool) =
        match replaceAtSelection findText replaceText caseSensitive model with
        | Some newModel ->
            model <- newModel
            recalculateFormulas()
            model <- pushCommand "Replace" model
            true
        | None -> false

    /// Replace every match on the sheet as one undoable step; returns the count.
    let replaceEverywhere (findText: string) (replaceText: string) (caseSensitive: bool) =
        let newModel, count, _, _ = replaceAllWithChanges findText replaceText caseSensitive model
        if count > 0 then
            model <- newModel
            recalculateFormulas()
            model <- pushCommand "Replace All" model
        count

    let recalculateSingleCellEdit (addr: CellAddress) (cellData: CellData) =
        // Build the graph before replacing the cell, not against its new value.
        session.UpdateView(currentWorkbookSnapshot())
        ensureRecalcEngine ()
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
    /// Where the pointer of a drag (of the selection, of the fill handle) was last seen, and
    /// the timer that scrolls the sheet while it is held at an edge of the grid (see Mount).
    let mutable dragPointer: obj option = None
    let mutable dragScrollTimer: int option = None
    let startDragScrollHook : (unit -> unit) ref = ref ignore
    let startDragScroll () = startDragScrollHook.Value ()
    let stopDragScroll () =
        match dragScrollTimer with
        | Some timer ->
            JS.clearInterval timer
            dragScrollTimer <- None
        | None -> ()
        dragPointer <- None
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
    /// The cell whose comment box is shown.
    let mutable commentBoxCell: (CellAddress * SheetOps.CommentInfo) option = None
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
        if not (obj.ReferenceEquals(outlineAxes, model.Axes)) then
            outlineAxes <- box model.Axes
            outlineRowGroups <- SheetOps.outlineGroups true model |> Array.ofList
            outlineColGroups <- SheetOps.outlineGroups false model |> Array.ofList
            outlineRowDepth <- SheetOps.outlineDepth true model
            outlineColDepth <- SheetOps.outlineDepth false model
        // (A sheet draws its strips with or without its headers, unless it hides the outline's symbols.)
        let strip (depth: int) = if showOutline && model.Axes.ShowSymbols && depth > 0 then float (depth + 1) * outlineStep + 3.0 else 0.0
        let wantedWidth = strip outlineRowDepth
        let wantedHeight = strip outlineColDepth
        if model.ShowHeaders <> appliedHeaders || wantedWidth <> gutterWidth || wantedHeight <> gutterHeight then
            appliedHeaders <- model.ShowHeaders
            gutterWidth <- wantedWidth
            gutterHeight <- wantedHeight
            headBand <- if appliedHeaders then 30.0 elif gutterWidth > 0.0 then outlineStep + 3.0 else 0.0
            sideBand <- if appliedHeaders then 50.0 elif gutterHeight > 0.0 then outlineStep + 3.0 else 0.0
            headerHeight <- headBand + gutterHeight
            rowHeaderWidth <- sideBand + gutterWidth
            match cornerElement with
            | Some corner ->
                setStyle corner "width" (string rowHeaderWidth + "px")
                setStyle corner "height" (string headerHeight + "px")
            | None -> ()
            match layers with
            | Some layer ->
                // (Without headers the layers still hold the outline strips; the header cells are hidden by a rule of the sheet style.)
                let display = if appliedHeaders || gutterWidth > 0.0 || gutterHeight > 0.0 then "" else "none"
                for headers in [ layer.ColHeaders; layer.RowHeaders; layer.FrozenHeaders ] do
                    setStyle headers "display" display
                    if appliedHeaders then removeAttribute headers "data-headers" else setAttribute headers "data-headers" "off"
            | None -> ()
        let zoom = if model.Zoom >= 0.1 && model.Zoom <= 4.0 then model.Zoom else 1.0
        if zoom <> appliedZoom then
            match contentElement, gridContainer with
            | Some content, Some grid ->
                appliedZoom <- zoom
                // (A cell's edges are rounded to the screen's pixels at the zoom: every cell is placed again.)
                styleEpoch <- styleEpoch + 1
                setStyle content "zoom" (if zoom = 1.0 then "" else string zoom)
                containerHeight <- gridClientHeight grid
                containerWidth <- gridClientWidth grid
            | _ -> appliedZoom <- zoom

    /// Give the scrolled content the extent of the active sheet, so that
    /// every row and column can be scrolled to. Done whenever the sheet's
    /// size or the size of a row / column changes.
    let syncExtent () =
        let oldZoom = appliedZoom
        let oldTop, oldLeft =
            match gridContainer with
            | Some grid -> gridScrollTop grid, gridScrollLeft grid
            | None -> scrollTop, scrollLeft
        syncView ()
        match contentElement, gridContainer with
        | Some content, Some grid ->
            let metrics = SpreadsheetLayout.metricsOf model
            let width = rowHeaderWidth + metrics.ColPrefix.[metrics.ColCount]
            let height = headerHeight + metrics.RowPrefix.[metrics.RowCount]
            if width <> extentWidth || height <> extentHeight || oldZoom <> appliedZoom then
                extentWidth <- width
                extentHeight <- height
                setStyle content "width" (string (min width (maxLogicalExtent ())) + "px")
                setStyle content "height" (string (min height (maxLogicalExtent ())) + "px")
                // Out-of-view descendants must not enlarge the bounded scroll area.
                setStyle content "overflow" (if width > maxLogicalExtent () || height > maxLogicalExtent () then "clip" else "")
                setGridScrollTop grid oldTop
                setGridScrollLeft grid oldLeft
                scrollTop <- gridScrollTop grid
                scrollLeft <- gridScrollLeft grid
        | _ -> ()

    /// Scroll the grid so that a cell is in view (frozen rows and columns
    /// stay put and cover the start of the scrolled area).
    let revealCell (row: int) (col: int) =
        match gridContainer with
        | None -> ()
        | Some grid ->
            // (A cell far past the end - Ctrl+Down in an empty column - is a row of the sheet first.)
            if not (isDeferredSheet model) then giveRoom ()
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
        // (With nothing ahead: the last row / column a sheet can have - row 1 048 576, column
        // XFD -, as in Excel; a viewer's sheet ends where it ends.)
        let limit =
            if readOnly then (if deltaRow <> 0 then model.RowCount else model.ColCount) - 1
            elif deltaRow <> 0 then 1048575 else 16383
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
        // (Rows an AutoFilter hides are not copied.)
        [| for row in filterShownRows top bottom model ->
            [| for col in left .. right ->
                let addr = { Row = row; Col = col }
                addr, getCell addr model |] |]

    let cellText (data: CellData option) : string =
        match data with
        | Some data -> SpreadsheetCellText.cellDisplayTextIn model.Date1904 None 1.0e6 data.Format data.Value
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

    /// Commit text to a cell: parse it the way Excel reads input, check it
    /// against the cell's validation (`validate`), store it and recalculate,
    /// as one undoable step. `typed`: the entry is the user's, and what
    /// follows from typing in or beside a table follows (a host's SetCell is
    /// not). False when a validation rule refused the entry.
    /// Set once the workbook's table steps exist (`changeTables`): enters a cell as a
    /// change of the tables; false when the entry concerns none.
    let tableEntryHook : (CellAddress * CellData option -> bool) ref = ref (fun _ -> false)

    let commitInput (addr: CellAddress) (text: string) (validate: bool) (typed: bool) : bool =
        giveRoomFor addr
        // (While a macro runs the grid shows its frames: an entry would be lost when it ends.)
        if sheetLocked || macroRunning.IsSome then false else
        let isClear = System.String.IsNullOrWhiteSpace text
        let hadCell = getCell addr model |> Option.isSome
        let mutable tableHandled = false
        let asTableEntry (cell: CellData option) =
            if typed && SheetOps.entryConcernsTable addr model && tableEntryHook.Value (addr, cell) then
                tableHandled <- true
                true
            else false
        let notify style title message =
            let styleName =
                match style with
                | StopStyle -> "stop"
                | WarningStyle -> "warning"
                | InformationStyle -> "information"
            validationNotice <- Some (styleName, defaultArg title "", message, addr)
            for listener in validationListeners.ToArray() do
                try listener styleName (defaultArg title "") message with error -> reportError "OnValidationNotice" (box error)
        let accept (cell: CellData) =
            if isClear then not (asTableEntry None)
            else
                let outcome =
                    if validate && not model.Validations.IsEmpty then
                        refreshRuleWorkbook()
                        checkEntryIn listReader ruleEvaluator addr cell.Value model
                    else EntryValid
                match outcome with
                | EntryInvalid (StopStyle, title, message) ->
                    notify StopStyle title message
                    false
                | _ ->
                    match outcome with
                    | EntryInvalid (style, title, message) -> notify style title message
                    | EntryValid -> validationNotice <- None
                    not (asTableEntry (Some cell))
        if isClear && not hadCell then true else
        session.UpdateView(currentWorkbookSnapshot())
        let prepareFull () =
            installWorkbookSnapshot session.Workbook
            convertFormulaSheets ()
            currentWorkbookSnapshot()
        // The sheet's Worksheet_Change (its module's, by code name) runs in the edit's step: undone with it.
        let changeEvent : (Workbook -> Workbook) option =
            match macroProject with
            | Some project when macrosEnabled ->
                let sheetId = sheetIdForIndex activeSheetIndex
                let handler = (Map.tryFind sheetId sheetCodeNames |> Option.defaultValue sheetNames.[activeSheetIndex]) + ".Worksheet_Change"
                if not (project.Has handler) then None
                else
                    Some (fun wb ->
                        match project.RunEvent(wb, macroHooks, handler, Some (sheetId, addr.Row, addr.Col, addr.Row, addr.Col)) with
                        | Some (Rbt.Spreadsheet.MacroApi.MacroDone (next, _)) -> next
                        | Some (Rbt.Spreadsheet.MacroApi.MacroFailed failure) ->
                            macroHooks.Output failure.Text
                            wb
                        | None -> wb)
            | _ -> None
        let accepted =
            if isClear then
                session.ClearCell(sheetIdForIndex activeSheetIndex, addr, accept = accept, ?afterCalculate = changeEvent, beforeFull = prepareFull)
            else
                session.EnterCell(sheetIdForIndex activeSheetIndex, addr, text, currentYear(),
                                  accept = accept, ?afterCalculate = changeEvent, beforeFull = prepareFull)
        if accepted then
            installWorkbookSnapshot session.Workbook
            adoptConvertedSheets ()
        accepted || tableHandled

    /// Continue the block `source` in `direction` by `count` rows/columns as
    /// one undoable step. `series` continues series (the fill handle);
    /// without it the block is copied (Ctrl+D / Ctrl+R).
    let fillFrom (source: int * int * int * int) (direction: SpreadsheetInput.FillDirection) (count: int) (series: bool) =
        if count > 0 && not readOnly then
            // (A fill past the last row or column of the sheet makes the sheet that large.)
            (let _, _, bottom, right = source
             match direction with
             | SpreadsheetInput.FillDown -> giveRoomFor { Row = bottom + count; Col = right }
             | SpreadsheetInput.FillRight -> giveRoomFor { Row = bottom; Col = right + count }
             | _ -> ())
            model <- SpreadsheetInput.fillCells source direction count series model
            recalculateFormulas()
            let top, left, bottom, right = source
            let selection =
                match direction with
                | SpreadsheetInput.FillDown -> Range (top, left, bottom + count, right)
                | SpreadsheetInput.FillUp -> Range (max 0 (top - count), left, bottom, right)
                | SpreadsheetInput.FillRight -> Range (top, left, bottom, right + count)
                | SpreadsheetInput.FillLeft -> Range (top, max 0 (left - count), bottom, right)
            model <- { model with Selection = selection }
            model <- pushCommand "Fill" model

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
            model <- pushCommand "Delete" model

    /// Paste external table cells through the shared workbook transaction.
    let pasteExternalRows (rows: string array array) : bool =
        if readOnly || not (rows |> Array.exists (fun row -> row.Length > 0)) then false
        else
            let top, left, _, _ = selectionBounds()
            let originalRows, originalCols = model.RowCount, model.ColCount
            session.UpdateView(currentWorkbookSnapshot())
            match session.PasteRows(sheetIdForIndex activeSheetIndex, rows, currentYear(), beforeFull = prepareClipboardFull) with
            | Result.Error problem -> failwith problem
            | Ok _ ->
                installWorkbookSnapshot session.Workbook
                adoptConvertedSheets ()
                anchorCell <- { Row = top; Col = left }
                if model.RowCount > originalRows || model.ColCount > originalCols then
                    let id = sheetKeyForRoom ()
                    if not (roomBefore.ContainsKey id) then roomBefore.[id] <- (originalRows, originalCols)
                true

    let pasteFlavours (text: string) (html: string) : bool =
        if readOnly then false
        else
            let pasteExternal () =
                let fromHtml = if System.String.IsNullOrEmpty html then null else htmlTableCells html
                if not (isNull (box fromHtml)) && fromHtml.Length > 0 then pasteExternalRows fromHtml
                elif not (System.String.IsNullOrEmpty text) then
                    pasteExternalRows (Rbt.Spreadsheet.Session.clipboardRowsOfText text)
                else false
            match session.ClipboardSource(Some text) with
            | Rbt.Spreadsheet.Session.InternalPaste -> pasteClipboardAtSelection()
            | Rbt.Spreadsheet.Session.IgnoredPaste when System.String.IsNullOrEmpty text -> pasteExternal ()
            | Rbt.Spreadsheet.Session.IgnoredPaste -> false
            | Rbt.Spreadsheet.Session.ExternalPaste -> pasteExternal ()

    let executeFormatting description block command =
        let sheetId = sheetIdForIndex activeSheetIndex
        session.UpdateView(currentWorkbookSnapshot())
        let prepareFull () =
            installWorkbookSnapshot session.Workbook
            convertFormulaSheets ()
            currentWorkbookSnapshot()
        match session.ExecuteFormatting(sheetId, command, bounds = block, activeCell = activeCellAddress(),
                                        description = description, beforeFull = prepareFull) with
        | Ok true ->
            installWorkbookSnapshot session.Workbook
            adoptConvertedSheets ()
            true
        | Ok false -> false
        | Result.Error problem -> failwith problem

    let formatSelection description update =
        executeFormatting description (selectionBounds()) (Rbt.Spreadsheet.Session.TransformFormat update)

    /// An argument a host must give: throws when it is undefined or null.
    let requiredArg (memberName: string) (what: string) (value: obj) : obj =
        if Host.isNullish value then Host.hostFail memberName (what + " is missing") else value

    let clearBlock description what block =
        executeFormatting description block (Rbt.Spreadsheet.Session.ClearCells what)

    let executeRule memberName command =
        session.UpdateView(currentWorkbookSnapshot())
        match session.ExecuteRule(sheetIdForIndex activeSheetIndex, command) with
        | Result.Error problem -> Host.hostFail memberName problem
        | Ok result ->
            if result.Changed then
                installWorkbookSnapshot session.Workbook
                match command with
                | Rbt.Spreadsheet.Session.SetValidation _ | Rbt.Spreadsheet.Session.ClearValidation _ ->
                    invalidCircles.Remove activeSheetIndex |> ignore
                | _ -> ()
            result

    let addRule ranges stopIfTrue format =
        let result = executeRule "AddConditionalFormat" (Rbt.Spreadsheet.Session.AddConditionalFormat (ranges, stopIfTrue, format))
        result.Id.Value

    /// Change the comments of the active sheet (notes are on the cells,
    /// threads beside them) as one undoable step; no value changes. False
    /// when nothing changed.
    let changeComments (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let cellsBefore, threadsBefore = model.Cells, model.Threads
        let changed = change model
        let cellsAfter, threadsAfter = changed.Cells, changed.Threads
        if sameRef cellsBefore cellsAfter && (sameRef threadsBefore threadsAfter || threadsBefore = threadsAfter) then false
        else
            model <- { model with Cells = cellsAfter; Threads = threadsAfter }
            pushValueNeutral description
            true
    /// The author of a comment as a host gives it: a text, or nothing.
    let commentAuthorArg (memberName: string) (value: obj) : string option =
        Host.optionalArg (Host.stringArg memberName "the author") value |> Option.filter (fun name -> name.Trim() <> "")

    // --- hyperlinks ---------------------------------------------------------------
    let hyperlinkListeners = ResizeArray<obj -> unit>()
    /// Change the hyperlinks of the active sheet, with the text and the look
    /// of the linked cells, as one undoable step. False when nothing changed.
    let changeHyperlinks (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let cellsBefore, linksBefore = model.Cells, model.Hyperlinks
        let changed = change model
        let cellsAfter, linksAfter = changed.Cells, changed.Hyperlinks
        if sameRef cellsBefore cellsAfter && (sameRef linksBefore linksAfter || linksBefore = linksAfter) then false
        else
            model <- { model with Cells = cellsAfter; Hyperlinks = linksAfter } |> invalidateFilterProjection
            // (A link can give its first cell a text: what reads the cell follows.)
            let sameValues =
                sameRef cellsBefore cellsAfter
                || (cellsBefore.Count = cellsAfter.Count
                    && Seq.forall2 (fun (KeyValue(a, before: CellData)) (KeyValue(b, after: CellData)) -> a = b && (sameRef before after || before.Value = after.Value)) cellsBefore cellsAfter)
            if sameValues then pushValueNeutral description
            else
                recalculateFormulas()
                model <- pushCommand description model
            true

    // --- pictures -------------------------------------------------------------------
    let executeDrawing memberName description command =
        session.UpdateView(currentWorkbookSnapshot())
        match session.ExecuteDrawing(sheetIdForIndex activeSheetIndex, command, description = description) with
        | Result.Error problem -> Host.hostFail memberName problem
        | Ok result ->
            if result.Changed then installWorkbookSnapshot session.Workbook
            result
    /// The cell a point of the sheet (pixels at 100%) is in, and how far into it.
    let cellAtSheetPoint (x: float) (y: float) : CellAddress * float * float =
        let metrics = SpreadsheetLayout.metricsOf model
        let x, y = max 0.0 x, max 0.0 y
        let col = max 0 (min (metrics.ColCount - 1) (SpreadsheetLayout.indexAtOffset metrics.ColPrefix metrics.ColCount x))
        let row = max 0 (min (metrics.RowCount - 1) (SpreadsheetLayout.indexAtOffset metrics.RowPrefix metrics.RowCount y))
        { Row = row; Col = col }, x - metrics.ColPrefix.[col], y - metrics.RowPrefix.[row]
    /// Give a picture a rectangle of the sheet as one undoable step: it is
    /// anchored to the cell its top-left corner is in. Throws (as the member
    /// named) when there is no such picture.
    let placePictureAt (memberName: string) (description: string) (id: int) (rect: SpreadsheetLayout.Rect) : bool =
        let cell, offsetX, offsetY = cellAtSheetPoint rect.X rect.Y
        (executeDrawing memberName description (Rbt.Spreadsheet.Session.PlacePicture (id, cell, offsetX, offsetY, max 1.0 rect.Width, max 1.0 rect.Height))).Changed
    /// The rectangle a drag has taken a picture to: moved by the pointer's
    /// travel, or sized from the handle held - a corner keeps the picture's
    /// proportions unless `free` (Shift), a side changes one dimension.
    let draggedRect (drag: PictureDrag) (dx: float) (dy: float) (free: bool) : SpreadsheetLayout.Rect =
        let start = drag.Start
        if drag.Handle = "move" then { start with X = max 0.0 (start.X + dx); Y = max 0.0 (start.Y + dy) }
        else
            let west, east = drag.Handle.Contains "w", drag.Handle.Contains "e"
            let north, south = drag.Handle.Contains "n", drag.Handle.Contains "s"
            let least = 8.0
            let mutable width = if east then start.Width + dx elif west then start.Width - dx else start.Width
            let mutable height = if south then start.Height + dy elif north then start.Height - dy else start.Height
            width <- max least width
            height <- max least height
            if drag.Handle.Length = 2 && not free && start.Width > 0.0 && start.Height > 0.0 then
                let scale = max (width / start.Width) (height / start.Height)
                width <- max least (start.Width * scale)
                height <- max least (start.Height * scale)
            // (The side opposite the handle stays where it is.)
            let x = if west then max 0.0 (start.X + start.Width - width) else start.X
            let y = if north then max 0.0 (start.Y + start.Height - height) else start.Y
            { X = x; Y = y; Width = width; Height = height }

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

    let executeName memberName command =
        session.UpdateView(currentWorkbookSnapshot())
        let prepareFull () =
            installWorkbookSnapshot session.Workbook
            convertFormulaSheets ()
            currentWorkbookSnapshot()
        match session.ExecuteName(command, beforeFull = prepareFull) with
        | Result.Error problem -> Host.hostFail memberName problem
        | Ok changed ->
            if changed then
                installWorkbookSnapshot session.Workbook
                adoptConvertedSheets ()
            changed

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
            pushValueNeutral description
            true

    // --- rows and columns: sizes, hidden ones, outline ---------------------------
    // One undo step holds the sizes of the rows and columns and what the
    // sheet keeps beside them (Model.SheetAxes) before and after.
    let axesState (m: SpreadsheetModel) = m.RowHeights, m.ColumnWidths, m.Axes
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
            pushValueNeutral description
            true
    let changeAxes (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let before = axesState model
        model <- change model
        commitAxes description before
    // --- AutoFilter ----------------------------------------------------------------
    // The filter of a sheet is its range, the column filters and the rows
    // they hide (Model.reapplyFilter works those out). A change of any of the
    // three is one undo step.
    let filterState (m: SpreadsheetModel) = m.AutoFilter, m.Filters, m.FilterHiddenRows
    /// Change the active sheet's AutoFilter as one undoable step. False when nothing changed.
    let changeFilter (description: string) (change: SpreadsheetModel -> SpreadsheetModel) : bool =
        let before = filterState model
        model <- change model
        let after = filterState model
        let (rangeBefore, filtersBefore, hiddenBefore), (rangeAfter, filtersAfter, hiddenAfter) = before, after
        let sameHidden = sameRef hiddenBefore hiddenAfter || hiddenBefore = hiddenAfter
        if rangeBefore = rangeAfter && (sameRef filtersBefore filtersAfter || filtersBefore = filtersAfter) && sameHidden then false
        else
            // (SUBTOTAL and AGGREGATE leave out the rows a filter hides.)
            if not sameHidden then recalculateFormulas()
            // In a viewer a filter is how the sheet is looked at: no undo step, no new revision.
            if readOnly then absorbViewState <- true
            else model <- pushCommand description model
            true
    /// Sort a range of the active sheet as one undoable step, then recalculate.
    let sortBlock (memberName: string) (range: CellRange) (keys: SheetOps.SortKey list) (options: SheetOps.SortOptions) (description: string) : bool =
        let sorted = SheetOps.sortRange memberName range keys options model
        if sameRef sorted.Cells model.Cells && sorted.SortState = model.SortState then false
        else
            model <- sorted
            recalculateFormulas()
            model <- { model with RowsSorted = None }
            model <- pushCommand description model
            true
    /// Sort the block of data around a cell by the cell's column (a quick
    /// sort), as one undoable step. With only a column in hand (a click on a
    /// column header) the block is the one around the active cell when that
    /// is in the column, else around the column's first entry.
    let sortActiveSheet (memberName: string) (col: int) (ascending: bool) (description: string) =
        // (A quick sort from a cell with no data around it does nothing; a
        // click on a column header has the column's own data to go to.)
        let active = activeCellAddress()
        let anchor = if memberName = "sort" then sortAnchorForColumn col (Some active) model else Some active
        match anchor |> Option.bind (fun anchor -> SheetOps.quickSortTarget anchor model |> Option.map (fun target -> anchor, target)) with
        | Some (anchor, (range, options)) ->
            let key : SheetOps.SortKey = { Column = anchor.Col; Descending = not ascending; By = SheetOps.SortByValue; Color = "" }
            sortBlock memberName range [ key ] options description |> ignore
        | None -> ()
    /// Text widths for auto-fit: as the grid draws the text (in the cell's font, `cellTextWidth`).
    let fitMeasure : SheetOps.TextMeasure = cellTextWidth
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


    /// What GetSelectionFormat last gave, and what it gave it for (the cells,
    /// the column styles and the selected block): asked again for the same,
    /// it answers without looking at a cell.
    let mutable selectionFormatFor : (obj * obj * (int * int * int * int)) option = None
    let mutable selectionFormatObj : obj = null

    /// What GetStatus last answered and for what (the sheet's cells, the rows
    /// its filter hides and the selection).
    let mutable statusFor : (obj * obj * (int * int * int * int)) option = None
    let mutable statusObj : obj = null
    /// The status of the selection (see GetStatus).
    let selectionStatusObj () : obj =
        let top, left, bottom, right = selectionBounds()
        let block = top, left, min bottom (model.RowCount - 1), min right (model.ColCount - 1)
        match statusFor with
        | Some (cells, hidden, known) when sameRef cells (box model.Cells) && sameRef hidden (box model.FilterHiddenRows) && known = block -> statusObj
        | _ ->
            let status = SheetOps.selectionStatus model block
            let numbers = status.NumericCount > 0
            let made =
                createObj [
                    "sum" ==> (if numbers then box status.Sum else null)
                    "average" ==> (if numbers then box (status.Sum / float status.NumericCount) else null)
                    "count" ==> status.Count
                    "numericCount" ==> status.NumericCount
                    "min" ==> Host.orNull status.Min
                    "max" ==> Host.orNull status.Max
                ]
            statusFor <- Some (box model.Cells, box model.FilterHiddenRows, block)
            statusObj <- made
            made

    /// The model of the sheet at `index` (converting a sheet of an opened
    /// file that was not needed so far).
    let sheetModelAt (index: int) : SpreadsheetModel =
        if index = activeSheetIndex then model
        else
            convertInPlace index
            sheets.[index]

    /// The model of the sheet a host names by an optional index (none: the active sheet).
    let sheetArg (memberName: string) (sheetIndex: obj) : SpreadsheetModel =
        match Host.optionalArg (Host.intArg memberName "sheetIndex") sheetIndex with
        | None -> model
        | Some index when index >= 0 && index < sheets.Length -> sheetModelAt index
        | Some index -> Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)

    // --- find and replace ---------------------------------------------------------
    /// The search Find started (None: none, or StartSearch started one since).
    let mutable hostSearch : HostSearch option = None
    /// The public search owned by the visible bar, not a later host Find call.
    let mutable findSession: HostSearch option = None
    /// The matches of that search on each sheet, for the cells they were found in.
    let searchCache = System.Collections.Generic.Dictionary<SheetId, obj * (CellAddress * string) list>()
    /// One search surface per grid; a shell may host the same node.
    let mutable findBar: Element option = None
    let mutable findHost: Element option = None
    let mutable findRefresh: (unit -> unit) = fun () -> ()
    let mutable findNextAction: (unit -> unit) = fun () -> ()
    let mutable findReplaceAllAction: (unit -> unit) = fun () -> ()
    let mutable findFocus: Element option = None
    let mutable findSelected = false
    let mutable findEpoch = -1

    /// Whether the sheet at an index has a tab.
    let sheetHasTab (index: int) : bool =
        match workbook |> Option.bind (fun wb -> Map.tryFind (sheetIdForIndex index) wb.Sheets) with
        | Some sheet -> sheet.Visible
        | None -> true

    /// The sheets a search reads: the one on display, or every sheet with a tab.
    let searchedSheets (search: HostSearch) : int list =
        if search.Request.Workbook then [ for index in 0 .. sheets.Length - 1 do if index = activeSheetIndex || sheetHasTab index then yield index ]
        else [ activeSheetIndex ]

    /// The matches of a search on one sheet (found again when its cells changed).
    let sheetMatches (search: HostSearch) (index: int) : (CellAddress * string) list =
        let sheetModel = sheetModelAt index
        let id = sheetIdForIndex index
        match Portable.Dict.tryFind id searchCache with
        | ValueSome (cells, found) when obj.ReferenceEquals(cells, sheetModel.Cells) -> found
        | _ ->
            let found = SheetOps.findInSheet search.Finder search.Request.Options search.Request.Range sheetModel
            searchCache.[id] <- (box sheetModel.Cells, found)
            found

    /// Every match of a search as (sheet index, cell, text), sheet by sheet in tab order.
    let allMatches (search: HostSearch) : (int * CellAddress * string) array =
        [| for index in searchedSheets search do
            for addr, text in sheetMatches search index do
                yield index, addr, text |]

    /// The picture of the active sheet a host names by its id.
    let pictureArg (memberName: string) (id: obj) : SheetDrawing =
        match SheetOps.pictureById (Host.intArg memberName "the picture id" id) model with
        | Ok drawing -> drawing
        | Result.Error message -> Host.hostFail memberName message
    /// A size { width, height } a host gives for a picture that is `width0`
    /// by `height0`: one of the two alone keeps the proportions, none (the
    /// argument left out) gives the picture's own size.
    let pictureSizeArg (memberName: string) (size: obj) (width0: float) (height0: float) : float * float =
        if Host.isNullish size then width0, height0
        elif not (Host.isJsObject size) then Host.hostFail memberName "the size is not an object ({ width, height })"
        else
            let read (key: string) =
                Host.optionalArg (Host.numberArg memberName key) (Host.getKey size key)
                |> Option.map (fun value -> if value > 0.0 && value <= 100000.0 then value else Host.hostFail memberName (sprintf "%s must be above 0 (sheet pixels)" key))
            match read "width", read "height" with
            | Some width, Some height -> width, height
            | Some width, None -> width, (if width0 > 0.0 then height0 * width / width0 else height0)
            | None, Some height -> (if height0 > 0.0 then width0 * height / height0 else width0), height
            | None, None -> Host.hostFail memberName "the size has neither a width nor a height"

    // --- the list of sheets ---------------------------------------------------------
    // Sheet-list edits join the same workbook journal as cell and formatting edits.
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
        convertInPlace active
        model <- sheets.[active]
        anchorCell <- sheetAnchors.[active]
        refreshSheetLock ()
        lookedUpCells.Clear()
        if order.[active] <> activeBefore then
            scrollHome ()
            for listener in sheetChangedListeners do
                try listener active with error -> reportError "OnSheetChanged" (box error)
    let executeSheetStructure memberName command =
        session.UpdateView(bookSnapshot ())
        let prepareFull () =
            installBook session.Workbook
            convertFormulaSheets ()
            bookSnapshot ()
        match session.ExecuteStructural(command, beforeFull = prepareFull) with
        | Ok true ->
            installBook session.Workbook
            adoptConvertedSheets ()
            true
        | Ok false -> false
        | Result.Error problem -> Host.hostFail memberName problem



    /// Register a workbook-level operation as one step in the shared journal.
    let runBookStep (description: string) (forward: unit -> unit) =
        sheets.[activeSheetIndex] <- model
        forward ()
        model <- pushCommand description model

    /// Change the page setup of a sheet (the active one, or the one a host
    /// names) as one undoable step. False when nothing changed. A step on
    /// another sheet than the one on display is a step of the workbook.
    let changePageSetup (memberName: string) (sheetIndex: obj) (description: string) (change: PageSetup -> PageSetup) : bool =
        let index =
            match Host.optionalArg (Host.intArg memberName "sheetIndex") sheetIndex with
            | None -> activeSheetIndex
            | Some index when index >= 0 && index < sheets.Length -> index
            | Some index -> Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        let before = (sheetModelAt index).PageSetup
        let after = change before
        if after = before then false
        elif index = activeSheetIndex then
            model <- { model with PageSetup = after }
            pushValueNeutral description
            true
        else
            let id = sheetIdForIndex index
            let put (value: PageSetup) () =
                let current = bookSnapshot ()
                match Map.tryFind id current.Sheets with
                | Some sheet -> installBook { current with Sheets = Map.add id { sheet with Model = { sheet.Model with PageSetup = value } } current.Sheets }
                | None -> ()
            runBookStep description (put after)
            true

    /// Apply a table change and its formula recalculation in one workbook step.
    /// The session retains its persistent before/after workbook state for undo.
    let changeTables (description: string) (op: Workbook -> SheetOps.TableChange) : bool =
        // Formulas of every sheet can name a table, and a table's sheet is changed.
        convertFormulaSheets ()
        for index in 0 .. sheets.Length - 1 do
            if index <> activeSheetIndex && not sheets.[index].Tables.IsEmpty then convertInPlace index
        sheets.[activeSheetIndex] <- model
        let before = bookSnapshot ()
        let edited = (op before).Book
        if obj.ReferenceEquals(edited, before) then false
        else
            runBookStep description (fun () ->
                installBook edited
                recalculateEverything ())
            true

    do
        tableEntryHook.Value <- fun (addr, entry) ->
            let sheetId = sheetIdForIndex activeSheetIndex
            changeTables (sprintf "Edit cell %s" (formatCellRef addr.Row addr.Col)) (fun wb ->
                match SheetOps.tableEntry sheetId addr entry wb with
                | Some change -> change
                | None -> { Book = wb; Cells = [] })

    /// The index of the sheet a table is on (throws for a table of no sheet).
    let tableSheetIndex (memberName: string) (table: SheetOps.TableInfo) : int =
        let index = sheetIndexOfId table.SheetId
        if index < 0 then Host.hostFail memberName (sprintf "the table \"%s\" is on no sheet" table.Name)
        index

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
                "kind" ==> sheetKindAt index
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
                    if sheetsListeners.Contains listener then (try listener (box infos) with error -> reportError "OnSheetsChange" (box error))


    let updateFormulaHelp () =
        match layers, model.EditingCell with
        | Some layer, Some _ when not disposed ->
            let caret = unbox<int> (getProp layer.EditInput "selectionStart")
            match formulaHelpAt model.EditBuffer caret with
            | Some (info, argument) ->
                let label = functionArgumentHelp info argument
                setTextContent layer.EditHelp (info.Name + " — " + label + " · " + info.Description)
                let grid = gridContainer |> Option.get
                let editRect = boundingRect layer.EditInput
                let gridRect = boundingRect grid
                let x = max 4.0 (min (getClientWidth grid - 306.0) (editRect.[0] - gridRect.[0]))
                let below = editRect.[1] - gridRect.[1] + editRect.[3] + 4.0
                setCssText layer.EditHelp
                    ("position:absolute;left:" + string (int (getScrollLeft grid + x)) + "px;top:" + string (int (unbox<float> (getProp grid "scrollTop") + below)) +
                     "px;box-sizing:border-box;max-width:min(300px,calc(100% - 12px));padding:5px 8px;background:#fff;border:1px solid #217346;" +
                     "box-shadow:0 3px 10px #0002;white-space:normal;font:12px Arial,sans-serif;color:#202124;z-index:30;pointer-events:none")
                let helpHeight = (boundingRect layer.EditHelp).[3]
                if below + helpHeight > getClientHeight grid then
                    let above = max 4.0 (editRect.[1] - gridRect.[1] - helpHeight - 4.0)
                    setStyle layer.EditHelp "top" (string (int (unbox<float> (getProp grid "scrollTop") + above)) + "px")
            | None -> setStyle layer.EditHelp "display" "none"
        | Some layer, _ -> setStyle layer.EditHelp "display" "none"
        | _ -> ()

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
                    if editStateListeners.Contains listener then (try listener state with error -> reportError "OnEditStateChange" (box error))


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
                // (A cell an array spilled into holds nothing of its own: the entry starts empty, as in Excel.)
                | None -> if Map.containsKey addr model.Spill then "" else SpreadsheetInput.editTextFor model.Date1904 (getCell addr model)
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
            updateFormulaHelp ()
            notifyEditState ()

    and startEditing (row: int) (col: int) = startEditingWith row col None

    /// Finish editing. `preserveInvalid` keeps a rejected entry visible for
    /// workflows (such as File > Save/Open) that must not discard its buffer.
    and finishEditingWith (preserveInvalid: bool) (save: bool) : bool =
        match model.EditingCell with
        | None -> true
        | Some addr ->
            let buffer = model.EditBuffer
            model <- { model with EditingCell = None; EditBuffer = "" }
            let accepted = if save then commitInput addr buffer true true else true
            if not accepted && preserveInvalid then
                model <- { model with EditingCell = Some addr; EditBuffer = buffer }
            scheduleNoticeDismiss()
            if accepted || not preserveInvalid then
                match layers, contentElement with
                | Some layer, Some content when isActiveElement layer.EditInput -> focusQuietly content
                | _ -> ()
            render()
            if not accepted && preserveInvalid then
                layers |> Option.iter (fun layer -> focus layer.EditInput)
            updateFormulaHelp ()
            notifyEditState ()
            accepted

    and finishEditing (save: bool) : bool = finishEditingWith false save

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

    /// Copy displayed flavours while the session owns clipboard identity.
    and copySelection (isCut: bool) : string * string =
        let text = selectionText ()
        let html = selectionHtml ()
        session.UpdateView(currentWorkbookSnapshot())
        match session.CaptureClipboard(sheetIdForIndex activeSheetIndex, isCut && not readOnly, text, beforeFull = prepareClipboardFull) with
        | Result.Error problem -> failwith problem
        | Ok _ ->
            installWorkbookSnapshot session.Workbook
            adoptConvertedSheets ()
        text, html

    /// Handle keyboard events
    and handleKeyDown (e: obj) =
        // A key a macro bound (Application.OnKey) runs its procedure while macros may run; "" does nothing.
        let bound =
            if macrosEnabled && macroKeys.Count > 0 && model.EditingCell.IsNone then
                pressedOnKey (getEventKey e) (eventFlag e "ctrlKey" || eventFlag e "metaKey") (eventFlag e "shiftKey") (eventFlag e "altKey")
                |> Option.bind (fun name -> match macroKeys.TryGetValue name with | true, p -> Some p | _ -> None)
            else None
        match bound with
        | Some procedure ->
            preventDefault e
            if procedure <> "" then runMacroWhenIdle procedure
        | None -> handleGridKeyDown e
    and handleGridKeyDown (e: obj) =
        let key = getEventKey e
        let isModifier = eventFlag e "ctrlKey" || eventFlag e "metaKey"
        let shift = eventFlag e "shiftKey"
        let lower = if key.Length = 1 then key.ToLowerInvariant() else key

        // A selected picture takes Delete and Escape.
        let onPicture =
            match selectedPicture, model.EditingCell with
            | Some id, None when not isModifier && (key = "Escape" || ((key = "Delete" || key = "Backspace") && not readOnly)) ->
                preventDefault e
                if key <> "Escape" then
                    (try executeDrawing "picture" "Delete picture" (Rbt.Spreadsheet.Session.DeletePicture id) |> ignore
                     with _ -> ())
                selectedPicture <- None
                render()
                true
            | _ -> false
        if onPicture then () else

        // The menu key / Shift+F10: the context menu of the active cell (of
        // the selected picture), when a host shows one.
        let onMenu =
            if contextListeners.Count > 0 && model.EditingCell.IsNone && (key = "ContextMenu" || (key = "F10" && shift)) then
                preventDefault e
                keyboardMenuAt <- nowMs ()
                let active = activeCellAddress()
                let pointOf (selector: string) =
                    match contentElement, gridContainer with
                    | Some content, Some grid ->
                        let node = queryOne content selector
                        if isNull (box node) then [| getBoundingLeft grid + 8.0; getBoundingTop grid + 8.0 |] else clientCenterOf node
                    | _ -> [| 0.0; 0.0 |]
                match selectedPicture |> Option.bind (fun id -> SheetOps.drawingById id model) with
                | Some drawing ->
                    let point = pointOf (sprintf ".spreadsheet-picture[data-drawing-id=\"%d\"]" drawing.Id)
                    raiseContextMenu "picture" drawing.Anchor.FromRow drawing.Anchor.FromCol (Some drawing.Id) true point.[0] point.[1] activeSheetIndex |> ignore
                | None ->
                    revealCell active.Row active.Col
                    render()
                    let point = pointOf (sprintf ".spreadsheet-cell[data-row=\"%d\"][data-col=\"%d\"]" active.Row active.Col)
                    raiseContextMenu "cell" active.Row active.Col None true point.[0] point.[1] activeSheetIndex |> ignore
                true
            else false
        if onMenu then () else

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
            | "Tab" when not shift && not isModifier && not (eventFlag e "altKey") && not (eventFlag e "isComposing") ->
                match layers with
                | Some layer ->
                    let caret = unbox<int> (getProp layer.EditInput "selectionStart")
                    let selectionEnd = unbox<int> (getProp layer.EditInput "selectionEnd")
                    match formulaFunctionCompletion model.EditBuffer caret with
                    | Some (start, info) when caret = selectionEnd ->
                        preventDefault e
                        let replacement = info.Name + "("
                        model <- { model with EditBuffer = model.EditBuffer.Substring(0, start) + replacement + model.EditBuffer.Substring(caret) }
                        setValue layer.EditInput model.EditBuffer
                        setProp layer.EditInput "selectionStart" (box (start + replacement.Length))
                        setProp layer.EditInput "selectionEnd" (box (start + replacement.Length))
                        notifyEditState ()
                        updateFormulaHelp ()
                    | _ ->
                        preventDefault e
                        if finishEditing true then moveWithinSelection 0 1
                | None ->
                    preventDefault e
                    if finishEditing true then moveWithinSelection 0 1
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
                | "f" | "h" ->
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
                // Alt+Down on a header cell with a filter button presses the button.
                | "ArrowDown" when eventFlag e "altKey" && (filterButtonAt active.Row active.Col).IsSome ->
                    preventDefault e
                    filterButtonAt active.Row active.Col |> Option.iter pressFilterButton
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
            let generation = clipboardGeneration
            let pasteOwn () =
                if pasteClipboardAtSelection() then
                    render()
                    true
                else false
            promiseThen (readSystemClipboard ()) (fun flavours ->
                if disposed || generation <> clipboardGeneration || readOnly then false
                elif isNull (box flavours) || (flavours.[0] = "" && flavours.[1] = "") then
                    match session.ClipboardSource(None) with
                    | Rbt.Spreadsheet.Session.InternalPaste -> pasteOwn ()
                    | _ -> false
                elif pasteFlavours flavours.[0] flavours.[1] then
                    render()
                    true
                else false)

    and undoStep () =
        observeRevision ()
        if session.Undo() then
            installBook session.Workbook
            render()

    and redoStep () =
        observeRevision ()
        if session.Redo() then
            installBook session.Workbook
            render()

    /// Handle input change (while editing)
    and handleInput (e: obj) =
        let target = getEventTarget e
        let value = getValue target
        model <- { model with EditBuffer = value }
        updateFormulaHelp ()
        notifyEditState ()

    /// Draw a cell into its node: position, look and text. Selection is not
    /// drawn here (it is an overlay), so moving it draws no cell. How the
    /// cell looks - fill, font, colour, alignment, indent, rotation, wrap,
    /// its number-formatted text, its table's look and its conditional
    /// formats - is worked out by the painters' rule (`Paint.cellLook`);
    /// this maps it to CSS.
    /// A face a cell is drawn in, asked of the host's loader (`SetFontLoader`) the first time it is
    /// drawn, unless the system has its family; once it is loaded every cell is drawn again.
    and wantFace (family: string option) (bold: bool) (italic: bool) =
        match family with
        | Some family when not (isNull fontLoader) && family.Trim() <> "" ->
            let family = family.Trim()
            let name = family.ToLowerInvariant()
            let fromSystem =
                match Portable.Dict.tryFind name systemFamilies with
                | ValueSome known -> known
                | _ ->
                    let known = fontFamilyAvailable family
                    systemFamilies.[name] <- known
                    known
            if not fromSystem && wantedFaces.Add (name + "|" + (if bold then "b" else "") + (if italic then "i" else "")) then
                loadFace fontLoader family bold italic (fun () ->
                    if not fontRedrawPending && not disposed then
                        fontRedrawPending <- true
                        // (The faces of a sheet arrive together: one drawing for them.)
                        setTimeoutMs (fun () ->
                            fontRedrawPending <- false
                            if not disposed then
                                renderedStyleInputs <- [||]
                                render ()) 30 |> ignore)
        | _ -> ()

    and drawCell (layer: GridLayers) (node: CellNode) (row: int) (col: int) (x: float) (y: float) (w: float) (h: float) (pane: int)
                 (cellData: CellData option) =
        let cell = node.Element
        // (`cellData` is what the cell holds; the look applies the sheet's view options.)
        let storedData = cellData
        // (The measurer as a lambda here, not one kept in a module value: Fable keeps a function
        // value stored in an option curried and wraps it again - uncurry5, a WeakMap entry - every
        // time it is passed; a lambda written at the call is passed as it is.)
        // The text is measured as Excel draws it at the sheet's zoom (`ExcelText`), in screen pixels, and
        // a number fitted to the cell's width on screen as Excel fits it (`CellText.NumberFit`).
        // (Excel's column on screen is its whole pixels at 100 % x the zoom, rounded - 75 px at 90 % are
        // 68 -, whatever the columns to its left.)
        let zoom = appliedZoom
        let screenWidth = System.Math.Floor (System.Math.Round (System.Math.Round w * zoom, 9) + 0.5)
        let look =
            SpreadsheetPaint.cellLookFitted (Some ({ Zoom = zoom } : SpreadsheetCellText.NumberFit)) model row col cellData
                (Some (fun text fontPt fontName bold italic -> excelTextWidth zoom text fontPt fontName bold italic)) screenWidth
        if node.Pane <> pane || not (isChildOf cell layer.Panes.[pane]) then appendChild layer.Panes.[pane] cell
        let addr = { Row = row; Col = col }
        // What every cell has in common is in the grid's style sheet (see
        // Mount); the node's own style is its place and what its format adds.
        let css = System.Text.StringBuilder()
        let px (value: float) = string (int (value + 0.5)) + "px"
        let add (name: string) (value: string) = css.Append(name).Append(':').Append(value).Append(';') |> ignore
        // (From the rounded left edge to the rounded right edge: rounding the
        // width on its own leaves a pixel between this cell and the next
        // when a column's width is not a whole number - a white line through
        // a row of filled cells, and through text that runs across it.)
        let span (start: float) (size: float) = string (int (start + size + 0.5) - int (start + 0.5)) + "px"
        if appliedZoom = 1.0 then
            add "left" (px x)
            add "top" (px y)
            add "width" (span x w)
            add "height" (span y h)
        else
            // Zoomed, the edges are rounded to the screen's pixels instead (the content is scaled by
            // CSS zoom): a whole sheet pixel scaled by 85 % is not a whole screen pixel, and two
            // cells' fills, each rounded on its own, left a white line between them (acre-apartment's
            // sheets at 85 %; Excel's fills meet).
            let device (value: float) = System.Math.Floor (value * appliedZoom + 0.5)
            let at (value: float) = string (device value / appliedZoom) + "px"
            let size (start: float) (length: float) = string ((device (start + length) - device start) / appliedZoom) + "px"
            add "left" (at x)
            add "top" (at y)
            add "width" (size x w)
            add "height" (size y h)
        // The look: fill and font.
        match look.Fill with Some fill -> add "background-color" fill | None -> ()
        match look.TextColor with Some color -> add "color" color | None -> ()
        if look.Bold then add "font-weight" "bold"
        if look.Italic then add "font-style" "italic"
        if look.Underline || look.UnderlineDouble || look.Strikethrough then
            let lines =
                (if look.Underline || look.UnderlineDouble then "underline " else "")
                + (if look.Strikethrough then "line-through " else "")
                + (if look.UnderlineDouble then "double" else "")
            add "text-decoration" (lines.TrimEnd())
        if not (isNull fontLoader) && look.Text <> "" then wantFace look.FontName look.Bold look.Italic
        let drawnFamily, drawnScale = drawnFace look.FontSize zoom look.FontName
        if ExcelText.isSmallFonts look.FontSize zoom then add "font-family" drawnFamily
        else
            match look.FontName with
            | Some family when family <> "" -> add "font-family" drawnFamily
            | _ -> ()
        // The cell's size in points drawn as Excel draws it: at a whole number of screen pixels,
        // round(points x 96/72 x zoom) (`ExcelText.pixelSize`: 11 pt is 15 px at 100 % and 12 px at
        // both 80 % and 85 %); in sheet pixels, which the content's zoom scales. (Calibri where the
        // browser has none smaller, as tall as Calibri's letters, `drawnFontScale`: a size of its own
        // rather than `font-size-adjust`, so that the text is measured as it is drawn.)
        let fontPx = float (ExcelText.pixelSize look.FontSize zoom) / zoom * drawnScale
        if look.Text <> "" then add "font-size" (cssPxFine fontPx)
        // Across the cell (the style sheet's default is the left).
        match look.Horizontal with
        | "distributed" | "justify" -> add "text-align" "justify"
        | _ ->
            match look.Align with
            | SpreadsheetLayout.AlignRight -> add "text-align" "right"
            | SpreadsheetLayout.AlignCenter -> add "text-align" "center"
            | SpreadsheetLayout.AlignLeft -> ()
        // Where the text sits in the height of the cell. A line of text is
        // as tall as the line of an auto-fitted row of its size (20 px for
        // 11 pt: `SheetOps.lineHeightPxOfPoints`). Bottom, Excel's default,
        // puts the last line at the bottom of the cell; top puts the first
        // a padding below the top; middle (and a row too low for a line)
        // centres the text by the cell's line height (the node's default
        // look). Wrapped text is a column of lines, placed as a block, from
        // edge to edge of the cell (a row auto-fitted to n lines is n lines tall).
        let lineHeight = SheetOps.lineHeightPxOfPoints look.FontSize
        if look.Wrap then
            // (Its own line breaks are kept.)
            add "white-space" "pre-line"
            add "word-wrap" "break-word"
            // Excel wraps text to the width of the cell less 3 px each side (wrap-threshold, checked
            // in Excel: bold Calibri 11 "Releasing Cost", 89 px of ink, is one line in a 95-px column
            // and two in a 93-px one), not the 5 px a line of text is set in from the cell's edge.
            if look.IndentLeft = 0.0 then add "padding-left" "3px"
            if look.IndentRight = 0.0 then add "padding-right" "3px"
            add "line-height" (px lineHeight)
            add "display" "flex"
            add "flex-direction" "column"
            add "justify-content"
                (match look.VerticalAlign with
                 | SpreadsheetPaint.VAlignTop -> "flex-start"
                 | SpreadsheetPaint.VAlignMiddle -> "center"
                 | SpreadsheetPaint.VAlignBottom -> "flex-end")
            if look.VerticalAlign <> SpreadsheetPaint.VAlignTop then add "padding-top" "0px"
            add "padding-bottom" "0px"
        else
            match look.VerticalAlign with
            | SpreadsheetPaint.VAlignTop -> add "line-height" (px lineHeight)
            | SpreadsheetPaint.VAlignBottom when h >= lineHeight ->
                add "line-height" (px lineHeight)
                // (A pixel lower than the line's own box: Excel's baseline is 4 px above the
                // bottom of a 20 px row, 5 px when the line ends at the edge.)
                add "padding-top" (px (h - lineHeight + 1.0))
                add "padding-bottom" "0px"
            | _ -> add "line-height" (px (h - 4.0))
        // An indent pads the edge the text is aligned to.
        // (Excel sets a line of text in 2 px from the cell's left side and 3 px from its right - the
        // grid line at its right is the cell's -, in sheet pixels at any zoom: the grid's style sheet.)
        if look.IndentRight > 0.0 then add "padding-right" (px (3.0 + look.IndentRight))
        elif look.IndentLeft > 0.0 then add "padding-left" (px (2.0 + look.IndentLeft))
        // The table's look of the cell, for its edges and theirs.
        let tableLook = look.Table
        // Borders and grid lines: one line per edge. A cell draws the edge
        // above it and the edge to its left - each the one border the two
        // cells that share it resolve to (Layout.edgeAbove / edgeLeft), or
        // their grid line - inside its own box; the edges below and to its
        // right are the next cells'. (A cell of the last row / column draws
        // those too.) Grid lines are an inset shadow of the cell; a border is
        // a child element over it, the stronger of the two on top where they
        // meet, so a line runs on through the corners of the cells it passes.
        let endRow, endCol =
            if model.MergedCells.IsEmpty then row, col
            else
                match findMergedCell addr model with
                | Some merge -> merge.EndRow, merge.EndCol
                | None -> row, col
        let edgeChildren = ResizeArray<int * string>()
        let borderWidth (border: BorderStyle) = match borderCssParts border.Style with Some (width, _) -> width | None -> 0
        // A line is as many screen pixels wide at every zoom: Excel draws a medium line 2 px wide at 70 %
        // too (choudhary-real-estate-leasing's "Inc St", the line above "End of Sheet"), where the
        // content's CSS zoom made it 1.4 px - drawn 1 px.
        let zoomedBorderCss (border: BorderStyle) =
            if appliedZoom = 1.0 then borderCss border
            else
                match borderCssParts border.Style with
                | Some (width, line) -> string (float width / appliedZoom) + "px " + line + " " + (border.Color |> Option.defaultValue "#000000")
                | None -> "none"
        // The runs of one kind along an edge of the cell: (offset, length, kind).
        let runsAlong = SpreadsheetLayout.edgeRuns
        let gridColor = defaultGridColor
        let shadows = ResizeArray<string>()
        // The colour of the grid line drawn for a thin line of a table's style, on the top (0) / left (1) edge.
        let tableShadow : string array = [| null; null |]
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
            | [ (_, _, SpreadsheetLayout.EdgeGrid true) ] -> shadows.Add (gridShadow + (match (if side < 2 then tableShadow.[side] else null) with null -> gridColor | color -> color))
            | _ ->
                for offset, length, kind in runs do
                    match kind with
                    | SpreadsheetLayout.EdgeBorder border -> edgeChildren.Add (borderPriority border.Style, place offset length (borderWidth border) (zoomedBorderCss border))
                    | SpreadsheetLayout.EdgeGrid true -> edgeChildren.Add (0, place offset length 1 ("1px solid " + gridColor))
                    | _ -> ()
        let columnStart (c: int) = SpreadsheetLayout.columnLeft model c
        let rowStart (r: int) = SpreadsheetLayout.rowTop model r
        // The cells the edges are shared with: the one above, the one to the
        // left, and the one above that (for the corner between them).
        let above = SpreadsheetLayout.rowShownBefore model row
        let before = SpreadsheetLayout.columnShownBefore model col
        let aboveRaw, leftRaw, cornerRaw = cellAt above col, cellAt row before, cellAt above before
        // The look of a neighbour's table, for the edges the cell shares with it.
        let lookOf (r: int) (c: int) : TableLookEntry option =
            if model.Tables.IsEmpty || r < 0 || c < 0 then None else SpreadsheetLayout.tableEntryAt model.Tables r c
        let aboveLook, leftLook, cornerLook = lookOf above col, lookOf row before, lookOf above before
        let plainSheet = model.MergedCells.IsEmpty
        // (Without merged cells an edge follows from the two cells' formats - each with its table's look
        // under it: its own line, else the look's - with no further look-ups: the painters' rule,
        // `Layout.edgeBetween`, of the cells at hand.)
        let border (data: CellData option) (pick: CellFormat -> BorderStyle option) = data |> Option.bind (fun data -> pick data.Format)
        let ownLines (first: CellData option) (second: CellData option) (horizontal: bool) (hasFirst: bool) =
            (if hasFirst then border first (fun format -> if horizontal then format.BorderBottom else format.BorderRight) else None),
            border second (fun format -> if horizontal then format.BorderTop else format.BorderLeft)
        let edgeBetween (first: CellData option, firstLook: TableLookEntry option) (second: CellData option, secondLook: TableLookEntry option) (horizontal: bool) (hasFirst: bool) =
            SpreadsheetLayout.edgeBetween model.ShowGridLines first firstLook second secondLook horizontal hasFirst
        // A cell's own edge (top or left) where the line is a thin one of a table's style and neither cell
        // states a line of its own: drawn as the grid line is, in the line's colour (no node of its own).
        let ownEdge (side: int) (first: CellData option, firstLook: TableLookEntry option) (horizontal: bool) (hasFirst: bool) =
            let edge = edgeBetween (first, firstLook) (cellData, tableLook) horizontal hasFirst
            match edge with
            | SpreadsheetLayout.EdgeBorder line when line.Style = "thin" && line.Color.IsSome && (match ownLines first cellData horizontal hasFirst with None, None -> true | _ -> false) ->
                tableShadow.[side] <- line.Color.Value
                SpreadsheetLayout.EdgeGrid true
            | _ -> edge
        let topRuns =
            if plainSheet then [ 0.0, w, ownEdge 0 (aboveRaw, aboveLook) true (above >= 0) ]
            else runsAlong col endCol columnStart (fun c -> SpreadsheetLayout.edgeAbove model row c)
        let leftRuns =
            if plainSheet then [ 0.0, h, ownEdge 1 (leftRaw, leftLook) false (before >= 0) ]
            else runsAlong row endRow rowStart (fun r -> SpreadsheetLayout.edgeLeft model r col)
        // The two edges that end at the cell's top-left corner without being its own:
        // the one coming down from above and the one coming in from the left.
        let noEdge = SpreadsheetLayout.EdgeGrid false
        let verticalEdge =
            if above < 0 || before < 0 then noEdge
            elif plainSheet then (if aboveRaw.IsSome || cornerRaw.IsSome || aboveLook.IsSome || cornerLook.IsSome then edgeBetween (cornerRaw, cornerLook) (aboveRaw, aboveLook) false true else noEdge)
            else SpreadsheetLayout.edgeLeft model above col
        let horizontalEdge =
            if above < 0 || before < 0 then noEdge
            elif plainSheet then (if leftRaw.IsSome || cornerRaw.IsSome || leftLook.IsSome || cornerLook.IsSome then edgeBetween (cornerRaw, cornerLook) (leftRaw, leftLook) true true else noEdge)
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
                    // (A thin line of a table's style drawn as the cell's grid line covers the corner pixel too:
                    // inside a table with inner lines that is every cell, and a node at each corner cost a
                    // screenful of cells a node each.)
                    || (not (isNull tableShadow.[0]) && cornerHeight <= 1)
                    || (not (isNull tableShadow.[1]) && cornerWidth <= 1)
                if not covered && cornerWidth > 0 && cornerHeight > 0 then
                    let stronger = if borderPriority vertical.Style > borderPriority horizontal.Style then vertical else horizontal
                    edgeChildren.Add (100, "position:absolute;pointer-events:none;left:0;top:0;width:" + string cornerWidth + "px;height:" + string cornerHeight
                                           + "px;background:" + (stronger.Color |> Option.defaultValue "#000000"))
            | _ -> ()
        // (The style sheet gives a cell the two grid lines most cells have.)
        if not (shadows.Count = 2 && gridColor = defaultGridColor && isNull tableShadow.[0] && isNull tableShadow.[1] && shadows.[0].StartsWith "inset 0 1px" && shadows.[1].StartsWith "inset 1px") then
            add "box-shadow" (if shadows.Count = 0 then "none" else String.concat "," shadows)
        // The data bar and the icon of its conditional formats (`Paint.cellLook`).
        let icon = look.Icon
        let mutable bar = -1
        let mutable barDraw : DataBarDraw option = None
        match look.DataBar with
        | Some draw when draw.From = 0.0 && not draw.Gradient && draw.Border.IsNone && draw.Axis.IsNone ->
            // A plain bar from the left edge, behind the text, two pixels in from the cell's top and
            // bottom as Excel draws it (the painter's too).
            let percent = int (System.Math.Round (draw.To * 100.0))
            add "background-image" ("linear-gradient(to right, " + draw.Color + " " + string percent + "%, transparent " + string percent + "%)")
            add "background-repeat" "no-repeat"
            add "background-position" "0 2px"
            add "background-size" "100% calc(100% - 4px)"
            bar <- percent
        | Some draw ->
            // A bar of Excel 2010+ (an axis, negative values, a border, a gradient): drawn by nodes of its own.
            bar <- int (System.Math.Round ((draw.To - draw.From) * 100.0))
            barDraw <- Some draw
        | None -> ()
        // The text: number-formatted (#### for a number that does not fit), none where a data bar
        // or an icon set hides the value.
        let displayText = look.Text
        // (In sheet pixels: what Excel draws on screen at the zoom, unscaled.)
        let measureText (text: string) = excelTextWidth zoom text look.FontSize look.FontName look.Bold look.Italic / zoom
        // Drawn as wide as Excel draws it: the browser's own advances evened out to Excel's over the
        // line (`excelLetterSpacing`), unless the text is drawn smaller to fit below.
        let mutable spacing = excelLetterSpacing zoom displayText look.FontSize look.FontName look.Bold look.Italic / zoom
        let mutable squeezed = false
        // Shrink to fit: one line of text wider than its cell is drawn smaller
        // (never larger), down to what fits between the cell's paddings.
        if look.Shrink && displayText <> "" then
            let room = w - 5.0 - look.IndentLeft - look.IndentRight
            let width = measureText displayText
            if room > 0.0 && width > room then
                spacing <- 0.0
                add "font-size" (string (System.Math.Floor (max 1.0 (fontPx * room / width * 0.98) * 100.0) / 100.0) + "px")
        // Turned text: a line through the middle of the cell at the format's
        // angle (stacked text: upright letters one under the other).
        let rotation = if displayText <> "" then look.Rotation else 0
        // A number is never cut off or run over its neighbours: its text was chosen to fit the
        // cell as Excel draws it (`CellText.NumberFit`: the room between the cell's paddings).
        // Where it is wider than that all the same (a merged cell's run, a rounding), the
        // paddings give way; where it is wider than the cell, it is drawn smaller.
        let isNumber = match unwrapFormulaValue look.Value with CellValue.Number _ -> true | _ -> false
        // An accounting format's fill (`_($* #,##0_)`): the currency sign at the cell's left edge and
        // the number at its right, whatever the font (`CellText` pads the text with spaces for it,
        // which the browser would run together). (The sign's side, the number's side.)
        let accounting =
            match look.Format.NumberFormat with
            | Some code when isNumber && rotation = 0 && not look.Wrap && code.Contains "* " ->
                let at = displayText.IndexOf "$ "
                if at < 0 then None
                else Some (displayText.Substring(0, at + 1), displayText.Substring(at + 1).TrimStart(' '))
            | _ -> None
        // (A number centred across the selection was fitted to the run's width: `Paint.cellLook`.)
        if isNumber && displayText <> "" && rotation = 0 && not look.Wrap && not look.Shrink && icon.IsNone && look.AcrossTo <= col then
            let indent = look.IndentLeft + look.IndentRight
            let width = measureText (match accounting with Some (sign, number) -> sign + " " + number | None -> displayText)
            if width > w - 5.0 - indent then
                squeezed <- true
                if look.IndentRight = 0.0 then add "padding-right" "1px"
                if look.IndentLeft = 0.0 then add "padding-left" "1px"
                let room = w - 2.0 - indent
                if room > 0.0 && width > room then
                    spacing <- 0.0
                    add "font-size" (cssPx (System.Math.Floor (max 1.0 (fontPx * room / width * 0.98) * 100.0) / 100.0))
        // Centred across the selection: the text is let out of its box over the run of cells it is
        // centred across (`Paint.cellLook`) and centred there, in a box of its own that spans them.
        let across =
            if look.AcrossTo > col && displayText <> "" && rotation = 0 && endRow = row && endCol = col then
                SpreadsheetLayout.columnLeft model (look.AcrossTo + 1) - SpreadsheetLayout.columnLeft model (col + 1)
            else 0.0
        if across > 0.0 then
            add "overflow" "visible"
            // (A cell node's `contain: strict` cuts what it paints at its box: the text centred across the
            // run is beyond it.)
            add "contain" "layout style"
            add "z-index" "1"
            add "clip-path" ("inset(0px " + string (-(int across)) + "px 0px 0px)")
        // Text overflow: text wider than its cell runs over the empty cells
        // beside it (to the right when left-aligned, to the left when
        // right-aligned, both ways when centred) up to the first cell that
        // holds something. The node keeps its place; its text is let out of
        // its box and clipped at the run of cells it may cover.
        let mutable wide = false
        let mutable crossed : (float * int) list = []
        if displayText <> "" && endRow = row && endCol = col && icon.IsNone && across = 0.0 then
            match SpreadsheetLayout.textOverflowSides look.Format look.Value with
            // (Not measured when even the widest glyphs would fit the cell. An indent is room the text
            // does not have in its cell: vertex42-home-mortgage's indented "Mortgage Information" was cut short.)
            | Some (toLeft, toRight) when float displayText.Length * fontPx * 1.05 > w - 5.0 - look.IndentLeft - look.IndentRight ->
                let excess = measureText displayText - (w - 5.0 - look.IndentLeft - look.IndentRight)
                if excess > 0.0 then
                    wide <- true
                    // (Past the cell's edge: its own padding on that side is its to use.)
                    let needLeft = if toLeft then (if toRight then excess / 2.0 else excess) - 2.0 else 0.0
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
        // Rich text: a text in runs, each in its own look (`RichRuns`).
        let runs =
            if model.RichRuns.IsEmpty || rotation <> 0 then None
            else
                match Map.tryFind addr model.RichRuns, unwrapFormulaValue look.Value with
                | Some runs, CellValue.String text when not runs.IsEmpty && displayText = text && (runs |> List.map fst |> String.concat "") = text -> Some runs
                | _ -> None
        if accounting.IsSome && across = 0.0 then
            add "display" "flex"
            add "justify-content" "space-between"
            add "white-space" "pre"
        node.Wide <- wide || across > 0.0
        // (Rich text is spaced run by run, below. The spacing follows a line's last character too: a
        // right-aligned or centred line has its right padding less it, so that its characters stand
        // where Excel's do.)
        if spacing <> 0.0 && runs.IsNone then
            add "letter-spacing" (cssPxFine spacing)
            if not look.Wrap && not squeezed && rotation = 0 && across = 0.0
               && (look.Align = SpreadsheetLayout.AlignRight || look.Align = SpreadsheetLayout.AlignCenter) then
                add "padding-right" (cssPxFine (3.0 + look.IndentRight - spacing))
        setCssText cell (css.ToString())
        setAttribute cell "data-row" (string row)
        setAttribute cell "data-col" (string col)
        if bar >= 0 then setAttribute cell "data-cf-bar" (string bar) else removeAttribute cell "data-cf-bar"
        // (This also drops the icon and comment marker of the node's previous cell.)
        let runSpan (text: string, run: RunOverride) =
            let span = createElement "span"
            setTextContent span text
            let size = run.FontSize |> Option.defaultValue look.FontSize
            let family = run.FontName |> Option.orElse look.FontName
            let runBold = run.Bold |> Option.defaultValue look.Bold
            let runItalic = run.Italic |> Option.defaultValue look.Italic
            if not (isNull fontLoader) then wantFace family (run.Bold |> Option.defaultValue look.Bold) (run.Italic |> Option.defaultValue look.Italic)
            let small = run.VertAlign = Some VertAlignSuperscript || run.VertAlign = Some VertAlignSubscript
            let lines =
                [ if run.Underline = Some true || run.UnderlineDouble = Some true then "underline"
                  if run.Strikethrough = Some true then "line-through" ]
            setCssText span
                (String.concat ";"
                    [ (match run.Bold with Some true -> "font-weight:bold" | Some false -> "font-weight:normal" | None -> "")
                      (match run.Italic with Some true -> "font-style:italic" | Some false -> "font-style:normal" | None -> "")
                      (if lines.IsEmpty then "" else "text-decoration:" + String.concat " " lines + (if run.UnderlineDouble = Some true then " double" else ""))
                      (match run.Color with Some color -> "color:" + color | None -> "")
                      (let points = size * (if small then 0.6 else 1.0)
                       if ExcelText.isSmallFonts points zoom || run.FontName.IsSome then "font-family:" + fst (drawnFace points zoom family) else "")
                      // (At Excel's whole pixel size, as wide as Excel draws it: as the cell's text.)
                      (let points = size * (if small then 0.6 else 1.0)
                       "font-size:" + cssPxFine (float (ExcelText.pixelSize points zoom) / zoom * snd (drawnFace points zoom family)))
                      (let points = size * (if small then 0.6 else 1.0)
                       let runSpacing = excelLetterSpacing zoom text points family runBold runItalic / zoom
                       if runSpacing = 0.0 then "" else "letter-spacing:" + cssPxFine runSpacing)
                      (match run.VertAlign with Some VertAlignSuperscript -> "vertical-align:super" | Some VertAlignSubscript -> "vertical-align:sub" | _ -> "") ]
                 |> fun css -> css.Trim(';'))
            span
        if rotation = 0 && across = 0.0 then
            match runs, accounting with
            | Some runs, _ ->
                // (In one line box of their own: a wrapped cell lays its content out as a column.)
                setTextContent cell ""
                let text = createElement "span"
                for run in runs do appendChild text (runSpan run)
                appendChild cell text
            | None, Some (sign, number) ->
                setTextContent cell ""
                for part in [ sign; number ] do
                    let span = createElement "span"
                    setTextContent span part
                    appendChild cell span
            | None, None -> setTextContent cell displayText
        elif rotation = 0 then
            setTextContent cell ""
            let centred = createElement "span"
            setClassName centred "spreadsheet-across"
            setTextContent centred displayText
            setCssText centred ("position:absolute;left:0;top:0;right:" + string (-(int across)) + "px;padding:inherit;text-align:center;white-space:nowrap;pointer-events:none")
            appendChild cell centred
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
        match barDraw with
        | Some draw ->
            let percent (fraction: float) = string (System.Math.Round (fraction * 10000.0) / 100.0) + "%"
            let fill = createElement "div"
            setClassName fill "spreadsheet-cf-bar"
            let background =
                if draw.Gradient then "linear-gradient(to " + (if draw.FadeRight then "right" else "left") + ", " + draw.Color + ", #FFFFFF)"
                else draw.Color
            setCssText fill
                ("position:absolute;box-sizing:border-box;top:2px;bottom:2px;left:" + percent draw.From + ";width:" + percent (draw.To - draw.From)
                 + ";background:" + background + (match draw.Border with Some line -> ";border:1px solid " + line | None -> "")
                 + ";z-index:-1;pointer-events:none")
            setAttribute fill "data-cf-bar-from" (percent draw.From)
            setAttribute fill "data-cf-bar-color" draw.Color
            if draw.To > draw.From then appendChild cell fill
            match draw.Axis with
            | Some (at, color) ->
                let axis = createElement "div"
                setClassName axis "spreadsheet-cf-axis"
                setCssText axis ("position:absolute;top:0;bottom:0;left:" + percent at + ";width:0;border-left:1px dashed " + color + ";z-index:-1;pointer-events:none")
                appendChild cell axis
            | None -> ()
        | None -> ()
        match icon with
        | Some draw ->
            // The icon as a character in the colour of the real one (arrows, flags, marks, lights).
            let glyph, color = Rbt.Spreadsheet.ConditionalFormatting.iconGlyph draw.Family draw.Index
            let mark = createElement "span"
            setTextContent mark glyph
            // (At the cell's left, whatever the alignment of its value, as Excel draws it.)
            setCssText mark ("float:left;color:" + color + ";margin-right:3px;font-size:10px")
            setAttribute mark "data-cf-icon" (string draw.Index)
            setAttribute mark "data-cf-icon-set" draw.Family
            prependChild cell mark
        | None -> ()
        // A marker and a tooltip for a comment or a comment thread.
        let thread = if Map.isEmpty model.Threads then None else Map.tryFind addr model.Threads
        let marker (color: string) (tooltip: string) =
            let indicator = createElement "div"
            setCssText indicator ("position:absolute;top:0;right:0;width:0;height:0;border-style:solid;border-width:0 6px 6px 0;border-color:transparent " + color + " transparent transparent;pointer-events:none")
            appendChild cell indicator
            // The comment is shown in a box beside the cell while the pointer
            // is over it (showCommentBox); the cell says its text.
            setAttribute cell "data-comment" tooltip
            node.HasNote <- true
            if node.HasTitle then
                removeAttribute cell "title"
                node.HasTitle <- false
        // A linked cell says where it leads (its screen tip, else the
        // target) unless it has a comment to show; in a viewer the pointer
        // shows that a click follows it.
        let link = if model.Hyperlinks.IsEmpty then None else SheetOps.hyperlinkAt model addr
        (match link with
         | Some link ->
            setAttribute cell "data-link" (if link.IsInternal then "internal" else "external")
            if readOnly then setStyle cell "cursor" "pointer"
            node.HasLink <- true
         | None ->
            if node.HasLink then
                removeAttribute cell "data-link"
                node.HasLink <- false)
        match thread, cellData with
        | Some thread, _ -> marker "#7b61ff" (threadText thread)
        | None, Some { Comment = Some comment } ->
            // (A note whose text begins with its author's name - as Excel writes notes - says the name once.)
            marker "#ff0000" (match comment.Author with
                              | Some author when author <> "" && not (comment.Text.StartsWith (author + ":")) -> author + ":\n" + comment.Text
                              | _ -> comment.Text)
        | _ ->
            if node.HasNote then
                removeAttribute cell "data-comment"
                node.HasNote <- false
            match link with
            | Some link ->
                setAttribute cell "title" (link.Tooltip |> Option.defaultValue link.Target)
                node.HasTitle <- true
            | None ->
                if node.HasTitle then
                    removeAttribute cell "title"
                    node.HasTitle <- false
        node.Pane <- pane
        node.Data <- (match storedData with Some data -> box data | None -> null)
        let near (data: CellData option) = match data with Some data -> box data | None -> null
        node.NearAbove <- near aboveRaw
        node.NearLeft <- near leftRaw
        node.NearCorner <- near cornerRaw
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
        // (A host that listens to the filter buttons has its own filter menu:
        // the button beside the column letter, which opens the grid's, goes.)
        let hosted = filterButtonListeners.Count > 0
        let signature = string col + "|" + string x + "|" + string w + "|" + sort + "|" + string filtered + "|" + string readOnly + "|" + string hosted + "|" + string gutterHeight
        if node.Signature <> signature then
            node.Signature <- signature
            node.Index <- col
            setAttribute node.Element "data-col" (string col)
            setCssText node.Element ("left:" + string (int (x + 0.5)) + "px;width:" + string (int (x + w + 0.5) - int (x + 0.5)) + "px;top:" + string gutterHeight + "px")
            setTextContent node.Label (columnToLetter col + sort)
            setStyle node.Label "cursor" (if readOnly then "default" else "pointer")
            match node.Button with
            | Some button ->
                // (A column too narrow for its letter and the button keeps the letter, unless it is filtered.)
                let narrow = w < 44.0 && not filtered
                setCssText button ("opacity:" + (if filtered then "1.0" else "0.5") + ";display:" + (if readOnly || hosted || narrow then "none" else "flex"))
            | None -> ()
            setStyle node.Handle "display" (if readOnly then "none" else "block")

    and drawRowHeader (node: HeaderNode) (row: int) (y: float) (h: float) =
        let signature = string row + "|" + string y + "|" + string h + "|" + string readOnly + "|" + string gutterWidth
        if node.Signature <> signature then
            node.Signature <- signature
            node.Index <- row
            setAttribute node.Element "data-row" (string row)
            setCssText node.Element ("top:" + string (int (y + 0.5)) + "px;height:" + string (int (y + h + 0.5) - int (y + 0.5)) + "px;left:" + string gutterWidth + "px")
            setTextContent node.Label (string (row + 1))
            setStyle node.Handle "display" (if readOnly then "none" else "block")

    /// Render the grid: bring what is on display in line with the model and
    /// the scroll position. Nodes of cells that stay in view and whose data
    /// did not change are left alone; nodes of cells that left the view are
    /// reused for the ones that came in.
    and render () =
        if not (isDeferredSheet model) then giveRoom ()
        match contentElement, layers with
        | Some content, Some layer ->
            if not model.ConditionalFormats.IsEmpty then refreshRuleWorkbook()
            syncExtent()
            match gridContainer with
            | Some grid ->
                originX <- if extentWidth > maxLogicalExtent () then scrollLeft - getScrollLeft grid / appliedZoom else 0.0
                originY <- if extentHeight > maxLogicalExtent () then scrollTop - getScrollTop grid / appliedZoom else 0.0
            | None -> originX <- 0.0; originY <- 0.0
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
            let styleInputs : obj array =
                [| box model.ConditionalFormats; box model.Threads; box model.Hyperlinks; box model.MergedCells; box model.ShowGridLines; box model.Tables
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
                    let pane = (if r < frozenRows then 1 else 0) + (if c < frozenCols then 2 else 0)
                    let x = paneX pane (rowHeaderWidth + colPrefix.[c])
                    let y = paneY pane (headerHeight + rowPrefix.[r])
                    match Portable.Dict.tryFind key cellNodes with
                    | ValueSome node ->
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
                                    drawCell layer node r c x y w h pane data
                    | ValueNone ->
                        let node =
                            if freeCellNodes.Count > 0 then
                                let last = freeCellNodes.[freeCellNodes.Count - 1]
                                Portable.Stack.dropLast freeCellNodes
                                last
                            else
                                let element = createElement "div"
                                setClassName element "spreadsheet-cell"
                                { Element = element; Key = key; Pane = -1; Data = null; NearAbove = null; NearLeft = null; NearCorner = null; X = 0.0; Y = 0.0; W = 0.0; H = 0.0; Epoch = -1; Seen = 0; HasTitle = false; HasNote = false; HasLink = false; Wide = false }
                        node.Key <- key
                        node.Seen <- generation
                        cellNodes.[key] <- node
                        drawCell layer node r c x y w h pane (cellAt r c)
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
                            match Portable.Dict.tryFind row overflowSources with
                            | ValueSome known -> known
                            | ValueNone ->
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
                             (create: unit -> HeaderNode) (draw: HeaderNode -> int -> unit) (frozen: int) (first: int) (last: int) (prefix: SpreadsheetLayout.AxisOffsets) =
                let place (index: int) =
                    if prefix.[index + 1] > prefix.[index] then
                        let node =
                            match Portable.Dict.tryFind index nodes with
                            | ValueSome node -> node
                            | ValueNone ->
                                let node =
                                    if free.Count > 0 then
                                        let last = free.[free.Count - 1]
                                        Portable.Stack.dropLast free
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
                (fun node col -> drawColumnHeader node col (paneX (if col < frozenCols then 2 else 0) (rowHeaderWidth + colPrefix.[col])) (colPrefix.[col + 1] - colPrefix.[col]))
                frozenCols firstCol lastCol colPrefix
            placeHeaders rowHeaderNodes freeRowHeaders layer.RowHeaders newRowHeader
                (fun node row -> drawRowHeader node row (paneY (if row < frozenRows then 1 else 0) (headerHeight + rowPrefix.[row])) (rowPrefix.[row + 1] - rowPrefix.[row]))
                frozenRows firstRow lastRow rowPrefix
            renderOutline layer metrics frozenRows frozenCols firstRow lastRow firstCol lastCol
            renderDrawings content metrics
            renderOverlays layer metrics frozenRows frozenCols
            renderPages layer metrics
            notifyChanges ()
        | _ ->
            // Not mounted: nothing to draw, but the revision still follows the content.
            observeRevision ()

    /// Observe content not already registered as an undoable transaction.
    /// View changes and deferred-sheet conversion never create revisions.
    and observeRevision () =
        let current = Array.init sheets.Length (fun i -> if i = activeSheetIndex then model else sheets.[i])
        let snapshot = currentWorkbookSnapshot()
        if absorbViewState then
            absorbViewState <- false
            session.AbsorbView(snapshot)
            if current.Length = notifiedSheets.Length then notifiedSheets <- Array.copy current
        else
            session.Observe(snapshot) |> ignore
        revisionSheets <- current

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
            let selectionChanged = selection <> notifiedSelection || selectedPicture <> notifiedPicture
            notifiedPicture <- selectedPicture
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
                        "Revision" ==> session.Revision
                        "Modified" ==> session.IsModified
                    ]
                for listener in changeListeners.ToArray() do
                    if changeListeners.Contains listener then (try listener change with error -> reportError "OnChange" (box error))
            if selectionChanged then
                let active = activeCellAddress()
                let change =
                    createObj [
                        "SheetIndex" ==> activeSheetIndex
                        "Selection" ==> selectionRefText ()
                        "ActiveCell" ==> (columnToLetter active.Col + string (active.Row + 1))
                        "ActiveRow" ==> active.Row
                        "ActiveCol" ==> active.Col
                        "Picture" ==> (match selectedPicture with Some id -> box id | None -> null)
                    ]
                // (Worked out when a subscriber reads it: see GetStatus.)
                defineLazy change "Status" selectionStatusObj
                for listener in selectionListeners.ToArray() do
                    if selectionListeners.Contains listener then (try listener change with error -> reportError "OnSelectionChange" (box error))

    /// The selection as an A1 reference ("B2" or "B2:D5").
    and selectionRefText () =
        match model.Selection with
        | SingleCell (row, col) -> sprintf "%s%d" (columnToLetter col) (row + 1)
        | Range (startRow, startCol, endRow, endCol) ->
            sprintf "%s%d:%s%d" (columnToLetter startCol) (startRow + 1) (columnToLetter endCol) (endRow + 1)

    /// Pictures, and labelled frames for charts and other kept objects, over
    /// the cells at their anchors. They are part of the scrolled content, so
    /// they move with the sheet; clicks go through to the cells under them.
    /// With frozen panes a drawing is cut along the frozen edges, as in
    /// Excel: the part over the frozen rows / columns sits in their pane and
    /// stays, the part beside them scrolls (and goes under the frozen cells).
    /// Rebuilt when the drawings, the sizes of rows / columns, their chart
    /// bindings or the panes change (and while a picture is dragged across panes).
    and renderDrawings (content: Element) (metrics: SpreadsheetLayout.SheetMetrics) =
        ensureChartBindings ()
        let frozenRows = max 0 (min metrics.RowCount model.FrozenRows)
        let frozenCols = max 0 (min metrics.ColCount model.FrozenCols)
        let panes = string frozenRows + "|" + string frozenCols + "|" + string rowHeaderWidth + "|" + string headerHeight
        // (A picture that is dragged over frozen panes is cut anew as it moves.)
        let dragged = match pictureDrag with Some drag when frozenRows > 0 || frozenCols > 0 -> Some drag | _ -> None
        if renderedChartBindingVersion <> chartBindingVersion
           || not (obj.ReferenceEquals(renderedDrawings, model.Drawings)) || not (obj.ReferenceEquals(renderedMetrics, metrics)) || renderedDrawingsReadOnly <> readOnly
           || renderedDrawingsPanes <> panes || dragged.IsSome
           || (drawingNodes.Count > 0 && (abs (originX - drawingOriginX) > 250000.0 || abs (originY - drawingOriginY) > 250000.0)) then
            renderedDrawings <- (if dragged.IsSome then null else box model.Drawings)
            renderedMetrics <- box metrics
            renderedDrawingsReadOnly <- readOnly
            renderedDrawingsPanes <- panes
            drawingOriginX <- originX
            drawingOriginY <- originY
            drawingShiftX <- originX
            drawingShiftY <- originY
            renderedChartBindingVersion <- chartBindingVersion
            for node in drawingNodes do
                clearImageSources node
                detach node
            drawingNodes.Clear()
            chartFrames.Clear()
            pendingObjects.Clear()
            chartStamp <- cellsStamp ()
            let frozenW = metrics.ColPrefix.[frozenCols]
            let frozenH = metrics.RowPrefix.[frozenRows]
            let paneElements = match layers with Some layer -> layer.Panes | None -> [| content; content; content; content |]
            for drawing in model.Drawings do
                let rect =
                    match dragged with
                    | Some drag when drag.Id = drawing.Id -> drag.Current
                    | _ -> SpreadsheetLayout.drawingRect model drawing
                if rect.Width >= 1.0 && rect.Height >= 1.0 then
                    let label = SpreadsheetLayout.drawingLabel drawing
                    // (A shape with a macro assigned - its OnAction - runs it when clicked, once macros may run.)
                    let assigned = assignedMacro drawing
                    let takesPointer = assigned.IsSome || (match drawing.Content with Picture _ -> not readOnly | KeptObject _ -> false)
                    // A chart, a shape or a group of them is drawn from the file's parts (None: a frame that says what it is).
                    // (Laid out after the grid is on screen, the objects in view first: until then its place is empty.)
                    let known =
                        match drawing.Content with
                        | KeptObject (_, xml) -> drawingSvgWhen false drawing xml rect.Width rect.Height
                        | Picture _ -> Some None
                    let pending = known.IsNone
                    if pending then pendingObjects.Add drawing.Id |> ignore
                    let drawn = known |> Option.bind id
                    // The drawing's node at (left, top) of its parent.
                    let build (left: float) (top: float) : Element =
                        let node =
                            match drawing.Content with
                            | KeptObject (kind, _) when drawn.IsSome ->
                                let svg, reads = drawn.Value
                                let frame = createElement "div"
                                setClassName frame ("spreadsheet-drawing spreadsheet-object spreadsheet-object-" + kind + " spreadsheet-object-drawn")
                                setStyles frame [ "overflow", "hidden"; "box-sizing", "border-box" ]
                                setInnerHtml frame svg
                                chartFrames.Add ((frame, drawing, rect.Width, rect.Height, reads))
                                frame
                            | Picture (contentType, _, data) ->
                                let image = createElement "img"
                                setAttribute image "src" (pictureUrl data contentType)
                                setAttribute image "alt" (drawing.Description |> Option.defaultValue label)
                                setAttribute image "draggable" "false"
                                setClassName image "spreadsheet-drawing spreadsheet-picture"
                                image
                            | KeptObject (kind, _) when pending ->
                                let frame = createElement "div"
                                setClassName frame ("spreadsheet-drawing spreadsheet-object spreadsheet-object-" + kind + " spreadsheet-object-pending")
                                frame
                            | KeptObject (kind, _) ->
                                let frame = createElement "div"
                                setClassName frame ("spreadsheet-drawing spreadsheet-object spreadsheet-object-" + kind)
                                // The label is one line inside the frame: cut with an ellipsis where the
                                // frame is narrower than it, in smaller letters in a low frame, and
                                // left out of a frame too small to say anything (the title still says it).
                                let caption = createElement "span"
                                setClassName caption "spreadsheet-object-label"
                                setTextContent caption label
                                let fontPx = if rect.Height >= 18.0 then 12 elif rect.Height >= 13.0 then 9 else 0
                                setStyles caption [
                                    "display", (if fontPx = 0 || rect.Width < 24.0 then "none" else "block")
                                    "min-width", "0"
                                    "max-width", "100%"
                                    "padding", "0 4px"
                                    "box-sizing", "border-box"
                                    "white-space", "nowrap"
                                    "overflow", "hidden"
                                    "text-overflow", "ellipsis"
                                    "line-height", "1.2"
                                ]
                                appendChild frame caption
                                setStyles frame [
                                    "display", "flex"
                                    "align-items", "center"
                                    "justify-content", "center"
                                    "box-sizing", "border-box"
                                    "background", "#f6f8fa"
                                    "border", "1px solid #8a94a6"
                                    "color", "#4a5568"
                                    "font", (string (max 9 fontPx) + "px Arial, sans-serif")
                                    "overflow", "hidden"
                                ]
                                frame
                        setAttribute node "title" label
                        setAttribute node "data-drawing-id" (string drawing.Id)
                        assigned |> Option.iter (setAttribute node "data-macro")
                        // A picture takes the pointer in the editor (it is selected,
                        // moved and sized with it); in a viewer, and for a chart or
                        // another kept object, clicks go through to the cells.
                        setStyles node [
                            "position", "absolute"
                            "left", sprintf "%.1fpx" left
                            "top", sprintf "%.1fpx" top
                            "width", sprintf "%.1fpx" rect.Width
                            "height", sprintf "%.1fpx" rect.Height
                            "max-width", "none"
                            "z-index", "3"
                            "pointer-events", (if takesPointer then "auto" else "none")
                            "cursor", (if assigned.IsSome then "pointer" elif takesPointer then "move" else "default")
                        ]
                        node
                    let right = rect.X + rect.Width
                    let bottom = rect.Y + rect.Height
                    if (frozenRows = 0 && frozenCols = 0) || (rect.X >= frozenW && rect.Y >= frozenH) then
                        let node = build (paneX 0 (rowHeaderWidth + rect.X)) (paneY 0 (headerHeight + rect.Y))
                        setAttribute node "data-pane" "0"
                        appendChild content node
                        drawingNodes.Add node
                    else
                        // One part per pane the drawing reaches: a box that cuts
                        // the drawing to the pane's share of it.
                        for pane in 0 .. 3 do
                            let x0, x1 = if pane >= 2 then rect.X, min right frozenW else max rect.X frozenW, right
                            let y0, y1 = if pane = 1 || pane = 3 then rect.Y, min bottom frozenH else max rect.Y frozenH, bottom
                            if x1 - x0 > 0.0 && y1 - y0 > 0.0 then
                                let part = createElement "div"
                                setClassName part "spreadsheet-drawing-part"
                                setAttribute part "data-drawing-id" (string drawing.Id)
                                setAttribute part "data-pane" (string pane)
                                setCssText part
                                    (sprintf "position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;overflow:hidden;z-index:3;pointer-events:none"
                                        (paneX pane (rowHeaderWidth + x0)) (paneY pane (headerHeight + y0)) (x1 - x0) (y1 - y0))
                                appendChild part (build (rect.X - x0) (rect.Y - y0))
                                appendChild paneElements.[pane] part
                                drawingNodes.Add part
        elif chartFrames.Count > 0 then
            // The same drawings at the same places: a chart whose cells changed is drawn again.
            let stamp = cellsStamp ()
            if stamp.Length <> chartStamp.Length || not (Array.forall2 (fun (a: obj) (b: obj) -> obj.ReferenceEquals(a, b)) stamp chartStamp) then
                let changed =
                    [| for index in 0 .. stamp.Length - 1 do
                        if index >= chartStamp.Length || not (obj.ReferenceEquals(stamp.[index], chartStamp.[index])) then
                            if index < sheetNames.Length then yield sheetNames.[index].ToUpperInvariant() |]
                chartStamp <- stamp
                for at in 0 .. chartFrames.Count - 1 do
                    let frame, drawing, width, height, reads = chartFrames.[at]
                    // (A sheet it reads changed: laid out again when the cells it reads did - an edit
                    // elsewhere on the sheet, or one that left them as they were, is none of its business.)
                    if reads |> Array.exists (fun ((name, _, _, _, _, _) as read) -> Array.contains name changed && not (chartBlockStands read)) then
                        match drawing.Content with
                        | KeptObject (_, xml) ->
                            (match drawingSvg drawing xml width height with
                             | Some (svg, readsNow) ->
                                chartRedraws <- chartRedraws + 1
                                setInnerHtml frame svg
                                chartFrames.[at] <- (frame, drawing, width, height, readsNow)
                             | None -> ())
                        | Picture _ -> ()
        // The objects that are not laid out yet: those in view, once the grid is on screen - a
        // few at a time, so that a sheet full of charts does not hold up its own display.
        if pendingObjects.Count > 0 && not objectsScheduled then
            let inView (drawing: SheetDrawing) =
                let rect = SpreadsheetLayout.drawingRect model drawing
                rect.Y + rect.Height >= scrollTop - 300.0 && rect.Y <= scrollTop + containerHeight + 300.0
                && rect.X + rect.Width >= scrollLeft - 300.0 && rect.X <= scrollLeft + containerWidth + 300.0
            if model.Drawings |> List.exists (fun drawing -> pendingObjects.Contains drawing.Id && inView drawing) then
                objectsScheduled <- true
                setTimeoutMs (fun () ->
                    objectsScheduled <- false
                    if contentElement.IsSome && layers.IsSome then
                        let started = nowMs ()
                        let mutable laidOut = 0
                        for drawing in model.Drawings do
                            if pendingObjects.Contains drawing.Id && inView drawing && (laidOut = 0 || nowMs () - started < 12.0) then
                                match drawing.Content with
                                | KeptObject (_, xml) ->
                                    let rect = SpreadsheetLayout.drawingRect model drawing
                                    drawingSvg drawing xml rect.Width rect.Height |> ignore
                                    laidOut <- laidOut + 1
                                | Picture _ -> ()
                        // (Drawn from what is known now; what is left in view follows in the next slice.)
                        if laidOut > 0 then
                            renderedDrawings <- null
                            render ()) 0 |> ignore
        if drawingNodes.Count > 0 && (drawingShiftX <> originX || drawingShiftY <> originY) then
            drawingShiftX <- originX
            drawingShiftY <- originY
            let dx = drawingOriginX - originX
            let dy = drawingOriginY - originY
            for node in drawingNodes do
                let pane = match getAttribute node "data-pane" with "1" -> 1 | "2" -> 2 | "3" -> 3 | _ -> 0
                let x = if pane >= 2 then 0.0 else dx
                let y = if pane = 1 || pane = 3 then 0.0 else dy
                setStyle node "transform" (sprintf "translate(%.2fpx,%.2fpx)" x y)
        renderPictureFrame content

    /// The page break preview (`viewMode`): every page the sheet prints on (`paginatePrinted`,
    /// what the PDF export lays out and `GetPrintPages` reports) that is in view is a frame over
    /// its cells with its number in the middle. The edge of a frame is solid where the print
    /// area begins or ends and at a manual break, dashed at a break the paper makes. The frames
    /// take no click: the grid under them is edited as in the normal view.
    and renderPages (layer: GridLayers) (metrics: SpreadsheetLayout.SheetMetrics) =
        for node in pageNodes do detach node
        pageNodes.Clear()
        if viewMode = "pageBreak" && not sheetLocked && not (isDeferredSheet model) then
            // (Worked out again when what the pages depend on changed: cells, sizes, merges, the page setup.)
            let key : obj array = [| box model.Cells; box model.PageSetup; box model.RowHeights; box model.ColumnWidths; box model.MergedCells; box model.Drawings |]
            let pagination =
                match pagesKnown with
                | Some (known, rowDefault, colDefault, pages)
                    when rowDefault = model.DefaultRowHeight && colDefault = model.DefaultColumnWidth
                         && known.Length = key.Length && Array.forall2 (fun (a: obj) (b: obj) -> obj.ReferenceEquals(a, b)) known key -> pages
                | _ ->
                    let pages = Rbt.Spreadsheet.Pdf.Pagination.paginatePrinted model
                    pagesKnown <- Some (key, model.DefaultRowHeight, model.DefaultColumnWidth, pages)
                    pages
            let setup = model.PageSetup
            let offset = if setup.UseFirstPageNumber then setup.FirstPageNumber - 1 else 0
            let rowPrefix = metrics.RowPrefix
            let colPrefix = metrics.ColPrefix
            let areas = List.toArray pagination.PrintAreas
            let viewTop, viewBottom = scrollTop - 100.0, scrollTop + containerHeight + 100.0
            let viewLeft, viewRight = scrollLeft - 100.0, scrollLeft + containerWidth + 100.0
            pagination.Pages
            |> List.iteri (fun at page ->
                let firstRow, lastRow = max 0 page.RowStart, min (metrics.RowCount - 1) page.RowEnd
                let firstCol, lastCol = max 0 page.ColStart, min (metrics.ColCount - 1) page.ColEnd
                if lastRow >= firstRow && lastCol >= firstCol then
                    let x, y = colPrefix.[firstCol], rowPrefix.[firstRow]
                    let width, height = colPrefix.[lastCol + 1] - x, rowPrefix.[lastRow + 1] - y
                    if width > 0.0 && height > 0.0 && y + height >= viewTop && y <= viewBottom && x + width >= viewLeft && x <= viewRight then
                        let area = if page.Area >= 0 && page.Area < areas.Length then Some areas.[page.Area] else None
                        let manualTop = List.contains page.RowStart setup.ManualRowBreaks
                        let manualLeft = List.contains page.ColStart setup.ManualColBreaks
                        let atAreaTop = (match area with Some range -> page.RowStart <= range.From.Row | None -> true)
                        let atAreaLeft = (match area with Some range -> page.ColStart <= range.From.Col | None -> true)
                        let atAreaBottom = (match area with Some range -> page.RowEnd >= range.To.Row | None -> true)
                        let atAreaRight = (match area with Some range -> page.ColEnd >= range.To.Col | None -> true)
                        let line (solid: bool) = if solid then "2px solid #1a73e8" else "2px dashed #1a73e8"
                        let frame = createElement "div"
                        setClassName frame "spreadsheet-page"
                        setAttribute frame "data-page" (string (at + 1 + offset))
                        setAttribute frame "data-range" (SheetOps.rangeText { From = { Row = page.RowStart; Col = page.ColStart }; To = { Row = page.RowEnd; Col = page.ColEnd } })
                        setAttribute frame "data-first-row" (string page.RowStart)
                        setAttribute frame "data-first-col" (string page.ColStart)
                        setAttribute frame "data-break-top" (if atAreaTop then "area" elif manualTop then "manual" else "automatic")
                        setAttribute frame "data-break-left" (if atAreaLeft then "area" elif manualLeft then "manual" else "automatic")
                        setCssText frame
                            (sprintf "position:absolute;left:%.1fpx;top:%.1fpx;width:%.1fpx;height:%.1fpx;box-sizing:border-box;pointer-events:none;z-index:3;display:flex;align-items:center;justify-content:center;overflow:hidden;border-top:%s;border-left:%s;border-bottom:%s;border-right:%s"
                                (paneX 0 (rowHeaderWidth + x)) (paneY 0 (headerHeight + y)) width height
                                (line (atAreaTop || manualTop)) (line (atAreaLeft || manualLeft))
                                (if atAreaBottom then line true else "0") (if atAreaRight then line true else "0"))
                        let label = createElement "span"
                        setClassName label "spreadsheet-page-label"
                        setTextContent label ("Page " + string (at + 1 + offset))
                        setCssText label (sprintf "font:bold %dpx Arial, sans-serif;color:rgba(60,64,67,0.28);white-space:nowrap;user-select:none" (if width < 160.0 || height < 60.0 then 14 else 32))
                        appendChild frame label
                        appendChild layer.Panes.[0] frame
                        pageNodes.Add frame)

    /// The frame around the selected picture, with its eight handles. While
    /// the picture is dragged, the frame and the picture are where the drag
    /// has them (the model changes when the mouse is released).
    and renderPictureFrame (content: Element) =
        let selected =
            match selectedPicture with
            | Some id ->
                (match SheetOps.drawingById id model with
                 | Some drawing when (match drawing.Content with Picture _ -> true | KeptObject _ -> false) -> Some drawing
                 | _ ->
                    // (Deleted, or undone away.)
                    selectedPicture <- None
                    None)
            | None -> None
        match selected with
        | None ->
            (match pictureFrame with Some frame -> setStyle frame "display" "none" | None -> ())
        | Some drawing ->
            let rect =
                match pictureDrag with
                | Some drag when drag.Id = drawing.Id -> drag.Current
                | _ -> SpreadsheetLayout.drawingRect model drawing
            let metrics = SpreadsheetLayout.metricsOf model
            let pane =
                (if rect.Y < metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)] then 1 else 0)
                + (if rect.X < metrics.ColPrefix.[max 0 (min metrics.ColCount model.FrozenCols)] then 2 else 0)
            // The frame goes with the pane the picture's corner is in (a
            // picture over frozen rows / columns stays there with its handles).
            let parent =
                match layers with
                | Some layer when model.FrozenRows > 0 || model.FrozenCols > 0 ->
                    let metrics = SpreadsheetLayout.metricsOf model
                    let frozenW = metrics.ColPrefix.[max 0 (min metrics.ColCount model.FrozenCols)]
                    let frozenH = metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)]
                    layer.Panes.[(if rect.Y < frozenH then 1 else 0) + (if rect.X < frozenW then 2 else 0)]
                | _ -> content
            let frame =
                match pictureFrame with
                | Some frame when not (isNull (box (parentOf frame))) ->
                    if not (isChildOf frame parent) then appendChild parent frame
                    frame
                | _ ->
                    let frame = createElement "div"
                    setClassName frame "spreadsheet-picture-frame"
                    setCssText frame "position:absolute;box-sizing:border-box;border:1px solid #1a73e8;z-index:4;pointer-events:none"
                    for handle, left, top, cursor in [ "nw", "0%", "0%", "nwse-resize"; "n", "50%", "0%", "ns-resize"; "ne", "100%", "0%", "nesw-resize"; "e", "100%", "50%", "ew-resize"
                                                       "se", "100%", "100%", "nwse-resize"; "s", "50%", "100%", "ns-resize"; "sw", "0%", "100%", "nesw-resize"; "w", "0%", "50%", "ew-resize" ] do
                        let grip = createElement "div"
                        setClassName grip "spreadsheet-picture-handle"
                        setAttribute grip "data-handle" handle
                        setCssText grip (sprintf "position:absolute;left:%s;top:%s;width:8px;height:8px;margin:-5px 0 0 -5px;box-sizing:border-box;background:#fff;border:1px solid #1a73e8;cursor:%s;pointer-events:auto" left top cursor)
                        appendChild frame grip
                    appendChild parent frame
                    pictureFrame <- Some frame
                    frame
            setAttribute frame "data-drawing-id" (string drawing.Id)
            setStyles frame [
                "display", "block"
                "left", sprintf "%.1fpx" (paneX pane (rowHeaderWidth + rect.X))
                "top", sprintf "%.1fpx" (paneY pane (headerHeight + rect.Y))
                "width", sprintf "%.1fpx" rect.Width
                "height", sprintf "%.1fpx" rect.Height
            ]
            match pictureDrag with
            | Some drag when drag.Id = drawing.Id && model.FrozenRows <= 0 && model.FrozenCols <= 0 ->
                let image = queryOne content (sprintf ".spreadsheet-picture[data-drawing-id=\"%d\"]" drawing.Id)
                if not (isNull (box image)) then
                    setStyles image [
                        "left", sprintf "%.1fpx" (paneX 0 (rowHeaderWidth + rect.X))
                        "top", sprintf "%.1fpx" (paneY 0 (headerHeight + rect.Y))
                        "transform", ""
                        "width", sprintf "%.1fpx" rect.Width
                        "height", sprintf "%.1fpx" rect.Height
                    ]
            | _ -> ()

    /// The outline strips (see `showOutline`): a background, a bracket and a
    /// +/- button per group in view, the level buttons in the corner. Rows'
    /// groups sit in the strip that scrolls with the row numbers, groups of
    /// frozen rows / columns with the frozen headers.
    and renderOutline (layer: GridLayers) (metrics: SpreadsheetLayout.SheetMetrics) (frozenRows: int) (frozenCols: int)
                      (firstRow: int) (lastRow: int) (firstCol: int) (lastCol: int) =
        if gutterWidth = 0.0 && gutterHeight = 0.0 then
            if outlineNodes.Count > 0 then
                for node in outlineNodes do detach node
                outlineNodes.Clear()
                outlineSignature <- [||]
        else
            let signature : obj array =
                [| outlineAxes; box metrics.RowPrefix; box metrics.ColPrefix; box frozenRows; box frozenCols; box firstRow; box lastRow; box firstCol; box lastCol
                   box gutterWidth; box gutterHeight; box readOnly; box headBand; box sideBand; box originX; box originY |]
            let same =
                signature.Length = outlineSignature.Length
                && Array.forall2 (fun (a: obj) (b: obj) -> obj.ReferenceEquals(a, b) || (not (isNull a) && a.Equals b)) signature outlineSignature
            if not same then
                outlineSignature <- signature
                let px (value: float) = string (System.Math.Round value) + "px"
                let mutable used = 0
                let put (parent: Element) (className: string) (css: string) : Element =
                    let node =
                        if used < outlineNodes.Count then outlineNodes.[used]
                        else
                            let node = createElement "div"
                            outlineNodes.Add node
                            node
                    used <- used + 1
                    if not (isChildOf node parent) then appendChild parent node
                    setClassName node ("spreadsheet-outline " + className)
                    setCssText node css
                    // (A node of the pool was something else in the last drawing - a
                    // level button with its digit, a +/- button: nothing of that stays.)
                    setTextContent node ""
                    for name in [ "data-axis"; "data-first"; "data-last"; "data-level"; "data-collapsed"; "title" ] do removeAttribute node name
                    node
                let stripBackground = "background:#f3f3f3;"
                let cursor = ""
                let axis (rows: bool) =
                    let prefix = if rows then metrics.RowPrefix else metrics.ColPrefix
                    let count = if rows then metrics.RowCount else metrics.ColCount
                    let frozen = if rows then frozenRows else frozenCols
                    let first, last = if rows then firstRow, lastRow else firstCol, lastCol
                    let groups = if rows then outlineRowGroups else outlineColGroups
                    let depth = if rows then outlineRowDepth else outlineColDepth
                    let thickness = if rows then gutterWidth else gutterHeight
                    // Where the sheet's rows / columns begin on this axis, and the other axis' header band.
                    let origin = if rows then headerHeight else rowHeaderWidth
                    let band = if rows then gutterHeight else gutterWidth
                    let bandSize = if rows then headBand else sideBand
                    let scrolled = if rows then layer.RowHeaders else layer.ColHeaders
                    let shift = if rows then originY else originX
                    // A box along the axis (`along`, `length`) and across the strip (`across`, `size`).
                    let rect (along: float) (length: float) (across: float) (size: float) =
                        if rows then "left:" + px across + ";top:" + px along + ";width:" + px size + ";height:" + px length + ";"
                        else "top:" + px across + ";left:" + px along + ";height:" + px size + ";width:" + px length + ";"
                    let edge = if rows then "border-right:1px solid #a6a6a6;" else "border-bottom:1px solid #a6a6a6;"
                    if thickness > 0.0 then
                        put scrolled "spreadsheet-outline-strip" (rect 0.0 (min (origin + prefix.[count]) (maxLogicalExtent ())) 0.0 thickness + stripBackground + edge + "z-index:1") |> ignore
                        // (Groups that scroll pass under the strip of the frozen rows / columns.)
                        if frozen > 0 then
                            put layer.FrozenHeaders "spreadsheet-outline-strip" (rect origin prefix.[frozen] 0.0 thickness + stripBackground + edge + "z-index:4") |> ignore
                        let after = SheetOps.summaryAfter rows model
                        for (groupFirst, groupLast, level, collapsed) in groups do
                            let inView = groupFirst < frozen || (groupLast >= first - 1 && groupFirst <= last + 1)
                            if inView && used < 1500 && groupLast < count then
                                let parent = if groupLast < frozen then layer.FrozenHeaders else scrolled
                                let across = float (level - 1) * outlineStep + 2.0
                                let start = origin + prefix.[groupFirst] - (if groupLast < frozen then 0.0 else shift)
                                let length = prefix.[groupLast + 1] - prefix.[groupFirst]
                                if length > 3.0 then
                                    // The bracket: a line along the group with a tick at its far end.
                                    let tick =
                                        if rows then (if after then "border-top:1px solid #666;" else "border-bottom:1px solid #666;") + "border-left:1px solid #666;"
                                        else (if after then "border-left:1px solid #666;" else "border-right:1px solid #666;") + "border-top:1px solid #666;"
                                    put parent "spreadsheet-outline-line" (rect (start + 1.0) (length - 2.0) (across + 6.0) 5.0 + tick + "pointer-events:none;z-index:5") |> ignore
                                let summary = SheetOps.summaryIndex rows groupFirst groupLast model
                                if summary >= 0 && summary < count && prefix.[summary + 1] > prefix.[summary] then
                                    let size = prefix.[summary + 1] - prefix.[summary]
                                    let buttonParent = if summary < frozen then layer.FrozenHeaders else scrolled
                                    let button =
                                        put buttonParent "spreadsheet-outline-button"
                                            ((if rows then "left:" + px across + ";top:" + px (origin + prefix.[summary] - (if summary < frozen then 0.0 else shift) + (size - 13.0) / 2.0) + ";"
                                              else "top:" + px across + ";left:" + px (origin + prefix.[summary] - (if summary < frozen then 0.0 else shift) + (size - 13.0) / 2.0) + ";") + cursor)
                                    setTextContent button (if collapsed then "+" else "\u2212")
                                    setAttribute button "data-axis" (if rows then "rows" else "columns")
                                    setAttribute button "data-first" (string groupFirst)
                                    setAttribute button "data-last" (string groupLast)
                                    setAttribute button "data-level" (string level)
                                    setAttribute button "data-collapsed" (if collapsed then "1" else "0")
                                    setAttribute button "title" (if collapsed then "Expand" else "Collapse")
                        // The level buttons, in the corner beside the other axis' headers.
                        for level in 1 .. depth + 1 do
                            let across = float (level - 1) * outlineStep + 2.0
                            let button =
                                put layer.FrozenHeaders "spreadsheet-outline-level"
                                    ((if rows then "left:" + px across + ";top:" + px (band + (bandSize - 13.0) / 2.0) + ";"
                                      else "top:" + px across + ";left:" + px (band + (bandSize - 13.0) / 2.0) + ";") + cursor)
                            setTextContent button (string level)
                            setAttribute button "data-axis" (if rows then "rows" else "columns")
                            setAttribute button "data-level" (string level)
                            removeAttribute button "data-first"
                            setAttribute button "title" ("Show level " + string level)
                axis true
                axis false
                while outlineNodes.Count > used do
                    detach outlineNodes.[outlineNodes.Count - 1]
                    Portable.Stack.dropLast outlineNodes

    /// A press of a button of the outline strips: a group's +/- or a level button.
    and pressOutline (node: Element) =
        let rows = getAttribute node "data-axis" = "rows"
        let number (name: string) = match System.Int32.TryParse (getAttribute node name) with | true, value -> value | _ -> -1
        // In a viewer what is collapsed is how the sheet is looked at, not a
        // change of the workbook: no undo step, no new revision.
        let apply (description: string) (change: SpreadsheetModel -> SpreadsheetModel) =
            if readOnly then
                // (What happened before is told first; this then counts as the state already known.)
                notifyChanges ()
                let before = axesState model
                model <- change model
                if before <> axesState model then
                    absorbViewState <- true
                    render()
            elif changeAxes description change then render()
        if number "data-first" >= 0 then
            let collapse = getAttribute node "data-collapsed" <> "1"
            apply (if collapse then "Collapse group" else "Expand group") (SheetOps.setGroupCollapsed "outline" rows (number "data-first") collapse)
        elif number "data-level" > 0 then
            apply "Show outline level" (SheetOps.showOutlineLevel rows (number "data-level"))

    /// The header cells that carry a filter button: (row, col, active, sorted)
    /// with sorted 1 ascending, -1 descending, 0 not sorted. The header row of
    /// the sheet's AutoFilter, then the header rows of the tables that have a
    /// filter (a cell is listed once).
    and filterButtonCells () : (int * int * bool * int) list =
        let sortedOf (col: int) =
            match model.SortState with
            | Some state when state.Column = col -> if state.Ascending then 1 else -1
            | _ -> 0
        let ofSheet =
            match model.AutoFilter with
            | Some filter ->
                [ for col in filter.Range.From.Col .. filter.Range.To.Col ->
                    filter.Range.From.Row, col, Map.containsKey col model.Filters, sortedOf col ]
            | None -> []
        let taken = ofSheet |> List.map (fun (row, col, _, _) -> row, col) |> Set.ofList
        let ofTables =
            [ for table in model.Tables do
                match table.AutoFilter with
                | Some (firstCol, lastCol) when table.HeaderRowCount > 0 ->
                    let row = table.StartRow - 1 + table.HeaderRowCount - 1
                    // (The sheet has one filter: a table's button is lit when that filter is over the table's rows.)
                    let filtersTable =
                        match model.AutoFilter with
                        | Some filter -> filter.Range.From.Row <= table.EndRow - 1 && filter.Range.To.Row >= row
                        | None -> false
                    for col in firstCol - 1 .. lastCol - 1 do
                        if not (Set.contains (row, col) taken) then
                            yield row, col, (filtersTable && Map.containsKey col model.Filters), sortedOf col
                | _ -> () ]
        ofSheet @ ofTables

    /// What OnFilterButton subscribers and GetFilterButtonRects give for a button.
    and filterButtonPayload (node: Element) : obj =
        let number (name: string) = match System.Int32.TryParse (getAttribute node name) with | true, value -> value | _ -> -1
        let rect: float array = boundingRect node
        createObj [ "column" ==> number "data-col"; "row" ==> number "data-row"; "sheetIndex" ==> activeSheetIndex
                    "clientRect" ==> createObj [ "left" ==> rect.[0]; "top" ==> rect.[1]; "width" ==> rect.[2]; "height" ==> rect.[3] ]
                    "active" ==> (getAttribute node "data-active" = "1")
                    "sorted" ==> (match getAttribute node "data-sorted" with
                                  | "asc" -> box "asc"
                                  | "desc" -> box "desc"
                                  | _ -> null) ]

    /// The button of a header cell, when it is on display.
    and filterButtonAt (row: int) (col: int) : Element option =
        filterButtonNodes |> Seq.tryFind (fun node -> getAttribute node "data-row" = string row && getAttribute node "data-col" = string col)

    /// A filter button was pressed (or Alt+Down on its cell): subscribers are
    /// told, and the grid's own checklist opens only when there is none.
    and pressFilterButton (node: Element) =
        if filterButtonListeners.Count > 0 then
            let payload = filterButtonPayload node
            for listener in filterButtonListeners.ToArray() do
                if filterButtonListeners.Contains listener then (try listener payload with error -> reportError "OnFilterButton" (box error))
        elif not readOnly then
            match System.Int32.TryParse (getAttribute node "data-col") with
            | true, col -> showFilterDropdown col (pageRight node) (pageBottom node)
            | _ -> ()

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
                let x0, x1 = rowHeaderWidth + colPrefix.[left], rowHeaderWidth + colPrefix.[right + 1]
                let y0, y1 = headerHeight + rowPrefix.[top], headerHeight + rowPrefix.[bottom + 1]
                // Huge selections are clipped outside the viewport before being
                // handed to CSS, whose coordinate and size ranges are finite.
                let x0, x1 =
                    if pane < 2 && extentWidth > maxLogicalExtent () then
                        max x0 (scrollLeft - 1000.0), min x1 (scrollLeft + containerWidth + 1000.0)
                    else x0, x1
                let y0, y1 =
                    if pane <> 1 && pane <> 3 && extentHeight > maxLogicalExtent () then
                        max y0 (scrollTop - 1000.0), min y1 (scrollTop + containerHeight + 1000.0)
                    else y0, y1
                if x1 <= x0 || y1 <= y0 then None
                else Some (paneX pane x0, paneY pane y0, x1 - x0, y1 - y0)
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
            let x = paneX pane (rowHeaderWidth + colPrefix.[addr.Col])
            let y = paneY pane (headerHeight + rowPrefix.[addr.Row])
            let w = colPrefix.[editRight + 1] - colPrefix.[addr.Col]
            let h = rowPrefix.[editBottom + 1] - rowPrefix.[addr.Row]
            // (Typed in the cell's font, at the size its text is drawn.)
            let editFormat = match getCell addr model with Some data -> data.Format | None -> defaultFormat
            setCssText layer.EditInput
                ("position:absolute;left:" + px x + ";top:" + px y + ";width:" + px (max 40.0 w) + ";height:" + px (max 18.0 h)
                 + ";border:2px solid #217346;outline:none;padding:0 3px;margin:0;font-family:" + cellFontFamily editFormat.FontName
                 + (if isCalibri editFormat.FontName then ";font-size-adjust:" + string calibriAspect else "")
                 + ";font-size:" + cssPx (SheetOps.pointsToPx (float (defaultArg editFormat.FontSize 11))) + ";background:white;box-sizing:border-box;z-index:11")
            if getValue layer.EditInput <> model.EditBuffer && not (isActiveElement layer.EditInput) then setValue layer.EditInput model.EditBuffer
        | _ -> hide layer.EditInput
        // The fill handle, at the bottom-right corner of the selection.
        if readOnly || model.EditingCell.IsSome || rowCount = 0 || colCount = 0 || bottom < 0 || right < 0 then hide layer.FillHandle
        else
            place layer.FillHandle (paneOf bottom right)
            let cornerX = paneX (paneOf bottom right) (rowHeaderWidth + colPrefix.[right + 1])
            let cornerY = paneY (paneOf bottom right) (headerHeight + rowPrefix.[bottom + 1])
            setCssText layer.FillHandle
                ("position:absolute;left:" + px (cornerX - 4.0) + ";top:" + px (cornerY - 4.0)
                 + ";width:7px;height:7px;background:#217346;border:1px solid white;box-sizing:border-box;cursor:crosshair;z-index:9")
        // The button that opens the list of a list validation, beside the active cell.
        if readOnly || model.EditingCell.IsSome || rowCount = 0 || colCount = 0 || active.Row >= rowCount || active.Col >= colCount
           || model.Validations.IsEmpty || not (showsListDropdown active model) then hide layer.ListButton
        else
            place layer.ListButton activePane
            let cellRight = paneX activePane (rowHeaderWidth + colPrefix.[min colCount (activeRight + 1)])
            let cellTop = paneY activePane (headerHeight + rowPrefix.[activeTop])
            let cellHeight = rowPrefix.[min rowCount (activeBottom + 1)] - rowPrefix.[activeTop]
            setCssText layer.ListButton
                ("position:absolute;left:" + px cellRight + ";top:" + px (cellTop + max 0.0 (cellHeight - 20.0)) + ";width:17px;height:" + px (min 20.0 (max 12.0 cellHeight))
                 + ";box-sizing:border-box;border:1px solid #8a8a8a;background:#f3f3f3;color:#444;font-size:9px;line-height:" + px (min 18.0 (max 10.0 (cellHeight - 2.0)))
                 + ";text-align:center;cursor:pointer;user-select:none;z-index:9")
            if getAttribute layer.ListButton "data-drawn" <> "1" then
                setAttribute layer.ListButton "data-drawn" "1"
                setTextContent layer.ListButton "\u25BC"
        // The dynamic array the active cell is part of: a thin outline around
        // the range it fills, as in Excel.
        let spilled = if Map.isEmpty model.Spill then None else SheetOps.spillRangeAt active model
        for pane in 0 .. 3 do
            let rect =
                match spilled with
                | Some (_, range) when range.To.Row < rowCount && range.To.Col < colCount -> paneRect pane range.From.Row range.From.Col range.To.Row range.To.Col
                | _ -> None
            match rect with
            | Some (x, y, w, h) when w > 0.0 && h > 0.0 ->
                if isNull (box spillOutlines.[pane]) then
                    let node = createElement "div"
                    setClassName node "spreadsheet-spill-outline"
                    spillOutlines.[pane] <- node
                let node = spillOutlines.[pane]
                place node pane
                setAttribute node "data-range" (match spilled with Some (_, range) -> SheetOps.rangeText range | None -> "")
                setCssText node
                    ("position:absolute;left:" + px x + ";top:" + px y + ";width:" + px (w + 1.0) + ";height:" + px (h + 1.0)
                     + ";box-sizing:border-box;border:1px solid #2b6cd4;pointer-events:none;z-index:4")
            | _ -> if not (isNull (box spillOutlines.[pane])) then hide spillOutlines.[pane]
        // The filter buttons: one in each header cell of the AutoFilter and of
        // a filtered table, in the cell's pane (so a frozen header keeps its
        // buttons and a scrolled one takes them along). A viewer has them only
        // for a host that listens (it decides what to offer).
        let mutable buttons = 0
        if (not readOnly || filterButtonListeners.Count > 0) && (model.AutoFilter.IsSome || not model.Tables.IsEmpty) then
            for (row, col, active, sorted) in filterButtonCells () do
                if row >= 0 && col >= 0 && row < rowCount && col < colCount then
                    let w = colPrefix.[col + 1] - colPrefix.[col]
                    let h = rowPrefix.[row + 1] - rowPrefix.[row]
                    if w > 0.0 && h > 0.0 then
                        let node =
                            if buttons < filterButtonNodes.Count then filterButtonNodes.[buttons]
                            else
                                let node = createElement "div"
                                setClassName node "spreadsheet-filter-button"
                                filterButtonNodes.Add node
                                node
                        buttons <- buttons + 1
                        place node (paneOf row col)
                        let size = max 6.0 (min 16.0 (min (w - 2.0) (h - 2.0)))
                        let signature = string row + "|" + string col + "|" + string active + "|" + string sorted
                        if getAttribute node "data-state" <> signature then
                            setAttribute node "data-state" signature
                            setAttribute node "data-row" (string row)
                            setAttribute node "data-col" (string col)
                            setAttribute node "data-active" (if active then "1" else "0")
                            setAttribute node "data-sorted" (if sorted > 0 then "asc" elif sorted < 0 then "desc" else "")
                            setAttribute node "title" (if active then "Filtered" else "Filter")
                            // An arrow for the sort order, a funnel-like mark for a filter, else the plain triangle.
                            setTextContent node ((if sorted > 0 then "\u2191" elif sorted < 0 then "\u2193" else "") + (if active then "\u25BC\u0332" else "\u25BE"))
                        setCssText node
                            ("position:absolute;left:" + px (paneX (paneOf row col) (rowHeaderWidth + colPrefix.[col + 1] - size - 1.0)) + ";top:" + px (paneY (paneOf row col) (headerHeight + rowPrefix.[row + 1] - size - 1.0))
                             + ";width:" + px size + ";height:" + px size + ";box-sizing:border-box;border:1px solid " + (if active then "#217346" else "#8a8a8a")
                             + ";background:" + (if active then "#d5eadd" else "#f3f3f3") + ";color:" + (if active then "#217346" else "#444")
                             + ";font-size:" + (if sorted <> 0 then "7px" else "9px") + ";line-height:" + px (size - 2.0)
                             + ";text-align:center;cursor:pointer;user-select:none;overflow:hidden;white-space:nowrap;z-index:8")
        while filterButtonNodes.Count > buttons do
            detach filterButtonNodes.[filterButtonNodes.Count - 1]
            Portable.Stack.dropLast filterButtonNodes
        // The circles of CircleInvalid: a red ellipse around each cell, in the cell's pane.
        let circled =
            match Portable.Dict.tryFind activeSheetIndex invalidCircles with
            | ValueSome cells -> cells
            | ValueNone -> [||]
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
                        ("position:absolute;left:" + px (paneX (paneOf cell.Row cell.Col) (rowHeaderWidth + colPrefix.[cell.Col] - 3.0)) + ";top:" + px (paneY (paneOf cell.Row cell.Col) (headerHeight + rowPrefix.[cell.Row] - 3.0))
                         + ";width:" + px (w + 6.0) + ";height:" + px (h + 6.0)
                         + ";box-sizing:border-box;border:2px solid #E00000;border-radius:50%;pointer-events:none;z-index:6")
        while circleNodes.Count > used do
            detach circleNodes.[circleNodes.Count - 1]
            Portable.Stack.dropLast circleNodes
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
        let below (row: int) = headerHeight + rowPrefix.[min rowCount (max 0 row + 1)] + 3.0 + (if row < frozenRows then scrollTop else 0.0) - originY
        let leftOf (col: int) = rowHeaderWidth + colPrefix.[min colCount (max 0 col)] + (if col < frozenCols then scrollLeft else 0.0) - originX
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
            let current = SpreadsheetCellText.cellDisplayTextIn model.Date1904 None 1.0e6 defaultFormat (getCellValue addr model)
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

        // The texts the column's cells show in the filter's rows ("" stands for the blank cells).
        let uniqueValues =
            let values, blanks = filterColumnValues col model
            if blanks then values @ [ "" ] else values

        // What the column's filter lets through now (everything without one,
        // and for a filter that is not a list of values).
        let visibleValues =
            match Map.tryFind col model.Filters with
            | Some { Criteria = FilterValues (values, blanks) } ->
                let wanted = values |> Set.map (fun value -> value.ToUpperInvariant())
                uniqueValues |> List.filter (fun value -> if value = "" then blanks else Set.contains (value.ToUpperInvariant()) wanted) |> Set.ofList
            | _ -> Set.ofList uniqueValues

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

            // Apply filter (one undo step)
            if checkedValues.Count = uniqueValues.Length then
                // All selected = no filter
                changeFilter (sprintf "Clear filter of column %s" (columnToLetter col)) (clearColumnFilter (Some col)) |> ignore
            else
                changeFilter (sprintf "Filter column %s" (columnToLetter col))
                    (setColumnFilter col (FilterValues (Set.remove "" checkedValues, Set.contains "" checkedValues))) |> ignore

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


    /// One search session per grid, hosted by its shell when one is mounted.
    and closeFindBar () =
        findBar |> Option.iter (fun bar ->
            setStyle bar "display" "none"
            setAttribute bar "data-hidden" "true")
        findFocus |> Option.iter (fun target -> if unbox<bool> (getProp target "isConnected") then focus target)
        findFocus <- None

    and ensureFindBar () =
        match findBar with
        | Some bar -> bar
        | None ->
            let bar = createElement "div"
            setClassName bar "sheet-find-bar"
            setAttribute bar "data-hidden" "true"
            setAttribute bar "role" "search"
            setAttribute bar "aria-label" "Find and replace"
            setStyles bar [
                "display", "none"; "flex-wrap", "wrap"; "align-items", "center"
                "gap", "8px"; "padding", "8px"; "box-sizing", "border-box"
                "background", "#f8f8f8"; "border-bottom", "1px solid #ccc"
                "font-family", "Arial, sans-serif"; "font-size", "12px"
            ]
            let input name placeholder =
                let field = createElement "input"
                setAttribute field "type" "text"
                setAttribute field "aria-label" name
                setAttribute field "placeholder" placeholder
                field
            let label caption field =
                let node = createElement "label"
                setTextContent node caption
                appendChild node field
                appendChild bar node
            let optionSelect name className choices =
                let field = createElement "select"
                setAttribute field "aria-label" name
                setClassName field className
                for value, caption in choices do
                    let option = createElement "option"
                    setAttribute option "value" value
                    setTextContent option caption
                    appendChild field option
                label (name + " ") field
                field
            let button icon caption action =
                let node = createElement "button"
                setAttribute node "type" "button"
                setAttribute node "data-sheet-icon" icon
                setTextContent node caption
                listen node "click" (fun _ -> action ())
                appendChild bar node
                node
            let findInput = input "Find in spreadsheet" "Search..."
            setClassName findInput "sheet-find-input"
            label "Find " findInput
            let replaceInput = input "Replace with" "Replace with..."
            setClassName replaceInput "sheet-replace-input"
            label "Replace " replaceInput
            let caseInput = createElement "input"
            setAttribute caseInput "type" "checkbox"
            setClassName caseInput "sheet-find-case"
            label "Match case " caseInput
            let wholeInput = createElement "input"
            setAttribute wholeInput "type" "checkbox"
            setClassName wholeInput "sheet-find-whole"
            label "Whole cell " wholeInput
            let regexInput = createElement "input"
            setAttribute regexInput "type" "checkbox"
            setClassName regexInput "sheet-find-regex"
            label "Regular expression " regexInput
            let lookIn = optionSelect "Look in" "sheet-find-look-in" [ "values", "Values"; "formulas", "Formulas" ]
            let scope = optionSelect "Scope" "sheet-find-scope" [ "sheet", "Sheet"; "workbook", "Workbook" ]
            let range = input "Search range" "B2:D9 (optional)"
            setClassName range "sheet-find-range"
            label "Range " range
            let result = createElement "span"
            setClassName result "sheet-find-result"
            setAttribute result "role" "status"
            setAttribute result "aria-live" "polite"
            setTextContent result "Enter text to find"
            let status message = setTextContent result message
            let options () =
                createObj [
                    "matchCase" ==> getChecked caseInput
                    "wholeCell" ==> getChecked wholeInput
                    "regex" ==> getChecked regexInput
                    "lookIn" ==> getValue lookIn
                    "scope" ==> getValue scope
                    if getValue range <> "" then "range" ==> getValue range
                ]
            let mutable matches: obj array = [||]
            let mutable query = ""
            let mutable settings = ""
            let mutable seenSerial = -1
            let mutable seenSheet = -1
            let mutable replaceButtons: Element array = [||]
            let refresh () =
                if not disposed && getAttribute bar "data-hidden" = "false" then
                    for node in replaceButtons do setProp node "disabled" (box readOnly)
                    let text = getValue findInput
                    let currentSettings =
                        string (getChecked caseInput) + "|" + string (getChecked wholeInput)
                        + "|" + string (getChecked regexInput) + "|" + getValue lookIn
                        + "|" + getValue scope + "|" + getValue range
                    let changed = text <> query || currentSettings <> settings || findEpoch <> session.Epoch
                    if changed then findSelected <- false
                    query <- text
                    settings <- currentSettings
                    findEpoch <- session.Epoch
                    try
                        if text = "" then
                            matches <- [||]
                            findSession <- None
                        elif changed || seenSerial <> session.ChangeSerial || seenSheet <> activeSheetIndex then
                            matches <- this.Find(box text, options ())
                            findSession <- hostSearch
                        seenSerial <- session.ChangeSerial
                        seenSheet <- activeSheetIndex
                        if text = "" then status "Enter text to find"
                        else
                            let active = activeCellAddress ()
                            let index =
                                if findSelected then
                                    matches |> Array.tryFindIndex (fun found ->
                                        unbox<int> (getProp found "sheetIndex") = activeSheetIndex
                                        && unbox<int> (getProp found "row") = active.Row
                                        && unbox<int> (getProp found "col") = active.Col)
                                else None
                            match index with
                            | Some index -> status (sprintf "%d of %d matches" (index + 1) matches.Length)
                            | None -> status (sprintf "%d matches" matches.Length)
                    with error ->
                        matches <- [||]
                        status ("Find failed: " + errorMessageOf (box error))
            let ensureOwnSearch () =
                match hostSearch, findSession with
                | Some current, Some own when obj.ReferenceEquals(current, own) -> ()
                | _ ->
                    this.Find(box query, options ()) |> ignore
                    findSession <- hostSearch
            findRefresh <- refresh
            let navigate forward =
                if not (this.TryCommitEdit()) then status "Find blocked: correct the cell entry first."
                else
                    refresh ()
                    if query <> "" && matches.Length > 0 then
                        try
                            ensureOwnSearch ()
                            let found = if forward then this.FindNext() else this.FindPrevious()
                            findSelected <- not (isNull found)
                            refresh ()
                        with error -> status ("Find failed: " + errorMessageOf (box error))
            let replace single =
                if readOnly then status "Replace unavailable in read-only mode."
                elif not (this.TryCommitEdit()) then status "Replace blocked: correct the cell entry first."
                else
                    refresh ()
                    if query <> "" && matches.Length > 0 then
                        try
                            if single then ensureOwnSearch ()
                            if single then
                                let outcome = this.ReplaceMatch(box (getValue replaceInput))
                                findSelected <- not (isNull (getProp outcome "next"))
                                refresh ()
                                if not (unbox<bool> (getProp outcome "replaced")) then
                                    status "No replaceable match at current selection"
                            else
                                let count = this.ReplaceAll(box query, box (getValue replaceInput), options ())
                                findSelected <- false
                                refresh ()
                                status (sprintf "Replaced %d cells" count)
                        with error -> status ("Replace failed: " + errorMessageOf (box error))
            button "find-run" "Next" (fun () -> navigate true) |> ignore
            button "find-previous" "Previous" (fun () -> navigate false) |> ignore
            let replaceButton = button "replace-one" "Replace" (fun () -> replace true)
            let replaceAllButton = button "replace-run" "Replace all" (fun () -> replace false)
            replaceButtons <- [| replaceButton; replaceAllButton |]
            findNextAction <- (fun () -> navigate true)
            findReplaceAllAction <- (fun () -> replace false)
            button "find-close" "Close" closeFindBar |> ignore
            for field in [ findInput; caseInput; wholeInput; regexInput; lookIn; scope; range ] do
                listen field (if obj.ReferenceEquals(field, findInput) || obj.ReferenceEquals(field, range) then "input" else "change")
                    (fun _ -> findSelected <- false; refresh ())
            listen bar "keydown" (fun event ->
                match getEventKey event with
                | "Escape" -> preventDefault event; closeFindBar ()
                | "Enter" when not (eventFlag event "isComposing") ->
                    preventDefault event
                    navigate (not (eventFlag event "shiftKey"))
                | _ -> ())
            for node in replaceButtons do setProp node "disabled" (box readOnly)
            let selectionSubscription = this.OnSelectionChange(System.Action<obj>(fun _ -> refresh ()))
            let changeSubscription = this.OnChange(System.Action<obj>(fun _ -> refresh ()))
            let readonlySubscription = this.OnReadOnlyChange(System.Action<bool>(fun _ -> refresh ()))
            cleanups.Add(fun () ->
                selectionSubscription.Invoke()
                changeSubscription.Invoke()
                readonlySubscription.Invoke()
                detach bar)
            appendChild bar result
            let host = findHost |> Option.defaultValue container
            if obj.ReferenceEquals(host, container) then
                let _ : obj = emitJsExpr (container, bar) "$0.insertBefore($1, $0.firstChild)"
                ()
            else appendChild host bar
            findBar <- Some bar
            bar

    and showFindReplaceDialog () =
        if not disposed && gridContainer.IsSome then
            let bar = ensureFindBar ()
            if getAttribute bar "data-hidden" = "true" then
                findFocus <- Some (activeElement ())
                setStyle bar "display" "flex"
                setAttribute bar "data-hidden" "false"
            findRefresh ()
            let input = queryOne bar ".sheet-find-input"
            if not (isNull input) then focus input


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

    /// A function module or a provider's answer arrived (`Formula.onFunctionsArrived`):
    /// the formulas that showed #BUSY! are calculated again - once for a burst of
    /// arrivals, and not while a cell is being edited.
    let mutable functionsArrivalScheduled = false
    let rec recalculateForArrivedFunctions () =
        if not disposed then
            if model.EditingCell.IsSome then
                setTimeoutMs recalculateForArrivedFunctions 250 |> ignore
            else
                functionsArrivalScheduled <- false
                recalculateFromScratch ()
                render ()
    do
        let stop =
            Rbt.Spreadsheet.Formula.onFunctionsArrived (fun () ->
                if not functionsArrivalScheduled && not disposed then
                    functionsArrivalScheduled <- true
                    setTimeoutMs recalculateForArrivedFunctions 0 |> ignore)
        cleanups.Add stop

    /// Mount the spreadsheet
    member this.Mount() =
        // A cell's text is measured in its font (`excelLetterSpacing`): drawn before a page's font has
        // loaded, it is drawn again once it has.
        let stopWatchingFonts =
            onFontsLoaded (fun () ->
                if not fontRedrawPending && not disposed then
                    fontRedrawPending <- true
                    setTimeoutMs (fun () ->
                        fontRedrawPending <- false
                        if not disposed then
                            renderedStyleInputs <- [||]
                            render ()) 30 |> ignore)
        cleanups.Add stopWatchingFonts
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
            // (Text a cell cannot show is cut at its edge, as Excel does: never "…".)
            (".spreadsheet-container .spreadsheet-cell{position:absolute;padding:3px 3px 3px 2px;overflow:hidden;white-space:nowrap;text-overflow:clip;box-sizing:border-box;font-kerning:none;"
             + "font-family:Arial, sans-serif;font-size:" + cssPx (SheetOps.pointsToPx 11.0) + ";cursor:default;background:white;border:none;box-shadow:inset 0 1px 0 0 " + defaultGridColor + ",inset 1px 0 0 0 " + defaultGridColor + ";contain:strict}"
             + ".spreadsheet-container .spreadsheet-col-header{position:absolute;top:0;height:30px;border-right:1px solid #d0d0d0;border-bottom:1px solid #a6a6a6;"
             + "background:linear-gradient(180deg, #f8f8f8 0%, #e8e8e8 100%);display:flex;align-items:center;justify-content:space-between;padding:0 4px;box-sizing:border-box;"
             + "font-weight:normal;font-size:11px;font-family:Arial, sans-serif;color:#444;user-select:none;overflow:hidden}"
             + ".spreadsheet-container .spreadsheet-row-header{position:absolute;left:0;width:50px;border-right:1px solid #a6a6a6;border-bottom:1px solid #d0d0d0;"
             + "background:linear-gradient(90deg, #f8f8f8 0%, #e8e8e8 100%);display:flex;align-items:center;justify-content:center;box-sizing:border-box;"
             + "font-weight:normal;font-size:11px;font-family:Arial, sans-serif;color:#444;user-select:none;cursor:default;overflow:hidden}"
             + ".spreadsheet-container [data-headers=off] .spreadsheet-col-header,.spreadsheet-container [data-headers=off] .spreadsheet-row-header{display:none}"
             + ".spreadsheet-container .col-header-text{flex:1;text-align:center}"
             + ".spreadsheet-container .col-filter-btn{width:16px;height:16px;align-items:center;justify-content:center;cursor:pointer;font-size:10px;border-radius:2px}"
             + ".spreadsheet-container .col-filter-btn:hover{background:#d0d0d0;opacity:1 !important}"
             + ".spreadsheet-container .spreadsheet-outline{position:absolute;box-sizing:border-box;user-select:none;font-family:Arial, sans-serif}"
             + ".spreadsheet-container .spreadsheet-outline-button,.spreadsheet-container .spreadsheet-outline-level{width:13px;height:13px;border:1px solid #777;background:#fff;color:#333;"
             + "font-size:10px;line-height:11px;text-align:center;cursor:pointer;z-index:5}"
             + ".spreadsheet-container .spreadsheet-outline-level{font-size:9px;background:#f8f8f8}"
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
        cornerElement <- Some corner
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
        listen editBox "keyup" (fun _ -> updateFormulaHelp ())
        listen editBox "mouseup" (fun _ -> updateFormulaHelp ())
        layers <-
            Some { Panes = panes
                   ColHeaders = colHeaderStrip
                   RowHeaders = rowHeaderStrip
                   FrozenHeaders = frozenHeaders
                   Selections = panes |> Array.map (overlay "spreadsheet-selection")
                   Actives = panes |> Array.map (overlay "spreadsheet-active-cell")
                   FillHandle = overlay "spreadsheet-fill-handle" content
                   EditInput = editBox
                   EditHelp = overlay "spreadsheet-formula-help" content
                   Notice = overlay "spreadsheet-validation-notice" content
                   InputMessage = overlay "spreadsheet-input-message" content
                   CommentBox = overlay "spreadsheet-comment-box" content
                   ListButton = overlay "spreadsheet-list-button" content }
        layers |> Option.iter (fun layer ->
            setAttribute layer.EditHelp "role" "status"
            setAttribute layer.EditHelp "aria-live" "polite")

        // The box of a comment, beside its cell while the pointer is over the
        // cell: a note in its runs (font, size, bold, italic, underline,
        // colour), a thread as its comments with their authors.
        let hideCommentBox () =
            if commentBoxCell.IsSome then
                commentBoxCell <- None
                match layers with
                | Some layer -> setStyle layer.CommentBox "display" "none"
                | None -> ()
        let showCommentBox (addr: CellAddress) =
            match layers, SheetOps.commentAt model addr with
            | Some layer, Some info when commentBoxCell <> Some (addr, info) ->
                commentBoxCell <- Some (addr, info)
                let panel = layer.CommentBox
                let metrics = SpreadsheetLayout.metricsOf model
                // (Beside the whole of a merged cell.)
                let right, top =
                    match findMergedCell addr model with
                    | Some merged -> merged.EndCol, merged.StartRow
                    | None -> addr.Col, addr.Row
                let x = rowHeaderWidth + metrics.ColPrefix.[min metrics.ColCount (right + 1)] + (if right < model.FrozenCols then scrollLeft else 0.0) + 6.0
                let y = headerHeight + metrics.RowPrefix.[min metrics.RowCount (max 0 top)] + (if top < model.FrozenRows then scrollTop else 0.0)
                setCssText panel
                    ("position:absolute;left:" + string (int (x - originX)) + "px;top:" + string (int (y - originY)) + "px;min-width:96px;max-width:280px;padding:3px 5px;background:"
                     + (if info.Thread.IsSome then "#FFFFFF" else "#FFFFE1") + ";border:1px solid " + (if info.Thread.IsSome then "#7b61ff" else "#000000")
                     + ";box-shadow:2px 2px 4px rgba(0,0,0,0.3);font-family:Tahoma, Arial, sans-serif;font-size:12px;color:#000000;white-space:pre-wrap;overflow-wrap:break-word;pointer-events:none;z-index:31")
                setAttribute panel "data-row" (string addr.Row)
                setAttribute panel "data-col" (string addr.Col)
                setInnerHTML panel ""
                let line (parent: Element) (text: string) (css: string) =
                    let element = createElement "div"
                    setTextContent element text
                    if css <> "" then setCssText element css
                    appendChild parent element
                match info.Thread with
                | Some thread ->
                    setAttribute panel "data-kind" "thread"
                    let entry (indent: bool) (comment: CommentThread) =
                        let block = createElement "div"
                        setClassName block "spreadsheet-comment-entry"
                        setCssText block ((if indent then "margin:6px 0 0 12px" else "margin:0"))
                        line block (comment.Author |> Option.defaultValue "") "font-weight:bold"
                        (match comment.Date with
                         | Some date -> line block (date.Replace("T", " ")) "font-size:10px;color:#666666"
                         | None -> ())
                        line block comment.Text ""
                        appendChild panel block
                    entry false thread
                    for reply in thread.Replies do entry true reply
                    if thread.Done then line panel "Resolved" "margin-top:6px;font-size:10px;color:#1E7B34;font-weight:bold"
                | None ->
                    setAttribute panel "data-kind" "note"
                    if info.Runs.IsEmpty then
                        // (A note that starts with its author's name says who wrote it.)
                        (match info.Author with
                         | Some author when author <> "" && not (info.Text.StartsWith author) -> line panel (author + ":") "font-weight:bold"
                         | _ -> ())
                        line panel info.Text ""
                    else
                        for text, run in info.Runs do
                            let span = createElement "span"
                            setTextContent span text
                            let css =
                                [ (if run.Bold = Some true then "font-weight:bold" else "")
                                  (if run.Italic = Some true then "font-style:italic" else "")
                                  (match run.Underline = Some true, run.Strikethrough = Some true with
                                   | true, true -> "text-decoration:underline line-through"
                                   | true, false -> "text-decoration:underline"
                                   | false, true -> "text-decoration:line-through"
                                   | false, false -> "")
                                  (match run.Color with Some color -> "color:" + color | None -> "")
                                  (match run.FontName with Some name -> "font-family:'" + name.Replace("'", "") + "', Tahoma, Arial, sans-serif" | None -> "")
                                  (match run.FontSize with Some size -> "font-size:" + string size + "pt" | None -> "") ]
                                |> List.filter (fun part -> part <> "")
                                |> String.concat ";"
                            if css <> "" then setCssText span css
                            appendChild panel span
                // The box stays inside what is in view: left of its cell when there is no room on
                // the right of it, and further up when it would reach below the last row in view.
                let boxWidth, boxHeight = getClientWidth panel + 2.0, getClientHeight panel + 2.0
                let viewLeft, viewTop = scrollLeft + rowHeaderWidth, scrollTop + headerHeight
                let viewRight, viewBottom = scrollLeft + gridClientWidth gridDiv - 2.0, scrollTop + gridClientHeight gridDiv - 2.0
                let placedX =
                    if x + boxWidth <= viewRight then x
                    else
                        let leftCol = match findMergedCell addr model with Some merged -> merged.StartCol | None -> addr.Col
                        let cellLeft = rowHeaderWidth + metrics.ColPrefix.[min metrics.ColCount (max 0 leftCol)] + (if leftCol < model.FrozenCols then scrollLeft else 0.0)
                        let flipped = cellLeft - 6.0 - boxWidth
                        if flipped >= viewLeft then flipped else max viewLeft (viewRight - boxWidth)
                let placedY = if y + boxHeight <= viewBottom then y else max viewTop (viewBottom - boxHeight)
                if placedX <> x then setStyle panel "left" (string (int (placedX - originX)) + "px")
                if placedY <> y then setStyle panel "top" (string (int (placedY - originY)) + "px")
            | _, None -> hideCommentBox ()
            | _ -> ()
        listen content "mousemove" (fun e ->
            // (Not while a button is down: a drag selects, resizes or moves.)
            if pointerButtons e <> 0 then hideCommentBox ()
            else
                match hitTestEvent e with
                | SpreadsheetLayout.Cell addr ->
                    // A merged cell's comment is on its first cell.
                    let addr = match findMergedCell addr model with Some merged -> { Row = merged.StartRow; Col = merged.StartCol } | None -> addr
                    let commented =
                        (not (Map.isEmpty model.Threads) && Map.containsKey addr model.Threads)
                        || (match Map.tryFind addr model.Cells with Some { Comment = Some _ } -> true | _ -> false)
                    if commented then showCommentBox addr else hideCommentBox ()
                | _ -> hideCommentBox ())
        listen content "mouseleave" (fun _ -> hideCommentBox ())
        listen content "mousedown" (fun _ -> hideCommentBox ())

        // Add event listeners
        listen gridDiv "scroll" handleScroll
        listen gridDiv "wheel" (fun e ->
            // A scrollbar thumb maps the entire sheet, but a wheel or
            // trackpad still travels in screen pixels, not in scrollbar ratios.
            if not (eventFlag e "ctrlKey") && (extentWidth > maxLogicalExtent () || extentHeight > maxLogicalExtent ()) then
                let mode = wheelDeltaMode e
                let dx, dy = wheelDeltaX e, wheelDeltaY e
                let dx, dy = if eventFlag e "shiftKey" && dx = 0.0 then dy, 0.0 else dx, dy
                let unitX = if mode = 1 then 16.0 elif mode = 2 then getClientWidth gridDiv else 1.0
                let unitY = if mode = 1 then 16.0 elif mode = 2 then getClientHeight gridDiv else 1.0
                preventDefault e
                if dx <> 0.0 then setGridScrollLeft gridDiv (gridScrollLeft gridDiv + dx * unitX / appliedZoom)
                if dy <> 0.0 then setGridScrollTop gridDiv (gridScrollTop gridDiv + dy * unitY / appliedZoom)
                handleScroll e)
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
            elif within target ".spreadsheet-outline" then
                // The outline strips: a button collapses / expands; nothing is selected.
                preventDefault e
                stopPropagation e
                let button = closest target ".spreadsheet-outline-button, .spreadsheet-outline-level"
                if mouseButton e = 0 && not (isNull (box button)) then pressOutline button
            elif within target ".spreadsheet-filter-button" then
                // (Not a click on the cell under it: the selection stays.)
                preventDefault e
                stopPropagation e
                if mouseButton e = 0 then pressFilterButton (closest target ".spreadsheet-filter-button")
            elif within target "[data-macro]" then
                // (A shape's macro runs on a click, as in Excel; it is not selected.)
                preventDefault e
                // (The keys stay the grid's: a macro may have bound some with OnKey.)
                if not (holdsFocus content) then focusQuietly content
                if mouseButton e = 0 && macrosEnabled then runMacroWhenIdle (getAttribute (closest target "[data-macro]") "data-macro")
            elif within target ".spreadsheet-picture-handle, .spreadsheet-picture" then
                // A picture is selected by a click on it, moved by dragging it
                // and sized by dragging a handle of its frame.
                preventDefault e
                let onHandle = within target ".spreadsheet-picture-handle"
                let holder = if onHandle then closest target ".spreadsheet-picture-frame" else closest target ".spreadsheet-picture"
                match System.Int32.TryParse (getAttribute holder "data-drawing-id") with
                | true, id when (SheetOps.drawingById id model).IsSome ->
                    if model.EditingCell.IsSome then finishEditing true |> ignore
                    if not (holdsFocus content) then focusQuietly content
                    selectedPicture <- Some id
                    if mouseButton e = 0 && not readOnly then
                        let rect = SpreadsheetLayout.drawingRect model (SheetOps.drawingById id model).Value
                        pictureDrag <-
                            Some { Id = id
                                   Handle = (if onHandle then getAttribute (closest target ".spreadsheet-picture-handle") "data-handle" else "move")
                                   StartX = getClientX e; StartY = getClientY e; Start = rect; Current = rect; Moved = false }
                    render()
                | _ -> ()
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
                    // (A click on a cell takes the selection off a picture.)
                    selectedPicture <- None
                    if eventFlag e "shiftKey" then
                        // Shift+click: from the active cell to here.
                        selectFromTo (activeCellAddress()) hitAddr
                    else
                        isDragging <- true
                        dragStartCell <- Some hitAddr
                        selectCellAddress hitAddr
                    render()
                | _ -> ())
        listen content "contextmenu" (fun e ->
            let target = getEventTarget e
            // (In the edit box the browser's menu stays: spelling, paste.)
            if contextListeners.Count > 0 && not (within target ".spreadsheet-edit-input") then
                preventDefault e
                // (The menu key was answered at its keydown.)
                if nowMs () - keyboardMenuAt > 400.0 then
                    let top, left, bottom, right = selectionBounds()
                    let lastRow = max 0 (model.RowCount - 1)
                    let lastCol = max 0 (model.ColCount - 1)
                    let tell (kind: string) (row: int) (col: int) (picture: int option) (inSelection: bool) =
                        raiseContextMenu kind row col picture inSelection (getClientX e) (getClientY e) activeSheetIndex |> ignore
                    // What is under the pointer becomes the selection, unless it is part of it.
                    let take (select: unit -> unit) =
                        if model.EditingCell.IsSome then finishEditing true |> ignore
                        if not (holdsFocus content) then focusQuietly content
                        selectedPicture <- None
                        select ()
                        render()
                    if within target ".spreadsheet-picture-handle, .spreadsheet-picture-frame, .spreadsheet-picture" then
                        let holder = closest target "[data-drawing-id]"
                        match (if isNull (box holder) then (false, 0) else System.Int32.TryParse (getAttribute holder "data-drawing-id")) with
                        | true, id ->
                            match SheetOps.drawingById id model with
                            | Some drawing ->
                                if selectedPicture <> Some id then
                                    selectedPicture <- Some id
                                    render()
                                tell "picture" drawing.Anchor.FromRow drawing.Anchor.FromCol (Some id) true
                            | None -> ()
                        | _ -> ()
                    elif within target ".spreadsheet-corner" then
                        tell "corner" -1 -1 None (selectedPicture.IsNone && top = 0 && left = 0 && bottom >= lastRow && right >= lastCol)
                    elif within target ".spreadsheet-col-header" then
                        match headerIndex target ".spreadsheet-col-header" "data-col" with
                        | Some col ->
                            let inside = selectedPicture.IsNone && top = 0 && bottom >= lastRow && col >= left && col <= right
                            if not inside then
                                take (fun () ->
                                    anchorCell <- { Row = 0; Col = col }
                                    model <- { model with Selection = Range (0, col, lastRow, col) })
                            tell "colHeader" -1 col None inside
                        | None -> ()
                    elif within target ".spreadsheet-row-header" then
                        match headerIndex target ".spreadsheet-row-header" "data-row" with
                        | Some row ->
                            let inside = selectedPicture.IsNone && left = 0 && right >= lastCol && row >= top && row <= bottom
                            if not inside then
                                take (fun () ->
                                    anchorCell <- { Row = row; Col = 0 }
                                    model <- { model with Selection = Range (row, 0, row, lastCol) })
                            tell "rowHeader" row -1 None inside
                        | None -> ()
                    else
                        match hitTestEvent e with
                        | SpreadsheetLayout.Cell addr ->
                            let inside = selectedPicture.IsNone && addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right
                            if not inside then take (fun () -> selectCellAddress addr)
                            tell "cell" addr.Row addr.Col None inside
                        | SpreadsheetLayout.HitColumnHeader _ | SpreadsheetLayout.HitRowHeader _ | SpreadsheetLayout.Corner | SpreadsheetLayout.Empty -> ())
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
            elif not (within target ".spreadsheet-edit-input, .spreadsheet-col-header, .spreadsheet-row-header, .spreadsheet-corner, .spreadsheet-fill-handle, .spreadsheet-list-button, .spreadsheet-filter-button, .spreadsheet-picture, .spreadsheet-picture-handle") then
                match hitTestEvent e with
                | SpreadsheetLayout.Cell hitAddr -> startEditing hitAddr.Row hitAddr.Col
                | _ -> ())
        listen content "click" (fun e ->
            let target = getEventTarget e
            // A link is followed by Ctrl+click (Cmd+click), in a viewer by a
            // plain click: to a place in the workbook the grid goes itself,
            // an address is handed to the host (OnHyperlinkActivate).
            if mouseButton e = 0 && (readOnly || eventFlag e "ctrlKey" || eventFlag e "metaKey") && not model.Hyperlinks.IsEmpty
               && not (within target ".spreadsheet-edit-input, .spreadsheet-col-header, .spreadsheet-row-header, .spreadsheet-corner, .spreadsheet-fill-handle, .spreadsheet-list-button, .spreadsheet-filter-button, .spreadsheet-picture, .spreadsheet-picture-handle") then
                match hitTestEvent e with
                | SpreadsheetLayout.Cell hitAddr when (SheetOps.hyperlinkAt model hitAddr).IsSome ->
                    preventDefault e
                    this.ActivateHyperlink (box hitAddr.Row, box hitAddr.Col) |> ignore
                | _ -> ()
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
                    (try sortActiveSheet "sort" col ascending (sprintf "Sort by column %s" (columnToLetter col)) with error -> reportError "sort" (box error))
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
        // A drag of the selection or of the fill handle goes where the pointer is ...
        let dragTo (e: obj) =
            if isFilling then
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
        // ... and while the pointer is held at an edge of the grid, or past it, the sheet scrolls
        // under it - faster the further out the pointer is - and the drag goes on to the rows and
        // columns that come into view (the sheet grows with the selection, see `giveRoom`): a
        // fill or a selection past what the grid shows needs no second drag.
        let dragScrollStep () =
            match gridContainer, dragPointer with
            | Some grid, Some e when isFilling || isDragging ->
                let width = gridClientWidth grid
                let height = gridClientHeight grid
                let x = pointerX e grid
                let y = pointerY e grid
                let metrics = SpreadsheetLayout.metricsOf model
                let frozenW = metrics.ColPrefix.[max 0 (min metrics.ColCount model.FrozenCols)]
                let frozenH = metrics.RowPrefix.[max 0 (min metrics.RowCount model.FrozenRows)]
                let edge = 4.0
                let speed (past: float) = min 600.0 (25.0 + max 0.0 past * 1.5)
                let dy =
                    if y >= height - edge then speed (y - height)
                    elif y <= headerHeight + frozenH + edge && gridScrollTop grid > 0.0 then - speed (headerHeight + frozenH - y)
                    else 0.0
                let dx =
                    if x >= width - edge then speed (x - width) * 2.0
                    elif x <= rowHeaderWidth + frozenW + edge && gridScrollLeft grid > 0.0 then - speed (rowHeaderWidth + frozenW - x) * 2.0
                    else 0.0
                if dy <> 0.0 || dx <> 0.0 then
                    if dy <> 0.0 then setGridScrollTop grid (max 0.0 (gridScrollTop grid + dy))
                    if dx <> 0.0 then setGridScrollLeft grid (max 0.0 (gridScrollLeft grid + dx))
                    scrollTop <- gridScrollTop grid
                    scrollLeft <- gridScrollLeft grid
                    dragTo e
            | _ -> stopDragScroll ()
        startDragScrollHook.Value <- (fun () ->
            if dragScrollTimer.IsNone then dragScrollTimer <- Some (JS.setInterval dragScrollStep 50))
        let onDocumentMouseMove = (fun (e: obj) ->
            if pictureDrag.IsSome then
                match pictureDrag with
                | Some drag ->
                    let dx = (getClientX e - drag.StartX) / appliedZoom
                    let dy = (getClientY e - drag.StartY) / appliedZoom
                    if drag.Moved || abs dx >= 2.0 || abs dy >= 2.0 then
                        drag.Moved <- true
                        drag.Current <- draggedRect drag dx dy (eventFlag e "shiftKey")
                        // (Only the frame and the picture move: no render of the cells.)
                        // (Over frozen panes the picture is cut anew where the drag has it.)
                        if model.FrozenRows > 0 || model.FrozenCols > 0 then renderDrawings content (SpreadsheetLayout.metricsOf model)
                        else renderPictureFrame content
                | None -> ()
            elif isResizingColumn then
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
            elif isFilling || isDragging then
                dragPointer <- Some e
                dragTo e
                startDragScroll ()
        )

        // Mouse up for both drag selection and resize
        let onDocumentMouseUp = (fun (_: obj) ->
            stopDragScroll ()
            if pictureDrag.IsSome then
                // The whole drag is one undo step.
                let drag = pictureDrag.Value
                pictureDrag <- None
                if drag.Moved then
                    (try placePictureAt "picture" (if drag.Handle = "move" then "Move picture" else "Resize picture") drag.Id drag.Current |> ignore with error -> reportError "picture" (box error))
                    // (Drawn again from the model, whatever came of it.)
                    renderedDrawings <- null
                render()
            elif isFilling then
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
        headBand <- 30.0
        sideBand <- 50.0
        headerHeight <- 30.0
        rowHeaderWidth <- 50.0
        gutterWidth <- 0.0
        gutterHeight <- 0.0
        outlineNodes.Clear()
        outlineSignature <- [||]

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
        clipboardGeneration <- clipboardGeneration + 1
        session.ClearClipboard()
        for cleanup in cleanups.ToArray() do
            try cleanup () with _ -> ()
        cleanups.Clear()
        findBar <- None
        findHost <- None
        findFocus <- None
        findRefresh <- (fun () -> ())
        findNextAction <- (fun () -> ())
        findReplaceAllAction <- (fun () -> ())
        findSelected <- false
        findSession <- None
        match noticeTimer with
        | Some id ->
            clearTimeoutId id
            noticeTimer <- None
        | None -> ()
        // Cancel image loads before revoking URLs, including a mount disposed
        // in the same event turn in which its picture nodes were created.
        match gridContainer with
        | Some gridDiv -> clearImageSources gridDiv
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
        session.InvalidateCalculation()
        engineGeneration <- engineGeneration + 1
        overflowSources.Clear()
        overflowCells <- null
        notifiedSheets <- [||]
        notifiedSelection <- None
        validationListeners.Clear()
        isDragging <- false
        stopDragScroll ()
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
        readOnlyListeners.Clear()
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
    /// another sheet). It receives { SheetIndex, Selection: "B2" | "B2:D5",
    /// ActiveCell, ActiveRow, ActiveCol, Picture (id of the selected picture
    /// or null), Status (what GetStatus() gives for the new selection) }.
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
        session.CaptureRevision()
    /// Unlike GetRevision, this changes even when undo restores an earlier revision.
    member _.GetChangeSerial() : int =
        observeRevision ()
        session.ChangeSerial

    /// Identifies the current workbook across replacements, even if its revision matches.
    member _.GetWorkbookEpoch() : int = session.Epoch
    /// Whether a keyboard event target is in this grid (not a neighbouring editor).
    member _.OwnsKeyboardTarget(target: Element) : bool =
        contains container target && not (isNull (gridClosest target ".spreadsheet-container"))



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
        session.MarkSaved(?revision = savedAt)

    /// Whether the workbook's content differs from the saved revision.
    member _.IsModified() : bool =
        observeRevision ()
        session.IsModified

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
        if viewerMode <> value then
            viewerMode <- value
            refreshSheetLock ()
            if readOnly && model.EditingCell.IsSome then
                finishEditing false |> ignore
            else render()
            if readOnlyListeners.Count > 0 then
                let current = viewerMode
                for listener in readOnlyListeners.ToArray() do
                    if readOnlyListeners.Contains listener then
                        try listener current with error -> reportError "OnReadOnlyChange" (box error)

    /// Whether the grid is in viewer mode.
    member _.IsReadOnly() : bool = viewerMode

    /// Call `callback` after viewer mode changes, without a workbook edit.
    /// Returns an unsubscribe function; Dispose removes all subscriptions.
    member _.OnReadOnlyChange(callback: System.Action<bool>) : System.Action =
        let listener = callback.Invoke
        readOnlyListeners.Add listener
        System.Action(fun () -> readOnlyListeners.Remove listener |> ignore)

    /// What kind of sheet the sheet at `index` (default: the one on display)
    /// is: "worksheet", or - for a sheet of an opened file that has no cells -
    /// "chartsheet" / "dialogsheet". Such a sheet is shown as an empty sheet
    /// (a chart sheet with the frame of its chart), is written back as the
    /// part it was read from, and takes no entry: while it is on display the
    /// grid behaves as in viewer mode, and SetCell / EnterText change nothing
    /// (EnterText returns false).
    member _.GetSheetKind(?index: obj) : string =
        match Host.optionalArg (Host.intArg "GetSheetKind" "index") (defaultArg index null) with
        | None -> sheetKindAt activeSheetIndex
        | Some index when index >= 0 && index < sheets.Length -> sheetKindAt index
        | Some index -> Host.hostFail "GetSheetKind" (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)

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
            refreshSheetLock ()
            selectedPicture <- None
            pictureDrag <- None
            // Scroll lives outside the SpreadsheetModel; a sheet starts at
            // its top-left corner (we don't persist scroll yet — reset-to-zero
            // matches Excel's behavior on workbook reload). The container is
            // put there too: the content is drawn for where it is scrolled to.
            resetScroll ()
            render()
            for listener in sheetChangedListeners do
                try listener index with error -> reportError "OnSheetChanged" (box error)

    /// Register a callback invoked after every successful sheet switch.
    /// Chrome uses this to update its tab-strip selection highlight when
    /// the switch is triggered from inside (e.g. via keyboard shortcut)
    /// rather than from the chrome's own click handler.
    member _.OnSheetChanged(listener: System.Action<int>) =
        sheetChangedListeners.Add(listener.Invoke)

    // ===== Sheets =====
    // Sheet-list edits join the shared workbook journal on the sheet displayed
    // after the operation. Undo/redo selects that sheet's steps without
    // dropping edits made later on another sheet.

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
        let name = Host.optionalArg (Host.stringArg memberName "name") name
        let at = Host.optionalArg (Host.intArg memberName "atIndex") atIndex
        executeSheetStructure memberName (Rbt.Spreadsheet.Session.AddSheet (name, at, 1000, 26)) |> ignore
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
        // Reject malformed input before committing an in-progress cell edit.
        match Rbt.Spreadsheet.Workbook.sheetNameProblem name with
        | Some problem -> Host.hostFail memberName problem
        | None -> ()
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let id = sheetIdForIndex index
        let oldName = sheetNames.[index]
        if oldName <> name then convertFormulaSheets ()
        if executeSheetStructure memberName (Rbt.Spreadsheet.Session.RenameSheet (id, name)) then
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
        convertFormulaSheets ()
        executeSheetStructure memberName (Rbt.Spreadsheet.Session.DeleteSheet id) |> ignore
        render()

    /// Move the sheet at index `from` so that it is at index `toIndex`
    /// afterwards. The sheet on display stays on display. One undo step.
    /// No formula is rewritten; a 3-D reference (`Jan:Mar!B1`) reads the
    /// sheets that are between its two ends afterwards, so a workbook that
    /// has one is calculated again.
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
            executeSheetStructure memberName (Rbt.Spreadsheet.Session.MoveSheet (id, target)) |> ignore
            render()

    /// Copy the sheet at `index`; the copy is put right after it, named
    /// "Name (2)" (then "(3)", ...) and shown. Returns the copy's index. The
    /// copy's formulas that named the sheet itself name the copy; the names
    /// of the sheet's scope and its tables are copied (a table under a new
    /// name, "Sales_2", which the copy's structured references follow), and
    /// so are its pictures, comments, formats, validations and conditional
    /// formats, and the drawing objects an opened file only kept for the
    /// sheet - charts, shapes: an export writes each copied chart as a chart
    /// part of its own whose series read the copy. Pivot tables and named
    /// cell styles stay with the original (the copy has a pivot table's
    /// cells as values). One undo step. Throws when there is no such sheet.
    member _.DuplicateSheet(index: obj) : int =
        let memberName = "DuplicateSheet"
        let index = Host.intArg memberName "index" index
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        if model.EditingCell.IsSome then finishEditing true |> ignore
        // (The copy is one of the sheet as it is: a sheet not converted yet is now.)
        sheetModelAt index |> ignore
        let sourceId = sheetIdForIndex index
        executeSheetStructure memberName (Rbt.Spreadsheet.Session.DuplicateSheet sourceId) |> ignore
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
        let id = sheetIdForIndex index
        if executeSheetStructure memberName (Rbt.Spreadsheet.Session.SetSheetVisibility (id, visible, veryHidden)) then
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
            runBookStep "Tab colour" (put color)
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
        // (Another workbook: its macros, if any, come with its import.)
        macroPreserved <- None
        macroProject <- None
        clearMacroCallbacks ()
        macroProjectLoad <- None
        macroLog.Clear()
        macroProjectEdited <- false
        sheetCodeNames <- Map.empty
        workbookCodeName <- "ThisWorkbook"
        macrosEnabled <- false
        workbookOpened <- false
        if newSheets.Length = 0 then () else
        // Session.Replace below advances the workbook epoch and clears history.
        clipboardGeneration <- clipboardGeneration + 1
        importedPackage <- None
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
        activeSheetIndex <- safeIndex
        model <- sheets.[safeIndex]
        refreshSheetLock ()
        workbook <-
            match context with
            | Some wb when List.length wb.SheetOrder = sheets.Length ->
                Some { wb with ActiveSheet = List.item safeIndex wb.SheetOrder }
            | _ ->
                Some (buildWorkbookContext sheets sheetNames (Array.create sheets.Length true) Map.empty Map.empty safeIndex)
        sheetAnchors <- Array.create sheets.Length { Row = 0; Col = 0 }
        anchorCell <- { Row = 0; Col = 0 }
        // The new workbook replaces the session's state and undo journal atomically.
        session.Replace(currentWorkbookSnapshot())
        // Conversion belongs to the newly installed session and stable sheet
        // ID, never to the replaced workbook's committed sheet order.
        convertInPlace safeIndex
        model <- sheets.[safeIndex]
        formulaSheetsConverted <- not (sheets |> Array.exists isDeferredSheet)
        // Browser imports carry cached formula values, including deferred sheets.
        session.RebaseCalculation()
        // (The subscribers of OnSheetsChange are told of the new workbook's sheets.)
        notifiedSheetList <- ""
        resetScroll ()
        render()
        // The workbook as it is installed is the calculated state edits start
        // from (a file's formulas come with their values). Its dependency
        // graph is built a moment later (or at the first edit, when that comes
        // first); the formulas of a sheet that is not converted are read off
        // the file, and the sheet stays as it is.
        let generation = engineGeneration
        setTimeoutMs (fun () ->
            if not disposed && session.IncrementalRecalculation && generation = engineGeneration
               && session.RecalcStats.Formulas < 0 then ensureRecalcEngine ()) 250 |> ignore
        // The first edit converts the sheets its formulas reach: the rows of
        // the sheets that have formulas are read from the file right after
        // the workbook is shown, a sheet at a time, so that the edit does not
        // wait for that. (Nothing is converted: see IsSheetLoaded.)
        (
            let rec readAhead (index: int) =
                if not disposed && not readOnly && generation = engineGeneration && index < sheets.Length then
                    let pending =
                        match sheets.[index].Deferred, sheets.[index].SourceSheet, xlsxSource with
                        | Some deferred, Some origin, Some source when deferred.Loaded.IsNone && deferred.HasFormulas && origin.Book = source.Id ->
                            source.Workbook.Sheets |> List.tryItem origin.Sheet |> Option.filter (fun sheet -> not sheet.RowSource.IsLoaded)
                        | _ -> None
                    match pending with
                    | Some sheet ->
                        (try sheet.Rows |> ignore with _ -> ())
                        setTimeoutMs (fun () -> readAhead (index + 1)) 0 |> ignore
                    | None -> readAhead (index + 1)
            setTimeoutMs (fun () -> readAhead 0) 0 |> ignore)
        for listener in sheetChangedListeners do
            try listener safeIndex with error -> reportError "OnSheetChanged" (box error)

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
        session.SetIncrementalRecalculation(enabled)

    /// What the last recalculation was: { Kind: "incremental" | "full" |
    /// "none", Reason (why it was a full one), ChangedCells, ReachedFormulas,
    /// EvaluatedFormulas (distinct formulas evaluated), Evaluations (how
    /// many evaluations: the formulas of an iterated circular reference are
    /// evaluated once per sweep) - of the last incremental one -, Formulas
    /// (in the dependency graph, -1 before it is built) }. For hosts and checks that want to see that an edit
    /// recalculated what it changed.
    member _.GetRecalcStats() : obj =
        let info = session.RecalcStats
        createObj [
            "Kind" ==> info.Kind
            "Reason" ==> info.Reason
            "ChangedCells" ==> info.Stats.ChangedCells
            "ReachedFormulas" ==> info.Stats.ReachedFormulas
            "EvaluatedFormulas" ==> info.Stats.EvaluatedFormulas
            "Evaluations" ==> info.Stats.Evaluations
            "ComparedCells" ==> SpreadsheetRecalc.comparedCells
            "Formulas" ==> info.Formulas
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
        // (The source first: the sheet's charts and shapes are drawn from its parts, and the
        // workbook is drawn when it is installed. InstallWorkbook keeps a source its sheets name.)
        if models.Length > 0 then
            xlsxSource <- Some imported.Source
        this.InstallWorkbook(models, names, active, Some imported.Context)
        if models.Length > 0 then
            xlsxSource <- Some imported.Source
        macroPreserved <- if Rbt.Spreadsheet.MacroApi.hasProject imported.Source.Workbook.Preserved then Some imported.Source.Workbook.Preserved else None
        let byName = Rbt.Spreadsheet.MacroApi.sheetCodeNames imported.Source.Workbook
        sheetCodeNames <- names |> Array.mapi (fun index name -> sheetIdForIndex index, Map.tryFind name byName |> Option.defaultValue (sprintf "Sheet%d" (index + 1))) |> Map.ofArray
        workbookCodeName <- Rbt.Spreadsheet.MacroApi.workbookCodeName imported.Source.Workbook
        // (What kind of sheet is on display is the file's to say.)
        let wasLocked = sheetLocked
        refreshSheetLock ()
        if sheetLocked <> wasLocked then render()
        // A workbook that was just opened is the saved one.
        this.MarkSaved()

    /// Parse before mutating the current workbook. A false result means a
    /// newer open, edit, or disposed shell vetoed installation at commit time.
    member private this.ImportXlsxCore(data: obj, canInstall: unit -> bool) : JS.Promise<bool> =
        observeRevision ()
        let epoch, serial = session.Epoch, session.ChangeSerial
        reportRejection
            ((ownBytesOf data).``then``(fun (bytes: byte[]) ->
                (loadReaderWorkbook data None).``then``(fun readerWorkbook ->
                    let imported = Rbt.Spreadsheet.XlsxBridge.Import.importWorkbookLazy (fun _ -> false) readerWorkbook
                    let saved = readerWorkbook.Preserved.ActiveTab
                    let active =
                        if saved >= 0 && saved < imported.Sheets.Length && imported.Sheets.[saved].Visible then saved
                        else imported.Sheets |> Array.tryFindIndex (fun sheet -> sheet.Visible) |> Option.defaultValue 0
                    // Validate the active sheet before the commit guard (which may
                    // commit an in-progress cell edit) touches the old workbook.
                    resolveSheet imported.Sheets.[active].Model |> ignore
                    if disposed || session.Epoch <> epoch || session.ChangeSerial <> serial || not (canInstall ()) then false
                    else
                        this.InstallImportedWorkbook(imported, active)
                        observeRevision ()
                        importedPackage <- Some (bytes, session.Revision, activeSheetIndex)
                        true)) |> unbox<JS.Promise<bool>>)
            (reportError "ImportXlsx")

    /// Open an XLSX file, leaving the previous workbook untouched if parsing
    /// fails. A rejected file leaves its previous revision and undo history.
    member this.ImportXlsx(data: obj): JS.Promise<unit> =
        this.ImportXlsxCore(data, fun () -> true).``then``(fun _ -> ())

    /// Open only if `canInstall` still holds after asynchronous parsing.
    /// Resolves false without touching the workbook when installation is vetoed.
    member this.ImportXlsxIfCurrent(data: obj, canInstall: System.Func<bool>): JS.Promise<bool> =
        this.ImportXlsxCore(data, canInstall.Invoke)

    /// Whether every sheet is shown as its file says (zoom, grid lines,
    /// headers, formulas, zeros: what an export writes besides the content),
    /// and the file is a package that says what its parts are (a writer
    /// makes a whole package of one that does not).
    member private _.ShownAsOpened() : bool =
        match xlsxSource with
        | None -> false
        | Some source ->
            let fileSheets = source.Workbook.Sheets |> Array.ofList
            let typed =
                source.Workbook.Preserved.Overrides
                |> List.exists (fun (_, contentType) -> contentType.EndsWith ".main+xml" && contentType.Contains "sheet")
            typed
            && Seq.forall (fun i ->
                    let sheet = if i = activeSheetIndex then model else sheets.[i]
                    // (A sheet that was never converted is the file's.)
                    isDeferredSheet sheet
                    || (match sheet.SourceSheet with
                        | Some from when from.Book = source.Id && from.Sheet = i && from.Sheet < fileSheets.Length ->
                            let file = fileSheets.[from.Sheet]
                            sheet.Zoom = float file.View.ZoomScale / 100.0 && sheet.ShowGridLines = file.ShowGridLines && sheet.ShowHeaders = file.View.ShowRowColHeaders
                            && sheet.ShowFormulas = file.View.ShowFormulas && sheet.ShowZeros = file.ShowZeros
                        | _ -> false))
                (seq { 0 .. sheets.Length - 1 })

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

    /// The editor's workbook as what the XLSX export is made from.
    member private _.XlsxExportInput() : Rbt.Spreadsheet.XlsxBridge.ExportInput =
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
        let input : Rbt.Spreadsheet.XlsxBridge.ExportInput =
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
              Macros = None
              MacroProject =
                match macroProject with
                | Some project when macroProjectEdited && not project.Modules.IsEmpty -> Some (project.ProjectBytes())
                | _ -> None }
        input

    /// The editor's workbook in the XLSX writer's shape. (ExportXlsx writes
    /// this, with the sheets nothing was changed in as the file's own parts
    /// where it has the file.)
    member this.ToXlsxWorkbook() : XlsxWriter.XlsxWorkbookOut =
        Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut (this.XlsxExportInput())
        // (Iterative calculation is a setting of the workbook, not of what was kept of its file.)
        |> Rbt.Spreadsheet.XlsxBridge.Export.withIteration (workbook |> Option.bind (fun wb -> wb.Iteration))

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
    ///
    /// A workbook that ImportXlsx opened and in which nothing was changed
    /// since (the revision is the one it opened with - also after undoing
    /// every change -, the same sheet is on display, and zoom, grid lines,
    /// headers, formulas and zeros are shown as when it opened) is not written
    /// anew: the promise resolves to a copy of the bytes that were opened, so
    /// every part of the package is the file's, byte for byte.
    member this.ExportXlsx() : JS.Promise<byte[]> =
        observeRevision ()
        match importedPackage with
        | Some (bytes, openedAt, shown) when session.Revision = openedAt && activeSheetIndex = shown && this.ShownAsOpened() ->
            resolvedCopy bytes
        | Some (bytes, _, _) ->
            // Sheets nothing was changed in go back as the parts they were
            // read from: their bytes, and the styles and shared strings they
            // index, are taken from the file that was opened.
            let sourceAtExport = xlsxSource
            let iterationAtExport = workbook |> Option.bind (fun wb -> wb.Iteration)
            let keptAtExport =
                match sourceAtExport with
                | Some opened -> opened.Workbook.Preserved.Parts @ opened.Workbook.Preserved.SourceParts
                | None -> []
            let input = this.XlsxExportInput()
            let write (source: obj) (parts: (string * byte[]) list) =
                let out =
                    Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOutWith parts input
                    |> Rbt.Spreadsheet.XlsxBridge.Export.withIteration iterationAtExport
                // (A part that goes back as it was read - a sheet, and what the reader kept of the
                // file: charts, pivot caches, media - is the file's compressed entry, not deflated again.)
                let kept = keptAtExport
                writeXlsxPackage source (parts @ kept) out
                |> packageBytes
            let wanted = try Rbt.Spreadsheet.XlsxBridge.Export.verbatimPartsWanted input with _ -> []
            let fetched : JS.Promise<obj * (string * byte[]) list> =
                (ZipInterop.loadRawArchive (box bytes)).``then``(fun source ->
                    (ZipInterop.rawArchiveBytes source wanted).``then``(fun parts -> source, parts))
                |> unbox
            // (A file that cannot be read again is written from the model, as before.)
            (fetched.``catch``(fun _ -> null, [])).``then``(fun (source, parts) -> write source parts) |> unbox<JS.Promise<byte[]>>
        | None ->
            // (The workbook is taken now, as it is when asked for; the writer may still be loading.)
            let out = this.ToXlsxWorkbook()
            packageBytes (writeXlsxPackage null [] out)

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

    /// How many times a chart (or a group with charts) on display was laid out again because
    /// cells it reads changed, since the component was made. An edit that changes none of the
    /// cells a chart reads leaves it as it is drawn.
    member _.GetChartRedraws() : int = chartRedraws

    /// How many times a chart, shape or group of an opened file was laid out for the grid at
    /// all: once when its sheet is first shown, again when its size or the cells it reads
    /// change. Showing a sheet again, scrolling and edits elsewhere lay nothing out.
    member _.GetChartLayouts() : int = chartLayouts

    /// The active sheet's pictures and other drawing objects, back to front:
    /// { id, kind: "picture" | "chart" | "shape" | "other", name, description,
    ///   contentType (a picture's, else null), row, col (anchor cell),
    ///   x, y, width, height (sheet pixels at 100%) }.
    member _.GetDrawings() : obj array =
        model.Drawings
        |> List.map (fun drawing ->
            let rect = SpreadsheetLayout.drawingRect model drawing
            Host.drawingToObj drawing rect.X rect.Y rect.Width rect.Height)
        |> List.toArray

    /// The active sheet's pictures, back to front (the objects of GetDrawings
    /// whose kind is "picture"): { id, kind, name, description, contentType,
    /// row, col, x, y, width, height }. `id` names the picture for
    /// MovePicture, ResizePicture, DeletePicture and SelectPicture; it stays
    /// the same through undo and redo. A file's objects are numbered 1..n
    /// back to front when it is opened.
    member _.GetPictures() : obj array =
        model.Drawings
        |> List.filter (fun drawing -> match drawing.Content with Picture _ -> true | KeptObject _ -> false)
        |> List.map (fun drawing ->
            let rect = SpreadsheetLayout.drawingRect model drawing
            Host.drawingToObj drawing rect.X rect.Y rect.Width rect.Height)
        |> List.toArray

    /// Capture the workbook, sheet and selection for a chart insertion dialog.
    /// Passing this to InsertChart rejects a stale dialog instead of placing a
    /// chart on a different sheet or from another selection.
    member this.GetChartSelectionSnapshot() : obj =
        observeRevision ()
        if readOnly then null
        else createObj [
            "epoch" ==> session.Epoch
            "serial" ==> session.ChangeSerial
            "sheetIndex" ==> activeSheetIndex
            "selection" ==> this.GetSelectionRef() ]

    /// Insert a classic chart from a contiguous bounded range. Columns are
    /// series; firstRowHeaders / firstColumnCategories explicitly say which
    /// cells are labels. Pie requires exactly one remaining value column.
    /// Returns its drawing id; the anchor and chart part are one undo step.
    member this.InsertChart(range: obj, options: obj, snapshot: obj) : int =
        let memberName = "InsertChart"
        if readOnly then Host.hostFail memberName "the sheet is read-only"
        observeRevision ()
        if not (Host.isNullish snapshot) then
            if not (Host.isJsObject snapshot) then Host.hostFail memberName "the selection snapshot is invalid"
            if Host.getKey snapshot "epoch" <> box session.Epoch
               || Host.getKey snapshot "serial" <> box session.ChangeSerial
               || Host.getKey snapshot "sheetIndex" <> box activeSheetIndex
               || Host.getKey snapshot "selection" <> box (this.GetSelectionRef()) then
                Host.hostFail memberName "the workbook, sheet or selection changed; reopen the chart dialog"
        let area = Host.rangeArg memberName range
        if not (Host.isJsObject options) then Host.hostFail memberName "options must be { type, title, firstRowHeaders, firstColumnCategories }"
        for key in Host.keysOf options do
            if not (List.contains key [ "type"; "title"; "firstRowHeaders"; "firstColumnCategories" ]) then
                Host.hostFail memberName (sprintf "unknown option \"%s\"" key)
        let chartKind =
            match Host.stringArg memberName "options.type" (Host.getKey options "type") with
            | "column" -> SpreadsheetChartAuthoring.Column
            | "bar" -> SpreadsheetChartAuthoring.Bar
            | "line" -> SpreadsheetChartAuthoring.Line
            | "pie" -> SpreadsheetChartAuthoring.Pie
            | other -> Host.hostFail memberName (sprintf "unknown chart type \"%s\"" other)
        let title = Host.optionalArg (Host.stringArg memberName "options.title") (Host.getKey options "title") |> Option.defaultValue ""
        let headers = Host.boolArg memberName "options.firstRowHeaders" (Host.getKey options "firstRowHeaders")
        let categories = Host.boolArg memberName "options.firstColumnCategories" (Host.getKey options "firstColumnCategories")
        let reserved =
            match xlsxSource, model.SourceSheet with
            | Some source, Some origin when origin.Book = source.Id ->
                (List.item origin.Sheet source.Workbook.Sheets).Extras.DrawingRelationships
                |> List.map (fun (relId, _, _, _) -> relId)
            | _ -> []
        let result = executeDrawing memberName "Insert chart" (Rbt.Spreadsheet.Session.InsertChart (area, chartKind, title, headers, categories, reserved))
        let id = result.Id.Value
        render ()
        id

    /// Put a picture on the active sheet, in front of its other objects, with
    /// its top-left corner at a cell (one undo step), and return its id.
    /// `source`: the image file's bytes (Uint8Array or ArrayBuffer) or a data
    /// URL ("data:image/png;base64,..."); PNG, JPEG and GIF are taken (the
    /// kind is read off the bytes, not off a name or the URL's type).
    /// `anchorCell`: "B3" or { row, col }. `size` (optional): { width, height }
    /// in sheet pixels - with one of the two the other follows the picture's
    /// proportions, with neither the picture has its own pixel size.
    /// `name` (optional) is what GetPictures calls it ("Picture 3" otherwise).
    /// The picture moves with its cell and keeps its size. Throws for
    /// anything that is not a PNG, JPEG or GIF image, a cell outside the
    /// sheet, or a size that is not positive.
    member _.InsertPicture(source: obj, anchorCell: obj, size: obj, name: obj) : int =
        let memberName = "InsertPicture"
        let bytes = Host.pictureBytesArg memberName source
        let info =
            match SheetOps.imageInfo bytes with
            | Some info -> info
            | None -> Host.hostFail memberName "the picture is not a PNG, JPEG or GIF image"
        let cell = Host.cellArg memberName "anchorCell" model anchorCell
        let width, height = pictureSizeArg memberName size (float info.Width) (float info.Height)
        let name = Host.optionalArg (Host.stringArg memberName "name") name |> Option.filter (fun text -> text.Trim() <> "")
        let id = (executeDrawing memberName "Insert picture" (Rbt.Spreadsheet.Session.InsertPicture (bytes, cell, width, height, name))).Id.Value
        render()
        id

    /// Move a picture of the active sheet (one undo step): to a cell ("B3" or
    /// { row, col }: its top-left corner goes to the cell's), or to a point of
    /// the sheet ({ x, y } in sheet pixels at 100%, from the sheet's top-left
    /// corner; it is then anchored to the cell the point is in). The size
    /// stays. Throws when there is no picture with the id (a chart or another
    /// kept object cannot be moved) or the place is not one.
    member _.MovePicture(id: obj, place: obj) : unit =
        let memberName = "MovePicture"
        let drawing = pictureArg memberName id
        let rect = SpreadsheetLayout.drawingRect model drawing
        let x, y =
            if Host.isJsObject place && Host.hasKey place "x" && Host.hasKey place "y" then
                let x, y = Host.numberArg memberName "x" (Host.getKey place "x"), Host.numberArg memberName "y" (Host.getKey place "y")
                if x < 0.0 || y < 0.0 then Host.hostFail memberName "x and y cannot be negative"
                x, y
            else
                let cell = Host.cellArg memberName "the place" model place
                SpreadsheetLayout.columnLeft model cell.Col, SpreadsheetLayout.rowTop model cell.Row
        if placePictureAt memberName "Move picture" drawing.Id { rect with X = x; Y = y } then render()

    /// Give a picture of the active sheet a size (one undo step):
    /// { width, height } in sheet pixels; with one of the two the other
    /// follows the picture's present proportions. Its top-left corner stays.
    /// Throws when there is no picture with the id or the size is not positive.
    member _.ResizePicture(id: obj, size: obj) : unit =
        let memberName = "ResizePicture"
        let drawing = pictureArg memberName id
        let rect = SpreadsheetLayout.drawingRect model drawing
        if Host.isNullish size then Host.hostFail memberName "the size is missing ({ width, height })"
        let width, height = pictureSizeArg memberName size rect.Width rect.Height
        if placePictureAt memberName "Resize picture" drawing.Id { rect with Width = width; Height = height } then render()

    /// Take a picture off the active sheet (one undo step). Throws when there
    /// is no picture with the id.
    member _.DeletePicture(id: obj) : unit =
        let drawing = pictureArg "DeletePicture" id
        let result = executeDrawing "DeletePicture" "Delete picture" (Rbt.Spreadsheet.Session.DeletePicture drawing.Id)
        if selectedPicture = Some drawing.Id then selectedPicture <- None
        if result.Changed then render()

    /// Set the alternative text of a picture (its description; "" removes
    /// it) and, when given, its name, as one undo step.
    member _.SetPictureAltText(id: obj, description: obj, name: obj) : unit =
        let memberName = "SetPictureAltText"
        let drawing = pictureArg memberName id
        let description = Host.stringArg memberName "the description" description
        let name = Host.optionalArg (Host.stringArg memberName "name") name
        if (executeDrawing memberName "Picture text" (Rbt.Spreadsheet.Session.DescribePicture (drawing.Id, name, Some description))).Changed then render()

    /// The id of the picture that is selected on the active sheet, or null.
    /// A click on a picture selects it (not in read-only mode): it gets a
    /// frame with eight handles, is moved by dragging it and sized by
    /// dragging a handle (a corner keeps its proportions, Shift frees them);
    /// Delete removes it, Escape or a click on a cell lets go of it. Each
    /// drag is one undo step. OnSelectionChange tells: its payload has
    /// `Picture` (the id, or null).
    member _.GetSelectedPicture() : obj =
        match selectedPicture with
        | Some id when (SheetOps.drawingById id model).IsSome -> box id
        | _ -> null

    /// Select a picture of the active sheet by its id, or with null none.
    /// Throws when there is no picture with the id.
    member _.SelectPicture(id: obj) : unit =
        if Host.isNullish id then selectedPicture <- None
        else
            let drawing = pictureArg "SelectPicture" id
            if model.EditingCell.IsSome then finishEditing true |> ignore
            selectedPicture <- Some drawing.Id
        render()

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
            let displayText = SpreadsheetCellText.cellDisplayTextIn model.Date1904 measure rect.Width format value
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
                "rowHeight" ==> pxToPt (Map.tryFind addr.Row model.RowHeights |> Option.defaultValue model.DefaultRowHeight)
                "columnWidth" ==> pxToPt (Map.tryFind addr.Col model.ColumnWidths |> Option.defaultValue model.DefaultColumnWidth)
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
    ///   comment: { text, author: string | null, thread } | null (see GetComment),
    ///   hyperlink: { target, isInternal: boolean, range } | null,
    ///   validation: the rule object of the validation on the cell with its `id` and `range`
    ///               (see GetValidations / GetValidationAt) | null,
    ///   spill: { anchor: "A1", range: "A1:C5" } | null - the dynamic array the cell is part of
    ///          (for its formula cell and for the cells it filled; see GetSpillRange) }.
    /// Throws when row / col are not whole numbers inside the sheet.
    member _.GetCell(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetCell" model row; Col = Host.colArg "GetCell" model col }
        Host.cellToObj model addr (cellText (getCell addr model))

    /// The dynamic array a cell of the active sheet is part of:
    /// { anchor: "A1" (the formula cell), range: "A1:C5" (what it fills, the
    /// anchor included) }, or null for a cell that is neither a spilling
    /// formula nor a cell one filled (a formula that gives one value, or
    /// whose array is blocked - #SPILL! - has no range). Throws when row /
    /// col are not whole numbers inside the sheet.
    member _.GetSpillRange(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetSpillRange" model row; Col = Host.colArg "GetSpillRange" model col }
        Host.spillToObj model addr

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

    /// What a status bar shows of the selection: { sum, average, count,
    /// numericCount, min, max }. `count` is the number of cells that hold
    /// something, `numericCount` how many of them are numbers (dates are
    /// numbers; a formula counts by its result; text, booleans and errors
    /// are not numbers), and sum / average / min / max are of the numbers -
    /// null when there is none. Rows an AutoFilter hides are left out, as in
    /// Excel's status bar; rows hidden by hand are counted. A selection of
    /// whole rows, columns or the sheet costs one pass over the cells that
    /// exist, and asking again for the same selection of unchanged cells
    /// costs nothing. The same object is the `Status` property of the
    /// OnSelectionChange payload (worked out when it is first read). The
    /// selection event is not raised when the cells of an unchanged
    /// selection change: read GetStatus() from OnChange for that.
    member _.GetStatus() : obj = selectionStatusObj ()

    /// The used range of a sheet (the active one, or `sheetIndex`): the
    /// smallest range that holds every cell with a value, a format or a
    /// comment and every merged cell, as Excel's UsedRange (it need not
    /// start at A1) - { range: "B2:F40", firstRow, firstCol, lastRow,
    /// lastCol } (0-based), or null for an empty sheet. A sheet of an opened
    /// file that was not needed so far is converted by this call.
    member _.GetUsedRange(?sheetIndex: obj) : obj =
        let source = sheetArg "GetUsedRange" (defaultArg sheetIndex null)
        match SheetOps.usedExtent source with
        | None -> null
        | Some range ->
            createObj [ "range" ==> SheetOps.rangeText range
                        "firstRow" ==> range.From.Row; "firstCol" ==> range.From.Col
                        "lastRow" ==> range.To.Row; "lastCol" ==> range.To.Col ]

    /// The page setup of a sheet (the active one, or `sheetIndex`), for a
    /// print dialog:
    /// { orientation: "portrait" | "landscape",
    ///   paperSize: "letter" | "tabloid" | "legal" | "executive" | "a3" | "a4" | "a5" | "b5" | { code } (any other OOXML paper code),
    ///   pageWidth, pageHeight (points, of the paper in portrait; read only),
    ///   margins: { top, bottom, left, right, header, footer } (inches),
    ///   scale (percent, 10..400; null while the sheet is fitted to pages),
    ///   fitToWidth, fitToHeight (pages, 0 = as many as it takes; null while a scale applies),
    ///   printArea: "A1:F40" | null (one area),
    ///   printTitles: { rows: "1:2" | null, columns: "A:B" | null },
    ///   header, footer, firstHeader, firstFooter, evenHeader, evenFooter:
    ///     { left, center, right, raw } - the three sections and the whole
    ///     string, in Excel's codes (&P page, &N pages, &D date, &T time,
    ///     &F file, &A sheet, && an ampersand, &"Font,Style", &12 a size);
    ///     the first / even ones apply when differentFirst / differentOddEven,
    ///   differentFirst, differentOddEven, scaleHeaderFooter,
    ///   gridlines, headings, centerHorizontally, centerVertically,
    ///   pageOrder: "downThenOver" | "overThenDown", blackAndWhite,
    ///   firstPageNumber (number | null: automatic),
    ///   comments: "none" | "asDisplayed" | "atEnd" }.
    /// Manual page breaks are GetPageBreaks.
    member _.GetPageSetup(?sheetIndex: obj) : obj =
        Host.pageSetupToObj (sheetArg "GetPageSetup" (defaultArg sheetIndex null)).PageSetup

    /// The pages a sheet (the active one, or `sheetIndex`) prints on, as the
    /// PDF export lays them out (the same pagination: manual page breaks,
    /// the scale or the fit to pages, a print area of several areas, print
    /// titles and the row / column headings taken off every page) - for a
    /// print preview or a page-break view:
    /// { sheetIndex, sheetName,
    ///   setup: the object of GetPageSetup,
    ///   scale: the scale the sheet is printed at, in percent (the one that
    ///     fits it to its pages when it is fitted),
    ///   printAreas: ["A1:F40", ...] (what is printed: the print area's areas
    ///     in order, or the used range),
    ///   printTitles: { rows: { first, last } | null, columns: { first, last } | null } (0-based),
    ///   breaks: { rows: [...], columns: [...] } (the manual ones, as GetPageBreaks),
    ///   pageCount,
    ///   pages: [{ index (0-based), number (as &P prints it: firstPageNumber counts),
    ///     area (index into printAreas), range: "A1:F40",
    ///     firstRow, lastRow, firstCol, lastCol (0-based, the body of the
    ///     page: print titles ride along beside it),
    ///     header, footer: { left, center, right } | null - the texts as
    ///     printed, codes resolved (&P &N &D &T &A &F ...; the first / even
    ///     ones where the setup has them) }] }.
    /// `fileName` is what &F prints (default ""). A sheet of an opened file
    /// that was not needed so far is converted by this call.
    member _.GetPrintPages(?sheetIndex: obj, ?fileName: obj) : obj =
        let memberName = "GetPrintPages"
        let index =
            match Host.optionalArg (Host.intArg memberName "sheetIndex") (defaultArg sheetIndex null) with
            | None -> activeSheetIndex
            | Some index -> index
        let source = sheetArg memberName (defaultArg sheetIndex null)
        let file = Host.optionalArg (Host.stringArg memberName "fileName") (defaultArg fileName null) |> Option.defaultValue ""
        let setup = source.PageSetup
        let pagination = Rbt.Spreadsheet.Pdf.Pagination.paginatePrinted source
        let pages = List.toArray pagination.Pages
        let total = max 1 pages.Length
        let offset = if setup.UseFirstPageNumber then setup.FirstPageNumber - 1 else 0
        let now = System.DateTime.Now
        let sectionObj (raw: string option) (number: int) : obj =
            match raw with
            | None -> null
            | Some raw ->
                let context : Rbt.Spreadsheet.Pdf.HeaderFooter.PageContext =
                    { PageNumber = number + offset; TotalPages = total; SheetName = sheetNames.[index]; FileName = file; FilePath = ""; Now = now }
                let resolved = Rbt.Spreadsheet.Pdf.HeaderFooter.resolve context raw
                createObj [ "left" ==> resolved.Left.Text; "center" ==> resolved.Center.Text; "right" ==> resolved.Right.Text ]
        let span (range: (int * int) option) : obj =
            match range with
            | Some (first, last) -> createObj [ "first" ==> first; "last" ==> last ]
            | None -> null
        createObj [
            "sheetIndex" ==> index
            "sheetName" ==> sheetNames.[index]
            "setup" ==> Host.pageSetupToObj setup
            "scale" ==> System.Math.Round (pagination.Scale * 10000.0) / 100.0
            "printAreas" ==> (pagination.PrintAreas |> List.map SheetOps.rangeText |> List.toArray)
            "printTitles" ==> createObj [ "rows" ==> span pagination.RepeatRows; "columns" ==> span pagination.RepeatCols ]
            "breaks" ==> Host.pageBreaksToObj setup
            "pageCount" ==> pages.Length
            "pages" ==>
                (pages |> Array.mapi (fun at page ->
                    let header, footer =
                        let hf = setup.HeaderFooter
                        if hf.DifferentFirst && at = 0 then hf.FirstHeader, hf.FirstFooter
                        elif hf.DifferentOddEven && (at + 1) % 2 = 0 then hf.EvenHeader, hf.EvenFooter
                        else hf.OddHeader, hf.OddFooter
                    createObj [
                        "index" ==> at
                        "number" ==> at + 1 + offset
                        "area" ==> page.Area
                        "range" ==> SheetOps.rangeText { From = { Row = page.RowStart; Col = page.ColStart }; To = { Row = page.RowEnd; Col = page.ColEnd } }
                        "firstRow" ==> page.RowStart; "lastRow" ==> page.RowEnd
                        "firstCol" ==> page.ColStart; "lastCol" ==> page.ColEnd
                        "header" ==> sectionObj header (at + 1)
                        "footer" ==> sectionObj footer (at + 1) ]))
        ]

    /// The faces the PDF export sets text in (`ExportPdf`): [{ faceName, family, bold, italic, data }] -
    /// the TrueType bytes (Uint8Array / ArrayBuffer) of each face of a family; `faceName` is the face's
    /// name in the PDF (default: the family with -Bold / -Italic / -BoldItalic, without spaces). The
    /// `FontFace` records a host makes for Word's PDF (`new FontFace(faceName, family, bold, italic, data)`
    /// of word-pdf/Pdf.js, fields FaceName, Family, Bold, Italic, Data) are taken as they are: one set of
    /// faces for both editors. Without faces a PDF has no text. null clears them.
    member _.SetPdfFonts(fonts: obj) : unit =
        pdfFonts <- pdfFontsArg "SetPdfFonts" fonts

    /// Where the faces cells are drawn in come from: `loader(family, bold, italic)` resolves to the
    /// TrueType bytes of that face (a Uint8Array / ArrayBuffer), or null when the host has none. The
    /// grid asks for a face the first time a cell in it is drawn, when the browser has no font of
    /// that family (Calibri, Georgia, Segoe UI ... on Linux), and draws every cell again once it is
    /// loaded. Only the faces a workbook shows are fetched. null: no loader.
    member _.SetFontLoader(loader: obj) : unit =
        if not (Host.isNullish loader) && Host.jsTypeOf loader <> "function" then
            Host.hostFail "SetFontLoader" "the loader is a function (family, bold, italic) => Promise<Uint8Array | ArrayBuffer | null>"
        fontLoader <- (if Host.isNullish loader then null else loader)
        wantedFaces.Clear()
        renderedStyleInputs <- [||]
        render ()

    /// The PDF of the sheet on display, of its selection or of the whole workbook, as Excel prints it: each
    /// sheet by its page setup - print area, print titles, scaling or fitting to pages, orientation, paper,
    /// margins, centring, headers and footers, grid lines and headings, page order, manual breaks (the pages of
    /// `GetPrintPages`) - with its charts, shapes and pictures. `options` (all optional): { what: "sheet"
    /// (default) | "selection" (the cells selected, as Excel's Print Selection: the sheet's setup without
    /// its print area) | "workbook" (every visible worksheet, one after the other), sheetIndex (for
    /// "sheet"; default the one on display), fileName (what &F prints; default ""), defaultFont (the family
    /// of cells that name none; default "Calibri"), fonts (as SetPdfFonts, for this export only) }.
    /// Resolves to the PDF's bytes (Uint8Array). The PDF writer is loaded with the first export
    /// (dist/spreadsheet-pdf), not with the grid.
    member this.ExportPdf(?options: obj) : JS.Promise<byte[]> =
        let memberName = "ExportPdf"
        let options = defaultArg options null
        if not (Host.isNullish options) then
            if not (Host.isJsObject options) then Host.hostFail memberName "the options are { what, sheetIndex, fileName, defaultFont, fonts }"
            for key in Host.keysOf options do
                if not (List.contains key [ "what"; "sheetIndex"; "fileName"; "defaultFont"; "fonts" ]) then
                    Host.hostFail memberName (sprintf "the options have no \"%s\" (they are what, sheetIndex, fileName, defaultFont, fonts)" key)
        let get (key: string) = if Host.isNullish options then null else Host.getKey options key
        let what = Host.optionalArg (Host.stringArg memberName "what") (get "what") |> Option.defaultValue "sheet"
        let fileName = Host.optionalArg (Host.stringArg memberName "fileName") (get "fileName") |> Option.defaultValue ""
        let defaultFont = Host.optionalArg (Host.stringArg memberName "defaultFont") (get "defaultFont") |> Option.defaultValue "Calibri"
        let fonts = if Host.isNullish (get "fonts") then pdfFonts else pdfFontsArg memberName (get "fonts")
        if model.EditingCell.IsSome then finishEditing true |> ignore
        // The sheets, each with the charts and shapes drawn over its cells.
        let withOverlays (index: int) (sheet: SpreadsheetModel) =
            let sceneOf (xml: string) (width: float) (height: float) =
                try
                    drawingPartsOf sheet
                    |> Option.bind (fun parts -> SpreadsheetCharts.objectScene parts (chartDataReading (ResizeArray())) xml width height)
                with error ->
                    reportError "drawing" (box error)
                    None
            sheet, sheetNames.[index], SpreadsheetPaint.drawingOverlays sceneOf (Rbt.Spreadsheet.Pdf.Pagination.withPrintSizes sheet)
        let chosen =
            match what with
            | "sheet" ->
                let index =
                    match Host.optionalArg (Host.intArg memberName "sheetIndex") (get "sheetIndex") with
                    | None -> activeSheetIndex
                    | Some index when index >= 0 && index < sheets.Length -> index
                    | Some index -> Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
                [ withOverlays index (sheetModelAt index) ]
            | "selection" ->
                let r0, c0, r1, c1 =
                    match model.Selection with
                    | SingleCell (r, c) -> r, c, r, c
                    | Range (sr, sc, er, ec) -> min sr er, min sc ec, max sr er, max sc ec
                let area = { From = { Row = r0; Col = c0 }; To = { Row = r1; Col = c1 } }
                let selected = { model with PageSetup = { model.PageSetup with PrintArea = Some area; ExtraPrintAreas = [] } }
                [ withOverlays activeSheetIndex selected ]
            | "workbook" ->
                let visible = this.SheetVisibility()
                [ for index in 0 .. sheets.Length - 1 do
                    if visible.[index] && sheetKindAt index = "worksheet" then yield withOverlays index (sheetModelAt index) ]
            | other -> Host.hostFail memberName (sprintf "\"%s\" is not what a PDF is made of (\"sheet\", \"selection\" or \"workbook\")" other)
        (importSheetPdf ()).``then``(fun (pdf: SheetPdfModule) ->
            let sheetList = chosen |> List.map (fun (sheet, name, overlays) -> pdf.pdfSheet (sheet, name, overlays))
            let pictures = pdf.pngPictures sheetList
            let converted = allPromises (pictures |> Array.map pngAsJpeg)
            converted.``then``(fun (jpegs: byte[] array) ->
                let replace (data: byte[]) =
                    match pictures |> Array.tryFindIndex (fun known -> obj.ReferenceEquals(known, data)) with
                    | Some at when not (isNull jpegs.[at]) -> Some ("image/jpeg", jpegs.[at])
                    | _ -> None
                let sheetList = if pictures.Length = 0 then sheetList else pdf.withPictures (replace, sheetList)
                let faces = fonts |> Array.map (fun (faceName, family, bold, italic, data) -> pdf.fontFace (faceName, family, bold, italic, data))
                pdf.exportSheetsWithDefaultFonts (defaultFont, fileName, faces, List.toArray sheetList)))
        |> unbox<JS.Promise<byte[]>>

    /// Print through the browser: the PDF of `ExportPdf(options)` in a hidden frame of the page, with the
    /// browser's print dialog (its pages are the PDF's, as Excel prints them). Resolves once the dialog was
    /// opened.
    member this.PrintPdf(?options: obj) : JS.Promise<unit> =
        (this.ExportPdf(defaultArg options null)).``then``(fun (pdf: byte[]) -> printPdfInFrame pdf) |> unbox<JS.Promise<unit>>

    // ---- macros ------------------------------------------------------------------------------------

    /// Whether the opened workbook has VBA macros (known without loading them).
    member _.HasMacros = macroPreserved.IsSome || (macroProject |> Option.exists (fun p -> not p.Modules.IsEmpty))

    /// Load the workbook's macros, with the VBA engine (fetched the first time in the browser). The
    /// members that read them (GetMacros, GetMacroModules, GetMacroProblems) answer from what is loaded;
    /// those that run them load first.
    member _.LoadMacros() : JS.Promise<unit> =
        let loaded =
            match macroProject, macroProjectLoad with
            | Some project, _ -> Fable.Core.JS.Constructors.Promise.resolve project
            | None, Some pending -> pending
            | None, None ->
                let preserved = macroPreserved
                let pending =
                    (loadMacroProject preserved).``then``(fun (project: Rbt.Spreadsheet.MacroApi.IMacroProject) ->
                        // (A newer workbook installed meanwhile has its own.)
                        if obj.ReferenceEquals(preserved, macroPreserved) && macroProject.IsNone then macroProject <- Some project
                        project)
                macroProjectLoad <- Some pending
                pending
        loaded.``then``(fun _ -> ())

    /// The macros a user can run (Excel's Macros dialog): the public Subs without required arguments, in
    /// order; those of the workbook's and sheets' own modules as `Module.Name`. Empty until LoadMacros.
    member _.GetMacros() : string array =
        match macroProject with
        | Some project -> List.toArray project.Macros
        | None -> [||]

    /// The modules of the workbook's VBA project: [{ name, kind ("standard" | "class" | "document"), source,
    /// code }] (`code` without the Attribute lines). Empty until LoadMacros.
    member _.GetMacroModules() : obj array =
        match macroProject with
        | Some project ->
            [| for m in project.Modules -> createObj [ "name" ==> m.Name; "kind" ==> m.Kind; "source" ==> m.Source; "code" ==> m.Code ] |]
        | None -> [||]

    /// What keeps the project from compiling: [{ module, line, column, description }] (none when it compiles).
    member _.GetMacroProblems() : obj array =
        match macroProject with
        | Some project ->
            [| for f in project.Problems -> createObj [ "module" ==> f.Module; "line" ==> f.Line; "column" ==> f.Column; "description" ==> f.Description ] |]
        | None -> [||]

    /// Whether macros may run (`EnableMacros`). Off for every workbook opened, as Excel's "Enable content".
    member _.MacrosEnabled = macrosEnabled

    /// Let the workbook's macros run (loading them); runs its Workbook_Open (of ThisWorkbook) or Auto_Open
    /// once, as opening the workbook in Excel with content enabled would. Resolves to its result, as RunMacro's.
    member this.EnableMacros() : JS.Promise<obj> =
        macrosEnabled <- true
        (this.LoadMacros()).``then``(fun () ->
            if workbookOpened then createObj [ "ok" ==> true ]
            else
                workbookOpened <- true
                match macroProject with
                | Some project when project.Has (workbookCodeName + ".Workbook_Open") -> this.RunLoaded(workbookCodeName + ".Workbook_Open", (fun project start hooks -> project.Run(start, hooks, workbookCodeName + ".Workbook_Open")), "Workbook_Open")
                | Some project when project.Has "Auto_Open" -> this.RunLoaded("Auto_Open", (fun project start hooks -> project.Run(start, hooks, "Auto_Open")), "Auto_Open")
                | _ -> createObj [ "ok" ==> true ])

    /// Where MsgBox, InputBox and Debug.Print go: { messageBox(prompt, buttons, title) -> VBA's answer
    /// (vbOK 1, vbCancel 2, vbYes 6, vbNo 7 ...), inputBox(prompt, title, default) -> text or null,
    /// output(text), wait(ms) -> false to go on at once from an Application.Wait of so many milliseconds
    /// (a test or a recording that does not wait) }. Missing members keep the browser's dialogs, the
    /// console and waiting; null restores them.
    member _.SetMacroHooks(hooks: obj) =
        if Host.isNullish hooks then
            macroHooks <- browserMacroHooks
            macroWait <- None
        else
            let fn (key: string) = let f = Host.getKey hooks key in if Host.isNullish f then None else Some f
            macroWait <- fn "wait" |> Option.map (fun f -> fun (ms: float) -> let r = callHook3 f (box ms) null null in Host.isNullish r || unbox<bool> r)
            macroHooks <-
                { browserMacroHooks with
                    MessageBox = (match fn "messageBox" with Some f -> (fun p b t -> let r = callHook3 f (box p) (box b) (box t) in if Host.isNullish r then 1 else unbox<int> r) | None -> browserMacroHooks.MessageBox)
                    InputBox = (match fn "inputBox" with Some f -> (fun p t d -> let r = callHook3 f (box p) (box t) (box d) in if Host.isNullish r then None else Some (string r)) | None -> browserMacroHooks.InputBox)
                    Output = (match fn "output" with Some f -> (fun text -> callHook1 f (box text)) | None -> browserMacroHooks.Output) }

    /// Run a macro of the workbook by name (one of GetMacros, or any Sub, as `Name` or `Module.Name`). What
    /// it changes is one undo step; nothing of a macro that stops on an error is kept. Resolves to
    /// { ok: true } or { ok: false, number, description, module, procedure, line } (Excel's Err.Number and
    /// Err.Description, and where it stopped); once the macros are loaded it runs at once. Throws when
    /// macros are not enabled.
    ///
    /// The macro runs a slice at a time: between slices the grid shows what it has done so far (unless it
    /// switched ScreenUpdating off), Application.Wait waits, and StopMacro stops it. Meanwhile the grid is
    /// read-only; its history gets the one step when the macro ends.
    member this.RunMacro(name: string) : JS.Promise<obj> =
        if not macrosEnabled then Host.hostFail "RunMacro" "macros are not enabled (EnableMacros)"
        if macroRunning.IsSome then Host.hostFail "RunMacro" "a macro is running (StopMacro)"
        match macroProject with
        | Some _ -> this.RunStepped(name)
        | None -> (this.LoadMacros()).``then``(fun () -> this.RunStepped(name)) |> unbox<JS.Promise<obj>>

    /// Whether a macro is running (RunMacro).
    member _.MacroRunning = macroRunning.IsSome

    /// Stop the macro that is running (also at an Application.Wait): it ends as one stopped by Esc in Excel
    /// (error 18), and nothing of it is kept.
    member _.StopMacro() = macroRunning |> Option.iter (fun run -> run.EndWait(); run.Cancel())

    member private this.RunStepped(name: string) : JS.Promise<obj> =
        match macroProject with
        | None -> Fable.Core.JS.Constructors.Promise.resolve (createObj [ "ok" ==> false; "number" ==> 0; "description" ==> "The macros are not loaded" ])
        | Some project ->
            if model.EditingCell.IsSome then finishEditing true |> ignore
            observeRevision ()
            let start = currentWorkbookSnapshot ()
            let hooks = { macroHooks with OnTime = scheduleMacro; OnKey = bindMacroKey; Output = (fun text -> macroLog.Add text; macroHooks.Output text) }
            let run = project.Start(start, hooks, name)
            macroRunning <- Some run
            refreshSheetLock ()
            let shown = ref false
            Fable.Core.JS.Constructors.Promise.Create(fun resolve _ ->
                let finish (outcome: Rbt.Spreadsheet.MacroApi.MacroOutcome) =
                    macroRunning <- None
                    match outcome with
                    | Rbt.Spreadsheet.MacroApi.MacroDone (result, selection) ->
                        if not (obj.ReferenceEquals(result, start)) then
                            session.CommitApplied({ Description = "Run macro " + name; Workbook = result })
                        installBook session.Workbook
                        match selection with
                        | Some (sheet, top, left, bottom, right) when sheet = sheetIdForIndex activeSheetIndex ->
                            selectFromTo { Row = top; Col = left } { Row = bottom; Col = right }
                        | _ -> ()
                        render()
                        resolve (createObj [ "ok" ==> true ])
                    | Rbt.Spreadsheet.MacroApi.MacroFailed f ->
                        // (Nothing of it is kept: the grid shows the workbook as it was. The frames shown may
                        // have reached the session's view - a selection, a scroll -, so it is the start again too.)
                        if shown.Value then
                            installBook start
                            session.UpdateView start
                        refreshSheetLock ()
                        render()
                        macroLog.Add f.Text
                        resolve (createObj [ "ok" ==> false; "number" ==> f.Number; "description" ==> f.Description; "module" ==> f.Module; "procedure" ==> f.Procedure; "line" ==> f.Line ])
                // (Showing a frame costs time the macro does not get: the slices run three times as long as
                // the last frame took to show, and at least 30 ms - about 25 frames a second at most.)
                let paintMs = ref 10.0
                // (A Wait the host's wait hook was asked about: it is asked once for each.)
                let asked = ref false
                let waiting () =
                    let ms = run.WaitMs()
                    if ms <= 0.0 then asked.Value <- false; false
                    elif asked.Value then true
                    else
                        asked.Value <- true
                        match macroWait with
                        | Some wait when not (wait ms) -> run.EndWait(); asked.Value <- false; false
                        | _ -> true
                let rec pump () =
                    // In steps small enough to notice a Wait soon after it.
                    let sliceEnd = nowMs () + max 30.0 (3.0 * paintMs.Value)
                    let mutable outcome = None
                    while outcome.IsNone && nowMs () < sliceEnd && not (waiting ()) do
                        outcome <- run.Step 4000
                    match outcome with
                    | Some ended -> finish ended
                    | None ->
                        let painting = nowMs ()
                        match run.Frame() with
                        | Some frame ->
                            installBook frame
                            shown.Value <- true
                            render()
                            paintMs.Value <- nowMs () - painting
                        | None -> ()
                        match run.WaitMs() with
                        | ms when ms > 0.0 -> setTimeoutMs pump (int (min 1000.0 ms)) |> ignore
                        | _ -> yieldThen pump
                pump ())

    /// Run a line as the VBA editor's Immediate window does (`?expression`, `Print expression`, or
    /// statements) on the workbook: what it changes is one undo step. Resolves as RunMacro.
    member this.RunMacroLine(line: string) : JS.Promise<obj> =
        if not macrosEnabled then Host.hostFail "RunMacroLine" "macros are not enabled (EnableMacros)"
        macroLog.Add("> " + line)
        this.RunWhenLoaded(line, (fun project start hooks -> project.Immediate(start, hooks, line)), "Immediate")

    member private this.RunWhenLoaded(label: string, run: Rbt.Spreadsheet.MacroApi.IMacroProject -> Workbook -> Rbt.Spreadsheet.MacroApi.MacroHooks -> Rbt.Spreadsheet.MacroApi.MacroOutcome, description: string) : JS.Promise<obj> =
        if not macrosEnabled then Host.hostFail "RunMacro" "macros are not enabled (EnableMacros)"
        match macroProject with
        | Some _ -> Fable.Core.JS.Constructors.Promise.resolve (this.RunLoaded(label, run, description))
        | None -> (this.LoadMacros()).``then``(fun () -> this.RunLoaded(label, run, description))

    member private this.RunLoaded(label: string, run: Rbt.Spreadsheet.MacroApi.IMacroProject -> Workbook -> Rbt.Spreadsheet.MacroApi.MacroHooks -> Rbt.Spreadsheet.MacroApi.MacroOutcome, description: string) : obj =
        let failed (f: Rbt.Spreadsheet.MacroApi.MacroFailure) =
            macroLog.Add f.Text
            createObj [ "ok" ==> false; "number" ==> f.Number; "description" ==> f.Description; "module" ==> f.Module; "procedure" ==> f.Procedure; "line" ==> f.Line ]
        match macroProject with
        | None -> createObj [ "ok" ==> false; "number" ==> 0; "description" ==> "The macros are not loaded" ]
        | Some project ->
            if model.EditingCell.IsSome then finishEditing true |> ignore
            observeRevision ()
            let start = currentWorkbookSnapshot ()
            let hooks = { macroHooks with OnTime = scheduleMacro; OnKey = bindMacroKey; Output = (fun text -> macroLog.Add text; macroHooks.Output text) }
            match run project start hooks with
            | Rbt.Spreadsheet.MacroApi.MacroDone (result, selection) ->
                if not (obj.ReferenceEquals(result, start)) then
                    session.CommitApplied({ Description = description; Workbook = result })
                    installBook session.Workbook
                match selection with
                | Some (sheet, top, left, bottom, right) when sheet = sheetIdForIndex activeSheetIndex ->
                    selectFromTo { Row = top; Col = left } { Row = bottom; Col = right }
                | _ -> ()
                render()
                createObj [ "ok" ==> true ]
            | Rbt.Spreadsheet.MacroApi.MacroFailed f -> failed f

    /// Replace the code of a module of the workbook's project (the code as the VBA editor shows it, without
    /// its Attribute lines; a module it does not have is added, as a standard module) and compile the
    /// project again; its macros run from the new code, and a save writes the project from the modules.
    /// Resolves to the problems, as GetMacroProblems.
    member this.SetMacroModule(name: string, code: string) : JS.Promise<obj array> =
        if System.String.IsNullOrWhiteSpace name then Host.hostFail "SetMacroModule" "the module needs a name"
        (this.LoadMacros()).``then``(fun () ->
            match macroProject with
            | None -> [||]
            | Some project ->
                let edited = project.WithModule(name, code)
                macroProject <- Some edited
                macroProjectEdited <- true
                let problems = this.GetMacroProblems()
                macroLog.Add(if problems.Length = 0 then sprintf "%s: compiled" name else sprintf "%s: %d problem(s)" name problems.Length)
                for f in edited.Problems do macroLog.Add f.Text
                problems)

    /// What macros printed and how the VBA editor's runs ended, newest last.
    member _.GetMacroOutput() : string array = macroLog.ToArray()

    /// Change the page setup of a sheet (the active one, or `sheetIndex`):
    /// `changes` is the object of GetPageSetup with only the keys to change
    /// (the whole object may be handed back). One undo step. `scale` switches
    /// fitting to pages off and `fitToWidth` / `fitToHeight` switch it on
    /// (both in one call throws); `margins` may name some sides; a header or
    /// footer is a string (the raw text with its codes), null (none) or
    /// { left?, center?, right? } (sections that are not given stay);
    /// `printArea` and the print titles are A1 references or null;
    /// `paperSize` a name, a code or { code }. `pageWidth` / `pageHeight` are
    /// ignored. Throws for an unknown key or a value Excel does not take.
    /// The print area and titles follow inserted and deleted rows / columns.
    member _.SetPageSetup(changes: obj, ?sheetIndex: obj) : unit =
        let memberName = "SetPageSetup"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let changes = requiredArg memberName "the changes" changes
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Page setup" (fun setup ->
                SheetOps.applyPageSetupChanges memberName (Host.pageSetupChangesOfObj memberName setup changes) setup)
        if changed then render()

    /// The manual page breaks of a sheet (the active one, or `sheetIndex`):
    /// { rows: number[], columns: number[] } - the 0-based rows / columns
    /// BEFORE which a break sits (rows: [20] starts a new page at row 20,
    /// the 21st), in ascending order.
    member _.GetPageBreaks(?sheetIndex: obj) : obj =
        Host.pageBreaksToObj (sheetArg "GetPageBreaks" (defaultArg sheetIndex null)).PageSetup

    /// Replace the manual page breaks of a sheet with { rows?, columns? }
    /// (as GetPageBreaks gives them; a list that is not given stays). One
    /// undo step. Throws for a break before row / column 0 or outside the sheet.
    member _.SetPageBreaks(breaks: obj, ?sheetIndex: obj) : unit =
        let memberName = "SetPageBreaks"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let breaks = requiredArg memberName "the breaks" breaks
        if not (Host.isJsObject breaks) then Host.hostFail memberName "the breaks must be { rows, columns }"
        for key in Host.keysOf breaks do
            if key <> "rows" && key <> "columns" then Host.hostFail memberName (sprintf "page breaks have no \"%s\" (they have rows, columns)" key)
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Page breaks" (fun setup ->
                let rows = if Host.hasKey breaks "rows" then Host.intListArg memberName "rows" (Host.getKey breaks "rows") else setup.ManualRowBreaks
                let cols = if Host.hasKey breaks "columns" then Host.intListArg memberName "columns" (Host.getKey breaks "columns") else setup.ManualColBreaks
                SheetOps.setPageBreaks memberName rows cols setup)
        if changed then render()

    /// Add a manual page break before row `row` and / or before column `col`
    /// (0-based; null for none of that kind) on a sheet. One undo step.
    /// False when the breaks were there already. Throws when neither is given.
    member _.AddPageBreak(row: obj, ?col: obj, ?sheetIndex: obj) : bool =
        let memberName = "AddPageBreak"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let row = Host.optionalArg (Host.intArg memberName "row") row
        let col = Host.optionalArg (Host.intArg memberName "col") (defaultArg col null)
        if row.IsNone && col.IsNone then Host.hostFail memberName "give a row, a column or both"
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Page break" (fun setup ->
                SheetOps.setPageBreaks memberName (Option.toList row @ setup.ManualRowBreaks) (Option.toList col @ setup.ManualColBreaks) setup)
        if changed then render()
        changed

    /// Remove the manual page break before row `row` and / or before column
    /// `col` (0-based; null for none of that kind). One undo step. False
    /// when there was no such break.
    member _.RemovePageBreak(row: obj, ?col: obj, ?sheetIndex: obj) : bool =
        let memberName = "RemovePageBreak"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let row = Host.optionalArg (Host.intArg memberName "row") row
        let col = Host.optionalArg (Host.intArg memberName "col") (defaultArg col null)
        if row.IsNone && col.IsNone then Host.hostFail memberName "give a row, a column or both"
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Page break" (fun setup ->
                { setup with
                    ManualRowBreaks = (match row with Some row -> setup.ManualRowBreaks |> List.filter (fun known -> known <> row) | None -> setup.ManualRowBreaks)
                    ManualColBreaks = (match col with Some col -> setup.ManualColBreaks |> List.filter (fun known -> known <> col) | None -> setup.ManualColBreaks) })
        if changed then render()
        changed

    /// Remove every manual page break of a sheet. One undo step. Returns
    /// how many there were.
    member _.ClearPageBreaks(?sheetIndex: obj) : int =
        let memberName = "ClearPageBreaks"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let mutable count = 0
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Clear page breaks" (fun setup ->
                count <- setup.ManualRowBreaks.Length + setup.ManualColBreaks.Length
                if count = 0 then setup else { setup with ManualRowBreaks = []; ManualColBreaks = [] })
        if changed then render()
        count

    /// Call `callback` when the user asks for a context menu: a right click
    /// in the grid, or the menu key / Shift+F10 (for the active cell or the
    /// selected picture). It receives { target: "cell" | "rowHeader" |
    /// "colHeader" | "corner" | "picture" | "sheetTab", row, col (0-based;
    /// -1 where the target has none: a column header has no row, the corner
    /// neither; for a picture its anchor cell), sheetIndex, clientX, clientY
    /// (where to open the menu: the pointer, or the middle of the cell for
    /// the keyboard), pictureId (number | null), inSelection (whether the
    /// target was part of the selection already) }. As in Excel, a right
    /// click outside the selection first selects what it is on - the cell,
    /// the whole row or column of a header, the picture - and one inside the
    /// selection keeps it; OnSelectionChange is told before the callback
    /// runs. While a subscriber is registered the browser's own menu is held
    /// back in the grid (not in the edit box of a cell); the component has no
    /// menu of its own, and without a subscriber nothing changes. "sheetTab"
    /// comes from RaiseSheetTabContextMenu only. Returns a function that
    /// removes the subscription.
    member _.OnContextMenu(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        contextListeners.Add listener
        System.Action(fun () -> contextListeners.Remove listener |> ignore)

    /// Subscribe to the AutoFilter buttons: the callback gets { column, row,
    /// sheetIndex, clientRect: { left, top, width, height }, active, sorted }
    /// when a button in a header cell is pressed, or Alt+Down is typed on its
    /// cell. While a subscriber is registered the grid's own checklist does
    /// not open (the host shows its menu) and the legacy button beside the
    /// column letters is hidden; the buttons in the cells stay and show
    /// whether the column is filtered / sorted. Returns a function that
    /// removes the subscription.
    member _.OnFilterButton(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        filterButtonListeners.Add listener
        // (The buttons of a viewer and the legacy header buttons depend on it.)
        let dropdown = getElementById "filter-dropdown"
        if not (isNull dropdown) then closePopup dropdown
        render ()
        System.Action(fun () ->
            if filterButtonListeners.Remove listener then render ())

    /// The filter buttons that can be seen now (not scrolled away, not under
    /// a frozen pane), each as the object OnFilterButton gives.
    member _.GetFilterButtonRects() : obj array =
        match gridContainer with
        | Some grid when filterButtonNodes.Count > 0 ->
            let zoom = appliedZoom
            let gridRect: float array = boundingRect grid
            let metrics = SpreadsheetLayout.metricsOf model
            let frozenRows = max 0 (min metrics.RowCount model.FrozenRows)
            let frozenCols = max 0 (min metrics.ColCount model.FrozenCols)
            let right = gridRect.[0] + gridClientWidth grid
            let bottom = gridRect.[1] + gridClientHeight grid
            // Where the scrolled cells begin, on each axis.
            let scrolledLeft = gridRect.[0] + (rowHeaderWidth + metrics.ColPrefix.[frozenCols]) * zoom
            let scrolledTop = gridRect.[1] + (headerHeight + metrics.RowPrefix.[frozenRows]) * zoom
            [| for node in filterButtonNodes do
                let payload = filterButtonPayload node
                let row: int = payload?row
                let col: int = payload?column
                let rect: float array = boundingRect node
                let centerX = rect.[0] + rect.[2] / 2.0
                let centerY = rect.[1] + rect.[3] / 2.0
                let left = if col < frozenCols then gridRect.[0] else scrolledLeft
                let top = if row < frozenRows then gridRect.[1] else scrolledTop
                if centerX >= left && centerX <= right && centerY >= top && centerY <= bottom then yield payload |]
        | _ -> [||]

    /// For a tab strip (the component draws none): tell the OnContextMenu
    /// subscribers of a right click on the tab of sheet `sheetIndex` at
    /// (clientX, clientY) - { target: "sheetTab", row: -1, col: -1,
    /// sheetIndex, clientX, clientY, pictureId: null, inSelection (the sheet
    /// is the one on display) }. The sheet on display is not changed. True
    /// when a subscriber was told. Throws for a sheet that is not there.
    member _.RaiseSheetTabContextMenu(sheetIndex: obj, clientX: obj, clientY: obj) : bool =
        let memberName = "RaiseSheetTabContextMenu"
        let index = Host.intArg memberName "sheetIndex" sheetIndex
        if index < 0 || index >= sheets.Length then Host.hostFail memberName (sprintf "there is no sheet %d (the workbook has %d)" index sheets.Length)
        let x = Host.numberArg memberName "clientX" clientX
        let y = Host.numberArg memberName "clientY" clientY
        raiseContextMenu "sheetTab" -1 -1 None (index = activeSheetIndex) x y index

    /// Call `callback` with errors the component catches instead of letting
    /// them reach the host: { source, message, error }. `source` says where:
    /// the name of a subscription whose callback threw ("OnChange",
    /// "OnSelectionChange", "OnSheetsChange", "OnSheetChanged",
    /// "OnEditStateChange", "OnHyperlinkActivate", "OnContextMenu"),
    /// "event:<name>" for an error in the grid's handling of a DOM event
    /// (a key, the pointer, scrolling, paste: it is reported and then thrown
    /// on, so the browser's console shows it as before), "picture" (a
    /// picture drag that could not be applied) or "ImportXlsx" (the file
    /// could not be read: the promise is rejected as well). What a member
    /// called by the host throws is thrown to the caller and is not reported
    /// here. An error thrown by an OnError callback is dropped. Returns a
    /// function that removes the subscription.
    member _.OnError(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        errorListeners.Add listener
        System.Action(fun () -> errorListeners.Remove listener |> ignore)

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
    /// checked against data validation and a table beside the cell does not
    /// grow over it. One undo step, like every other change of the workbook.
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
    /// Commit without dropping an entry rejected by a stop validation rule.
    /// False leaves the edit buffer visible for correction and does not save.
    member _.TryCommitEdit() : bool = finishEditingWith true true


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
            updateFormulaHelp ()
    /// A snapshot of the active entry and its caret for an external insert
    /// chooser. A later insert rejects changed text, selection or workbook.
    member _.GetEditSnapshot() : obj =
        if disposed || readOnly then null else
        let active = activeCellAddress()
        let text, start, finish =
            match model.EditingCell, layers with
            | Some _, Some layer ->
                model.EditBuffer,
                unbox<int> (getProp layer.EditInput "selectionStart"),
                unbox<int> (getProp layer.EditInput "selectionEnd")
            | Some _, None -> model.EditBuffer, model.EditBuffer.Length, model.EditBuffer.Length
            | None, _ ->
                let initial = if Map.containsKey active model.Spill then "" else SpreadsheetInput.editTextFor model.Date1904 (getCell active model)
                initial, initial.Length, initial.Length
        createObj [ "Text" ==> text; "Start" ==> start; "End" ==> finish
                    "Row" ==> active.Row; "Col" ==> active.Col; "Sheet" ==> activeSheetIndex
                    "Epoch" ==> session.Epoch; "Serial" ==> session.ChangeSerial ]

    /// Insert into an unchanged active edit without committing it. False
    /// means the chooser was stale; no pending input or cell is discarded.
    member _.InsertFunctionAt(name: string, snapshot: obj) : bool =
        match findFunctionHelp name with
        | None -> false
        | Some info when disposed || readOnly || isNull snapshot -> false
        | Some info ->
            let active = activeCellAddress()
            let expected = unbox<string> (getProp snapshot "Text")
            let first = unbox<int> (getProp snapshot "Start")
            let last = unbox<int> (getProp snapshot "End")
            let current =
                match model.EditingCell with
                | Some _ -> model.EditBuffer
                | None -> if Map.containsKey active model.Spill then "" else SpreadsheetInput.editTextFor model.Date1904 (getCell active model)
            if unbox<int> (getProp snapshot "Epoch") <> session.Epoch ||
               unbox<int> (getProp snapshot "Serial") <> session.ChangeSerial ||
               unbox<int> (getProp snapshot "Sheet") <> activeSheetIndex ||
               unbox<int> (getProp snapshot "Row") <> active.Row ||
               unbox<int> (getProp snapshot "Col") <> active.Col ||
               first < 0 || last < first || last > current.Length || current <> expected then false
            else
                if model.EditingCell.IsNone then startEditing active.Row active.Col
                let prefix = if current.StartsWith("=") then "" else "="
                let inserted = info.Name + "()"
                let value = prefix + current.Substring(0, first) + inserted + current.Substring(last)
                model <- { model with EditBuffer = value }
                match layers with
                | Some layer ->
                    setValue layer.EditInput value
                    focus layer.EditInput
                    let position = prefix.Length + first + inserted.Length - 1
                    setProp layer.EditInput "selectionStart" (box position)
                    setProp layer.EditInput "selectionEnd" (box position)
                | None -> ()
                notifyEditState ()
                updateFormulaHelp ()
                true

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

    /// Whether the workbook session has a step to undo or redo.
    member _.CanUndo() : bool = session.CanUndo
    member _.CanRedo() : bool = session.CanRedo

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
    member _.ApplyFormat(formatType: string) =
        let command, description =
            match formatType with
            | "bold" -> Rbt.Spreadsheet.Session.ToggleBold, "Bold"
            | "italic" -> Rbt.Spreadsheet.Session.ToggleItalic, "Italic"
            | "underline" -> Rbt.Spreadsheet.Session.ToggleUnderline, "Underline"
            | "doubleUnderline" -> Rbt.Spreadsheet.Session.ToggleDoubleUnderline, "Double underline"
            | "strikethrough" -> Rbt.Spreadsheet.Session.ToggleStrikethrough, "Strikethrough"
            | "wrap" -> Rbt.Spreadsheet.Session.ToggleWrapText, "Wrap text"
            | "shrinkToFit" -> Rbt.Spreadsheet.Session.ToggleShrinkToFit, "Shrink to fit"
            | "fontIncrease" -> Rbt.Spreadsheet.Session.ChangeFontSize 2, "Font size"
            | "fontDecrease" -> Rbt.Spreadsheet.Session.ChangeFontSize -2, "Font size"
            | "indentIncrease" -> Rbt.Spreadsheet.Session.ChangeIndent 1, "Increase indent"
            | "indentDecrease" -> Rbt.Spreadsheet.Session.ChangeIndent -1, "Decrease indent"
            | "decimalIncrease" -> Rbt.Spreadsheet.Session.ChangeDecimals 1, "Increase decimals"
            | "decimalDecrease" -> Rbt.Spreadsheet.Session.ChangeDecimals -1, "Decrease decimals"
            | "alignLeft" -> Rbt.Spreadsheet.Session.HorizontalAlign (Some "left"), "Align left"
            | "alignCenter" -> Rbt.Spreadsheet.Session.HorizontalAlign (Some "center"), "Align center"
            | "alignRight" -> Rbt.Spreadsheet.Session.HorizontalAlign (Some "right"), "Align right"
            | "alignJustify" -> Rbt.Spreadsheet.Session.HorizontalAlign (Some "justify"), "Justify"
            | "alignGeneral" -> Rbt.Spreadsheet.Session.HorizontalAlign None, "General alignment"
            | "alignTop" -> Rbt.Spreadsheet.Session.VerticalAlign (Some "top"), "Align top"
            | "alignMiddle" -> Rbt.Spreadsheet.Session.VerticalAlign (Some "middle"), "Align middle"
            | "alignBottom" -> Rbt.Spreadsheet.Session.VerticalAlign (Some "bottom"), "Align bottom"
            | other -> Host.hostFail "ApplyFormat" (sprintf "unknown command \"%s\"" (if isNull other then "null" else other))
        if executeFormatting description (selectionBounds()) command then render()

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
        if executeFormatting "Number format" (selectionBounds()) (Rbt.Spreadsheet.Session.NumberFormat code) then render()

    /// One decimal place more for the numbers of the selection (one undo
    /// step): every cell that holds a number gets its own number format with
    /// a decimal more in each section ("0.0" -> "0.00", "0%" -> "0.0%"); a
    /// cell in the General format starts from the decimals its value shows
    /// (12.5 -> "0.00"). Dates, times, fractions and text are left alone.
    member _.IncreaseDecimals() =
        if executeFormatting "Increase decimals" (selectionBounds()) (Rbt.Spreadsheet.Session.ChangeDecimals 1) then render()

    /// One decimal place fewer (see IncreaseDecimals); a number without
    /// decimals stays as it is.
    member _.DecreaseDecimals() =
        if executeFormatting "Decrease decimals" (selectionBounds()) (Rbt.Spreadsheet.Session.ChangeDecimals -1) then render()

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
        if executeFormatting "Paint format" (range.From.Row, range.From.Col, range.To.Row, range.To.Col) (Rbt.Spreadsheet.Session.PaintFormat wanted) then render()

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
        if executeFormatting "Border" (selectionBounds()) (Rbt.Spreadsheet.Session.Border (side, style, color)) then render()

    /// Merge cells in a range (one undo step).
    member _.MergeCells(startRow: int, startCol: int, endRow: int, endCol: int) =
        executeFormatting "Merge cells" (startRow, startCol, endRow, endCol) Rbt.Spreadsheet.Session.MergeCells |> ignore
        render()

    /// Unmerge cells at address (one undo step).
    member _.UnmergeCells(row: int, col: int) =
        executeFormatting "Unmerge cells" (selectionBounds()) (Rbt.Spreadsheet.Session.UnmergeCells { Row = row; Col = col }) |> ignore
        render()

    /// The hyperlink on a cell of the active sheet, or null:
    /// { range: "B2" | "B2:C3", target, isInternal, tooltip: string | null,
    ///   display: string | null, safe } - `isInternal`: the target is a place
    /// in the workbook ("Sheet2!A1", a defined name), else an address;
    /// `safe`: an address with a scheme a host may open (http, https, mailto,
    /// tel). GetCell's `hyperlink` is the same object.
    member _.GetHyperlink(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetHyperlink" model row; Col = Host.colArg "GetHyperlink" model col }
        match SheetOps.hyperlinkAt model addr with
        | Some link -> Host.hyperlinkToObj link
        | None -> null

    /// Every hyperlink of a sheet (the active one, or the one at
    /// `sheetIndex`), in the order they were defined (see GetHyperlink).
    member _.GetHyperlinks(sheetIndex: obj) : obj array =
        (sheetArg "GetHyperlinks" sheetIndex).Hyperlinks |> List.map Host.hyperlinkToObj |> List.toArray

    /// Put a hyperlink on cells of the active sheet, in place of whatever
    /// links they had (one undo step). `range`: "B2", "B2:C3" or several
    /// areas. `target`: an address ("https://example.com", "mailto:a@b.c";
    /// any text is stored, but only http, https, mailto and tel addresses are
    /// ever reported `safe`), or a place in the workbook: "Sheet2!A1",
    /// "'My sheet'!B2:C3", "A1" (this sheet), a defined name, or any of them
    /// after "#". `display` (optional) becomes the text of the first cell;
    /// without it an empty first cell shows the target. `tooltip` (optional)
    /// is the screen tip. The cells get the look of a link (blue,
    /// underlined) unless they have a text colour of their own. Throws when
    /// the range or the target is missing or empty.
    member _.SetHyperlink(range: obj, target: obj, display: obj, tooltip: obj) : unit =
        let ranges = SheetOps.parseRanges "SetHyperlink" (Host.stringArg "SetHyperlink" "the range" range)
        let target = (Host.stringArg "SetHyperlink" "the target" target).Trim()
        if target = "" || target = "#" then Host.hostFail "SetHyperlink" "the target is empty (RemoveHyperlink removes a link)"
        let display = Host.optionalArg (Host.stringArg "SetHyperlink" "display") display |> Option.filter (fun text -> text <> "")
        let tooltip = Host.optionalArg (Host.stringArg "SetHyperlink" "tooltip") tooltip |> Option.filter (fun text -> text <> "")
        let isName (name: string) =
            let prefix, rest = SheetOps.splitSheetPrefix name
            let home = match prefix with Some sheetName -> sheetIndexOfName sheetName | None -> activeSheetIndex
            home >= 0 && (workbook |> Option.bind (Rbt.Spreadsheet.Workbook.getName (Some (sheetIdForIndex home)) rest)).IsSome
        let stored, isInternal = SheetOps.classifyHyperlinkTarget isName target
        if changeHyperlinks "Hyperlink" (SheetOps.setHyperlink ranges stored isInternal display tooltip) then render()

    /// Take the hyperlinks off cells of the active sheet (one undo step), and
    /// with them the look of a link, as Excel's "Remove Hyperlink" does; the
    /// texts stay. A link that covered more cells keeps the others. Returns
    /// how many links were removed or cut.
    member _.RemoveHyperlink(range: obj) : int =
        let ranges = SheetOps.parseRanges "RemoveHyperlink" (Host.stringArg "RemoveHyperlink" "the range" range)
        let mutable count = 0
        let changed = changeHyperlinks "Remove hyperlink" (fun m -> let removed, n = SheetOps.removeHyperlinks ranges m in count <- n; removed)
        if changed then render()
        count

    /// Call `callback` when the user activates a link that leads out of the
    /// workbook (Ctrl+click or Cmd+click on its cell, a plain click in
    /// read-only mode, or ActivateHyperlink), with
    /// { target, row, col, sheetIndex, isInternal: false, safe, tooltip }.
    /// The grid never opens an address itself: the host decides, and should
    /// open only what is `safe` (http, https, mailto, tel; a "javascript:"
    /// address is reported with safe: false). Links to a place in the
    /// workbook are followed by the grid and not reported. Returns a
    /// function that removes the subscription.
    member _.OnHyperlinkActivate(callback: System.Action<obj>) : System.Action =
        let listener = callback.Invoke
        hyperlinkListeners.Add listener
        System.Action(fun () -> hyperlinkListeners.Remove listener |> ignore)

    /// Activate the hyperlink on a cell of the active sheet, as a Ctrl+click
    /// on it does. A link to a place in the workbook is followed (the sheet
    /// switched to, the place selected and scrolled to); a link out of the
    /// workbook is handed to the OnHyperlinkActivate subscribers and nothing
    /// else happens. Returns { target, row, col, sheetIndex, isInternal,
    /// safe, tooltip, followed } - `followed`: the grid went to the place -
    /// or null when the cell has no link.
    member this.ActivateHyperlink(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "ActivateHyperlink" model row; Col = Host.colArg "ActivateHyperlink" model col }
        match SheetOps.hyperlinkAt model addr with
        | None -> null
        | Some link ->
            let sheetIndex = activeSheetIndex
            let payload (followed: bool) =
                createObj [ "target" ==> link.Target; "row" ==> addr.Row; "col" ==> addr.Col; "sheetIndex" ==> sheetIndex
                            "isInternal" ==> link.IsInternal; "safe" ==> SheetOps.hyperlinkIsSafe link
                            "tooltip" ==> Host.orNull link.Tooltip; "followed" ==> followed ]
            if link.IsInternal then
                // (A link to a place that is not there any more leads nowhere.)
                let followed = try this.GoTo (box link.Target) |> ignore; true with _ -> false
                payload followed
            else
                let event = payload false
                for listener in hyperlinkListeners.ToArray() do
                    if hyperlinkListeners.Contains listener then (try listener event with error -> reportError "OnHyperlinkActivate" (box error))
                event

    /// The comment on a cell of the active sheet, or null:
    /// { text, author: string | null, thread } - `thread` is null for a note
    /// and, for a comment thread, { id, author, date, text, resolved,
    /// replies: [{ id, author, date, text }] } (`text` and `author` are then
    /// those of the thread's first comment). GetCell's `comment` is the same object.
    member _.GetComment(row: obj, col: obj) : obj =
        let addr = { Row = Host.rowArg "GetComment" model row; Col = Host.colArg "GetComment" model col }
        match SheetOps.commentAt model addr with
        | Some info -> Host.commentToObj info
        | None -> null

    /// Every comment of a sheet (the active one, or the one at `sheetIndex`),
    /// row by row: [{ row, col, address, text, author, runs, thread }] (see GetComment).
    member _.GetComments(sheetIndex: obj) : obj array =
        SheetOps.allComments (sheetArg "GetComments" sheetIndex)
        |> List.map (fun info ->
            let comment = Host.commentToObj info
            createObj [ "row" ==> info.Cell.Row; "col" ==> info.Cell.Col; "address" ==> formatCellRef info.Cell.Row info.Cell.Col
                        "text" ==> info.Text; "author" ==> Host.orNull info.Author
                        "runs" ==> Host.getKey comment "runs"
                        "thread" ==> Host.getKey comment "thread" ])
        |> List.toArray

    /// Put a comment on a cell of the active sheet (one undo step). A cell
    /// without a comment, or with a note, gets a note with this text and
    /// author. On a cell with a comment thread the text of the thread's first
    /// comment changes (and its author, when one is given); the replies stay.
    /// Throws when the text is empty (DeleteComment removes a comment).
    member _.SetComment(row: obj, col: obj, text: obj, author: obj) =
        let addr = { Row = Host.rowArg "SetComment" model row; Col = Host.colArg "SetComment" model col }
        let text = Host.stringArg "SetComment" "the text" text
        if System.String.IsNullOrWhiteSpace text then Host.hostFail "SetComment" "the text is empty (DeleteComment removes a comment)"
        let author = commentAuthorArg "SetComment" author
        if changeComments "Comment" (SheetOps.setComment addr text author) then render()

    /// Put a note whose text is formatted in runs on a cell of the active
    /// sheet (one undo step): `runs` is an array of { text, bold?, italic?,
    /// underline?, strikethrough?, color? ("#RRGGBB"), fontName?, fontSize?
    /// (points) }; what a run leaves out is the note's usual look. The note's
    /// text is the runs' texts joined. On a cell with a comment thread the
    /// first comment gets that text (threads hold plain text). Throws when
    /// the runs hold no text.
    member _.SetCommentRuns(row: obj, col: obj, runs: obj, author: obj) =
        let addr = { Row = Host.rowArg "SetCommentRuns" model row; Col = Host.colArg "SetCommentRuns" model col }
        let runs = Host.commentRunsArg "SetCommentRuns" runs
        if runs |> List.forall (fun (text, _) -> System.String.IsNullOrWhiteSpace text) then Host.hostFail "SetCommentRuns" "the runs hold no text (DeleteComment removes a comment)"
        let author = commentAuthorArg "SetCommentRuns" author
        if changeComments "Comment" (SheetOps.setCommentRuns addr runs author) then render()

    /// Remove the comment of a cell of the active sheet - a note, or a thread
    /// with its replies - as one undo step. True when the cell had one.
    member _.DeleteComment(row: obj, col: obj) : bool =
        let addr = { Row = Host.rowArg "DeleteComment" model row; Col = Host.colArg "DeleteComment" model col }
        let removed = changeComments "Delete comment" (SheetOps.deleteComment addr)
        if removed then render()
        removed

    /// The same as DeleteComment (the older name).
    member this.RemoveComment(row: obj, col: obj) : unit = this.DeleteComment(row, col) |> ignore

    /// Reply to the comment of a cell of the active sheet (one undo step) and
    /// return the reply's id. A note becomes the first comment of a thread, as
    /// Excel's "convert to comments" does. The reply is dated now. Throws when
    /// the cell has no comment or the text is empty.
    member _.AddCommentReply(row: obj, col: obj, text: obj, author: obj) : string =
        let addr = { Row = Host.rowArg "AddCommentReply" model row; Col = Host.colArg "AddCommentReply" model col }
        let text = Host.stringArg "AddCommentReply" "the text" text
        if System.String.IsNullOrWhiteSpace text then Host.hostFail "AddCommentReply" "the text is empty"
        let author = commentAuthorArg "AddCommentReply" author
        match SheetOps.addCommentReply addr text author (Some (commentDateNow ())) model with
        | Result.Error message -> Host.hostFail "AddCommentReply" message
        | Ok (changed, id) ->
            changeComments "Reply to comment" (fun _ -> changed) |> ignore
            render()
            id

    /// Remove one reply (by its id, see GetComment) from the comment thread
    /// of a cell of the active sheet, as one undo step. Throws when the cell
    /// has no thread or the thread no such reply.
    member _.DeleteCommentReply(row: obj, col: obj, replyId: obj) : unit =
        let addr = { Row = Host.rowArg "DeleteCommentReply" model row; Col = Host.colArg "DeleteCommentReply" model col }
        match SheetOps.deleteCommentReply addr (Host.stringArg "DeleteCommentReply" "the reply id" replyId) model with
        | Result.Error message -> Host.hostFail "DeleteCommentReply" message
        | Ok changed ->
            changeComments "Delete reply" (fun _ -> changed) |> ignore
            render()

    /// Mark the comment thread of a cell resolved (the default) or open again,
    /// as one undo step. Throws when the cell has no thread (a note cannot be resolved).
    member _.ResolveComment(row: obj, col: obj, resolved: obj) : unit =
        let addr = { Row = Host.rowArg "ResolveComment" model row; Col = Host.colArg "ResolveComment" model col }
        let resolved = Host.optionalArg (Host.boolArg "ResolveComment" "resolved") resolved |> Option.defaultValue true
        match SheetOps.setCommentResolved addr resolved model with
        | Result.Error message -> Host.hostFail "ResolveComment" message
        | Ok changed -> if changeComments "Resolve comment" (fun _ -> changed) then render()

    /// Set workbook context for table reference support.
    member _.SetWorkbookContext(wb: Workbook) =
        invalidateCalcEngine()
        workbook <- Some wb
        session.UpdateView(currentWorkbookSnapshot())
        session.RebaseCalculation()
        render()  // Re-render to apply table styles

    /// Get workbook
    member _.GetWorkbook() =
        workbook

    /// The tables of the workbook, sheet by sheet: [{ name, sheetIndex,
    /// range ("A1:D9": header row, data and total row), hasHeaders,
    /// hasTotals, columns: [names, left to right], style: { name,
    /// showRowStripes, showColumnStripes, showFirstColumn, showLastColumn }
    /// | null, hasAutoFilter (the header cells carry filter buttons) }].
    member _.GetTables() : obj array =
        match workbook with
        | None -> [||]
        | Some _ ->
            SheetOps.tablesOf (bookSnapshot ())
            |> List.choose (fun table ->
                let index = sheetIndexOfId table.SheetId
                if index < 0 then None else Some (Host.tableToObj index table))
            |> List.toArray

    /// Make a table of a range of the active sheet:
    /// `DefineTable(range, name?, hasHeaders?, style?)` -> the table's name.
    /// `range` "A1:D9"; `name` (default the first free "TableN") follows
    /// the rules of defined names and is not the name of another table or
    /// defined name; `hasHeaders` (default true): the first row names the
    /// columns - blank header cells become "ColumnN", repeated names get a
    /// number appended, and every header cell becomes text; false: every row
    /// is data and the columns are Column1, Column2 ... (no row is inserted,
    /// unlike Excel); `style`: a built-in style name or { name,
    /// showRowStripes?, showColumnStripes?, showFirstColumn?,
    /// showLastColumn? }, "none" for no style (default
    /// "TableStyleMedium2"). Structured references to the table resolve from
    /// then on. One undo step. Throws for a bad name, a range that is
    /// outside the sheet, whole rows / columns, overlaps another table or
    /// holds merged cells, and a header row without a row of data.
    /// The earlier call `DefineTable(name, startRow, startCol, endRow,
    /// endCol, columns, hasHeaders)` (0-based, the column names given) keeps
    /// working: it registers the table for structured references only and
    /// is now an undo step too.
    member _.DefineTable(first: obj, second: obj, third: obj, fourth: obj, fifth: obj, sixth: obj, seventh: obj) : string =
        let memberName = "DefineTable"
        if Host.isJsNumber second then
            let name = Host.stringArg memberName "name" first
            let columns : string list =
                if Host.isJsArray sixth then List.ofArray (unbox<string array> sixth)
                elif Host.isNullish sixth then []
                else unbox<string list> sixth
            let table : Rbt.Spreadsheet.Workbook.Table =
                { Name = name
                  SheetId = sheetIdForIndex activeSheetIndex
                  StartRow = Host.intArg memberName "startRow" second
                  StartCol = Host.intArg memberName "startCol" third
                  EndRow = Host.intArg memberName "endRow" fourth
                  EndCol = Host.intArg memberName "endCol" fifth
                  Columns = columns
                  HasHeaders = (Host.optionalArg (Host.boolArg memberName "hasHeaders") seventh |> Option.defaultValue true)
                  HasTotals = false }
            if changeTables "Define table" (fun wb -> { Book = Rbt.Spreadsheet.Workbook.defineTable table wb; Cells = [] }) then render()
            name
        else
            let range = Host.rangeArg memberName first
            let name = Host.optionalArg (Host.stringArg memberName "name") second
            let hasHeaders = Host.optionalArg (Host.boolArg memberName "hasHeaders") third |> Option.defaultValue true
            let style = Host.tableStyleArg memberName (Some (SheetOps.tableStyleNamed SheetOps.defaultTableStyleName)) fourth
            let made : string ref = ref ""
            let sheetId = sheetIdForIndex activeSheetIndex
            changeTables "Define table" (fun wb ->
                let change, name = SheetOps.defineTable memberName sheetId range name hasHeaders style wb
                made.Value <- name
                change) |> ignore
            render()
            made.Value

    /// Give a table another range ("A1:F20") on its sheet: the first row
    /// stays where it is and the new range overlaps the old one; with a
    /// total row the range includes it (it moves to the new last row, which
    /// has to be free). Columns that come in are named by their header
    /// cells. One undo step. Throws for an unknown table or a bad range.
    member _.ResizeTable(name: obj, range: obj) : unit =
        let memberName = "ResizeTable"
        let name = Host.stringArg memberName "name" name
        let range = Host.rangeArg memberName range
        if changeTables "Resize table" (SheetOps.resizeTable memberName name range) then render()

    /// Rename a table; structured references to it in every formula of the
    /// workbook (cells, conditional formats, validations, defined names)
    /// follow. One undo step. Throws for an unknown table and for a name
    /// that is not valid or taken.
    member _.RenameTable(name: obj, newName: obj) : unit =
        let memberName = "RenameTable"
        let name = Host.stringArg memberName "name" name
        let newName = Host.stringArg memberName "newName" newName
        if changeTables "Rename table" (SheetOps.renameTable memberName name newName) then render()

    /// Remove a table. `keepData` (default true) converts it to a range, as
    /// Excel does: the cells stay and structured references to the table
    /// become the plain references they stood for ("Sales[Amount]" ->
    /// "$B$2:$B$9"); false also empties the table's cells (formats stay),
    /// and references to it show #REF!. One undo step. Throws for an unknown
    /// table.
    member _.DeleteTable(name: obj, keepData: obj) : unit =
        let memberName = "DeleteTable"
        let name = Host.stringArg memberName "name" name
        let keepData = Host.optionalArg (Host.boolArg memberName "keepData") keepData |> Option.defaultValue true
        if changeTables "Delete table" (SheetOps.deleteTable memberName name keepData) then render()

    /// Show (true) or remove (false) the total row of a table: the row below
    /// its data, with "Total" in the first column and in the last the sum
    /// of that column when it holds numbers, else the count of its entries
    /// (SUBTOTAL: filtered rows are left out). The row below the table has
    /// to be empty (nothing is moved down); removing empties the row. One
    /// undo step. Throws for an unknown table and a row that is in use.
    member _.SetTableTotalRow(name: obj, show: obj) : unit =
        let memberName = "SetTableTotalRow"
        let name = Host.stringArg memberName "name" name
        let show = Host.boolArg memberName "show" show
        if changeTables (if show then "Total row" else "Remove total row") (SheetOps.setTableTotalRow memberName name show) then render()

    /// Give a table another style: a built-in style name
    /// ("TableStyleLight9"), { name, showRowStripes?, showColumnStripes?,
    /// showFirstColumn?, showLastColumn? }, or null / "none" for no style.
    /// One undo step. (The grid draws the header fill and the row bands of a
    /// style in the Office theme's colours.)
    member _.SetTableStyle(name: obj, style: obj) : unit =
        let memberName = "SetTableStyle"
        let name = Host.stringArg memberName "name" name
        let style = Host.tableStyleArg memberName None style
        if changeTables "Table style" (SheetOps.setTableStyle memberName name style) then render()

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

    /// Start a search of the active sheet the old way (text contained in a
    /// cell's value or formula): FindNext() then steps through the results
    /// and answers true / false. Hosts use Find.
    member _.StartSearch(query: string, caseSensitive: bool) =
        hostSearch <- None
        searchCache.Clear()
        model <- startSearch query caseSensitive model
        render()

    /// Search the workbook. `text` is what to look for; `options` (all
    /// optional) is { matchCase: boolean (false), wholeCell: boolean (false:
    /// a part of the cell's text is enough), regex: boolean (false; `text`
    /// is then a regular expression - a JavaScript one in the browser),
    /// lookIn: "values" (what cells show; the default) | "formulas" (what
    /// was entered: a formula with its "=", any other cell as it would be
    /// typed), scope: "sheet" (the sheet on display; the default) |
    /// "workbook" (every sheet with a tab, hidden sheets left out),
    /// range: "B2:D9" (on the sheet on display; not with scope "workbook") }.
    /// Returns the matches, [{ sheetIndex, row, col, address, text (the
    /// text that was searched: shown or entered) }], row by row, sheets in
    /// tab order starting at the one on display. Nothing is selected yet:
    /// the search is remembered and FindNext() / FindPrevious() /
    /// ReplaceMatch() go by it (and find again in cells that changed
    /// since). A workbook search converts the sheets of an opened file that
    /// were not needed so far. Throws for an empty text, an unreadable
    /// regular expression and unknown options.
    member _.Find(text: obj, options: obj) : obj array =
        let memberName = "Find"
        let text = Host.stringArg memberName "text" text
        let request = Host.findRequestOfObj memberName options
        let search = { Finder = SheetOps.finderOf memberName text request.Options; Request = request }
        hostSearch <- Some search
        searchCache.Clear()
        let found = allMatches search
        // (From the sheet on display on, then the sheets before it.)
        Array.append (found |> Array.filter (fun (index, _, _) -> index >= activeSheetIndex)) (found |> Array.filter (fun (index, _, _) -> index < activeSheetIndex))
        |> Array.map (fun (index, addr, text) -> Host.findMatchToObj index addr text)

    /// Step from the active cell to the next (`forward`) or previous match of
    /// the search Find started: the match's sheet is shown, its cell selected
    /// and scrolled into view. null when nothing matches.
    member private this.StepSearch(search: HostSearch, forward: bool) : obj =
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let found = allMatches search
        let active = activeCellAddress ()
        let places = found |> Array.map (fun (index, addr, _) -> index, addr.Row, addr.Col)
        match SheetOps.stepMatch forward (activeSheetIndex, active.Row, active.Col) places with
        | None -> null
        | Some at ->
            let index, addr, text = found.[at]
            if index <> activeSheetIndex then this.SwitchSheet index
            selectCellAddress addr
            revealCell addr.Row addr.Col
            render()
            Host.findMatchToObj index addr text

    /// Go to the next match. After Find: the first match after the active
    /// cell (row by row, then the following sheets for a workbook search,
    /// starting over after the last one); its sheet is shown and its cell
    /// selected; returns the match { sheetIndex, row, col, address, text },
    /// or null when nothing matches. After StartSearch (the old search):
    /// selects the next result and returns true, or false when there is none.
    member this.FindNext() : obj =
        match hostSearch with
        | Some search -> this.StepSearch(search, true)
        | None ->
            match nextSearchResult model with
            | Some newModel ->
                model <- newModel
                render()
                box true
            | None -> box false

    /// Go to the match before the active cell of the search Find started
    /// (the last one after the first): returns the match, or null when
    /// nothing matches. Throws when Find was not called.
    member this.FindPrevious() : obj =
        match hostSearch with
        | Some search -> this.StepSearch(search, false)
        | None -> Host.hostFail "FindPrevious" "there is no search (call Find first)"

    /// Replace in the active cell when it is a match of the search Find
    /// started, then go to the next match. Returns { replaced: boolean,
    /// next: match | null }. `replacement` replaces every occurrence in the
    /// cell's text ($1, $2 ... are the groups of a regular expression); the
    /// result is read the way typed input is (a number, a date, a formula)
    /// and the cell keeps its format; it is not checked against data
    /// validation. Searching "formulas" the replacement is made in the
    /// formula text; searching "values" a formula cell is skipped (replaced:
    /// false) and the replacement is made in the text other cells show. One
    /// undo step. When the active cell is not a match nothing is replaced
    /// and the next match is selected. Throws when Find was not called.
    member this.ReplaceMatch(replacement: obj) : obj =
        let memberName = "ReplaceMatch"
        let replacement = Host.stringArg memberName "replacement" replacement
        match hostSearch with
        | None -> Host.hostFail memberName "there is no search (call Find first)"
        | Some search ->
            if model.EditingCell.IsSome then finishEditing true |> ignore
            let addr = activeCellAddress ()
            let inRange =
                match search.Request.Range with
                | Some range -> addr.Row >= range.From.Row && addr.Row <= range.To.Row && addr.Col >= range.From.Col && addr.Col <= range.To.Col
                | None -> true
            let changed =
                match getCell addr model with
                | Some data when inRange && not (Map.containsKey addr model.Spill) ->
                    SheetOps.replaceInCellIn model.Date1904 search.Finder search.Request.Options replacement (currentYear()) data
                | _ -> None
            let replaced =
                match changed with
                | Some cellData ->
                    // The shared session snapshots the pre-edit cell.
                    let cellData =
                        match cellData.Value with
                        | CellValue.Formula (formula, _) -> { cellData with Value = CellValue.Formula (formula, evaluateFormula formula model workbook) }
                        | _ -> cellData
                    recalculateSingleCellEdit addr cellData
                    model <- pushCommand "Replace" model
                    true
                | None -> false
            let next = this.StepSearch(search, true)
            if isNull next then render()
            createObj [ "replaced" ==> replaced; "next" ==> next ]

    /// Replace current selection
    member _.Replace(findText: string, replaceText: string, caseSensitive: bool) =
        let replaced = replaceInSelection findText replaceText caseSensitive
        if replaced then render()
        replaced

    /// Replace every match and return how many cells changed, as ONE undo
    /// step - also across sheets. `ReplaceAll(text, replacement, options)`
    /// takes the options of Find ({ matchCase, wholeCell, regex, lookIn,
    /// scope, range }) and replaces as ReplaceMatch does (formula cells only
    /// when looking in formulas; the cells of a spilled array never). With
    /// scope "workbook" every sheet with a tab is changed and one Undo() on
    /// the sheet on display restores them all. Formulas are recalculated.
    /// The old call `ReplaceAll(text, replacement, caseSensitive: boolean)`
    /// keeps working: the active sheet, text contained in a cell's value or
    /// formula. Throws as Find does.
    member _.ReplaceAll(findText: obj, replaceText: obj, options: obj) : int =
        let memberName = "ReplaceAll"
        if Host.jsTypeOf options = "boolean" then
            let count = replaceEverywhere (Host.stringArg memberName "text" findText) (Host.stringArg memberName "replacement" replaceText) (unbox<bool> options)
            render()
            count
        else
            let text = Host.stringArg memberName "text" findText
            let replacement = Host.stringArg memberName "replacement" replaceText
            let request = Host.findRequestOfObj memberName options
            let finder = SheetOps.finderOf memberName text request.Options
            if model.EditingCell.IsSome then finishEditing true |> ignore
            let year = currentYear()
            let changesOf (index: int) =
                SheetOps.replaceInSheet finder request.Options replacement year request.Range (sheetModelAt index)
            let put (index: int) (cells: (CellAddress * CellData option) list) =
                if index = activeSheetIndex then model <- restoreCellOptions cells model
                else sheets.[index] <- restoreCellOptions cells (sheetModelAt index)
            if not request.Workbook then
                let changes = changesOf activeSheetIndex
                if not changes.IsEmpty then
                    put activeSheetIndex (changes |> List.map (fun (addr, _, after) -> addr, Some after))
                    recalculateFormulas()
                    model <- pushCommand "Replace All" model
                render()
                changes.Length
            else
                // (Formulas of every sheet may read the cells that change.)
                convertFormulaSheets ()
                let changes =
                    [ for index in 0 .. sheets.Length - 1 do
                        if index = activeSheetIndex || sheetHasTab index then
                            match changesOf index with
                            | [] -> ()
                            | found -> yield sheetIdForIndex index, found ]
                let count = changes |> List.sumBy (fun (_, found) -> found.Length)
                if count > 0 then
                    runBookStep "Replace All" (fun () ->
                        for id, found in changes do
                            let index = sheetIndexOfId id
                            if index >= 0 then
                                put index (found |> List.map (fun (addr, _, after) -> addr, Some after))
                        recalculateFormulas())
                render()
                count

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
        let formatChange = if (given "rule").IsSome || (given "format").IsSome then Some format else None
        let result = executeRule name (Rbt.Spreadsheet.Session.UpdateConditionalFormat (id, ranges, stopIfTrue, formatChange, priority))
        if result.Changed then render()

    /// Remove the conditional format with an id from the active sheet (one
    /// undo step). Throws when the sheet has none with that id.
    member _.RemoveConditionalFormat(id: obj) : unit =
        let id = Host.intArg "RemoveConditionalFormat" "id" id
        if (executeRule "RemoveConditionalFormat" (Rbt.Spreadsheet.Session.RemoveConditionalFormat id)).Changed then render()

    /// Put the active sheet's conditional formats in a new order of
    /// precedence (one undo step): `ids` lists every rule of the sheet once,
    /// the one that takes precedence first. Throws on an unknown or repeated
    /// id and when a rule is left out.
    member _.ReorderConditionalFormats(ids: obj) : unit =
        let name = "ReorderConditionalFormats"
        if not (Host.isJsArray ids) then Host.hostFail name "ids must be an array of rule ids"
        let ids = unbox<obj array> ids |> Array.map (Host.intArg name "an id") |> Array.toList
        if (executeRule name (Rbt.Spreadsheet.Session.ReorderConditionalFormats ids)).Changed then render()

    /// Take the conditional formats off the active sheet (no argument), or
    /// off the cells of `range` (areas as in AddConditionalFormat): a rule
    /// that covers more cells keeps them, a rule left with none is removed.
    /// One undo step. Returns how many rules were removed or cut.
    member _.ClearConditionalFormats(range: obj) : int =
        let name = "ClearConditionalFormats"
        let holes = Host.optionalArg (fun value -> SheetOps.parseRanges name (Host.stringArg name "the range" value)) range
        let result = executeRule name (Rbt.Spreadsheet.Session.ClearConditionalFormats holes)
        if result.Changed then render()
        result.Affected

    /// Remove the conditional formats from one cell of the active sheet (one
    /// undo step): a rule covering more than the cell keeps the rest of its range.
    member _.RemoveConditionalFormats(row: int, col: int) =
        let addr = { Row = row; Col = col }
        let result = executeRule "RemoveConditionalFormats" (Rbt.Spreadsheet.Session.ClearConditionalFormats (Some [ { From = addr; To = addr } ]))
        if result.Changed then render()

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
        let validation = SheetOps.validationOfDescriptionIn model.Date1904 name (Host.validationOfObj name (requiredArg name "the rule" rule))
        let result = executeRule name (Rbt.Spreadsheet.Session.SetValidation (ranges, validation))
        if result.Changed then render()
        result.Id.Value

    /// Take the validations off cells of the active sheet (one undo step):
    /// a validation that covers more cells keeps them, one left with no cell
    /// is removed. `range` as in SetValidation. Returns how many validations
    /// were removed or cut.
    member _.ClearValidation(range: obj) : int =
        let name = "ClearValidation"
        let holes = SheetOps.parseRanges name (Host.stringArg name "the range" (requiredArg name "the range" range))
        let result = executeRule name (Rbt.Spreadsheet.Session.ClearValidation holes)
        if result.Changed then render()
        result.Affected

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
    /// 7. The grid draws the outline strip with its +/- and level buttons
    /// (see SetShowOutline). One undo step.
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

    /// Excel's outline level buttons: show the rows / columns (`axis` "rows"
    /// or "columns") of the levels below `level` and collapse every group
    /// from `level` on (1 = only what is in no group; one more than the
    /// deepest level = everything). One undo step.
    member _.ShowOutlineLevel(axis: obj, level: obj) : unit =
        let rows =
            match Host.stringArg "ShowOutlineLevel" "the axis" axis with
            | "rows" -> true
            | "columns" -> false
            | other -> Host.hostFail "ShowOutlineLevel" (sprintf "the axis must be \"rows\" or \"columns\", got \"%s\"" other)
        let level = Host.intArg "ShowOutlineLevel" "level" level
        if level < 1 || level > SheetOps.maxOutlineLevel + 1 then Host.hostFail "ShowOutlineLevel" (sprintf "the level has to be 1..%d, not %d" (SheetOps.maxOutlineLevel + 1) level)
        if changeAxes "Show outline level" (SheetOps.showOutlineLevel rows level) then render()

    /// Where the summary rows / columns of the active sheet's groups are:
    /// `below` true (Excel's default) = the row after a group, false = the
    /// row before it; `right` likewise for columns (the sheet's outlinePr).
    /// The +/- button of a group is on its summary row / column. Collapsed
    /// groups stay collapsed. One undo step.
    member _.SetOutlineSummary(below: obj, right: obj) : unit =
        let below = Host.boolArg "SetOutlineSummary" "below" below
        let right = Host.boolArg "SetOutlineSummary" "right" right
        if changeAxes "Outline settings" (SheetOps.setOutlineSummary below right) then render()

    /// Show (the default) or hide the outline strips the grid draws for
    /// grouped rows / columns. A view setting of the component (every sheet;
    /// not saved, no undo step).
    member _.SetShowOutline(show: obj) : unit =
        let show = Host.boolArg "SetShowOutline" "show" show
        if show <> showOutline then
            showOutline <- show
            render()

    member _.GetShowOutline() : bool = showOutline

    /// Show (the default) or hide the outline's symbols of the active sheet:
    /// its brackets, +/- and level buttons. A setting of the sheet, saved with
    /// it (the sheet view's showOutlineSymbols); the groups themselves stay.
    /// One undo step. (SetShowOutline hides the strips of every sheet in this
    /// component without changing any sheet.)
    member _.SetShowOutlineSymbols(show: obj) : unit =
        let show = Host.boolArg "SetShowOutlineSymbols" "show" show
        if changeAxes "Outline symbols" (SheetOps.setOutlineSymbols show) then render()

    member _.GetShowOutlineSymbols() : bool = model.Axes.ShowSymbols

    /// The outline of the active sheet: { rows: [group], columns: [group],
    /// rowLevels, columnLevels (the deepest level, 0 = no groups),
    /// summaryBelow, summaryRight, showSymbols } with
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
            "summaryBelow" ==> model.Axes.SummaryBelow
            "summaryRight" ==> model.Axes.SummaryRight
            "showSymbols" ==> model.Axes.ShowSymbols
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

    /// The view of the grid: "normal", or "pageBreak" - the page break preview, in which every
    /// page the sheet prints on (GetPrintPages) is outlined over its cells with its number
    /// (`div.spreadsheet-page[data-page][data-range]`; an edge is solid at a manual break and
    /// where the print area ends, dashed at a break the paper makes). The grid is edited as
    /// before; the frames follow the cells, the sizes and the page setup. The mode is the
    /// component's, for every sheet, and is not saved. Throws for another mode.
    member _.SetViewMode(mode: obj) : unit =
        let wanted = Host.stringArg "SetViewMode" "the mode" mode
        if wanted <> "normal" && wanted <> "pageBreak" then Host.hostFail "SetViewMode" (sprintf "\"%s\" is no view mode (\"normal\", \"pageBreak\")" wanted)
        if wanted <> viewMode then
            viewMode <- wanted
            render()

    /// "normal" or "pageBreak" (see SetViewMode).
    member _.GetViewMode() : string = viewMode

    /// Move a manual page break, as dragging its line does: `kind` "row" or "column", the break
    /// before row / column `from` (0-based) is from now on before `target`. One undo step.
    /// False when there is no manual break at `from`, or it stays where it is; throws for a
    /// place a break cannot be at. A break moved onto another one is one break.
    member _.MovePageBreak(kind: obj, from: obj, target: obj, ?sheetIndex: obj) : bool =
        let memberName = "MovePageBreak"
        if model.EditingCell.IsSome then finishEditing true |> ignore
        let kind = Host.stringArg memberName "kind" kind
        if kind <> "row" && kind <> "column" then Host.hostFail memberName (sprintf "\"%s\" is no kind of break (\"row\", \"column\")" kind)
        let from = Host.intArg memberName "from" from
        let target = Host.intArg memberName "target" target
        let changed =
            changePageSetup memberName (defaultArg sheetIndex null) "Move page break" (fun setup ->
                let moved (breaks: int list) = if List.contains from breaks then target :: (breaks |> List.filter (fun known -> known <> from)) else breaks
                if kind = "row" then SheetOps.setPageBreaks memberName (moved setup.ManualRowBreaks) setup.ManualColBreaks setup
                else SheetOps.setPageBreaks memberName setup.ManualRowBreaks (moved setup.ManualColBreaks) setup)
        if changed then render()
        changed

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
        let scope = scopeIndex |> Option.map sheetIdForIndex
        if executeName memberName (Rbt.Spreadsheet.Session.SetDefinedName (scope, name, refersTo, comment)) then render()

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
        let scope = scopeIndex |> Option.map sheetIdForIndex
        if executeName memberName (Rbt.Spreadsheet.Session.DeleteDefinedName (scope, name)) then render()

    /// The workbook's date system: 1900, or 1904 for a workbook opened from a
    /// file that says `date1904` (serial 0 is then 1904-01-01). The serials in
    /// the cells are the file's; number formats, the date functions and typed
    /// dates work in this system, and a save keeps it.
    member _.GetDateSystem() : int =
        if model.Date1904 then 1904 else 1900

    /// How the workbook calculates circular references: { enabled,
    /// maxIterations, maxChange } - Excel's "Enable iterative calculation"
    /// (File > Options > Formulas), kept in the file. Off (the default, and
    /// what a new workbook has): the formulas on a circular reference, whether
    /// the cycle stays on one sheet or crosses sheets, and what reads them are
    /// not calculated - they keep the values they have (0, one just entered),
    /// as in Excel. On: they and what reads them are calculated in sweeps, in
    /// reading order, from the values they have, until none changes by
    /// `maxChange` or more, at most `maxIterations` times (100 and 0.001 when
    /// off) - and swept once more at every calculation, as in Excel.
    member _.GetIterativeCalculation() : obj =
        match workbook |> Option.bind (fun wb -> wb.Iteration) with
        | Some setting -> createObj [ "enabled" ==> true; "maxIterations" ==> setting.MaxIterations; "maxChange" ==> setting.MaxChange ]
        | None -> createObj [ "enabled" ==> false; "maxIterations" ==> 100; "maxChange" ==> 0.001 ]

    // ---- functions of the function table and their providers ----------------------------------
    /// Answer a provider function (WEBSERVICE, STOCKHISTORY, CUBEVALUE, RTD, CALL ... - see
    /// `GetProviderFunctions`): `provider({ name, args, sheet, row, col })` gives the value - a
    /// number, text, a boolean, null, { error: "#N/A" }, an array or an array of rows (a spill) - or
    /// a Promise of it. `args` holds each argument evaluated: one value, or an array of rows for a
    /// range or array. A formula shows #BUSY! until the Promise settles (rejected: #CONNECT!); an
    /// answer is kept for the same arguments. null takes the provider away: the formulas keep the
    /// values they have. The workbook is calculated again with the provider.
    member _.RegisterFunctionProvider(name: obj, provider: obj) : unit =
        let memberName = "RegisterFunctionProvider"
        let name = Host.stringArg memberName "name" name
        if not (Rbt.Spreadsheet.Formula.isProviderFunction name) then
            Host.hostFail memberName (sprintf "%s is not a provider function (%s)" name
                (System.String.Join(", ", Rbt.Spreadsheet.Formula.providerFunctionSpecs |> List.map (fun (n, _, _) -> n))))
        if Host.isNullish provider then Rbt.Spreadsheet.Formula.unregisterFunctionProvider name
        else
            if Host.jsTypeOf provider <> "function" then
                Host.hostFail memberName "the provider is a function ({ name, args, sheet, row, col }) => value | Promise<value>"
            let invoke (call: Rbt.Spreadsheet.Formula.ProviderCall) (answer: CellValue[][] -> unit) =
                let sheetName =
                    match call.Sheet, workbook with
                    | Some id, Some wb -> Map.tryFind id wb.Sheets |> Option.map (fun sheet -> box sheet.Name) |> Option.defaultValue null
                    | _ -> null
                let request =
                    createObj [ "name" ==> call.Name
                                "args" ==> (call.Arguments |> Array.map providerArgumentJs)
                                "sheet" ==> sheetName
                                "row" ==> call.Row
                                "col" ==> call.Col ]
                let result = callHostFunction provider request
                if isThenable result then
                    whenSettled result (fun value -> answer (providerRowsOf value))
                        (fun _ -> answer [| [| CellValue.Error Rbt.Spreadsheet.Formula.Err.connect |] |])
                else answer (providerRowsOf result)
            Rbt.Spreadsheet.Formula.registerFunctionProvider name { Invoke = invoke } |> ignore
        Rbt.Spreadsheet.Formula.notifyFunctionsArrived ()

    /// The functions a host can answer with `RegisterFunctionProvider`, and whether one has a provider now.
    member _.GetProviderFunctions() : obj[] =
        [| for (name, error, spills) in Rbt.Spreadsheet.Formula.providerFunctionSpecs ->
            createObj [ "name" ==> name
                        "provided" ==> Rbt.Spreadsheet.Formula.hasFunctionProvider name
                        "errorWithout" ==> ("#" + error)
                        "spills" ==> spills ] |]

    /// What INFO(field) gives - "directory", "origin", "osversion", "system", "release", "recalc",
    /// "numfile" - and "filename", the workbook's file name CELL("filename") puts with "directory"
    /// ("C:\\Books\\[Plan.xlsx]Sheet1"). null: INFO's default again. Shared by every grid of the page.
    member _.SetHostInfo(field: obj, value: obj) : unit =
        let memberName = "SetHostInfo"
        let field = Host.stringArg memberName "field" field
        Rbt.Spreadsheet.Formula.setHostInfo field (if Host.isNullish value then None else Some (Host.stringArg memberName "value" value))
        Rbt.Spreadsheet.Formula.notifyFunctionsArrived ()

    /// Load the modules of these function names (["ROMAN", "XIRR", ...]) ahead of their first use;
    /// resolves once they are installed. Opening a workbook and evaluating a formula load what they
    /// need by themselves (a function not loaded yet shows #BUSY! meanwhile).
    member _.LoadFunctions(names: obj) : JS.Promise<unit> =
        let memberName = "LoadFunctions"
        if not (Host.isJsArray names) then Host.hostFail memberName "names is an array of function names"
        let names = unbox<obj[]> names |> Array.map (Host.stringArg memberName "names[]")
        promiseOfReady (fun ready -> Rbt.Spreadsheet.Formula.requestFunctions names ready)

    /// Switch the iterative calculation of circular references on or off
    /// (see GetIterativeCalculation). `maxIterations` (1 .. 32767, default
    /// 100) and `maxChange` (0 or more, default 0.001) are Excel's two
    /// numbers. The workbook is calculated again. One undo step; saved with
    /// the workbook. Throws on arguments out of range.
    member _.SetIterativeCalculation(enabled: obj, maxIterations: obj, maxChange: obj) : unit =
        let memberName = "SetIterativeCalculation"
        let enabled = Host.boolArg memberName "enabled" enabled
        let count = Host.optionalArg (Host.intArg memberName "maxIterations") maxIterations |> Option.defaultValue 100
        let change = Host.optionalArg (Host.numberArg memberName "maxChange") maxChange |> Option.defaultValue 0.001
        if count < 1 || count > 32767 then Host.hostFail memberName (sprintf "maxIterations is 1 to 32767, not %d" count)
        if not (change >= 0.0) then Host.hostFail memberName "maxChange is 0 or more"
        let wb = ensureWorkbook ()
        let before = wb.Iteration
        let after = if enabled then Some { MaxIterations = count; MaxChange = change } else None
        if before <> after then
            if model.EditingCell.IsSome then finishEditing true |> ignore
            workbook <- workbook |> Option.map (fun current -> { current with Iteration = after })
            recalculateFromScratch()
            model <- pushCommand "Iterative calculation" model
            render()

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
        // (Any cell a sheet can have - to row 1 048 576 and column XFD -, as SelectCell: the grid is given room for it.)
        if range.From.Row < 0 || range.From.Col < 0 || range.To.Row > 1048575 || range.To.Col > 16383 then
            Host.hostFail memberName (sprintf "%s is outside the sheet" (SheetOps.rangeText range))
        if not sheetLocked then
            giveRoomFor range.From
            // (A whole row or column is that of the grid as it is.)
            if range.To.Row < 1048575 && range.To.Col < 16383 then giveRoomFor range.To
        let range =
            match SheetOps.clipRange model.RowCount model.ColCount range with
            | Some clipped -> clipped
            | None -> Host.hostFail memberName (sprintf "%s is outside the sheet" (SheetOps.rangeText range))
        if model.EditingCell.IsSome then finishEditing true |> ignore
        selectFromTo range.From range.To
        revealCell range.From.Row range.From.Col
        render()
        createObj [ "sheetIndex" ==> sheetIndex; "range" ==> SheetOps.rangeText range ]

    /// Open the shared conditional-rule manager. The dialog is mounted lazily
    /// on the grid so this public entry point also works without a shell.
#if FABLE_COMPILER
    member private this.ManagerPort() : Rbt.Spreadsheet.Managers.SpreadsheetPort = {
        GetWorkbookEpoch = fun () -> this.GetWorkbookEpoch()
        GetChangeSerial = fun () -> this.GetChangeSerial()
        GetActiveSheetIndex = fun () -> this.GetActiveSheetIndex()
        IsReadOnly = fun () -> this.IsReadOnly()
        IsEditing = fun () -> this.IsEditing()
        GetConditionalFormats = fun index -> this.GetConditionalFormats(index)
        GetSelectionRef = fun () -> this.GetSelectionRef()
        UpdateConditionalFormat = fun (id, changes) -> this.UpdateConditionalFormat(id, changes)
        AddConditionalFormat = fun (target, rule, options, fourth, fifth) ->
            this.AddConditionalFormat(target, rule, options, fourth, fifth)
        RemoveConditionalFormat = fun id -> this.RemoveConditionalFormat(id)
        GetDefinedNames = fun () -> this.GetDefinedNames()
        GetSheetNames = fun () -> this.GetSheetNames()
        SetDefinedName = fun (name, refersTo, scope, comment) -> this.SetDefinedName(name, refersTo, scope, comment)
        DeleteDefinedName = fun (name, scope) -> this.DeleteDefinedName(name, scope)
    }

    member private this.ManagerDialog() =
        match managerDialog with
        | Some dialog -> dialog
        | None ->
            let dialog = Rbt.FableUI.OfficeUiDialog.createState ()
            let unmount = Rbt.FableUI.Dom.mount (unbox (getBody())) (Rbt.FableUI.OfficeUiDialog.renderDialog dialog)
            managerDialog <- Some dialog
            cleanups.Add(fun () ->
                Rbt.FableUI.OfficeUiDialog.closeDialog dialog
                unmount()
                managerDialog <- None)
            dialog

    member this.ShowConditionalFormattingDialog() =
        if not disposed && not (this.IsReadOnly()) && this.TryCommitEdit() then
            Rbt.Spreadsheet.Managers.openConditionalManager (this.ManagerPort()) (this.ManagerDialog())

    /// Open the workbook/sheet-scoped name manager.
    member this.ShowDefinedNamesDialog() =
        if not disposed && not (this.IsReadOnly()) && this.TryCommitEdit() then
            Rbt.Spreadsheet.Managers.openNameManager (this.ManagerPort()) (this.ManagerDialog())
#else
    member _.ShowConditionalFormattingDialog() =
        invalidOp "The conditional-format dialog requires the browser editor."
    member _.ShowDefinedNamesDialog() =
        invalidOp "The name manager requires the browser editor."
#endif

    /// Open the grid's single find/replace surface (also used by its shell).
    member _.ShowFindReplaceDialog() = showFindReplaceDialog ()
    /// Shell command entrypoints use the same search controls and options.
    member _.FindNextInBar() =
        showFindReplaceDialog ()
        findNextAction ()

    member _.ReplaceAllInBar() =
        showFindReplaceDialog ()
        findReplaceAllAction ()

    /// Move the grid-owned search controls into a mounted shell.
    member _.AttachFindReplace(host: Element) =
        findHost <- Some host
        if not disposed then appendChild host (ensureFindBar ())

    /// A shell leaves the controls with their grid when that shell unmounts.
    member _.DetachFindReplace(host: Element) =
        if findHost |> Option.exists (fun owner -> obj.ReferenceEquals(owner, host)) then
            closeFindBar ()
            findHost <- None
            findBar |> Option.iter (fun bar ->
                let _ : obj = emitJsExpr (container, bar) "$0.insertBefore($1, $0.firstChild)"
                ())

    member _.ToggleFindReplace() =
        if findBar |> Option.exists (fun bar -> getAttribute bar "data-hidden" = "false") then closeFindBar ()
        else showFindReplaceDialog ()

    /// Import CSV data
    member _.ImportCSV(csvText: string, startRow: int, startCol: int) =
        // The session retains a persistent pre-import workbook snapshot.
        model <- importCSV csvText startRow startCol model
        recalculateFormulas()

        model <- pushCommand "Import CSV" model
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
        reportRejection
            ((importWorkbookPromise buffer).``then``(fun imported ->
                // Land on the requested sheet, else the first one Excel shows.
                let active =
                    match sheetName with
                    | Some n -> imported.Sheets |> Array.tryFindIndex (fun sheet -> sheet.Name = n) |> Option.defaultValue -1
                    | None -> -1
                // The imported context (defined names, tables, hidden sheets) is
                // what lets later recalculation resolve cross-sheet references.
                this.InstallImportedWorkbook(imported, active)))
            (reportError "ImportXlsx")

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

    /// The formula or value of the active cell, as it would be typed: what
    /// an entry in it starts with. "" for a cell a dynamic array spilled
    /// into - the value it shows is not its own (see GetGhostFormula).
    member _.GetCurrentCellFormula(): string =
        let active = activeCellAddress()
        if Map.containsKey active model.Spill then "" else SpreadsheetInput.editTextFor model.Date1904 (getCell active model)

    /// For a cell a dynamic array spilled into: the formula of the array's
    /// anchor cell, which Excel's formula bar shows greyed for such a cell
    /// (it cannot be edited there). null for any other cell.
    member _.GetGhostFormula(): string =
        let active = activeCellAddress()
        match Map.tryFind active model.Spill with
        | Some anchor -> SpreadsheetInput.editTextFor model.Date1904 (getCell anchor model)
        | None -> null

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

    /// Quick sort, ascending: the block of data around the active cell
    /// (Excel's current region - bounded by empty rows and columns) is
    /// sorted by the active cell's column. Only the cells of the block move.
    /// Its first row stays when it is a header: a table's header row, a row
    /// of labels over values of another kind (text over numbers), or a row
    /// formatted unlike the row under it. Blanks are last. One undo step;
    /// nothing happens when there is no data around the cell. Throws like
    /// SortRange for merged cells or a dynamic array in the block.
    member _.SortCurrentColumnAsc() =
        let col = (activeCellAddress()).Col
        sortActiveSheet "SortCurrentColumnAsc" col true (sprintf "Sort %s ascending" (columnToLetter col))
        render()

    /// Quick sort, descending (see SortCurrentColumnAsc); blanks are last
    /// here too.
    member _.SortCurrentColumnDesc() =
        let col = (activeCellAddress()).Col
        sortActiveSheet "SortCurrentColumnDesc" col false (sprintf "Sort %s descending" (columnToLetter col))
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


    // ===== Sort and filter =====

    /// Sort the rows of a range of the active sheet, as one undo step.
    /// `range`: A1 ("A1:D20"; whole columns "A:D" are clipped to the used rows).
    /// `keys`: [{ column, descending?, by?, color? }] - the first key decides,
    /// the next ones break ties, rows equal on every key keep their order.
    /// `column` is a 0-based sheet column or its letters ("C") inside the
    /// range; `by` is "value" (default), "cellColor" or "fontColor", and with
    /// a colour sort `color` ("#RRGGBB", null = no fill / automatic) is the
    /// colour whose rows come first (last with `descending`).
    /// `options`: { hasHeader?: boolean (the first row stays), caseSensitive?: boolean }.
    /// Values sort as in Excel: numbers, then text, FALSE, TRUE, errors;
    /// descending reverses that; blank cells are last either way.
    /// Only the cells of the range move, as whole rows of it, each with its
    /// value, format, comment and single-cell hyperlink. A formula that moves
    /// is re-pointed like a copy of it (relative references shift with the
    /// row, `$` rows stay); formulas elsewhere are not re-pointed. Row
    /// heights, validations and conditional formats stay in place, and so do
    /// the rows an AutoFilter hides. Throws for a bad range or key, a key
    /// column outside the range, merged cells in the rows to sort, or a
    /// dynamic array (spill) among them.
    member _.SortRange(range: obj, keys: obj, ?options: obj) : unit =
        let area = Host.rangeArg "SortRange" (requiredArg "SortRange" "range" range)
        let sortKeys = Host.sortKeysOfObj "SortRange" model keys
        let sortOptions = Host.sortOptionsOfObj "SortRange" (match options with Some value -> value | None -> null)
        if sortBlock "SortRange" area sortKeys sortOptions (sprintf "Sort %s" (SheetOps.rangeText area)) then render()

    /// Give the active sheet an AutoFilter over `range` (A1; its first row is
    /// the header row, the rows below are what gets filtered), replacing the
    /// one it has: column filters inside the new range stay, the others go.
    /// A sheet has one AutoFilter. One undo step. Throws for a bad range.
    /// (A table's filter buttons are not this filter; filter a table by
    /// setting the AutoFilter over the table's range.)
    member _.SetAutoFilter(range: obj) : unit =
        let area = SheetOps.autoFilterRange "SetAutoFilter" (Host.rangeArg "SetAutoFilter" (requiredArg "SetAutoFilter" "range" range)) model
        if changeFilter "Set filter" (setAutoFilter (Some area)) then render()
    /// Set the header range and one column's criteria together as a single
    /// undoable change. Validate every argument before replacing the old
    /// AutoFilter, so a rejected criterion cannot leave a new range behind.
    /// Existing criteria on columns within the new range are retained.
    member _.SetAutoFilterAndFilter(range: obj, column: obj, criteria: obj) : unit =
        let name = "SetAutoFilterAndFilter"
        let area = SheetOps.autoFilterRange name (Host.rangeArg name (requiredArg name "range" range)) model
        let col = Host.columnArg name "column" model (requiredArg name "column" column)
        if col < area.From.Col || col > area.To.Col then
            Host.hostFail name (sprintf "column %s is outside the filter range %s" (columnToLetter col) (SheetOps.rangeText area))
        let given = requiredArg name "criteria" criteria
        let wanted = SheetOps.filterOfDescriptionIn model.Date1904 name (Host.filterOfObj name given)
        let copied = Host.cloneGiven given
        if changeFilter (sprintf "Filter column %s" (columnToLetter col))
                (setAutoFilterAndColumnFilter area col wanted) then render()
        filterGiven.[sprintf "%s|%d" (sheetIdForIndex activeSheetIndex) col] <- (wanted, copied)


    /// Remove the active sheet's AutoFilter with its column filters; every
    /// row it hid is shown again. One undo step.
    member _.ClearAutoFilter() : unit =
        if changeFilter "Remove filter" (setAutoFilter None) then render()

    /// Filter a column of the AutoFilter: the rows whose cell in `column` (a
    /// 0-based sheet column or its letters) fails `criteria` are hidden -
    /// hidden, not removed: their cells stay, and copy, cut, fill, clear and
    /// format commands on a selection skip them. Rows hidden by the user
    /// (HideRows) are another matter and stay hidden. `criteria` is one of
    ///   { values: string[], blanks?: boolean }    cells showing one of the texts (case-insensitive), and blank cells
    ///   { conditions: [{ operator, value }], and?: boolean }   one or two, joined by AND (default) or OR;
    ///       operator "equals" | "notEquals" | "greaterThan" | "greaterOrEqual" | "lessThan" | "lessOrEqual"
    ///                | "contains" | "notContains" | "beginsWith" | "endsWith";
    ///       value: number, text ("equals" / "notEquals" take the wildcards * and ?), Date or "2024-03-31"
    ///   { top: { count, percent?: boolean, bottom?: boolean } }   the largest / smallest numbers (ties included)
    ///   { color: "#RRGGBB" | null, by?: "cellColor" | "fontColor" }
    ///   { dynamic: "aboveAverage" | "belowAverage" }
    /// It replaces the column's filter; the filters of several columns
    /// combine (a row has to pass all). The rows are worked out now: a later
    /// edit of a cell does not hide or show its row until ReapplyFilter().
    /// Without an AutoFilter the sheet gets one over the block of data
    /// around the column. One undo step. Throws for a column outside the
    /// AutoFilter's range and for criteria that are not understood.
    member _.SetFilter(column: obj, criteria: obj) : unit =
        let col = Host.columnArg "SetFilter" "column" model (requiredArg "SetFilter" "column" column)
        let wanted = SheetOps.filterOfDescriptionIn model.Date1904 "SetFilter" (Host.filterOfObj "SetFilter" (requiredArg "SetFilter" "criteria" criteria))
        match model.AutoFilter with
        | Some filter when not filter.Implicit && (col < filter.Range.From.Col || col > filter.Range.To.Col) ->
            Host.hostFail "SetFilter" (sprintf "column %s is outside the filter range %s" (columnToLetter col) (SheetOps.rangeText filter.Range))
        | _ -> ()
        // (What the host gave is what GetFilters hands back for as long as
        // the column's filter is this one - through undo and redo too.)
        filterGiven.[sprintf "%s|%d" (sheetIdForIndex activeSheetIndex) col] <- (wanted, Host.cloneGiven criteria)
        if changeFilter (sprintf "Filter column %s" (columnToLetter col)) (setColumnFilter col wanted) then render()

    /// Take the filter off `column` (0-based or letters), or off every column
    /// when called without one; the AutoFilter itself stays (ClearAutoFilter
    /// removes it). One undo step.
    member _.ClearFilter(?column: obj) : unit =
        let col =
            match column with
            | Some value when not (Host.isNullish value) -> Some (Host.columnArg "ClearFilter" "column" model value)
            | _ -> None
        if changeFilter "Clear filter" (clearColumnFilter col) then render()

    /// Apply the column filters again to the cells as they are now (after
    /// edits, a paste or a sort changed what the rows hold). One undo step
    /// when the hidden rows change.
    member _.ReapplyFilter() : unit =
        if changeFilter "Reapply filter" reapplyFilter then render()

    /// The AutoFilter of a sheet (the active one by default), or null:
    /// { range: "A1:D20", columns: [{ column (0-based), letter, criteria }],
    ///   hiddenRows: number[] (0-based rows the filter hides), hiddenRowCount,
    ///   visibleRowCount (rows of the range below the header that are shown) }.
    /// `criteria` is what SetFilter was given for the column - the same keys
    /// and the same types (a number stays a number, a text a text, a Date a
    /// Date; nothing is added), as a copy. For a filter that did not come
    /// through SetFilter (an opened file, the grid's own filter buttons) it
    /// has the shape SetFilter takes with the keys at their default left
    /// out, a comparison's operand a number when it is written as one, and
    /// { kept: true } for a filter the editor cannot evaluate.
    member _.GetFilters(?sheetIndex: obj) : obj =
        let sheetAt =
            match sheetIndex with
            | Some index when not (Host.isNullish index) ->
                let index = Host.intArg "GetFilters" "sheetIndex" index
                if index < 0 || index >= sheets.Length then Host.hostFail "GetFilters" (sprintf "there is no sheet %d (0 to %d)" index (sheets.Length - 1))
                index
            | _ -> activeSheetIndex
        let sheet = if sheetAt = activeSheetIndex then model else sheetModelAt sheetAt
        let criteriaOf (col: int) (criteria: FilterCriteria) : obj =
            match Portable.Dict.tryFind (sprintf "%s|%d" (sheetIdForIndex sheetAt) col) filterGiven with
            | ValueSome (set, given) when set = criteria -> Host.cloneGiven given
            | _ -> Host.filterToObj criteria
        match sheet.AutoFilter with
        | None -> null
        | Some filter ->
            let rows = match filterDataRows sheet with Some (first, last) -> last - first + 1 | None -> 0
            createObj [
                "range" ==> SheetOps.rangeText filter.Range
                "columns" ==> [| for KeyValue(col, columnFilter) in sheet.Filters ->
                                    createObj [ "column" ==> col; "letter" ==> columnToLetter col; "criteria" ==> criteriaOf col columnFilter.Criteria ] |]
                // (A plain array: an int array is a typed array in JS.)
                "hiddenRows" ==> [| for row in sheet.FilterHiddenRows -> box row |]
                "hiddenRowCount" ==> sheet.FilterHiddenRows.Length
                "visibleRowCount" ==> max 0 (rows - sheet.FilterHiddenRows.Length)
            ]

    /// What a filter's checklist offers for a column (0-based or letters) of
    /// the active sheet: { values: string[] (the distinct texts its cells
    /// show in the filter's rows, numbers first and in order), blanks:
    /// boolean (it has blank cells there) }.
    member _.GetFilterValues(column: obj) : obj =
        let col = Host.columnArg "GetFilterValues" "column" model (requiredArg "GetFilterValues" "column" column)
        let values, blanks = filterColumnValues col model
        createObj [ "values" ==> Array.ofList values; "blanks" ==> blanks ]

/// Export for JavaScript
[<ExportDefault>]
let createSpreadsheet (container: Element): Spreadsheet =
    new Spreadsheet(container)
