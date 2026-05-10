module BrowserRunner

open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Browser.Types
open ZipInterop

module WordApi = RedBat.Word.API
module Formats = Main

type RunnerState = {
    Canvas: HTMLCanvasElement
    Status: HTMLElement
    mutable Editor: WordApi.Editor option
}

[<Emit("fetch($0).then(function(r) { if (!r.ok) throw new Error('fetch failed ' + r.status + ': ' + $0); return r.arrayBuffer(); })")>]
let private fetchArrayBuffer (_url: string) : JS.Promise<obj> = jsNative

[<Emit("JSON.stringify($0, null, 2)")>]
let private stringify (_value: obj) : string = jsNative

[<Emit("Array.from($0)")>]
let private arrayFrom (_value: obj) : obj array = jsNative

[<Emit("window[$0] = $1")>]
let private setWindowValue (_name: string) (_value: obj) : unit = jsNative

[<Emit("$0.toDataURL('image/png')")>]
let private canvasToPngDataUrl (_canvas: HTMLCanvasElement) : string = jsNative

[<Emit("new Blob([$0], { type: $1 })")>]
let private createBlob (_text: string) (_mimeType: string) : Blob = jsNative

[<Emit("URL.createObjectURL($0)")>]
let private createObjectUrl (_blob: Blob) : string = jsNative

[<Emit("setTimeout($0, $1)")>]
let private setTimeout (_callback: unit -> unit) (_ms: int) : unit = jsNative

[<Emit("$0.style.width = $1")>]
let private setStyleWidth (_element: HTMLElement) (_value: string) : unit = jsNative

[<Emit("$0.style.height = $1")>]
let private setStyleHeight (_element: HTMLElement) (_value: string) : unit = jsNative

[<Emit("$0.download = $1")>]
let private setDownload (_element: HTMLAnchorElement) (_value: string) : unit = jsNative

[<Emit("new URLSearchParams(window.location.search).get($0)")>]
let private getQueryParam (_name: string) : string = jsNative

[<Emit("(document.fonts ? document.fonts.ready.then(function(){}) : Promise.resolve())")>]
let private documentFontsReady () : JS.Promise<unit> = jsNative

[<Emit("$0.catch($1)")>]
let private promiseCatch (_promise: JS.Promise<'a>) (_handler: obj -> obj) : JS.Promise<'a> = jsNative

let private queryParam name =
    let value = getQueryParam name
    if isNull value || value = "" then None else Some value

let private status (state: RunnerState) text =
    state.Status.textContent <- text

let private ensureEditor (state: RunnerState) =
    match state.Editor with
    | Some editor -> editor
    | None ->
        let editor = WordApi.initializeCanvas state.Canvas
        state.Editor <- Some editor
        editor

let private setCanvasSize (state: RunnerState) (width: int) (height: int) =
    state.Canvas.width <- width
    state.Canvas.height <- height
    setStyleWidth (state.Canvas :?> HTMLElement) (sprintf "%dpx" width)
    setStyleHeight (state.Canvas :?> HTMLElement) (sprintf "%dpx" height)

let private getLayoutSegments (state: RunnerState) =
    let editor = ensureEditor state
    WordApi.getLayoutSegments editor

let private download name text mimeType =
    let blob = createBlob text mimeType
    let url = createObjectUrl blob
    let a = document.createElement("a") :?> HTMLAnchorElement
    a.href <- url
    setDownload a name
    a.click()

let private loadArrayBuffer (state: RunnerState) sourceName (buffer: obj) =
    promise {
        status state (sprintf "Loading %s..." sourceName)
        do! documentFontsReady ()
        let editor = ensureEditor state
        let! package = Formats.Processors.forEditorPackage buffer
        WordApi.importFromDocxPackage editor package
        setTimeout (fun () ->
            let segments = getLayoutSegments state
            status state (sprintf "Loaded %s: %d layout segments" sourceName segments.Length)) 0
        return package
    }

let private loadDocxUrl (state: RunnerState) url =
    promise {
        let! buffer = fetchArrayBuffer url
        return! loadArrayBuffer state url buffer
    }

let private loadFile (state: RunnerState) (file: File) =
    promise {
        let! buffer = file.arrayBuffer()
        return! loadArrayBuffer state file.name buffer
    }

let private selectedFile (input: HTMLInputElement) =
    if input.files.length > 0 then Some input.files.[0] else None

let private renderShell () =
    document.body.innerHTML <- """
<div class="toolbar">
  <input id="docx-file" type="file" accept=".docx" />
  <input id="docx-url" type="text" placeholder="DOCX URL or path" />
  <button id="load-url">Load URL</button>
  <button id="dump-layout">Download Layout JSON</button>
  <button id="dump-png">Download PNG</button>
  <span id="status">Starting...</span>
</div>
<main>
  <canvas id="word-canvas" width="900" height="1600"></canvas>
</main>
"""

let private installStyles () =
    let style = document.createElement("style")
    let fontRoot = queryParam "fontRoot" |> Option.map (fun root -> root.TrimEnd([|'/'|]))
    let wordFonts =
        match fontRoot with
        | Some root ->
            let face family file weight style =
                sprintf "@font-face { font-family: \"%s\"; src: url(\"%s/%s\") format(\"truetype\"); font-weight: %d; font-style: %s; }" family root file weight style
            String.concat "\n" [
                face "Calibri" "calibri.ttf" 400 "normal"
                face "Calibri" "calibrib.ttf" 700 "normal"
                face "Calibri" "calibrii.ttf" 400 "italic"
                face "Calibri" "calibriz.ttf" 700 "italic"
                face "Arial" "arial.ttf" 400 "normal"
                face "Arial" "arialbd.ttf" 700 "normal"
                face "Arial" "ariali.ttf" 400 "italic"
                face "Arial" "arialbi.ttf" 700 "italic"
                face "Times New Roman" "times.ttf" 400 "normal"
                face "Times New Roman" "timesbd.ttf" 700 "normal"
                face "Times New Roman" "timesi.ttf" 400 "italic"
                face "Times New Roman" "timesbi.ttf" 700 "italic"
                ""
            ]
        | None -> ""
    style.textContent <- wordFonts + """
html, body { margin: 0; min-height: 100%; font-family: system-ui, sans-serif; background: #646b75; }
.toolbar { position: sticky; top: 0; z-index: 10; display: flex; gap: 8px; align-items: center; padding: 8px; background: #f3f4f6; border-bottom: 1px solid #cbd5e1; }
.toolbar input[type='text'] { width: 420px; max-width: 45vw; padding: 5px 7px; }
.toolbar button { padding: 5px 9px; }
#status { margin-left: auto; font-size: 12px; color: #334155; }
main { padding: 24px; }
#word-canvas { display: block; background: white; box-shadow: 0 2px 18px rgba(0,0,0,.25); }
"""
    document.head.appendChild(style) |> ignore

let private boot () =
    installStyles ()
    renderShell ()

    let canvas = document.getElementById("word-canvas") :?> HTMLCanvasElement
    let statusEl = document.getElementById("status") :?> HTMLElement
    let state = { Canvas = canvas; Status = statusEl; Editor = None }
    setCanvasSize state 900 1600
    let editor = ensureEditor state
    WordApi.setContent editor "Browser Word runner ready."
    status state "Ready"

    let fileInput = document.getElementById("docx-file") :?> HTMLInputElement
    fileInput.onchange <- fun _ ->
        match selectedFile fileInput with
        | Some file ->
            promiseCatch (loadFile state file) (fun err ->
                status state (sprintf "Load failed: %A" err)
                null)
            |> ignore
        | None -> ()

    let urlInput = document.getElementById("docx-url") :?> HTMLInputElement
    let loadUrlButton = document.getElementById("load-url") :?> HTMLButtonElement
    loadUrlButton.onclick <- fun _ ->
        if urlInput.value.Trim() <> "" then
            promiseCatch (loadDocxUrl state urlInput.value) (fun err ->
                status state (sprintf "Load failed: %A" err)
                null)
            |> ignore

    let dumpLayoutButton = document.getElementById("dump-layout") :?> HTMLButtonElement
    dumpLayoutButton.onclick <- fun _ ->
        getLayoutSegments state
        |> box
        |> stringify
        |> fun json -> download "browser-layout.json" json "application/json"

    let dumpPngButton = document.getElementById("dump-png") :?> HTMLButtonElement
    dumpPngButton.onclick <- fun _ ->
        let a = document.createElement("a") :?> HTMLAnchorElement
        a.href <- canvasToPngDataUrl state.Canvas
        setDownload a "browser-render.png"
        a.click()

    setWindowValue "rbWordRunner" (createObj [
        "loadDocxUrl" ==> fun (url: string) -> loadDocxUrl state url
        "getLayoutSegments" ==> fun () -> getLayoutSegments state
        "getLayoutJson" ==> fun () -> stringify (box (getLayoutSegments state))
        "clearSelection" ==> fun () -> ensureEditor state |> WordApi.clearSelection
        "setCanvasSize" ==> fun (width: int) (height: int) -> setCanvasSize state width height
        "canvasToPngDataUrl" ==> fun () -> canvasToPngDataUrl state.Canvas
        "editor" ==> fun () -> state.Editor
    ])

    match queryParam "doc" with
    | Some url ->
        urlInput.value <- url
        promiseCatch (loadDocxUrl state url) (fun err ->
            status state (sprintf "Load failed: %A" err)
            null)
        |> ignore
    | None -> ()

boot ()
