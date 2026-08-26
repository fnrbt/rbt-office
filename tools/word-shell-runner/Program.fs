module WordShellRunner

#if FABLE_COMPILER

// Minimal Fable page that exercises the OfficeUi dialog framework
// (`Rbt.FableUI.OfficeUiDialog`) end-to-end in a real browser. Used
// by `scripts/check-shell-runner.mjs` under Playwright to verify the
// dialogs that the native `tools/word-editor` already drives.
//
// The page itself is intentionally barebones — four buttons that open
// each `WordUi.*Dialog` and a status line that records the result of
// the last submitted action. Playwright reads `window.rbDialog.*` and
// the dialog's data-testid attributes.

open System
open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom
open Rbt.FableUI.OfficeUiDialog

module WordUi = Rbt.Word.Ui

[<Emit("window[$0] = $1")>]
let private setWindowValue (_name: string) (_value: obj) : unit = jsNative

let private resultSignal = signal "(no action yet)"

let private record message =
    resultSignal.Value <- message

let private openFind dialogState () =
    WordUi.findReplaceDialog
        (fun q _ ->
            record (sprintf "find:%s" q)
            "")
        (fun q r _ ->
            record (sprintf "replace:%s->%s" q r)
            "")
    |> openDialog dialogState

let private openInsertTable dialogState () =
    WordUi.insertTableDialog (fun rows cols ->
        record (sprintf "table:%dx%d" rows cols))
    |> openDialog dialogState

let private openLink dialogState () =
    WordUi.insertLinkDialog (fun url ->
        record (sprintf "link:%s" url))
    |> openDialog dialogState

let private openPageSetup dialogState () =
    WordUi.pageSetupDialog "portrait" 72.0 72.0 72.0 72.0 (fun (orient, ml, mr, mt, mb) ->
        record (sprintf "page:%s ml=%g mr=%g mt=%g mb=%g" orient ml mr mt mb))
    |> openDialog dialogState

let private openImage dialogState () =
    WordUi.insertImageDialog (fun src ->
        record (sprintf "image:%s" src))
    |> openDialog dialogState

let private boot () =
    let dialogState = createState ()

    // Inject the default dialog stylesheet and a tiny page style.
    let style = document.createElement("style")
    style.textContent <- defaultStylesheet + """
body { font-family: system-ui, sans-serif; margin: 32px; color: #202124; background: #FFFFFF; }
h1 { font-size: 18px; margin: 0 0 16px; }
.toolbar { display: flex; gap: 8px; margin-bottom: 16px; }
.toolbar button {
    padding: 8px 14px; border: 1px solid #DADCE0; border-radius: 4px;
    background: #FFFFFF; color: #202124; font-size: 13px; cursor: pointer;
}
.toolbar button:hover { background: #F1F3F4; }
#result {
    padding: 12px; background: #F1F3F4; border-radius: 4px;
    font-family: ui-monospace, monospace; font-size: 13px; color: #5F6368;
    min-height: 1.2em;
}
"""
    document.head.appendChild(style) |> ignore

    // Render the shell: one button per dialog + a status line + the
    // dialog overlay slot (visible only when a dialog is open).
    let shell =
        div [] [
            h2 [] [ text "OfficeUi dialog runner" ]
            div [ cls "toolbar" ] [
                button [ attr "data-testid" "btn-find"; onClick (fun _ -> openFind dialogState ()) ] [ text "Find / Replace" ]
                button [ attr "data-testid" "btn-table"; onClick (fun _ -> openInsertTable dialogState ()) ] [ text "Insert Table" ]
                button [ attr "data-testid" "btn-link"; onClick (fun _ -> openLink dialogState ()) ] [ text "Insert Link" ]
                button [ attr "data-testid" "btn-page-setup"; onClick (fun _ -> openPageSetup dialogState ()) ] [ text "Page Setup" ]
                button [ attr "data-testid" "btn-image"; onClick (fun _ -> openImage dialogState ()) ] [ text "Insert Image" ]
            ]
            div [ attr "id" "result"; attr "data-testid" "result" ] [
                textR (fun () -> resultSignal.Value)
            ]
            renderDialog dialogState
        ]

    mountApp "#app" shell

    // Playwright hooks. The data-testid attributes are also sufficient
    // for most flows; these are a fallback for hard-to-target cases.
    setWindowValue "rbDialog" (createObj [
        "openFind" ==> openFind dialogState
        "openInsertTable" ==> openInsertTable dialogState
        "openLink" ==> openLink dialogState
        "openPageSetup" ==> openPageSetup dialogState
        "openImage" ==> openImage dialogState
        "lastResult" ==> fun () -> resultSignal.Value
        "close" ==> fun () -> closeDialog dialogState
    ])

boot ()

#endif
