module WordShellRunner

#if FABLE_COMPILER

// Browser fixture for the shared OfficeUi renderer and the WordUi dialog specs.
// The focused fixture deliberately mixes every DialogFieldKind across tabs.
// `scripts/check-shell-runner.mjs` drives its actual DOM inputs with Playwright.

open System
open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom
open Rbt.FableUI.OfficeUiDialog

module WordUi = Rbt.Word.Ui
module Dialog = Rbt.Office.Ui
module WordApi = Rbt.Word.API

[<Emit("window[$0] = $1")>]
let private setWindowValue (_name: string) (_value: obj) : unit = jsNative

let private resultSignal = signal "(no action yet)"

let private record message =
    resultSignal.Value <- message

let private attempts = ResizeArray<string>()

let private fieldSnapshot (values: Map<string, string>) =
    values
    |> Map.toList
    |> List.map (fun (key, value) -> sprintf "%s=%s" key value)
    |> String.concat "|"

let private fixtureSpec dialogState : Dialog.DialogSpec =
    let preview values =
        sprintf "%s / %s / %s"
            (Map.find "name" values)
            (Map.find "notes" values)
            (Map.find "mode" values)
    {
        TestId = "fixture-dialog"
        Title = "All field kinds"
        Subtitle = Some "Change fields on both tabs before applying."
        // With Tabs present this field must never appear or reach the callback.
        Fields = [ Dialog.dialogText "ignored" "Ignored field" "" "do-not-submit" ]
        Tabs = [
            "Basics", [
                Dialog.dialogSection "Identity"
                Dialog.dialogText "name" "Name" "Enter name" "Initial"
                Dialog.dialogNumber "quantity" "Quantity" (Some "1") (Some "20") "2"
                Dialog.dialogCheckbox "enabled" "Enabled" false
                Dialog.dialogSelect "size" "Size"
                    [ Dialog.choice "small" "Small"; Dialog.choice "large" "Large" ] "large"
                Dialog.dialogStatus "status" "Status" "Ready"
            ]
            "Details", [
                Dialog.dialogColor "color" "Color" "#336699"
                Dialog.dialogSlider "strength" "Strength" 0.0 10.0 2.0
                Dialog.dialogTextArea "notes" "Notes" "Write notes" 3 "Original"
                Dialog.dialogRadio "mode" "Mode"
                    [ Dialog.choice "plain" "Plain"; Dialog.choice "rich" "Rich" ] "plain"
                Dialog.dialogComputed "preview" "Preview" preview
            ]
        ]
        Buttons = [
            Dialog.dialogCancel
            Dialog.dialogButton "fixture-replace" "Open link dialog" false (fun _ ->
                WordUi.insertLinkDialog (fun url -> record (sprintf "replacement-link:%s" url))
                |> openDialog dialogState
                true)
            Dialog.dialogButton "fixture-apply" "Apply" true (fun values ->
                let snapshot = fieldSnapshot values
                attempts.Add snapshot
                record (sprintf "fixture:%d:%s" attempts.Count snapshot)
                Map.find "name" values = "Accept")
        ]
    }

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

let private openFixture dialogState () =
    fixtureSpec dialogState |> openDialog dialogState

let private boot () =
    let dialogState = createState ()

    // Inject the default dialog stylesheet and a tiny page style.
    let style = document.createElement("style")
    style.textContent <- defaultStylesheet + """
body { font-family: system-ui, sans-serif; margin: 32px; color: #202124; background: #FFFFFF; }
h1 { font-size: 18px; margin: 0 0 16px; }
.toolbar { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 16px; }
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
#word-fixture { display: block; width: 500px; height: 350px; margin-top: 16px; }
"""
    document.head.appendChild(style) |> ignore

    // A genuine Word editor, not a recorder: the link dialog below calls
    // the public editing API and Playwright checks its document/undo history.
    let canvas = document.createElement("canvas") :?> Browser.Types.HTMLCanvasElement
    canvas.id <- "word-fixture"
    canvas.width <- 500
    canvas.height <- 350
    canvas.tabIndex <- 0
    document.body.appendChild(canvas) |> ignore
    let editor = WordApi.initializeCanvas canvas
    WordApi.setContent editor "Start"
    let openEditorLink () =
        WordUi.insertLinkDialog (fun url ->
            WordApi.insertHyperlink editor url "linked"
            record (sprintf "editor-link:%s" url))
        |> openDialog dialogState
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
                button [ attr "data-testid" "btn-fixture"; onClick (fun _ -> openFixture dialogState ()) ] [ text "All fields" ]
                button [ attr "data-testid" "btn-editor-link"; onClick (fun _ -> openEditorLink ()) ] [ text "Insert link in editor" ]
            ]
            div [ attr "id" "result"; attr "data-testid" "result" ] [
                textR (fun () -> resultSignal.Value)
            ]
            renderDialog dialogState
        ]

    let disposeShell = mountApp "#app" shell

    // Playwright hooks. The data-testid attributes are also sufficient
    // for most flows; these are a fallback for hard-to-target cases.
    setWindowValue "rbDialog" (createObj [
        "openFind" ==> openFind dialogState
        "openInsertTable" ==> openInsertTable dialogState
        "openLink" ==> openLink dialogState
        "openPageSetup" ==> openPageSetup dialogState
        "openImage" ==> openImage dialogState
        "openFixture" ==> openFixture dialogState
        "openEditorLink" ==> openEditorLink
        "attempts" ==> fun () -> attempts.ToArray()
        "editorText" ==> fun () -> WordApi.getText editor
        "editorCanUndo" ==> fun () -> WordApi.canUndo editor
        "editorUndo" ==> fun () -> WordApi.undo editor
        "editorCanRedo" ==> fun () -> WordApi.canRedo editor
        "lastResult" ==> fun () -> resultSignal.Value
        "close" ==> fun () -> closeDialog dialogState
        "disposeShell" ==> disposeShell
    ])

boot ()

#endif
