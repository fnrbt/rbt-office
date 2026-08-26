module Rbt.FableUI.OfficeUiDialog

#if FABLE_COMPILER

// Backend-specific renderer for `Rbt.Office.Ui.DialogSpec` values.
// Native rendering lives in OfficeShell.fs; this module is the Fable/DOM
// twin so apps can hand the same DialogSpec to both shells.
//
// Usage from a Fable shell:
//
//     let dialogState = OfficeUiDialog.createState ()
//
//     // Wire menu/toolbar actions to open dialogs from Rbt.Word.Ui:
//     let openFind () =
//         WordUi.findReplaceDialog onFindNext onReplaceAll
//         |> OfficeUiDialog.openDialog dialogState
//
//     // Render the overlay once, alongside the rest of the shell tree:
//     div [] [
//         shellChrome ()
//         documentArea ()
//         OfficeUiDialog.renderDialog dialogState
//     ]
//
// Drop `OfficeUiDialog.defaultStylesheet` into a <style> block (or copy
// the rules into your existing CSS) for the out-of-the-box look.

open System.Collections.Generic
open Rbt.Office.Ui
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom

/// Captures the open dialog plus per-field value signals so the renderer
/// can read and write field values reactively.
type OpenDialogState(spec: DialogSpec, values: Dictionary<string, Signal<string>>) =
    member _.Spec = spec
    member _.Values = values

type DialogState = {
    Open: Signal<OpenDialogState option>
}

let createState () : DialogState = { Open = signal None }

let openDialog (state: DialogState) (spec: DialogSpec) =
    let values = Dictionary<string, Signal<string>>()
    for field in spec.Fields do
        values.[field.Id] <- signal field.Initial
    state.Open.Value <- Some (OpenDialogState(spec, values))

let closeDialog (state: DialogState) =
    state.Open.Value <- None

// ── Field renderers ──────────────────────────────────────────

let private renderField (values: Dictionary<string, Signal<string>>) (field: DialogField) =
    let valueSignal = values.[field.Id]
    let writeValue v = valueSignal.Value <- v
    let fieldTestId = sprintf "dlg-field-%s" field.Id
    match field.Kind with
    | DialogTextField placeholder ->
        div [ cls "dlg-field" ] [
            label [ cls "dlg-label" ] [ text field.Label ]
            input [
                cls "dlg-input"
                attr "data-testid" fieldTestId
                typ "text"
                attr "placeholder" placeholder
                propR "value" (fun () -> box valueSignal.Value)
                onInput writeValue
            ]
        ]
    | DialogNumberField(minV, maxV) ->
        let extraAttrs =
            (minV |> Option.map (fun v -> attr "min" v) |> Option.toList)
            @ (maxV |> Option.map (fun v -> attr "max" v) |> Option.toList)
        div [ cls "dlg-field" ] [
            label [ cls "dlg-label" ] [ text field.Label ]
            input ([
                cls "dlg-input"
                attr "data-testid" fieldTestId
                typ "number"
                propR "value" (fun () -> box valueSignal.Value)
                onInput writeValue
            ] @ extraAttrs)
        ]
    | DialogCheckboxField ->
        let toggle (event: Browser.Types.Event) =
            let target = event.target :?> Browser.Types.HTMLInputElement
            valueSignal.Value <- if target.``checked`` then "true" else "false"
        div [ cls "dlg-field"; cls "dlg-field-checkbox" ] [
            label [ cls "dlg-label-inline" ] [
                input [
                    cls "dlg-checkbox"
                    attr "data-testid" fieldTestId
                    typ "checkbox"
                    propR "checked" (fun () -> box (valueSignal.Value = "true"))
                    onChange toggle
                ]
                span [] [ text field.Label ]
            ]
        ]
    | DialogSelectField options ->
        let onSelect (event: Browser.Types.Event) =
            let target = event.target :?> Browser.Types.HTMLSelectElement
            valueSignal.Value <- target.value
        div [ cls "dlg-field" ] [
            label [ cls "dlg-label" ] [ text field.Label ]
            select [
                cls "dlg-select"
                attr "data-testid" fieldTestId
                propR "value" (fun () -> box valueSignal.Value)
                onChange onSelect
            ] (options |> List.map (fun choice ->
                option [ attr "value" choice.Value ] [ text choice.Label ]))
        ]
    | DialogStatusField ->
        div [ cls "dlg-field"; cls "dlg-field-status" ] [
            label [ cls "dlg-label" ] [ text field.Label ]
            div [ cls "dlg-status"; attr "data-testid" fieldTestId ] [
                textR (fun () -> valueSignal.Value)
            ]
        ]
    | DialogComputedField compute ->
        // Live preview: recompute from every field's current value on change.
        div [ cls "dlg-field"; cls "dlg-field-status" ] [
            label [ cls "dlg-label" ] [ text field.Label ]
            div [ cls "dlg-status"; cls "dlg-preview"; attr "data-testid" fieldTestId ] [
                textR (fun () ->
                    let vals = values |> Seq.map (fun kv -> kv.Key, kv.Value.Value) |> Map.ofSeq
                    compute vals)
            ]
        ]

let private renderButton (state: DialogState) (values: Dictionary<string, Signal<string>>) (btn: DialogButton) =
    let activate (_: Browser.Types.MouseEvent) =
        let snapshot =
            values
            |> Seq.map (fun kv -> kv.Key, kv.Value.Value)
            |> Map.ofSeq
        let shouldClose = btn.Action snapshot
        if shouldClose then closeDialog state

    let primaryClass =
        if btn.IsPrimary then [ cls "dlg-button-primary" ] else []
    button (
        [ cls "dlg-button"; attr "data-testid" btn.TestId; onClick activate ]
        @ primaryClass)
        [ text btn.Label ]

/// Render the dialog overlay. The caller mounts this once and the renderer
/// becomes active/inactive based on the DialogState's open-signal.
let renderDialog (state: DialogState) : View =
    showIf (fun () -> state.Open.Value.IsSome) (
        div [ cls "dlg-overlay"; attr "role" "dialog"; attr "aria-modal" "true" ] [
            ReactiveFragment(fun () ->
                match state.Open.Value with
                | None -> []
                | Some openState ->
                    let spec = openState.Spec
                    let values = openState.Values
                    [
                        div [
                            cls "dlg-card"
                            attr "data-testid" spec.TestId
                        ] [
                            div [ cls "dlg-header" ] [
                                h2 [ cls "dlg-title" ] [ text spec.Title ]
                                (match spec.Subtitle with
                                 | Some s -> p [ cls "dlg-subtitle" ] [ text s ]
                                 | None -> fragment [])
                            ]
                            div [ cls "dlg-fields" ] (spec.Fields |> List.map (renderField values))
                            div [ cls "dlg-buttons" ] (spec.Buttons |> List.map (renderButton state values))
                        ]
                    ])
        ])

/// Default stylesheet for the dialog. Apps may inline this in their host
/// document or replace it with their own CSS.
let defaultStylesheet = """
.rb-dlg-overlay {
    position: fixed; inset: 0;
    background: rgba(0,0,0,0.25);
    display: flex; align-items: center; justify-content: center;
    z-index: 1000;
}
.rb-dlg-card {
    background: #FFFFFF;
    border: 1px solid #DADCE0;
    border-radius: 8px;
    box-shadow: 0 8px 24px rgba(0,0,0,0.16);
    padding: 24px;
    min-width: 360px;
    max-width: 480px;
    font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
    color: #202124;
}
.rb-dlg-header { margin-bottom: 16px; }
.rb-dlg-title { margin: 0; font-size: 16px; font-weight: 600; }
.rb-dlg-subtitle { margin: 4px 0 0; font-size: 12px; color: #5F6368; }
.rb-dlg-fields { display: flex; flex-direction: column; gap: 14px; }
.rb-dlg-field { display: flex; flex-direction: column; gap: 4px; }
.rb-dlg-field-checkbox { gap: 0; }
.rb-dlg-label { font-size: 12px; color: #5F6368; }
.rb-dlg-label-inline { display: inline-flex; align-items: center; gap: 8px; font-size: 13px; color: #202124; cursor: pointer; }
.rb-dlg-input, .rb-dlg-select {
    height: 30px; padding: 0 8px;
    border: 1px solid #DADCE0; border-radius: 4px;
    font-size: 13px; color: #202124; background: #FFFFFF;
    outline: none;
}
.rb-dlg-input:focus, .rb-dlg-select:focus { border-color: #1A73E8; }
.rb-dlg-status { font-size: 13px; color: #5F6368; min-height: 18px; }
.rb-dlg-buttons {
    display: flex; justify-content: flex-end; gap: 8px;
    margin-top: 18px;
}
.rb-dlg-button {
    height: 32px; padding: 0 16px;
    border: 1px solid #DADCE0; border-radius: 4px;
    background: #FFFFFF; color: #202124;
    font-size: 13px; cursor: pointer;
}
.rb-dlg-button:hover { background: #F1F3F4; }
.rb-dlg-button-primary {
    background: #1A73E8; color: #FFFFFF; border-color: #1A73E8;
}
.rb-dlg-button-primary:hover { background: #1764C6; }
"""

#endif
