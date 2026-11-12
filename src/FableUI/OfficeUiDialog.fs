module Rbt.FableUI.OfficeUiDialog

#if FABLE_COMPILER

open System.Collections.Generic
open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Browser.Types
open Rbt.Office.Ui
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom

/// The open spec and the values shared by all its tabs.
type OpenDialogState(spec: DialogSpec, values: Dictionary<string, Signal<string>>) =
    member _.Spec = spec
    member _.Values = values

type DialogState = { Open: Signal<OpenDialogState option> }

let createState () : DialogState = { Open = signal None }

let openDialog (state: DialogState) (spec: DialogSpec) =
    let values = Dictionary<string, Signal<string>>()
    let fields = if spec.Tabs.IsEmpty then spec.Fields else spec.Tabs |> List.collect snd
    for field in fields do
        values.[field.Id] <- signal field.Initial
    state.Open.Value <- Some (OpenDialogState(spec, values))

/// Reuse application-owned field signals when migrating a form without
/// maintaining a second copy of its values in the renderer.
let openDialogWithValues (state: DialogState) (spec: DialogSpec) (values: Dictionary<string, Signal<string>>) =
    state.Open.Value <- Some (OpenDialogState(spec, values))

let closeDialog (state: DialogState) = state.Open.Value <- None

[<Emit("setTimeout($0, 0)")>]
let private defer (f: unit -> unit) : int = jsNative

[<Emit("clearTimeout($0)")>]
let private cancelDeferred (handle: int) : unit = jsNative

[<Emit("$0.isConnected")>]
let private isConnected (element: Element) : bool = jsNative

// Follow native radio tab order as well as excluding hidden and disabled controls.
[<Emit("((nodes) => nodes.filter(e => e.getClientRects().length && e.tabIndex >= 0 && (e.type !== 'radio' || !e.name || e.checked || !nodes.some(r => r.type === 'radio' && r.name === e.name && r.form === e.form && r.checked && r.getClientRects().length && !r.disabled))))(Array.from($0.querySelectorAll('input:not([disabled]):not([type=hidden]), select:not([disabled]), textarea:not([disabled]), button:not([disabled]), [tabindex]:not([tabindex=\"-1\"])')))")>]
let private focusable (element: Element) : HTMLElement array = jsNative

let mutable private nextId = 0
let private freshId () =
    nextId <- nextId + 1
    sprintf "rb-dialog-%d" nextId

let private snapshot (values: Dictionary<string, Signal<string>>) =
    values |> Seq.map (fun kv -> kv.Key, kv.Value.Value) |> Map.ofSeq

let private renderField (prefix: string) (values: Dictionary<string, Signal<string>>) (field: DialogField) =
    let value = values.[field.Id]
    let fieldId = prefix + "-field-" + field.Id
    let testId = field.TestId |> Option.defaultValue ("dlg-field-" + field.Id)
    let kind, update =
        match field.Kind with
        | DialogUpdatingField (kind, update) -> kind, Some update
        | kind -> kind, None
    let write v =
        match update with
        | None -> value.Value <- v
        | Some change ->
            let next = change (snapshot values) v
            for KeyValue(id, text) in next do
                match values.TryGetValue id with
                | true, target when target.Value <> text -> target.Value <- text
                | _ -> ()
    let labelled control =
        div [ cls "dlg-field" ] [
            label [ cls "dlg-label"; attr "for" fieldId ] [ text field.Label ]
            control
        ]
    let common = [ id fieldId; attr "data-testid" testId ]
    let valueProp = propR "value" (fun () -> box value.Value)
    match kind with
    | DialogUpdatingField _ -> failwith "Nested dialog update is not supported"
    | DialogTextField placeholder ->
        labelled (input (common @ [ cls "dlg-input"; typ "text"; attr "placeholder" placeholder; valueProp; onInput write ]))
    | DialogNumberField (low, high) ->
        let bounds = (low |> Option.map (attr "min") |> Option.toList) @ (high |> Option.map (attr "max") |> Option.toList)
        labelled (input (common @ [ cls "dlg-input"; typ "number"; valueProp; onInput write ] @ bounds))
    | DialogTextAreaField (placeholder, lines) ->
        labelled (el "textarea" (common @ [ cls "dlg-textarea"; attr "placeholder" placeholder; attr "rows" (string (max 1 lines)); valueProp; onInput write ]) [])
    | DialogCheckboxField ->
        div [ cls "dlg-field"; cls "dlg-field-checkbox" ] [
            label [ cls "dlg-label-inline" ] [
                input (common @ [ cls "dlg-checkbox"; typ "checkbox"; propR "checked" (fun () -> box (value.Value = "true")); onChange (fun e -> write (if (e.target :?> HTMLInputElement).``checked`` then "true" else "false")) ])
                span [] [ text field.Label ]
            ]
        ]
    | DialogSelectField options ->
        labelled (select (common @ [ cls "dlg-select"; onChange (fun e -> write ((e.target :?> HTMLSelectElement).value)) ])
            (options |> List.map (fun choice ->
                option ([ attr "value" choice.Value; propR "selected" (fun () -> box (value.Value = choice.Value)) ]
                        @ (if choice.IsEnabled then [] else [ attr "disabled" "disabled" ])) [ text choice.Label ])))
    | DialogColorField ->
        labelled (input (common @ [ cls "dlg-color"; typ "color"; valueProp; onInput write ]))
    | DialogSliderField (low, high) ->
        let bounds = [ attr "min" (sprintf "%g" low); attr "max" (sprintf "%g" high); attr "step" "any" ]
        labelled (div [ cls "dlg-slider" ] [
            input (common @ [ typ "range"; valueProp; onInput write ] @ bounds)
            input ([ cls "dlg-input"; typ "number"; attr "aria-label" field.Label; valueProp; onInput write ] @ bounds)
        ])
    | DialogRadioField options ->
        div [ cls "dlg-field" ] [
            div [ cls "dlg-label"; id (fieldId + "-label") ] [ text field.Label ]
            div [ cls "dlg-radios"; attr "role" "radiogroup"; attr "aria-labelledby" (fieldId + "-label"); attr "data-testid" testId ]
                (options |> List.map (fun choice ->
                    label [ cls "dlg-label-inline" ] [
                        input [ typ "radio"; attr "name" fieldId; attr "value" choice.Value
                                propR "checked" (fun () -> box (value.Value = choice.Value))
                                onChange (fun _ -> write choice.Value) ]
                        span [] [ text choice.Label ]
                    ]))
        ]
    | DialogSectionHeader ->
        div [ cls "dlg-section"; attr "data-testid" testId ] [ text field.Label ]
    | DialogStatusField ->
        div [ cls "dlg-field" ] [
            if field.Label <> "" then div [ cls "dlg-label"; id (fieldId + "-label") ] [ text field.Label ]
            div ([ cls "dlg-status"; attr "data-testid" testId ]
                 @ (if field.Label = "" then [] else [ attr "aria-labelledby" (fieldId + "-label") ])
                 @ (if field.Id.EndsWith("-error") then [ attr "role" "alert" ] else []))
                [ textR (fun () -> value.Value) ]
        ]
    | DialogComputedField compute ->
        div [ cls "dlg-field" ] [
            div [ cls "dlg-label"; id (fieldId + "-label") ] [ text field.Label ]
            div [ cls "dlg-status"; cls "dlg-preview"; attr "data-testid" testId; attr "aria-labelledby" (fieldId + "-label") ] [ textR (fun () -> compute (snapshot values)) ]
        ]

let private isCurrent (state: DialogState) (opened: OpenDialogState) =
    match state.Open.Peek() with
    | Some current -> obj.ReferenceEquals(current, opened)
    | None -> false

let private activate (state: DialogState) (opened: OpenDialogState) (btn: DialogButton) =
    if isCurrent state opened then
        let shouldClose = btn.Action (snapshot opened.Values)
        // An action can synchronously replace this dialog.
        if shouldClose && isCurrent state opened then closeDialog state

/// Default CSS, which the host can replace with its own theme.
let defaultStylesheet = """
.rb-dlg-overlay { position: fixed; inset: 0; background: rgba(0,0,0,.25); display: flex; align-items: center; justify-content: center; padding: 16px; z-index: 1000; }
.rb-dlg-card { box-sizing: border-box; display: flex; flex-direction: column; background: var(--office-surface, #fff); border: 1px solid var(--office-border, #dadce0); border-radius: 10px; box-shadow: 0 12px 40px rgb(0 0 0 / 18%); padding: 24px; width: min(480px, 100%); max-height: 100%; font-family: system-ui, sans-serif; color: var(--office-text, #202124); }
.rb-dlg-header { margin-bottom: 16px; }
.rb-dlg-title { margin: 0; font-size: 16px; font-weight: 600; }
.rb-dlg-subtitle { margin: 4px 0 0; font-size: 13px; color: var(--office-muted, #5f6368); }
.rb-dlg-body { min-height: 0; overflow: auto; }
.rb-dlg-fields { display: flex; flex-direction: column; gap: 14px; }
.rb-dlg-panel[hidden] { display: none; }
.rb-dlg-tabs { display: flex; gap: 4px; overflow-x: auto; margin-bottom: 12px; }
.rb-dlg-tab { border: 0; border-bottom: 2px solid transparent; background: transparent; padding: 8px; cursor: pointer; white-space: nowrap; }
.rb-dlg-tab[aria-selected=true] { border-color: var(--office-accent, #1a73e8); color: var(--office-accent, #1a73e8); }
.rb-dlg-field { display: flex; flex-direction: column; gap: 4px; }
.rb-dlg-label { font-size: 13px; color: var(--office-muted, #5f6368); }
.rb-dlg-label-inline { display: inline-flex; align-items: center; gap: 8px; font-size: 13px; cursor: pointer; }
.rb-dlg-input, .rb-dlg-select, .rb-dlg-textarea { box-sizing: border-box; padding: 6px 8px; border: 1px solid var(--office-border, #dadce0); border-radius: var(--office-radius, 6px); font: inherit; color: var(--office-text, #202124); background: var(--office-surface, #fff); }
.rb-dlg-slider { display: flex; align-items: center; gap: 8px; }
.rb-dlg-slider input[type=range] { flex: 1; min-width: 0; }
.rb-dlg-slider input[type=number] { width: 80px; }
.rb-dlg-radios { display: flex; flex-direction: column; gap: 4px; }
.rb-dlg-section { font-weight: 600; border-bottom: 1px solid var(--office-border, #dadce0); padding-bottom: 4px; }
.rb-dlg-status { font-size: 13px; color: var(--office-muted, #5f6368); min-height: 18px; white-space: pre-wrap; }
.rb-dlg-buttons { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; margin-top: 18px; }
.rb-dlg-button { min-height: 32px; padding: 0 16px; border: 1px solid var(--office-border, #dadce0); border-radius: var(--office-radius, 6px); background: var(--office-surface, #fff); color: var(--office-text, #202124); font: inherit; cursor: pointer; }
.rb-dlg-button:hover { background: var(--office-hover, #f1f3f4); }
.rb-office-dialog .rb-dlg-button-primary { background: var(--office-accent, #1a73e8); color: #fff; border-color: var(--office-accent, #1a73e8); }
.rb-office-dialog .rb-dlg-button-primary:hover { filter: brightness(.9); }
.rb-dlg-card :focus-visible, .rb-dlg-card:focus-visible { outline: 2px solid var(--office-accent, #1a73e8); outline-offset: 2px; }
"""
/// Mount once alongside the shell. Dialogs are created and disposed as the state changes.
let renderDialog (state: DialogState) : View =
    Chrome.installStyles "dialog" defaultStylesheet
    let mutable previousFocus: Element option = None
    ReactiveFragment(fun () ->
        match state.Open.Value with
        | None -> []
        | Some opened ->
            let spec = opened.Spec
            let prefix = freshId ()
            let activeTab = signal 0
            let titleId = prefix + "-title"
            let subtitleId = prefix + "-description"
            let selectTab index = activeTab.Value <- index
            let tabViews =
                if spec.Tabs.IsEmpty then
                    [ div [ cls "dlg-fields" ] (spec.Fields |> List.map (renderField prefix opened.Values)) ]
                else
                    let count = spec.Tabs.Length
                    let tabs = spec.Tabs |> List.mapi (fun index (name, _) ->
                        button [ cls "dlg-tab"; typ "button"; id (sprintf "%s-tab-%d" prefix index)
                                 attr "role" "tab"; attr "data-dlg-tab" name
                                 attr "aria-controls" (sprintf "%s-panel-%d" prefix index)
                                 attrR "aria-selected" (fun () -> if activeTab.Value = index then "true" else "false")
                                 attrR "tabindex" (fun () -> if activeTab.Value = index then "0" else "-1")
                                 onClick (fun _ -> selectTab index)
                                 onKeyDown (fun e ->
                                     let next =
                                         match e.key with
                                         | "ArrowRight" -> Some ((index + 1) % count)
                                         | "ArrowLeft" -> Some ((index + count - 1) % count)
                                         | "Home" -> Some 0
                                         | "End" -> Some (count - 1)
                                         | _ -> None
                                     match next with
                                     | Some i ->
                                         e.preventDefault ()
                                         selectTab i
                                         let target = document.getElementById(sprintf "%s-tab-%d" prefix i)
                                         if not (isNull target) then target.focus ()
                                     | None -> ()) ] [ text name ])
                    let panels = spec.Tabs |> List.mapi (fun index (_, fields) ->
                        div [ cls "dlg-fields"; cls "dlg-panel"; id (sprintf "%s-panel-%d" prefix index)
                              attr "role" "tabpanel"; attr "aria-labelledby" (sprintf "%s-tab-%d" prefix index)
                              attr "tabindex" "0"; propR "hidden" (fun () -> box (activeTab.Value <> index)) ]
                            (fields |> List.map (renderField prefix opened.Values)))
                    div [ cls "dlg-tabs"; attr "role" "tablist"; attr "aria-label" spec.Title ] tabs :: panels
            let keydown (e: KeyboardEvent) =
                e.stopPropagation ()
                let tag = (e.target :?> Element).tagName.ToLowerInvariant()
                let composing: bool = e?isComposing
                if e.key = "Tab" then
                    let card = e.currentTarget :?> Element
                    let controls = focusable card
                    if controls.Length = 0 then
                        e.preventDefault ()
                        (card :?> HTMLElement).focus ()
                    else
                        let current = document.activeElement
                        let index = controls |> Array.tryFindIndex (fun node -> (node :> Element) = current)
                        match index with
                        | None -> e.preventDefault (); (if e.shiftKey then controls.[controls.Length - 1] else controls.[0]).focus ()
                        | Some i when e.shiftKey && i = 0 -> e.preventDefault (); controls.[controls.Length - 1].focus ()
                        | Some i when not e.shiftKey && i = controls.Length - 1 -> e.preventDefault (); controls.[0].focus ()
                        | _ -> ()
                elif not composing && e.key = "Escape" then
                    e.preventDefault ()
                    match spec.Buttons |> List.tryFind (fun b -> b.TestId = "dlg-cancel" || b.Label = "Cancel" || b.Label = "Close") with
                    | Some cancel -> activate state opened cancel
                    | None -> if isCurrent state opened then closeDialog state
                elif not composing && e.key = "Enter" && not e.repeat && not e.ctrlKey && not e.altKey && not e.metaKey && not e.shiftKey
                     && tag <> "textarea" && tag <> "select" && tag <> "button"
                     && not (tag = "input" && ((e.target :?> HTMLInputElement).``type`` = "radio" || (e.target :?> HTMLInputElement).``type`` = "checkbox" || (e.target :?> HTMLInputElement).``type`` = "color")) then
                    match spec.Buttons |> List.tryFind (fun b -> b.IsPrimary) with
                    | Some primary -> e.preventDefault (); activate state opened primary
                    | None -> ()
            let dialogAttrs =
                [ cls "dlg-card"; attr "role" "dialog"; attr "aria-modal" "true"; attr "aria-labelledby" titleId
                  attr "data-testid" spec.TestId; attr "tabindex" "-1"; onKeyDown keydown
                  on "keyup" (fun e -> e.stopPropagation ()); on "keypress" (fun e -> e.stopPropagation ())
                  refEffect (fun card ->
                      if previousFocus.IsNone then
                          let active = document.activeElement
                          if not (isNull active) then previousFocus <- Some active
                      let focusInside () =
                          let controls = focusable card
                          if controls.Length > 0 then controls.[0].focus () else (card :?> HTMLElement).focus ()
                      let onFocusIn (event: Event) =
                          if isCurrent state opened && isConnected card && not (card.contains(event.target :?> Node)) then
                              focusInside ()
                      document.addEventListener("focusin", onFocusIn)
                      let timer = defer (fun () ->
                          if isConnected card && isCurrent state opened then
                              focusInside ())
                      fun () ->
                          document.removeEventListener("focusin", onFocusIn)
                          cancelDeferred timer
                          let replacing = match state.Open.Peek() with Some current -> not (obj.ReferenceEquals(current, opened)) | None -> false
                          if not replacing then
                              previousFocus |> Option.iter (fun previous -> if isConnected previous then (previous :?> HTMLElement).focus ())
                              previousFocus <- None) ]
                @ (if spec.Subtitle.IsSome then [ attr "aria-describedby" subtitleId ] else [])
            [ div [ cls "office-dialog"; cls "dlg-overlay"; attr "data-testid" (spec.TestId + "-overlay") ] [
                div dialogAttrs [
                    div [ cls "dlg-header" ] [
                        h2 [ cls "dlg-title"; id titleId ] [ text spec.Title ]
                        match spec.Subtitle with
                        | Some subtitle -> p [ cls "dlg-subtitle"; id subtitleId ] [ text subtitle ]
                        | None -> fragment []
                    ]
                    div [ cls "dlg-body" ] tabViews
                    div [ cls "dlg-buttons" ] (spec.Buttons |> List.map (fun btn ->
                        button ([ cls "dlg-button"; typ "button"; attr "data-testid" btn.TestId; onClick (fun _ -> activate state opened btn) ]
                                @ (if btn.IsPrimary then [ cls "dlg-button-primary" ] else [])) [ text btn.Label ]))
                ]
            ] ])


#endif
