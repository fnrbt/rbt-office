module Rbt.FableUI.UI

open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Browser.Types
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom

[<Emit("Array.from($0.querySelectorAll('button:not([disabled]), [href], input:not([disabled]):not([type=hidden]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex=\"-1\"])')).filter(e => e.getClientRects().length && e.tabIndex >= 0)")>]
let private queryFocusable (element: Element) : HTMLElement array = jsNative

[<Emit("$0.focus()")>]
let private focusElement (element: Element) : unit = jsNative

let private focusFirst (container: Element) =
    let nodes = queryFocusable container
    if nodes.Length > 0 then focusElement nodes.[0]

let private focusTrap (container: Element) : unit -> unit =
    let previouslyFocused = document.activeElement
    let handleKeyDown (event: Event) =
        let keyboardEvent = event :?> KeyboardEvent
        if keyboardEvent.key = "Tab" then
            let nodes = queryFocusable container
            if nodes.Length > 0 then
                let first = nodes.[0] :> Element
                let last = nodes.[nodes.Length - 1] :> Element
                if keyboardEvent.shiftKey && document.activeElement = first then
                    keyboardEvent.preventDefault()
                    focusElement last
                elif not keyboardEvent.shiftKey && document.activeElement = last then
                    keyboardEvent.preventDefault()
                    focusElement first

    let timer = JS.setTimeout (fun () -> if container?isConnected then focusFirst container) 0
    container.addEventListener("keydown", handleKeyDown)

    fun () ->
        window.clearTimeout(timer)
        container.removeEventListener("keydown", handleKeyDown)
        if not (isNull previouslyFocused) && previouslyFocused?isConnected then focusElement previouslyFocused

let private dismissible (container: Element) (onDismiss: unit -> unit) : unit -> unit =
    let handleMouseDown (event: Event) =
        let target = event.target :?> Node
        if not (container.contains(target)) then onDismiss ()

    let handleKeyDown (event: Event) =
        let keyboardEvent = event :?> KeyboardEvent
        if keyboardEvent.key = "Escape" then onDismiss ()

    let mouseTimer = JS.setTimeout (fun () -> document.addEventListener("mousedown", handleMouseDown)) 0
    document.addEventListener("keydown", handleKeyDown)

    fun () ->
        window.clearTimeout(mouseTimer)
        document.removeEventListener("mousedown", handleMouseDown)
        document.removeEventListener("keydown", handleKeyDown)

let Button attrs content = button (cls "btn" :: attrs) content
let Input attrs = input (cls "input" :: attrs)

let mutable private nextPortalId = 0

let rec private freshPortalId () =
    nextPortalId <- nextPortalId + 1
    let candidate = sprintf "rbt-office-dialog-%d" nextPortalId
    if isNull (document.getElementById candidate) then candidate else freshPortalId ()

let DialogWithTestIds overlayTestId dialogTestId (isOpen: Signal<bool>) (content: View list) =
    showIf (fun () -> isOpen.Value)
        (viewComponent (fun () ->
            Chrome.ensureStyles ()
            let portal = document.createElement("div")
            portal.id <- freshPortalId ()
            portal.className <- "rb-office-dialog"
            portal.setAttribute("data-fui-dialog-portal", "")
            document.body.appendChild(portal) |> ignore
            let close () = isOpen.Value <- false

            let dialog =
                div [ cls "dialog-overlay"; attr "data-testid" overlayTestId; onClick (fun _ -> close()) ] [
                    div [ cls "dialog-backdrop" ] []
                    div [ cls "dialog-content"
                          attr "role" "dialog"
                          attr "aria-modal" "true"
                          attr "data-testid" dialogTestId
                          onClick (fun event -> event.stopPropagation())
                          refEffect (fun element ->
                              let cleanupFocus = focusTrap element
                              let cleanupDismiss = dismissible element close
                              fun () ->
                                  cleanupDismiss()
                                  cleanupFocus()) ]
                        content
                ]

            let portalCleanup = mount portal dialog
            let title = portal.querySelector(".rb-dialog-title")
            let card = portal.querySelector("[role=dialog]")
            if not (isNull title) && not (isNull card) then
                title.id <- portal.id + "-title"
                card.setAttribute("aria-labelledby", title.id)

            div [ style "display" "none"
                  onUnmount (fun () ->
                      portalCleanup ()
                      if not (isNull portal.parentNode) then portal.parentNode.removeChild(portal) |> ignore) ] []))

let Dialog isOpen content = DialogWithTestIds "fui-dialog-overlay" "fui-dialog" isOpen content

let DialogHeader attrs content = div (cls "dialog-header" :: attrs) content
let DialogTitle attrs content = h2 (cls "dialog-title" :: attrs) content
let DialogBody attrs content = div (cls "dialog-body" :: attrs) content
let DialogFooter attrs content = div (cls "dialog-footer" :: attrs) content
