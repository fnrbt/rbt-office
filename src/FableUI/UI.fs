module Rbt.FableUI.UI

open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Browser.Types
open Rbt.FableUI.Reactive
open Rbt.FableUI.Dom

[<Emit("$0.querySelectorAll('button, [href], input, select, textarea, [tabindex]:not([tabindex=\"-1\"])')")>]
let private queryFocusable (element: Element) : NodeList = jsNative

[<Emit("$0.focus()")>]
let private focusElement (element: Element) : unit = jsNative

let private focusFirst (container: Element) =
    let nodes = queryFocusable container
    if nodes.length > 0 then focusElement (nodes.[0] :?> Element)

let private focusTrap (container: Element) : unit -> unit =
    let previouslyFocused = document.activeElement
    let handleKeyDown (event: Event) =
        let keyboardEvent = event :?> KeyboardEvent
        if keyboardEvent.key = "Tab" then
            let nodes = queryFocusable container
            if nodes.length > 0 then
                let first = nodes.[0] :?> Element
                let last = nodes.[int nodes.length - 1] :?> Element
                if keyboardEvent.shiftKey && document.activeElement = first then
                    keyboardEvent.preventDefault()
                    focusElement last
                elif not keyboardEvent.shiftKey && document.activeElement = last then
                    keyboardEvent.preventDefault()
                    focusElement first

    JS.setTimeout (fun () -> focusFirst container) 0 |> ignore
    container.addEventListener("keydown", handleKeyDown)

    fun () ->
        container.removeEventListener("keydown", handleKeyDown)
        if not (isNull previouslyFocused) then focusElement previouslyFocused

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

let DialogWithTestIds overlayTestId dialogTestId (isOpen: Signal<bool>) (content: View list) =
    showIf (fun () -> isOpen.Value)
        (viewComponent (fun () ->
            let portal = document.createElement("div")
            portal.setAttribute("data-fui-dialog-portal", "")
            document.body.appendChild(portal) |> ignore
            let close () = isOpen.Value <- false
            let mutable contentCleanup: (unit -> unit) option = None

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
                              contentCleanup <- Some (fun () -> cleanupDismiss(); cleanupFocus())
                              fun () ->
                                  cleanupDismiss()
                                  cleanupFocus()) ]
                        content
                ]

            let portalCleanup = mount portal dialog

            div [ style "display" "none"
                  onUnmount (fun () ->
                      contentCleanup |> Option.iter (fun cleanup -> cleanup())
                      portalCleanup ()
                      if not (isNull portal.parentNode) then portal.parentNode.removeChild(portal) |> ignore) ] []))

let Dialog isOpen content = DialogWithTestIds "fui-dialog-overlay" "fui-dialog" isOpen content

let DialogHeader attrs content = div (cls "dialog-header" :: attrs) content
let DialogTitle attrs content = h2 (cls "dialog-title" :: attrs) content
let DialogBody attrs content = div (cls "dialog-body" :: attrs) content
let DialogFooter attrs content = div (cls "dialog-footer" :: attrs) content
