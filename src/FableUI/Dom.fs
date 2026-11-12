module Rbt.FableUI.Dom

open Fable.Core
open Fable.Core.JsInterop
open Browser.Dom
open Browser.Types
open Rbt.FableUI.Reactive

type Attr =
    | StaticAttr of string * string
    | ReactiveAttr of string * (unit -> string)
    | StaticProp of string * obj
    | ReactiveProp of string * (unit -> obj)
    | EventHandler of string * (Event -> unit)
    | StaticStyle of string * string
    | ReactiveStyle of string * (unit -> string)
    | StaticClass of string
    | ReactiveClass of (unit -> string)
    | Ref of (Element -> unit)
    | RefEffect of (Element -> (unit -> unit))
    | OnUnmount of (unit -> unit)

type View =
    | Element of tag: string * attrs: Attr list * children: View list
    | Text of string
    | ReactiveText of (unit -> string)
    | Fragment of View list
    | ReactiveFragment of (unit -> View list)
    | Component of (unit -> View)
    | Empty

let private setStyle (element: Element) (prop: string) (value: string) =
    let style: obj = element?style
    if prop.StartsWith("--") then style?setProperty(prop, value)
    else style?(prop) <- value

[<Fable.Core.Emit("if ($0[$1] !== $2) $0[$1] = $2")>]
let private setChangedProperty (_: Element) (_: string) (_: obj) : unit = jsNative

let rec private mountWith (parent: Node) (view: View) : unit -> unit =
    let cleanups = ResizeArray<unit -> unit>()
    let addCleanup cleanup = cleanups.Add(cleanup)
    let runCleanups () =
        for cleanup in cleanups do cleanup()
        cleanups.Clear()

    match view with
    | Empty -> ignore
    | Text content ->
        let node = document.createTextNode(content)
        parent.appendChild(node) |> ignore
        fun () ->
            if not (isNull node.parentNode) then node.parentNode.removeChild(node) |> ignore
    | ReactiveText getText ->
        let node = document.createTextNode("")
        parent.appendChild(node) |> ignore
        let disposable = effect (fun () -> node.textContent <- getText())
        fun () ->
            disposable.Dispose()
            if not (isNull node.parentNode) then node.parentNode.removeChild(node) |> ignore
    | Element (tag, attrs, children) ->
        let element = document.createElement(tag)

        List.iter (fun attr ->
            match attr with
            | StaticAttr (name, value) -> element.setAttribute(name, value)
            | ReactiveAttr (name, getValue) ->
                let disposable = effect (fun () -> element.setAttribute(name, getValue()))
                addCleanup disposable.Dispose
            | StaticProp (name, value) -> element?(name) <- value
            | ReactiveProp (name, getValue) ->
                // Rewriting an input's own value destroys transient number input
                // such as "-" and resets text selection during two-way binding.
                let disposable = effect (fun () ->
                    let value = getValue()
                    setChangedProperty element name value)
                addCleanup disposable.Dispose
            | EventHandler (eventName, handler) ->
                element.addEventListener(eventName, handler)
                addCleanup (fun () -> element.removeEventListener(eventName, handler))
            | StaticStyle (prop, value) -> setStyle element prop value
            | ReactiveStyle (prop, getValue) ->
                let disposable = effect (fun () -> setStyle element prop (getValue()))
                addCleanup disposable.Dispose
            | StaticClass className -> element.classList.add(className)
            | ReactiveClass getClass ->
                let mutable previous: string array = [||]
                let disposable = effect (fun () ->
                    for className in previous do
                        if className <> "" then element.classList.remove(className)
                    let next = getClass().Split(' ') |> Array.filter ((<>) "")
                    for className in next do element.classList.add(className)
                    previous <- next)
                addCleanup disposable.Dispose
            | Ref callback -> callback element
            | RefEffect callback ->
                let cleanup = callback element
                addCleanup cleanup
            | OnUnmount cleanup -> addCleanup cleanup
        ) attrs

        // Keep returned disposers in separate callback scopes: Fable hoists
        // function-result adapters out of a for-loop and aliases their captures.
        List.iter (fun child -> addCleanup (mountWith element child)) children

        parent.appendChild(element) |> ignore

        fun () ->
            runCleanups ()
            if not (isNull element.parentNode) then element.parentNode.removeChild(element) |> ignore
    | Fragment children ->
        let childCleanups = children |> List.map (mountWith parent)
        fun () -> for cleanup in childCleanups do cleanup()
    | ReactiveFragment getChildren ->
        let marker = document.createComment("fui")
        parent.appendChild(marker) |> ignore
        let mutable currentCleanups: (unit -> unit) list = []
        let disposable = effect (fun () ->
            for cleanup in currentCleanups do cleanup()
            let fragment = document.createDocumentFragment()
            currentCleanups <- getChildren() |> List.map (mountWith fragment)
            parent.insertBefore(fragment, marker) |> ignore)
        fun () ->
            disposable.Dispose()
            for cleanup in currentCleanups do cleanup()
            if not (isNull marker.parentNode) then marker.parentNode.removeChild(marker) |> ignore
    | Component render -> mountWith parent (render ())

let mount (parent: Node) (view: View) : unit -> unit = mountWith parent view

let el tag attrs children = Element(tag, attrs, children)
let div attrs children = el "div" attrs children
let span attrs children = el "span" attrs children
let p attrs children = el "p" attrs children
let h2 attrs children = el "h2" attrs children
let button attrs children = el "button" attrs children
let canvas attrs children = el "canvas" attrs children
let input attrs = el "input" attrs []
let label attrs children = el "label" attrs children
let select attrs children = el "select" attrs children
let option attrs children = el "option" attrs children

let text value = Text value
let textR getText = ReactiveText getText
let fragment children = Fragment children
let children getChildren = ReactiveFragment getChildren
let showIf condition view = ReactiveFragment(fun () -> if condition() then [view] else [])
let viewComponent render = Component render

let attr name value = StaticAttr(name, value)
let attrR name getValue = ReactiveAttr(name, getValue)
let prop name value = StaticProp(name, value)
let propR name getValue = ReactiveProp(name, getValue)
let propS name (signal: Signal<'T>) = ReactiveProp(name, fun () -> box signal.Value)
let style name value = StaticStyle(name, value)
let styleR name getValue = ReactiveStyle(name, getValue)
let cls className = StaticClass("rb-" + className)
let rawCls className = StaticClass className
let clsR getClass = ReactiveClass getClass
let refEl callback = Ref callback
let refEffect callback = RefEffect callback
let on eventName handler = EventHandler(eventName, handler)
let onClick handler = on "click" (fun event -> handler (event :?> MouseEvent))
let onInput handler = on "input" (fun event -> handler ((event.target :?> HTMLInputElement).value))
let onChange handler = on "change" (fun event -> handler event)
let onKeyDown handler = on "keydown" (fun event -> handler (event :?> KeyboardEvent))
let onUnmount cleanup = OnUnmount cleanup

let id value = attr "id" value
let typ value = attr "type" value
let valueS signal = propS "value" signal
let disabled value = prop "disabled" value

let mountApp selector view =
    let container = document.querySelector(selector)
    if isNull container then failwith (sprintf "Could not find element: %s" selector)
    mount container view
