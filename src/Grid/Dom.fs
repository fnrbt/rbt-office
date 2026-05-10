/// DOM manipulation helpers using Emit
module RedBat.Grid.Dom

open Fable.Core
open Fable.Core.JsInterop

/// HTML Element (opaque)
type Element = obj
type EventListener = obj -> unit

// Document access
[<Emit("document")>]
let getDocument (): Element = jsNative

[<Emit("document.body")>]
let getBody (): Element = jsNative

[<Emit("document.getElementById($0)")>]
let getElementById (id: string): Element = jsNative

[<Emit("$0.querySelectorAll($1)")>]
let querySelectorAll (el: Element) (selector: string): obj = jsNative

[<Emit("$0.contains($1)")>]
let contains (el: Element) (other: Element): bool = jsNative

// Element creation
[<Emit("document.createElement($0)")>]
let createElement (tag: string): Element = jsNative

[<Emit("document.createTextNode($0)")>]
let createTextNode (text: string): Element = jsNative

// Properties
[<Emit("$0.className = $1")>]
let setClassName (el: Element) (className: string): unit = jsNative

[<Emit("$0.textContent = $1")>]
let setTextContent (el: Element) (text: string): unit = jsNative

[<Emit("$0.innerHTML = $1")>]
let setInnerHTML (el: Element) (html: string): unit = jsNative

[<Emit("$0.title = $1")>]
let setTitle (el: Element) (title: string): unit = jsNative

[<Emit("$0.type = $1")>]
let setType (el: Element) (typ: string): unit = jsNative

[<Emit("$0.value = $1")>]
let setValue (el: Element) (value: string): unit = jsNative

[<Emit("$0.value")>]
let getValue (el: Element): string = jsNative

[<Emit("$0.checked = $1")>]
let setChecked (el: Element) (checked: bool): unit = jsNative

[<Emit("$0.checked")>]
let getChecked (el: Element): bool = jsNative

// Attributes
[<Emit("$0.setAttribute($1, $2)")>]
let setAttribute (el: Element) (name: string) (value: string): unit = jsNative

[<Emit("$0.getAttribute($1)")>]
let getAttribute (el: Element) (name: string): string = jsNative

// Style
[<Emit("$0.style[$1] = $2")>]
let setStyle (el: Element) (prop: string) (value: string): unit = jsNative

[<Emit("$0.style.cssText = $1")>]
let setCssText (el: Element) (css: string): unit = jsNative

// DOM manipulation
[<Emit("$0.appendChild($1)")>]
let appendChild (parent: Element) (child: Element): unit = jsNative

[<Emit("$0.removeChild($1)")>]
let removeChild (parent: Element) (child: Element): unit = jsNative

[<Emit("$0.insertBefore($1, $2)")>]
let insertBefore (parent: Element) (newChild: Element) (refChild: Element): unit = jsNative

[<Emit("$0.children")>]
let getChildren (el: Element): Element array = jsNative

[<Emit("$0.children.length")>]
let getChildCount (el: Element): int = jsNative

[<Emit("$0.firstChild")>]
let getFirstChild (el: Element): Element option = jsNative

[<Emit("$0.querySelector($1)")>]
let querySelector (el: Element) (selector: string): Element option = jsNative

// Events
[<Emit("$0.addEventListener($1, $2)")>]
let addEventListener (el: Element) (eventName: string) (listener: EventListener): unit = jsNative

[<Emit("$0.removeEventListener($1, $2)")>]
let removeEventListener (el: Element) (eventName: string) (listener: EventListener): unit = jsNative

[<Emit("$0[$1]")>]
let getProp (obj: obj) (key: string): obj = jsNative

[<Emit("$0[$1] = $2")>]
let setProp (obj: obj) (key: string) (value: obj): unit = jsNative

// Event properties
[<Emit("$0.target")>]
let getEventTarget (e: obj): Element = jsNative

[<Emit("$0.currentTarget")>]
let getCurrentTarget (e: obj): Element = jsNative

[<Emit("$0.key")>]
let getEventKey (e: obj): string = jsNative

[<Emit("$0.preventDefault()")>]
let preventDefault (e: obj): unit = jsNative

[<Emit("$0.stopPropagation()")>]
let stopPropagation (e: obj): unit = jsNative

// Scrolling
[<Emit("$0.scrollTop")>]
let getScrollTop (el: Element): float = jsNative

[<Emit("$0.scrollTop = $1")>]
let setScrollTop (el: Element) (value: float): unit = jsNative

[<Emit("$0.clientHeight")>]
let getClientHeight (el: Element): float = jsNative

// Focus
[<Emit("$0.focus()")>]
let focus (el: Element): unit = jsNative

[<Emit("$0.select()")>]
let select (el: Element): unit = jsNative

[<Emit("$0.blur()")>]
let blur (el: Element): unit = jsNative

// Null checks
[<Emit("$0 == null")>]
let isNull (obj: obj): bool = jsNative

[<Emit("$0 != null")>]
let isNotNull (obj: obj): bool = jsNative

// Helper to set multiple styles
let setStyles (el: Element) (styles: (string * string) list): unit =
    for (prop, value) in styles do
        setStyle el prop value

// Request animation frame
[<Emit("requestAnimationFrame($0)")>]
let requestAnimationFrame (callback: unit -> unit): unit = jsNative
