namespace Rbt.Word

#if FABLE_COMPILER

/// The Word shell's navigation pane (the headings as a tree beside the pages) and outline view (the
/// document as Word's outline, over the pages, with the outlining toolbar). Outline.fs is the model
/// and the commands; both go through the editor's API (one undo step each). Both are drawn with the
/// DOM directly and only when the outline changed: typing in a long document costs the pane a look
/// at the block list, and the outline is computed again at most every few hundred milliseconds.
module OutlineUi =
    open System.Collections.Generic
    open Fable.Core
    open Fable.Core.JsInterop
    open Browser.Types
    open Browser.Dom
    open Rbt.FableUI.Reactive
    open Rbt.FableUI.Dom

    module WordApi = Rbt.Word.API
    module WordHost = Rbt.Word.Host

    let css = """
.rb-word-shell .rb-word-main { display: flex; flex: 1 1 70vh; min-height: 240px; min-width: 0; }
.rb-word-shell .rb-word-pages { position: relative; display: flex; flex-direction: column; flex: 1 1 auto; min-width: 0; min-height: 0; }
.rb-word-shell .rb-word-nav-pane { display: flex; flex-direction: column; flex: 0 0 240px; max-width: 40%; min-height: 0; box-sizing: border-box; border-right: 1px solid var(--office-border); background: var(--office-surface); }
.rb-word-shell .rb-word-nav-header { display: flex; justify-content: space-between; align-items: center; gap: 8px; padding: 6px 8px 6px 12px; border-bottom: 1px solid var(--office-border); font-weight: 600; }
.rb-word-shell .rb-word-nav-tree { flex: 1 1 auto; min-height: 0; overflow: auto; padding: 4px 0; outline: none; }
.rb-word-shell .rb-word-nav-item { display: flex; align-items: baseline; gap: 4px; padding: 3px 8px; cursor: pointer; white-space: nowrap; user-select: none; }
.rb-word-shell .rb-word-nav-item:hover { background: var(--office-desk); }
.rb-word-shell .rb-word-nav-item:focus-visible { outline: 2px solid var(--office-accent); outline-offset: -2px; }
.rb-word-shell .rb-word-nav-item.rb-word-nav-current { background: var(--office-selected); color: var(--office-accent); box-shadow: inset 3px 0 0 var(--office-accent); }
.rb-word-shell .rb-word-nav-item.rb-word-nav-dragging { opacity: 0.5; }
.rb-word-shell .rb-word-nav-item.rb-word-nav-drop-before { box-shadow: inset 0 2px 0 var(--office-accent); }
.rb-word-shell .rb-word-nav-item.rb-word-nav-drop-after { box-shadow: inset 0 -2px 0 var(--office-accent); }
.rb-word-shell .rb-word-nav-toggle { flex: 0 0 14px; text-align: center; color: var(--office-muted); }
.rb-word-shell .rb-word-nav-label { overflow: hidden; text-overflow: ellipsis; }
.rb-word-shell .rb-word-nav-empty { padding: 10px 12px; color: var(--office-muted); white-space: normal; }
.rb-word-shell .rb-word-outline-view { position: absolute; inset: 0; z-index: 5; display: flex; flex-direction: column; background: var(--office-surface); }
.rb-word-shell .rb-word-outline-toolbar { display: flex; flex-wrap: wrap; align-items: center; gap: 3px; padding: 6px 10px; border-bottom: 1px solid var(--office-border); }
.rb-word-shell .rb-word-outline-toolbar label { display: inline-flex; align-items: center; gap: 4px; margin-inline: 4px; }
.rb-word-shell .rb-word-outline-lines { flex: 1 1 auto; min-height: 0; overflow: auto; padding: 12px 16px 48px; outline: none; }
.rb-word-shell .rb-word-outline-line { position: relative; padding: 2px 6px; min-height: 1.3em; white-space: pre-wrap; overflow-wrap: anywhere; cursor: default; user-select: none; }
.rb-word-shell .rb-word-outline-line[aria-selected="true"] { background: var(--office-selected); }
.rb-word-shell .rb-word-outline-lines:focus-visible .rb-word-outline-focus { outline: 1px dotted var(--office-accent); }
.rb-word-shell .rb-word-outline-symbol { position: absolute; top: 2px; width: 18px; text-align: center; color: var(--office-muted); cursor: pointer; }
.rb-word-shell .rb-word-outline-collapsed .rb-word-outline-text { border-bottom: 1px dashed var(--office-muted); }
.rb-word-shell .rb-word-outline-table .rb-word-outline-text { font-style: italic; color: var(--office-muted); }
.rb-word-shell .rb-word-outline-first-line .rb-word-outline-line { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
@media (max-width: 600px) {
  .rb-word-shell .rb-word-main { flex-direction: column; }
  .rb-word-shell .rb-word-nav-pane { flex: 0 0 auto; max-width: none; max-height: 30vh; border-right: 0; border-bottom: 1px solid var(--office-border); }
}
"""

    [<Emit("window.setTimeout($0, $1)")>]
    let private setTimer (_: unit -> unit) (_: int) : float = jsNative

    [<Emit("window.clearTimeout($0)")>]
    let private clearTimer (_: float) : unit = jsNative

    [<Emit("$0 && $0.focus({ preventScroll: $1 })")>]
    let private focusElement (_: obj) (_: bool) : unit = jsNative

    [<Emit("$0 && $0.closest ? $0.closest($1) : null")>]
    let private closest (_: EventTarget) (_: string) : HTMLElement = jsNative

    [<Emit("$0.scrollIntoView({ block: 'nearest' })")>]
    let private scrollNearest (_: HTMLElement) : unit = jsNative

    [<Emit("$0.dataTransfer")>]
    let private dataTransfer (_: Event) : obj = jsNative

    let private element (tag: string) (className: string) : HTMLElement =
        let e = document.createElement tag
        if className <> "" then e.className <- className
        e

    /// A popup menu of actions at a point (the navigation pane's context menu): the shell's menu look,
    /// keyboard Up / Down / Home / End / Escape, closed by a click elsewhere.
    type private Popup(owner: HTMLElement) =
        let mutable menu: HTMLElement = null
        let mutable returnFocus: obj = null
        let onOutside (event: Event) =
            if not (isNull menu) && not (menu.contains(event.target :?> Node)) then
                menu.remove()
                menu <- null
        member this.Close(restore: bool) =
            if not (isNull menu) then
                menu.remove()
                menu <- null
                document.removeEventListener("pointerdown", onOutside, true)
                if restore then focusElement returnFocus true
        member this.Open(x: float, y: float, testId: string, items: (string * string * bool * (unit -> unit)) list, focusBack: obj) =
            this.Close false
            returnFocus <- focusBack
            let m = element "div" "rb-word-menu rb-word-context-menu rb-word-nav-menu"
            m.setAttribute("role", "menu")
            m.setAttribute("data-testid", testId)
            m.setAttribute("aria-label", "Heading actions")
            for itemTestId, label, enabled, action in items do
                let b = element "button" "rb-word-menu-item" :?> HTMLButtonElement
                b.textContent <- label
                b.setAttribute("role", "menuitem")
                b.setAttribute("data-testid", itemTestId)
                b.disabled <- not enabled
                b.addEventListener("click", fun _ -> this.Close true; action ())
                m.appendChild b |> ignore
            m.addEventListener("keydown", fun (event: Event) ->
                let key = event :?> KeyboardEvent
                let buttons = m.querySelectorAll("button:not(:disabled)")
                if key.key = "Escape" || key.key = "Tab" then
                    this.Close true
                    key.preventDefault(); key.stopPropagation()
                elif buttons.length > 0 && (key.key = "ArrowDown" || key.key = "ArrowUp" || key.key = "Home" || key.key = "End") then
                    let mutable index = -1
                    for i = 0 to buttons.length - 1 do
                        if obj.ReferenceEquals(buttons.[i], document.activeElement) then index <- i
                    let next =
                        match key.key with
                        | "Home" -> 0
                        | "End" -> buttons.length - 1
                        | "ArrowUp" -> (index + buttons.length - 1) % buttons.length
                        | _ -> (index + 1) % buttons.length
                    focusElement buttons.[next] false
                    key.preventDefault(); key.stopPropagation())
            owner.appendChild m |> ignore
            let bounds = m.getBoundingClientRect()
            let left = max 8.0 (min x (window.innerWidth - bounds.width - 8.0))
            let top = max 8.0 (min y (window.innerHeight - bounds.height - 8.0))
            m.setAttribute("style", "position:fixed;z-index:40;left:" + string (int left) + "px;top:" + string (int top) + "px")
            menu <- m
            document.addEventListener("pointerdown", onOutside, true)
            focusElement (m.querySelector "button:not(:disabled)") false

    // ---- the navigation pane ---------------------------------------------------------------------------

    /// The headings of the document as a tree beside the pages (View > Navigation pane): indented by
    /// their place in the outline, each with its own expand / collapse; the heading the caret is under
    /// highlighted; a click goes to a heading; a context menu (right click, Shift+F10, the menu key)
    /// promotes, demotes, moves (with the subtext) and deletes; a heading dragged onto another moves
    /// there with its subtext.
    let navigationPane (host: WordHost.Host) (isOpen: Signal<bool>) : View =
        let collapsed = HashSet<string>()
        let items = Dictionary<string, HTMLElement>()
        let pane: HTMLElement ref = ref null
        let tree: HTMLElement ref = ref null
        let popup: Popup option ref = ref None
        let outline: Outline.DocumentOutline option ref = ref None
        let shownBlocks: obj ref = ref null
        let shownStyles: obj ref = ref null
        // (Per heading of the outline: the heading it is under, its depth in the tree, whether it has subheadings.)
        let parents: int array ref = ref [||]
        let hasChildren: bool array ref = ref [||]
        let signature = ref ""
        let current: string ref = ref null
        let timer = ref 0.0
        let dragged: string ref = ref null

        let positionOf (o: Outline.DocumentOutline) (id: string) =
            let mutable found = -1
            for i = 0 to o.Headings.Length - 1 do
                if found < 0 && o.BlockIds.[o.Headings.[i]] = id then found <- i
            found

        let highlight () =
            match host.Editor, outline.Value with
            | Some editor, Some o when obj.ReferenceEquals(o.Blocks, editor.Whole.Blocks) ->
                let id = WordApi.getCurrentHeading editor
                // (A heading inside a collapsed one: the one shown that holds it.)
                let mutable shown = id
                if not (isNull id) && not (items.ContainsKey id) then
                    let mutable at = positionOf o id
                    while at >= 0 && not (items.ContainsKey o.BlockIds.[o.Headings.[at]]) do at <- parents.Value.[at]
                    shown <- if at >= 0 then o.BlockIds.[o.Headings.[at]] else null
                if shown <> current.Value then
                    if not (isNull current.Value) then
                        match Portable.Dict.tryFind current.Value items with
                        | ValueSome item -> item.classList.remove "rb-word-nav-current"; item.setAttribute("aria-selected", "false")
                        | ValueNone -> ()
                    current.Value <- shown
                    if not (isNull shown) then
                        match Portable.Dict.tryFind shown items with
                        | ValueSome item ->
                            item.classList.add "rb-word-nav-current"
                            item.setAttribute("aria-selected", "true")
                            scrollNearest item
                        | ValueNone -> ()
            | _ -> ()

        let build (force: bool) =
            if not (isNull tree.Value) then
                match host.Editor with
                | None ->
                    tree.Value.textContent <- ""
                    items.Clear()
                    outline.Value <- None
                | Some editor ->
                    let doc = editor.Whole
                    let o = Outline.ofDocument doc
                    outline.Value <- Some o
                    shownBlocks.Value <- box doc.Blocks
                    shownStyles.Value <- box doc.Styles
                    let n = o.Headings.Length
                    // (Drawn again only when the headings are other headings: their ids, levels and text.)
                    let parts = ResizeArray<string>(n * 3)
                    for i = 0 to n - 1 do
                        let index = o.Headings.[i]
                        parts.Add o.BlockIds.[index]
                        parts.Add (string o.Levels.[index])
                        parts.Add o.HeadingTexts.[i]
                    let next = System.String.Join("\u0001", parts)
                    if force || next <> signature.Value then
                        signature.Value <- next
                        parents.Value <- Array.create n -1
                        hasChildren.Value <- Array.create n false
                        let depths = Array.zeroCreate n
                        let stack = ResizeArray<int>()
                        for i = 0 to n - 1 do
                            let level = o.Levels.[o.Headings.[i]]
                            while stack.Count > 0 && o.Levels.[o.Headings.[stack.[stack.Count - 1]]] >= level do Portable.Stack.dropLast stack
                            if stack.Count > 0 then
                                parents.Value.[i] <- stack.[stack.Count - 1]
                                hasChildren.Value.[stack.[stack.Count - 1]] <- true
                            depths.[i] <- stack.Count
                            stack.Add i
                        items.Clear()
                        current.Value <- null
                        tree.Value.textContent <- ""
                        if n = 0 then
                            let empty = element "div" "rb-word-nav-empty"
                            empty.setAttribute("data-testid", "word-nav-empty")
                            empty.textContent <- "No headings. Text in a heading style (Heading 1, Heading 2, ...) or with an outline level is listed here."
                            tree.Value.appendChild empty |> ignore
                        else
                            let fragment = document.createDocumentFragment()
                            let mutable hiddenBelow = -1
                            for i = 0 to n - 1 do
                                if hiddenBelow >= 0 && depths.[i] > hiddenBelow then ()
                                else
                                    hiddenBelow <- -1
                                    let index = o.Headings.[i]
                                    let id = o.BlockIds.[index]
                                    let item = element "div" "rb-word-nav-item"
                                    item.setAttribute("role", "treeitem")
                                    item.setAttribute("data-testid", "word-nav-item")
                                    item.setAttribute("data-block-id", id)
                                    item.setAttribute("data-level", string o.Levels.[index])
                                    item.setAttribute("aria-level", string (depths.[i] + 1))
                                    item.setAttribute("aria-selected", "false")
                                    item.setAttribute("draggable", "true")
                                    item.setAttribute("tabindex", if i = 0 then "0" else "-1")
                                    item.setAttribute("style", "padding-left:" + string (8 + depths.[i] * 14) + "px")
                                    let text = o.HeadingTexts.[i]
                                    item.setAttribute("title", text)
                                    let toggle = element "span" "rb-word-nav-toggle"
                                    toggle.setAttribute("data-testid", "word-nav-toggle")
                                    toggle.setAttribute("aria-hidden", "true")
                                    if hasChildren.Value.[i] then
                                        let isCollapsed = collapsed.Contains id
                                        item.setAttribute("aria-expanded", if isCollapsed then "false" else "true")
                                        toggle.textContent <- if isCollapsed then "▸" else "▾"
                                        if isCollapsed then hiddenBelow <- depths.[i]
                                    let label = element "span" "rb-word-nav-label"
                                    label.textContent <- if text = "" then "(empty heading)" else text
                                    item.appendChild toggle |> ignore
                                    item.appendChild label |> ignore
                                    fragment.appendChild item |> ignore
                                    items.[id] <- item
                            tree.Value.appendChild fragment |> ignore
                    highlight ()

        let refresh () =
            if isOpen.Value then
                match host.Editor with
                | Some editor when not (isNull tree.Value) ->
                    let doc = editor.Whole
                    if obj.ReferenceEquals(box doc.Blocks, shownBlocks.Value) && obj.ReferenceEquals(box doc.Styles, shownStyles.Value) then highlight ()
                    elif isNull shownBlocks.Value then build false
                    // (While the document is typed in: once in a while, not at each key.)
                    elif timer.Value = 0.0 then
                        timer.Value <- setTimer (fun () -> timer.Value <- 0.0; if isOpen.Value then build false) 300
                | _ -> ()

        let run (status: string) (command: WordApi.Editor -> string -> bool) (id: string) =
            match host.Editor with
            | Some editor ->
                let mutable changed = false
                WordHost.runCommand host (fun e -> changed <- command e id)
                if changed then host.Status.Value <- status
                build false
                match Portable.Dict.tryFind id items with
                | ValueSome item -> focusElement item true
                | ValueNone -> ()
            | None -> ()

        let goTo (id: string) =
            match host.Editor with
            | Some editor ->
                if WordApi.goToHeading editor id then
                    host.Status.Value <- "Went to heading"
                    highlight ()
            | None -> ()

        let toggle (id: string) =
            if not (collapsed.Remove id) then collapsed.Add id |> ignore
            build true
            match Portable.Dict.tryFind id items with
            | ValueSome item -> focusElement item true
            | ValueNone -> ()

        let openMenu (id: string) (x: float) (y: float) (item: HTMLElement) =
            match host.Editor, outline.Value, popup.Value with
            | Some editor, Some o, Some menu ->
                let doc = editor.Whole
                let index = Outline.indexOf doc id
                let canUp = (Outline.siblingMove true index doc).IsSome
                let canDown = (Outline.siblingMove false index doc).IsSome
                let at = positionOf o id
                let branch = [ for i in index .. Outline.subtextEnd o index - 1 do if o.Levels.[i] < Outline.BodyLevel then yield o.Levels.[i] ]
                menu.Open(x, y, "word-nav-menu",
                          [ "word-nav-menu-promote", "Promote", (branch |> List.forall (fun l -> l > 1)), (fun () -> run "Heading promoted" WordApi.promoteHeading id)
                            "word-nav-menu-demote", "Demote", (branch |> List.forall (fun l -> l < 9)), (fun () -> run "Heading demoted" WordApi.demoteHeading id)
                            "word-nav-menu-move-up", "Move up", canUp, (fun () -> run "Heading moved up" WordApi.moveHeadingUp id)
                            "word-nav-menu-move-down", "Move down", canDown, (fun () -> run "Heading moved down" WordApi.moveHeadingDown id)
                            "word-nav-menu-delete", "Delete heading and content", true, (fun () -> run "Heading and its content deleted" WordApi.deleteHeading id)
                            "word-nav-menu-expand-all", "Expand all", collapsed.Count > 0, (fun () -> collapsed.Clear(); build true)
                            "word-nav-menu-collapse-all", "Collapse all", (at >= 0 && hasChildren.Value |> Array.exists (fun b -> b)),
                                (fun () -> (for i = 0 to hasChildren.Value.Length - 1 do if hasChildren.Value.[i] then collapsed.Add o.BlockIds.[o.Headings.[i]] |> ignore); build true) ],
                          item)
            | _ -> ()

        let visibleItems () = tree.Value.querySelectorAll ".rb-word-nav-item"

        let onTreeKey (event: Event) =
            let key = event :?> KeyboardEvent
            let item = closest event.target ".rb-word-nav-item"
            if not (isNull item) && not key.altKey && not key.ctrlKey && not key.metaKey then
                let id = item.getAttribute "data-block-id"
                let all = visibleItems ()
                let mutable index = -1
                for i = 0 to all.length - 1 do
                    if obj.ReferenceEquals(all.[i], item) then index <- i
                let focusAt (i: int) =
                    if i >= 0 && i < all.length then
                        for j = 0 to all.length - 1 do (all.[j] :?> HTMLElement).setAttribute("tabindex", if j = i then "0" else "-1")
                        focusElement all.[i] false
                let handled =
                    match key.key with
                    | "ArrowDown" -> focusAt (index + 1); true
                    | "ArrowUp" -> focusAt (index - 1); true
                    | "Home" -> focusAt 0; true
                    | "End" -> focusAt (all.length - 1); true
                    | "ArrowRight" ->
                        if item.getAttribute "aria-expanded" = "false" then toggle id else focusAt (index + 1)
                        true
                    | "ArrowLeft" ->
                        if item.getAttribute "aria-expanded" = "true" then toggle id
                        else
                            match outline.Value with
                            | Some o ->
                                let at = positionOf o id
                                if at >= 0 && parents.Value.[at] >= 0 then
                                    match Portable.Dict.tryFind o.BlockIds.[o.Headings.[parents.Value.[at]]] items with
                                    | ValueSome parent -> focusElement parent false
                                    | ValueNone -> ()
                            | None -> ()
                        true
                    | "Enter" | " " -> goTo id; true
                    | "ContextMenu" -> (let r = item.getBoundingClientRect() in openMenu id (r.left + 16.0) r.bottom item); true
                    | "F10" when key.shiftKey -> (let r = item.getBoundingClientRect() in openMenu id (r.left + 16.0) r.bottom item); true
                    | _ -> false
                if handled then
                    key.preventDefault()
                    key.stopPropagation()

        let clearDrop () =
            for item in items.Values do
                item.classList.remove "rb-word-nav-drop-before"
                item.classList.remove "rb-word-nav-drop-after"

        /// Where a heading dragged onto `target` goes: before it (upper half), or after it and its subtext.
        let dropPlace (event: Event) (target: HTMLElement) =
            match host.Editor, outline.Value with
            | Some editor, Some o when not (isNull dragged.Value) && obj.ReferenceEquals(o.Blocks, editor.Whole.Blocks) ->
                let doc = editor.Whole
                let first = Outline.indexOf doc dragged.Value
                let targetIndex = Outline.indexOf doc (target.getAttribute "data-block-id")
                if first < 0 || targetIndex < 0 then None
                else
                    let last = Outline.subtextEnd o first
                    let rect = target.getBoundingClientRect()
                    let mouse = event :?> MouseEvent
                    let above = mouse.clientY < rect.top + rect.height / 2.0
                    let before = if above then targetIndex else Outline.subtextEnd o targetIndex
                    if (targetIndex >= first && targetIndex < last) || (before >= first && before <= last) then None
                    else Some (o, first, last, before, above)
            | _ -> None

        let onDragStart (event: Event) =
            let item = closest event.target ".rb-word-nav-item"
            if not (isNull item) then
                dragged.Value <- item.getAttribute "data-block-id"
                item.classList.add "rb-word-nav-dragging"
                let transfer = dataTransfer event
                if not (isNull transfer) then
                    transfer?setData("text/plain", dragged.Value) |> ignore
                    transfer?effectAllowed <- "move"

        let onDragOver (event: Event) =
            let item = closest event.target ".rb-word-nav-item"
            clearDrop ()
            if not (isNull item) then
                match dropPlace event item with
                | Some (_, _, _, _, above) ->
                    event.preventDefault()
                    item.classList.add (if above then "rb-word-nav-drop-before" else "rb-word-nav-drop-after")
                    let transfer = dataTransfer event
                    if not (isNull transfer) then transfer?dropEffect <- "move"
                | None -> ()

        let onDragEnd (_: Event) =
            clearDrop ()
            for item in items.Values do item.classList.remove "rb-word-nav-dragging"
            dragged.Value <- null

        let onDrop (event: Event) =
            let item = closest event.target ".rb-word-nav-item"
            if not (isNull item) then
                match dropPlace event item with
                | Some (o, first, last, before, _) ->
                    event.preventDefault()
                    let id = dragged.Value
                    let beforeId = if before >= o.BlockIds.Length then null else o.BlockIds.[before]
                    let firstId, lastId = o.BlockIds.[first], o.BlockIds.[last - 1]
                    onDragEnd event
                    run "Heading moved" (fun editor _ -> WordApi.moveBlocks editor firstId lastId beforeId) id
                | None -> onDragEnd event

        div [ cls "word-nav-pane"; attr "role" "navigation"; attr "aria-label" "Navigation"; attr "data-testid" "word-nav-pane"
              styleR "display" (fun () -> if isOpen.Value then "flex" else "none")
              refEffect (fun element ->
                  pane.Value <- element :?> HTMLElement
                  popup.Value <- Some (Popup pane.Value)
                  let subscription = effect (fun () ->
                      let _ = host.SelectionVersion.Value
                      let _ = isOpen.Value
                      untrack refresh)
                  fun () ->
                      subscription.Dispose()
                      if timer.Value <> 0.0 then clearTimer timer.Value
                      popup.Value |> Option.iter (fun p -> p.Close false)
                      pane.Value <- null) ] [
            div [ cls "word-nav-header" ] [
                span [] [ text "Navigation" ]
                button [ attr "type" "button"; attr "data-testid" "word-nav-close"; attr "aria-label" "Close the navigation pane"
                         onClick (fun _ -> isOpen.Value <- false) ] [ text "Close" ]
            ]
            div [ cls "word-nav-tree"; attr "role" "tree"; attr "aria-label" "Headings"; attr "data-testid" "word-nav-tree"
                  refEffect (fun element ->
                      tree.Value <- element :?> HTMLElement
                      shownBlocks.Value <- null
                      tree.Value.addEventListener("click", fun (event: Event) ->
                          let item = closest event.target ".rb-word-nav-item"
                          if not (isNull item) then
                              let id = item.getAttribute "data-block-id"
                              if not (isNull (closest event.target ".rb-word-nav-toggle")) && item.hasAttribute "aria-expanded" then toggle id
                              else goTo id)
                      tree.Value.addEventListener("contextmenu", fun (event: Event) ->
                          let item = closest event.target ".rb-word-nav-item"
                          if not (isNull item) then
                              event.preventDefault()
                              let mouse = event :?> MouseEvent
                              openMenu (item.getAttribute "data-block-id") mouse.clientX mouse.clientY item)
                      tree.Value.addEventListener("keydown", onTreeKey)
                      tree.Value.addEventListener("dragstart", onDragStart)
                      tree.Value.addEventListener("dragover", onDragOver)
                      tree.Value.addEventListener("drop", onDrop)
                      tree.Value.addEventListener("dragend", onDragEnd)
                      refresh ()
                      fun () -> tree.Value <- null) ] []
        ]

    // ---- outline view ---------------------------------------------------------------------------------------

    /// What outline view's commands are run with (the shell's keys and menus): "promote", "demote",
    /// "promoteToHeading1", "demoteToBody", "level:N", "moveUp", "moveDown", "expand", "collapse".
    type OutlineControl() =
        member val Run : string -> unit = ignore with get, set

    /// Outline view (View > Outline): every block of the body a line, indented by its level, with
    /// Word's symbols - a plus for a heading with subtext, a minus for one without, a small circle for
    /// body text - and the outlining toolbar. A structure editor: lines are selected (click, Shift+click,
    /// the arrow keys), and promoted, demoted, given a level, moved (a collapsed heading with what is
    /// under it), expanded and collapsed (in the view only); a double click or Enter on a line goes back
    /// to print layout with the caret there.
    let outlineView (host: WordHost.Host) (viewType: Signal<int>) (showLevel: Signal<int>) (control: OutlineControl) : View =
        let collapsed = HashSet<string>()
        let expanded = HashSet<string>()
        let firstLineOnly = signal false
        let showFormatting = signal true
        let focusLevel = signal ""
        // (The lines shown, in order: their places among the body's blocks; which ids they are; the
        // element of each place, once made; the headings shown whose subtext is not.)
        let visible: int array ref = ref [||]
        let rows = HashSet<string>()
        let elementAt: HTMLElement array ref = ref [||]
        let hiddenUnder = HashSet<string>()
        // (A line's element is kept while its block and the way it is shown are the same: a command
        // makes again only the lines it changed, and the elements are moved, not made, when blocks move.)
        let rowCache = Dictionary<string, struct (obj * string * HTMLElement)>()
        let lines: HTMLElement ref = ref null
        let outline: Outline.DocumentOutline option ref = ref None
        let shownBlocks: obj ref = ref null
        let shownStyles: obj ref = ref null
        let shownFlags = ref ""
        let anchor: string ref = ref null
        let focus: string ref = ref null
        let wasActive = ref false
        let generation = ref 0
        let painted: (int * int) ref = ref (-1, -1)
        let paintedFocus: HTMLElement ref = ref null
        let active () = viewType.Value = 2

        let indexIn (o: Outline.DocumentOutline) (id: string) =
            if isNull id then -1
            else
                let mutable found = -1
                let mutable i = 0
                while found < 0 && i < o.BlockIds.Length do
                    if o.BlockIds.[i] = id then found <- i
                    i <- i + 1
                found

        let selectedRange () =
            match outline.Value with
            | Some o ->
                let a = indexIn o anchor.Value
                let f = indexIn o focus.Value
                if a < 0 || f < 0 then None else Some (o, min a f, max a f)
            | None -> None

        let elementOf (id: string) =
            match outline.Value with
            | Some o ->
                let i = indexIn o id
                if i >= 0 && i < elementAt.Value.Length then elementAt.Value.[i] else null
            | None -> null

        let markSelected (element: HTMLElement) (selected: bool) =
            if not (isNull element) && (element?__selected = true) <> selected then
                element?__selected <- selected
                element.setAttribute("aria-selected", if selected then "true" else "false")

        /// The selected lines and the focused one shown so (only the lines whose state changes are touched).
        let paintSelection () =
            let lo, hi = match selectedRange () with Some (_, lo, hi) -> lo, hi | None -> -1, -1
            let plo, phi = painted.Value
            let at = elementAt.Value
            if plo >= 0 then
                for i in plo .. min phi (at.Length - 1) do
                    if i < lo || i > hi then markSelected at.[i] false
            if lo >= 0 then
                for i in lo .. min hi (at.Length - 1) do markSelected at.[i] true
            painted.Value <- (lo, hi)
            let focused = elementOf focus.Value
            if not (obj.ReferenceEquals(focused, paintedFocus.Value)) then
                if not (isNull paintedFocus.Value) then paintedFocus.Value.classList.remove "rb-word-outline-focus"
                if not (isNull focused) then focused.classList.add "rb-word-outline-focus"
                paintedFocus.Value <- focused
            if isNull focused then focusLevel.Value <- ""
            else
                scrollNearest focused
                focusLevel.Value <- focused.getAttribute "data-level"

        /// The document's selection over the selected lines and what a collapsed heading among them holds.
        let selectInDocument (withHidden: bool) =
            match host.Editor, selectedRange () with
            | Some editor, Some (o, lo, hi) ->
                let mutable last = hi
                if withHidden then
                    for i in lo .. hi do
                        if hiddenUnder.Contains o.BlockIds.[i] then last <- max last (Outline.subtextEnd o i - 1)
                WordApi.selectBlocks editor o.BlockIds.[lo] o.BlockIds.[last] |> ignore
                Some (o, lo, last)
            | _ -> None

        let runsOf (resolve: Rbt.Word.Document.ParagraphBlock -> TextBuffer.TextStyle -> TextBuffer.TextStyle) (para: Rbt.Word.Document.ParagraphBlock) (body: HTMLElement) =
            let text = TextBuffer.toString para.Buffer
            let mutable any = false
            for run in para.Buffer.Runs do
                match run with
                | TextBuffer.Text textRun when textRun.Length > 0 && textRun.Start + textRun.Length <= text.Length && not textRun.Style.Hidden ->
                    let style = resolve para textRun.Style
                    let piece = element "span" ""
                    let chars = text.Substring(textRun.Start, textRun.Length).ToCharArray()
                    for k = 0 to chars.Length - 1 do
                        if chars.[k] < ' ' then chars.[k] <- ' '
                    piece.textContent <- System.String(chars)
                    let mutable css = "font-size:" + string (style.FontSize * 0.75) + "pt"
                    match style.FontFamily with
                    | Some family -> css <- css + ";font-family:\"" + family + "\""
                    | None -> ()
                    if style.Bold then css <- css + ";font-weight:700"
                    if style.Italic then css <- css + ";font-style:italic"
                    let decorations = (if TextBuffer.underlineShown style then "underline " else "") + (if style.Strike || style.DoubleStrike then "line-through" else "")
                    if decorations <> "" then css <- css + ";text-decoration:" + decorations
                    match style.Color with
                    | Some color -> css <- css + ";color:" + color
                    | None -> ()
                    piece.setAttribute("style", css)
                    body.appendChild piece |> ignore
                    any <- true
                | _ -> ()
            any

        let makeRow (block: Rbt.Word.Document.Block) (id: string) (level: int) (depth: int) (symbolKind: string) (collapsedMark: bool) (resolve: unit -> (Rbt.Word.Document.ParagraphBlock -> TextBuffer.TextStyle -> TextBuffer.TextStyle) option) =
            let isHeading = level < Outline.BodyLevel
            let kind = match block with Rbt.Word.Document.Table _ -> "table" | Rbt.Word.Document.TextBox _ -> "textbox" | _ -> "paragraph"
            let row = element "div" ("rb-word-outline-line" + (if isHeading then " rb-word-outline-heading" else "") + (if kind = "table" then " rb-word-outline-table" else "") + (if collapsedMark then " rb-word-outline-collapsed" else ""))
            row.setAttribute("role", "option")
            row.setAttribute("data-testid", "word-outline-line")
            row.setAttribute("data-block-id", id)
            row.setAttribute("data-level", string level)
            row.setAttribute("data-kind", kind)
            row.setAttribute("aria-selected", "false")
            row.setAttribute("style", "padding-left:" + string (depth * 24 + 30) + "px")
            let symbol = element "span" "rb-word-outline-symbol"
            symbol.setAttribute("data-testid", "word-outline-symbol")
            symbol.setAttribute("style", "left:" + string (depth * 24 + 6) + "px")
            symbol.setAttribute("data-symbol", symbolKind)
            symbol.setAttribute("aria-hidden", "true")
            symbol.textContent <- (match symbolKind with "plus" -> "⊕" | "minus" -> "⊖" | _ -> "◦")
            let body = element "span" "rb-word-outline-text"
            let plain (content: string) = body.textContent <- if content = "" then " " else content
            match block with
            | Rbt.Word.Document.Paragraph para | Rbt.Word.Document.Heading (_, para) ->
                match resolve () with
                | Some resolver -> if not (runsOf resolver para body) then plain ""
                | None -> plain (Outline.lineText para)
            | Rbt.Word.Document.Table table ->
                let cells =
                    table.Rows |> List.truncate 1 |> List.collect (fun r -> r.Cells)
                    |> List.map (fun cell -> cell.Content |> List.map Outline.lineText |> String.concat " ")
                plain ("Table: " + String.concat " | " cells)
            | Rbt.Word.Document.TextBox box -> plain (box.Content |> List.map Outline.lineText |> String.concat " ")
            row.appendChild symbol |> ignore
            row.appendChild body |> ignore
            row

        /// Put the elements of the lines shown from `first` to `last` (positions in `visible`) in their places.
        let place (first: int) (last: int) =
            let container = lines.Value
            let shown = visible.Value
            let at = elementAt.Value
            let mutable cursor: Node = if first = 0 then container.firstChild else (at.[shown.[first - 1]] :> Node).nextSibling
            for k in first .. last - 1 do
                let row = at.[shown.[k]]
                if obj.ReferenceEquals(cursor, row) then cursor <- cursor.nextSibling
                else container.insertBefore(row, cursor) |> ignore

        let clearLines () =
            generation.Value <- generation.Value + 1
            if not (isNull lines.Value) then lines.Value.textContent <- ""
            rowCache.Clear()
            rows.Clear()
            elementAt.Value <- [||]
            visible.Value <- [||]
            shownBlocks.Value <- null
            painted.Value <- (-1, -1)
            paintedFocus.Value <- null

        let render (force: bool) =
            if not (isNull lines.Value) then
                match host.Editor with
                | None -> clearLines ()
                | Some editor ->
                    let doc = editor.Whole
                    let formatting = showFormatting.Value
                    let flags = string showLevel.Value + (if formatting then "f" else "") + (if firstLineOnly.Value then "1" else "")
                    if force || not (obj.ReferenceEquals(box doc.Blocks, shownBlocks.Value) && obj.ReferenceEquals(box doc.Styles, shownStyles.Value) && flags = shownFlags.Value) then
                        if not (obj.ReferenceEquals(box doc.Styles, shownStyles.Value)) || flags <> shownFlags.Value then rowCache.Clear()
                        shownBlocks.Value <- box doc.Blocks
                        shownStyles.Value <- box doc.Styles
                        shownFlags.Value <- flags
                        let o = Outline.ofDocument doc
                        outline.Value <- Some o
                        let blocks = List.toArray doc.Blocks
                        let n = blocks.Length
                        let levels = o.Levels
                        let limit = showLevel.Value
                        if firstLineOnly.Value then lines.Value.classList.add "rb-word-outline-first-line" else lines.Value.classList.remove "rb-word-outline-first-line"
                        // Which lines are shown: none under a collapsed heading; under Show level, the
                        // levels to it and what an expanded heading has one level further.
                        rows.Clear()
                        hiddenUnder.Clear()
                        let shown = ResizeArray<int>()
                        let depths = Array.zeroCreate n
                        let stackLevel = ResizeArray<int>()
                        let stackShown = ResizeArray<bool>()
                        let stackCollapsed = ResizeArray<bool>()
                        let stackExpanded = ResizeArray<bool>()
                        let stackIds = ResizeArray<string>()
                        let mutable headingLevel = 0
                        for i = 0 to n - 1 do
                            let level = levels.[i]
                            let id = o.BlockIds.[i]
                            let isHeading = level < Outline.BodyLevel
                            if isHeading then
                                headingLevel <- level
                                while stackLevel.Count > 0 && stackLevel.[stackLevel.Count - 1] >= level do
                                    Portable.Stack.dropLast stackLevel; Portable.Stack.dropLast stackShown
                                    Portable.Stack.dropLast stackCollapsed; Portable.Stack.dropLast stackExpanded
                                    Portable.Stack.dropLast stackIds
                            depths.[i] <- if isHeading then level - 1 else headingLevel
                            let top = stackLevel.Count - 1
                            let parentShown = top < 0 || stackShown.[top]
                            let isShown =
                                parentShown && not (top >= 0 && stackCollapsed.[top])
                                && (limit = 0 || level <= limit || (top >= 0 && stackExpanded.[top]))
                            // (A heading shown with something under it that is not: shown as collapsed.)
                            if top >= 0 && not isShown && parentShown then hiddenUnder.Add stackIds.[top] |> ignore
                            if isHeading then
                                stackLevel.Add level
                                stackShown.Add isShown
                                stackCollapsed.Add (collapsed.Contains id)
                                stackExpanded.Add (expanded.Contains id)
                                stackIds.Add id
                            if isShown then
                                shown.Add i
                                rows.Add id |> ignore
                        visible.Value <- shown.ToArray()
                        let previous = elementAt.Value
                        for element in previous do
                            if not (isNull element) then
                                markSelected element false
                        if not (isNull paintedFocus.Value) then paintedFocus.Value.classList.remove "rb-word-outline-focus"
                        paintedFocus.Value <- null
                        painted.Value <- (-1, -1)
                        elementAt.Value <- Array.create n null
                        // (The selection on a line that is shown.)
                        if isNull focus.Value || not (rows.Contains focus.Value) then
                            let start = if shown.Count > 0 then o.BlockIds.[shown.[0]] else null
                            if isNull focus.Value || isNull start then (focus.Value <- start; anchor.Value <- start)
                            else
                                let at = indexIn o focus.Value
                                let mutable best = start
                                for v in shown do if v <= at then best <- o.BlockIds.[v]
                                focus.Value <- best
                                anchor.Value <- best
                        if isNull anchor.Value || not (rows.Contains anchor.Value) then anchor.Value <- focus.Value
                        let resolver = ref None
                        let resolve () =
                            if not formatting then None
                            else
                                if resolver.Value.IsNone then resolver.Value <- Some (WordStyles.effectiveRunStyleResolver doc)
                                resolver.Value
                        let used = HashSet<string>()
                        let make (k: int) =
                            let i = shown.[k]
                            let id = o.BlockIds.[i]
                            let level = levels.[i]
                            let symbolKind = if level >= Outline.BodyLevel then "body" elif i + 1 < n && levels.[i + 1] > level then "plus" else "minus"
                            let mark = hiddenUnder.Contains id
                            let key = string level + "|" + string depths.[i] + "|" + symbolKind + (if mark then "|c" else "")
                            let block = blocks.[i]
                            let row =
                                match Portable.Dict.tryFind id rowCache with
                                | ValueSome (struct (known, knownKey, row)) when obj.ReferenceEquals(known, box block) && knownKey = key -> row
                                | _ ->
                                    let row = makeRow block id level depths.[i] symbolKind mark resolve
                                    rowCache.[id] <- struct (box block, key, row)
                                    row
                            used.Add id |> ignore
                            elementAt.Value.[i] <- row
                        let finishAll () =
                            // (What is left after the last line: lines no longer shown.)
                            let at = elementAt.Value
                            let container = lines.Value
                            let mutable after: Node = if shown.Count = 0 then container.firstChild else (at.[shown.[shown.Count - 1]] :> Node).nextSibling
                            while not (isNull after) do
                                let next = after.nextSibling
                                container.removeChild after |> ignore
                                after <- next
                            for id in rowCache.Keys |> Seq.toArray do
                                if not (used.Contains id) then rowCache.Remove id |> ignore
                        // The lines at the start at once; the rest in slices (a long document's thousands of lines).
                        let gen = generation.Value + 1
                        generation.Value <- gen
                        let first = min shown.Count 400
                        for k in 0 .. first - 1 do make k
                        place 0 first
                        paintSelection ()
                        let rec slice (from: int) =
                            if generation.Value = gen && not (isNull lines.Value) then
                                let until = min shown.Count (from + 1500)
                                for k in from .. until - 1 do make k
                                place from until
                                if until < shown.Count then setTimer (fun () -> slice until) 0 |> ignore
                                else
                                    finishAll ()
                                    painted.Value <- (-1, -1)
                                    paintSelection ()
                        if first < shown.Count then setTimer (fun () -> slice first) 0 |> ignore
                        else finishAll ()
                    else paintSelection ()

        let visiblePosition (id: string) =
            match outline.Value with
            | Some o ->
                let index = indexIn o id
                let mutable at = -1
                for i = 0 to visible.Value.Length - 1 do
                    if visible.Value.[i] = index then at <- i
                at
            | None -> -1

        let selectLine (id: string) (extend: bool) =
            if not (isNull id) then
                focus.Value <- id
                if not extend || isNull anchor.Value then anchor.Value <- id
                paintSelection ()
                selectInDocument false |> ignore

        let toPrintLayout (id: string) =
            match host.Editor with
            | Some editor ->
                WordApi.setView editor 3
                if not (isNull id) then WordApi.goToHeading editor id |> ignore
                focusElement (WordApi.getInputElement editor) false
                host.Status.Value <- "Print layout"
            | None -> ()

        let afterCommand () =
            render false
            if rows.Contains focus.Value then focusElement lines.Value true

        let levelCommand (status: string) (command: WordApi.Editor -> bool) =
            match host.Editor with
            | Some editor ->
                selectInDocument true |> ignore
                let mutable changed = false
                WordHost.runCommand host (fun e -> changed <- command e)
                if changed then host.Status.Value <- status
                selectInDocument false |> ignore
                afterCommand ()
            | None -> ()

        let move (up: bool) =
            match host.Editor, selectInDocument true with
            | Some editor, Some (o, lo, last) ->
                let firstVisible = visiblePosition o.BlockIds.[lo]
                let target =
                    if up then
                        if firstVisible > 0 then Some visible.Value.[firstVisible - 1] else None
                    else
                        let next = visible.Value |> Array.tryFind (fun v -> v > last)
                        next |> Option.map (fun n -> if hiddenUnder.Contains o.BlockIds.[n] then Outline.subtextEnd o n else n + 1)
                match target with
                | Some before ->
                    let beforeId = if before >= o.BlockIds.Length then null else o.BlockIds.[before]
                    let mutable changed = false
                    WordHost.runCommand host (fun e -> changed <- WordApi.moveBlocks e o.BlockIds.[lo] o.BlockIds.[last] beforeId)
                    if changed then host.Status.Value <- (if up then "Moved up" else "Moved down")
                    selectInDocument false |> ignore
                    afterCommand ()
                | None -> selectInDocument false |> ignore
            | _ -> ()

        let expandOrCollapse (expand: bool) =
            match selectedRange () with
            | Some (o, lo, hi) ->
                let headings = [ for i in lo .. hi do if o.Levels.[i] < Outline.BodyLevel && rows.Contains o.BlockIds.[i] then yield i ]
                // (On body text, Collapse is of the heading it is under.)
                let headings =
                    if headings.IsEmpty && not expand then
                        let at = Outline.headingAt o lo
                        if at >= 0 then [ at ] else []
                    else headings
                for i in headings do
                    let id = o.BlockIds.[i]
                    let last = Outline.subtextEnd o i
                    let children = [ for j in i + 1 .. last - 1 do if o.Levels.[j] < Outline.BodyLevel then yield o.BlockIds.[j] ]
                    if expand then
                        if collapsed.Remove id then ()
                        elif showLevel.Value > 0 then
                            // (One level further each time, as Word.)
                            if expanded.Add id then ()
                            else for child in children do if rows.Contains child then expanded.Add child |> ignore
                    else
                        let deeper = children |> List.filter expanded.Contains
                        if not deeper.IsEmpty then (for child in deeper do expanded.Remove child |> ignore)
                        elif expanded.Remove id then ()
                        elif last > i + 1 then
                            collapsed.Add id |> ignore
                            if not (isNull focus.Value) && indexIn o focus.Value > i && indexIn o focus.Value < last then (focus.Value <- id; anchor.Value <- id)
                if not headings.IsEmpty then
                    // (The selection stays on lines that are shown.)
                    render true
                    selectInDocument false |> ignore
            | None -> ()

        let runCommand (name: string) =
            match name with
            | "promote" -> levelCommand "Promoted" WordApi.outlinePromote
            | "demote" -> levelCommand "Demoted" WordApi.outlineDemote
            | "promoteToHeading1" -> levelCommand "Promoted to Heading 1" WordApi.outlinePromoteToHeading1
            | "demoteToBody" -> levelCommand "Demoted to body text" WordApi.outlineDemoteToBody
            | "moveUp" -> move true
            | "moveDown" -> move false
            | "expand" -> expandOrCollapse true
            | "collapse" -> expandOrCollapse false
            | other when other.StartsWith "level:" ->
                match System.Int32.TryParse(other.Substring 6) with
                | true, level -> levelCommand "Outline level set" (fun e -> WordApi.setOutlineLevel e level)
                | _ -> ()
            | _ -> ()
        control.Run <- runCommand

        let setShowLevel (level: int) =
            collapsed.Clear()
            expanded.Clear()
            match host.Editor with
            | Some editor -> WordApi.setOutlineShowLevel editor level
            | None -> ()

        let onKey (event: Event) =
            let key = event :?> KeyboardEvent
            if not event.defaultPrevented then
                let move (step: int) =
                    let at = visiblePosition focus.Value
                    let next = max 0 (min (visible.Value.Length - 1) (at + step))
                    match outline.Value with
                    | Some o when visible.Value.Length > 0 -> selectLine o.BlockIds.[visible.Value.[next]] key.shiftKey
                    | _ -> ()
                let handled =
                    if key.altKey && key.shiftKey && not key.ctrlKey then
                        match key.key, key.code with
                        | ("+" | "=" | "Add"), _ | _, "NumpadAdd" -> runCommand "expand"; true
                        | ("-" | "_" | "Subtract"), _ | _, "NumpadSubtract" -> runCommand "collapse"; true
                        | "ArrowLeft", _ -> runCommand "promote"; true
                        | "ArrowRight", _ -> runCommand "demote"; true
                        | "ArrowUp", _ -> runCommand "moveUp"; true
                        | "ArrowDown", _ -> runCommand "moveDown"; true
                        | _, code when code.StartsWith "Digit" && code.Length = 6 && code.[5] >= '1' && code.[5] <= '9' -> setShowLevel (int code.[5] - int '0'); true
                        | _, "KeyA" -> setShowLevel 0; true
                        | _ -> false
                    elif not key.altKey && not key.ctrlKey && not key.metaKey then
                        match key.key with
                        | "ArrowDown" -> move 1; true
                        | "ArrowUp" -> move -1; true
                        | "PageDown" -> move 20; true
                        | "PageUp" -> move -20; true
                        | "Home" -> move -visible.Value.Length; true
                        | "End" -> move visible.Value.Length; true
                        | "Enter" -> toPrintLayout focus.Value; true
                        | _ -> false
                    else false
                if handled then
                    key.preventDefault()
                    key.stopPropagation()

        let toolButton (testId: string) (label: string) (title: string) (action: unit -> unit) =
            button [ cls "word-tool-button"; attr "type" "button"; attr "data-testid" testId; attr "title" title; attr "aria-label" title
                     onClick (fun _ -> action (); focusElement lines.Value true) ] [ text label ]

        div [ cls "word-outline-view"; attr "role" "region"; attr "aria-label" "Outline view"; attr "data-testid" "word-outline-view"
              styleR "display" (fun () -> if active () then "flex" else "none")
              refEffect (fun _ ->
                  let subscription = effect (fun () ->
                      let _ = host.SelectionVersion.Value
                      let isActive = active ()
                      let _ = showLevel.Value
                      let _ = showFormatting.Value
                      let _ = firstLineOnly.Value
                      untrack (fun () ->
                          if isActive then
                              if not wasActive.Value then
                                  // (Coming into outline view: the line the caret is in.)
                                  wasActive.Value <- true
                                  match host.Editor with
                                  | Some editor ->
                                      let doc = editor.Whole
                                      let caret = match doc.Selection with Some s -> Outline.indexOf doc s.Focus.BlockId | None -> -1
                                      let o = Outline.ofDocument doc
                                      focus.Value <- if caret >= 0 then o.BlockIds.[caret] else null
                                      anchor.Value <- focus.Value
                                      render true
                                      focusElement lines.Value true
                                  | None -> ()
                              else render false
                          elif wasActive.Value then
                              // (Leaving outline view: its lines go, a long document's thousands of elements with them.)
                              wasActive.Value <- false
                              clearLines ()))
                  fun () -> subscription.Dispose()) ] [
            div [ cls "word-outline-toolbar"; attr "role" "toolbar"; attr "aria-label" "Outlining"; attr "data-testid" "word-outline-toolbar" ] [
                toolButton "word-outline-promote-h1" "«" "Promote to Heading 1" (fun () -> runCommand "promoteToHeading1")
                toolButton "word-outline-promote" "‹" "Promote (Alt+Shift+Left)" (fun () -> runCommand "promote")
                select [ cls "word-tool-select"; attr "data-testid" "word-outline-level"; attr "aria-label" "Outline level"; attr "title" "Outline level"
                         propR "value" (fun () -> box focusLevel.Value)
                         onChange (fun event ->
                             let value = (event.target :?> HTMLSelectElement).value
                             if value <> "" then runCommand ("level:" + value)) ] [
                    yield option [ attr "value" "" ] [ text "—" ]
                    yield! [ for level in 1 .. 9 -> option [ attr "value" (string level) ] [ text ("Level " + string level) ] ]
                    yield option [ attr "value" "10" ] [ text "Body Text" ]
                ]
                toolButton "word-outline-demote" "›" "Demote (Alt+Shift+Right)" (fun () -> runCommand "demote")
                toolButton "word-outline-demote-body" "»" "Demote to body text" (fun () -> runCommand "demoteToBody")
                div [ cls "word-toolbar-separator"; attr "aria-hidden" "true" ] []
                toolButton "word-outline-move-up" "▲" "Move up (Alt+Shift+Up)" (fun () -> runCommand "moveUp")
                toolButton "word-outline-move-down" "▼" "Move down (Alt+Shift+Down)" (fun () -> runCommand "moveDown")
                toolButton "word-outline-expand" "+" "Expand (Alt+Shift+Plus)" (fun () -> runCommand "expand")
                toolButton "word-outline-collapse" "−" "Collapse (Alt+Shift+Minus)" (fun () -> runCommand "collapse")
                div [ cls "word-toolbar-separator"; attr "aria-hidden" "true" ] []
                label [] [
                    text "Show level"
                    select [ cls "word-tool-select"; attr "data-testid" "word-outline-show-level"; attr "aria-label" "Show level"
                             propR "value" (fun () -> box (string showLevel.Value))
                             onChange (fun event -> setShowLevel (int (event.target :?> HTMLSelectElement).value)) ] [
                        yield option [ attr "value" "0" ] [ text "All levels" ]
                        yield! [ for level in 1 .. 9 -> option [ attr "value" (string level) ] [ text ("Level " + string level) ] ]
                    ]
                ]
                label [] [
                    input [ typ "checkbox"; attr "data-testid" "word-outline-first-line"; propR "checked" (fun () -> box firstLineOnly.Value)
                            onChange (fun event -> firstLineOnly.Value <- (event.target :?> HTMLInputElement).``checked``) ]
                    text "Show first line only"
                ]
                label [] [
                    input [ typ "checkbox"; attr "data-testid" "word-outline-formatting"; propR "checked" (fun () -> box showFormatting.Value)
                            onChange (fun event -> showFormatting.Value <- (event.target :?> HTMLInputElement).``checked``) ]
                    text "Show text formatting"
                ]
                div [ cls "word-toolbar-separator"; attr "aria-hidden" "true" ] []
                button [ cls "word-tool-button"; attr "type" "button"; attr "data-testid" "word-outline-close"
                         onClick (fun _ -> toPrintLayout focus.Value) ] [ text "Close outline view" ]
            ]
            div [ cls "word-outline-lines"; attr "role" "listbox"; attr "aria-multiselectable" "true"; attr "aria-label" "Outline"
                  attr "tabindex" "0"; attr "data-testid" "word-outline-lines"
                  refEffect (fun element ->
                      lines.Value <- element :?> HTMLElement
                      lines.Value.addEventListener("mousedown", fun (event: Event) ->
                          let mouse = event :?> MouseEvent
                          let row = closest event.target ".rb-word-outline-line"
                          if not (isNull row) && mouse.button = 0.0 then
                              let id = row.getAttribute "data-block-id"
                              let symbol = closest event.target ".rb-word-outline-symbol"
                              if not (isNull symbol) && symbol.getAttribute "data-symbol" <> "body" then
                                  // (The symbol of a heading: its subtext shown or hidden.)
                                  selectLine id false
                                  if hiddenUnder.Contains id || collapsed.Contains id then runCommand "expand" else runCommand "collapse"
                              else selectLine id mouse.shiftKey
                              focusElement lines.Value true
                              event.preventDefault())
                      lines.Value.addEventListener("dblclick", fun (event: Event) ->
                          let row = closest event.target ".rb-word-outline-line"
                          if not (isNull row) && isNull (closest event.target ".rb-word-outline-symbol") then
                              toPrintLayout (row.getAttribute "data-block-id"))
                      lines.Value.addEventListener("keydown", onKey)
                      fun () -> lines.Value <- null) ] []
        ]

#endif
