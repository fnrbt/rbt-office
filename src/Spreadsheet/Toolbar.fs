module Rbt.Spreadsheet.Toolbar

#if FABLE_COMPILER
open Fable.Core
open Fable.Core.JsInterop
open Rbt.Spreadsheet.Dom
open Rbt.Office.Commands
open Rbt.FableUI.OfficeUiDialog
open Rbt.FableUI.Chrome
open Rbt.Office.Ui
open Rbt.Spreadsheet.SheetOps
open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Formula

[<Emit("!!($0.isComposing || $0.keyCode === 229 || $0.repeat || $0.defaultPrevented)")>]
let private unavailableKey (event: obj) : bool = jsNative

[<Emit("$0.ctrlKey || $0.metaKey")>]
let private ctrlOrMeta (event: obj) : bool = jsNative

[<Emit("!!$0.shiftKey")>]
let private keyShift (event: obj) : bool = jsNative

[<Emit("!!$0.altKey")>]
let private keyAlt (event: obj) : bool = jsNative

[<Emit("$0.tagName && /^(INPUT|TEXTAREA|SELECT)$/.test($0.tagName)")>]
let private isTextControl (target: Element) : bool = jsNative

[<Emit("$0.stopPropagation()")>]
let private stopEvent (event: obj) : unit = jsNative

[<Emit("Object.keys($0).length")>]
let private objectSize (value: obj) : int = jsNative

[<Emit("$0 == null")>]
let private isNullish (value: obj) : bool = jsNative

[<Emit("String($0)")>]
let private jsString (value: obj) : string = jsNative

[<Emit("$0.addEventListener('keydown', $1, true)")>]
let private addCaptureKey (target: Element) (listener: EventListener) : unit = jsNative
[<Emit("$0.removeEventListener('keydown', $1, true)")>]
let private removeCaptureKey (target: Element) (listener: EventListener) : unit = jsNative

[<Emit("$0.isConnected")>]
let private isConnected (element: Element) : bool = jsNative

[<Emit("String(($0 && $0.message) || $0)")>]
let private printError (error: obj) : string = jsNative
[<Emit("window.confirm($0)")>]
let private confirmReplace (message: string) : bool = jsNative

[<Emit("window")>]
let private browserWindow () : Element = jsNative

[<Emit("($0.files && $0.files[0]) || null")>]
let private selectedFile (input: Element) : obj = jsNative

[<Emit("(function(bytes, name) { const url = URL.createObjectURL(new Blob([bytes], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); try { const a = document.createElement('a'); a.href = url; a.download = name; a.style.display = 'none'; document.body.appendChild(a); try { a.click(); } finally { a.remove(); } } finally { setTimeout(() => URL.revokeObjectURL(url), 60000); } })($0, $1)")>]
let private downloadXlsx (bytes: byte[]) (name: string) : unit = jsNative

[<Emit("$0.click()")>]
let private clickElement (element: Element) : unit = jsNative

[<Emit("!!($0.ctrlKey || $0.metaKey) && !$0.altKey && !$0.shiftKey && !$0.isComposing && $0.keyCode !== 229")>]
let private fileShortcut (event: obj) : bool = jsNative

[<Emit("$0.isContentEditable || !!$0.closest('[contenteditable=true]')")>]
let private editingOtherContent (target: Element) : bool = jsNative
[<Emit("setTimeout($0, 0)")>]
let private deferFocus (action: unit -> unit) : unit = jsNative

[<Emit("typeof $0 === 'number' && Number.isInteger($0)")>]
let private isIntegerArg (value: obj) : bool = jsNative

[<Emit("(function(panel,trigger){const r=trigger.getBoundingClientRect(), w=Math.min(240,window.innerWidth-16); panel.style.width=w+'px';panel.style.left=Math.max(8,Math.min(r.left,window.innerWidth-w-8))+'px';panel.style.top=Math.max(8,Math.min(r.bottom+2,window.innerHeight-panel.offsetHeight-8))+'px';})($0,$1)")>]
let private positionSheetMenu (panel: Element) (trigger: Element) : unit = jsNative

[<Emit("($0.getBoundingClientRect().left + $0.getBoundingClientRect().width / 2)")>]
let private centerX (element: Element) : float = jsNative

[<Emit("($0.getBoundingClientRect().top + $0.getBoundingClientRect().height / 2)")>]
let private centerY (element: Element) : float = jsNative

[<Emit("!!($0.shiftKey && $0.key === 'F10') || $0.key === 'ContextMenu'")>]
let private sheetMenuKey (event: obj) : bool = jsNative

[<Emit("Number($0)")>]
let private asNumber (value: obj) : float = jsNative



// ── Spec types ──────────────────────────────────────────────

type MenuItem = {
    Id: string
    Label: string
    Shortcut: string option
    Action: unit -> unit
}

type MenuRow =
    | MItem of MenuItem
    | MSep

type Menu = {
    Id: string
    Label: string
    Rows: MenuRow list
}

type ToolbarItem =
    | TButton of id: string * label: string * action: (unit -> unit)
    | TSep

let private selectionStatusFields = [|
    "count", "Count"
    "numericCount", "Numeric count"
    "sum", "Sum"
    "average", "Average"
    "min", "Min"
    "max", "Max"
|]

// Only the shell and standalone formula bar carry these rules. The grid, host
// page and other mounted editor instances keep their own layout and styles.
let private shellStyles = """
.rb-office-shell.sheet-shell, .rb-office-shell.sheet-formula-bar { width: 100%; min-width: 0; max-width: 100%; }
.sheet-shell .sheet-file-name { display: block; padding: 8px 12px 4px; font-size: 13px; font-weight: 600; overflow-wrap: anywhere; }
.sheet-shell .sheet-menubar, .sheet-shell .sheet-toolbar, .sheet-shell .sheet-find-bar,
.rb-office-shell.sheet-formula-bar { display: flex; align-items: center; min-width: 0; max-width: 100%; gap: 4px; padding: 4px 8px; border-bottom: 1px solid var(--office-border); background: var(--office-surface); }
.sheet-shell .sheet-menubar { flex-wrap: wrap; }
.sheet-shell .sheet-toolbar { flex-wrap: wrap; gap: 4px; }
.sheet-shell .sheet-find-bar { flex-wrap: wrap; gap: 8px; }
.sheet-shell .sheet-tool-button, .sheet-shell .sheet-menu-button { white-space: nowrap; }
.sheet-shell .sheet-tool-button { min-width: 32px; }
.sheet-shell .sheet-menu-panel { position: absolute; top: 100%; z-index: 100; padding: 4px; background: var(--office-surface); border: 1px solid var(--office-border); border-radius: var(--office-radius); box-shadow: 0 4px 14px rgb(32 33 36 / 18%); }
.sheet-shell .sheet-menu-item { display: flex; align-items: center; justify-content: space-between; width: 100%; text-align: left; gap: 12px; white-space: normal; }
.sheet-shell .sheet-menu-item span + span { color: var(--office-muted); }
.sheet-shell .sheet-separator { width: 1px; height: 20px; margin: 0 4px; background: var(--office-border); flex: 0 0 auto; }
.sheet-shell .sheet-menu-separator { height: 1px; margin: 4px; background: var(--office-border); }
.sheet-shell .sheet-find-bar > label { display: inline-flex; align-items: center; flex: 1 1 210px; min-width: 0; gap: 6px; color: var(--office-muted); }
.sheet-shell .sheet-find-input, .sheet-shell .sheet-replace-input { flex: 1 1 120px; width: min(180px, 100%); min-width: 0; }
.sheet-shell .sheet-find-result, .sheet-shell .sheet-print-status { color: var(--office-muted); font-size: 12px; }
.sheet-shell .sheet-print-status { padding: 2px 8px; overflow-wrap: anywhere; }
.sheet-shell .sheet-selection-status { display: flex; flex-wrap: wrap; min-width: 0; max-width: 100%; gap: 4px 14px; padding: 4px 8px; border-bottom: 1px solid var(--office-border); background: var(--office-surface); font-size: 12px; }
.sheet-shell .sheet-selection-status > span { min-width: 0; max-width: 100%; overflow-wrap: anywhere; }
.rb-office-shell.sheet-shell input.sheet-color-picker { position: absolute; width: 1px; height: 1px; min-height: 0; padding: 0; border: 0; opacity: 0; pointer-events: none; }
.sheet-shell .sheet-tab-bar { display: flex; align-items: center; gap: 4px; min-width: 0; max-width: 100%; padding: 4px 8px; border-bottom: 1px solid var(--office-border); background: var(--office-surface); }
.sheet-shell .sheet-tab-list { display: flex; align-items: center; flex: 1 1 auto; min-width: 0; overflow-x: auto; gap: 2px; scrollbar-width: thin; }
.sheet-shell .sheet-tab, .sheet-shell .sheet-tab-action { flex: 0 0 auto; white-space: nowrap; min-height: 32px; }
.sheet-shell .sheet-tab[data-active='true'] { color: #217346; border-bottom: 2px solid #217346; font-weight: 600; }
.sheet-shell .sheet-tab-menu { position: fixed; z-index: 120; max-height: min(65dvh, 340px); overflow-y: auto; padding: 4px; background: var(--office-surface); border: 1px solid var(--office-border); border-radius: var(--office-radius); box-shadow: 0 4px 14px rgb(32 33 36 / 18%); }
.sheet-shell .sheet-tab-menu button { display: block; width: 100%; min-height: 32px; text-align: left; white-space: normal; }
.rb-office-shell.sheet-formula-bar .cell-ref-input { width: 68px; text-align: center; }
.rb-office-shell.sheet-formula-bar .formula-input { flex: 1 1 0; min-width: 0; }
.rb-office-shell.sheet-formula-bar { flex-wrap: wrap; }
.rb-office-shell.sheet-formula-bar .sheet-formula-argument-help { flex: 1 0 100%; font-size: 12px; color: var(--office-muted); overflow-wrap: anywhere; }
.sheet-shell .sheet-function-backdrop { position: fixed; inset: 0; z-index: 210; background: rgb(0 0 0 / 35%); display: flex; align-items: center; justify-content: center; padding: 12px; }
.sheet-shell .sheet-function-dialog { width: min(520px, 100%); max-height: min(80dvh, 660px); display: flex; flex-direction: column; gap: 8px; padding: 16px; background: var(--office-surface); border: 1px solid var(--office-border); border-radius: var(--office-radius); box-shadow: 0 8px 24px #0004; box-sizing: border-box; }
.sheet-shell .sheet-function-dialog h2 { font-size: 18px; margin: 0; }
.sheet-shell .sheet-function-search { width: 100%; min-width: 0; box-sizing: border-box; }
.sheet-shell .sheet-function-results { overflow-y: auto; min-height: 80px; max-height: 48dvh; display: flex; flex-direction: column; gap: 3px; }
.sheet-shell .sheet-function-choice { width: 100%; height: auto; flex-shrink: 0; text-align: left; white-space: normal; overflow-wrap: anywhere; padding: 7px; border: 1px solid var(--office-border); background: var(--office-surface); }
.sheet-shell .sheet-function-choice:focus-visible { outline: 2px solid #217346; }
.sheet-shell .sheet-function-description { display: block; font-size: 12px; color: var(--office-muted); }
.sheet-shell .sheet-function-cancel { align-self: flex-end; }
@media (max-width: 600px) {
  .sheet-shell .sheet-toolbar { max-height: 124px; overflow-y: auto; }
  .sheet-shell .sheet-menu-panel { max-height: min(60dvh, 320px); overflow-y: auto; }
}
"""

[<Emit("$0.removeAttribute($1)")>]
let private removeAttribute (element: Element) (name: string) : unit = jsNative

[<Emit("(function(panel,trigger) { const host = trigger.closest('.sheet-shell').getBoundingClientRect(); const button = trigger.getBoundingClientRect(); const width = Math.max(0, Math.min(240, host.width - 16)); panel.style.width = width + 'px'; panel.style.left = (Math.min(0, host.right - 8 - button.left - width)) + 'px'; })($0,$1)")>]
let private positionMenu (panel: Element) (trigger: Element) : unit = jsNative

// ── Document-level menu state ───────────────────────────────
// Only one menu panel can be open at a time; clicking outside closes it.

let mutable private openTrigger: Element option = None
let mutable private openPanel: Element option = None
let mutable private sheetMenuPanel: Element option = None
let mutable private sheetMenuTrigger: Element option = None

let private closeSheetMenu () =
    sheetMenuPanel |> Option.iter (fun panel -> setStyle panel "display" "none")
    sheetMenuTrigger |> Option.iter (fun trigger -> setAttribute trigger "aria-expanded" "false")
    sheetMenuPanel <- None
    sheetMenuTrigger <- None

let mutable private docHandler: EventListener option = None
let mutable private mountedMenus = 0

let private closeOpenMenu () =
    match openPanel with
    | Some panel -> setStyle panel "display" "none"
    | None -> ()
    match openTrigger with
    | Some trigger -> setAttribute trigger "aria-expanded" "false"
    | None -> ()
    openPanel <- None
    openTrigger <- None

let private installDocCloseHandler () =
    mountedMenus <- mountedMenus + 1
    if docHandler.IsNone then
        let handler: EventListener = fun _ ->
            closeOpenMenu ()
            closeSheetMenu ()
        addEventListener (getDocument ()) "click" handler
        docHandler <- Some handler

let private releaseDocCloseHandler (root: Element) =
    match openTrigger with
    | Some trigger when contains root trigger -> closeOpenMenu ()
    | _ -> ()
    match sheetMenuTrigger with
    | Some trigger when contains root trigger -> closeSheetMenu ()
    | _ -> ()
    mountedMenus <- mountedMenus - 1
    if mountedMenus = 0 then
        docHandler |> Option.iter (removeEventListener (getDocument ()) "click")
        docHandler <- None

// ── Toolbar primitives ──────────────────────────────────────

let private toolbarButton id label accessibleName action =
    let btn = createElement "button"
    setClassName btn "sheet-tool-button"
    setAttribute btn "data-sheet-icon" id
    setAttribute btn "type" "button"
    setAttribute btn "aria-label" accessibleName
    setTitle btn accessibleName
    setTextContent btn label
    addEventListener btn "mousedown" (fun e -> preventDefault e)
    addEventListener btn "click" (fun _ -> action ())
    btn

let private toolbarSeparator () =
    let sep = createElement "div"
    setClassName sep "sheet-separator"
    setAttribute sep "aria-hidden" "true"
    sep

// ── Menu rendering ──────────────────────────────────────────

let private renderMenuItem (item: MenuItem) (onClose: unit -> unit) =
    let row = createElement "button"
    setClassName row "sheet-menu-item"
    setAttribute row "data-sheet-menu-item" item.Id
    setAttribute row "type" "button"
    setTitle row item.Label
    let labelSpan = createElement "span"
    setTextContent labelSpan item.Label
    appendChild row labelSpan
    match item.Shortcut with
    | Some shortcut ->
        let sh = createElement "span"
        setAttribute sh "aria-hidden" "true"
        setTextContent sh shortcut
        appendChild row sh
    | None -> ()
    addEventListener row "click" (fun _ ->
        onClose ()
        item.Action ())
    row

let private renderMenuSeparator () =
    let sep = createElement "div"
    setClassName sep "sheet-menu-separator"
    setAttribute sep "aria-hidden" "true"
    sep

let private renderMenu (menu: Menu) =
    let wrap = createElement "div"
    setStyle wrap "position" "relative"

    let trigger = createElement "button"
    setClassName trigger "sheet-menu-button"
    setAttribute trigger "data-sheet-menu" menu.Id
    setAttribute trigger "type" "button"
    setAttribute trigger "aria-expanded" "false"
    setTitle trigger (menu.Label + " menu")
    setTextContent trigger menu.Label

    let panel = createElement "div"
    setClassName panel "sheet-menu-panel"
    setAttribute panel "data-sheet-menu-panel" menu.Id
    setStyle panel "display" "none"
    // Clicks inside the panel must not bubble to the document-level handler
    // that closes the menu. The item-click handler closes it explicitly.
    addEventListener panel "click" (fun e -> stopPropagation e)

    for row in menu.Rows do
        match row with
        | MItem item -> appendChild panel (renderMenuItem item (fun () -> closeOpenMenu (); focus trigger))
        | MSep -> appendChild panel (renderMenuSeparator ())

    addEventListener panel "keydown" (fun e ->
        if getEventKey e = "Escape" then
            preventDefault e
            stopEvent e
            closeOpenMenu ()
            focus trigger)
    addEventListener trigger "keydown" (fun e ->
        if getEventKey e = "Escape" && openTrigger = Some trigger then
            preventDefault e
            closeOpenMenu ())
    addEventListener trigger "click" (fun e ->
        stopPropagation e
        let wasOpen = openTrigger = Some trigger
        closeOpenMenu ()
        if not wasOpen then
            positionMenu panel trigger
            setStyle panel "display" "block"
            setAttribute trigger "aria-expanded" "true"
            openTrigger <- Some trigger
            openPanel <- Some panel)

    appendChild wrap trigger
    appendChild wrap panel
    wrap

// ── Unified spreadsheet shell ───────────────────────────────
// Replaces the per-category toolbar classes (FormatToolbar,
// AlignmentToolbar, NumberFormatToolbar, InsertDeleteToolbar,
// DataToolbar, ClipboardToolbar, FindReplaceToolbar) with one shell
// that renders a menubar and focused toolbar row. The grid owns the
// find/replace session and its controls, which the shell hosts.

[<AttachMembers>]
type SpreadsheetShell
        (grid: Rbt.Spreadsheet.Spreadsheet.Spreadsheet,
         onAction: string -> unit) =

    let mutable rootElement: Element option = None
    let mutable statusElement: Element option = None
    let mutable selectionStats: Element array = [||]
    let mutable statusSnapshot: obj = null
    let mutable printing = false
    let mutable generation = 0
    let mutable disposed = false
    let mutable fileName = "Workbook1.xlsx"
    let mutable nameElement: Element option = None
    let mutable fileInput: Element option = None
    let mutable saving = false
    let mutable openTicket = 0
    let mutable editSerial = 0
    let mutable fileSubscription: System.Action option = None
    let mutable editSubscription: System.Action option = None
    let mutable readOnlySubscription: System.Action option = None
    let mutable sheetTabsElement: Element option = None
    let mutable sheetPopup: Element option = None
    let mutable sheetsSubscription: System.Action option = None
    let sheetTabButtons = System.Collections.Generic.Dictionary<string, Element>()
    let mutable keyListener: EventListener option = None
    let mutable unloadListener: EventListener option = None
    let mutable functionChooser: Element option = None
    let mutable functionOpener: Element option = None

    let live () = not disposed && (rootElement |> Option.exists isConnected)
    let unsaved () = grid.IsModified() || grid.IsEditing()

    let refreshFileState () =
        if live () then
            nameElement |> Option.iter (fun element ->
                let dirty = unsaved ()
                let title = fileName + (if dirty then " • Unsaved changes" else "")
                setTextContent element title
                setTitle element title
                setAttribute element "data-sheet-filename" fileName
                setAttribute element "data-sheet-dirty" (if dirty then "true" else "false"))

    let setFileMessage message =
        if live () then
            statusElement |> Option.iter (fun status -> setTextContent status message)
    let refreshSelectionStatus () =
        if live () then
            let status = grid.GetStatus()
            if not (obj.ReferenceEquals(statusSnapshot, status)) then
                statusSnapshot <- status
                for index in 0 .. selectionStatusFields.Length - 1 do
                    let key, label = selectionStatusFields.[index]
                    let field = selectionStats.[index]
                    let value = getProp status key
                    let present = not (isNullish value)
                    setTextContent field (if present then label + ": " + jsString value else "")
                    setStyle field "display" (if present then "" else "none")

    let commands = CommandRegistry()
    // Export a stable, plain-JS command surface. Fable emits ordinary F#
    // members with transformed names; browser callers need callable methods.
    let jsCommands =
        createObj [
            "Run" ==> System.Func<string, obj, bool>(fun id argument ->
                commands.Run(id, argument = argument))
            "State" ==> System.Func<string, obj>(fun id ->
                match commands.State id with
                | None -> null
                | Some state ->
                    createObj [
                        "Enabled" ==> state.Enabled
                        "Checked" ==> (state.Checked |> Option.map box |> Option.defaultValue null)
                        "Mixed" ==> state.Mixed ])
            "OnChange" ==> System.Func<System.Action, System.Action>(fun callback ->
                System.Action(commands.OnChange(callback.Invoke))) ]
    let dialog = createState ()
    let mutable unmountDialog: (unit -> unit) option = None
    let mutable selectionSubscription: System.Action option = None
    let mutable formatSummary: obj = null
    let buttons = ResizeArray<string * Element>()
    let inputs = ResizeArray<string * Element>()
    let closeFunctionChooser restoreFocus =
        let opener = functionOpener
        functionChooser |> Option.iter (fun backdrop ->
            let parent = getProp backdrop "parentNode"
            if isNotNull parent then removeChild parent backdrop)
        functionChooser <- None
        functionOpener <- None
        if restoreFocus then opener |> Option.filter isConnected |> Option.iter focus

    let openFunctionChooser () =
        if live () && not (grid.IsReadOnly()) then
            let snapshot = grid.GetEditSnapshot()
            let entrySerial = editSerial
            if not (isNullish snapshot) then
                closeFunctionChooser false
                let currentFocus = getProp (getDocument ()) "activeElement"
                functionOpener <- if isNullish currentFocus then None else Some (unbox<Element> currentFocus)
                let backdrop = createElement "div"
                setClassName backdrop "sheet-function-backdrop"
                setAttribute backdrop "data-sheet-function-chooser" "true"
                let panel = createElement "section"
                setClassName panel "sheet-function-dialog"
                setAttribute panel "role" "dialog"
                setAttribute panel "aria-modal" "true"
                setAttribute panel "aria-label" "Insert function"
                let heading = createElement "h2"
                setTextContent heading "Insert function"
                appendChild panel heading
                let search = createElement "input"
                setClassName search "sheet-function-search"
                setAttribute search "type" "search"
                setAttribute search "aria-label" "Search supported functions"
                setAttribute search "placeholder" "Search name, description or arguments"
                appendChild panel search
                let results = createElement "div"
                setClassName results "sheet-function-results"
                setAttribute results "role" "group"
                appendChild panel results
                let cancel = createElement "button"
                setClassName cancel "sheet-function-cancel"
                setAttribute cancel "type" "button"
                setAttribute cancel "data-sheet-function-cancel" "true"
                setTextContent cancel "Cancel"
                addEventListener cancel "click" (fun _ -> closeFunctionChooser true)
                appendChild panel cancel
                appendChild backdrop panel
                let choices =
                    functionHelpCatalog |> Array.map (fun info ->
                        let button = createElement "button"
                        setClassName button "sheet-function-choice"
                        setAttribute button "type" "button"
                        setAttribute button "data-sheet-function" info.Name
                        let syntax = createElement "strong"
                        setTextContent syntax (info.Name + "(" + (info.Arguments |> Array.map (fun arg -> arg.Split(':').[0]) |> String.concat ", ") + ")")
                        appendChild button syntax
                        let detail = createElement "span"
                        setClassName detail "sheet-function-description"
                        setTextContent detail (info.Description + " " + String.concat "; " info.Arguments)
                        appendChild button detail
                        addEventListener button "click" (fun _ ->
                            let opener = functionOpener
                            closeFunctionChooser false
                            if editSerial <> entrySerial || not (grid.InsertFunctionAt(info.Name, snapshot)) then
                                setFileMessage "Function insertion cancelled: the cell or workbook changed."
                                opener |> Option.filter isConnected |> Option.iter focus)
                        appendChild results button
                        info, button)
                let filtered () =
                    let term = (getValue search).Trim().ToLowerInvariant()
                    choices |> Array.choose (fun (info, button) ->
                        let haystack = (info.Name + " " + info.Description + " " + String.concat " " info.Arguments).ToLowerInvariant()
                        let visible = term = "" || haystack.Contains(term)
                        setStyle button "display" (if visible then "" else "none")
                        if visible then Some button else None)
                addEventListener search "input" (fun _ -> filtered () |> ignore)
                addEventListener backdrop "click" (fun e ->
                    if obj.ReferenceEquals(box (getEventTarget e), box backdrop) then closeFunctionChooser true)
                addEventListener panel "keydown" (fun e ->
                    match getEventKey e with
                    | "Escape" ->
                        preventDefault e
                        stopEvent e
                        closeFunctionChooser true
                    | "Enter" when obj.ReferenceEquals(box (getEventTarget e), box search) ->
                        preventDefault e
                        let matches = filtered ()
                        if matches.Length > 0 then clickElement matches.[0]
                    | "ArrowDown" when obj.ReferenceEquals(box (getEventTarget e), box search) ->
                        let matches = filtered ()
                        if matches.Length > 0 then
                            preventDefault e
                            focus matches.[0]
                    | "Tab" ->
                        let matches = filtered ()
                        let sequence = Array.concat [| [| search |]; matches; [| cancel |] |]
                        let focused = getProp (getDocument ()) "activeElement"
                        let current = sequence |> Array.tryFindIndex (fun item -> obj.ReferenceEquals(box item, focused))
                        match current with
                        | Some index when keyShift e && index = 0 -> preventDefault e; focus cancel
                        | Some index when not (keyShift e) && index = sequence.Length - 1 -> preventDefault e; focus search
                        | _ -> ()
                    | _ -> ())
                functionChooser <- Some backdrop
                rootElement |> Option.iter (fun root -> appendChild root backdrop)
                focus search


    let currentFormat () =
        if isNull formatSummary then grid.GetSelectionFormat() else formatSummary
    let formatValue name = getProp (currentFormat ()) name
    let formatText name =
        let value = formatValue name
        if isNullish value then "" else jsString value
    let checkedFlag name () =
        match formatText name with
        | "mixed" -> None
        | "true" -> Some true
        | _ -> Some false
    let refreshCommands () =
        formatSummary <- grid.GetSelectionFormat()
        for id, button in buttons do
            commands.State(id) |> Option.iter (fun state ->
                setProp button "disabled" (box (not state.Enabled))
                match id with
                | "format.bold" | "format.italic" | "format.underline" ->
                    setAttribute button "aria-pressed"
                        (if state.Mixed then "mixed" elif state.Checked = Some true then "true" else "false")
                | _ -> removeAttribute button "aria-pressed")
        for key, input in inputs do
            let value = formatText key
            setValue input (if value = "mixed" then "" elif key = "fontSize" && value = "" then "11" else value)
        commands.NotifyChanged()

    let openFormatDialog () =
        if not (grid.IsReadOnly()) && grid.TryCommitEdit() then
            let summary = grid.GetSelectionFormat()
            let read name =
                let value = getProp summary name
                if isNullish value then ""
                else
                    let text = jsString value
                    if text = "mixed" then "<mixed>" else text
            let border = getProp summary "borders"
            let readBorder side key =
                let value = getProp border side
                if isNullish value then (if key = "style" then "none" else "")
                elif jsString value = "mixed" then "<mixed>"
                else jsString (getProp value key)
            let active = grid.GetActiveCell()
            let cell = grid.GetCell(unbox<int> (getProp active "Row"), unbox<int> (getProp active "Col"))
            let singleCell = grid.GetSelectionRef() = unbox<string> (getProp cell "address")
            let diagonalFlag key =
                if singleCell then jsString (getProp (getProp (getProp cell "format") "borders") key)
                else "<mixed>"
            let initial =
                [ for key in [ "bold"; "italic"; "strikethrough"; "wrap"; "shrinkToFit"
                               "fontFamily"; "fontSize"; "textColor"; "fill"; "horizontalAlign"
                               "verticalAlign"; "numberFormat"; "indent"; "textRotation" ] do
                      yield key, read key
                  yield "underline", (if read "underline" = "<mixed>" || read "doubleUnderline" = "<mixed>" then "<mixed>"
                                      elif read "doubleUnderline" = "true" then "double"
                                      elif read "underline" = "true" then "single" else "none")
                  for side in [ "top"; "bottom"; "left"; "right"; "diagonal" ] do
                      yield side + "-style", readBorder side "style"
                      yield side + "-color", readBorder side "color"
                  yield "diagonalUp", diagonalFlag "diagonalUp"
                  yield "diagonalDown", diagonalFlag "diagonalDown" ]
                |> Map.ofList
            let spec = Rbt.Spreadsheet.Ui.selectionFormatDialog initial (fun fields ->
                try
                    let patch = createObj []
                    let changed key = Map.find key fields <> Map.find key initial && Map.find key fields <> "<mixed>"
                    for key in [ "bold"; "italic"; "strikethrough"; "wrap"; "shrinkToFit" ] do
                        if changed key then setProp patch key (box (Map.find key fields = "true"))
                    for key in [ "fontFamily"; "fontSize"; "textColor"; "fill"; "horizontalAlign"; "verticalAlign"
                                 "numberFormat"; "indent"; "textRotation" ] do
                        if changed key then
                            let value = (Map.find key fields).Trim()
                            let item =
                                if key = "indent" || key = "textRotation" then box (System.Double.Parse value)
                                elif key = "fontSize" then if value = "" then null else box (System.Double.Parse value)
                                else if value = "" then null else box value
                            setProp patch key item
                    if changed "underline" then
                        let value = Map.find "underline" fields
                        setProp patch "underline" (box (value <> "none"))
                        setProp patch "doubleUnderline" (box (value = "double"))
                    let borders = createObj []
                    for side in [ "top"; "bottom"; "left"; "right"; "diagonal" ] do
                        let styleKey, colorKey = side + "-style", side + "-color"
                        if changed styleKey || changed colorKey then
                            let style, color = Map.find styleKey fields, (Map.find colorKey fields).Trim()
                            if style = "<mixed>" && changed colorKey then
                                invalidArg colorKey "Choose a border style before changing the color of mixed borders."
                            if style = "none" && changed colorKey then
                                invalidArg colorKey "Choose a border style before setting its color."
                            if changed styleKey && style <> "none" && color = "<mixed>" then
                                invalidArg colorKey "Choose a border color before changing the style of mixed colors."
                            let existing = getProp border side
                            let oldStyle key fallback =
                                if isNullish existing || jsString existing = "mixed" then fallback
                                else jsString (getProp existing key)
                            if style = "none" then setProp borders side null
                            else setProp borders side (createObj [
                                "style" ==> (if style = "<mixed>" then oldStyle "style" "thin" else style)
                                "color" ==> (if color = "<mixed>" || color = "" then oldStyle "color" "#000000" else color) ])
                    for key in [ "diagonalUp"; "diagonalDown" ] do
                        if changed key then setProp borders key (box (Map.find key fields = "true"))
                    if objectSize borders > 0 then setProp patch "borders" borders
                    if objectSize patch > 0 then grid.ApplyFormatPatch patch
                    true
                with error ->
                    setFileMessage ("Format failed: " + printError (box error))
                    match dialog.Open.Peek() with
                    | Some opened ->
                        for tab in [ "number"; "alignment"; "font"; "border"; "fill" ] do
                            opened.Values.["error-" + tab].Value <- printError (box error)
                    | None -> ()
                    false)
            openDialog dialog spec
        elif grid.IsEditing() then setFileMessage "Format blocked: correct the cell entry first."

    let dialogError (error: obj) =
        let message = printError error
        match dialog.Open.Peek() with
        | Some opened -> opened.Values.["error"].Value <- message
        | None -> ()
        false

    let dataReady epoch serial sheetIndex =
        if grid.IsReadOnly() then invalidOp "The workbook is read-only."
        if grid.GetWorkbookEpoch() <> epoch || grid.GetChangeSerial() <> serial
           || grid.GetActiveSheetIndex() <> sheetIndex then
            invalidOp "The workbook or active sheet changed while this dialog was open; reopen it for the current sheet."
        if not (grid.TryCommitEdit()) then invalidOp "Correct the cell entry before applying."

    let openDataDialog (make: int -> int -> int -> DialogSpec) =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                try openDialog dialog (make (grid.GetWorkbookEpoch()) (grid.GetChangeSerial()) (grid.GetActiveSheetIndex()))
                with error -> setFileMessage ("Data dialog failed: " + printError (box error))
            else setFileMessage "Data blocked: correct the cell entry first."

    let textOf (value: obj) =
        if isNullish value then "" else jsString value
    let property (value: obj) name =
        if isNullish value then null else getProp value name
    let flagOf value name fallback =
        let item = property value name
        if isNullish item then fallback else unbox<bool> item
    let optionalText (text: string) =
        if text.Trim() = "" then null else box (text.Trim())
    let selectedRange () = grid.GetSelectionRef()
    let checkedRange label (value: string) =
        let area = parseRange label (value.Trim())
        let model = grid.GetModel()
        if area.From.Row >= model.RowCount || area.From.Col >= model.ColCount then
            invalidArg "range" (sprintf "%s is outside this sheet." value)
        area
    let options names = names |> List.map (fun (value, label) -> choice value label)
    let select id label names value = dialogSelect id label (options names) value

    let chartDialog () =
        let snapshot = grid.GetChartSelectionSnapshot()
        if isNullish snapshot then invalidOp "The active worksheet is read-only."
        let range = selectedRange ()
        let area = checkedRange "Chart" range
        let valueType row col = textOf (property (grid.GetCell(row, col)) "valueType")
        let headers = area.From.Row < area.To.Row && valueType area.From.Row area.From.Col = "string"
        let categories = area.From.Col < area.To.Col && valueType (area.From.Row + (if headers then 1 else 0)) area.From.Col = "string"
        let submit (fields: Map<string, string>) =
            try
                let source = fields.["range"].Trim()
                let options = createObj [
                    "type" ==> fields.["type"]
                    "title" ==> fields.["title"]
                    "firstRowHeaders" ==> (fields.["headers"] = "true")
                    "firstColumnCategories" ==> (fields.["categories"] = "true") ]
                grid.InsertChart(box source, options, snapshot) |> ignore
                true
            with error -> dialogError (box error)
        { TestId = "xl-dlg-insert-chart"
          Title = "Insert chart"
          Subtitle = Some "Select one rectangular source. Each value column is a series; a pie takes exactly one."
          Fields = [
              dialogText "range" "Source range" "A1:C8" range
              select "type" "Chart type" [
                  "column", "Column"; "bar", "Horizontal bar"; "line", "Line"; "pie", "Pie" ] "column"
              dialogText "title" "Chart title (optional)" "Sales by quarter" ""
              dialogCheckbox "headers" "First row contains series names" headers
              dialogCheckbox "categories" "First column contains categories" categories
              dialogStatus "error" "Chart" "" ]
          Tabs = []
          Buttons = [ dialogCancel; dialogButton "xl-dlg-insert-chart-apply" "Insert" true submit ] }

    let validationDialog epoch serial sheetIndex =
        let range = selectedRange ()
        let area = checkedRange "Data validation" range
        let active = grid.GetActiveCell()
        let current = grid.GetValidationAt(getProp active "Row", getProp active "Col")
        let covered (target: CellRange) =
            not (isNullish current) &&
            ((textOf (property current "range")).Split([| ' '; ',' |], System.StringSplitOptions.RemoveEmptyEntries)
             |> Array.exists (fun text ->
                 match tryParseRange text with
                 | Some existing ->
                     existing.From.Row <= target.From.Row && existing.From.Col <= target.From.Col
                     && existing.To.Row >= target.To.Row && existing.To.Col >= target.To.Col
                 | None -> false))
        let coversSelection = covered area
        let rule = if coversSelection then current else null
        let field key = textOf (property rule key)
        let kind = if isNullish rule then "any" else field "type"
        let listValues =
            let given = property rule "values"
            if isNullish given then "" else String.concat "\n" (unbox<string array> given)
        let source = field "source"
        let mutable initialFields = Map.empty<string, string>
        let submit (fields: Map<string, string>) =
            try
                dataReady epoch serial sheetIndex
                let target = (Map.find "range" fields).Trim()
                let targetArea = checkedRange "Data validation" target
                let kind = Map.find "type" fields
                let rule = createObj [
                    "type" ==> kind
                    "allowBlank" ==> (Map.find "allowBlank" fields = "true")
                    "showInput" ==> (Map.find "showInput" fields = "true")
                    "inputTitle" ==> optionalText fields.["inputTitle"]
                    "inputMessage" ==> optionalText fields.["inputMessage"]
                    "showError" ==> (Map.find "showError" fields = "true")
                    "errorStyle" ==> fields.["errorStyle"]
                    "errorTitle" ==> optionalText fields.["errorTitle"]
                    "errorMessage" ==> optionalText fields.["errorMessage"] ]
                match kind with
                | "list" ->
                    let source = fields.["source"].Trim()
                    let values = fields.["values"].Split([| '\n'; '\r' |], System.StringSplitOptions.RemoveEmptyEntries)
                                 |> Array.map (fun value -> value.Trim()) |> Array.filter ((<>) "")
                    if source <> "" && values.Length > 0 then
                        invalidArg "source" "Choose either list values or a source, not both."
                    if source <> "" then setProp rule "source" (box source)
                    else
                        if values.Length = 0 then invalidArg "values" "Enter list values, one per line, or a source."
                        if values |> Array.exists (fun value -> value.Contains ",") then
                            invalidArg "values" "List values cannot contain commas; use a source range."
                        setProp rule "values" (box values)
                    setProp rule "dropdown" (box (fields.["dropdown"] = "true"))
                | "custom" ->
                    if fields.["formula"].Trim() = "" then invalidArg "formula" "Enter a custom validation formula."
                    setProp rule "formula" (box (fields.["formula"].Trim()))
                | "any" -> ()
                | _ ->
                    let operator = fields.["operator"]
                    if fields.["formula1"].Trim() = "" then invalidArg "formula1" "Enter the first comparison operand."
                    if (operator = "between" || operator = "notBetween") && fields.["formula2"].Trim() = "" then
                        invalidArg "formula2" "Enter the second comparison operand."
                    setProp rule "operator" (box operator)
                    setProp rule "formula1" (box (fields.["formula1"].Trim()))
                    if operator = "between" || operator = "notBetween" then
                        setProp rule "formula2" (box (fields.["formula2"].Trim()))
                let desired = validationOfDescriptionIn (grid.GetModel().Date1904) "Data validation"
                                  (Rbt.Spreadsheet.HostShapes.validationOfObj "Data validation" rule)
                let unchanged =
                    covered targetArea
                    && desired = validationOfDescriptionIn (grid.GetModel().Date1904) "Data validation"
                                     (Rbt.Spreadsheet.HostShapes.validationOfObj "Data validation" current)
                if Map.remove "error" fields <> Map.remove "error" initialFields && not unchanged then
                    grid.SetValidation(box target, rule) |> ignore
                true
            with error -> dialogError (box error)
        let spec = {
            TestId = "xl-dlg-data-validation"
            Title = "Data validation"
            Subtitle = Some "Apply replaces validation on the specified cells only. Other cells keep their rules."
            Fields = [
                dialogText "range" "Apply to range" "B2:B20" range
                select "type" "Allow" [
                    "any", "Any value"; "list", "List"; "whole", "Whole number"; "decimal", "Decimal"
                    "date", "Date"; "time", "Time"; "textLength", "Text length"; "custom", "Custom formula" ] kind
                select "operator" "Comparison (number/date/time/text length)" [
                    "between", "Between"; "notBetween", "Not between"; "equal", "Equal"
                    "notEqual", "Not equal"; "greaterThan", "Greater than"; "lessThan", "Less than"
                    "greaterThanOrEqual", "Greater than or equal"; "lessThanOrEqual", "Less than or equal" ]
                    (if field "operator" = "" then "between" else field "operator")
                dialogText "formula1" "First operand" "Number, date, reference or formula" (field "formula1")
                dialogText "formula2" "Second operand (between only)" "Number, date, reference or formula" (field "formula2")
                dialogTextArea "values" "List values (one per line)" "Yes\nNo" 3 listValues
                dialogText "source" "List source (instead of values)" "A1:A20 or named range" source
                dialogText "formula" "Custom formula" "A1>0" (if kind = "custom" then field "formula" else "")
                dialogCheckbox "dropdown" "Show list dropdown" (flagOf rule "dropdown" true)
                dialogCheckbox "allowBlank" "Allow blank" (flagOf rule "allowBlank" true)
                dialogCheckbox "showInput" "Show input message" (flagOf rule "showInput" true)
                dialogText "inputTitle" "Input title" "" (field "inputTitle")
                dialogTextArea "inputMessage" "Input message" "" 2 (field "inputMessage")
                dialogCheckbox "showError" "Show error alert" (flagOf rule "showError" true)
                select "errorStyle" "Error style" [
                    "stop", "Stop"; "warning", "Warning"; "information", "Information" ]
                    (if field "errorStyle" = "" then "stop" else field "errorStyle")
                dialogText "errorTitle" "Error title" "" (field "errorTitle")
                dialogTextArea "errorMessage" "Error message" "" 2 (field "errorMessage")
                dialogStatus "error" "Validation" "" ]
            Tabs = []
            Buttons = [
                dialogCancel
                dialogButton "xl-dlg-data-validation-clear" "Clear validation" false (fun fields ->
                    try
                        dataReady epoch serial sheetIndex
                        let target = fields.["range"].Trim()
                        checkedRange "Clear validation" target |> ignore
                        grid.ClearValidation(box target) |> ignore
                        true
                    with error -> dialogError (box error))
                dialogButton "xl-dlg-data-validation-apply" "Apply" true submit ]
        }
        initialFields <- spec.Fields |> List.map (fun item -> item.Id, item.Initial) |> Map.ofList
        spec

    let sortDialog epoch serial sheetIndex =
        let selected = selectedRange ()
        let active = grid.GetActiveCell()
        let addr = { Row = unbox<int> (getProp active "Row"); Col = unbox<int> (getProp active "Col") }
        let model = grid.GetModel()
        let quick = if tryParseCell selected |> Option.isSome then quickSortTarget addr model else None
        let range = match quick with Some (region, _) -> rangeText region | None -> selected
        let area = checkedRange "Sort" range
        let col = addr.Col
        let initialColumn =
            columnToLetter (if col >= area.From.Col && col <= area.To.Col then col else area.From.Col)
        let sortColumn (text: string) =
            let text = text.Trim().ToUpperInvariant()
            match tryParseRange (text + ":" + text) with
            | Some area when isWholeColumns area -> area.From.Col
            | _ -> invalidArg "column" (sprintf "Enter a column letter (for example A or BC), not \"%s\"." text)
        let submit (fields: Map<string, string>) =
            try
                dataReady epoch serial sheetIndex
                let target = fields.["range"].Trim()
                let area = checkedRange "Sort" target
                let key index =
                    let col = sortColumn fields.["column" + index]
                    if col < area.From.Col || col > area.To.Col then
                        invalidArg "column" (sprintf "Sort column %s must be inside %s." (columnToLetter col) target)
                    let by = fields.["by" + index]
                    let color = fields.["color" + index].Trim()
                    if by <> "value" && color <> "" && not (System.Text.RegularExpressions.Regex.IsMatch(color, "^#[0-9a-fA-F]{6}$")) then
                        invalidArg "color" "A sort color must be #RRGGBB (empty means no fill or automatic text color)."
                    let item = createObj [ "column" ==> col; "descending" ==> (fields.["direction" + index] = "desc"); "by" ==> by ]
                    if by <> "value" then setProp item "color" (if color = "" then null else box color)
                    item
                let keys = [| key "1" |]
                let second = fields.["column2"].Trim()
                let keys = if second = "" then keys else [| keys.[0]; key "2" |]
                let options = createObj [
                    "hasHeader" ==> (fields.["hasHeader"] = "true")
                    "caseSensitive" ==> (fields.["caseSensitive"] = "true") ]
                // The engine checks merged cells and spills before any mutation.
                grid.SortRange(box target, box keys, options = options)
                true
            with error -> dialogError (box error)
        let keyFields index column =
            let direction =
                if index = "1" && (model.SortState |> Option.exists (fun state -> state.Column = col && not state.Ascending))
                then "desc" else "asc"
            [ dialogText ("column" + index) ("Key " + index + " column (letters)") "A" column
              select ("direction" + index) ("Key " + index + " order") [ "asc", "Ascending"; "desc", "Descending" ] direction
              select ("by" + index) ("Key " + index + " sort by") [
                  "value", "Value"; "cellColor", "Cell color"; "fontColor", "Font color" ] "value"
              dialogText ("color" + index) ("Key " + index + " color (when sorting by color)") "#RRGGBB; empty means none" "" ]
        {
            TestId = "xl-dlg-data-sort"; Title = "Custom sort"
            Subtitle = Some "Rows of the range move together. An optional second key breaks ties."
            Fields =
                [ dialogText "range" "Range to sort" "A1:D20" range
                  dialogCheckbox "hasHeader" "First row is a header"
                      (quick |> Option.exists (fun (_, options) -> options.HasHeader))
                  dialogCheckbox "caseSensitive" "Case-sensitive text" false ]
                @ keyFields "1" initialColumn @ keyFields "2" "" @ [ dialogStatus "error" "Validation" "" ]
            Tabs = []
            Buttons = [ dialogCancel; dialogButton "xl-dlg-data-sort-apply" "Sort" true submit ]
        }

    let filterDialog epoch serial sheetIndex =
        let existing = grid.GetFilters()
        let selected = selectedRange ()
        let active = grid.GetActiveCell()
        let addr = { Row = unbox<int> (getProp active "Row"); Col = unbox<int> (getProp active "Col") }
        let range =
            if not (isNullish existing) then textOf (property existing "range")
            elif (tryParseCell selected).IsSome then
                match currentRegion addr (grid.GetModel()) with
                | Some region -> rangeText region
                | None -> selected
            else selected
        let area = checkedRange "Filter" range
        let activeCol = if addr.Col >= area.From.Col && addr.Col <= area.To.Col then addr.Col else area.From.Col
        let columns = property existing "columns"
        let entries = if isNullish columns then [||] else unbox<obj array> columns
        let current = entries |> Array.tryFind (fun entry -> unbox<int> (property entry "column") = activeCol)
                      |> Option.map (fun entry -> property entry "criteria") |> Option.defaultValue null
        let kind =
            if isNullish current then "values"
            elif Rbt.Spreadsheet.HostShapes.hasKey current "kept" then "kept"
            else [ "values"; "conditions"; "top"; "color"; "dynamic" ]
                 |> List.tryFind (Rbt.Spreadsheet.HostShapes.hasKey current)
                 |> Option.defaultValue "values"
        let condition index name =
            let conditions = property current "conditions"
            if isNullish conditions then null
            else
                let items = unbox<obj array> conditions
                if index >= items.Length then null else property items.[index] name
        let top = property current "top"
        let values =
            let items = property current "values"
            if isNullish items then ""
            else String.concat "\n" (unbox<obj array> items |> Array.map textOf)
        let submit (fields: Map<string, string>) =
            try
                dataReady epoch serial sheetIndex
                let target = fields.["range"].Trim()
                let area = checkedRange "Filter" target
                let columnText = fields.["column"].Trim().ToUpperInvariant()
                let col =
                    match tryParseRange (columnText + ":" + columnText) with
                    | Some colRange when isWholeColumns colRange -> colRange.From.Col
                    | _ -> invalidArg "column" "Enter a column letter, such as A or BC."
                if col < area.From.Col || col > area.To.Col then
                    invalidArg "column" "The filter column must be inside the filter range."
                let kind = fields.["kind"]
                if kind = "kept" then invalidArg "kind" "This imported filter cannot be edited; choose a supported filter type first."
                let criteria =
                    match kind with
                    | "values" ->
                        let values = fields.["values"].Split([| '\n'; '\r' |], System.StringSplitOptions.RemoveEmptyEntries)
                                     |> Array.map (fun text -> text.Trim()) |> Array.filter ((<>) "")
                        if values.Length = 0 && fields.["blanks"] <> "true" then
                            invalidArg "values" "Enter at least one visible value or select blanks."
                        createObj [ "values" ==> values; "blanks" ==> (fields.["blanks"] = "true") ]
                    | "conditions" ->
                        let one index =
                            let value = fields.["value" + index].Trim()
                            if value = "" && fields.["operator" + index] <> "equals" && fields.["operator" + index] <> "notEquals" then
                                invalidArg "value" "Enter a value for the condition."
                            createObj [ "operator" ==> fields.["operator" + index]; "value" ==> value ]
                        let first = one "1"
                        let second = if fields.["value2"].Trim() = "" then [| first |] else [| first; one "2" |]
                        createObj [ "conditions" ==> second; "and" ==> (fields.["and"] = "true") ]
                    | "top" ->
                        let count =
                            match System.Int32.TryParse(fields.["count"].Trim()) with
                            | true, value when value >= 1 && (fields.["percent"] <> "true" || value <= 100) -> value
                            | _ -> invalidArg "count" "Enter a positive whole count (1–100 for percentages)."
                        createObj [ "top" ==> createObj [
                            "count" ==> count; "percent" ==> (fields.["percent"] = "true")
                            "bottom" ==> (fields.["bottom"] = "true") ] ]
                    | "color" ->
                        let color = fields.["color"].Trim()
                        if color <> "" && not (System.Text.RegularExpressions.Regex.IsMatch(color, "^#[0-9a-fA-F]{6}$")) then
                            invalidArg "color" "Enter #RRGGBB or leave empty for no fill/automatic text color."
                        createObj [ "color" ==> (if color = "" then null else box color); "by" ==> fields.["by"] ]
                    | "dynamic" -> createObj [ "dynamic" ==> fields.["dynamic"] ]
                    | _ -> invalidArg "kind" "Choose a supported filter type."
                // The grid validates range, column and criteria, then records
                // the combined range/criteria change as one undo step.
                grid.SetAutoFilterAndFilter(box target, box col, criteria)
                true
            with error -> dialogError (box error)
        let operators = [
            "equals", "Equals"; "notEquals", "Does not equal"; "greaterThan", "Greater than"
            "greaterOrEqual", "Greater than or equal"; "lessThan", "Less than"
            "lessOrEqual", "Less than or equal"; "contains", "Contains"; "notContains", "Does not contain"
            "beginsWith", "Begins with"; "endsWith", "Ends with" ]
        let op index =
            let value = textOf (condition index "operator")
            if value = "" then "equals" else value
        let currentColor = textOf (property current "color")
        {
            TestId = "xl-dlg-data-filter"; Title = "Filter"
            Subtitle = Some "First row is the header. Filters on other columns remain unless you change the range."
            Fields = [
                dialogText "range" "Filter range" "A1:D20 (header in first row)" range
                dialogText "column" "Column to filter (letters)" "A" (columnToLetter activeCol)
                select "kind" "Criteria" [
                    yield! (if kind = "kept" then [ "kept", "Imported filter (read-only)" ] else [])
                    yield "values", "Visible values"; yield "conditions", "Conditions"
                    yield "top", "Top or bottom"; yield "color", "Color"; yield "dynamic", "Average" ] kind
                dialogTextArea "values" "Visible values (one per line)" "One value per line" 3 values
                dialogCheckbox "blanks" "Include blanks" (flagOf current "blanks" false)
                select "operator1" "First condition" operators (op 0)
                dialogText "value1" "First value" "Text, number or date" (textOf (condition 0 "value"))
                select "operator2" "Second condition (optional)" operators (op 1)
                dialogText "value2" "Second value (empty: one condition)" "Text, number or date" (textOf (condition 1 "value"))
                dialogCheckbox "and" "Both conditions must match (AND)" (flagOf current "and" true)
                dialogText "count" "Top/bottom count" "1" (if isNullish top then "10" else textOf (property top "count"))
                dialogCheckbox "percent" "Count is a percentage" (flagOf top "percent" false)
                dialogCheckbox "bottom" "Bottom instead of top" (flagOf top "bottom" false)
                select "by" "Color source" [ "cellColor", "Cell fill"; "fontColor", "Text color" ]
                    (if textOf (property current "by") = "fontColor" then "fontColor" else "cellColor")
                dialogText "color" "Color (#RRGGBB; empty: none)" "#RRGGBB" currentColor
                select "dynamic" "Average criterion" [ "aboveAverage", "Above average"; "belowAverage", "Below average" ]
                    (if textOf (property current "dynamic") = "belowAverage" then "belowAverage" else "aboveAverage")
                dialogStatus "error" "Validation" "" ]
            Tabs = []
            Buttons = [
                dialogCancel
                dialogButton "xl-dlg-data-filter-clear" "Clear column" false (fun fields ->
                    try
                        dataReady epoch serial sheetIndex
                        let column = fields.["column"].Trim()
                        grid.ClearFilter(column = box column)
                        true
                    with error -> dialogError (box error))
                dialogButton "xl-dlg-data-filter-apply" "Apply" true submit ]
        }

    let quickData label action =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                try action () with error -> setFileMessage (label + " failed: " + printError (box error))
            else setFileMessage (label + " blocked: correct the cell entry first.")

    let openFile () =
        if live () then
            try
                fileInput |> Option.iter (fun input ->
                    setValue input ""
                    clickElement input)
            with error -> setFileMessage ("Open failed: " + printError (box error))

    let onFilePicked () =
        let file = fileInput |> Option.map selectedFile |> Option.defaultValue null
        if live () && not (isNull file) then
            try
                if not (unsaved ()) || confirmReplace "Discard unsaved changes and open another workbook?" then
                    openTicket <- openTicket + 1
                    let ticket = openTicket
                    let epoch = grid.GetWorkbookEpoch()
                    let serial = grid.GetChangeSerial()
                    let editAtPick = editSerial
                    let mutable invalidEntry = false
                    let importedName = unbox<string> (getProp file "name")
                    setFileMessage ("Opening " + importedName + "...")
                    let finish message =
                        if ticket = openTicket && not disposed then
                            setFileMessage message
                            refreshFileState ()
                    grid.ImportXlsxIfCurrent(file, System.Func<bool>(fun () ->
                        if not (live () && ticket = openTicket
                                && grid.GetWorkbookEpoch() = epoch && grid.GetChangeSerial() = serial
                                && editSerial = editAtPick) then false
                        elif grid.IsEditing() && not (grid.TryCommitEdit()) then
                            invalidEntry <- true
                            false
                        else true))
                        .``then``(fun installed ->
                            if installed && live () && ticket = openTicket then
                                fileName <- importedName
                                finish ("Opened " + importedName)
                            else finish (if invalidEntry then "Open blocked: correct the cell entry first." else "Open cancelled: workbook changed."))
                        .``catch``(fun error -> finish ("Open failed: " + printError (box error)))
                    |> ignore
            with error ->
                setFileMessage ("Open failed: " + printError (box error))

    let saveFile () =
        if live () && not saving then
            if not (grid.TryCommitEdit()) then
                setFileMessage "Save blocked: correct the cell entry first."
                refreshFileState ()
            else
                saving <- true
                let version = generation
                let epoch = grid.GetWorkbookEpoch()
                let revision = grid.GetRevision()
                let name = fileName
                setFileMessage ("Preparing " + name + "...")
                let finish message =
                    if not disposed && version = generation then
                        saving <- false
                        if live () && grid.GetWorkbookEpoch() = epoch then
                            setFileMessage message
                            refreshFileState ()
                try
                    grid.ExportXlsx()
                        .``then``(fun bytes ->
                            if live () && version = generation && grid.GetWorkbookEpoch() = epoch then
                                downloadXlsx bytes name
                                // The browser accepted a download request; later edits still differ.
                                if grid.GetWorkbookEpoch() = epoch then grid.MarkSaved(revision)
                                finish ("Download requested: " + name)
                            else finish "")
                        .``catch``(fun error -> finish ("Save failed: " + printError (box error)))
                    |> ignore
                with error -> finish ("Save failed: " + printError (box error))

    let onKeyDown (event: obj) =
        if live () && dialog.Open.Peek().IsNone && ctrlOrMeta event && not (unavailableKey event) && not (keyAlt event) then
            let target = getEventTarget event
            let inShell = rootElement |> Option.exists (fun root -> contains root target)
            let inGrid = grid.OwnsKeyboardTarget(target)
            let key = (getEventKey event).ToLowerInvariant()
            if (inShell || (inGrid && not (editingOtherContent target))) && fileShortcut event
               && (key = "o" || key = "s") then
                preventDefault event
                stopEvent event
                commands.Run (if key = "o" then "file.open" else "file.save") |> ignore
            elif (inShell || inGrid) && not (isTextControl target || editingOtherContent target)
                 && not (grid.IsEditing()) then
                let chord = chordOf key true (keyShift event) false
                match commands.ForChord chord with
                | Some command when command.Id = "edit.find" || command.Id = "format.cells"
                                   || command.Id = "format.bold" || command.Id = "format.italic"
                                   || command.Id = "format.underline" ->
                    preventDefault event
                    stopEvent event
                    commands.Run command.Id |> ignore
                | _ when key = "h" && not (keyShift event) ->
                    preventDefault event
                    stopEvent event
                    grid.ShowFindReplaceDialog()
                | _ -> ()

    let onBeforeUnload (event: obj) =
        if live () && unsaved () then
            preventDefault event
            setProp event "returnValue" (box "")


    let setPrintStatus version message =
        if not disposed && version = generation then
            match statusElement with
            | Some status when isConnected status -> setTextContent status message
            | _ -> ()

    let printPdf () =
        if not disposed && not printing && (rootElement |> Option.exists isConnected) then
            printing <- true
            let version = generation
            setPrintStatus version "Preparing PDF for print..."
            let finish message =
                if version = generation then
                    printing <- false
                    setPrintStatus version message
            try
                grid.PrintPdf()
                    .``then``(fun () -> finish "Print dialog requested.")
                    .``catch``(fun error -> finish ("Print PDF failed: " + printError (box error)))
                |> ignore
            with error ->
                finish ("Print PDF failed: " + printError (box error))



    // A sheet operation captures a stable sheet ID, not the tab index (which
    // can change while a dialog is open). A rejected edit stays in the grid.
    let sheetInfoById id =
        grid.GetSheetInfos()
        |> Array.tryFind (fun info -> textOf (property info "id") = id)

    let sheetReady epoch serial =
        if not (live ()) || grid.IsReadOnly() then invalidOp "The workbook is read-only or the shell was closed."
        if not (grid.TryCommitEdit()) then invalidOp "Correct the cell entry before applying."
        if grid.GetWorkbookEpoch() <> epoch || grid.GetChangeSerial() <> serial then
            invalidOp "The workbook changed; reopen the sheet menu or dialog."

    let sheetIndex id =
        match sheetInfoById id with
        | Some info -> unbox<int> (property info "index")
        | None -> invalidOp "This sheet no longer exists; reopen the sheet menu."

    let focusSheetById id =
        deferFocus (fun () ->
            if live () then
                match sheetTabButtons.TryGetValue id with
                | true, button -> focus button
                | _ ->
                    let active = grid.GetSheetInfos().[grid.GetActiveSheetIndex()]
                    match sheetTabButtons.TryGetValue (textOf (property active "id")) with
                    | true, button -> focus button
                    | _ -> ())

    let sheetOperation label focusActive epoch serial id action =
        try
            sheetReady epoch serial
            let index = sheetIndex id
            action index
            if focusActive then
                let active = grid.GetSheetInfos().[grid.GetActiveSheetIndex()]
                focusSheetById (textOf (property active "id"))
            else focusSheetById id
        with error -> setFileMessage (label + " failed: " + printError (box error))

    let openSheetNameDialog rename id =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                let epoch, serial = grid.GetWorkbookEpoch(), grid.GetChangeSerial()
                let current = if rename then sheetInfoById id else None
                if rename && current.IsNone then setFileMessage "Rename failed: this sheet no longer exists."
                else
                    let oldName = current |> Option.map (fun info -> textOf (property info "name")) |> Option.defaultValue ""
                    let submit (fields: Map<string, string>) =
                        try
                            sheetReady epoch serial
                            let name = fields.["name"].Trim()
                            if rename then
                                let index = sheetIndex id
                                if name = "" then invalidArg "name" "Enter a sheet name."
                                if name <> oldName then grid.RenameSheet(box index, box name)
                            else
                                grid.AddSheet((if name = "" then null else box name), null) |> ignore
                            let focusId =
                                if rename then id
                                else
                                    let info = grid.GetSheetInfos().[grid.GetActiveSheetIndex()]
                                    textOf (property info "id")
                            focusSheetById focusId
                            true
                        with error -> dialogError (box error)
                    openDialog dialog {
                        TestId = if rename then "xl-dlg-rename-sheet" else "xl-dlg-insert-sheet"
                        Title = if rename then "Rename worksheet" else "New worksheet"
                        Subtitle = Some (if rename then "Change this sheet's name. References and defined names follow it."
                                         else "Leave the name blank for the next available Sheet name.")
                        Fields = [ dialogText "name" "Sheet name" "Sheet1" oldName; dialogStatus "error" "Validation" "" ]
                        Tabs = []
                        Buttons = [ dialogCancel
                                    dialogButton (if rename then "xl-dlg-rename-sheet-apply" else "xl-dlg-insert-sheet-apply")
                                        (if rename then "Rename" else "Add") true submit ] }
            else setFileMessage "Sheet operation blocked: correct the cell entry first."

    let openUnhideDialog () =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                let epoch, serial = grid.GetWorkbookEpoch(), grid.GetChangeSerial()
                let hidden =
                    grid.GetSheetInfos()
                    |> Array.filter (fun info -> textOf (property info "visibility") = "hidden")
                if hidden.Length > 0 then
                    let entries = hidden |> Array.map (fun info -> textOf (property info "id"), textOf (property info "name")) |> Array.toList
                    let submit (fields: Map<string, string>) =
                        try
                            sheetReady epoch serial
                            let id = fields.["sheet"]
                            let index = sheetIndex id
                            let info = grid.GetSheetInfos().[index]
                            if textOf (property info "visibility") <> "hidden" then invalidOp "The selected sheet is no longer hidden."
                            grid.SetSheetVisibility(box index, box "visible")
                            focusSheetById id
                            true
                        with error -> dialogError (box error)
                    openDialog dialog {
                        TestId = "xl-dlg-unhide-sheet"; Title = "Unhide worksheet"
                        Subtitle = Some "Sheets marked veryHidden by the workbook cannot be unhidden here."
                        Fields = [ select "sheet" "Hidden sheet" entries (fst entries.Head)
                                   dialogStatus "error" "Validation" "" ]
                        Tabs = []
                        Buttons = [ dialogCancel; dialogButton "xl-dlg-unhide-sheet-apply" "Unhide" true submit ] }
            else setFileMessage "Unhide blocked: correct the cell entry first."

    let axisSelection rows =
        let selection = grid.GetSelectionRef()
        let area = parseRange "Selected rows or columns" selection
        let model = grid.GetModel()
        let first = if rows then area.From.Row else area.From.Col
        let last = min (if rows then area.To.Row else area.To.Col)
                       ((if rows then model.RowCount else model.ColCount) - 1)
        if first > last then invalidOp "The selection is outside this sheet."
        selection, first, last, model

    let axisReady epoch serial sheet selection =
        dataReady epoch serial sheet
        if grid.GetSelectionRef() <> selection then
            invalidOp "The selected rows or columns changed; reopen the size dialog."

    let openAxisDialog rows =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                let kind = if rows then "Row height" else "Column width"
                try
                    let epoch, serial, sheet = grid.GetWorkbookEpoch(), grid.GetChangeSerial(), grid.GetActiveSheetIndex()
                    let selection, first, last, model = axisSelection rows
                    let initialSize = axisSize rows first model
                    let mutable mixed = false
                    let mutable i = first + 1
                    while i <= last && not mixed do
                        mixed <- axisSize rows i model <> initialSize
                        i <- i + 1
                    let initial = if mixed then "" else string initialSize
                    let maximum = if rows then maxRowHeightPx else maxColumnWidthPx
                    let submit (fields: Map<string, string>) =
                        try
                            axisReady epoch serial sheet selection
                            let value = fields.["size"].Trim()
                            // Blank means "leave unchanged", including a mixed selection.
                            if value <> "" then
                                match System.Double.TryParse value with
                                | true, size when not (System.Double.IsNaN size || System.Double.IsInfinity size)
                                                  && size >= 0.0 && size <= maximum ->
                                    if rows then grid.SetRowHeight(box selection, box size)
                                    else grid.SetColumnWidth(box selection, box size)
                                | _ -> invalidArg "size" (sprintf "Enter a size from 0 to %g pixels." maximum)
                            true
                        with error -> dialogError (box error)
                    openDialog dialog {
                        TestId = if rows then "xl-dlg-row-height" else "xl-dlg-column-width"
                        Title = kind
                        Subtitle = Some (sprintf "%s. Sizes are sheet pixels at 100%% zoom (0 hides). %s"
                            selection (if mixed then "Mixed sizes: leave blank to keep each one." else "Leave blank to keep the current size."))
                        Fields = [ dialogNumber "size" (kind + " (px)") (Some "0") (Some (string maximum)) initial
                                   dialogStatus "error" "Validation" "" ]
                        Tabs = []
                        Buttons = [ dialogCancel
                                    dialogButton (if rows then "xl-dlg-row-height-apply" else "xl-dlg-column-width-apply")
                                        "Apply" true submit ] }
                with error -> setFileMessage (kind + " failed: " + printError (box error))
            else setFileMessage "Sizing blocked: correct the cell entry first."

    let quickAxis rows action label =
        if live () && not (grid.IsReadOnly()) then
            if grid.TryCommitEdit() then
                try
                    let target, _, _, _ = axisSelection rows
                    match action with
                    | "autofit" -> if rows then grid.AutoFitRows(range = box target) else grid.AutoFitColumns(range = box target)
                    | "hide" -> if rows then grid.HideRows(box target) else grid.HideColumns(box target)
                    | _ -> if rows then grid.UnhideRows(box target) else grid.UnhideColumns(box target)
                with error -> setFileMessage (label + " failed: " + printError (box error))
            else setFileMessage (label + " blocked: correct the cell entry first.")

    let applySheetAction action id epoch serial =
        if action = "rename" then
            try
                sheetReady epoch serial
                sheetIndex id |> ignore
                openSheetNameDialog true id
            with error -> setFileMessage ("Rename failed: " + printError (box error))
        else
            sheetOperation ("Sheet " + action) (action = "duplicate") epoch serial id (fun index ->
                match action with
                | "duplicate" -> grid.DuplicateSheet(box index) |> ignore
                | "delete" ->
                    let name = sheetInfoById id |> Option.map (fun info -> textOf (property info "name")) |> Option.defaultValue "this sheet"
                    if confirmReplace (sprintf "Delete worksheet \"%s\"? This removes its contents and references." name) then
                        grid.DeleteSheet(box index)
                | "moveLeft" when index > 0 -> grid.MoveSheet(box index, box (index - 1))
                | "moveRight" when index < grid.GetSheetInfos().Length - 1 -> grid.MoveSheet(box index, box (index + 1))
                | "hide" ->
                    if sheetInfoById id |> Option.exists (fun info -> textOf (property info "visibility") = "visible") then
                        grid.SetSheetVisibility(box index, box "hidden")
                | _ -> ())

    let registerCommands () =
        let enabled () = live () && not (grid.IsReadOnly()) && not (grid.IsEditing())
        let apply name _ = grid.ApplyFormat name
        let add id label shortcut run =
            commands.Add (command id label shortcut enabled run)
        let toggleFormat id label shortcut propertyName name =
            commands.Add (toggle id label shortcut enabled (checkedFlag propertyName) (apply name))
        commands.Add(command "file.open" "Open" (Some "Ctrl+O") live (fun _ -> openFile()))
        commands.Add(command "file.save" "Save" (Some "Ctrl+S") live (fun _ -> saveFile()))
        commands.Add(command "file.print" "Print PDF" None live (fun _ -> printPdf()))
        commands.Add(command "edit.undo" "Undo" None (fun () -> grid.CanUndo() && enabled()) (fun _ -> grid.Undo()))
        commands.Add(command "edit.redo" "Redo" None (fun () -> grid.CanRedo() && enabled()) (fun _ -> grid.Redo()))
        commands.Add(command "edit.find" "Find" (Some "Ctrl+F") live (fun _ -> grid.ToggleFindReplace()))
        commands.Add(command "edit.findNext" "Find next" None live (fun _ -> grid.FindNextInBar()))
        commands.Add(command "edit.replaceAll" "Replace all" None (fun () -> live () && not (grid.IsReadOnly())) (fun _ -> grid.ReplaceAllInBar()))
        commands.Add(command "insert.function" "Insert function" None
            (fun () -> live () && not (grid.IsReadOnly()) && not (isNullish (grid.GetEditSnapshot())))
            (fun _ -> openFunctionChooser ()))
        commands.Add(command "insert.chart" "Insert chart" None
            (fun () -> live () && not (grid.IsReadOnly()) && grid.GetSheetKind() = "worksheet")
            (fun _ -> openDataDialog (fun _ _ _ -> chartDialog ())))
        commands.Add(command "format.cells" "Format cells" (Some "Ctrl+1") (fun () -> live () && not (grid.IsReadOnly())) (fun _ -> openFormatDialog()))
        let dataEnabled () = live () && not (grid.IsReadOnly())
        commands.Add(command "data.validation" "Data validation" None dataEnabled (fun _ ->
            openDataDialog validationDialog))
        commands.Add(command "data.sort" "Custom sort" None dataEnabled (fun _ ->
            openDataDialog sortDialog))
        commands.Add(command "data.filter" "Filter" None dataEnabled (fun _ ->
            openDataDialog filterDialog))
        commands.Add(command "data.conditionalFormatting" "Conditional formatting" None dataEnabled (fun _ ->
            quickData "Conditional formatting" (fun () -> grid.ShowConditionalFormattingDialog())))
        commands.Add(command "data.names" "Name manager" None dataEnabled (fun _ ->
            quickData "Name manager" (fun () -> grid.ShowDefinedNamesDialog())))
        // The VBA editor: the workbook's modules (a blank Module1 when it has none), its macros, an
        // Immediate line and their output. Running from it enables the workbook's macros.
        commands.Add(command "tools.vba" "VBA editor" None dataEnabled (fun _ ->
            let attempt (action: unit -> unit) =
                try action () with error -> setFileMessage ("VBA: " + printError (box error))
            let mutable reopen : string -> unit = ignore
            let mutable refresh : unit -> unit = ignore
            // (Macros load before the editor opens: its actions then run at once and the output follows.)
            let enabled (action: unit -> JS.Promise<obj>) =
                attempt (fun () ->
                    let run () = (action ()).``then``(fun _ -> refresh ()) |> ignore
                    if grid.MacrosEnabled then run ()
                    else (grid.EnableMacros()).``then``(fun _ -> run ()) |> ignore)
            let host : Rbt.Spreadsheet.Ui.MacroEditorHost = {
                Modules = fun () ->
                    match grid.GetMacroModules() |> Array.map (fun m -> textOf (property m "name"), textOf (property m "kind"), textOf (property m "code")) |> List.ofArray with
                    | [] -> [ "Module1", "standard", "" ]
                    | modules -> modules
                Macros = fun () -> grid.GetMacros() |> List.ofArray
                Run = fun name -> if name <> "" then enabled (fun () -> grid.RunMacro name)
                Immediate = fun line -> if line.Trim() <> "" then enabled (fun () -> grid.RunMacroLine line)
                Apply = fun name code ->
                    // (Shown again when compiled: the macro list follows the module's new code.)
                    attempt (fun () -> (grid.SetMacroModule(name, code)).``then``(fun _ -> reopen name) |> ignore)
                Output = fun () ->
                    let lines = grid.GetMacroOutput()
                    String.concat "\n" (if lines.Length > 200 then lines.[lines.Length - 200 ..] else lines) }
            reopen <- fun name -> openDialog dialog (Rbt.Spreadsheet.Ui.macroEditor host (Some name))
            refresh <- fun () ->
                // (An action that settled after its button returned: the output follows on a fresh view.)
                match dialog.Open.Peek() with
                | Some opened when opened.Spec.TestId = "xl-dlg-vba" ->
                    match opened.Values.TryGetValue "__revision" with
                    | true, revision -> revision.Value <- string (int revision.Value + 1)
                    | _ -> ()
                | _ -> ()
            (grid.LoadMacros()).``then``(fun () -> openDialog dialog (Rbt.Spreadsheet.Ui.macroEditor host None)) |> ignore))
        let sheetEnabled () = live () && not (grid.IsReadOnly())
        let targetId argument =
            let infos = grid.GetSheetInfos()
            let index =
                if isNullish argument then grid.GetActiveSheetIndex()
                elif isIntegerArg argument then unbox<int> argument
                else invalidArg "index" "Expected a 0-based sheet index."
            if index < 0 || index >= infos.Length then invalidArg "index" "The sheet index is outside this workbook."
            textOf (property infos.[index] "id")
        commands.Add(command "sheet.add" "New sheet" None sheetEnabled (fun _ -> openSheetNameDialog false ""))
        commands.Add(command "sheet.unhide" "Unhide sheet" None
            (fun () -> sheetEnabled () && (grid.GetSheetInfos() |> Array.exists (fun info -> textOf (property info "visibility") = "hidden")))
            (fun _ -> openUnhideDialog ()))
        for action, label in [
            "rename", "Rename sheet"; "duplicate", "Duplicate sheet"; "delete", "Delete sheet"
            "moveLeft", "Move sheet left"; "moveRight", "Move sheet right"; "hide", "Hide sheet" ] do
            commands.Add(command ("sheet." + action) label None sheetEnabled (fun argument ->
                let id = targetId argument
                applySheetAction action id (grid.GetWorkbookEpoch()) (grid.GetChangeSerial())))
        let sizingEnabled () = sheetEnabled () && grid.GetSheetKind() = "worksheet"
        for rows, axis in [ true, "row"; false, "column" ] do
            commands.Add(command ("axis." + axis + "Size") (if rows then "Row height" else "Column width")
                None sizingEnabled (fun _ -> openAxisDialog rows))
            for action, label in [
                "autofit", "AutoFit"; "hide", "Hide"; "unhide", "Unhide" ] do
                commands.Add(command ("axis." + axis + "." + action) (label + " " + axis + "s")
                    None sizingEnabled (fun _ -> quickAxis rows action (label + " " + axis + "s")))
        commands.Add(command "data.sortAsc" "Sort A to Z" None dataEnabled (fun _ ->
            quickData "Sort" (fun () -> grid.SortCurrentColumnAsc())))
        commands.Add(command "data.sortDesc" "Sort Z to A" None dataEnabled (fun _ ->
            quickData "Sort" (fun () -> grid.SortCurrentColumnDesc())))
        commands.Add(command "data.clearFilter" "Clear filters" None
            (fun () -> dataEnabled () && not (isNullish (grid.GetFilters())))
            (fun _ -> quickData "Clear filters" (fun () -> grid.ClearFilter())))
        toggleFormat "format.bold" "Bold" (Some "Ctrl+B") "bold" "bold"
        toggleFormat "format.italic" "Italic" (Some "Ctrl+I") "italic" "italic"
        toggleFormat "format.underline" "Underline" (Some "Ctrl+U") "underline" "underline"
        for id, label, action in [
            "format.fontIncrease", "Increase font size", "fontIncrease"
            "format.fontDecrease", "Decrease font size", "fontDecrease"
            "format.decimalIncrease", "Increase decimals", "decimalIncrease"
            "format.decimalDecrease", "Decrease decimals", "decimalDecrease" ] do
            add id label None (apply action)
        for id, label, alignment, action in [
            "format.alignLeft", "Align left", "left", "alignLeft"
            "format.alignCenter", "Align center", "center", "alignCenter"
            "format.alignRight", "Align right", "right", "alignRight" ] do
            commands.Add(toggle id label None enabled
                (fun () -> match formatText "horizontalAlign" with
                           | "mixed" -> None
                           | value -> Some (value = alignment)) (apply action))
        for id, label, code in [
            "format.number", "Number", "0.00"
            "format.currency", "Currency", "$#,##0.00"
            "format.percent", "Percentage", "0%" ] do
            commands.Add(toggle id label None enabled
                (fun () -> match formatText "numberFormat" with
                           | "mixed" -> None
                           | value -> Some (value = code))
                (fun _ -> grid.ApplyFormatPatch(createObj [ "numberFormat" ==> code ])))
        for id, label, propertyName, color in [
            "format.textColor", "Text color", "textColor", "#000000"
            "format.bgColor", "Fill color", "fill", "#FFFF00" ] do
            add id label None (fun argument ->
                let value = if isNullish argument then color else unbox<string> argument
                grid.ApplyFormatPatch(createObj [ propertyName ==> value ]))
        add "format.fontFamily" "Font family" None (fun argument ->
            grid.ApplyFormatPatch(createObj [ "fontFamily" ==> argument ]))
        add "format.fontSize" "Font size" None (fun argument ->
            grid.ApplyFormatPatch(createObj [ "fontSize" ==> argument ]))
    let run id () = commands.Run id |> ignore


    let menus () = [
        {
            Id = "file"
            Label = "File"
            Rows = [
                MItem { Id = "open"; Label = "Open..."; Shortcut = Some "Ctrl+O"; Action = run "file.open" }
                MItem { Id = "save"; Label = "Save"; Shortcut = Some "Ctrl+S"; Action = run "file.save" }
                MSep
                MItem { Id = "print-pdf"; Label = "Print PDF"; Shortcut = None; Action = run "file.print" }
                MItem { Id = "new-sheet"; Label = "New worksheet..."; Shortcut = None; Action = run "sheet.add" }
            ]
        }
        {
            Id = "edit"
            Label = "Edit"
            Rows = [
                MItem { Id = "undo"; Label = "Undo"; Shortcut = Some "Ctrl+Z"; Action = run "edit.undo" }
                MItem { Id = "redo"; Label = "Redo"; Shortcut = Some "Ctrl+Y"; Action = run "edit.redo" }
                MSep
                MItem { Id = "cut"; Label = "Cut"; Shortcut = Some "Ctrl+X"; Action = fun () -> onAction "cut" }
                MItem { Id = "copy"; Label = "Copy"; Shortcut = Some "Ctrl+C"; Action = fun () -> onAction "copy" }
                MItem { Id = "paste"; Label = "Paste"; Shortcut = Some "Ctrl+V"; Action = fun () -> onAction "paste" }
                MSep
                MItem { Id = "find-replace"; Label = "Find and replace..."; Shortcut = Some "Ctrl+F"; Action = run "edit.find" }
            ]
        }
        {
            Id = "insert"
            Label = "Insert"
            Rows = [
                MItem { Id = "insert-row"; Label = "Insert row"; Shortcut = None; Action = fun () -> onAction "insertRow" }
                MItem { Id = "insert-column"; Label = "Insert column"; Shortcut = None; Action = fun () -> onAction "insertColumn" }
                MItem { Id = "insert-function"; Label = "Function..."; Shortcut = None; Action = run "insert.function" }
                MItem { Id = "insert-chart"; Label = "Chart..."; Shortcut = None; Action = run "insert.chart" }
                MSep
                MItem { Id = "delete-row"; Label = "Delete row"; Shortcut = None; Action = fun () -> onAction "deleteRow" }
                MItem { Id = "delete-column"; Label = "Delete column"; Shortcut = None; Action = fun () -> onAction "deleteColumn" }
            ]
        }
        {
            Id = "format"
            Label = "Format"
            Rows = [
                MItem { Id = "bold"; Label = "Bold"; Shortcut = Some "Ctrl+B"; Action = run "format.bold" }
                MItem { Id = "italic"; Label = "Italic"; Shortcut = Some "Ctrl+I"; Action = run "format.italic" }
                MItem { Id = "underline"; Label = "Underline"; Shortcut = Some "Ctrl+U"; Action = run "format.underline" }
                MSep
                MItem { Id = "format-cells"; Label = "Format cells..."; Shortcut = Some "Ctrl+1"; Action = run "format.cells" }
                MItem { Id = "format-number"; Label = "Number"; Shortcut = None; Action = run "format.number" }
                MItem { Id = "format-currency"; Label = "Currency"; Shortcut = None; Action = run "format.currency" }
                MItem { Id = "format-percent"; Label = "Percentage"; Shortcut = None; Action = run "format.percent" }
                MSep
                MItem { Id = "align-left"; Label = "Align left"; Shortcut = None; Action = run "format.alignLeft" }
                MItem { Id = "align-center"; Label = "Align center"; Shortcut = None; Action = run "format.alignCenter" }
                MItem { Id = "align-right"; Label = "Align right"; Shortcut = None; Action = run "format.alignRight" }
                MSep
                MItem { Id = "row-height"; Label = "Row height..."; Shortcut = None; Action = run "axis.rowSize" }
                MItem { Id = "column-width"; Label = "Column width..."; Shortcut = None; Action = run "axis.columnSize" }
                MItem { Id = "autofit-rows"; Label = "AutoFit selected rows"; Shortcut = None; Action = run "axis.row.autofit" }
                MItem { Id = "autofit-columns"; Label = "AutoFit selected columns"; Shortcut = None; Action = run "axis.column.autofit" }
                MItem { Id = "hide-rows"; Label = "Hide selected rows"; Shortcut = None; Action = run "axis.row.hide" }
                MItem { Id = "unhide-rows"; Label = "Unhide selected rows"; Shortcut = None; Action = run "axis.row.unhide" }
                MItem { Id = "hide-columns"; Label = "Hide selected columns"; Shortcut = None; Action = run "axis.column.hide" }
                MItem { Id = "unhide-columns"; Label = "Unhide selected columns"; Shortcut = None; Action = run "axis.column.unhide" }
            ]
        }
        {
            Id = "data"
            Label = "Data"
            Rows = [
                MItem { Id = "sort-asc"; Label = "Sort A to Z"; Shortcut = None; Action = run "data.sortAsc" }
                MItem { Id = "sort-desc"; Label = "Sort Z to A"; Shortcut = None; Action = run "data.sortDesc" }
                MItem { Id = "sort-custom"; Label = "Custom sort..."; Shortcut = None; Action = run "data.sort" }
                MSep
                MItem { Id = "filter"; Label = "Filter..."; Shortcut = None; Action = run "data.filter" }
                MItem { Id = "clear-filter"; Label = "Clear filters"; Shortcut = None; Action = run "data.clearFilter" }
                MSep
                MItem { Id = "validation"; Label = "Data validation..."; Shortcut = None; Action = run "data.validation" }
                MItem { Id = "conditional-formatting"; Label = "Conditional formatting..."; Shortcut = None; Action = run "data.conditionalFormatting" }
                MItem { Id = "name-manager"; Label = "Name manager..."; Shortcut = None; Action = run "data.names" }
                MSep
                MItem { Id = "vba-editor"; Label = "VBA editor..."; Shortcut = None; Action = run "tools.vba" }
            ]
        }
    ]

    // Single, focused toolbar row. Niche actions (insert/delete row/col,
    // sort, filter, find) live in the menus.
    let toolbarItems () = [
        TButton ("insert-function", "fx", run "insert.function")
        TSep
        TButton ("bold", "B", run "format.bold")
        TButton ("italic", "I", run "format.italic")
        TButton ("underline", "U", run "format.underline")
        TSep
        TButton ("font-increase", "A+", run "format.fontIncrease")
        TButton ("font-decrease", "A-", run "format.fontDecrease")
        TSep
        TButton ("text-color", "A", fun () -> ())
        TButton ("bg-color", "Fill", fun () -> ())
        TSep
        TButton ("align-left", "L", run "format.alignLeft")
        TButton ("align-center", "C", run "format.alignCenter")
        TButton ("align-right", "R", run "format.alignRight")
        TSep
        TButton ("format-currency", "$", run "format.currency")
        TButton ("format-percent", "%", run "format.percent")
        TButton ("decimal-increase", ".0+", run "format.decimalIncrease")
        TButton ("decimal-decrease", ".0-", run "format.decimalDecrease")
        TSep
        TButton ("cut", "Cut", fun () -> onAction "cut")
        TButton ("copy", "Copy", fun () -> onAction "copy")
        TButton ("paste", "Paste", fun () -> onAction "paste")
        TSep
        TButton ("undo", "Undo", run "edit.undo")
        TButton ("redo", "Redo", run "edit.redo")
    ]

    let buildMenubar () =
        let bar = createElement "div"
        setClassName bar "sheet-menubar"
        setAttribute bar "role" "group"
        setAttribute bar "aria-label" "Spreadsheet menus"
        for menu in menus () do
            let rendered = renderMenu menu
            appendChild bar rendered
            for oldId, commandId in [
                "undo", "edit.undo"; "redo", "edit.redo"; "bold", "format.bold"
                "italic", "format.italic"; "underline", "format.underline"
                "align-left", "format.alignLeft"; "align-center", "format.alignCenter"
                "align-right", "format.alignRight"; "format-cells", "format.cells"
                "format-number", "format.number"; "format-currency", "format.currency"
                "format-percent", "format.percent"
                "new-sheet", "sheet.add"
                "insert-function", "insert.function"
                "insert-chart", "insert.chart"
                "row-height", "axis.rowSize"; "column-width", "axis.columnSize"
                "autofit-rows", "axis.row.autofit"; "autofit-columns", "axis.column.autofit"
                "hide-rows", "axis.row.hide"; "unhide-rows", "axis.row.unhide"
                "hide-columns", "axis.column.hide"; "unhide-columns", "axis.column.unhide"
                "sort-asc", "data.sortAsc"; "sort-desc", "data.sortDesc"; "sort-custom", "data.sort"
                "filter", "data.filter"; "clear-filter", "data.clearFilter"
                "validation", "data.validation"
                "conditional-formatting", "data.conditionalFormatting"
                "name-manager", "data.names" ] do
                querySelector rendered (sprintf "[data-sheet-menu-item='%s']" oldId)
                |> Option.iter (fun button -> buttons.Add(commandId, button))
        bar

    let buildToolbar () =
        let bar = createElement "div"
        setClassName bar "sheet-toolbar"
        setAttribute bar "role" "toolbar"
        setAttribute bar "aria-label" "Spreadsheet formatting and editing"
        for key, commandId, label in [
            "fontFamily", "format.fontFamily", "Font family"
            "fontSize", "format.fontSize", "Font size (pt)" ] do
            let input = createElement "input"
            setAttribute input "type" (if key = "fontSize" then "number" else "text")
            setAttribute input "aria-label" label
            setAttribute input "placeholder" label
            setAttribute input "data-sheet-format" key
            setStyle input "width" (if key = "fontSize" then "68px" else "120px")
            addEventListener input "change" (fun _ ->
                let value = (getValue input).Trim()
                if key = "fontFamily" then
                    commands.Run(commandId, argument = (if value = "" then null else box value)) |> ignore
                else
                    match System.Double.TryParse value with
                    | true, points -> commands.Run(commandId, argument = box points) |> ignore
                    | _ when value = "" -> commands.Run(commandId, argument = null) |> ignore
                    | _ -> setFileMessage "Enter a valid font size."
                refreshCommands())
            appendChild bar input
            inputs.Add(key, input)
        appendChild bar (toolbarSeparator ())
        for item in toolbarItems () do
            match item with
            | TButton (id, label, action) ->
                let accessibleName =
                    match id with
                    | "bold" -> "Bold"
                    | "italic" -> "Italic"
                    | "underline" -> "Underline"
                    | "insert-function" -> "Insert function"
                    | "font-increase" -> "Increase font size"
                    | "font-decrease" -> "Decrease font size"
                    | "text-color" -> "Choose text color"
                    | "bg-color" -> "Choose fill color"
                    | "align-left" -> "Align left"
                    | "align-center" -> "Align center"
                    | "align-right" -> "Align right"
                    | "format-currency" -> "Currency format"
                    | "format-percent" -> "Percentage format"
                    | "decimal-increase" -> "Increase decimal places"
                    | "decimal-decrease" -> "Decrease decimal places"
                    | _ -> label
                let button = toolbarButton id label accessibleName action
                appendChild bar button
                let commandId =
                    match id with
                    | "bold" | "italic" | "underline" -> "format." + id
                    | "undo" | "redo" -> "edit." + id
                    | "insert-function" -> "insert.function"
                    | "font-increase" -> "format.fontIncrease"
                    | "font-decrease" -> "format.fontDecrease"
                    | "text-color" -> "format.textColor"
                    | "bg-color" -> "format.bgColor"
                    | "align-left" -> "format.alignLeft"
                    | "align-center" -> "format.alignCenter"
                    | "align-right" -> "format.alignRight"
                    | "format-currency" -> "format.currency"
                    | "format-percent" -> "format.percent"
                    | "decimal-increase" -> "format.decimalIncrease"
                    | "decimal-decrease" -> "format.decimalDecrease"
                    | _ -> ""
                if commandId <> "" then buttons.Add(commandId, button)
                if id = "text-color" || id = "bg-color" then
                    let colorKey = if id = "text-color" then "textColor" else "fill"
                    let picker = createElement "input"
                    setAttribute picker "type" "color"
                    setAttribute picker "aria-label" (if id = "text-color" then "Text color" else "Fill color")
                    setAttribute picker "data-sheet-color" colorKey
                    setClassName picker "sheet-color-picker"
                    setAttribute picker "tabindex" "-1"
                    addEventListener picker "change" (fun _ ->
                        commands.Run(commandId, argument = box (getValue picker)) |> ignore)
                    addEventListener button "click" (fun _ ->
                        let current = formatText colorKey
                        setValue picker (if current.StartsWith("#") && current.Length = 7 then current else "#000000")
                        clickElement picker)
                    appendChild bar picker
            | TSep -> appendChild bar (toolbarSeparator ())
        bar


    let mutable lastTabSheet = -1

    let popupButton popup id label disabled action =
        let button = createElement "button"
        setAttribute button "type" "button"
        setAttribute button "data-sheet-tab-action" id
        setAttribute button "role" "menuitem"
        setTextContent button label
        setProp button "disabled" (box disabled)
        addEventListener button "click" (fun e ->
            stopPropagation e
            let trigger = sheetMenuTrigger
            closeSheetMenu ()
            trigger |> Option.iter focus
            action ())
        appendChild popup button

    let showSheetPopup id (trigger: Element) epoch notifyHost x y =
        try
            if epoch <> grid.GetWorkbookEpoch() then invalidOp "The workbook changed; use its current tabs."
            let serial = grid.GetChangeSerial()
            let index = sheetIndex id
            closeSheetMenu ()
            if notifyHost && grid.RaiseSheetTabContextMenu(box index, box x, box y) then ()
            else
                match sheetPopup with
                | Some popup ->
                    closeOpenMenu ()
                    setInnerHTML popup ""
                    let editable = not (grid.IsReadOnly())
                    let visible = grid.GetSheetInfos() |> Array.filter (fun item -> textOf (property item "visibility") = "visible")
                    let lastVisible = visible.Length <= 1
                    let item action label disabled =
                        popupButton popup action label (not editable || disabled)
                            (fun () -> applySheetAction action id epoch serial)
                    popupButton popup "add" "New worksheet..." (not editable)
                        (fun () -> commands.Run "sheet.add" |> ignore)
                    item "rename" "Rename..." false
                    item "duplicate" "Duplicate" false
                    item "delete" "Delete..." lastVisible
                    item "moveLeft" "Move left" (index = 0)
                    item "moveRight" "Move right" (index = grid.GetSheetInfos().Length - 1)
                    item "hide" "Hide" lastVisible
                    popupButton popup "unhide" "Unhide..." (not editable ||
                        not (grid.GetSheetInfos() |> Array.exists (fun item -> textOf (property item "visibility") = "hidden")))
                        (fun () -> openUnhideDialog ())
                    setStyle popup "display" "block"
                    positionSheetMenu popup trigger
                    setAttribute trigger "aria-expanded" "true"
                    sheetMenuTrigger <- Some trigger
                    sheetMenuPanel <- Some popup
                    querySelector popup "button:not([disabled])" |> Option.iter focus
                | None -> ()
        with error -> setFileMessage ("Sheet menu failed: " + printError (box error))

    let refreshSheetTabs () =
        match sheetTabsElement with
        | Some bar when live () ->
            closeSheetMenu ()
            match querySelector bar ".sheet-tab-list" with
            | None -> ()
            | Some list ->
                setInnerHTML list ""
                sheetTabButtons.Clear()
                let infos = grid.GetSheetInfos()
                lastTabSheet <- grid.GetActiveSheetIndex()
                let visible = infos |> Array.filter (fun info -> textOf (property info "visibility") = "visible")
                for position = 0 to visible.Length - 1 do
                    let info = visible.[position]
                    let id = textOf (property info "id")
                    let name = textOf (property info "name")
                    let index = unbox<int> (property info "index")
                    let active = index = lastTabSheet
                    let tab = createElement "button"
                    setClassName tab "sheet-tab"
                    setAttribute tab "type" "button"
                    setAttribute tab "role" "tab"
                    setAttribute tab "data-sheet-tab-id" id
                    setAttribute tab "data-active" (if active then "true" else "false")
                    setAttribute tab "aria-selected" (if active then "true" else "false")
                    setAttribute tab "aria-label" ("Switch to " + name)
                    setTextContent tab name
                    let color = textOf (property info "tabColor")
                    if color.StartsWith "#" then setStyle tab "border-top" ("3px solid " + color)
                    let epoch = grid.GetWorkbookEpoch()
                    addEventListener tab "click" (fun _ ->
                        try
                            if epoch <> grid.GetWorkbookEpoch() then invalidOp "The workbook changed; use its current tabs."
                            if sheetIndex id <> grid.GetActiveSheetIndex() then
                                if not (grid.TryCommitEdit()) then invalidOp "Correct the cell entry before switching sheets."
                                let current = sheetIndex id
                                let currentInfo = grid.GetSheetInfos().[current]
                                if textOf (property currentInfo "visibility") <> "visible" then invalidOp "This sheet is no longer visible."
                                grid.SwitchSheet current
                            match sheetTabButtons.TryGetValue id with
                            | true, button -> focus button
                            | _ -> ()
                        with error -> setFileMessage ("Switch sheet failed: " + printError (box error)))
                    addEventListener tab "keydown" (fun e ->
                        let key = getEventKey e
                        if sheetMenuKey e then
                            preventDefault e
                            stopEvent e
                            showSheetPopup id tab epoch true (centerX tab) (centerY tab)
                        elif key = "ArrowLeft" || key = "ArrowRight" || key = "Home" || key = "End" then
                            preventDefault e
                            let next =
                                match key with
                                | "Home" -> 0
                                | "End" -> visible.Length - 1
                                | "ArrowLeft" -> max 0 (position - 1)
                                | _ -> min (visible.Length - 1) (position + 1)
                            let nextId = textOf (property visible.[next] "id")
                            match sheetTabButtons.TryGetValue nextId with
                            | true, button -> focus button
                            | _ -> ())
                    addEventListener tab "contextmenu" (fun e ->
                        preventDefault e
                        showSheetPopup id tab epoch true (asNumber (getProp e "clientX")) (asNumber (getProp e "clientY")))
                    appendChild list tab
                    sheetTabButtons.[id] <- tab
                    let more = createElement "button"
                    setClassName more "sheet-tab-action"
                    setAttribute more "type" "button"
                    setAttribute more "data-sheet-tab-menu-trigger" id
                    setAttribute more "aria-label" ("Options for " + name)
                    setAttribute more "aria-haspopup" "menu"
                    setAttribute more "aria-expanded" "false"
                    setTextContent more "⋯"
                    addEventListener more "click" (fun e ->
                        stopPropagation e
                        showSheetPopup id more epoch false 0.0 0.0)
                    addEventListener more "keydown" (fun e ->
                        if sheetMenuKey e then
                            preventDefault e
                            stopEvent e
                            showSheetPopup id more epoch false 0.0 0.0)
                    appendChild list more
        | _ -> ()

    let buildSheetTabs () =
        let bar = createElement "div"
        setClassName bar "sheet-tab-bar"
        setAttribute bar "aria-label" "Worksheets"
        let add = createElement "button"
        setClassName add "sheet-tab-action"
        setAttribute add "type" "button"
        setAttribute add "data-sheet-tab-add" "true"
        setAttribute add "aria-label" "Add worksheet"
        setTextContent add "+"
        addEventListener add "click" (fun _ -> commands.Run "sheet.add" |> ignore)
        appendChild bar add
        buttons.Add("sheet.add", add)
        let list = createElement "div"
        setClassName list "sheet-tab-list"
        setAttribute list "role" "tablist"
        setAttribute list "aria-label" "Visible worksheets"
        appendChild bar list
        let unhide = createElement "button"
        setClassName unhide "sheet-tab-action"
        setAttribute unhide "type" "button"
        setAttribute unhide "data-sheet-tab-unhide" "true"
        setTextContent unhide "Unhide"
        addEventListener unhide "click" (fun _ -> commands.Run "sheet.unhide" |> ignore)
        appendChild bar unhide
        buttons.Add("sheet.unhide", unhide)
        sheetTabsElement <- Some bar
        bar

    let buildSheetPopup () =
        let popup = createElement "div"
        setClassName popup "sheet-tab-menu"
        setAttribute popup "data-sheet-tab-menu" "true"
        setAttribute popup "role" "menu"
        setAttribute popup "aria-label" "Worksheet actions"
        setStyle popup "display" "none"
        addEventListener popup "click" stopEvent
        addEventListener popup "keydown" (fun e ->
            let key = getEventKey e
            if key = "Escape" then
                preventDefault e
                stopEvent e
                let trigger = sheetMenuTrigger
                closeSheetMenu ()
                trigger |> Option.iter focus
            elif key = "ArrowUp" || key = "ArrowDown" || key = "Home" || key = "End" then
                preventDefault e
                let available =
                    getChildren popup
                    |> Array.filter (fun button -> not (unbox<bool> (getProp button "disabled")))
                if available.Length > 0 then
                    let current = getProp (getDocument ()) "activeElement"
                    let index = available |> Array.tryFindIndex (fun button -> obj.ReferenceEquals(box button, current))
                    let next =
                        match key with
                        | "Home" -> 0
                        | "End" -> available.Length - 1
                        | "ArrowUp" -> max 0 ((Option.defaultValue 0 index) - 1)
                        | _ -> min (available.Length - 1) ((Option.defaultValue -1 index) + 1)
                    focus available.[next])
        sheetPopup <- Some popup
        popup

    let build () =
        let shell = createElement "div"
        setClassName shell "rb-office-shell sheet-shell"
        let name = createElement "span"
        setClassName name "sheet-file-name"
        setAttribute name "aria-live" "polite"
        setTitle name "Workbook filename and unsaved status"
        appendChild shell name
        nameElement <- Some name
        let input = createElement "input"
        setAttribute input "type" "file"
        setAttribute input "accept" ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        setAttribute input "aria-label" "Open workbook"
        setAttribute input "data-sheet-file-input" "true"
        setStyle input "display" "none"
        addEventListener input "change" (fun _ -> onFilePicked ())
        appendChild shell input
        fileInput <- Some input
        appendChild shell (buildMenubar ())
        appendChild shell (buildToolbar ())
        grid.AttachFindReplace(shell)
        appendChild shell (buildSheetTabs ())
        appendChild shell (buildSheetPopup ())
        let selectionStatus = createElement "div"
        setClassName selectionStatus "sheet-selection-status"
        setAttribute selectionStatus "role" "status"
        setAttribute selectionStatus "aria-live" "polite"
        setAttribute selectionStatus "aria-label" "Selection statistics"
        selectionStats <-
            [| for key, _ in selectionStatusFields do
                   let field = createElement "span"
                   setAttribute field "data-sheet-selection-stat" key
                   appendChild selectionStatus field
                   yield field |]
        appendChild shell selectionStatus
        let status = createElement "div"
        setClassName status "sheet-print-status"
        setAttribute status "role" "status"
        setAttribute status "aria-live" "polite"
        setAttribute status "aria-label" "Spreadsheet status"
        appendChild shell status
        statusElement <- Some status
        shell

    let detach () =
        rootElement |> Option.iter (fun shell -> grid.DetachFindReplace(shell))
        closeFunctionChooser false
        openTicket <- openTicket + 1
        saving <- false
        selectionSubscription |> Option.iter (fun unsubscribe -> unsubscribe.Invoke())
        selectionSubscription <- None
        sheetsSubscription |> Option.iter (fun unsubscribe -> unsubscribe.Invoke())
        sheetsSubscription <- None
        sheetTabsElement <- None
        sheetPopup <- None
        sheetTabButtons.Clear()
        lastTabSheet <- -1
        unmountDialog |> Option.iter (fun unmount -> unmount())
        unmountDialog <- None
        closeDialog dialog
        buttons.Clear()
        inputs.Clear()
        fileSubscription |> Option.iter (fun unsubscribe -> unsubscribe.Invoke())
        fileSubscription <- None
        editSubscription |> Option.iter (fun unsubscribe -> unsubscribe.Invoke())
        editSubscription <- None
        readOnlySubscription |> Option.iter (fun unsubscribe -> unsubscribe.Invoke())
        readOnlySubscription <- None
        keyListener |> Option.iter (removeCaptureKey (getDocument ()))
        keyListener <- None
        unloadListener |> Option.iter (removeEventListener (browserWindow ()) "beforeunload")
        unloadListener <- None
        rootElement |> Option.iter (fun root ->
            releaseDocCloseHandler root
            let parent = getProp root "parentNode"
            if isNotNull parent then removeChild parent root)
        rootElement <- None
        statusElement <- None
        selectionStats <- [||]
        statusSnapshot <- null
        nameElement <- None
        fileInput <- None

    member _.Mount(container: Element) =
        generation <- generation + 1
        printing <- false
        detach ()
        disposed <- false
        registerCommands ()
        installStyles "spreadsheet" shellStyles
        installDocCloseHandler ()
        let shell = build ()
        appendChild container shell
        rootElement <- Some shell
        unmountDialog <- Some (Rbt.FableUI.Dom.mount (unbox shell) (renderDialog dialog))
        refreshSheetTabs ()
        refreshFileState ()
        fileSubscription <- Some (grid.OnChange(System.Action<obj>(fun _ ->
            refreshFileState ()
            refreshCommands ()
            refreshSelectionStatus ())))
        selectionSubscription <- Some (grid.OnSelectionChange(System.Action<obj>(fun _ ->
            refreshCommands ()
            refreshSelectionStatus ()
            if grid.GetActiveSheetIndex() <> lastTabSheet then refreshSheetTabs ())))
        sheetsSubscription <- Some (grid.OnSheetsChange(System.Action<obj>(fun _ ->
            refreshSheetTabs ()
            refreshSelectionStatus ()
            refreshCommands ())))
        editSubscription <- Some (grid.OnEditStateChange(System.Action<obj>(fun _ ->
            editSerial <- editSerial + 1
            refreshFileState ()
            refreshCommands ())))
        readOnlySubscription <- Some (grid.OnReadOnlyChange(System.Action<bool>(fun _ ->
            closeSheetMenu ()
            closeFunctionChooser false
            refreshCommands ())))
        let key: EventListener = onKeyDown
        let unload: EventListener = onBeforeUnload
        addCaptureKey (getDocument ()) key
        addEventListener (browserWindow ()) "beforeunload" unload
        keyListener <- Some key
        unloadListener <- Some unload
        refreshCommands ()
        refreshSelectionStatus ()

    member _.Dispose() =
        disposed <- true
        generation <- generation + 1
        printing <- false
        detach ()

    member _.GetElement() = rootElement
    member _.Commands = jsCommands

    member _.ShowFindReplace() =
        if live () then grid.ShowFindReplaceDialog()

// ── Formula bar (host callbacks retained; local argument help) ──

[<AttachMembers>]
type FormulaBar
        (onCellRefChange: string -> unit,
         onFormulaChange: string -> unit,
         onFormulaSubmit: unit -> unit) =
    let mutable formulaBarElement: Element option = None
    let mutable cellRefInput: Element option = None
    let mutable formulaInput: Element option = None
    let mutable argumentHelp: Element option = None
    let updateArgumentHelp () =
        match formulaInput, argumentHelp with
        | Some input, Some help ->
            let caret = unbox<int> (getProp input "selectionStart")
            match formulaHelpAt (getValue input) caret with
            | Some (info, index) ->
                setTextContent help (info.Name + " — " + functionArgumentHelp info index + " · " + info.Description)
                setStyle help "display" ""
            | None -> setStyle help "display" "none"
        | _ -> ()

    let createFormulaBar () =
        let container = createElement "div"
        setClassName container "rb-office-shell sheet-formula-bar"
        let cellRefLabel = createElement "span"

        setStyle cellRefLabel "color" "var(--office-muted)"
        setTextContent cellRefLabel "Cell"
        appendChild container cellRefLabel

        let cellRef = createElement "input"
        setClassName cellRef "cell-ref-input"
        setAttribute cellRef "type" "text"
        setAttribute cellRef "readonly" "true"
        setAttribute cellRef "value" "A1"
        setAttribute cellRef "aria-label" "Active cell reference"
        appendChild container cellRef
        cellRefInput <- Some cellRef

        let fxLabel = createElement "span"
        setStyle fxLabel "color" "var(--office-muted)"
        setTextContent fxLabel "fx"
        appendChild container fxLabel

        let formula = createElement "input"
        setClassName formula "formula-input"
        setAttribute formula "type" "text"
        setAttribute formula "placeholder" "Enter value or formula"
        setAttribute formula "aria-label" "Cell value or formula"
        addEventListener formula "input" (fun e ->
            let target = getEventTarget e
            let value = getProp target "value"
            onFormulaChange (unbox value)
            updateArgumentHelp ())
        addEventListener formula "keydown" (fun e ->
            let key = getProp e "key"
            if key = box "Enter" then
                preventDefault e
                onFormulaSubmit ())
        addEventListener formula "keyup" (fun _ -> updateArgumentHelp ())
        addEventListener formula "mouseup" (fun _ -> updateArgumentHelp ())
        appendChild container formula
        formulaInput <- Some formula
        let help = createElement "div"
        setClassName help "sheet-formula-argument-help"
        setAttribute help "role" "status"
        setAttribute help "aria-live" "polite"
        setStyle help "display" "none"
        appendChild container help
        argumentHelp <- Some help

        container

    member _.Mount(container: Element) =
        installStyles "spreadsheet" shellStyles
        let formulaBar = createFormulaBar ()
        appendChild container formulaBar
        formulaBarElement <- Some formulaBar

    member _.SetCellRef(ref: string) =
        match cellRefInput with
        | Some input -> setProp input "value" (box ref)
        | None -> ()

    member _.SetFormula(formula: string) =
        match formulaInput with
        | Some input ->
            setProp input "value" (box formula)
            updateArgumentHelp ()
        | None -> ()

    member _.GetElement() = formulaBarElement

// ── Exports ─────────────────────────────────────────────────

[<ExportDefault>]
let createSpreadsheetShell
        (grid: Rbt.Spreadsheet.Spreadsheet.Spreadsheet)
        (onAction: string -> unit) : SpreadsheetShell =
    new SpreadsheetShell(grid, onAction)

let createFormulaBar
        (onCellRefChange: string -> unit)
        (onFormulaChange: string -> unit)
        (onFormulaSubmit: unit -> unit) : FormulaBar =
    new FormulaBar(onCellRefChange, onFormulaChange, onFormulaSubmit)

#endif
