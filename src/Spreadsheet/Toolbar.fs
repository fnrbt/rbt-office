module RedBat.Spreadsheet.Toolbar

open Fable.Core
open Fable.Core.JsInterop
open RedBat.Grid.Dom

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

// ── Flat chrome palette ─────────────────────────────────────
// Mirrors the Word web shell's Google-Docs-flat look: white surfaces,
// hairline borders, soft hover/pressed tints. Variables are also exposed
// as CSS custom properties on the root so downstream stylesheets can
// override individual tokens without forking the F#.

module private Palette =
    let chromeBg = "#ffffff"
    let hover = "#f1f3f4"
    let pressed = "#e8f0fe"
    let border = "#dadce0"
    let separator = "#e1e5ea"
    let text = "#202124"
    let mutedText = "#5f6368"
    let menuPanelShadow = "0 1px 3px rgba(60,64,67,0.16), 0 4px 8px rgba(60,64,67,0.08)"

// ── Document-level menu state ───────────────────────────────
// Only one menu panel can be open at a time; clicking outside closes it.

let mutable private openTrigger: Element option = None
let mutable private openPanel: Element option = None
let mutable private docHandlerInstalled = false

let private closeOpenMenu () =
    match openPanel with
    | Some panel -> setStyle panel "display" "none"
    | None -> ()
    match openTrigger with
    | Some trigger ->
        setStyle trigger "background" "transparent"
        setAttribute trigger "aria-expanded" "false"
    | None -> ()
    openPanel <- None
    openTrigger <- None

let private installDocCloseHandler () =
    if not docHandlerInstalled then
        let doc = getDocument ()
        addEventListener doc "click" (fun _ -> closeOpenMenu ())
        docHandlerInstalled <- true

// ── Toolbar primitives ──────────────────────────────────────

let private toolbarButton id label action =
    let btn = createElement "button"
    setClassName btn "sheet-tool-button"
    setAttribute btn "data-sheet-icon" id
    setAttribute btn "type" "button"
    setStyles btn [
        "padding", "0 8px"
        "border", "1px solid transparent"
        "background", "transparent"
        "font-family", "Arial, sans-serif"
        "font-size", "11px"
        "color", Palette.text
        "cursor", "pointer"
        "border-radius", "3px"
        "user-select", "none"
        "min-width", "24px"
        "height", "24px"
        "margin", "0 1px"
    ]
    setTextContent btn label
    addEventListener btn "mouseenter" (fun _ -> setStyle btn "background" Palette.hover)
    addEventListener btn "mouseleave" (fun _ -> setStyle btn "background" "transparent")
    addEventListener btn "mousedown" (fun e -> preventDefault e)
    addEventListener btn "click" (fun _ -> action ())
    btn

let private toolbarSeparator () =
    let sep = createElement "div"
    setStyles sep [
        "width", "1px"
        "height", "16px"
        "background", Palette.separator
        "margin", "0 6px"
        "flex", "0 0 auto"
    ]
    sep

// ── Menu rendering ──────────────────────────────────────────

let private renderMenuItem (item: MenuItem) (onClose: unit -> unit) =
    let row = createElement "button"
    setClassName row "sheet-menu-item"
    setAttribute row "data-sheet-menu-item" item.Id
    setAttribute row "type" "button"
    setStyles row [
        "display", "flex"
        "align-items", "center"
        "justify-content", "space-between"
        "width", "100%"
        "padding", "6px 16px"
        "border", "none"
        "background", "transparent"
        "font-family", "Arial, sans-serif"
        "font-size", "12px"
        "color", Palette.text
        "cursor", "pointer"
        "text-align", "left"
        "white-space", "nowrap"
    ]
    let labelSpan = createElement "span"
    setTextContent labelSpan item.Label
    appendChild row labelSpan
    match item.Shortcut with
    | Some shortcut ->
        let sh = createElement "span"
        setStyles sh [
            "color", Palette.mutedText
            "margin-left", "32px"
            "font-size", "11px"
        ]
        setTextContent sh shortcut
        appendChild row sh
    | None -> ()
    addEventListener row "mouseenter" (fun _ -> setStyle row "background" Palette.hover)
    addEventListener row "mouseleave" (fun _ -> setStyle row "background" "transparent")
    addEventListener row "click" (fun _ ->
        onClose ()
        item.Action ())
    row

let private renderMenuSeparator () =
    let sep = createElement "div"
    setStyles sep [
        "height", "1px"
        "background", Palette.separator
        "margin", "4px 0"
    ]
    sep

let private renderMenu (menu: Menu) =
    let wrap = createElement "div"
    setStyles wrap [
        "position", "relative"
        "display", "inline-block"
    ]

    let trigger = createElement "button"
    setClassName trigger "sheet-menu-button"
    setAttribute trigger "data-sheet-menu" menu.Id
    setAttribute trigger "type" "button"
    setAttribute trigger "aria-haspopup" "true"
    setAttribute trigger "aria-expanded" "false"
    setStyles trigger [
        "padding", "0 12px"
        "border", "1px solid transparent"
        "background", "transparent"
        "font-family", "Arial, sans-serif"
        "font-size", "13px"
        "color", Palette.text
        "cursor", "pointer"
        "border-radius", "3px"
        "user-select", "none"
        "height", "26px"
    ]
    setTextContent trigger menu.Label

    let panel = createElement "div"
    setClassName panel "sheet-menu-panel"
    setAttribute panel "data-sheet-menu-panel" menu.Id
    setStyles panel [
        "position", "absolute"
        "top", "100%"
        "left", "0"
        "min-width", "220px"
        "background", Palette.chromeBg
        "border", sprintf "1px solid %s" Palette.border
        "border-radius", "4px"
        "box-shadow", Palette.menuPanelShadow
        "padding", "4px 0"
        "z-index", "100"
        "display", "none"
    ]
    // Clicks inside the panel must not bubble to the document-level handler
    // that closes the menu. The item-click handler closes it explicitly.
    addEventListener panel "click" (fun e -> stopPropagation e)

    for row in menu.Rows do
        match row with
        | MItem item -> appendChild panel (renderMenuItem item closeOpenMenu)
        | MSep -> appendChild panel (renderMenuSeparator ())

    addEventListener trigger "mouseenter" (fun _ ->
        if openTrigger <> Some trigger then
            setStyle trigger "background" Palette.hover)
    addEventListener trigger "mouseleave" (fun _ ->
        if openTrigger <> Some trigger then
            setStyle trigger "background" "transparent")
    addEventListener trigger "mousedown" (fun e -> preventDefault e)
    addEventListener trigger "click" (fun e ->
        stopPropagation e
        let wasOpen = openTrigger = Some trigger
        closeOpenMenu ()
        if not wasOpen then
            setStyle panel "display" "block"
            setStyle trigger "background" Palette.pressed
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
// that renders a menubar, a single focused toolbar row, and an
// inline find/replace bar that toggles via Edit > Find.

[<AttachMembers>]
type SpreadsheetShell
        (onAction: string -> unit,
         onFind: string -> unit,
         onReplace: string * string -> unit) =

    let mutable rootElement: Element option = None
    let mutable findBar: Element option = None
    let mutable findInputEl: Element option = None
    let mutable replaceInputEl: Element option = None

    let toggleFindBar () =
        match findBar with
        | Some bar ->
            let isHidden = getAttribute bar "data-hidden" = "true"
            if isHidden then
                setStyle bar "display" "flex"
                setAttribute bar "data-hidden" "false"
                findInputEl |> Option.iter (fun el -> (unbox<obj> el)?focus())
            else
                setStyle bar "display" "none"
                setAttribute bar "data-hidden" "true"
        | None -> ()

    let menus () = [
        {
            Id = "file"
            Label = "File"
            Rows = [
                MItem { Id = "open"; Label = "Open..."; Shortcut = Some "Ctrl+O"; Action = fun () -> onAction "open" }
                MItem { Id = "save"; Label = "Save"; Shortcut = Some "Ctrl+S"; Action = fun () -> onAction "save" }
                MSep
                MItem { Id = "print-pdf"; Label = "Print PDF"; Shortcut = None; Action = fun () -> onAction "printPdf" }
            ]
        }
        {
            Id = "edit"
            Label = "Edit"
            Rows = [
                MItem { Id = "undo"; Label = "Undo"; Shortcut = Some "Ctrl+Z"; Action = fun () -> onAction "undo" }
                MItem { Id = "redo"; Label = "Redo"; Shortcut = Some "Ctrl+Y"; Action = fun () -> onAction "redo" }
                MSep
                MItem { Id = "cut"; Label = "Cut"; Shortcut = Some "Ctrl+X"; Action = fun () -> onAction "cut" }
                MItem { Id = "copy"; Label = "Copy"; Shortcut = Some "Ctrl+C"; Action = fun () -> onAction "copy" }
                MItem { Id = "paste"; Label = "Paste"; Shortcut = Some "Ctrl+V"; Action = fun () -> onAction "paste" }
                MSep
                MItem { Id = "find-replace"; Label = "Find and replace..."; Shortcut = Some "Ctrl+F"; Action = toggleFindBar }
            ]
        }
        {
            Id = "insert"
            Label = "Insert"
            Rows = [
                MItem { Id = "insert-row"; Label = "Insert row"; Shortcut = None; Action = fun () -> onAction "insertRow" }
                MItem { Id = "insert-column"; Label = "Insert column"; Shortcut = None; Action = fun () -> onAction "insertColumn" }
                MSep
                MItem { Id = "delete-row"; Label = "Delete row"; Shortcut = None; Action = fun () -> onAction "deleteRow" }
                MItem { Id = "delete-column"; Label = "Delete column"; Shortcut = None; Action = fun () -> onAction "deleteColumn" }
            ]
        }
        {
            Id = "format"
            Label = "Format"
            Rows = [
                MItem { Id = "bold"; Label = "Bold"; Shortcut = Some "Ctrl+B"; Action = fun () -> onAction "bold" }
                MItem { Id = "italic"; Label = "Italic"; Shortcut = Some "Ctrl+I"; Action = fun () -> onAction "italic" }
                MItem { Id = "underline"; Label = "Underline"; Shortcut = Some "Ctrl+U"; Action = fun () -> onAction "underline" }
                MSep
                MItem { Id = "format-number"; Label = "Number"; Shortcut = None; Action = fun () -> onAction "formatNumber" }
                MItem { Id = "format-currency"; Label = "Currency"; Shortcut = None; Action = fun () -> onAction "formatCurrency" }
                MItem { Id = "format-percent"; Label = "Percentage"; Shortcut = None; Action = fun () -> onAction "formatPercent" }
                MSep
                MItem { Id = "align-left"; Label = "Align left"; Shortcut = None; Action = fun () -> onAction "alignLeft" }
                MItem { Id = "align-center"; Label = "Align center"; Shortcut = None; Action = fun () -> onAction "alignCenter" }
                MItem { Id = "align-right"; Label = "Align right"; Shortcut = None; Action = fun () -> onAction "alignRight" }
            ]
        }
        {
            Id = "data"
            Label = "Data"
            Rows = [
                MItem { Id = "sort-asc"; Label = "Sort A to Z"; Shortcut = None; Action = fun () -> onAction "sortAsc" }
                MItem { Id = "sort-desc"; Label = "Sort Z to A"; Shortcut = None; Action = fun () -> onAction "sortDesc" }
                MSep
                MItem { Id = "filter"; Label = "Filter"; Shortcut = None; Action = fun () -> onAction "filter" }
                MItem { Id = "clear-filter"; Label = "Clear filters"; Shortcut = None; Action = fun () -> onAction "clearFilter" }
            ]
        }
    ]

    // Single, focused toolbar row. Niche actions (insert/delete row/col,
    // sort, filter, find) live in the menus.
    let toolbarItems () = [
        TButton ("bold", "B", fun () -> onAction "bold")
        TButton ("italic", "I", fun () -> onAction "italic")
        TButton ("underline", "U", fun () -> onAction "underline")
        TSep
        TButton ("font-increase", "A+", fun () -> onAction "fontIncrease")
        TButton ("font-decrease", "A-", fun () -> onAction "fontDecrease")
        TSep
        TButton ("text-color", "A", fun () -> onAction "textColor")
        TButton ("bg-color", "Fill", fun () -> onAction "bgColor")
        TSep
        TButton ("align-left", "L", fun () -> onAction "alignLeft")
        TButton ("align-center", "C", fun () -> onAction "alignCenter")
        TButton ("align-right", "R", fun () -> onAction "alignRight")
        TSep
        TButton ("format-currency", "$", fun () -> onAction "formatCurrency")
        TButton ("format-percent", "%", fun () -> onAction "formatPercent")
        TButton ("decimal-increase", ".0+", fun () -> onAction "decimalIncrease")
        TButton ("decimal-decrease", ".0-", fun () -> onAction "decimalDecrease")
        TSep
        TButton ("cut", "Cut", fun () -> onAction "cut")
        TButton ("copy", "Copy", fun () -> onAction "copy")
        TButton ("paste", "Paste", fun () -> onAction "paste")
        TSep
        TButton ("undo", "Undo", fun () -> onAction "undo")
        TButton ("redo", "Redo", fun () -> onAction "redo")
    ]

    let buildMenubar () =
        let bar = createElement "div"
        setClassName bar "sheet-menubar"
        setStyles bar [
            "display", "flex"
            "align-items", "center"
            "padding", "2px 6px"
            "background", Palette.chromeBg
            "border-bottom", sprintf "1px solid %s" Palette.separator
            "gap", "2px"
        ]
        for menu in menus () do appendChild bar (renderMenu menu)
        bar

    let buildToolbar () =
        let bar = createElement "div"
        setClassName bar "sheet-toolbar"
        setStyles bar [
            "display", "flex"
            "align-items", "center"
            "padding", "4px 6px"
            "background", Palette.chromeBg
            "border-bottom", sprintf "1px solid %s" Palette.separator
            "user-select", "none"
            "flex-wrap", "wrap"
        ]
        for item in toolbarItems () do
            match item with
            | TButton (id, label, action) -> appendChild bar (toolbarButton id label action)
            | TSep -> appendChild bar (toolbarSeparator ())
        bar

    let buildFindBar () =
        let bar = createElement "div"
        setClassName bar "sheet-find-bar"
        setAttribute bar "data-hidden" "true"
        setStyles bar [
            "display", "none"
            "align-items", "center"
            "padding", "6px 8px"
            "background", Palette.chromeBg
            "border-bottom", sprintf "1px solid %s" Palette.separator
            "gap", "8px"
        ]

        let findLabel = createElement "span"
        setStyles findLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", Palette.mutedText
        ]
        setTextContent findLabel "Find"
        appendChild bar findLabel

        let findInput = createElement "input"
        setAttribute findInput "type" "text"
        setAttribute findInput "placeholder" "Search..."
        setClassName findInput "sheet-find-input"
        setStyles findInput [
            "width", "180px"
            "padding", "4px 8px"
            "border", sprintf "1px solid %s" Palette.border
            "border-radius", "3px"
            "font-family", "Arial, sans-serif"
            "font-size", "12px"
        ]
        appendChild bar findInput
        findInputEl <- Some findInput

        let findBtn = toolbarButton "find-run" "Find" (fun () ->
            findInputEl |> Option.iter (fun el ->
                let value = getValue el
                onFind value))
        appendChild bar findBtn

        let replaceLabel = createElement "span"
        setStyles replaceLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", Palette.mutedText
            "margin-left", "12px"
        ]
        setTextContent replaceLabel "Replace"
        appendChild bar replaceLabel

        let replaceInput = createElement "input"
        setAttribute replaceInput "type" "text"
        setAttribute replaceInput "placeholder" "Replace with..."
        setClassName replaceInput "sheet-replace-input"
        setStyles replaceInput [
            "width", "180px"
            "padding", "4px 8px"
            "border", sprintf "1px solid %s" Palette.border
            "border-radius", "3px"
            "font-family", "Arial, sans-serif"
            "font-size", "12px"
        ]
        appendChild bar replaceInput
        replaceInputEl <- Some replaceInput

        let replaceBtn = toolbarButton "replace-run" "Replace all" (fun () ->
            match findInputEl, replaceInputEl with
            | Some f, Some r -> onReplace (getValue f, getValue r)
            | _ -> ())
        appendChild bar replaceBtn

        let closeBtn = toolbarButton "find-close" "Close" (fun () ->
            setStyle bar "display" "none"
            setAttribute bar "data-hidden" "true")
        // Push the close button to the right edge.
        let spacer = createElement "div"
        setStyle spacer "flex" "1"
        appendChild bar spacer
        appendChild bar closeBtn

        bar

    let build () =
        installDocCloseHandler ()
        let shell = createElement "div"
        setClassName shell "sheet-shell"
        setStyles shell [
            "background", Palette.chromeBg
            "color", Palette.text
            "border-bottom", sprintf "1px solid %s" Palette.border
        ]
        appendChild shell (buildMenubar ())
        appendChild shell (buildToolbar ())
        let fb = buildFindBar ()
        appendChild shell fb
        findBar <- Some fb
        shell

    member _.Mount(container: Element) =
        let shell = build ()
        appendChild container shell
        rootElement <- Some shell

    member _.GetElement() = rootElement

    member _.ShowFindReplace() = toggleFindBar ()

// ── Formula bar (unchanged behavior, flat chrome) ───────────

[<AttachMembers>]
type FormulaBar
        (onCellRefChange: string -> unit,
         onFormulaChange: string -> unit,
         onFormulaSubmit: unit -> unit) =
    let mutable formulaBarElement: Element option = None
    let mutable cellRefInput: Element option = None
    let mutable formulaInput: Element option = None

    let createFormulaBar () =
        let container = createElement "div"
        setClassName container "sheet-formula-bar"
        setStyles container [
            "display", "flex"
            "align-items", "center"
            "padding", "4px 6px"
            "background", Palette.chromeBg
            "border-bottom", sprintf "1px solid %s" Palette.separator
            "gap", "8px"
        ]

        let cellRefLabel = createElement "span"
        setStyles cellRefLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", Palette.mutedText
        ]
        setTextContent cellRefLabel "Cell"
        appendChild container cellRefLabel

        let cellRef = createElement "input"
        setClassName cellRef "cell-ref-input"
        setAttribute cellRef "type" "text"
        setAttribute cellRef "readonly" "true"
        setAttribute cellRef "value" "A1"
        setStyles cellRef [
            "width", "60px"
            "padding", "4px 6px"
            "border", sprintf "1px solid %s" Palette.border
            "border-radius", "3px"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "text-align", "center"
            "background", Palette.chromeBg
        ]
        appendChild container cellRef
        cellRefInput <- Some cellRef

        let fxLabel = createElement "span"
        setStyles fxLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "font-weight", "bold"
            "color", Palette.mutedText
        ]
        setTextContent fxLabel "fx"
        appendChild container fxLabel

        let formula = createElement "input"
        setClassName formula "formula-input"
        setAttribute formula "type" "text"
        setAttribute formula "placeholder" "Enter value or formula"
        setStyles formula [
            "flex", "1"
            "padding", "4px 8px"
            "border", sprintf "1px solid %s" Palette.border
            "border-radius", "3px"
            "font-family", "Consolas, 'Courier New', monospace"
            "font-size", "11px"
        ]
        addEventListener formula "input" (fun e ->
            let target = getEventTarget e
            let value = getProp target "value"
            onFormulaChange (unbox value))
        addEventListener formula "keydown" (fun e ->
            let key = getProp e "key"
            if key = box "Enter" then
                preventDefault e
                onFormulaSubmit ())
        appendChild container formula
        formulaInput <- Some formula

        container

    member _.Mount(container: Element) =
        let formulaBar = createFormulaBar ()
        appendChild container formulaBar
        formulaBarElement <- Some formulaBar

    member _.SetCellRef(ref: string) =
        match cellRefInput with
        | Some input -> setProp input "value" (box ref)
        | None -> ()

    member _.SetFormula(formula: string) =
        match formulaInput with
        | Some input -> setProp input "value" (box formula)
        | None -> ()

    member _.GetElement() = formulaBarElement

// ── Exports ─────────────────────────────────────────────────

[<ExportDefault>]
let createSpreadsheetShell
        (onAction: string -> unit)
        (onFind: string -> unit)
        (onReplace: string * string -> unit) : SpreadsheetShell =
    new SpreadsheetShell(onAction, onFind, onReplace)

let createFormulaBar
        (onCellRefChange: string -> unit)
        (onFormulaChange: string -> unit)
        (onFormulaSubmit: unit -> unit) : FormulaBar =
    new FormulaBar(onCellRefChange, onFormulaChange, onFormulaSubmit)
