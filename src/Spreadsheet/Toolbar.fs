module RedBat.Spreadsheet.Toolbar

open Fable.Core
open Fable.Core.JsInterop
open RedBat.Grid.Dom

/// Toolbar button
type ToolbarButton = {
    Id: string
    Label: string
    Icon: string option
    Action: unit -> unit
}

/// Toolbar separator
type ToolbarItem =
    | Button of ToolbarButton
    | Separator

/// Toolbar configuration
type ToolbarConfig = {
    Items: ToolbarItem list
    OnAction: string -> unit
}

/// Create a toolbar button element
let private createButton (btn: ToolbarButton): Element =
    let button = createElement "button"
    setClassName button "toolbar-button"
    setStyles button [
        "padding", "6px 12px"
        "border", "1px solid #d0d0d0"
        "background", "linear-gradient(180deg, #fefefe 0%, #f0f0f0 100%)"
        "font-family", "Arial, sans-serif"
        "font-size", "11px"
        "cursor", "pointer"
        "margin-right", "2px"
        "border-radius", "2px"
        "color", "#444"
        "user-select", "none"
    ]

    match btn.Icon with
    | Some icon -> setTextContent button icon
    | None -> setTextContent button btn.Label

    addEventListener button "mousedown" (fun e ->
        preventDefault e
    )

    addEventListener button "click" (fun _ ->
        btn.Action()
    )

    // Hover effect
    addEventListener button "mouseenter" (fun _ ->
        setStyle button "background" "linear-gradient(180deg, #ffffff 0%, #e8e8e8 100%)"
        setStyle button "border-color" "#b0b0b0"
    )

    addEventListener button "mouseleave" (fun _ ->
        setStyle button "background" "linear-gradient(180deg, #fefefe 0%, #f0f0f0 100%)"
        setStyle button "border-color" "#d0d0d0"
    )

    button

/// Create a toolbar separator
let private createSeparator (): Element =
    let sep = createElement "div"
    setClassName sep "toolbar-separator"
    setStyles sep [
        "width", "1px"
        "height", "24px"
        "background", "#d0d0d0"
        "margin", "0 4px"
        "display", "inline-block"
        "vertical-align", "middle"
    ]
    sep

/// Create toolbar container
let createToolbar (items: ToolbarItem list): Element =
    let toolbar = createElement "div"
    setClassName toolbar "toolbar"
    setStyles toolbar [
        "display", "flex"
        "align-items", "center"
        "padding", "4px"
        "background", "#f8f8f8"
        "border-bottom", "1px solid #d0d0d0"
        "user-select", "none"
    ]

    for item in items do
        match item with
        | Button btn ->
            let button = createButton btn
            appendChild toolbar button
        | Separator ->
            let sep = createSeparator()
            appendChild toolbar sep

    toolbar

/// Format toolbar for spreadsheet
[<AttachMembers>]
type FormatToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createFormatToolbar () =
        let items = [
            Button {
                Id = "bold"
                Label = "B"
                Icon = Some "B"
                Action = fun () -> onAction "bold"
            }
            Button {
                Id = "italic"
                Label = "I"
                Icon = Some "I"
                Action = fun () -> onAction "italic"
            }
            Button {
                Id = "underline"
                Label = "U"
                Icon = Some "U"
                Action = fun () -> onAction "underline"
            }
            Separator
            Button {
                Id = "font-increase"
                Label = "A+"
                Icon = Some "A+"
                Action = fun () -> onAction "fontIncrease"
            }
            Button {
                Id = "font-decrease"
                Label = "A-"
                Icon = Some "A-"
                Action = fun () -> onAction "fontDecrease"
            }
            Separator
            Button {
                Id = "text-color"
                Label = "Color"
                Icon = Some "🎨"
                Action = fun () -> onAction "textColor"
            }
            Button {
                Id = "bg-color"
                Label = "Fill"
                Icon = Some "⬛"
                Action = fun () -> onAction "bgColor"
            }
        ]

        // Update bold button style
        let toolbar = createToolbar items
        let boldBtn = querySelector toolbar ".toolbar-button"
        if boldBtn.IsSome then
            setStyle boldBtn.Value "font-weight" "bold"

        toolbar

    member _.Mount(container: Element) =
        let toolbar = createFormatToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Alignment toolbar
[<AttachMembers>]
type AlignmentToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createAlignmentToolbar () =
        let items = [
            Button {
                Id = "align-left"
                Label = "Left"
                Icon = Some "⬅"
                Action = fun () -> onAction "alignLeft"
            }
            Button {
                Id = "align-center"
                Label = "Center"
                Icon = Some "↔"
                Action = fun () -> onAction "alignCenter"
            }
            Button {
                Id = "align-right"
                Label = "Right"
                Icon = Some "➡"
                Action = fun () -> onAction "alignRight"
            }
        ]
        createToolbar items

    member _.Mount(container: Element) =
        let toolbar = createAlignmentToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Number format toolbar
[<AttachMembers>]
type NumberFormatToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createNumberToolbar () =
        let items = [
            Button {
                Id = "format-number"
                Label = "1,234.56"
                Icon = None
                Action = fun () -> onAction "formatNumber"
            }
            Button {
                Id = "format-currency"
                Label = "$"
                Icon = Some "$"
                Action = fun () -> onAction "formatCurrency"
            }
            Button {
                Id = "format-percent"
                Label = "%"
                Icon = Some "%"
                Action = fun () -> onAction "formatPercent"
            }
            Separator
            Button {
                Id = "decimal-increase"
                Label = ".0+"
                Icon = None
                Action = fun () -> onAction "decimalIncrease"
            }
            Button {
                Id = "decimal-decrease"
                Label = ".0-"
                Icon = None
                Action = fun () -> onAction "decimalDecrease"
            }
        ]
        createToolbar items

    member _.Mount(container: Element) =
        let toolbar = createNumberToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Formula bar for cell editing
[<AttachMembers>]
type FormulaBar(onCellRefChange: string -> unit, onFormulaChange: string -> unit, onFormulaSubmit: unit -> unit) =
    let mutable formulaBarElement: Element option = None
    let mutable cellRefInput: Element option = None
    let mutable formulaInput: Element option = None

    let createFormulaBar () =
        let container = createElement "div"
        setClassName container "formula-bar"
        setStyles container [
            "display", "flex"
            "align-items", "center"
            "padding", "4px"
            "background", "#ffffff"
            "border-bottom", "1px solid #d0d0d0"
            "gap", "8px"
        ]

        // Cell reference display/input (e.g., "A1")
        let cellRefLabel = createElement "div"
        setStyles cellRefLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", "#666"
            "white-space", "nowrap"
        ]
        setTextContent cellRefLabel "Cell:"
        appendChild container cellRefLabel

        let cellRef = createElement "input"
        setClassName cellRef "cell-ref-input"
        setAttribute cellRef "type" "text"
        setAttribute cellRef "readonly" "true"
        setAttribute cellRef "value" "A1"
        setStyles cellRef [
            "width", "60px"
            "padding", "4px 8px"
            "border", "1px solid #d0d0d0"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "text-align", "center"
            "background", "#f8f8f8"
        ]
        appendChild container cellRef
        cellRefInput <- Some cellRef

        // Formula input (shows/edits cell content)
        let fxLabel = createElement "div"
        setStyles fxLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "font-weight", "bold"
            "color", "#444"
            "white-space", "nowrap"
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
            "border", "1px solid #d0d0d0"
            "font-family", "Consolas, 'Courier New', monospace"
            "font-size", "11px"
        ]

        addEventListener formula "input" (fun e ->
            let target = getEventTarget e
            let value = getProp target "value"
            onFormulaChange (unbox value)
        )

        addEventListener formula "keydown" (fun e ->
            let key = getProp e "key"
            if key = box "Enter" then
                preventDefault e
                onFormulaSubmit()
        )

        appendChild container formula
        formulaInput <- Some formula

        container

    member _.Mount(container: Element) =
        let formulaBar = createFormulaBar()
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

/// Insert/Delete toolbar
[<AttachMembers>]
type InsertDeleteToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createInsertDeleteToolbar () =
        let items = [
            Button {
                Id = "insert-row"
                Label = "Insert Row"
                Icon = Some "⬇+"
                Action = fun () -> onAction "insertRow"
            }
            Button {
                Id = "delete-row"
                Label = "Delete Row"
                Icon = Some "⬇−"
                Action = fun () -> onAction "deleteRow"
            }
            Separator
            Button {
                Id = "insert-column"
                Label = "Insert Column"
                Icon = Some "➡+"
                Action = fun () -> onAction "insertColumn"
            }
            Button {
                Id = "delete-column"
                Label = "Delete Column"
                Icon = Some "➡−"
                Action = fun () -> onAction "deleteColumn"
            }
        ]
        createToolbar items

    member _.Mount(container: Element) =
        let toolbar = createInsertDeleteToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Data operations toolbar (sort, filter)
[<AttachMembers>]
type DataToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createDataToolbar () =
        let items = [
            Button {
                Id = "sort-asc"
                Label = "Sort A→Z"
                Icon = Some "↓A"
                Action = fun () -> onAction "sortAsc"
            }
            Button {
                Id = "sort-desc"
                Label = "Sort Z→A"
                Icon = Some "↓Z"
                Action = fun () -> onAction "sortDesc"
            }
            Separator
            Button {
                Id = "filter"
                Label = "Filter"
                Icon = Some "⚲"
                Action = fun () -> onAction "filter"
            }
            Button {
                Id = "clear-filter"
                Label = "Clear Filters"
                Icon = Some "✕"
                Action = fun () -> onAction "clearFilter"
            }
        ]
        createToolbar items

    member _.Mount(container: Element) =
        let toolbar = createDataToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Find/Replace toolbar
[<AttachMembers>]
type FindReplaceToolbar(onFind: string -> unit, onReplace: string * string -> unit) =
    let mutable toolbarElement: Element option = None
    let mutable findInput: Element option = None
    let mutable replaceInput: Element option = None

    let createFindReplaceToolbar () =
        let container = createElement "div"
        setClassName container "find-replace-toolbar"
        setStyles container [
            "display", "flex"
            "align-items", "center"
            "padding", "4px"
            "background", "#ffffff"
            "border-bottom", "1px solid #d0d0d0"
            "gap", "8px"
        ]

        // Find input
        let findLabel = createElement "div"
        setStyles findLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", "#666"
            "white-space", "nowrap"
        ]
        setTextContent findLabel "Find:"
        appendChild container findLabel

        let find = createElement "input"
        setClassName find "find-input"
        setAttribute find "type" "text"
        setAttribute find "placeholder" "Search..."
        setStyles find [
            "width", "150px"
            "padding", "4px 8px"
            "border", "1px solid #d0d0d0"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
        ]
        appendChild container find
        findInput <- Some find

        // Find button
        let findBtn = createElement "button"
        setClassName findBtn "toolbar-button"
        setTextContent findBtn "Find"
        setStyles findBtn [
            "padding", "4px 12px"
            "border", "1px solid #d0d0d0"
            "background", "linear-gradient(180deg, #fefefe 0%, #f0f0f0 100%)"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "cursor", "pointer"
            "border-radius", "2px"
        ]
        addEventListener findBtn "click" (fun _ ->
            match findInput with
            | Some input ->
                let value = getProp input "value"
                onFind (unbox value)
            | None -> ()
        )
        appendChild container findBtn

        // Replace input
        let replaceLabel = createElement "div"
        setStyles replaceLabel [
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "color", "#666"
            "white-space", "nowrap"
            "margin-left", "16px"
        ]
        setTextContent replaceLabel "Replace:"
        appendChild container replaceLabel

        let replace = createElement "input"
        setClassName replace "replace-input"
        setAttribute replace "type" "text"
        setAttribute replace "placeholder" "Replace with..."
        setStyles replace [
            "width", "150px"
            "padding", "4px 8px"
            "border", "1px solid #d0d0d0"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
        ]
        appendChild container replace
        replaceInput <- Some replace

        // Replace button
        let replaceBtn = createElement "button"
        setClassName replaceBtn "toolbar-button"
        setTextContent replaceBtn "Replace"
        setStyles replaceBtn [
            "padding", "4px 12px"
            "border", "1px solid #d0d0d0"
            "background", "linear-gradient(180deg, #fefefe 0%, #f0f0f0 100%)"
            "font-family", "Arial, sans-serif"
            "font-size", "11px"
            "cursor", "pointer"
            "border-radius", "2px"
        ]
        addEventListener replaceBtn "click" (fun _ ->
            match findInput, replaceInput with
            | Some findInp, Some replInp ->
                let findVal = getProp findInp "value"
                let replVal = getProp replInp "value"
                onReplace (unbox findVal, unbox replVal)
            | _ -> ()
        )
        appendChild container replaceBtn

        container

    member _.Mount(container: Element) =
        let toolbar = createFindReplaceToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Clipboard operations toolbar
[<AttachMembers>]
type ClipboardToolbar(onAction: string -> unit) =
    let mutable toolbarElement: Element option = None

    let createClipboardToolbar () =
        let items = [
            Button {
                Id = "cut"
                Label = "Cut"
                Icon = Some "✂"
                Action = fun () -> onAction "cut"
            }
            Button {
                Id = "copy"
                Label = "Copy"
                Icon = Some "📋"
                Action = fun () -> onAction "copy"
            }
            Button {
                Id = "paste"
                Label = "Paste"
                Icon = Some "📄"
                Action = fun () -> onAction "paste"
            }
            Separator
            Button {
                Id = "undo"
                Label = "Undo"
                Icon = Some "↶"
                Action = fun () -> onAction "undo"
            }
            Button {
                Id = "redo"
                Label = "Redo"
                Icon = Some "↷"
                Action = fun () -> onAction "redo"
            }
        ]
        createToolbar items

    member _.Mount(container: Element) =
        let toolbar = createClipboardToolbar()
        appendChild container toolbar
        toolbarElement <- Some toolbar

    member _.GetElement() = toolbarElement

/// Export for JavaScript
[<ExportDefault>]
let createFormatToolbar (onAction: string -> unit): FormatToolbar =
    new FormatToolbar(onAction)

let createAlignmentToolbar (onAction: string -> unit): AlignmentToolbar =
    new AlignmentToolbar(onAction)

let createNumberFormatToolbar (onAction: string -> unit): NumberFormatToolbar =
    new NumberFormatToolbar(onAction)

let createInsertDeleteToolbar (onAction: string -> unit): InsertDeleteToolbar =
    new InsertDeleteToolbar(onAction)

let createDataToolbar (onAction: string -> unit): DataToolbar =
    new DataToolbar(onAction)

let createClipboardToolbar (onAction: string -> unit): ClipboardToolbar =
    new ClipboardToolbar(onAction)

let createFormulaBar (onCellRefChange: string -> unit) (onFormulaChange: string -> unit) (onFormulaSubmit: unit -> unit): FormulaBar =
    new FormulaBar(onCellRefChange, onFormulaChange, onFormulaSubmit)

let createFindReplaceToolbar (onFind: string -> unit) (onReplace: string * string -> unit): FindReplaceToolbar =
    new FindReplaceToolbar(onFind, onReplace)
