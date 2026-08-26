module Rbt.Grid.Hooks

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Grid.Dom

/// Cell context for rendering/editing
type CellContext = {
    row: int
    col: int
    value: obj
    columnName: string
    columnType: string
    isEditing: bool
}

/// Selection mode
type SelectionMode =
    | Row          // Select entire rows (Airtable)
    | Cell         // Single cell selection
    | Range        // Multi-cell range (Excel)

/// Selection state
type Selection = {
    mode: SelectionMode
    rows: int array option
    cells: (int * int) array option
    range: (int * int * int * int) option  // (startRow, startCol, endRow, endCol)
}

/// Cell editor interface
type ICellEditor =
    abstract GetValue: unit -> obj
    abstract Focus: unit -> unit
    abstract Destroy: unit -> unit

/// Grid customization hooks
type GridHooks = {
    /// Format value for display
    formatValue: CellContext -> string

    /// Parse user input back to value
    parseValue: string * CellContext -> obj

    /// Custom cell rendering (optional, called after default rendering)
    renderCell: (CellContext * Element -> unit) option

    /// Custom header rendering (optional)
    renderHeader: (int * string * Element -> unit) option

    /// Create cell editor (return None for non-editable)
    createEditor: (CellContext * Element -> ICellEditor option) option

    /// Check if cell can be edited
    canEdit: CellContext -> bool

    /// Selection mode
    selectionMode: SelectionMode

    /// Selection change callback
    onSelectionChange: (Selection -> unit) option
}

/// Default text editor
type TextEditor(container: Element, initialValue: string, onSave: obj -> unit, onCancel: unit -> unit) =
    let input = createElement "input"

    do
        setType input "text"
        setValue input initialValue
        setStyles input [
            "width", "100%"
            "height", "100%"
            "border", "2px solid hsl(var(--primary))"
            "padding", "4px"
            "font-size", "14px"
            "outline", "none"
        ]

        // Key handlers
        addEventListener input "keydown" (fun e ->
            match getEventKey e with
            | "Enter" ->
                preventDefault e
                onSave (box (getValue input))
            | "Escape" ->
                preventDefault e
                onCancel()
            | _ -> ()
        )

        // Blur handler
        addEventListener input "blur" (fun _ ->
            onSave (box (getValue input))
        )

        appendChild container input
        focus input
        select input

    interface ICellEditor with
        member _.GetValue() = box (getValue input)
        member _.Focus() = focus input
        member _.Destroy() = removeChild container input

/// Airtable-style default hooks
let airtableHooks: GridHooks = {
    formatValue = fun ctx ->
        if isNull ctx.value then ""
        else
            match ctx.columnType with
            | "checkbox" -> if unbox<bool> ctx.value then "Yes" else "No"
            | "currency" -> sprintf "$%.2f" (unbox<float> ctx.value)
            | "percent" -> sprintf "%.0f%%" (unbox<float> ctx.value)
            | "date" ->
                // TODO: Format date properly
                string ctx.value
            | _ -> string ctx.value

    parseValue = fun (str, ctx) ->
        if System.String.IsNullOrEmpty(str) then null
        else
            match ctx.columnType with
            | "number" | "autonumber" ->
                match System.Int32.TryParse(str) with
                | true, i -> box i
                | false, _ -> null
            | "currency" | "percent" ->
                match System.Double.TryParse(str) with
                | true, d -> box d
                | false, _ -> null
            | "checkbox" ->
                box (str.ToLower() = "true" || str = "1")
            | _ -> box str

    renderCell = None  // Use default rendering
    renderHeader = None  // Use default rendering

    createEditor = Some (fun (ctx, container) ->
        if ctx.columnType = "checkbox" then
            None  // TODO: Checkbox editor
        else
            let onSave = fun value ->
                // Callback handled by grid
                ()
            let onCancel = fun () -> ()
            Some (new TextEditor(container, string ctx.value, onSave, onCancel) :> ICellEditor)
    )

    canEdit = fun ctx -> ctx.columnName <> "id"  // Don't edit ID column

    selectionMode = Row
    onSelectionChange = None
}

/// Excel-style default hooks
let excelHooks: GridHooks = {
    formatValue = fun ctx ->
        if isNull ctx.value then ""
        else
            let str = string ctx.value
            // Check for formula
            if str.StartsWith("=") then
                // TODO: Evaluate formula
                str
            else
                str

    parseValue = fun (str, ctx) ->
        if System.String.IsNullOrEmpty(str) then null
        else
            // Try to parse as number
            match System.Double.TryParse(str) with
            | true, d -> box d
            | false, _ -> box str

    renderCell = None
    renderHeader = None

    createEditor = Some (fun (ctx, container) ->
        let onSave = fun value -> ()
        let onCancel = fun () -> ()
        Some (new TextEditor(container, string ctx.value, onSave, onCancel) :> ICellEditor)
    )

    canEdit = fun _ -> true  // All cells editable

    selectionMode = Range
    onSelectionChange = None
}

/// Create hooks object from JavaScript
[<Emit("$0 || $1")>]
let orDefault (hooks: GridHooks option) (defaultHooks: GridHooks): GridHooks = jsNative
