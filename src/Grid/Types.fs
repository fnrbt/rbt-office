module Rbt.Grid.Types

open System
open Fable.Core
open Browser.Types

/// Column definition
[<AllowNullLiteral>]
type ColumnDef =
    abstract name: string
    abstract displayName: string option
    abstract width: float
    abstract columnType: string

/// Row data (generic object)
type RowData = obj

/// Cell value
type CellValue = obj

/// Grid configuration
type GridConfig = {
    RowHeight: float
    HeaderHeight: float
    BufferRows: int
    EnableVirtualScroll: bool
}

/// Default configuration
[<Emit("undefined")>]
let undefined: obj = jsNative

let defaultConfig = {
    RowHeight = 36.0
    HeaderHeight = 40.0
    BufferRows = 5
    EnableVirtualScroll = true
}

/// Editing state
type EditingCell = {
    RowIndex: int
    ColumnName: string
    Value: string
}

/// Sort state
type SortState = {
    Column: string
    Descending: bool
}

/// Grid state
type GridState = {
    ScrollTop: float
    SelectedRows: Set<int>
    EditingCell: EditingCell option
    SortState: SortState option
    HoveredRow: int option
}

let emptyState = {
    ScrollTop = 0.0
    SelectedRows = Set.empty
    EditingCell = None
    SortState = None
    HoveredRow = None
}

/// Event callbacks
[<AllowNullLiteral>]
type GridCallbacks =
    abstract onCellChange: rowIndex: int * columnName: string * value: obj -> unit
    abstract onAddRow: unit -> unit
    abstract onDeleteRow: rowIndex: int -> unit
    abstract onAddColumn: unit -> unit
    abstract onDeleteColumn: columnName: string -> unit
    abstract onSort: columnName: string * descending: bool -> unit

/// Helper to get property from JS object
[<Emit("$0[$1]")>]
let getProp (obj: obj) (key: string): obj = jsNative

/// Helper to set property on JS object
[<Emit("$0[$1] = $2")>]
let setProp (obj: obj) (key: string) (value: obj): unit = jsNative

/// Helper to check if value is null or undefined
[<Emit("$0 == null")>]
let isNullOrUndefined (value: obj): bool = jsNative

/// Helper to format value for display
let formatValue (value: obj) (colType: string): string =
    if isNullOrUndefined value then
        ""
    else
        match colType with
        | "checkbox" -> if unbox<bool> value then "Yes" else "No"
        | "currency" ->
            let num = unbox<float> value
            sprintf "$%.2f" num
        | "percent" ->
            let num = unbox<float> value
            sprintf "%.0f%%" num
        | "date" ->
            try
                let date = unbox<DateTime> value
                date.ToLocaleDateString()
            with _ -> string value
        | "datetime" ->
            try
                let date = unbox<DateTime> value
                date.ToLocaleString()
            with _ -> string value
        | _ -> string value

/// Helper to parse value from string
let parseValue (str: string) (colType: string): obj =
    if String.IsNullOrEmpty(str) then
        null
    else
        match colType with
        | "number" | "autonumber" ->
            match Int32.TryParse(str) with
            | true, i -> box i
            | false, _ -> null
        | "decimal" | "currency" | "percent" ->
            match Double.TryParse(str) with
            | true, d -> box d
            | false, _ -> null
        | "checkbox" ->
            box (str.ToLower() = "true" || str = "1")
        | _ -> box str
