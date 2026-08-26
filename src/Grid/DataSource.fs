module Rbt.Grid.DataSource

open Fable.Core
open Fable.Core.JsInterop

/// Column definition for Airtable-style grids
type ColumnDef = {
    name: string
    displayName: string option
    width: float
    columnType: string
}

/// Abstract data source interface
type IDataSource =
    /// Get value at row, column index
    abstract GetValue: row: int * col: int -> obj
    /// Set value at row, column index
    abstract SetValue: row: int * col: int * value: obj -> unit
    /// Total number of rows
    abstract GetRowCount: unit -> int
    /// Total number of columns
    abstract GetColumnCount: unit -> int
    /// Get column name/identifier (e.g., "name" or "A")
    abstract GetColumnName: col: int -> string
    /// Get column display name (e.g., "Full Name" or "A")
    abstract GetColumnDisplayName: col: int -> string option
    /// Get column type (e.g., "text", "number", "select")
    abstract GetColumnType: col: int -> string
    /// Get column width in pixels
    abstract GetColumnWidth: col: int -> float

// Helper functions for object property access
[<Emit("$0[$1]")>]
let private getProp (obj: obj) (key: string): obj = jsNative

[<Emit("$0[$1] = $2")>]
let private setProp (obj: obj) (key: string) (value: obj): unit = jsNative

/// Airtable-style: row objects with column schema
type ColumnDataSource(rows: obj array, columns: ColumnDef array) =
    interface IDataSource with
        member _.GetValue(row, col) =
            if row >= 0 && row < rows.Length && col >= 0 && col < columns.Length then
                let rowObj = rows.[row]
                let colName = columns.[col].name
                getProp rowObj colName
            else
                null

        member _.SetValue(row, col, value) =
            if row >= 0 && row < rows.Length && col >= 0 && col < columns.Length then
                let rowObj = rows.[row]
                let colName = columns.[col].name
                setProp rowObj colName value

        member _.GetRowCount() = rows.Length
        member _.GetColumnCount() = columns.Length
        member _.GetColumnName(col) = columns.[col].name
        member _.GetColumnDisplayName(col) = columns.[col].displayName
        member _.GetColumnType(col) = columns.[col].columnType
        member _.GetColumnWidth(col) = columns.[col].width

/// Excel-style: 2D cell array
type CellDataSource(cells: obj array array) =
    let rowCount = cells.Length
    let colCount = if rowCount > 0 then cells.[0].Length else 0

    /// Convert column index to letter (0 = "A", 25 = "Z", 26 = "AA")
    let columnIndexToLetter (index: int): string =
        let rec build (n: int) (acc: string) =
            if n < 0 then acc
            else
                let remainder = n % 26
                let letter = char (int 'A' + remainder) |> string
                build ((n / 26) - 1) (letter + acc)
        build index ""

    interface IDataSource with
        member _.GetValue(row, col) =
            if row >= 0 && row < rowCount && col >= 0 && col < colCount then
                cells.[row].[col]
            else
                null

        member _.SetValue(row, col, value) =
            if row >= 0 && row < rowCount && col >= 0 && col < colCount then
                cells.[row].[col] <- value

        member _.GetRowCount() = rowCount
        member _.GetColumnCount() = colCount
        member _.GetColumnName(col) = columnIndexToLetter col
        member _.GetColumnDisplayName(col) = None
        member _.GetColumnType(col) = "text"  // All cells are text by default
        member _.GetColumnWidth(col) = 100.0  // Default width

/// Helper to create column data source from JS
[<Emit("new $0($1, $2)")>]
let createColumnDataSource (rows: obj array) (columns: ColumnDef array): IDataSource = jsNative

/// Helper to create cell data source from JS
[<Emit("new $0($1)")>]
let createCellDataSource (cells: obj array array): IDataSource = jsNative
