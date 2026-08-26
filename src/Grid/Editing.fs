module Rbt.Grid.Editing

open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
open Browser.Dom
open Rbt.Grid.Types

/// Active editor state
type Editor = {
    RowIndex: int
    ColumnName: string
    ColumnType: string
    Input: HTMLInputElement
    Cell: HTMLElement
    OriginalValue: obj
}

let mutable currentEditor: Editor option = None

/// Start editing a cell
let startEditing (rowIndex: int) (columnName: string) (colType: string) (cell: HTMLElement) (currentValue: obj) (onSave: int -> string -> obj -> unit) (onCancel: unit -> unit): unit =
    // Close any existing editor
    match currentEditor with
    | Some editor -> editor.Cell.textContent <- formatValue editor.OriginalValue editor.ColumnType
    | None -> ()

    // Create input element
    let input = document.createElement("input") :?> HTMLInputElement
    input.className <- "redbat-grid-editor"
    input.style.width <- "100%"
    input.style.height <- "100%"
    input.style.border <- "2px solid hsl(var(--primary))"
    input.style.padding <- "4px"
    input.style.fontSize <- "14px"
    input.style.outline <- "none"
    input.value <- formatValue currentValue colType

    // Save function
    let save () =
        let newValue = parseValue input.value colType
        onSave rowIndex columnName newValue
        cell.textContent <- formatValue newValue colType
        currentEditor <- None

    // Cancel function
    let cancel () =
        cell.textContent <- formatValue currentValue colType
        onCancel()
        currentEditor <- None

    // Event handlers
    input.onkeydown <- fun e ->
        match e.key with
        | "Enter" ->
            e.preventDefault()
            save()
        | "Escape" ->
            e.preventDefault()
            cancel()
        | "Tab" ->
            e.preventDefault()
            save()
            // TODO: Move to next cell
        | _ -> ()

    input.onblur <- fun _ -> save()

    // Replace cell content with input
    cell.textContent <- ""
    cell.appendChild(input) |> ignore
    input.focus()
    input.select()

    // Store editor state
    currentEditor <- Some {
        RowIndex = rowIndex
        ColumnName = columnName
        ColumnType = colType
        Input = input
        Cell = cell
        OriginalValue = currentValue
    }

/// Stop editing current cell
let stopEditing (): unit =
    match currentEditor with
    | Some editor ->
        editor.Cell.textContent <- formatValue editor.OriginalValue editor.ColumnType
        currentEditor <- None
    | None -> ()

/// Check if currently editing
let isEditing (): bool =
    currentEditor.IsSome
