/// Grid Logic Module - Pure computational functions for virtual scrolling
/// No DOM manipulation - meant to be called from SolidJS
module RedBat.Grid.Logic

open Fable.Core
open Fable.Core.JsInterop

/// Configuration for grid behavior
type GridConfig = {
    rowHeight: float
    headerHeight: float
    bufferRows: int
}

[<Emit("{ rowHeight: 36, headerHeight: 40, bufferRows: 5 }")>]
let defaultConfig: GridConfig = jsNative

/// Visible range of rows (result of virtual scroll calculation)
type VisibleRange = {
    startIndex: int
    endIndex: int
    totalHeight: float
    offsetTop: float
}

/// Calculate which rows should be visible based on scroll position
let calculateVisibleRows (scrollTop: float) (containerHeight: float) (totalRows: int) (config: GridConfig): VisibleRange =
    // Calculate start index with buffer
    let rawStart = scrollTop / config.rowHeight
    let startIndex = max 0.0 (rawStart - float config.bufferRows) |> int

    // Calculate how many rows fit in viewport
    let visibleCount = (containerHeight / config.rowHeight) |> ceil |> int

    // Calculate end index with buffer
    let endIndex = min totalRows (startIndex + visibleCount + config.bufferRows * 2)

    // Total content height
    let totalHeight = config.headerHeight + (float totalRows * config.rowHeight)

    // Offset from top (for absolute positioning)
    let offsetTop = float startIndex * config.rowHeight

    {
        startIndex = startIndex
        endIndex = endIndex
        totalHeight = totalHeight
        offsetTop = offsetTop
    }

/// Calculate the Y position for a specific row
let getRowTop (rowIndex: int) (config: GridConfig): float =
    config.headerHeight + (float rowIndex * config.rowHeight)

/// Get total content height
let getTotalHeight (rowCount: int) (config: GridConfig): float =
    config.headerHeight + (float rowCount * config.rowHeight)

/// Export all functions for JavaScript
[<ExportDefault>]
let exports =
    createObj [
        "defaultConfig" ==> defaultConfig
        "calculateVisibleRows" ==> calculateVisibleRows
        "getRowTop" ==> getRowTop
        "getTotalHeight" ==> getTotalHeight
    ]
