module RedBat.Grid.VirtualScroll

open RedBat.Grid.Types

/// Calculate which rows should be visible based on scroll position
type VisibleRange = {
    StartIndex: int
    EndIndex: int
    OffsetTop: float
}

let calculateVisibleRange (scrollTop: float) (containerHeight: float) (totalRows: int) (config: GridConfig): VisibleRange =
    if not config.EnableVirtualScroll then
        // No virtualization - render all rows
        {
            StartIndex = 0
            EndIndex = totalRows
            OffsetTop = 0.0
        }
    else
        // Calculate visible range with buffer
        let startIndex = int (scrollTop / config.RowHeight) - config.BufferRows
        let visibleCount = int (containerHeight / config.RowHeight) + 1
        let endIndex = startIndex + visibleCount + (config.BufferRows * 2)

        // Clamp to valid range
        let startClamped = max 0 startIndex
        let endClamped = min totalRows endIndex

        {
            StartIndex = startClamped
            EndIndex = endClamped
            OffsetTop = float startClamped * config.RowHeight
        }

/// Calculate total content height (for scroll container)
let calculateTotalHeight (rowCount: int) (config: GridConfig): float =
    config.HeaderHeight + (float rowCount * config.RowHeight)

/// Calculate row position (absolute top offset)
let calculateRowTop (rowIndex: int) (config: GridConfig): float =
    config.HeaderHeight + (float rowIndex * config.RowHeight)
