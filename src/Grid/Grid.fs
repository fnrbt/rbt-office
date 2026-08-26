module Rbt.Grid.Grid

open Fable.Core
open Fable.Core.JsInterop
open Rbt.Grid.Dom
open Rbt.Grid.DataSource
open Rbt.Grid.Hooks

/// Grid configuration
type GridConfig = {
    rowHeight: float
    headerHeight: float
    bufferRows: int
}

let defaultConfig = {
    rowHeight = 36.0
    headerHeight = 40.0
    bufferRows = 5
}

/// Calculate visible row range
type VisibleRange = {
    startIndex: int
    endIndex: int
}

let calculateVisibleRange (scrollTop: float) (containerHeight: float) (totalRows: int) (config: GridConfig): VisibleRange =
    let startIndex = int (scrollTop / config.rowHeight) - config.bufferRows |> max 0
    let visibleCount = int (containerHeight / config.rowHeight) + 1
    let endIndex = startIndex + visibleCount + (config.bufferRows * 2) |> min totalRows
    { startIndex = startIndex; endIndex = endIndex }

/// Main Grid class
type Grid(container: Element, dataSource: IDataSource, hooks: GridHooks, config: GridConfig) =
    let mutable scrollTop = 0.0
    let mutable containerHeight = 600.0
    let mutable currentSort: (string * bool) option = None
    let mutable isReadOnly = false
    let mutable selectedRows: Set<int> = Set.empty
    let mutable currentEditor: ICellEditor option = None

    let mutable gridContainer: Element option = None
    let mutable contentElement: Element option = None
    let mutable headerElement: Element option = None

    // Callbacks
    let mutable onCellChangeCallback: (int * string * obj -> unit) option = None
    let mutable onAddRowCallback: (unit -> unit) option = None
    let mutable onDeleteRowCallback: (int -> unit) option = None
    let mutable onAddColumnCallback: (unit -> unit) option = None
    let mutable onDeleteColumnCallback: (string -> unit) option = None
    let mutable onSortCallback: (string * bool -> unit) option = None

    /// Create header
    let createHeader () =
        let header = createElement "div"
        setClassName header "redbat-grid-header"
        setStyles header [
            "position", "sticky"
            "top", "0"
            "display", "flex"
            "height", sprintf "%.0fpx" config.headerHeight
            "background-color", "hsl(var(--muted) / 0.8)"
            "border-bottom", "1px solid hsl(var(--border))"
            "z-index", "10"
            "backdrop-filter", "blur(8px)"
        ]

        // Selection checkbox column (only for Row selection mode)
        if hooks.selectionMode = Row then
            let selectCol = createElement "div"
            setStyles selectCol ["width", "40px"; "padding", "8px"]
            appendChild header selectCol

        // Row number column
        let numCol = createElement "div"
        setStyles numCol [
            "width", "48px"
            "padding", "8px"
            "font-size", "12px"
            "color", "hsl(var(--muted-foreground))"
        ]
        setTextContent numCol "#"
        appendChild header numCol

        // Data columns
        let colCount = dataSource.GetColumnCount()
        for colIdx = 0 to colCount - 1 do
            let colName = dataSource.GetColumnName(colIdx)
            let displayName =
                match dataSource.GetColumnDisplayName(colIdx) with
                | Some name -> name
                | None -> colName
            let width = dataSource.GetColumnWidth(colIdx)

            let headerCell = createElement "div"
            setClassName headerCell "redbat-grid-header-cell"
            setStyles headerCell [
                "min-width", sprintf "%.0fpx" width
                "padding", "8px 12px"
                "border-left", "1px solid hsl(var(--border))"
                "display", "flex"
                "align-items", "center"
                "gap", "8px"
                "flex-shrink", "0"
            ]

            // Custom header rendering or default
            match hooks.renderHeader with
            | Some customRender ->
                customRender(colIdx, colName, headerCell)
            | None ->
                // Default header: column name button
                let nameBtn = createElement "button"
                setStyles nameBtn [
                    "flex", "1"
                    "text-align", "left"
                    "font-weight", "500"
                    "cursor", "pointer"
                    "background", "none"
                    "border", "none"
                    "padding", "0"
                    "color", "inherit"
                ]

                let sortIndicator =
                    match currentSort with
                    | Some (sortCol, desc) when sortCol = colName -> if desc then " ↓" else " ↑"
                    | _ -> ""
                setTextContent nameBtn (displayName + sortIndicator)

                addEventListener nameBtn "click" (fun _ ->
                    match onSortCallback with
                    | Some cb ->
                        let isCurrentSort = match currentSort with Some (c, _) -> c = colName | None -> false
                        let desc = if isCurrentSort then not (snd currentSort.Value) else false
                        cb(colName, desc)
                    | None -> ()
                )

                appendChild headerCell nameBtn

            appendChild header headerCell

        // Add column button
        if not isReadOnly then
            let addBtn = createElement "button"
            setStyles addBtn [
                "width", "40px"
                "padding", "8px"
                "border-left", "1px solid hsl(var(--border))"
                "cursor", "pointer"
                "background", "none"
                "border", "none"
                "flex-shrink", "0"
            ]
            setTextContent addBtn "+"
            setTitle addBtn "Add column"
            addEventListener addBtn "click" (fun _ ->
                match onAddColumnCallback with
                | Some cb -> cb()
                | None -> ()
            )
            appendChild header addBtn

        header

    /// Start editing a cell
    let startEditing (rowIdx: int) (colIdx: int) (cell: Element) =
        // Close current editor if any
        match currentEditor with
        | Some editor -> editor.Destroy()
        | None -> ()

        let value = dataSource.GetValue(rowIdx, colIdx)
        let colName = dataSource.GetColumnName(colIdx)
        let colType = dataSource.GetColumnType(colIdx)

        let ctx = {
            row = rowIdx
            col = colIdx
            value = value
            columnName = colName
            columnType = colType
            isEditing = true
        }

        // Check if editable
        if not (hooks.canEdit ctx) then
            ()
        else
            // Clear cell
            setTextContent cell ""

            // Create editor
            match hooks.createEditor with
            | Some createFn ->
                match createFn(ctx, cell) with
                | Some editor ->
                    currentEditor <- Some editor
                    editor.Focus()
                | None -> ()
            | None -> ()

    /// Create a single row
    let createRow (rowIndex: int) =
        let row = createElement "div"
        setClassName row "redbat-grid-row"
        let top = config.headerHeight + (float rowIndex * config.rowHeight)
        setStyles row [
            "position", "absolute"
            "top", sprintf "%.0fpx" top
            "display", "flex"
            "height", sprintf "%.0fpx" config.rowHeight
            "border-bottom", "1px solid hsl(var(--border))"
            "left", "0"
            "right", "0"
        ]

        // Hover effect
        addEventListener row "mouseenter" (fun _ ->
            if not (selectedRows.Contains rowIndex) then
                setStyle row "background-color" "hsl(var(--accent) / 0.3)"
        )
        addEventListener row "mouseleave" (fun _ ->
            if not (selectedRows.Contains rowIndex) then
                setStyle row "background-color" ""
        )

        // Selection checkbox (Row mode only)
        if hooks.selectionMode = Row then
            let selectCell = createElement "div"
            setStyles selectCell ["width", "40px"; "padding", "6px 8px"; "display", "flex"; "align-items", "center"]
            let checkbox = createElement "input"
            setType checkbox "checkbox"
            setChecked checkbox (selectedRows.Contains rowIndex)
            addEventListener checkbox "change" (fun _ ->
                let isChecked = getChecked checkbox
                selectedRows <-
                    if isChecked then selectedRows.Add rowIndex
                    else selectedRows.Remove rowIndex

                // Trigger selection change callback
                match hooks.onSelectionChange with
                | Some cb ->
                    let selection = {
                        mode = Row
                        rows = Some (Set.toArray selectedRows)
                        cells = None
                        range = None
                    }
                    cb selection
                | None -> ()
            )
            appendChild selectCell checkbox
            appendChild row selectCell

        // Row number
        let numCell = createElement "div"
        setStyles numCell [
            "width", "48px"
            "padding", "6px 8px"
            "font-size", "12px"
            "color", "hsl(var(--muted-foreground))"
            "display", "flex"
            "align-items", "center"
        ]
        setTextContent numCell (string (rowIndex + 1))
        appendChild row numCell

        // Data cells
        let colCount = dataSource.GetColumnCount()
        for colIdx = 0 to colCount - 1 do
            let value = dataSource.GetValue(rowIndex, colIdx)
            let colName = dataSource.GetColumnName(colIdx)
            let colType = dataSource.GetColumnType(colIdx)
            let width = dataSource.GetColumnWidth(colIdx)

            let cell = createElement "div"
            setClassName cell "redbat-grid-cell"
            setStyles cell [
                "min-width", sprintf "%.0fpx" width
                "padding", "6px 12px"
                "border-left", "1px solid hsl(var(--border))"
                "cursor", "text"
                "display", "flex"
                "align-items", "center"
                "flex-shrink", "0"
                "overflow", "hidden"
                "text-overflow", "ellipsis"
                "white-space", "nowrap"
            ]

            let ctx = {
                row = rowIndex
                col = colIdx
                value = value
                columnName = colName
                columnType = colType
                isEditing = false
            }

            // Custom rendering or default
            match hooks.renderCell with
            | Some customRender ->
                customRender(ctx, cell)
            | None ->
                let displayValue = hooks.formatValue ctx
                setTextContent cell displayValue

            // Click to edit
            addEventListener cell "click" (fun _ ->
                startEditing rowIndex colIdx cell
            )

            appendChild row cell

        // Empty cell for add column button alignment
        if not isReadOnly then
            let emptyCell = createElement "div"
            setStyles emptyCell ["width", "40px"; "border-left", "1px solid hsl(var(--border))"]
            appendChild row emptyCell

        row

    /// Render visible rows
    let render () =
        match contentElement with
        | None -> ()
        | Some content ->
            // Remove old rows (keep header)
            let children = getChildren content
            if children.Length > 1 then
                for i = children.Length - 1 downto 1 do
                    removeChild content children.[i]

            // Calculate visible range
            let rowCount = dataSource.GetRowCount()
            let range = calculateVisibleRange scrollTop containerHeight rowCount config

            // Render visible rows
            for i = range.startIndex to range.endIndex - 1 do
                if i < rowCount then
                    let rowElement = createRow i
                    appendChild content rowElement

    /// Handle scroll
    let handleScroll (e: obj) =
        let target = getCurrentTarget e
        scrollTop <- getScrollTop target
        render()

    /// Mount the grid
    member _.Mount() =
        // Create container
        let gridDiv = createElement "div"
        setClassName gridDiv "redbat-grid-container"
        setStyles gridDiv [
            "height", "600px"
            "overflow", "auto"
            "position", "relative"
        ]

        // Create content wrapper
        let content = createElement "div"
        setClassName content "redbat-grid-content"
        let rowCount = dataSource.GetRowCount()
        let totalHeight = config.headerHeight + (float rowCount * config.rowHeight)
        setStyles content [
            "height", sprintf "%.0fpx" totalHeight
            "position", "relative"
        ]

        // Create header
        let header = createHeader()
        appendChild content header

        // Add scroll listener
        addEventListener gridDiv "scroll" handleScroll

        appendChild gridDiv content
        appendChild container gridDiv

        gridContainer <- Some gridDiv
        contentElement <- Some content
        headerElement <- Some header

        // Initial render
        containerHeight <- getClientHeight gridDiv
        render()

    /// Refresh the grid (when data source changes)
    member _.Refresh() =
        match contentElement with
        | Some content ->
            let rowCount = dataSource.GetRowCount()
            let totalHeight = config.headerHeight + (float rowCount * config.rowHeight)
            setStyle content "height" (sprintf "%.0fpx" totalHeight)
        | None -> ()
        render()

    /// Update sort state
    member _.SetSort(column: string option, desc: bool) =
        currentSort <- column |> Option.map (fun c -> (c, desc))
        match contentElement, headerElement with
        | Some content, Some oldHeader ->
            removeChild content oldHeader
            let newHeader = createHeader()
            match getFirstChild content with
            | Some firstChild -> insertBefore content newHeader firstChild
            | None -> appendChild content newHeader
            headerElement <- Some newHeader
        | _ -> ()

    /// Set readonly mode
    member _.SetReadOnly(readOnly: bool) =
        isReadOnly <- readOnly

    /// Set callbacks
    member _.OnCellChange(cb: int * string * obj -> unit) =
        onCellChangeCallback <- Some cb

    member _.OnAddRow(cb: unit -> unit) =
        onAddRowCallback <- Some cb

    member _.OnDeleteRow(cb: int -> unit) =
        onDeleteRowCallback <- Some cb

    member _.OnAddColumn(cb: unit -> unit) =
        onAddColumnCallback <- Some cb

    member _.OnDeleteColumn(cb: string -> unit) =
        onDeleteColumnCallback <- Some cb

    member _.OnSort(cb: string * bool -> unit) =
        onSortCallback <- Some cb

    /// Destroy the grid
    member _.Destroy() =
        match currentEditor with
        | Some editor -> editor.Destroy()
        | None -> ()

        match gridContainer with
        | Some grid -> removeChild container grid
        | None -> ()
        gridContainer <- None
        contentElement <- None
        headerElement <- None

/// Export for JavaScript
[<ExportDefault>]
let exports =
    createObj [
        "createGrid" ==> fun (container: Element) (dataSource: IDataSource) (hooks: obj) (config: obj) ->
            let gridHooks =
                if isNull hooks then airtableHooks
                else unbox<GridHooks> hooks
            let gridConfig =
                if isNull config then defaultConfig
                else unbox<GridConfig> config
            new Grid(container, dataSource, gridHooks, gridConfig)

        "ColumnDataSource" ==> fun (rows: obj array) (columns: ColumnDef array) ->
            new ColumnDataSource(rows, columns) :> IDataSource

        "CellDataSource" ==> fun (cells: obj array array) ->
            new CellDataSource(cells) :> IDataSource

        "airtableHooks" ==> airtableHooks
        "excelHooks" ==> excelHooks
    ]
