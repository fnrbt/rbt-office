module RedBat.Grid.Rendering

open Fable.Core
open Fable.Core.JsInterop
open Browser.Types
open Browser.Dom
open RedBat.Grid.Types
open RedBat.Grid.VirtualScroll

/// Create the main container element
let createContainer (containerHeight: float): HTMLElement =
    let div = document.createElement("div") :?> HTMLElement
    div.className <- "redbat-grid-container"
    div.style.height <- sprintf "%.0fpx" containerHeight
    div.style.overflow <- "auto"
    div.style.position <- "relative"
    div

/// Create the content wrapper (sized for all rows)
let createContent (totalHeight: float): HTMLElement =
    let div = document.createElement("div") :?> HTMLElement
    div.className <- "redbat-grid-content"
    div.style.height <- sprintf "%.0fpx" totalHeight
    div.style.position <- "relative"
    div

/// Create header row
let createHeader (columns: ColumnDef array) (config: GridConfig) (onSort: string -> bool -> unit) (onDeleteColumn: string -> unit) (onAddColumn: unit -> unit) (currentSort: SortState option) (isReadOnly: bool): HTMLElement =
    let header = document.createElement("div") :?> HTMLElement
    header.className <- "redbat-grid-header"
    header.style.position <- "sticky"
    header.style.top <- "0"
    header.style.display <- "flex"
    header.style.height <- sprintf "%.0fpx" config.HeaderHeight
    header.style.backgroundColor <- "hsl(var(--muted) / 0.8)"
    header.style.borderBottom <- "1px solid hsl(var(--border))"
    header.style.zIndex <- "10"
    header.style.backdropFilter <- "blur(8px)"

    // Selection checkbox column
    let selectCol = document.createElement("div") :?> HTMLElement
    selectCol.style.width <- "40px"
    selectCol.style.padding <- "8px"
    selectCol.style.display <- "flex"
    selectCol.style.alignItems <- "center"
    header.appendChild(selectCol) |> ignore

    // Row number column
    let numCol = document.createElement("div") :?> HTMLElement
    numCol.style.width <- "48px"
    numCol.style.padding <- "8px"
    numCol.style.fontSize <- "12px"
    numCol.style.color <- "hsl(var(--muted-foreground))"
    numCol.textContent <- "#"
    header.appendChild(numCol) |> ignore

    // Data columns
    for col in columns do
        let headerCell = document.createElement("div") :?> HTMLElement
        headerCell.className <- "redbat-grid-header-cell"
        headerCell.style.minWidth <- sprintf "%.0fpx" col.width
        headerCell.style.padding <- "8px 12px"
        headerCell.style.borderLeft <- "1px solid hsl(var(--border))"
        headerCell.style.display <- "flex"
        headerCell.style.alignItems <- "center"
        headerCell.style.gap <- "8px"
        headerCell.style.flexShrink <- "0"

        // Column name button (for sorting)
        let nameBtn = document.createElement("button") :?> HTMLButtonElement
        nameBtn.className <- "redbat-grid-column-name"
        nameBtn.style.flex <- "1"
        nameBtn.style.textAlign <- "left"
        nameBtn.style.fontWeight <- "500"
        nameBtn.style.cursor <- "pointer"
        nameBtn.style.background <- "none"
        nameBtn.style.border <- "none"
        nameBtn.style.padding <- "0"
        nameBtn.style.color <- "inherit"

        let displayName = match col.displayName with Some n -> n | None -> col.name
        let sortIndicator =
            match currentSort with
            | Some s when s.Column = col.name -> if s.Descending then " ↓" else " ↑"
            | _ -> ""
        nameBtn.textContent <- displayName + sortIndicator

        nameBtn.onclick <- fun _ ->
            let isCurrentSort = match currentSort with Some s -> s.Column = col.name | None -> false
            let desc = if isCurrentSort then not currentSort.Value.Descending else false
            onSort col.name desc

        headerCell.appendChild(nameBtn) |> ignore

        // Delete column button (if not readonly)
        if not isReadOnly then
            let deleteBtn = document.createElement("button") :?> HTMLButtonElement
            deleteBtn.className <- "redbat-grid-delete-column"
            deleteBtn.textContent <- "×"
            deleteBtn.style.width <- "20px"
            deleteBtn.style.height <- "20px"
            deleteBtn.style.padding <- "0"
            deleteBtn.style.border <- "none"
            deleteBtn.style.background <- "none"
            deleteBtn.style.cursor <- "pointer"
            deleteBtn.style.opacity <- "0"
            deleteBtn.style.color <- "hsl(var(--destructive))"
            deleteBtn.title <- "Delete column"
            deleteBtn.onclick <- fun _ -> onDeleteColumn col.name

            headerCell.onmouseenter <- fun _ -> deleteBtn.style.opacity <- "1"
            headerCell.onmouseleave <- fun _ -> deleteBtn.style.opacity <- "0"

            headerCell.appendChild(deleteBtn) |> ignore

        header.appendChild(headerCell) |> ignore

    // Add column button
    if not isReadOnly then
        let addBtn = document.createElement("button") :?> HTMLButtonElement
        addBtn.className <- "redbat-grid-add-column"
        addBtn.textContent <- "+"
        addBtn.style.width <- "40px"
        addBtn.style.padding <- "8px"
        addBtn.style.borderLeft <- "1px solid hsl(var(--border))"
        addBtn.style.cursor <- "pointer"
        addBtn.style.background <- "none"
        addBtn.style.border <- "none"
        addBtn.style.flexShrink <- "0"
        addBtn.title <- "Add column"
        addBtn.onclick <- fun _ -> onAddColumn()
        header.appendChild(addBtn) |> ignore

    header

/// Create a single data row
let createRow (rowIndex: int) (rowData: obj) (columns: ColumnDef array) (config: GridConfig) (state: GridState) (onClick: int -> string -> unit) (onCheck: int -> bool -> unit) (onDeleteRow: int -> unit): HTMLElement =
    let row = document.createElement("div") :?> HTMLElement
    row.className <- "redbat-grid-row"
    row.style.position <- "absolute"
    row.style.top <- sprintf "%.0fpx" (calculateRowTop rowIndex config)
    row.style.display <- "flex"
    row.style.height <- sprintf "%.0fpx" config.RowHeight
    row.style.borderBottom <- "1px solid hsl(var(--border))"

    let isSelected = state.SelectedRows.Contains rowIndex
    if isSelected then
        row.style.backgroundColor <- "hsl(var(--accent) / 0.5)"

    row.onmouseenter <- fun _ ->
        if not isSelected then
            row.style.backgroundColor <- "hsl(var(--accent) / 0.3)"

    row.onmouseleave <- fun _ ->
        if not isSelected then
            row.style.backgroundColor <- ""

    // Selection checkbox
    let selectCell = document.createElement("div") :?> HTMLElement
    selectCell.style.width <- "40px"
    selectCell.style.padding <- "6px 8px"
    selectCell.style.display <- "flex"
    selectCell.style.alignItems <- "center"

    let checkbox = document.createElement("input") :?> HTMLInputElement
    checkbox.``type`` <- "checkbox"
    checkbox.checked <- isSelected
    checkbox.onchange <- fun _ -> onCheck rowIndex checkbox.checked
    selectCell.appendChild(checkbox) |> ignore
    row.appendChild(selectCell) |> ignore

    // Row number
    let numCell = document.createElement("div") :?> HTMLElement
    numCell.style.width <- "48px"
    numCell.style.padding <- "6px 8px"
    numCell.style.fontSize <- "12px"
    numCell.style.color <- "hsl(var(--muted-foreground))"
    numCell.style.display <- "flex"
    numCell.style.alignItems <- "center"
    numCell.textContent <- string (rowIndex + 1)
    row.appendChild(numCell) |> ignore

    // Data cells
    for col in columns do
        let cell = document.createElement("div") :?> HTMLElement
        cell.className <- "redbat-grid-cell"
        cell.style.minWidth <- sprintf "%.0fpx" col.width
        cell.style.padding <- "6px 12px"
        cell.style.borderLeft <- "1px solid hsl(var(--border))"
        cell.style.cursor <- "text"
        cell.style.display <- "flex"
        cell.style.alignItems <- "center"
        cell.style.flexShrink <- "0"
        cell.style.overflow <- "hidden"
        cell.style.textOverflow <- "ellipsis"
        cell.style.whiteSpace <- "nowrap"

        let value = getProp rowData col.name
        let displayValue = formatValue value col.columnType

        cell.textContent <- displayValue
        cell.onclick <- fun _ -> onClick rowIndex col.name

        row.appendChild(cell) |> ignore

    row

/// Render visible rows
let renderRows (container: HTMLElement) (rows: obj array) (columns: ColumnDef array) (visibleRange: VisibleRange) (config: GridConfig) (state: GridState) (onCellClick: int -> string -> unit) (onCheck: int -> bool -> unit) (onDeleteRow: int -> unit): unit =
    // Clear existing rows (except header)
    let children = container.children
    let mutable i = 1 // Skip header at index 0
    while i < children.length do
        container.removeChild(children.[i]) |> ignore

    // Render visible rows
    for i in visibleRange.StartIndex .. visibleRange.EndIndex - 1 do
        if i < rows.Length then
            let rowElement = createRow i rows.[i] columns config state onCellClick onCheck onDeleteRow
            container.appendChild(rowElement) |> ignore
