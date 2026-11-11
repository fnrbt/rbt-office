module Rbt.Widgets.Composit.Ui.SpreadsheetEditor

open Rbt.Renderer
open Rbt.Widgets.Composit.Ui.Reactive
open Rbt.Widgets.Composit.Ui.Widget

open Rbt.Spreadsheet.Model

module Layout = Rbt.Spreadsheet.Layout
module SpreadsheetGrid = Rbt.Widgets.Composit.Ui.SpreadsheetGrid
module Workbook = Rbt.Spreadsheet.Workbook
module Recalc = Rbt.Spreadsheet.Recalc

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

/// Workbook is authoritative; arrays are read-only projections for the UI.
/// History snapshots include names, tables, iteration, sheet identity and order.
type SpreadsheetEditorState = {
    Workbook: RefSignal<Workbook.Workbook>
    Sheets: RefSignal<SpreadsheetModel array>
    SheetNames: Signal<string array>
    ActiveSheet: Signal<int>
    ScrollX: Signal<float>
    ScrollY: Signal<float>
    VisibleWidth: Signal<float>
    VisibleHeight: Signal<float>
    EditingCell: Signal<(int * int) option>
    EditBuffer: Signal<string>
    DragAnchor: Signal<(int * int) option>
    UndoStack: RefSignal<Workbook.Workbook list>
    RedoStack: RefSignal<Workbook.Workbook list>
    CalcEngine: RefSignal<Recalc.Engine option>
    Clipboard: RefSignal<(Workbook.SheetId * ClipboardData * string) option>
    IssuedSheetIds: System.Collections.Generic.HashSet<Workbook.SheetId>
    InvalidatedClipboardText: Signal<string option>
    FocusId: int
}

let private sheetId (state: SpreadsheetEditorState) =
    state.Workbook.Value.ActiveSheet

let private projectWorkbook (state: SpreadsheetEditorState) (book: Workbook.Workbook) =
    let previous = state.Workbook.Value
    let oldSheets = state.Sheets.Value
    let oldNames = state.SheetNames.Value
    if obj.ReferenceEquals(previous.SheetOrder, book.SheetOrder) && oldSheets.Length = book.SheetOrder.Length && oldNames.Length = oldSheets.Length then
        let mutable models = oldSheets
        let mutable names = oldNames
        book.SheetOrder |> List.iteri (fun index id ->
            let sheet = book.Sheets.[id]
            if not (obj.ReferenceEquals(models.[index], sheet.Model)) then
                if obj.ReferenceEquals(models, oldSheets) then models <- Array.copy oldSheets
                models.[index] <- sheet.Model
            if names.[index] <> sheet.Name then
                if obj.ReferenceEquals(names, oldNames) then names <- Array.copy oldNames
                names.[index] <- sheet.Name)
        if not (obj.ReferenceEquals(models, oldSheets)) then state.Sheets.Value <- models
        if not (obj.ReferenceEquals(names, oldNames)) then state.SheetNames.Value <- names
    else
        state.Sheets.Value <- book.SheetOrder |> List.map (fun id -> book.Sheets.[id].Model) |> List.toArray
        state.SheetNames.Value <- book.SheetOrder |> List.map (fun id -> book.Sheets.[id].Name) |> List.toArray
    state.Workbook.Value <- book
    state.ActiveSheet.Value <- Workbook.sheetIndex book.ActiveSheet book
    match state.Clipboard.Value with
    | Some (source, _, text) when not (Map.containsKey source book.Sheets) ->
        state.InvalidatedClipboardText.Value <- Some text
        state.Clipboard.Value <- None
    | _ -> ()

/// Imported workbooks can be supplied without losing sheet IDs or metadata.
let createStateFromWorkbook (book: Workbook.Workbook) =
    let state = {
        Workbook = refSignal book
        Sheets = refSignal [||]
        SheetNames = signal [||]
        ActiveSheet = signal 0
        ScrollX = signal 0.0
        ScrollY = signal 0.0
        VisibleWidth = signal 800.0
        VisibleHeight = signal 600.0
        EditingCell = signal None
        EditBuffer = signal ""
        DragAnchor = signal None
        UndoStack = refSignal []
        RedoStack = refSignal []
        CalcEngine = refSignal None
        Clipboard = refSignal None
        IssuedSheetIds = System.Collections.Generic.HashSet<Workbook.SheetId>(book.SheetOrder)
        InvalidatedClipboardText = signal None
        FocusId = Focus.allocId()
    }
    projectWorkbook state book
    state

/// Fresh models have no imported workbook metadata.
let createState (sheets: SpreadsheetModel array) (names: string array) (activeIndex: int) =
    if sheets.Length = 0 then invalidArg "sheets" "A workbook must have a sheet"
    if names.Length <> sheets.Length then invalidArg "names" "Sheet names must match sheet models"
    Rbt.Spreadsheet.XlsxBridge.buildWorkbookContext sheets names (Array.create sheets.Length true) Map.empty Map.empty activeIndex
    |> createStateFromWorkbook

let activeModel (state: SpreadsheetEditorState) : SpreadsheetModel =
    state.Workbook.Value.Sheets.[sheetId state].Model

let activeSheetName (state: SpreadsheetEditorState) : string =
    state.Workbook.Value.Sheets.[sheetId state].Name

let private recalculate (state: SpreadsheetEditorState) (book: Workbook.Workbook) hint =
    match state.CalcEngine.Value with
    | Some engine ->
        match Recalc.recalculate engine book hint with
        | Some calculated -> calculated
        | None ->
            let calculated, engine = Recalc.recalculateAll book
            state.CalcEngine.Value <- Some engine
            calculated
    | None ->
        let calculated, engine = Recalc.recalculateAll book
        state.CalcEngine.Value <- Some engine
        calculated

let private installCalculatedEdit (state: SpreadsheetEditorState) (before: Workbook.Workbook) (calculated: Workbook.Workbook) =
    state.UndoStack.Value <- before :: state.UndoStack.Value
    state.RedoStack.Value <- []
    projectWorkbook state calculated
    Platform.requestRedraw()
    true

let private installEdit (state: SpreadsheetEditorState) (before: Workbook.Workbook) (after: Workbook.Workbook) hint =
    if obj.ReferenceEquals(before, after) then false
    else installCalculatedEdit state before (recalculate state after hint)

let applyWorkbookChange (state: SpreadsheetEditorState) (transform: Workbook.Workbook -> Workbook.Workbook) =
    let before = state.Workbook.Value
    installEdit state before (transform before) None

/// Model transforms may write cells (including through macros), so always
/// recalculate across the entire workbook. Selection-only changes use the
/// no-history path below instead.
let applyChange (state: SpreadsheetEditorState) (transform: SpreadsheetModel -> SpreadsheetModel) =
    let before = state.Workbook.Value
    let id = before.ActiveSheet
    let oldModel = before.Sheets.[id].Model
    let changed = transform oldModel
    if obj.ReferenceEquals(oldModel, changed) then false
    else installEdit state before (Workbook.updateSheetModelWithTables id changed before) None

let applyChangeNoHistory (state: SpreadsheetEditorState) (transform: SpreadsheetModel -> SpreadsheetModel) =
    let before = state.Workbook.Value
    let id = before.ActiveSheet
    let oldModel = before.Sheets.[id].Model
    let changed = transform oldModel
    if obj.ReferenceEquals(oldModel, changed) then false
    else
        let book = Workbook.updateSheetModelWithTables id changed before
        if obj.ReferenceEquals(oldModel.Cells, changed.Cells) && obj.ReferenceEquals(oldModel.Tables, changed.Tables) then
            projectWorkbook state book
        else projectWorkbook state (recalculate state book None)
        Platform.requestRedraw()
        true

let private restore (state: SpreadsheetEditorState) (book: Workbook.Workbook) =
    state.CalcEngine.Value <- None
    match state.Clipboard.Value with
    | Some (_, data, _) when data.IsCut -> state.Clipboard.Value <- None
    | _ -> ()
    projectWorkbook state book
    state.EditingCell.Value <- None
    state.EditBuffer.Value <- ""
    Platform.requestRedraw()

let undo (state: SpreadsheetEditorState) =
    match state.UndoStack.Value with
    | previous :: rest ->
        state.RedoStack.Value <- state.Workbook.Value :: state.RedoStack.Value
        state.UndoStack.Value <- rest
        restore state previous
        true
    | [] -> true

let redo (state: SpreadsheetEditorState) =
    match state.RedoStack.Value with
    | next :: rest ->
        state.UndoStack.Value <- state.Workbook.Value :: state.UndoStack.Value
        state.RedoStack.Value <- rest
        restore state next
        true
    | [] -> true

let switchSheet (state: SpreadsheetEditorState) (index: int) =
    let order = state.Workbook.Value.SheetOrder
    if index >= 0 && index < order.Length && index <> state.ActiveSheet.Value then
        projectWorkbook state (Workbook.setActiveSheet order.[index] state.Workbook.Value)
        state.ScrollX.Value <- 0.0
        state.ScrollY.Value <- 0.0
        state.EditingCell.Value <- None
        state.EditBuffer.Value <- ""
        Platform.requestRedraw()

/// Replace an open document without carrying pending edits, selection drag,
/// clipboard provenance, history or calculated state into the new workbook.
let replaceWorkbook (state: SpreadsheetEditorState) (book: Workbook.Workbook) =
    state.EditingCell.Value <- None
    state.EditBuffer.Value <- ""
    state.DragAnchor.Value <- None
    state.ScrollX.Value <- 0.0
    state.ScrollY.Value <- 0.0
    state.UndoStack.Value <- []
    state.RedoStack.Value <- []
    state.Clipboard.Value <- None
    state.InvalidatedClipboardText.Value <- None
    state.CalcEngine.Value <- None
    state.IssuedSheetIds.Clear()
    for id in book.SheetOrder do state.IssuedSheetIds.Add id |> ignore
    projectWorkbook state book
    Platform.requestRedraw()

let insertSheet (state: SpreadsheetEditorState) (name: string) =
    let current = state.Workbook.Value
    if (Workbook.sheetNameProblem name).IsSome || Workbook.sheetNameTaken None name current then false
    else applyWorkbookChange state (fun book ->
        let id = Workbook.freshSheetId state.IssuedSheetIds.Contains book
        state.IssuedSheetIds.Add id |> ignore
        let sheet : Workbook.Sheet =
            { Id = id; Name = name; Model = { createEmpty 1000 26 with Date1904 = Workbook.date1904 book }; Visible = true }
        Workbook.insertSheet sheet book.SheetOrder.Length book)

let renameSheet (state: SpreadsheetEditorState) (id: Workbook.SheetId) (name: string) =
    let current = state.Workbook.Value
    if (Workbook.sheetNameProblem name).IsSome || Workbook.sheetNameTaken (Some id) name current then false
    else applyWorkbookChange state (Workbook.renameSheetEverywhere id name)

let deleteSheet (state: SpreadsheetEditorState) (id: Workbook.SheetId) =
    applyWorkbookChange state (fun book ->
        match Workbook.lastVisibleSheetProblem id book with
        | Some _ -> book
        | None -> match Workbook.removeSheet id book with Some (changed, _) -> changed | None -> book)

let moveSheet (state: SpreadsheetEditorState) (id: Workbook.SheetId) (index: int) =
    applyWorkbookChange state (Workbook.moveSheet id index)
// ── Selection helpers ──────────────────────────────────────

let private clampRow (model: SpreadsheetModel) row = max 0 (min (model.RowCount - 1) row)
let private clampCol (model: SpreadsheetModel) col = max 0 (min (model.ColCount - 1) col)

/// Return the focus cell (row, col) of the current selection. Defaults
/// to (0, 0) for an empty range.
let focusCell (model: SpreadsheetModel) =
    match model.Selection with
    | SingleCell (r, c) -> r, c
    | Range (_, _, er, ec) -> er, ec

/// Snap a navigation target to the next valid landing cell, honoring
/// merged ranges. The merge primary (top-left) "owns" the whole range
/// visually; landing on a secondary cell would hide focus behind the
/// merge's fill, so arrow keys jump:
///   • forward (dr/dc > 0) → past the merge end
///   • backward (dr/dc < 0) → to the merge start
/// Excel does the same — once a range is merged, the merge acts as a
/// single logical cell during keyboard navigation.
let private resolveNavTarget (model: SpreadsheetModel) (row: int) (col: int) (dr: int) (dc: int) =
    match findMergedCell { Row = row; Col = col } model with
    | None -> row, col
    | Some merge ->
        let r =
            if dr > 0 then merge.EndRow + 1
            elif dr < 0 then merge.StartRow
            elif row = merge.StartRow then row
            else merge.StartRow
        let c =
            if dc > 0 then merge.EndCol + 1
            elif dc < 0 then merge.StartCol
            elif col = merge.StartCol then col
            else merge.StartCol
        r, c

/// Adjust ScrollX/Y so cell (row, col) sits inside the visible viewport.
/// Header strip (20 px tall) and row-header strip (28 px wide) stay
/// painted at the top/left of the canvas — we offset the visible band by
/// those amounts when checking whether the cell needs to be scrolled
/// into view. No-op when VisibleHeight/Width haven't been written yet
/// (first frame before scrollWidget runs).
let ensureCellVisible (state: SpreadsheetEditorState) (row: int) (col: int) =
    let model = activeModel state
    let viewW = state.VisibleWidth.Peek()
    let viewH = state.VisibleHeight.Peek()
    if viewW <= 0.0 || viewH <= 0.0 then () else
    let headerH = 20.0
    let rowHeaderW = 28.0
    let cellY = Layout.rowTop model row
    let cellH = getRowHeight row model
    let cellX = Layout.columnLeft model col
    let cellW = getColumnWidth col model
    let cellTopCanvas = headerH + cellY
    let cellBottomCanvas = cellTopCanvas + cellH
    let cellLeftCanvas = rowHeaderW + cellX
    let cellRightCanvas = cellLeftCanvas + cellW
    let sy = state.ScrollY.Peek()
    let sx = state.ScrollX.Peek()
    let newSy =
        if cellTopCanvas < sy + headerH then max 0.0 (cellTopCanvas - headerH)
        elif cellBottomCanvas > sy + viewH then cellBottomCanvas - viewH
        else sy
    let newSx =
        if cellLeftCanvas < sx + rowHeaderW then max 0.0 (cellLeftCanvas - rowHeaderW)
        elif cellRightCanvas > sx + viewW then cellRightCanvas - viewW
        else sx
    if newSy <> sy then state.ScrollY.Value <- newSy
    if newSx <> sx then state.ScrollX.Value <- newSx

let private setFocus (state: SpreadsheetEditorState) (row: int) (col: int) =
    let model = activeModel state
    let r0 = clampRow model row
    let c0 = clampCol model col
    // Snap to the merge primary when the target lands inside a merged range.
    // Arrow-key callers already route through `resolveNavTarget` so they
    // pass a non-secondary cell; this handles click/Home/End/etc.
    let r, c =
        match findMergedCell { Row = r0; Col = c0 } model with
        | Some merge -> merge.StartRow, merge.StartCol
        | None -> r0, c0
    let selection = SingleCell (r, c)
    if model.Selection <> selection then
        applyChangeNoHistory state (fun m -> { m with Selection = selection }) |> ignore
    ensureCellVisible state r c

let private extendFocus (state: SpreadsheetEditorState) (row: int) (col: int) =
    let model = activeModel state
    let r = clampRow model row
    let c = clampCol model col
    let anchor =
        match model.Selection with
        | SingleCell (ar, ac) -> ar, ac
        | Range (sr, sc, _, _) -> sr, sc
    let ar, ac = anchor
    let sel =
        if ar = r && ac = c then SingleCell (r, c)
        else Range (min ar r, min ac c, max ar r, max ac c)
    if model.Selection <> sel then
        applyChangeNoHistory state (fun m -> { m with Selection = sel }) |> ignore
    ensureCellVisible state r c

/// Number of rows that fit in the current visible viewport — used by
/// PageUp/PageDown. Walks heights from the current focus row outward
/// until cumulative height reaches the viewport.
let private pageRowJump (state: SpreadsheetEditorState) (fromRow: int) (direction: int) : int =
    let model = activeModel state
    let usableH = max 60.0 (state.VisibleHeight.Peek() - 20.0)
    let mutable rows = 0
    let mutable acc = 0.0
    let mutable r = fromRow
    while acc < usableH && r >= 0 && r < model.RowCount do
        acc <- acc + getRowHeight r model
        if acc < usableH then
            rows <- rows + 1
            r <- r + direction
    max 1 rows

// ── Cell content helpers ───────────────────────────────────

/// Cell text shown in the formula bar / inline editor. Formula cells
/// surface their `=expr` source; literal cells surface their typed value.
let cellEditText (model: SpreadsheetModel) row col =
    // What, typed back unchanged, gives the same cell: the formula's text, a
    // date as a date, forced text with its apostrophe.
    Rbt.Spreadsheet.Input.editTextFor model.Date1904 (getCell { Row = row; Col = col } model)

/// Typed text → cell value, read the way Excel reads input (shared with the
/// web editor: `Rbt.Spreadsheet.Input`). This host stores formulas without
/// the leading `=`.
let private parseCellInput (date1904: bool) (text: string) : CellValue =
    match (Rbt.Spreadsheet.Input.parseInputFor date1904 System.DateTime.UtcNow.Year text).Value with
    | Formula (formula, cached) when formula.StartsWith "=" -> Formula (formula.Substring 1, cached)
    | value -> value

/// The cell's format after an entry: an entry written as a percentage, a
/// currency amount, a date… gives a General cell the matching number format.
let private formatAfterInput (date1904: bool) (text: string) (format: CellFormat) : CellFormat =
    match (Rbt.Spreadsheet.Input.parseInputFor date1904 System.DateTime.UtcNow.Year text).NumberFormat, format.NumberFormat with
    | Some code, (None | Some "" | Some "General") -> { format with NumberFormat = Some code }
    | _ -> format

let private fireChange (selection: Selection) (model: SpreadsheetModel) =
    let address row col = sprintf "%s%d" (columnToLetter col) (row + 1)
    let target =
        match selection with
        | SingleCell (row, col) -> sprintf "Range(\"%s\")" (address row col)
        | Range (r0, c0, r1, c1) -> sprintf "Range(\"%s:%s\")" (address r0 c0) (address r1 c1)
    Rbt.Spreadsheet.Vba.fireEvent "Worksheet_Change" target model

let private writeCell (row: int) (col: int) (text: string) (model: SpreadsheetModel) =
    let value = parseCellInput model.Date1904 text
    let addr = { Row = row; Col = col }
    match validateEdit addr value model with
    | Result.Error _ -> model
    | Result.Ok () ->
        let edited =
            match value, getCell addr model with
            | Empty, _ -> clearCell addr model
            | v, Some existing ->
                setCell addr { existing with Value = v; Format = formatAfterInput model.Date1904 text existing.Format } model
            | v, None ->
                setCell addr { Value = v; Format = formatAfterInput model.Date1904 text defaultFormat; Comment = None } model
        edited

/// Run Worksheet_Change against freshly calculated values; macro writes are
/// included in the same history entry and recalculated once more if needed.
let private commitValueEdit (state: SpreadsheetEditorState) (before: Workbook.Workbook)
                            (edited: Workbook.Workbook) hint (selection: Selection) =
    let calculated = recalculate state edited hint
    let id = edited.ActiveSheet
    let model = calculated.Sheets.[id].Model
    let afterEvent = fireChange selection model
    let final =
        if obj.ReferenceEquals(model, afterEvent) then calculated
        else
            let withMacro = Workbook.updateSheetModelWithTables id afterEvent calculated
            recalculate state withMacro None
    installCalculatedEdit state before final

let private applyCellEdit (state: SpreadsheetEditorState) (row: int) (col: int) (text: string) =
    let before = state.Workbook.Value
    let id = before.ActiveSheet
    let oldModel = before.Sheets.[id].Model
    let edited = writeCell row col text oldModel
    if obj.ReferenceEquals(oldModel, edited) then false
    else
        let book = Workbook.updateSheetModelWithTables id edited before
        commitValueEdit state before book (Some [ id, { Row = row; Col = col } ]) (SingleCell (row, col))

/// Active cell reference like "B7" for the status bar / cell-ref box.
let activeCellRef (state: SpreadsheetEditorState) =
    let model = activeModel state
    let r, c = focusCell model
    sprintf "%s%d" (columnToLetter c) (r + 1)

/// Formula bar text: edit buffer when in edit mode, otherwise the cell
/// value's source text.
let formulaText (state: SpreadsheetEditorState) =
    match state.EditingCell.Value with
    | Some _ -> state.EditBuffer.Value
    | None ->
        let model = activeModel state
        let r, c = focusCell model
        cellEditText model r c

/// Selection summary for the status bar. "Cell B7 (12 cells filled)".
let selectionSummary (state: SpreadsheetEditorState) =
    let model = activeModel state
    let cellsFilled = Map.count model.Cells
    sprintf "Cell %s   %d filled" (activeCellRef state) cellsFilled

// ── Edit mode ──────────────────────────────────────────────

let startEditing (state: SpreadsheetEditorState) (row: int) (col: int) (seed: string option) =
    let model = activeModel state
    let r = clampRow model row
    let c = clampCol model col
    let initial = seed |> Option.defaultWith (fun () -> cellEditText model r c)
    state.EditingCell.Value <- Some (r, c)
    state.EditBuffer.Value <- initial
    setFocus state r c
    Platform.requestRedraw()

let commitEdit (state: SpreadsheetEditorState) =
    match state.EditingCell.Value with
    | Some (r, c) ->
        let text = state.EditBuffer.Value
        applyCellEdit state r c text |> ignore
        state.EditingCell.Value <- None
        state.EditBuffer.Value <- ""
    | None -> ()

let cancelEdit (state: SpreadsheetEditorState) =
    state.EditingCell.Value <- None
    state.EditBuffer.Value <- ""
    Platform.requestRedraw()

let editFormula (state: SpreadsheetEditorState) (text: string) =
    match state.EditingCell.Value with
    | Some _ -> state.EditBuffer.Value <- text
    | None ->
        // Click into the formula bar with no active edit — start one on
        // the current focus cell and seed it with the typed text.
        let model = activeModel state
        let r, c = focusCell model
        state.EditingCell.Value <- Some (r, c)
        state.EditBuffer.Value <- text

let submitFormula (state: SpreadsheetEditorState) =
    commitEdit state
    // After commit, advance to the next row (Excel's default Enter behavior).
    let model = activeModel state
    let r, c = focusCell model
    setFocus state (r + 1) c

// ── Clipboard ──────────────────────────────────────────────

let private clipboardText (data: ClipboardData) (model: SpreadsheetModel) =
    let cells = data.Cells |> List.map (fun (r, c, cell) ->
        (r, c), Rbt.Spreadsheet.Input.editTextFor model.Date1904 (Some cell)) |> Map.ofList
    [ for row in 0 .. data.SourceEndRow - data.SourceStartRow ->
        [ for col in 0 .. data.SourceEndCol - data.SourceStartCol ->
            Map.tryFind (row, col) cells |> Option.defaultValue "" ] |> String.concat "\t" ]
    |> String.concat "\n"

let private clearSelected (state: SpreadsheetEditorState) =
    let before = state.Workbook.Value
    let id = before.ActiveSheet
    let model = before.Sheets.[id].Model
    let cleared = clearSelectionCells model.Selection model
    if obj.ReferenceEquals(model, cleared) then false
    else commitValueEdit state before (Workbook.updateSheetModelWithTables id cleared before) None model.Selection

let copySelection (state: SpreadsheetEditorState) =
    let model = activeModel state
    let data = serializeSelection model.Selection model false
    let text = clipboardText data model
    state.Clipboard.Value <- Some (sheetId state, data, text)
    state.InvalidatedClipboardText.Value <- None
    Platform.clipboard.SetText text
    text

let cutSelection (state: SpreadsheetEditorState) =
    let model = activeModel state
    let data = serializeSelection model.Selection model true
    let text = clipboardText data model
    state.Clipboard.Value <- Some (sheetId state, data, text)
    state.InvalidatedClipboardText.Value <- None
    Platform.clipboard.SetText text
    clearSelected state |> ignore
    text

let private pasteInternal (state: SpreadsheetEditorState) (source: Workbook.SheetId) (data: ClipboardData) =
    let before = state.Workbook.Value
    if not (Map.containsKey source before.Sheets) then
        state.InvalidatedClipboardText.Value <- (state.Clipboard.Value |> Option.map (fun (_, _, text) -> text))
        state.Clipboard.Value <- None
        false
    else
        let id = before.ActiveSheet
        let oldModel = before.Sheets.[id].Model
        let anchors = pasteAnchors oldModel.Selection data
        let pasted = pasteClipboardIntoSelection oldModel.Selection data oldModel
        let selection = pasteSelection anchors data oldModel.Selection
        if obj.ReferenceEquals(oldModel, pasted) then false
        else
            let book = Workbook.updateSheetModelWithTables id { pasted with Selection = selection } before
            let book =
                match anchors with
                | [ anchor ] when data.IsCut ->
                    let area = data.SourceStartRow, data.SourceStartCol, data.SourceEndRow, data.SourceEndCol
                    let delta = anchor.Row - data.SourceStartRow, anchor.Col - data.SourceStartCol
                    Workbook.repointMovedCells source area id delta book |> fst
                | _ -> book
            let committed = commitValueEdit state before book None selection
            if committed && data.IsCut then state.Clipboard.Value <- None
            committed

let pasteFromClipboard (state: SpreadsheetEditorState) =
    let systemText = Platform.clipboard.GetText()
    match state.Clipboard.Value, systemText with
    | Some (source, data, ownText), text when text.IsNone || text = Some ownText -> pasteInternal state source data
    | _, Some text when state.InvalidatedClipboardText.Value = Some text -> false
    | _, Some text when text <> "" ->
        let r, c = focusCell (activeModel state)
        applyCellEdit state r c text
    | _ -> false

// ── Hit testing ────────────────────────────────────────────

let private viewportOf (state: SpreadsheetEditorState) width height : Layout.Viewport =
    { ScrollX = state.ScrollX.Value
      ScrollY = state.ScrollY.Value
      Width = width
      Height = height
      HeaderHeight = 20.0
      RowHeaderWidth = 28.0
      Measurer = None }

let private hitCell (state: SpreadsheetEditorState) width height (x: float) (y: float) =
    let model = activeModel state
    let viewport = viewportOf state width height
    match Layout.hitTest model viewport x y with
    | Layout.Cell addr -> Some (addr.Row, addr.Col)
    | Layout.HitRowHeader row -> Some (row, 0)
    | Layout.HitColumnHeader col -> Some (0, col)
    | _ -> None

// ── Key handling ───────────────────────────────────────────

let handleKey (state: SpreadsheetEditorState) (ev: KeyEvent) : bool =
    let editing = state.EditingCell.Value.IsSome
    match ev with
    | { Key = "Escape" } when editing ->
        cancelEdit state
        true
    | { Key = "Enter" } when editing ->
        submitFormula state
        true
    | { Key = "Tab"; Shift = false } when editing ->
        commitEdit state
        let r, c = focusCell (activeModel state)
        setFocus state r (c + 1)
        true
    | { Key = "Tab"; Shift = true } when editing ->
        commitEdit state
        let r, c = focusCell (activeModel state)
        setFocus state r (c - 1)
        true
    | _ when editing ->
        // While editing, route printable keys to the edit buffer and
        // arrows to plain cursor moves within the buffer (TODO). For v1
        // we let printables append and arrows fall through unhandled.
        match ev with
        | { Key = "Backspace" } ->
            let buf = state.EditBuffer.Value
            if buf.Length > 0 then state.EditBuffer.Value <- buf.Substring(0, buf.Length - 1)
            true
        | { Ctrl = false; Alt = false; Key = key } when key.Length = 1 ->
            state.EditBuffer.Value <- state.EditBuffer.Value + key
            true
        | _ -> false
    | { Ctrl = true; Key = key } when key = "z" || key = "Z" -> undo state
    | { Ctrl = true; Key = key } when key = "y" || key = "Y" -> redo state
    | { Ctrl = true; Shift = true; Key = key } when key = "z" || key = "Z" -> redo state
    | { Ctrl = true; Key = key } when key = "c" || key = "C" ->
        let _ = copySelection state
        true
    | { Ctrl = true; Key = key } when key = "x" || key = "X" ->
        let _ = cutSelection state
        true
    | { Ctrl = true; Key = key } when key = "v" || key = "V" ->
        let _ = pasteFromClipboard state
        true
    | { Ctrl = true; Key = "Home" } ->
        setFocus state 0 0
        true
    | { Ctrl = true; Key = "End" } ->
        let model = activeModel state
        let maxRow, maxCol =
            usedBounds model |> Option.defaultValue (0, 0)
        setFocus state maxRow maxCol
        true
    | { Key = "Home"; Shift = false } ->
        let r, _ = focusCell (activeModel state)
        setFocus state r 0
        true
    | { Key = "End"; Shift = false } ->
        let model = activeModel state
        let r, _ = focusCell model
        let mutable lastCol = 0
        for KeyValue(addr, _) in model.Cells do
            if addr.Row = r && addr.Col > lastCol then
                lastCol <- addr.Col
        setFocus state r lastCol
        true
    | { Key = "PageDown"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        let jump = pageRowJump state r 1
        setFocus state (r + jump) c
        true
    | { Key = "PageUp"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        let jump = pageRowJump state r -1
        setFocus state (r - jump) c
        true
    | { Key = "PageDown"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        let jump = pageRowJump state r 1
        extendFocus state (r + jump) c
        true
    | { Key = "PageUp"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        let jump = pageRowJump state r -1
        extendFocus state (r - jump) c
        true
    | { Key = "ArrowLeft"; Shift = false } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c - 1) 0 -1
        setFocus state nr nc
        true
    | { Key = "ArrowRight"; Shift = false } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c + 1) 0 1
        setFocus state nr nc
        true
    | { Key = "ArrowUp"; Shift = false } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r - 1) c -1 0
        setFocus state nr nc
        true
    | { Key = "ArrowDown"; Shift = false } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r + 1) c 1 0
        setFocus state nr nc
        true
    | { Key = "ArrowLeft"; Shift = true } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c - 1) 0 -1
        extendFocus state nr nc
        true
    | { Key = "ArrowRight"; Shift = true } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c + 1) 0 1
        extendFocus state nr nc
        true
    | { Key = "ArrowUp"; Shift = true } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r - 1) c -1 0
        extendFocus state nr nc
        true
    | { Key = "ArrowDown"; Shift = true } ->
        let model = activeModel state
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r + 1) c 1 0
        extendFocus state nr nc
        true
    | { Key = "Enter" } ->
        let r, c = focusCell (activeModel state)
        setFocus state (r + 1) c
        true
    | { Key = "Tab"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        setFocus state r (c + 1)
        true
    | { Key = "Tab"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        setFocus state r (c - 1)
        true
    | { Key = "F2" } ->
        let r, c = focusCell (activeModel state)
        startEditing state r c None
        true
    | { Key = "Delete" } | { Key = "Backspace" } ->
        clearSelected state |> ignore
        true
    | { Ctrl = false; Alt = false; Key = key } when key.Length = 1 ->
        // Printable key with no edit in flight: start a new edit and
        // seed the buffer with the typed character (Excel's behavior).
        let r, c = focusCell (activeModel state)
        startEditing state r c (Some key)
        true
    | _ -> false

// ── Pointer handling ───────────────────────────────────────

let private handlePointer (state: SpreadsheetEditorState) width height = function
    | PointerDown (x, y) ->
        Focus.request state.FocusId
        // If the user was mid-edit, commit it before moving the selection.
        if state.EditingCell.Value.IsSome then commitEdit state
        match hitCell state width height x y with
        | Some (row, col) ->
            state.DragAnchor.Value <- Some (row, col)
            setFocus state row col
        | None -> ()
    | PointerMove (x, y) ->
        match state.DragAnchor.Value with
        | Some (ar, ac) ->
            match hitCell state width height x y with
            | Some (row, col) when (row, col) <> (ar, ac) ->
                extendFocus state row col
            | _ -> ()
        | None -> ()
    | PointerUp _ ->
        state.DragAnchor.Value <- None

// ── Rendering ──────────────────────────────────────────────

let private gridCommands (state: SpreadsheetEditorState) width height =
    let model = activeModel state
    SpreadsheetGrid.widget width height model

/// Build the interactive grid widget. The widget intentionally does NOT
/// scroll itself — wrap it with `scrollWidget` to get the standard
/// scroll-view behavior.
let widget (width: float) (height: float) (state: SpreadsheetEditorState) =
    dynamic (fun () ->
        let grid = gridCommands state width height
        let pointer = handlePointer state width height
        let keys = handleKey state
        Focusable(state.FocusId,
            onKey keys (
                pointerRegion pointer grid)))

/// Wrap the editor in a horizontal+vertical scroll view that tracks
/// state.ScrollX/Y.
let scrollWidget (width: float) (height: float) (state: SpreadsheetEditorState) =
    // Stash the visible viewport dimensions so key handlers (PageUp/Down,
    // ensureCellVisible after arrow keys) can read them. Done outside the
    // `dynamic` builder because the builder may re-run for other reasons
    // and we want the geometry to follow the actual passed dims.
    if state.VisibleWidth.Peek() <> width then state.VisibleWidth.Value <- width
    if state.VisibleHeight.Peek() <> height then state.VisibleHeight.Value <- height
    dynamic (fun () ->
        let model = activeModel state
        // Size the scrollable canvas to the data range plus a viewport-covering
        // margin, not the full 1,048,576 × 16,384 virtual grid. The model is
        // sized to Excel's real bounds so visibleRange can virtualize, but
        // summing every row/col here on each frame freezes the first paint.
        let dataMaxRow, dataMaxCol =
            usedBounds model |> Option.defaultValue (-1, -1)
        let rowBudget = min (max 199 (dataMaxRow + 200)) (model.RowCount - 1)
        let colBudget = min (max 49 (dataMaxCol + 50)) (model.ColCount - 1)
        let contentWidth = 28.0 + Layout.columnLeft model (colBudget + 1)
        let contentHeight = 20.0 + Layout.rowTop model (rowBudget + 1)
        let inner =
            SpreadsheetGrid.viewportWidget
                (max width contentWidth)
                (max height contentHeight)
                width
                height
                state.ScrollX.Value
                state.ScrollY.Value
                model
            |> pointerRegion (handlePointer state width height)
            |> onKey (handleKey state)
            |> fun grid -> Focusable(state.FocusId, grid)
        scrollWith true true state.ScrollX state.ScrollY inner)
