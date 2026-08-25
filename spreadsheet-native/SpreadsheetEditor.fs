module Composit.Ui.SpreadsheetEditor

open Renderer
open Composit.Ui.Reactive
open Composit.Ui.Widget

open RedBat.Spreadsheet.Model

module Layout = RedBat.Spreadsheet.Layout
module SpreadsheetGrid = Composit.Ui.SpreadsheetGrid

let inline private px (v: float) : float<px> = LanguagePrimitives.FloatWithMeasure v

/// Interactive editor state. Holds the full workbook (sheets array +
/// names + active index) plus per-app UI state: selection drag anchor,
/// in-cell edit buffer, scroll position, formula bar text, undo/redo
/// snapshots. Mirrors the structure of Composit.Ui.WordEditor.
/// SpreadsheetModel can't satisfy the `equality` constraint because its
/// History.Past/Future hold Command records with function-valued Undo/Redo
/// fields. RefSignal sidesteps the constraint — we re-render on every
/// assignment regardless of whether the new value is structurally equal,
/// which is fine for the editor since assignments are explicit.
type SpreadsheetEditorState = {
    Sheets: RefSignal<SpreadsheetModel array>
    SheetNames: Signal<string array>
    ActiveSheet: Signal<int>
    ScrollX: Signal<float>
    ScrollY: Signal<float>
    /// Visible viewport size in CSS pixels — written by `scrollWidget` on
    /// every layout pass so key handlers (PageUp/PageDown, ensure-visible
    /// after arrow moves) know how big the visible window is.
    VisibleWidth: Signal<float>
    VisibleHeight: Signal<float>
    /// (row, col) of the cell currently being edited. None when no cell
    /// is in edit mode. Set by double-click, F2, or a printable key.
    EditingCell: Signal<(int * int) option>
    /// Live edit-buffer text for the cell in EditingCell. Reused as the
    /// formula bar contents when the user clicks into the cell ref box.
    EditBuffer: Signal<string>
    /// Drag-select anchor: the (row, col) where the pointer went down,
    /// used to compute Range selections on PointerMove.
    DragAnchor: Signal<(int * int) option>
    UndoStack: RefSignal<(SpreadsheetModel array * int) list>
    RedoStack: RefSignal<(SpreadsheetModel array * int) list>
    CalcEngine: RefSignal<RedBat.Spreadsheet.Calc.CalcEngine option>
    FocusId: int
}

let createState (sheets: SpreadsheetModel array) (names: string array) (activeIndex: int) =
    let idx = max 0 (min (sheets.Length - 1) activeIndex)
    {
        Sheets = refSignal sheets
        SheetNames = signal names
        ActiveSheet = signal idx
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
        FocusId = Focus.allocId()
    }

/// Read the currently-active sheet's model.
let activeModel (state: SpreadsheetEditorState) : SpreadsheetModel =
    let sheets = state.Sheets.Value
    let idx = state.ActiveSheet.Value
    if idx >= 0 && idx < sheets.Length then sheets.[idx]
    else sheets.[0]

let activeSheetName (state: SpreadsheetEditorState) : string =
    let names = state.SheetNames.Value
    let idx = state.ActiveSheet.Value
    if idx >= 0 && idx < names.Length then names.[idx] else "Sheet1"

/// Replace the active sheet's model. Does not push undo; callers that
/// want history should go through applyChange.
let private putActiveModel (state: SpreadsheetEditorState) (model: SpreadsheetModel) =
    let sheets = state.Sheets.Value |> Array.copy
    let idx = state.ActiveSheet.Value
    if idx >= 0 && idx < sheets.Length then
        sheets.[idx] <- model
        state.Sheets.Value <- sheets

/// Push the current workbook snapshot onto the undo stack, then apply
/// the transform to the active sheet. Returns true when the change
/// actually altered the model.
let applyChange (state: SpreadsheetEditorState) (transform: SpreadsheetModel -> SpreadsheetModel) =
    state.CalcEngine.Value <- None
    let before = activeModel state
    let after = transform before
    if System.Object.ReferenceEquals(before, after) then false
    else
        state.UndoStack.Value <- (state.Sheets.Value, state.ActiveSheet.Value) :: state.UndoStack.Value
        state.RedoStack.Value <- []
        putActiveModel state after
        Platform.requestRedraw()
        true

/// Apply a transform without recording undo — used for selection and
/// scroll-only changes that should not eat into the user's undo budget.
let applyChangeNoHistory (state: SpreadsheetEditorState) (transform: SpreadsheetModel -> SpreadsheetModel) =
    let before = activeModel state
    let after = transform before
    if System.Object.ReferenceEquals(before, after) then false
    else
        putActiveModel state after
        Platform.requestRedraw()
        true

let undo (state: SpreadsheetEditorState) =
    match state.UndoStack.Value with
    | (sheets, idx) :: rest ->
        state.CalcEngine.Value <- None
        state.RedoStack.Value <- (state.Sheets.Value, state.ActiveSheet.Value) :: state.RedoStack.Value
        state.UndoStack.Value <- rest
        state.Sheets.Value <- sheets
        state.ActiveSheet.Value <- idx
        Platform.requestRedraw()
        true
    | [] -> true

let redo (state: SpreadsheetEditorState) =
    match state.RedoStack.Value with
    | (sheets, idx) :: rest ->
        state.CalcEngine.Value <- None
        state.UndoStack.Value <- (state.Sheets.Value, state.ActiveSheet.Value) :: state.UndoStack.Value
        state.RedoStack.Value <- rest
        state.Sheets.Value <- sheets
        state.ActiveSheet.Value <- idx
        Platform.requestRedraw()
        true
    | [] -> true

let switchSheet (state: SpreadsheetEditorState) (index: int) =
    let sheets = state.Sheets.Value
    if index >= 0 && index < sheets.Length && index <> state.ActiveSheet.Value then
        state.CalcEngine.Value <- None
        state.ActiveSheet.Value <- index
        state.ScrollX.Value <- 0.0
        state.ScrollY.Value <- 0.0
        state.EditingCell.Value <- None
        state.EditBuffer.Value <- ""
        Platform.requestRedraw()

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
    match getCell { Row = row; Col = col } model with
    | None -> ""
    | Some data ->
        match data.Value with
        | Number n -> string n
        | String s -> s
        | Boolean b -> if b then "TRUE" else "FALSE"
        | Formula (expr, _) -> "=" + expr
        | Empty -> ""
        | Error e -> "#" + e

let private parseCellInput (text: string) : CellValue =
    let trimmed = text.Trim()
    if trimmed = "" then Empty
    elif trimmed.StartsWith "=" then Formula (trimmed.Substring(1), Empty)
    elif System.String.Equals(trimmed, "TRUE", System.StringComparison.OrdinalIgnoreCase) then Boolean true
    elif System.String.Equals(trimmed, "FALSE", System.StringComparison.OrdinalIgnoreCase) then Boolean false
    else
        match System.Double.TryParse trimmed with
        | true, n -> Number n
        | _ -> String trimmed

let private singleSheetWorkbook (name: string) (model: SpreadsheetModel) : RedBat.Spreadsheet.Workbook.Workbook =
    let sheet : RedBat.Spreadsheet.Workbook.Sheet =
        { Id = "sheet1"; Name = name; Model = model; Visible = true }
    { Sheets = Map.ofList [ "sheet1", sheet ]
      SheetOrder = [ "sheet1" ]
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names = Map.empty }

let private recalcEditedCell (state: SpreadsheetEditorState) (addr: CellAddress) (cellData: CellData) (edited: SpreadsheetModel) (before: SpreadsheetModel) =
    if state.Sheets.Value.Length <> 1 || cellData.Value = CellValue.Empty then
        state.CalcEngine.Value <- None
        RedBat.Spreadsheet.Formula.recalcAndSpill None None edited
    else
        let wb = singleSheetWorkbook (activeSheetName state) before
        let sheetCell : RedBat.Spreadsheet.Workbook.SheetCell =
            { Sheet = wb.ActiveSheet; Cell = addr }
        match RedBat.Spreadsheet.Calc.trySingleCellEdit sheetCell cellData wb state.CalcEngine.Value with
        | RedBat.Spreadsheet.Calc.IncrementalApplied (updatedWorkbook, updatedEngine) ->
            state.CalcEngine.Value <- Some updatedEngine
            match Map.tryFind updatedWorkbook.ActiveSheet updatedWorkbook.Sheets with
            | Some sheet -> sheet.Model
            | None -> edited
        | RedBat.Spreadsheet.Calc.IncrementalFallback _ ->
            state.CalcEngine.Value <- None
            RedBat.Spreadsheet.Formula.recalcAndSpill None None edited

let private writeCell (state: SpreadsheetEditorState) (row: int) (col: int) (text: string) (model: SpreadsheetModel) =
    let value = parseCellInput text
    let addr = { Row = row; Col = col }
    // Reject edits that violate the cell's data-validation rule (keep old value).
    match validateEdit addr value model with
    | Result.Error _ -> model
    | Result.Ok() ->
        let edited =
            match value, getCell addr model with
            | Empty, _ -> clearCell addr model
            | v, Some existing -> setCell addr { existing with Value = v } model
            | v, None -> setCell addr { Value = v; Format = defaultFormat; Comment = None } model
        if System.Object.ReferenceEquals(model, edited) then model
        else
            // Recompute formulas and (re)materialize dynamic-array spills so the edit
            // is live — a spilling formula (FILTER/SORT/GROUPBY/PIVOTBY) refreshes when
            // its source changes. (Same-sheet scope; cross-sheet recalc is separate.)
            let cellData =
                match value, getCell addr model with
                | Empty, _ -> { Value = Empty; Format = defaultFormat; Comment = None }
                | v, Some existing -> { existing with Value = v }
                | v, None -> { Value = v; Format = defaultFormat; Comment = None }
            let recomputed = recalcEditedCell state addr cellData edited model
            // Fire the workbook's Worksheet_Change handler (if any), passing the edited
            // cell as Target. A no-op unless the macro module defines the sub; its
            // writes apply directly, so this can't recurse back through writeCell.
            let target = sprintf "Range(\"%s%d\")" (columnToLetter col) (row + 1)
            let afterEvent = RedBat.Spreadsheet.Vba.fireEvent "Worksheet_Change" target recomputed
            if not (System.Object.ReferenceEquals(recomputed, afterEvent)) then
                state.CalcEngine.Value <- None
            afterEvent

let private applyCellEdit (state: SpreadsheetEditorState) (row: int) (col: int) (text: string) =
    let before = activeModel state
    let after = writeCell state row col text before
    if System.Object.ReferenceEquals(before, after) then false
    else
        state.UndoStack.Value <- (state.Sheets.Value, state.ActiveSheet.Value) :: state.UndoStack.Value
        state.RedoStack.Value <- []
        putActiveModel state after
        Platform.requestRedraw()
        true

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

let copySelection (state: SpreadsheetEditorState) =
    let model = activeModel state
    let r, c = focusCell model
    let text = cellEditText model r c
    if text <> "" then Platform.clipboard.SetText text
    text

let cutSelection (state: SpreadsheetEditorState) =
    let text = copySelection state
    if text <> "" then
        let r, c = focusCell (activeModel state)
        applyChange state (clearCell { Row = r; Col = c }) |> ignore
    text

let pasteFromClipboard (state: SpreadsheetEditorState) =
    match Platform.clipboard.GetText() with
    | Some text when text <> "" ->
        let r, c = focusCell (activeModel state)
        applyCellEdit state r c text |> ignore
        true
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

let private handleKey (state: SpreadsheetEditorState) (ev: KeyEvent) : bool =
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
        let r, c = focusCell (activeModel state)
        applyChange state (clearCell { Row = r; Col = c }) |> ignore
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
