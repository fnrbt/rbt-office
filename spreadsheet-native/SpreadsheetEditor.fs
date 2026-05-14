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
        EditingCell = signal None
        EditBuffer = signal ""
        DragAnchor = signal None
        UndoStack = refSignal []
        RedoStack = refSignal []
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

let private setFocus (state: SpreadsheetEditorState) (row: int) (col: int) =
    let model = activeModel state
    let r = clampRow model row
    let c = clampCol model col
    applyChangeNoHistory state (fun m -> { m with Selection = SingleCell (r, c) }) |> ignore

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
    applyChangeNoHistory state (fun m -> { m with Selection = sel }) |> ignore

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

let private writeCell (row: int) (col: int) (text: string) (model: SpreadsheetModel) =
    let value = parseCellInput text
    let addr = { Row = row; Col = col }
    match value, getCell addr model with
    | Empty, _ -> clearCell addr model
    | v, Some existing -> setCell addr { existing with Value = v } model
    | v, None -> setCell addr { Value = v; Format = defaultFormat; Comment = None } model

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
        applyChange state (writeCell r c text) |> ignore
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
        applyChange state (writeCell r c text) |> ignore
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
    | { Key = "ArrowLeft"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        setFocus state r (c - 1)
        true
    | { Key = "ArrowRight"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        setFocus state r (c + 1)
        true
    | { Key = "ArrowUp"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        setFocus state (r - 1) c
        true
    | { Key = "ArrowDown"; Shift = false } ->
        let r, c = focusCell (activeModel state)
        setFocus state (r + 1) c
        true
    | { Key = "ArrowLeft"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        extendFocus state r (c - 1)
        true
    | { Key = "ArrowRight"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        extendFocus state r (c + 1)
        true
    | { Key = "ArrowUp"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        extendFocus state (r - 1) c
        true
    | { Key = "ArrowDown"; Shift = true } ->
        let r, c = focusCell (activeModel state)
        extendFocus state (r + 1) c
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
    let scrolled = { model with Cells = model.Cells }
    // SpreadsheetGrid.widget reads the model's selection directly, so we
    // hand the active model in as-is; scrolling is applied by the
    // parent ScrollView wrapper.
    SpreadsheetGrid.widget width height scrolled

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
    dynamic (fun () ->
        let model = activeModel state
        let contentWidth =
            [ 0 .. model.ColCount - 1 ]
            |> List.sumBy (fun c -> getColumnWidth c model)
            |> (+) 28.0
        let contentHeight =
            [ 0 .. model.RowCount - 1 ]
            |> List.sumBy (fun r -> getRowHeight r model)
            |> (+) 20.0
        let inner = widget (max width contentWidth) (max height contentHeight) state
        scrollWith true true state.ScrollX state.ScrollY inner)
