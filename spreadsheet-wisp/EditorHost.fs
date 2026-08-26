module Rbt.SpreadsheetWisp.EditorHost

open System
open Wisp.Reactive
open Wisp.Layout

module M = Rbt.Spreadsheet.Model
module L = Rbt.Spreadsheet.Layout
module Calc = Rbt.Spreadsheet.Calc
module Formula = Rbt.Spreadsheet.Formula
module Workbook = Rbt.Spreadsheet.Workbook
module GR = Rbt.SpreadsheetWisp.GridRender

type KeyEvent = { Key: string; Shift: bool; Ctrl: bool; Alt: bool }

module Clip =
    let mutable Data : M.ClipboardData option = None
    let mutable Text = ""

type EditorState =
    { Sheets: RefSignal<M.SpreadsheetModel array>
      SheetNames: RefSignal<string array>
      ActiveSheet: Signal<int>
      ScrollX: Signal<float>
      ScrollY: Signal<float>
      VisibleWidth: Signal<float>
      VisibleHeight: Signal<float>
      EditingCell: Signal<(int * int) option>
      EditBuffer: Signal<string>
      DragAnchor: Signal<(int * int) option>
      UndoStack: RefSignal<(M.SpreadsheetModel array * int) list>
      RedoStack: RefSignal<(M.SpreadsheetModel array * int) list>
      CalcEngine: RefSignal<Calc.CalcEngine option>
      mutable LastRecalcWasIncremental: bool }

let create (sheets: M.SpreadsheetModel array) (names: string array) (activeIndex: int) =
    let safeSheets = if sheets.Length = 0 then [| M.createEmpty 100 26 |] else sheets
    let safeNames =
        if names.Length = safeSheets.Length then names
        else [| for i in 0 .. safeSheets.Length - 1 -> sprintf "Sheet%d" (i + 1) |]
    let idx = max 0 (min (safeSheets.Length - 1) activeIndex)
    { Sheets = refSignal safeSheets
      SheetNames = refSignal safeNames
      ActiveSheet = signal idx
      ScrollX = signal 0.0
      ScrollY = signal 0.0
      VisibleWidth = signal 900.0
      VisibleHeight = signal 520.0
      EditingCell = signal None
      EditBuffer = signal ""
      DragAnchor = signal None
      UndoStack = refSignal []
      RedoStack = refSignal []
      CalcEngine = refSignal None
      LastRecalcWasIncremental = false }

let activeModel (st: EditorState) : M.SpreadsheetModel =
    let sheets = st.Sheets.Value
    let idx = st.ActiveSheet.Value
    if idx >= 0 && idx < sheets.Length then sheets.[idx] else sheets.[0]

let activeSheetName (st: EditorState) =
    let names = st.SheetNames.Value
    let idx = st.ActiveSheet.Value
    if idx >= 0 && idx < names.Length then names.[idx] else "Sheet1"

let private putActiveModel (st: EditorState) (model: M.SpreadsheetModel) =
    let sheets = st.Sheets.Value |> Array.copy
    let idx = st.ActiveSheet.Value
    if idx >= 0 && idx < sheets.Length then
        sheets.[idx] <- model
        st.Sheets.Value <- sheets

let private invalidateCalc (st: EditorState) =
    st.CalcEngine.Value <- None
    st.LastRecalcWasIncremental <- false

let applyNoHistory (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    let before = activeModel st
    let after = transform before
    if not (Object.ReferenceEquals(before, after)) then putActiveModel st after
    true

let applyChange (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    invalidateCalc st
    let before = activeModel st
    let after = transform before
    if not (Object.ReferenceEquals(before, after)) then
        st.UndoStack.Value <- (st.Sheets.Value, st.ActiveSheet.Value) :: st.UndoStack.Value |> List.truncate 50
        st.RedoStack.Value <- []
        putActiveModel st after
    true

let undo (st: EditorState) =
    match st.UndoStack.Value with
    | (sheets, idx) :: rest ->
        invalidateCalc st
        st.RedoStack.Value <- (st.Sheets.Value, st.ActiveSheet.Value) :: st.RedoStack.Value |> List.truncate 50
        st.UndoStack.Value <- rest
        st.Sheets.Value <- sheets
        st.ActiveSheet.Value <- idx
        true
    | [] -> true

let redo (st: EditorState) =
    match st.RedoStack.Value with
    | (sheets, idx) :: rest ->
        invalidateCalc st
        st.UndoStack.Value <- (st.Sheets.Value, st.ActiveSheet.Value) :: st.UndoStack.Value |> List.truncate 50
        st.RedoStack.Value <- rest
        st.Sheets.Value <- sheets
        st.ActiveSheet.Value <- idx
        true
    | [] -> true

let switchSheet (st: EditorState) index =
    if index >= 0 && index < st.Sheets.Value.Length && index <> st.ActiveSheet.Value then
        invalidateCalc st
        st.ActiveSheet.Value <- index
        st.ScrollX.Value <- 0.0
        st.ScrollY.Value <- 0.0
        st.EditingCell.Value <- None
        st.EditBuffer.Value <- ""

let private clampRow (model: M.SpreadsheetModel) row = max 0 (min (model.RowCount - 1) row)
let private clampCol (model: M.SpreadsheetModel) col = max 0 (min (model.ColCount - 1) col)

let focusCell (model: M.SpreadsheetModel) =
    match model.Selection with
    | M.SingleCell (r, c) -> r, c
    | M.Range (_, _, er, ec) -> er, ec

let private resolveNavTarget (model: M.SpreadsheetModel) row col dr dc =
    match M.findMergedCell { Row = row; Col = col } model with
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

let ensureCellVisible (st: EditorState) row col =
    let model = activeModel st
    let viewW = st.VisibleWidth.Peek()
    let viewH = st.VisibleHeight.Peek()
    if viewW > 0.0 && viewH > 0.0 then
        let cellY = L.rowTop model row
        let cellH = M.getRowHeight row model
        let cellX = L.columnLeft model col
        let cellW = M.getColumnWidth col model
        let sy = st.ScrollY.Peek()
        let sx = st.ScrollX.Peek()
        let headerH = 20.0
        let rowHeaderW = 28.0
        let newSy =
            if headerH + cellY < sy + headerH then max 0.0 cellY
            elif headerH + cellY + cellH > sy + viewH then max 0.0 (headerH + cellY + cellH - viewH)
            else sy
        let newSx =
            if rowHeaderW + cellX < sx + rowHeaderW then max 0.0 cellX
            elif rowHeaderW + cellX + cellW > sx + viewW then max 0.0 (rowHeaderW + cellX + cellW - viewW)
            else sx
        if newSy <> sy then st.ScrollY.Value <- newSy
        if newSx <> sx then st.ScrollX.Value <- newSx

let setFocus (st: EditorState) row col =
    let model = activeModel st
    let r0 = clampRow model row
    let c0 = clampCol model col
    let r, c =
        match M.findMergedCell { Row = r0; Col = c0 } model with
        | Some merge -> merge.StartRow, merge.StartCol
        | None -> r0, c0
    applyNoHistory st (fun m -> { m with Selection = M.SingleCell (r, c) }) |> ignore
    ensureCellVisible st r c

let extendFocus (st: EditorState) row col =
    let model = activeModel st
    let r = clampRow model row
    let c = clampCol model col
    let ar, ac =
        match model.Selection with
        | M.SingleCell (sr, sc) -> sr, sc
        | M.Range (sr, sc, _, _) -> sr, sc
    let sel =
        if ar = r && ac = c then M.SingleCell (r, c)
        else M.Range (min ar r, min ac c, max ar r, max ac c)
    applyNoHistory st (fun m -> { m with Selection = sel }) |> ignore
    ensureCellVisible st r c

let cellEditText (model: M.SpreadsheetModel) row col =
    match M.getCell { Row = row; Col = col } model with
    | None -> ""
    | Some data ->
        match data.Value with
        | M.Number n -> n.ToString("G", Globalization.CultureInfo.InvariantCulture)
        | M.String s -> s
        | M.Boolean b -> if b then "TRUE" else "FALSE"
        | M.Formula (expr, _) -> "=" + expr
        | M.Empty -> ""
        | M.Error e -> if e.StartsWith "#" then e else "#" + e

let parseCellInput (text: string) : M.CellValue =
    let trimmed = text.Trim()
    if trimmed = "" then M.Empty
    elif trimmed.StartsWith "=" then M.Formula(trimmed.Substring 1, M.Empty)
    elif String.Equals(trimmed, "TRUE", StringComparison.OrdinalIgnoreCase) then M.Boolean true
    elif String.Equals(trimmed, "FALSE", StringComparison.OrdinalIgnoreCase) then M.Boolean false
    else
        match Double.TryParse(trimmed, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, n -> M.Number n
        | _ -> M.String trimmed

let private singleSheetWorkbook (name: string) (model: M.SpreadsheetModel) : Workbook.Workbook =
    let sheet : Workbook.Sheet = { Id = "sheet1"; Name = name; Model = model; Visible = true }
    { Sheets = Map.ofList [ "sheet1", sheet ]
      SheetOrder = [ "sheet1" ]
      ActiveSheet = "sheet1"
      Tables = Map.empty
      Names = Map.empty }

let private recalcEditedCell (st: EditorState) (addr: M.CellAddress) (cellData: M.CellData) (edited: M.SpreadsheetModel) (before: M.SpreadsheetModel) =
    if st.Sheets.Value.Length <> 1 || cellData.Value = M.Empty then
        invalidateCalc st
        Formula.recalcAndSpill None None edited
    else
        let wb = singleSheetWorkbook (activeSheetName st) before
        let sheetCell : Workbook.SheetCell = { Sheet = wb.ActiveSheet; Cell = addr }
        match Calc.trySingleCellEdit sheetCell cellData wb st.CalcEngine.Value with
        | Calc.IncrementalApplied (updatedWorkbook, updatedEngine) ->
            st.CalcEngine.Value <- Some updatedEngine
            st.LastRecalcWasIncremental <- true
            match Map.tryFind updatedWorkbook.ActiveSheet updatedWorkbook.Sheets with
            | Some sheet -> sheet.Model
            | None -> edited
        | Calc.IncrementalFallback _ ->
            invalidateCalc st
            Formula.recalcAndSpill None None edited

let private writeCell (st: EditorState) row col text model =
    let value = parseCellInput text
    let addr : M.CellAddress = { Row = row; Col = col }
    match M.validateEdit addr value model with
    | Result.Error _ -> model
    | Result.Ok () ->
        let edited =
            match value, M.getCell addr model with
            | M.Empty, _ -> M.clearCell addr model
            | v, Some existing -> M.setCell addr { existing with Value = v } model
            | v, None -> M.setCell addr { Value = v; Format = M.defaultFormat; Comment = None } model
        if Object.ReferenceEquals(model, edited) then model
        else
            let cellData =
                match value, M.getCell addr model with
                | M.Empty, _ -> ({ Value = M.Empty; Format = M.defaultFormat; Comment = None } : M.CellData)
                | v, Some existing -> { existing with Value = v }
                | v, None -> { Value = v; Format = M.defaultFormat; Comment = None }
            let recomputed = recalcEditedCell st addr cellData edited model
            let target = sprintf "Range(\"%s%d\")" (M.columnToLetter col) (row + 1)
            let afterEvent = Rbt.Spreadsheet.Vba.fireEvent "Worksheet_Change" target recomputed
            if not (Object.ReferenceEquals(recomputed, afterEvent)) then invalidateCalc st
            afterEvent

let applyCellEdit (st: EditorState) row col text =
    let before = activeModel st
    let after = writeCell st row col text before
    if not (Object.ReferenceEquals(before, after)) then
        st.UndoStack.Value <- (st.Sheets.Value, st.ActiveSheet.Value) :: st.UndoStack.Value |> List.truncate 50
        st.RedoStack.Value <- []
        putActiveModel st after
        true
    else false

let startEditing (st: EditorState) row col seed =
    let model = activeModel st
    let r = clampRow model row
    let c = clampCol model col
    st.EditingCell.Value <- Some (r, c)
    st.EditBuffer.Value <- seed |> Option.defaultWith (fun () -> cellEditText model r c)
    setFocus st r c

let commitEdit (st: EditorState) =
    match st.EditingCell.Value with
    | Some (r, c) ->
        let changed = applyCellEdit st r c st.EditBuffer.Value
        st.EditingCell.Value <- None
        st.EditBuffer.Value <- ""
        changed
    | None -> false

let cancelEdit (st: EditorState) =
    st.EditingCell.Value <- None
    st.EditBuffer.Value <- ""

let editFormula (st: EditorState) text =
    match st.EditingCell.Value with
    | Some _ -> st.EditBuffer.Value <- text
    | None ->
        let r, c = focusCell (activeModel st)
        st.EditingCell.Value <- Some (r, c)
        st.EditBuffer.Value <- text

let submitFormula (st: EditorState) =
    commitEdit st |> ignore
    let r, c = focusCell (activeModel st)
    setFocus st (r + 1) c

let formulaText (st: EditorState) =
    match st.EditingCell.Value with
    | Some _ -> st.EditBuffer.Value
    | None ->
        let model = activeModel st
        let r, c = focusCell model
        cellEditText model r c

let activeCellRef (st: EditorState) =
    let r, c = focusCell (activeModel st)
    M.formatCellRef r c

let selectionSummary (st: EditorState) =
    sprintf "Cell %s   %d filled" (activeCellRef st) (Map.count (activeModel st).Cells)

let copySelection (st: EditorState) =
    let model = activeModel st
    let data = M.serializeSelection model.Selection model false
    Clip.Data <- Some data
    let r, c = focusCell model
    Clip.Text <- cellEditText model r c
    true

let cutSelection (st: EditorState) =
    let model = activeModel st
    Clip.Data <- Some (M.serializeSelection model.Selection model true)
    let r, c = focusCell model
    Clip.Text <- cellEditText model r c
    applyChange st (M.clearSelectionCells model.Selection) |> ignore
    true

let pasteFromClipboard (st: EditorState) =
    match Clip.Data with
    | Some data ->
        let r, c = focusCell (activeModel st)
        applyChange st (M.pasteClipboard { Row = r; Col = c } data) |> ignore
        true
    | None when Clip.Text <> "" ->
        let r, c = focusCell (activeModel st)
        applyCellEdit st r c Clip.Text |> ignore
        true
    | _ -> false

let private pageRowJump (st: EditorState) fromRow direction =
    let model = activeModel st
    let usableH = max 60.0 (st.VisibleHeight.Peek() - 20.0)
    let mutable rows = 0
    let mutable acc = 0.0
    let mutable r = fromRow
    while acc < usableH && r >= 0 && r < model.RowCount do
        acc <- acc + M.getRowHeight r model
        if acc < usableH then
            rows <- rows + 1
            r <- r + direction
    max 1 rows

let handleKey (st: EditorState) (ev: KeyEvent) : bool =
    let editing = st.EditingCell.Value.IsSome
    let isCtrl k = ev.Ctrl && (ev.Key = k || ev.Key = k.ToUpperInvariant())
    match ev with
    | { Key = "Escape" } when editing -> cancelEdit st; true
    | { Key = "Enter" } when editing -> submitFormula st; true
    | { Key = "Tab"; Shift = false } when editing ->
        commitEdit st |> ignore
        let r, c = focusCell (activeModel st)
        setFocus st r (c + 1)
        true
    | { Key = "Tab"; Shift = true } when editing ->
        commitEdit st |> ignore
        let r, c = focusCell (activeModel st)
        setFocus st r (c - 1)
        true
    | _ when editing ->
        match ev.Key with
        | "Backspace" ->
            let buf = st.EditBuffer.Value
            if buf.Length > 0 then st.EditBuffer.Value <- buf.Substring(0, buf.Length - 1)
            true
        | k when k.Length = 1 && not ev.Ctrl && not ev.Alt ->
            st.EditBuffer.Value <- st.EditBuffer.Value + k
            true
        | _ -> false
    | _ when isCtrl "z" && ev.Shift -> redo st
    | _ when isCtrl "z" -> undo st
    | _ when isCtrl "y" -> redo st
    | _ when isCtrl "c" -> copySelection st
    | _ when isCtrl "x" -> cutSelection st
    | _ when isCtrl "v" -> pasteFromClipboard st
    | { Ctrl = true; Key = "Home" } -> setFocus st 0 0; true
    | { Ctrl = true; Key = "End" } ->
        let maxRow, maxCol = M.usedBounds (activeModel st) |> Option.defaultValue (0, 0)
        setFocus st maxRow maxCol
        true
    | { Key = "Home"; Shift = false } ->
        let r, _ = focusCell (activeModel st)
        setFocus st r 0
        true
    | { Key = "End"; Shift = false } ->
        let model = activeModel st
        let r, _ = focusCell model
        let lastCol =
            model.Cells
            |> Seq.filter (fun kv -> kv.Key.Row = r)
            |> Seq.map (fun kv -> kv.Key.Col)
            |> Seq.fold max 0
        setFocus st r lastCol
        true
    | { Key = "PageDown"; Shift = false } ->
        let r, c = focusCell (activeModel st)
        setFocus st (r + pageRowJump st r 1) c
        true
    | { Key = "PageUp"; Shift = false } ->
        let r, c = focusCell (activeModel st)
        setFocus st (r - pageRowJump st r -1) c
        true
    | { Key = "PageDown"; Shift = true } ->
        let r, c = focusCell (activeModel st)
        extendFocus st (r + pageRowJump st r 1) c
        true
    | { Key = "PageUp"; Shift = true } ->
        let r, c = focusCell (activeModel st)
        extendFocus st (r - pageRowJump st r -1) c
        true
    | { Key = "ArrowLeft"; Shift = false } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c - 1) 0 -1
        setFocus st nr nc
        true
    | { Key = "ArrowRight"; Shift = false } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c + 1) 0 1
        setFocus st nr nc
        true
    | { Key = "ArrowUp"; Shift = false } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r - 1) c -1 0
        setFocus st nr nc
        true
    | { Key = "ArrowDown"; Shift = false } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r + 1) c 1 0
        setFocus st nr nc
        true
    | { Key = "ArrowLeft"; Shift = true } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c - 1) 0 -1
        extendFocus st nr nc
        true
    | { Key = "ArrowRight"; Shift = true } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model r (c + 1) 0 1
        extendFocus st nr nc
        true
    | { Key = "ArrowUp"; Shift = true } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r - 1) c -1 0
        extendFocus st nr nc
        true
    | { Key = "ArrowDown"; Shift = true } ->
        let model = activeModel st
        let r, c = focusCell model
        let nr, nc = resolveNavTarget model (r + 1) c 1 0
        extendFocus st nr nc
        true
    | { Key = "Enter" } ->
        let r, c = focusCell (activeModel st)
        setFocus st (r + 1) c
        true
    | { Key = "Tab"; Shift = false } ->
        let r, c = focusCell (activeModel st)
        setFocus st r (c + 1)
        true
    | { Key = "Tab"; Shift = true } ->
        let r, c = focusCell (activeModel st)
        setFocus st r (c - 1)
        true
    | { Key = "F2" } ->
        let r, c = focusCell (activeModel st)
        startEditing st r c None
        true
    | { Key = "Delete" } ->
        let model = activeModel st
        applyChange st (M.clearSelectionCells model.Selection) |> ignore
        true
    | { Key = "Backspace" } ->
        let model = activeModel st
        applyChange st (M.clearSelectionCells model.Selection) |> ignore
        true
    | { Ctrl = false; Alt = false; Key = key } when key.Length = 1 ->
        let r, c = focusCell (activeModel st)
        startEditing st r c (Some key)
        true
    | _ -> false

let viewportOf (st: EditorState) width height =
    if st.VisibleWidth.Peek() <> width then st.VisibleWidth.Value <- width
    if st.VisibleHeight.Peek() <> height then st.VisibleHeight.Value <- height
    GR.viewport st.ScrollX.Value st.ScrollY.Value width height

let private hitCell st width height x y =
    let model = activeModel st
    match L.hitTest model (viewportOf st width height) x y with
    | L.Cell addr -> Some (addr.Row, addr.Col)
    | L.HitRowHeader row -> Some (row, 0)
    | L.HitColumnHeader col -> Some (0, col)
    | _ -> None

let handlePointer st width height kind x y =
    match kind with
    | Down ->
        if st.EditingCell.Value.IsSome then commitEdit st |> ignore
        match hitCell st width height x y with
        | Some (row, col) ->
            st.DragAnchor.Value <- Some (row, col)
            setFocus st row col
        | None -> ()
    | Move ->
        match st.DragAnchor.Value, hitCell st width height x y with
        | Some _, Some (row, col) -> extendFocus st row col
        | _ -> ()
    | Up -> st.DragAnchor.Value <- None

let scrollBy st dx dy =
    st.ScrollX.Value <- max 0.0 (st.ScrollX.Peek() + dx)
    st.ScrollY.Value <- max 0.0 (st.ScrollY.Peek() + dy)

let renderGrid st width height =
    GR.commands (activeModel st) (viewportOf st width height)
