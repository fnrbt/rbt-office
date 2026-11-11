module Rbt.SpreadsheetWisp.EditorHost

open System
open Wisp.Reactive
open Wisp.Layout

module M = Rbt.Spreadsheet.Model
module L = Rbt.Spreadsheet.Layout
module Recalc = Rbt.Spreadsheet.Recalc
module Workbook = Rbt.Spreadsheet.Workbook
module GR = Rbt.SpreadsheetWisp.GridRender

type KeyEvent = { Key: string; Shift: bool; Ctrl: bool; Alt: bool }


type EditorState =
    { Workbook: RefSignal<Workbook.Workbook>
      Sheets: RefSignal<M.SpreadsheetModel array>
      SheetNames: RefSignal<string array>
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
      mutable UsedSheetIds: Set<string>
      mutable Clipboard: (M.ClipboardData * string) option
      mutable ClipboardText: string }

let createWithWorkbook (wb: Workbook.Workbook) =
    let entries = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id]) |> List.toArray
    { Workbook = refSignal wb
      Sheets = refSignal (entries |> Array.map (fun s -> s.Model))
      SheetNames = refSignal (entries |> Array.map (fun s -> s.Name))
      ActiveSheet = signal (max 0 (Workbook.sheetIndex wb.ActiveSheet wb))
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
      UsedSheetIds = Set.ofList wb.SheetOrder
      Clipboard = None
      ClipboardText = "" }

let create (sheets: M.SpreadsheetModel array) (names: string array) (activeIndex: int) =
    let sheets = if sheets.Length = 0 then [| M.createEmpty 100 26 |] else sheets
    Rbt.Spreadsheet.XlsxBridge.buildWorkbookContext sheets names (Array.create sheets.Length true) Map.empty Map.empty activeIndex
    |> createWithWorkbook

let activeModel (st: EditorState) : M.SpreadsheetModel =
    let sheets = st.Sheets.Value
    let idx = st.ActiveSheet.Value
    if idx >= 0 && idx < sheets.Length then sheets.[idx] else sheets.[0]

let activeSheetName (st: EditorState) =
    let names = st.SheetNames.Value
    let idx = st.ActiveSheet.Value
    if idx >= 0 && idx < names.Length then names.[idx] else "Sheet1"

let private install (st: EditorState) (wb: Workbook.Workbook) =
    st.Workbook.Value <- wb
    let entries = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id]) |> List.toArray
    st.Sheets.Value <- entries |> Array.map (fun s -> s.Model)
    st.SheetNames.Value <- entries |> Array.map (fun s -> s.Name)
    st.ActiveSheet.Value <- max 0 (Workbook.sheetIndex wb.ActiveSheet wb)

let replaceWorkbook (st: EditorState) wb =
    install st wb
    st.UndoStack.Value <- []
    st.RedoStack.Value <- []
    st.CalcEngine.Value <- None
    st.UsedSheetIds <- Set.ofList wb.SheetOrder
    st.Clipboard <- None
    st.ClipboardText <- ""
    st.EditingCell.Value <- None
    st.EditBuffer.Value <- ""
    st.DragAnchor.Value <- None
    st.ScrollX.Value <- 0.0
    st.ScrollY.Value <- 0.0

let private updateModel (wb: Workbook.Workbook) (model: M.SpreadsheetModel) =
    Workbook.updateSheetModelWithTables wb.ActiveSheet model wb

let exportWorkbook (st: EditorState) (source: Rbt.Spreadsheet.XlsxBridge.XlsxSource option) =
    let wb = st.Workbook.Value
    let sheets = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id]) |> List.toArray
    let ids = wb.SheetOrder |> List.mapi (fun i id -> id, Rbt.Spreadsheet.XlsxBridge.sheetIdOfIndex i) |> Map.ofList
    let remap entries =
        entries |> Map.toList |> List.choose (fun ((scope, name), value) ->
            match scope with
            | None -> Some ((None, name), value)
            | Some id -> Map.tryFind id ids |> Option.map (fun id -> (Some id, name), value)) |> Map.ofList
    let models = sheets |> Array.map (fun sheet -> sheet.Model)
    let input : Rbt.Spreadsheet.XlsxBridge.ExportInput =
        { Sheets = models
          Names = sheets |> Array.map (fun sheet -> sheet.Name)
          Visible = sheets |> Array.map (fun sheet -> sheet.Visible)
          ActiveSheetIndex = max 0 (Workbook.sheetIndex wb.ActiveSheet wb)
          DefinedNames = remap wb.Names
          NameInfo = remap wb.NameInfo
          Source = source
          Macros = models |> Array.tryPick (fun model -> if String.IsNullOrWhiteSpace model.Macros then None else Some model.Macros) }
    Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut input
    |> Rbt.Spreadsheet.XlsxBridge.Export.withIteration wb.Iteration

let private calculate (st: EditorState) wb hint =
    let result = st.CalcEngine.Value |> Option.bind (fun engine -> Recalc.recalculate engine wb hint)
    match result with
    | Some calculated -> calculated
    | None ->
        let calculated, engine = Recalc.recalculateAll wb
        st.CalcEngine.Value <- Some engine
        calculated

let private commitCalculated (st: EditorState) before after =
    if not (Object.ReferenceEquals(before, after)) then
        st.UndoStack.Value <- before :: st.UndoStack.Value |> List.truncate 50
        st.RedoStack.Value <- []
        install st after
        true
    else false

let private commit (st: EditorState) before after hint =
    if Object.ReferenceEquals(before, after) then false
    else commitCalculated st before (calculate st after hint)

let applyNoHistory (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    let wb = st.Workbook.Value
    let before = wb.Sheets.[wb.ActiveSheet].Model
    let after = transform before
    if not (Object.ReferenceEquals(before, after)) then install st (Workbook.updateSheetModel wb.ActiveSheet after wb)
    true

let applyChange (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    let wb = st.Workbook.Value
    let before = wb.Sheets.[wb.ActiveSheet].Model
    let after = transform before
    if not (Object.ReferenceEquals(before, after)) then commit st wb (updateModel wb after) None |> ignore
    true

let applyWorkbookChange (st: EditorState) (transform: Workbook.Workbook -> Workbook.Workbook) =
    let before = st.Workbook.Value
    commit st before (transform before) None |> ignore
    true

let applyStructureEdit (st: EditorState) (edit: M.StructureEdit) =
    let id = st.Workbook.Value.ActiveSheet
    applyWorkbookChange st (fun wb -> Workbook.applyStructureEdit id edit wb |> fst)

let addSheet (st: EditorState) name =
    let wb = st.Workbook.Value
    match Workbook.sheetNameProblem name with
    | Some problem -> Error problem
    | None when Workbook.sheetNameTaken None name wb -> Error "a sheet with that name already exists"
    | None ->
        let id = Workbook.freshSheetId (fun id -> Set.contains id st.UsedSheetIds) wb
        st.UsedSheetIds <- Set.add id st.UsedSheetIds
        let model = { M.createEmpty 100 26 with Date1904 = Workbook.date1904 wb }
        let sheet : Workbook.Sheet = { Id = id; Name = name; Model = model; Visible = true }
        applyWorkbookChange st (Workbook.insertSheet sheet wb.SheetOrder.Length) |> ignore
        Ok ()

let renameSheet (st: EditorState) name =
    let wb = st.Workbook.Value
    match Workbook.sheetNameProblem name with
    | Some problem -> Error problem
    | None when Workbook.sheetNameTaken (Some wb.ActiveSheet) name wb -> Error "a sheet with that name already exists"
    | None ->
        applyWorkbookChange st (Workbook.renameSheetEverywhere wb.ActiveSheet name) |> ignore
        Ok ()

let deleteSheet (st: EditorState) =
    let wb = st.Workbook.Value
    match Workbook.lastVisibleSheetProblem wb.ActiveSheet wb with
    | Some problem -> Error problem
    | None ->
        match Workbook.removeSheet wb.ActiveSheet wb with
        | None -> Error "cannot delete the last sheet"
        | Some (next, _) ->
            applyWorkbookChange st (fun _ -> next) |> ignore
            Ok ()

let undo (st: EditorState) =
    match st.UndoStack.Value with
    | wb :: rest ->
        st.RedoStack.Value <- st.Workbook.Value :: st.RedoStack.Value |> List.truncate 50
        st.UndoStack.Value <- rest
        st.CalcEngine.Value <- None
        install st wb
        true
    | [] -> true

let redo (st: EditorState) =
    match st.RedoStack.Value with
    | wb :: rest ->
        st.UndoStack.Value <- st.Workbook.Value :: st.UndoStack.Value |> List.truncate 50
        st.RedoStack.Value <- rest
        st.CalcEngine.Value <- None
        install st wb
        true
    | [] -> true

let switchSheet (st: EditorState) index =
    let wb = st.Workbook.Value
    match wb.SheetOrder |> List.tryItem index with
    | Some id when id <> wb.ActiveSheet ->
        install st (Workbook.setActiveSheet id wb)
        st.ScrollX.Value <- 0.0
        st.ScrollY.Value <- 0.0
        st.EditingCell.Value <- None
        st.EditBuffer.Value <- ""
    | _ -> ()

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
    // What, typed back unchanged, gives the same cell: the formula's text, a
    // date as a date, forced text with its apostrophe.
    Rbt.Spreadsheet.Input.editTextFor model.Date1904 (M.getCell { Row = row; Col = col } model)

/// Typed text → cell value, read the way Excel reads input (shared with the
/// web editor: `Rbt.Spreadsheet.Input`). This host stores formulas without
/// the leading `=`.
let parseCellInput (date1904: bool) (text: string) : M.CellValue =
    match (Rbt.Spreadsheet.Input.parseInputFor date1904 System.DateTime.UtcNow.Year text).Value with
    | M.Formula (formula, cached) when formula.StartsWith "=" -> M.Formula (formula.Substring 1, cached)
    | value -> value

/// The cell's format after an entry: an entry written as a percentage, a
/// currency amount, a date… gives a General cell the matching number format.
let private formatAfterInput (date1904: bool) (text: string) (format: M.CellFormat) : M.CellFormat =
    match (Rbt.Spreadsheet.Input.parseInputFor date1904 System.DateTime.UtcNow.Year text).NumberFormat, format.NumberFormat with
    | Some code, (None | Some "" | Some "General") -> { format with NumberFormat = Some code }
    | _ -> format

let private writeCell row col text (model: M.SpreadsheetModel) =
    let value = parseCellInput model.Date1904 text
    let addr : M.CellAddress = { Row = row; Col = col }
    match M.validateEdit addr value model with
    | Result.Error _ -> model
    | Result.Ok () ->
        match value, M.getCell addr model with
        | M.Empty, _ -> M.clearCell addr model
        | v, Some existing -> M.setCell addr { existing with Value = v; Format = formatAfterInput model.Date1904 text existing.Format } model
        | v, None -> M.setCell addr { Value = v; Format = formatAfterInput model.Date1904 text M.defaultFormat; Comment = None } model

let applyCellEdit (st: EditorState) row col text =
    let wb = st.Workbook.Value
    let before = activeModel st
    let after = writeCell row col text before
    if Object.ReferenceEquals(before, after) then false
    else
        let calculated = calculate st (updateModel wb after) None
        let model = calculated.Sheets.[calculated.ActiveSheet].Model
        let target = sprintf "Range(\"%s%d\")" (M.columnToLetter col) (row + 1)
        let afterEvent = Rbt.Spreadsheet.Vba.fireEvent "Worksheet_Change" target model
        let result =
            if Object.ReferenceEquals(model, afterEvent) then calculated
            else calculate st (updateModel calculated afterEvent) None
        commitCalculated st wb result

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
    st.Clipboard <- Some (data, st.Workbook.Value.ActiveSheet)
    let r, c = focusCell model
    st.ClipboardText <- cellEditText model r c
    true

let cutSelection (st: EditorState) =
    let model = activeModel st
    st.Clipboard <- Some (M.serializeSelection model.Selection model true, st.Workbook.Value.ActiveSheet)
    let r, c = focusCell model
    st.ClipboardText <- cellEditText model r c
    applyChange st (M.clearSelectionCells model.Selection) |> ignore
    true

let pasteFromClipboard (st: EditorState) =
    match st.Clipboard with
    | Some (_, sourceId) when not (Map.containsKey sourceId st.Workbook.Value.Sheets) -> false
    | Some (data, sourceId) ->
        let model = activeModel st
        if data.IsCut then
            // A pasted cut is a move: references to the moved cells follow
            // them, on every sheet.
            let r, c = focusCell model
            let targetId = st.Workbook.Value.ActiveSheet
            let source = data.SourceStartRow, data.SourceStartCol, data.SourceEndRow, data.SourceEndCol
            let delta = r - data.SourceStartRow, c - data.SourceStartCol
            applyWorkbookChange st (fun wb ->
                match Map.tryFind targetId wb.Sheets with
                | Some sheet ->
                    wb
                    |> Workbook.updateSheetModel targetId (M.pasteClipboard { Row = r; Col = c } data sheet.Model)
                    |> Workbook.repointMovedCells sourceId source targetId delta
                    |> fst
                | None -> wb) |> ignore
            st.Clipboard <- None
        else
            applyChange st (M.pasteClipboardIntoSelection model.Selection data) |> ignore
        true
    | None when st.ClipboardText <> "" ->
        let r, c = focusCell (activeModel st)
        applyCellEdit st r c st.ClipboardText |> ignore
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
