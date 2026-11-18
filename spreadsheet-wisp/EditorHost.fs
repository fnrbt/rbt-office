module Rbt.SpreadsheetWisp.EditorHost

open System
open Wisp.Reactive
open Wisp.Layout

module M = Rbt.Spreadsheet.Model
module L = Rbt.Spreadsheet.Layout
module Session = Rbt.Spreadsheet.Session
module Workbook = Rbt.Spreadsheet.Workbook
module GR = Rbt.SpreadsheetNative.GridRender

type KeyEvent = { Key: string; Shift: bool; Ctrl: bool; Alt: bool }

type ClipboardTransport =
    { WriteText: string -> Result<string, string>
      ReadText: unit -> Result<Wisp.OfficeClipboard.Contents, string> }


type EditorState =
    { Session: Session.Session
      Workbook: RefSignal<Workbook.Workbook>
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
      ClipboardStatus: Signal<string>
      mutable ClipboardToken: string option
      mutable ClipboardTransport: ClipboardTransport
      /// Whether the workbook's macros may run on their own (Worksheet_Change ...): off for a workbook
      /// opened, on once the user runs or enables its macros, as Excel's "Enable content".
      mutable MacrosEnabled: bool }

let createWithWorkbook (wb: Workbook.Workbook) =
    let entries = wb.SheetOrder |> List.map (fun id -> wb.Sheets.[id]) |> List.toArray
    { Session = Session.Session(wb)
      Workbook = refSignal wb
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
      ClipboardStatus = signal ""
      ClipboardToken = None
      ClipboardTransport = { WriteText = Wisp.OfficeClipboard.writeText; ReadText = Wisp.OfficeClipboard.readText }
      MacrosEnabled = false }

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
    st.Session.Replace(wb, markSaved = true)
    st.ClipboardToken <- None
    st.MacrosEnabled <- false
    install st st.Session.Workbook
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
          Macros = models |> Array.tryPick (fun model -> if String.IsNullOrWhiteSpace model.Macros then None else Some model.Macros)
          MacroProject = None }
    Rbt.Spreadsheet.XlsxBridge.Export.toWorkbookOut input
    |> Rbt.Spreadsheet.XlsxBridge.Export.withIteration wb.Iteration

let canUndo (st: EditorState) =
    st.Workbook.Value |> ignore
    st.Session.CanUndo
let canRedo (st: EditorState) =
    st.Workbook.Value |> ignore
    st.Session.CanRedo
let revision (st: EditorState) = st.Session.Revision
let isModified (st: EditorState) = st.Session.IsModified
let captureSave (st: EditorState) = st.Session.CaptureSave()
let markSaved (st: EditorState) (capture: Session.SaveToken) =
    st.Session.MarkSaved(revision = capture.Revision, epoch = capture.Epoch)

let private applyTransaction (st: EditorState) (transform: Workbook.Workbook -> Workbook.Workbook) =
    let before = st.Session.Workbook
    let after = transform before
    if Object.ReferenceEquals(before, after) then false
    else
        st.Session.Apply
            { Description = "Edit workbook"
              Apply = fun _ -> after
              Recalculate = true
              Hint = None
              AfterCalculate = None }
        install st st.Session.Workbook
        true

let applyNoHistory (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    let wb = st.Session.Workbook
    let before = wb.Sheets.[wb.ActiveSheet].Model
    let after = transform before
    if not (Object.ReferenceEquals(before, after)) then
        st.Session.UpdateView(Workbook.updateSheetModel wb.ActiveSheet after wb)
        install st st.Session.Workbook
    true

let applyChange (st: EditorState) (transform: M.SpreadsheetModel -> M.SpreadsheetModel) =
    applyTransaction st (fun wb ->
        let before = wb.Sheets.[wb.ActiveSheet].Model
        let after = transform before
        if Object.ReferenceEquals(before, after) then wb else updateModel wb after) |> ignore
    true

let applyWorkbookChange (st: EditorState) (transform: Workbook.Workbook -> Workbook.Workbook) =
    applyTransaction st transform |> ignore
    true
/// Session performs the structural transaction; signals only project its result.
let executeStructural (st: EditorState) (command: Session.StructuralCommand) =
    match st.Session.ExecuteStructural(command, recalculate = true) with
    | Ok true ->
        install st st.Session.Workbook
        Ok true
    | result -> result

let applyStructureEdit (st: EditorState) (edit: M.StructureEdit) =
    executeStructural st (Session.EditStructure (st.Session.Workbook.ActiveSheet, edit))
    |> ignore
    true

let addSheet (st: EditorState) name =
    executeStructural st (Session.AddSheet (Some name, None, 100, 26))
    |> Result.map (fun _ -> ())

let renameSheet (st: EditorState) name =
    executeStructural st (Session.RenameSheet (st.Session.Workbook.ActiveSheet, name))
    |> Result.map (fun _ -> ())

let deleteSheet (st: EditorState) =
    executeStructural st (Session.DeleteSheet st.Session.Workbook.ActiveSheet)
    |> Result.map (fun _ -> ())

let undo (st: EditorState) =
    if st.Session.Undo(restoreActiveSheet = true) then install st st.Session.Workbook
    true

let redo (st: EditorState) =
    if st.Session.Redo(restoreActiveSheet = true) then install st st.Session.Workbook
    true

let switchSheet (st: EditorState) index =
    let wb = st.Workbook.Value
    match wb.SheetOrder |> List.tryItem index with
    | Some id when id <> wb.ActiveSheet ->
        st.Session.UpdateView(Workbook.setActiveSheet id wb)
        install st st.Session.Workbook
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

/// Formatting is a single session transaction; publish only committed changes.
let executeFormatting (st: EditorState) (command: Session.FormattingCommand) =
    let row, col = focusCell (activeModel st)
    let activeCell : M.CellAddress = { Row = row; Col = col }
    match st.Session.ExecuteFormatting(st.Session.Workbook.ActiveSheet, command, activeCell = activeCell) with
    | Ok true ->
        install st st.Session.Workbook
        Ok true
    | result -> result
/// Defined names use the shared workbook transaction, including recalculation.
let executeName (st: EditorState) (command: Session.NameCommand) =
    match st.Session.ExecuteName command with
    | Ok true ->
        install st st.Session.Workbook
        Ok true
    | result -> result

/// Rules are sheet-owned; only a committed change updates UI projections.
let executeRule (st: EditorState) (command: Session.RuleCommand) =
    match st.Session.ExecuteRule(st.Session.Workbook.ActiveSheet, command) with
    | Ok result when result.Changed ->
        install st st.Session.Workbook
        Ok result
    | result -> result

/// Publish a shared drawing transaction without recalculating cell values.
let executeDrawing (st: EditorState) (command: Session.DrawingCommand) =
    match st.Session.ExecuteDrawing(st.Session.Workbook.ActiveSheet, command) with
    | Ok result when result.Changed ->
        install st st.Session.Workbook
        Ok result
    | result -> result


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

/// The workbook's macros compiled (the macro dialog's text, one standard module), by their text.
let private programs = Collections.Generic.Dictionary<string, Result<Rbt.Spreadsheet.MacroHost.MacroProgram, Rbt.Vb.Api.VbFailure list>>()
let private programOf (source: string) =
    if String.IsNullOrWhiteSpace source then None
    else
        match programs.TryGetValue source with
        | true, program -> Some program
        | _ ->
            let program = Rbt.Spreadsheet.MacroHost.MacroProgram.Compile (Rbt.Spreadsheet.MacroHost.editedModules source)
            programs.[source] <- program
            Some program

let applyCellEdit (st: EditorState) row col text =
    let id = st.Session.Workbook.ActiveSheet
    let addr : M.CellAddress = { Row = row; Col = col }
    let accept (cell: M.CellData) = M.validateEdit addr cell.Value st.Session.Workbook.Sheets.[id].Model |> Result.isOk
    // Worksheet_Change runs over the whole workbook, in the edit's step: undone with it.
    let afterCalculate (wb: Workbook.Workbook) =
        if not st.MacrosEnabled then wb
        else
            match programOf wb.Sheets.[id].Model.Macros with
            | Some (Result.Ok program) ->
                match program.RunEvent(wb, Rbt.Spreadsheet.MacroHost.defaultHooks, "Worksheet_Change", Some (id, row, col, row, col)) with
                | Some (Rbt.Spreadsheet.MacroHost.MacroDone (changed, _)) -> changed
                | _ -> wb
            | _ -> wb
    let changed = st.Session.EnterCell(id, addr, text, DateTime.UtcNow.Year, accept = accept, afterCalculate = afterCalculate)
    if changed then install st st.Session.Workbook
    changed

/// Run a macro of the source on the workbook: `procedure` (a Sub of it), else Main, else its first macro.
/// The source becomes the workbook's macros (a save writes it as its VBA project when it differs from
/// the one opened). What the macro changes is one undo step; Error (Excel's number and description, with
/// where it stopped) when it stops, and then nothing of it is kept. Enables the workbook's macros.
let runMacro (st: EditorState) (hooks: Rbt.Spreadsheet.MacroHost.MacroHooks) (source: string) (procedure: string option) : Result<unit, int * string> =
    st.MacrosEnabled <- true
    let start = st.Session.Workbook
    let withSource (wb: Workbook.Workbook) =
        { wb with Sheets = wb.Sheets |> Map.map (fun _ sheet -> if sheet.Model.Macros = source then sheet else { sheet with Model = { sheet.Model with Macros = source } }) }
    let outcome =
        match programOf source with
        | None -> Rbt.Spreadsheet.MacroHost.MacroFailed { Detail = ""; Kind = Rbt.Vb.Api.Compile; Number = 0; Description = "There are no macros"; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" }
        | Some (Result.Error failures) -> Rbt.Spreadsheet.MacroHost.MacroFailed (List.head failures)
        | Some (Result.Ok program) ->
            let name =
                match procedure with
                | Some name -> Some name
                | None -> if program.Has "Main" then Some "Main" else List.tryHead program.Macros
            match name with
            | Some name -> program.Run(start, hooks, name)
            | None -> Rbt.Spreadsheet.MacroHost.MacroFailed { Detail = ""; Kind = Rbt.Vb.Api.Compile; Number = 0; Description = "There is no Sub to run"; Source = ""; Line = 0; Column = 0; Procedure = ""; Module = "" }
    match outcome with
    | Rbt.Spreadsheet.MacroHost.MacroDone (result, selection) ->
        let result =
            match selection with
            | Some (sheetId, top, left, bottom, right) when Map.containsKey sheetId result.Sheets ->
                let sheet = result.Sheets.[sheetId]
                let selected = if top = bottom && left = right then M.SingleCell (top, left) else M.Range (top, left, bottom, right)
                { result with Sheets = Map.add sheetId { sheet with Model = { sheet.Model with Selection = selected } } result.Sheets }
            | _ -> result
        st.Session.CommitApplied({ Description = "Run macro"; Workbook = withSource result })
        install st st.Session.Workbook
        Ok ()
    | Rbt.Spreadsheet.MacroHost.MacroFailed failure ->
        if start.Sheets |> Map.exists (fun _ sheet -> sheet.Model.Macros <> source) then
            st.Session.CommitApplied({ Description = "Edit macros"; Workbook = withSource start })
            install st st.Session.Workbook
        Error (failure.Number, Rbt.Vb.Api.VbFailure.format failure)

/// Make `source` the workbook's macros (every sheet holds the same copy; a save writes it): one undo step.
let setMacros (st: EditorState) (source: string) =
    let start = st.Session.Workbook
    if start.Sheets |> Map.exists (fun _ sheet -> sheet.Model.Macros <> source) then
        let edited = { start with Sheets = start.Sheets |> Map.map (fun _ sheet -> { sheet with Model = { sheet.Model with Macros = source } }) }
        st.Session.CommitApplied({ Description = "Edit macros"; Workbook = edited })
        install st st.Session.Workbook

/// The macros of the source (the editor's macro list).
let macroNames (source: string) : string list =
    match programOf source with
    | Some (Result.Ok program) -> program.Macros
    | _ -> []

/// What keeps the source from compiling, one line each (none when it compiles).
let macroProblems (source: string) : string list =
    match programOf source with
    | Some (Result.Error failures) -> failures |> List.map Rbt.Vb.Api.VbFailure.format
    | _ -> []

/// Run an Immediate-window line (`?expression`, `Print expression`, statements) with the source's
/// program on the workbook; what it changes is one undo step. Enables the workbook's macros.
let runMacroLine (st: EditorState) (hooks: Rbt.Spreadsheet.MacroHost.MacroHooks) (source: string) (line: string) : Result<unit, string> =
    st.MacrosEnabled <- true
    let program =
        match programOf source with
        | Some program -> program
        | None -> Rbt.Spreadsheet.MacroHost.MacroProgram.Compile []
    match program with
    | Result.Error failures -> Error (Rbt.Vb.Api.VbFailure.format (List.head failures))
    | Result.Ok program ->
        let start = st.Session.Workbook
        match program.Immediate(start, hooks, line) with
        | Rbt.Spreadsheet.MacroHost.MacroDone (result, _) ->
            if not (Object.ReferenceEquals(result, start)) then
                st.Session.CommitApplied({ Description = "Immediate"; Workbook = result })
                install st st.Session.Workbook
            Ok ()
        | Rbt.Spreadsheet.MacroHost.MacroFailed failure -> Error (Rbt.Vb.Api.VbFailure.format failure)

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


let private clipboardText (model: M.SpreadsheetModel) (data: M.ClipboardData) =
    // Bound whole-axis copies, but retain the blank cells of an explicit range.
    let wholeRows, wholeCols =
        match model.Selection with
        | M.Range (r0, c0, r1, c1) -> r0 = 0 && r1 = model.RowCount - 1, c0 = 0 && c1 = model.ColCount - 1
        | _ -> false, false
    let usedRow, usedCol =
        data.Cells |> List.fold (fun (lastRow, lastCol) (row, col, _) -> max lastRow row, max lastCol col) (0, 0)
    let height = data.SourceEndRow - data.SourceStartRow
    let width = data.SourceEndCol - data.SourceStartCol
    let maxRow = if wholeRows then min height usedRow else height
    let maxCol = if wholeCols then min width usedCol else width
    let result = System.Text.StringBuilder()
    // serializeSelection supplies sparse cells in row/column order.
    let mutable cells = data.Cells
    // Terminate every record: otherwise a final empty single-column row disappears.
    for row in 0 .. maxRow do
        for col in 0 .. maxCol do
            if col > 0 then result.Append('\t') |> ignore
            match cells with
            | (cellRow, cellCol, cell) :: rest when cellRow = row && cellCol = col ->
                let value = Rbt.Spreadsheet.CellText.cellDisplayTextIn model.Date1904 None 1.0e6 cell.Format cell.Value
                if value.Contains('\t') || value.Contains('\r') || value.Contains('\n') || value.Contains('"') then
                    result.Append('"') |> ignore
                    for ch in value do
                        if ch = '"' then result.Append('"') |> ignore
                        result.Append(ch) |> ignore
                    result.Append('"') |> ignore
                else result.Append(value) |> ignore
                cells <- rest
            | _ -> ()
        result.Append("\r\n") |> ignore
    result.ToString()

let private captureSelection (st: EditorState) isCut =
    let mutable token = None
    let transfer text =
        match st.ClipboardTransport.WriteText text with
        | Ok written -> token <- Some written; Ok ()
        | Error problem -> Error problem
    let sheetId = st.Session.Workbook.ActiveSheet
    match st.Session.CaptureClipboard(sheetId, isCut, "",
                                      textOf = (fun data -> clipboardText st.Session.Workbook.Sheets.[sheetId].Model data),
                                      transfer = transfer) with
    | Ok changed ->
        st.ClipboardToken <- token
        if changed || not (Object.ReferenceEquals(st.Workbook.Value, st.Session.Workbook)) then
            install st st.Session.Workbook
        st.ClipboardStatus.Value <- if isCut then "Cut" else "Copied"
        true
    | Error problem ->
        st.ClipboardStatus.Value <- "Clipboard: " + problem
        false

let copySelection st = captureSelection st false
let cutSelection st = captureSelection st true

let pasteFromClipboard (st: EditorState) =
    let result =
        match st.ClipboardTransport.ReadText () with
        | Error problem -> Error problem
        | Ok contents when contents.Token.IsSome && contents.Token = st.ClipboardToken ->
            st.Session.PasteClipboard(st.Session.Workbook.ActiveSheet)
        | Ok contents when String.IsNullOrEmpty contents.Text -> Ok false
        | Ok contents ->
            st.Session.PasteRows(st.Session.Workbook.ActiveSheet, Session.clipboardRowsOfText contents.Text, DateTime.Now.Year)
    match result with
    | Ok changed ->
        if changed || not (Object.ReferenceEquals(st.Workbook.Value, st.Session.Workbook)) then
            install st st.Session.Workbook
        st.ClipboardStatus.Value <- if changed then "Pasted" else "Clipboard empty"
        changed
    | Error problem ->
        st.ClipboardStatus.Value <- "Clipboard: " + problem
        false

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
