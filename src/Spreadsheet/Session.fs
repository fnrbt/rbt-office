module Rbt.Spreadsheet.Session

open Rbt.Spreadsheet.Model
open Rbt.Spreadsheet.Workbook
open Rbt.Spreadsheet.Recalc

/// A workbook edit and an optional host event pass, performed once on commit.
/// Undo/redo restore persistent pre/post-calculation workbook snapshots instead
/// of rerunning host events or stale edit closures.
type Transaction = {
    Description: string
    Apply: Workbook -> Workbook
    Recalculate: bool
    Hint: (SheetId * CellAddress) list option
    AfterCalculate: (Workbook -> Workbook) option
}

/// Content a host command has already produced; it joins the same history.
type AppliedTransaction = { Description: string; Workbook: Workbook }
/// Structural workbook commands shared by desktop and browser hosts.
type StructuralCommand =
    | EditStructure of SheetId * StructureEdit
    | AddSheet of name: string option * atIndex: int option * rows: int * columns: int
    | RenameSheet of SheetId * string
    | DeleteSheet of SheetId
    | MoveSheet of SheetId * int
    | DuplicateSheet of SheetId
    | SetSheetVisibility of SheetId * visible: bool * veryHidden: bool
/// Formatting and selection edits shared by all spreadsheet hosts.
type FormattingCommand =
    | ToggleBold | ToggleItalic | ToggleUnderline | ToggleDoubleUnderline
    | ToggleStrikethrough | ToggleWrapText | ToggleShrinkToFit
    | ChangeFontSize of int | ChangeIndent of int | ChangeDecimals of int
    | HorizontalAlign of string option | VerticalAlign of string option
    | NumberFormat of string option | TransformFormat of (CellFormat -> CellFormat)
    | PaintFormat of CellFormat | Border of side: string * style: string * color: string
    | ClearCells of SheetOps.ClearWhat | MergeCells | UnmergeCells of CellAddress

type NameCommand =
    | SetDefinedName of scope: SheetId option * name: string * refersTo: string * comment: string option
    | DeleteDefinedName of scope: SheetId option * name: string

type RuleCommand =
    | AddConditionalFormat of CellRange list * bool * ConditionalFormat
    | UpdateConditionalFormat of id: int * ranges: CellRange list option * stopIfTrue: bool option * format: ConditionalFormat option * priority: int option
    | RemoveConditionalFormat of int
    | ReorderConditionalFormats of int list
    | ClearConditionalFormats of CellRange list option
    | SetValidation of CellRange list * DataValidation
    | ClearValidation of CellRange list

type RuleCommandResult = { Changed: bool; Id: int option; Affected: int }

/// Drawing commands deliberately edit pictures only; imported opaque objects
/// remain intact. Relationship IDs reserved by the source package are supplied
/// by the package adapter when authoring a chart.
type DrawingCommand =
    | InsertPicture of data: byte[] * cell: CellAddress * width: float * height: float * name: string option
    | PlacePicture of id: int * cell: CellAddress * offsetX: float * offsetY: float * width: float * height: float
    | DeletePicture of id: int
    | DescribePicture of id: int * name: string option * description: string option
    | InsertChart of range: CellRange * kind: ChartAuthoring.StandardChartKind * title: string * headers: bool * categories: bool * reservedRelationships: string list

type DrawingCommandResult = { Changed: bool; Id: int option }



type SaveToken = { Epoch: int; Revision: int }
type CalculationInfo = { Kind: string; Reason: string; Stats: RecalcStats; Formulas: int }
type ClipboardCapture = { SourceSheet: SheetId; Data: ClipboardData; Text: string }
type ClipboardPasteKind = InternalPaste | ExternalPaste | IgnoredPaste

/// Parse the quoted tab-separated text used by spreadsheet clipboards.
let clipboardRowsOfText (text: string) : string array array =
    let rows = ResizeArray<string array>()
    let row = ResizeArray<string>()
    let field = System.Text.StringBuilder()
    let mutable quoted = false
    let mutable atFieldStart = true
    let mutable i = 0
    let endField () =
        row.Add (field.ToString())
        field.Clear() |> ignore
        atFieldStart <- true
    let endRow () =
        endField ()
        rows.Add (row.ToArray())
        row.Clear()
    while i < text.Length do
        let ch = text.[i]
        if quoted then
            if ch = '"' then
                if i + 1 < text.Length && text.[i + 1] = '"' then
                    field.Append '"' |> ignore
                    i <- i + 1
                else quoted <- false
            else field.Append ch |> ignore
        elif ch = '"' && atFieldStart then
            quoted <- true
            atFieldStart <- false
        elif ch = '\t' then endField ()
        elif ch = '\r' then
            if i + 1 < text.Length && text.[i + 1] = '\n' then i <- i + 1
            endRow ()
        elif ch = '\n' then endRow ()
        else
            field.Append ch |> ignore
            atFieldStart <- false
        i <- i + 1
    if field.Length > 0 || row.Count > 0 then endRow ()
    rows.ToArray()


type private Step = { Description: string; Owner: SheetId; Before: int; After: int
                      BeforeWorkbook: Workbook; AfterWorkbook: Workbook }

/// Persisted sheet identity used by both session revisions and host change
/// notifications. Persistent maps and lists are compared by identity, without
/// traversing cell maps or materializing deferred sheets.
let sameSheetContent (x: SpreadsheetModel) (y: SpreadsheetModel) =
    obj.ReferenceEquals(x, y) ||
    (obj.ReferenceEquals(x.Cells, y.Cells) && obj.ReferenceEquals(x.RichRuns, y.RichRuns)
     && obj.ReferenceEquals(x.RowHeights, y.RowHeights) && obj.ReferenceEquals(x.ColumnWidths, y.ColumnWidths)
     && obj.ReferenceEquals(x.BestFitColumns, y.BestFitColumns)
     && x.DefaultRowHeight = y.DefaultRowHeight
     && x.DefaultColumnWidth = y.DefaultColumnWidth
     && sameAxes x.Axes y.Axes && obj.ReferenceEquals(x.Validations, y.Validations)
     && (obj.ReferenceEquals(x.MergedCells, y.MergedCells) || x.MergedCells = y.MergedCells)
     && obj.ReferenceEquals(x.ConditionalFormats, y.ConditionalFormats)
     && obj.ReferenceEquals(x.Tables, y.Tables) && obj.ReferenceEquals(x.Filters, y.Filters)
     && x.AutoFilter = y.AutoFilter
     && (obj.ReferenceEquals(x.FilterHiddenRows, y.FilterHiddenRows) || x.FilterHiddenRows = y.FilterHiddenRows)
     && (obj.ReferenceEquals(x.PageSetup, y.PageSetup) || x.PageSetup = y.PageSetup)
     && obj.ReferenceEquals(x.ArrayMarks, y.ArrayMarks)
     && obj.ReferenceEquals(x.CfBoldCells, y.CfBoldCells)
     && x.Date1904 = y.Date1904
     && x.StructureVersion = y.StructureVersion
     && x.FrozenRows = y.FrozenRows && x.FrozenCols = y.FrozenCols
     && x.ShowGridLines = y.ShowGridLines && x.ShowHeaders = y.ShowHeaders
     && x.ShowFormulas = y.ShowFormulas && x.ShowZeros = y.ShowZeros
     && x.TabColor = y.TabColor && x.VeryHidden = y.VeryHidden
     && obj.ReferenceEquals(x.Drawings, y.Drawings) && obj.ReferenceEquals(x.Threads, y.Threads)
     && (obj.ReferenceEquals(x.Hyperlinks, y.Hyperlinks) || x.Hyperlinks = y.Hyperlinks) && x.Macros = y.Macros
     && obj.ReferenceEquals(x.Pivots, y.Pivots)
     && obj.ReferenceEquals(x.AuthoredCharts, y.AuthoredCharts)
     && obj.ReferenceEquals(x.ColumnStyles, y.ColumnStyles) && obj.ReferenceEquals(x.RowStyles, y.RowStyles)
     && obj.ReferenceEquals(x.DataTables, y.DataTables)
     && obj.ReferenceEquals(x.KeptRanges, y.KeptRanges)
     && x.SourceSheet = y.SourceSheet && x.KeptFrom = y.KeptFrom
     && x.StructureMove = y.StructureMove
     && obj.ReferenceEquals(x.StructureLog, y.StructureLog)
     && obj.ReferenceEquals(x.RowsSorted, y.RowsSorted))

/// A workbook session owns the calculated state, undo journal and save identity.
/// Workbook records and sheet cell maps are persistent; history never copies cell maps.
type Session(initial: Workbook, ?historyBySheet: bool) =
    let mutable workbook = initial
    let mutable committed = initial
    let mutable calculated: Workbook option = None
    let mutable engine: Engine option = None
    let mutable incremental = true
    let mutable epoch = 0
    let mutable revision = 0
    let mutable nextRevision = 0
    let mutable epochStartRevision = 0
    let mutable savedRevision = 0
    let mutable serial = 0
    let mutable kind = "none"
    let mutable reason = ""
    let mutable stats: RecalcStats = { ChangedCells = 0; ReachedFormulas = 0; EvaluatedFormulas = 0; Evaluations = 0 }
    let past = ResizeArray<Step>()
    let future = ResizeArray<Step>()
    let issuedSheetIds = System.Collections.Generic.HashSet<SheetId>()
    let mutable lastRuleId = 0
    let mutable lastValidationId = 0
    let mutable clipboard: ClipboardCapture option = None
    // Only deleted sources leave a tombstone; consumed cuts remain ordinary text.
    let mutable deletedClipboardText: string option = None
    let mutable cutRevision: int option = None
    let clearClipboard () =
        clipboard <- None
        deletedClipboardText <- None
        cutRevision <- None
    let updateClipboard (before: Workbook) (after: Workbook) =
        match clipboard with
        | None -> ()
        | Some captured when not (Map.containsKey captured.SourceSheet after.Sheets) ->
            deletedClipboardText <- Some captured.Text
            clipboard <- None
            cutRevision <- None
        | Some captured ->
            let mutable cells = captured.Data.Cells
            for KeyValue(id, oldSheet) in before.Sheets do
                match Map.tryFind id after.Sheets with
                | Some next when oldSheet.Name <> next.Name ->
                    cells <- cells |> List.map (fun (row, col, cell) ->
                        let value =
                            match cell.Value with
                            | Formula (formula, cached) ->
                                Formula (rewriteFormulaForSheets (RenameSheetRefs (oldSheet.Name, next.Name)) formula, cached)
                            | other -> other
                        row, col, { cell with Value = value })
                | _ -> ()
            if not (obj.ReferenceEquals(cells, captured.Data.Cells)) then
                clipboard <- Some { captured with Data = { captured.Data with Cells = cells } }
    let reserveModelRuleIds (model: SpreadsheetModel) =
        lastRuleId <- max lastRuleId (SheetOps.highestConditionalFormatId model)
        lastValidationId <- max lastValidationId (nextValidationId model - 1)
    let reserveRuleIds (book: Workbook) =
        for KeyValue(_, sheet) in book.Sheets do reserveModelRuleIds sheet.Model
    let reserveChangedRuleIds (previous: Workbook) (next: Workbook) =
        if not (obj.ReferenceEquals(previous.Sheets, next.Sheets)) then
            for KeyValue(id, sheet) in previous.Sheets do
                match Map.tryFind id next.Sheets with
                | Some current ->
                    if not (obj.ReferenceEquals(sheet.Model.ConditionalFormats, current.Model.ConditionalFormats)) then
                        lastRuleId <- max lastRuleId (max (SheetOps.highestConditionalFormatId sheet.Model) (SheetOps.highestConditionalFormatId current.Model))
                    if not (obj.ReferenceEquals(sheet.Model.Validations, current.Model.Validations)) then
                        lastValidationId <- max lastValidationId (max (nextValidationId sheet.Model - 1) (nextValidationId current.Model - 1))
                | None -> reserveModelRuleIds sheet.Model
            for KeyValue(id, sheet) in next.Sheets do
                if not (Map.containsKey id previous.Sheets) then reserveModelRuleIds sheet.Model
    let reserveSheets (book: Workbook) =
        for KeyValue(id, _) in book.Sheets do issuedSheetIds.Add id |> ignore
    let reserveNewOrder (previous: Workbook) (next: Workbook) =
        if previous.Sheets.Count <> next.Sheets.Count
           || (not (obj.ReferenceEquals(previous.SheetOrder, next.SheetOrder))
               && previous.SheetOrder <> next.SheetOrder) then reserveSheets next
    do reserveSheets initial
    do reserveRuleIds initial
    let historyBySheet = defaultArg historyBySheet false
    // Persistent step identities describe a saved document even when an
    // older edit on one sheet is undone after edits on another.
    let mutable applied = Set.empty<int>
    let mutable savedCheckpoint: Set<int> option = None
    let captures = ResizeArray<int * Set<int>>()
    let captureRevision () =
        if historyBySheet && (captures.Count = 0 || fst captures.[captures.Count - 1] <> revision) then
            captures.Add(revision, applied)
            if captures.Count > 200 then captures.RemoveAt 0
        revision
    let indexFor (steps: ResizeArray<Step>) =
        if historyBySheet then
            let mutable index = steps.Count - 1
            while index >= 0 && steps.[index].Owner <> workbook.ActiveSheet do index <- index - 1
            index
        else steps.Count - 1
    let currentEpoch () = epoch
    let currentRevision () = revision
    let fresh () =
        nextRevision <- nextRevision + 1
        nextRevision
    let calculate (hint: (SheetId * CellAddress) list option) (beforeFull: (unit -> Workbook) option) =
        let available =
            if not incremental then None
            else
                match engine, calculated with
                | Some current, _ -> Some current
                | None, Some previous ->
                    let built = build previous
                    engine <- Some built
                    Some built
                | None, None -> None
        let result =
            match available with
            | Some current ->
                match recalculate current workbook hint with
                | Some updated ->
                    kind <- "incremental"
                    stats <- lastStats current
                    Some updated
                | None ->
                    reason <- lastDecline current
                    None
            | None ->
                reason <- if incremental then "no calculated state to start from" else "incremental recalculation is switched off"
                None
        match result with
        | Some updated -> workbook <- updated; calculated <- Some updated
        | None ->
            match beforeFull with
            | Some prepare -> workbook <- prepare()
            | None -> ()
            let updated, built = recalculateAll workbook
            workbook <- updated
            calculated <- Some updated
            engine <- if incremental then Some built else None
            kind <- "full"
    let changed (previous: Workbook) (next: Workbook) =
        let content (a: Sheet) (b: Sheet) =
            a.Name = b.Name && a.Visible = b.Visible && sameSheetContent a.Model b.Model
        previous.SheetOrder <> next.SheetOrder || not (sameMap previous.Names next.Names)
        || not (sameMap previous.NameInfo next.NameInfo) || not (sameMap previous.Tables next.Tables)
        || previous.Iteration <> next.Iteration || previous.Sheets.Count <> next.Sheets.Count
        || (next.Sheets |> Map.exists (fun id sheet ->
            match Map.tryFind id previous.Sheets with Some old -> not (content old sheet) | None -> true))
    let recordFor owner (description: string) (previous: Workbook) =
        updateClipboard previous workbook
        reserveNewOrder previous workbook
        let before = revision
        revision <- fresh ()
        if historyBySheet then applied <- Set.add revision applied
        serial <- serial + 1
        if historyBySheet then
            for index = future.Count - 1 downto 0 do
                if future.[index].Owner = owner then future.RemoveAt index
        else future.Clear()
        past.Add { Description = description; Owner = owner
                   Before = before; After = revision
                   BeforeWorkbook = previous; AfterWorkbook = workbook }
        committed <- workbook
        if past.Count > 200 then
            let discarded = past.[0].After
            past.RemoveAt 0
            if historyBySheet then
                applied <- Set.remove discarded applied
                savedCheckpoint <- savedCheckpoint |> Option.map (Set.remove discarded)
    let record description previous = recordFor workbook.ActiveSheet description previous
    let changeSheet id (change: SpreadsheetModel -> SpreadsheetModel) (wb: Workbook) =
        match Map.tryFind id wb.Sheets with
        | Some sheet -> { wb with Sheets = Map.add id { sheet with Model = change sheet.Model } wb.Sheets }
        | None -> wb
    let loadClipboardSheet id =
        match Map.tryFind id workbook.Sheets with
        | Some sheet when sheet.Model.Deferred.IsSome ->
            let loaded = resolveSheet sheet.Model
            let loaded = { loaded with TabColor = sheet.Model.TabColor; VeryHidden = sheet.Model.VeryHidden
                                       Selection = sheet.Model.Selection; EditingCell = sheet.Model.EditingCell
                                       EditBuffer = sheet.Model.EditBuffer; Zoom = sheet.Model.Zoom }
            let replace wb = changeSheet id (fun _ -> loaded) wb
            workbook <- replace workbook
            committed <- replace committed
            calculated <- calculated |> Option.map replace
            engine <- None
        | _ -> ()
    let finishClipboard owner description previous beforeFull afterCalculate =
        if changed previous workbook then
            calculate None beforeFull
            match afterCalculate with
            | Some callback ->
                let updated = callback workbook
                reserveChangedRuleIds workbook updated
                reserveNewOrder workbook updated
                if changed workbook updated then
                    workbook <- updated
                    calculate None None
            | None -> ()
            let contentChanged = changed previous workbook
            if contentChanged then recordFor owner description previous
            contentChanged
        else false

    let restoreSnapshot (step: Step) (snapshot: Workbook) restoreActiveSheet =
        // Hosts choose whether history navigates to the edited sheet. Content,
        // history and revision restoration are identical under either view policy.
        let restoreActive =
            restoreActiveSheet ||
            step.BeforeWorkbook.SheetOrder <> step.AfterWorkbook.SheetOrder ||
            step.BeforeWorkbook.ActiveSheet <> step.AfterWorkbook.ActiveSheet ||
            not (Map.containsKey workbook.ActiveSheet snapshot.Sheets)
        let sheets =
            snapshot.Sheets |> Map.map (fun id sheet ->
                match Map.tryFind id workbook.Sheets with
                | None -> sheet
                | Some current ->
                    let view = current.Model
                    { sheet with
                        Model =
                            { sheet.Model with
                                Selection = view.Selection
                                EditingCell = view.EditingCell
                                EditBuffer = view.EditBuffer
                                Clipboard = view.Clipboard
                                History = view.History
                                SearchQuery = view.SearchQuery
                                SearchCaseSensitive = view.SearchCaseSensitive
                                SearchResults = view.SearchResults
                                SearchIndex = view.SearchIndex
                                FilterVisibleRows = None
                                Zoom = view.Zoom } })
        { snapshot with Sheets = sheets
                        ActiveSheet = if restoreActive then snapshot.ActiveSheet else workbook.ActiveSheet }
    let patchWorkbook (fromBook: Workbook) (toBook: Workbook) (current: Workbook) =
        let choose was becomes now =
            if was = becomes || now <> was then now else becomes
        let chooseRef was becomes now =
            if obj.ReferenceEquals(was, becomes) || not (obj.ReferenceEquals(now, was)) then now else becomes
        let patchMap same (was: Map<'k, 'v>) (becomes: Map<'k, 'v>) (now: Map<'k, 'v>) =
            if obj.ReferenceEquals(was, becomes) then now
            else
                let mutable result = now
                let put key old next =
                    if not (same old next) && same (Map.tryFind key result) old then
                        result <- match next with Some value -> Map.add key value result | None -> Map.remove key result
                for KeyValue(key, value) in was do put key (Some value) (Map.tryFind key becomes)
                for KeyValue(key, value) in becomes do
                    if not (Map.containsKey key was) then put key None (Some value)
                result
        let eqOption a b = a = b
        let sameCell (a: CellData option) (b: CellData option) =
            match a, b with
            | Some old, Some next when old.Format = next.Format && old.Comment = next.Comment ->
                match old.Value, next.Value with
                | CellValue.Formula (f, _), CellValue.Formula (g, _) -> f = g
                | _ -> old.Value = next.Value
            | _ -> a = b
        let patchModel (was: SpreadsheetModel) (becomes: SpreadsheetModel) (now: SpreadsheetModel) =
            { now with
                Cells = patchMap sameCell was.Cells becomes.Cells now.Cells
                RichRuns = patchMap eqOption was.RichRuns becomes.RichRuns now.RichRuns
                RowHeights = patchMap eqOption was.RowHeights becomes.RowHeights now.RowHeights
                ColumnWidths = patchMap eqOption was.ColumnWidths becomes.ColumnWidths now.ColumnWidths
                BestFitColumns = chooseRef was.BestFitColumns becomes.BestFitColumns now.BestFitColumns
                DefaultRowHeight = choose was.DefaultRowHeight becomes.DefaultRowHeight now.DefaultRowHeight
                DefaultColumnWidth = choose was.DefaultColumnWidth becomes.DefaultColumnWidth now.DefaultColumnWidth
                Axes = choose was.Axes becomes.Axes now.Axes
                Validations = chooseRef was.Validations becomes.Validations now.Validations
                MergedCells = chooseRef was.MergedCells becomes.MergedCells now.MergedCells
                ConditionalFormats = chooseRef was.ConditionalFormats becomes.ConditionalFormats now.ConditionalFormats
                Tables = chooseRef was.Tables becomes.Tables now.Tables
                Filters = patchMap eqOption was.Filters becomes.Filters now.Filters
                AutoFilter = choose was.AutoFilter becomes.AutoFilter now.AutoFilter
                FilterHiddenRows = chooseRef was.FilterHiddenRows becomes.FilterHiddenRows now.FilterHiddenRows
                PageSetup = choose was.PageSetup becomes.PageSetup now.PageSetup
                ArrayMarks = patchMap eqOption was.ArrayMarks becomes.ArrayMarks now.ArrayMarks
                CfBoldCells = chooseRef was.CfBoldCells becomes.CfBoldCells now.CfBoldCells
                Date1904 = choose was.Date1904 becomes.Date1904 now.Date1904
                StructureVersion = choose was.StructureVersion becomes.StructureVersion now.StructureVersion
                FrozenRows = choose was.FrozenRows becomes.FrozenRows now.FrozenRows
                FrozenCols = choose was.FrozenCols becomes.FrozenCols now.FrozenCols
                ShowGridLines = choose was.ShowGridLines becomes.ShowGridLines now.ShowGridLines
                ShowHeaders = choose was.ShowHeaders becomes.ShowHeaders now.ShowHeaders
                ShowFormulas = choose was.ShowFormulas becomes.ShowFormulas now.ShowFormulas
                ShowZeros = choose was.ShowZeros becomes.ShowZeros now.ShowZeros
                TabColor = choose was.TabColor becomes.TabColor now.TabColor
                VeryHidden = choose was.VeryHidden becomes.VeryHidden now.VeryHidden
                Drawings = chooseRef was.Drawings becomes.Drawings now.Drawings
                Threads = patchMap eqOption was.Threads becomes.Threads now.Threads
                Hyperlinks = chooseRef was.Hyperlinks becomes.Hyperlinks now.Hyperlinks
                Macros = choose was.Macros becomes.Macros now.Macros
                Pivots = chooseRef was.Pivots becomes.Pivots now.Pivots
                AuthoredCharts = patchMap eqOption was.AuthoredCharts becomes.AuthoredCharts now.AuthoredCharts
                ColumnStyles = chooseRef was.ColumnStyles becomes.ColumnStyles now.ColumnStyles
                RowStyles = chooseRef was.RowStyles becomes.RowStyles now.RowStyles
                DataTables = chooseRef was.DataTables becomes.DataTables now.DataTables
                KeptRanges = chooseRef was.KeptRanges becomes.KeptRanges now.KeptRanges
                SourceSheet = choose was.SourceSheet becomes.SourceSheet now.SourceSheet
                KeptFrom = choose was.KeptFrom becomes.KeptFrom now.KeptFrom
                StructureMove = choose was.StructureMove becomes.StructureMove now.StructureMove
                StructureLog = chooseRef was.StructureLog becomes.StructureLog now.StructureLog
                RowsSorted = chooseRef was.RowsSorted becomes.RowsSorted now.RowsSorted }
        let mutable sheets = current.Sheets
        for KeyValue(id, previous) in fromBook.Sheets do
            let next = Map.tryFind id toBook.Sheets
            let existing = Map.tryFind id sheets
            match next, existing with
            | Some target, Some present ->
                let updated =
                    { present with Name = choose previous.Name target.Name present.Name
                                   Visible = choose previous.Visible target.Visible present.Visible
                                   Model = patchModel previous.Model target.Model present.Model }
                sheets <- Map.add id updated sheets
            | None, Some _ -> sheets <- Map.remove id sheets
            | _ -> ()
        for KeyValue(id, next) in toBook.Sheets do
            if not (Map.containsKey id fromBook.Sheets) && not (Map.containsKey id sheets) then
                sheets <- Map.add id next sheets
        let order =
            if fromBook.SheetOrder = toBook.SheetOrder then current.SheetOrder
            else
                // Keep sheets introduced by later transactions while applying
                // the edited order to the sheets this operation touched.
                let additions = current.SheetOrder |> List.filter (fun id -> not (List.contains id fromBook.SheetOrder) && not (List.contains id toBook.SheetOrder))
                (toBook.SheetOrder @ additions) |> List.filter (fun id -> Map.containsKey id sheets)
        { current with
            Sheets = sheets
            SheetOrder = order
            ActiveSheet =
                if Map.containsKey current.ActiveSheet sheets then current.ActiveSheet
                else toBook.ActiveSheet
            Names = patchMap eqOption fromBook.Names toBook.Names current.Names
            NameInfo = patchMap eqOption fromBook.NameInfo toBook.NameInfo current.NameInfo
            Tables = patchMap eqOption fromBook.Tables toBook.Tables current.Tables
            Iteration = choose fromBook.Iteration toBook.Iteration current.Iteration }
    let rebaseRenamedSheets (source: Workbook) (target: Workbook) (current: Workbook) =
        (current, source.Sheets)
        ||> Map.fold (fun state id previous ->
            match Map.tryFind id target.Sheets, Map.tryFind id state.Sheets with
            | Some next, Some existing when previous.Name <> next.Name && existing.Name = previous.Name ->
                renameSheetEverywhere id next.Name state
            | _ -> state)
    let editOfMove inverse (move: StructureMove) =
        match move.OnRows, (if inverse then not move.Insert else move.Insert) with
        | true, true -> InsertRows (move.At, move.Count)
        | true, false -> DeleteRows (move.At, move.Count)
        | false, true -> InsertCols (move.At, move.Count)
        | false, false -> DeleteCols (move.At, move.Count)
    let rebaseStructure (source: Workbook) (target: Workbook) (current: Workbook) =
        (current, source.Sheets)
        ||> Map.fold (fun state id previous ->
            match Map.tryFind id target.Sheets, Map.tryFind id state.Sheets with
            | Some next, Some existing when abs (next.Model.StructureVersion - previous.Model.StructureVersion) = 1 ->
                let undo = previous.Model.StructureVersion > next.Model.StructureVersion
                let move = (if undo then previous.Model else next.Model).StructureMove
                match move with
                | Some move ->
                    let edit = editOfMove undo move
                    let rewritten, _ = Rbt.Spreadsheet.Workbook.applyStructureEdit id edit state
                    // The journal restores the edited sheet's deleted data;
                    // applyStructureEdit here rebases newer foreign references.
                    { rewritten with Sheets = Map.add id existing rewritten.Sheets }
                | None -> state
            | _ -> state)
    let alignHistoryContext (source: Workbook) (target: Workbook) (current: Workbook) =
        // The older transaction may have been recorded before a rename or
        // row move that has since been undone on another sheet. Rewrite BOTH
        // sides of its delta into the current reference/coordinate system;
        // applying that delta directly to current would mistake a rebased
        // formula for an unrelated edit and skip its undo.
        let mutable before = source
        let mutable after = target
        let both transform =
            before <- transform before
            after <- transform after
        for KeyValue(id, original) in source.Sheets do
            match Map.tryFind id current.Sheets with
            | None -> ()
            | Some now ->
                if original.Name <> now.Name then
                    both (renameSheetEverywhere id now.Name)
                let originalVersion = original.Model.StructureVersion
                let currentVersion = now.Model.StructureVersion
                if originalVersion > currentVersion then
                    for move in original.Model.StructureLog do
                        if move.Version > currentVersion then
                            both (fun wb -> Rbt.Spreadsheet.Workbook.applyStructureEdit id (editOfMove true move) wb |> fst)
                elif originalVersion < currentVersion then
                    for move in now.Model.StructureLog |> List.rev do
                        if move.Version > originalVersion then
                            both (fun wb -> Rbt.Spreadsheet.Workbook.applyStructureEdit id (editOfMove false move) wb |> fst)
        before, after
    let input (sheetId: SheetId) (addr: CellAddress) (text: string) (year: int)
              (accept: (CellData -> bool) option) (afterCalculate: (Workbook -> Workbook) option)
              (beforeFull: (unit -> Workbook) option) =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> false
        | Some sheet ->
            // Entry into a source-backed sheet loads only that sheet. Loading
            // is not an edit; the committed and calculated baselines see it too.
            let sheet =
                if sheet.Model.Deferred.IsNone then sheet
                else
                    let loaded = resolveSheet sheet.Model
                    let loaded =
                        if loaded.TabColor = sheet.Model.TabColor && loaded.VeryHidden = sheet.Model.VeryHidden then loaded
                        else { loaded with TabColor = sheet.Model.TabColor; VeryHidden = sheet.Model.VeryHidden }
                    let replace wb = changeSheet sheetId (fun _ -> loaded) wb
                    committed <- replace committed
                    workbook <- replace workbook
                    calculated <- calculated |> Option.map replace
                    engine <- None
                    { sheet with Model = loaded }
            let model = sheet.Model
            let previous = getCell addr model
            let clear = System.String.IsNullOrWhiteSpace text
            if clear && previous.IsNone then false
            else
                let parsed =
                    if clear then None
                    else
                        let cell = Rbt.Spreadsheet.Input.cellFromInputFor model.Date1904 year previous text
                        let cell =
                            match cell.Value, cell.Format.NumberFormat with
                            | CellValue.Formula (formula, _), (None | Some "" | Some "General") ->
                                let formatAt name row col =
                                    let target =
                                        match name with
                                        | None -> Some sheet
                                        | Some name -> workbook.Sheets |> Seq.tryPick (fun (KeyValue(_, s)) ->
                                            if s.Name.Equals(name, System.StringComparison.OrdinalIgnoreCase) then Some s else None)
                                    target |> Option.bind (fun s -> getCell { Row = row; Col = col } (resolveSheet s.Model))
                                    |> Option.bind (fun c -> c.Format.NumberFormat)
                                match (try Rbt.Spreadsheet.SheetOps.inheritedNumberFormat formatAt formula with _ -> None) with
                                | Some code -> { cell with Format = { cell.Format with NumberFormat = Some code } }
                                | None -> cell
                            | _ -> cell
                        Some cell
                let requested =
                    parsed |> Option.defaultValue
                        { Value = CellValue.Empty
                          Format = previous |> Option.map (fun old -> old.Format) |> Option.defaultValue defaultFormat
                          Comment = None }
                // Validation needs the prospective value, but the stored formula
                // must start cold: recalculation decides whether it can run or
                // must retain zero when this entry closes a circular reference.
                let rejected =
                    accept |> Option.exists (fun check ->
                        let candidate =
                            match requested.Value with
                            | CellValue.Formula (formula, _) ->
                                { requested with Value = CellValue.Formula(formula, Rbt.Spreadsheet.Formula.evaluateFormula formula model (Some workbook)) }
                            | _ -> requested
                        not (check candidate))
                if rejected then false
                else
                    let put value m =
                        match value with Some cell -> setCell addr cell m | None -> clearCell addr m
                    let txn = {
                        Description = sprintf "%s cell %s" (if clear then "Clear" else "Edit") (formatCellRef addr.Row addr.Col)
                        Apply = changeSheet sheetId (put parsed)
                        Recalculate = true
                        Hint = Some [ sheetId, addr ]
                        AfterCalculate = afterCalculate
                    }
                    let before = workbook
                    workbook <- txn.Apply workbook
                    calculate txn.Hint beforeFull
                    match txn.AfterCalculate with
                    | Some callback ->
                        let updated = callback workbook
                        reserveChangedRuleIds workbook updated
                        reserveNewOrder workbook updated
                        if changed workbook updated then
                            workbook <- updated
                            calculate None None
                    | None -> ()
                    if changed before workbook then record txn.Description before
                    true

    member _.Workbook = workbook
    member _.Clipboard = clipboard
    member _.ClearClipboard() = clearClipboard ()
    member _.ClipboardSource(text: string option) =
        let normalise (value: string) =
            if isNull value then "" else value.Replace("\r\n", "\n").TrimEnd('\n')
        match text with
        | None -> if clipboard.IsSome then InternalPaste else IgnoredPaste
        | Some value ->
            let text = normalise value
            if clipboard |> Option.exists (fun captured -> normalise captured.Text = text && captured.Text <> "") then InternalPaste
            elif deletedClipboardText |> Option.exists (fun old -> normalise old = text && old <> "") then IgnoredPaste
            elif System.String.IsNullOrEmpty value then IgnoredPaste
            else ExternalPaste

    member _.CaptureClipboard(sheetId: SheetId, isCut: bool, text: string,
                              ?textOf: ClipboardData -> string, ?transfer: string -> Result<unit, string>,
                              ?beforeFull: unit -> Workbook, ?afterCalculate: Workbook -> Workbook) : Result<bool, string> =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> Result.Error "no such sheet"
        | Some _ ->
            loadClipboardSheet sheetId
            let model = workbook.Sheets.[sheetId].Model
            let capturedWorkbook = workbook
            let capturedEpoch, capturedRevision = epoch, revision
            let data = serializeSelection model.Selection model isCut
            let text = match textOf with Some render -> render data | None -> text
            let transferResult = match transfer with Some send -> send text | None -> Ok ()
            match transferResult with
            | Result.Error problem -> Result.Error problem
            | Ok () when not (obj.ReferenceEquals(workbook, capturedWorkbook))
                         || epoch <> capturedEpoch || revision <> capturedRevision ->
                Result.Error "workbook changed during clipboard transfer"
            | Ok () ->
                clearClipboard ()
                clipboard <- Some { SourceSheet = sheetId; Data = data; Text = text }
                if not isCut then Ok false
                else
                    let previous = workbook
                    workbook <- changeSheet sheetId (clearSelectionCells model.Selection) workbook
                    let didChange = finishClipboard sheetId "Cut" previous beforeFull afterCalculate
                    if didChange then cutRevision <- Some revision
                    Ok didChange

    member _.PasteClipboard(sheetId: SheetId, ?beforeFull: unit -> Workbook,
                            ?afterCalculate: Workbook -> Workbook) : Result<bool, string> =
        match Map.tryFind sheetId workbook.Sheets, clipboard with
        | None, _ -> Result.Error "no such sheet"
        | _, None -> Ok false
        | Some _, Some captured when not (Map.containsKey captured.SourceSheet workbook.Sheets) ->
            deletedClipboardText <- Some captured.Text
            clipboard <- None
            cutRevision <- None
            Ok false
        | Some _, Some captured ->
            loadClipboardSheet sheetId
            if captured.Data.IsCut then
                let ids = workbook.Sheets |> Map.toList |> List.choose (fun (id, sheet) ->
                    match sheet.Model.Deferred with Some deferred when deferred.HasFormulas -> Some id | _ -> None)
                for id in ids do loadClipboardSheet id
            let previous = workbook
            let model = workbook.Sheets.[sheetId].Model
            let anchors = pasteAnchors model.Selection captured.Data
            let pasted = anchors |> List.fold (fun state anchor -> pasteClipboard anchor captured.Data state) model
            let sameCellContent (a: CellData option) (b: CellData option) =
                match a, b with
                | Some x, Some y when x.Format = y.Format && x.Comment = y.Comment ->
                    match x.Value, y.Value with
                    | Formula (left, _), Formula (right, _) -> left = right
                    | left, right -> left = right
                | None, None -> true
                | _ -> false
            let unchanged =
                anchors |> List.forall (fun anchor ->
                    captured.Data.Cells |> List.forall (fun (row, col, _) ->
                        let addr = { Row = anchor.Row + row; Col = anchor.Col + col }
                        sameCellContent (getCell addr model) (getCell addr pasted)))
            let pasted = if unchanged then model else pasted
            let pasted = { pasted with Selection = pasteSelection anchors captured.Data model.Selection }
            workbook <- changeSheet sheetId (fun _ -> pasted) workbook
            match anchors with
            | [ anchor ] when captured.Data.IsCut ->
                let data = captured.Data
                let source = data.SourceStartRow, data.SourceStartCol, data.SourceEndRow, data.SourceEndCol
                let delta = anchor.Row - data.SourceStartRow, anchor.Col - data.SourceStartCol
                workbook <- repointMovedCells captured.SourceSheet source sheetId delta workbook |> fst
            | _ -> ()
            let didChange = finishClipboard sheetId "Paste" previous beforeFull afterCalculate
            if not didChange && pasted.Selection = model.Selection then workbook <- previous
            if captured.Data.IsCut then clearClipboard ()
            Ok didChange

    member _.PasteRows(sheetId: SheetId, rows: string array array, year: int,
                       ?beforeFull: unit -> Workbook, ?afterCalculate: Workbook -> Workbook) : Result<bool, string> =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> Result.Error "no such sheet"
        | Some _ ->
            let rows = rows |> Array.filter (fun row -> row.Length > 0)
            if rows.Length = 0 then Ok false
            else
                loadClipboardSheet sheetId
                let previous = workbook
                let model = workbook.Sheets.[sheetId].Model
                let top, left, bottom, right =
                    match model.Selection with
                    | SingleCell (r, c) -> r, c, r, c
                    | Range (r1, c1, r2, c2) -> min r1 r2, min c1 c2, max r1 r2, max c1 c2
                let height = rows.Length
                let width = rows |> Array.map Array.length |> Array.max
                let isRange =
                    match model.Selection with
                    | SingleCell _ -> false
                    | Range _ ->
                        match findMergedCell { Row = top; Col = left } model with
                        | Some merge -> not (merge.StartRow = top && merge.StartCol = left && merge.EndRow = bottom && merge.EndCol = right)
                        | None -> top <> bottom || left <> right
                let repeat = height = 1 && width = 1 && isRange
                let endRow = if repeat then bottom else min 1048575 (top + height - 1)
                let endCol = if repeat then right else min 16383 (left + width - 1)
                let mutable next = { model with RowCount = max model.RowCount (endRow + 1); ColCount = max model.ColCount (endCol + 1) }
                for row in top .. endRow do
                    for col in left .. endCol do
                        let source = if repeat then rows.[0] else rows.[row - top]
                        let text = if repeat then source.[0] elif col - left < source.Length then source.[col - left] else ""
                        let addr = { Row = row; Col = col }
                        match findMergedCell addr next with
                        | Some merge when merge.StartRow <> row || merge.StartCol <> col -> ()
                        | _ ->
                            let existing = getCell addr next
                            if System.String.IsNullOrEmpty text then
                                match existing with
                                | Some cell when cell.Value <> CellValue.Empty -> next <- setCell addr { cell with Value = CellValue.Empty } next
                                | _ -> ()
                            else
                                let cell = Rbt.Spreadsheet.Input.cellFromInputFor next.Date1904 year existing text
                                if existing <> Some cell then next <- setCell addr cell next
                let mutable t, l, b, r = top, left, endRow, endCol
                let mutable grew = true
                while grew do
                    grew <- false
                    for merge in mergedCellsIntersectingRange t (b + 1) l (r + 1) next do
                        if merge.StartRow < t then (t <- merge.StartRow; grew <- true)
                        if merge.StartCol < l then (l <- merge.StartCol; grew <- true)
                        if merge.EndRow > b then (b <- merge.EndRow; grew <- true)
                        if merge.EndCol > r then (r <- merge.EndCol; grew <- true)
                let selection = if t = b && l = r then SingleCell (t, l) else Range (t, l, b, r)
                next <- { next with Selection = selection }
                workbook <- changeSheet sheetId (fun _ -> next) workbook
                let didChange = finishClipboard sheetId "Paste" previous beforeFull afterCalculate
                if not didChange && selection = model.Selection
                   && next.RowCount = model.RowCount && next.ColCount = model.ColCount then
                    workbook <- previous
                Ok didChange

    member _.CommittedWorkbook = committed
    member _.Revision = revision
    member _.SavedRevision = savedRevision
    member _.IsModified =
        if historyBySheet then
            match savedCheckpoint with Some saved -> applied <> saved | None -> revision <> savedRevision
        else revision <> savedRevision
    member _.Epoch = epoch
    member _.ChangeSerial = serial
    member _.CanUndo = indexFor past >= 0
    member _.CanRedo = indexFor future >= 0
    member _.UndoDescription =
        let index = indexFor past
        if index < 0 then None else Some past.[index].Description
    member _.RedoDescription =
        let index = indexFor future
        if index < 0 then None else Some future.[index].Description
    member _.CaptureRevision() = captureRevision ()
    member _.CaptureSave() =
        captureRevision () |> ignore
        { Epoch = epoch; Revision = revision }
    member _.MarkSaved(?revision: int, ?epoch: int) =
        // Browser callers capture a revision alone; revisions never repeat and
        // a completion from a replaced workbook cannot save the new workbook.
        let captured = defaultArg revision (currentRevision())
        if Option.forall ((=) (currentEpoch())) epoch
           && captured >= epochStartRevision && captured <= nextRevision then
            savedRevision <- captured
            if historyBySheet then
                let recorded =
                    if captured = currentRevision() then Some applied
                    else
                        let mutable found = None
                        let mutable index = captures.Count - 1
                        while found.IsNone && index >= 0 do
                            let known, checkpoint = captures.[index]
                            if known = captured then found <- Some checkpoint
                            index <- index - 1
                        found
                let retained =
                    seq { for step in past do yield step.After
                          for step in future do yield step.After } |> Set.ofSeq
                savedCheckpoint <- recorded |> Option.map (fun snapshot -> Set.intersect snapshot retained)
    member _.IncrementalRecalculation = incremental
    member _.SetIncrementalRecalculation(enabled: bool) =
        incremental <- enabled
        if not enabled then engine <- None
    member _.RecalcStats =
        { Kind = kind
          Reason = if kind = "full" then reason else ""
          Stats = stats
          Formulas = match engine with Some value -> formulaCount value | None -> -1 }
    member _.EnsureRecalcEngine() =
        if incremental && engine.IsNone then
            match calculated with Some wb -> engine <- Some (build wb) | None -> ()
    member _.Recalculate(?hint: (SheetId * CellAddress) list, ?beforeFull: unit -> Workbook) =
        calculate hint beforeFull
    member _.RecalculateFromScratch() =
        let updated, built = recalculateAll workbook
        workbook <- updated
        calculated <- Some updated
        engine <- if incremental then Some built else None
        kind <- "full"
        reason <- "the workbook was calculated from scratch"
    member _.InvalidateCalculation() = engine <- None; calculated <- None
    member _.RebaseCalculation() = engine <- None; calculated <- Some workbook
    member _.Replace(next: Workbook, ?markSaved: bool) =
        clearClipboard ()
        epoch <- epoch + 1
        workbook <- next
        committed <- next
        issuedSheetIds.Clear()
        lastRuleId <- 0
        lastValidationId <- 0
        reserveRuleIds next
        reserveSheets next
        calculated <- None
        engine <- None
        past.Clear()
        future.Clear()
        if historyBySheet then
            applied <- Set.empty
            captures.Clear()
            savedCheckpoint <- if defaultArg markSaved false then Some Set.empty else None
        serial <- serial + 1
        revision <- fresh ()
        if defaultArg markSaved false then savedRevision <- revision
        epochStartRevision <- revision
    member _.UpdateView(next: Workbook) =
        // Selection, edit buffer and display fields are projections; never dirty a file.
        reserveNewOrder workbook next
        workbook <- next
    member _.AbsorbView(next: Workbook) =
        reserveNewOrder workbook next
        workbook <- next
        committed <- next
    member _.ResolveDeferredSheet(id: SheetId, model: SpreadsheetModel) =
        let replace (wb: Workbook) =
            match Map.tryFind id wb.Sheets with
            | Some sheet when sheet.Model.Deferred.IsSome ->
                { wb with Sheets = Map.add id { sheet with Model = model } wb.Sheets }
            | _ -> wb
        committed <- replace committed
        workbook <- replace workbook
        reserveModelRuleIds model
        calculated <- calculated |> Option.map replace
        engine <- None
    member _.Observe(next: Workbook) =
        let old = committed
        workbook <- next
        committed <- next
        reserveChangedRuleIds old next
        updateClipboard old next
        reserveNewOrder old next
        if changed old next then
            serial <- serial + 1
            revision <- fresh ()
            future.Clear()
            past.Clear()
            if historyBySheet then
                applied <- Set.empty
                captures.Clear()
                savedCheckpoint <- None
            true
        else false
    member _.Apply(txn: Transaction) =
        let previous = workbook
        workbook <- txn.Apply workbook
        reserveChangedRuleIds previous workbook
        reserveNewOrder previous workbook
        if txn.Recalculate then calculate txn.Hint None
        match txn.AfterCalculate with
        | Some callback ->
            let updated = callback workbook
            reserveChangedRuleIds workbook updated
            reserveNewOrder workbook updated
            if changed workbook updated then
                workbook <- updated
                calculate None None
        | None -> ()
        if changed previous workbook then record txn.Description previous
    member _.CommitApplied(txn: AppliedTransaction) =
        let previous = committed
        workbook <- txn.Workbook
        reserveChangedRuleIds previous txn.Workbook
        reserveNewOrder previous txn.Workbook
        if changed previous txn.Workbook then
            calculated <- Some txn.Workbook
            record txn.Description previous
        else committed <- txn.Workbook
    /// Apply a selection command as one workbook journal entry.
    member _.ExecuteFormatting(sheetId: SheetId, command: FormattingCommand,
                               ?bounds: (int * int * int * int), ?activeCell: CellAddress,
                               ?description: string, ?beforeFull: unit -> Workbook) : Result<bool, string> =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> Result.Error "no such sheet"
        | Some sheet ->
            let model = sheet.Model
            let sr, sc, er, ec =
                match bounds, model.Selection with
                | Some block, _ -> block
                | None, SingleCell (r, c) -> r, c, r, c
                | None, Range (r1, c1, r2, c2) -> r1, c1, r2, c2
            let top, left = max 0 (min sr er), max 0 (min sc ec)
            let bottom, right = min (model.RowCount - 1) (max sr er), min (model.ColCount - 1) (max sc ec)
            let block = top, left, bottom, right
            let anchor = defaultArg activeCell { Row = sr; Col = sc }
            let active = getCell anchor model |> Option.map (fun cell -> cell.Format) |> Option.defaultValue defaultFormat
            let toggle get set =
                let on = not (get active)
                fun (format: CellFormat) -> if get format = on then format else set on format
            let update, label =
                match command with
                | ToggleBold -> toggle (fun f -> f.Bold) (fun v f -> { f with Bold = v }), "Bold"
                | ToggleItalic -> toggle (fun f -> f.Italic) (fun v f -> { f with Italic = v }), "Italic"
                | ToggleUnderline -> toggle (fun f -> f.Underline) (fun v f -> if v then { f with Underline = true } else { f with Underline = false; UnderlineDouble = false }), "Underline"
                | ToggleDoubleUnderline -> toggle (fun f -> f.UnderlineDouble) (fun v f -> if v then { f with Underline = true; UnderlineDouble = true } else { f with UnderlineDouble = false }), "Double underline"
                | ToggleStrikethrough -> toggle (fun f -> f.Strikethrough) (fun v f -> { f with Strikethrough = v }), "Strikethrough"
                | ToggleWrapText -> toggle (fun f -> f.WrapText) (fun v f -> { f with WrapText = v }), "Wrap text"
                | ToggleShrinkToFit -> toggle (fun f -> f.ShrinkToFit) (fun v f -> { f with ShrinkToFit = v }), "Shrink to fit"
                | ChangeFontSize n ->
                    (fun f ->
                        if n = 0 then f
                        else
                            let size = defaultArg f.FontSize 11.0 + float n
                            { f with FontSize = Some (if n > 0 then min 409.0 size else max 8.0 size) }), "Font size"
                | ChangeIndent n -> (fun f -> { f with Indent = max 0 (min 250 (f.Indent + n)) }), "Indent"
                | HorizontalAlign value -> (fun f -> { f with HorizontalAlign = value }), "Horizontal alignment"
                | VerticalAlign value -> (fun f -> { f with VerticalAlign = value }), "Vertical alignment"
                | NumberFormat value -> (fun f -> { f with NumberFormat = value }), "Number format"
                | TransformFormat transform -> transform, "Format"
                | PaintFormat wanted ->
                    (fun f -> { wanted with FilterButton = f.FilterButton; HasComment = f.HasComment; HasListDropdown = f.HasListDropdown; DataBar = f.DataBar; IconSet = f.IconSet }), "Paint format"
                | Border _ -> id, "Border"
                | ChangeDecimals _ -> id, "Decimals"
                | ClearCells _ -> id, "Clear cells"
                | MergeCells -> id, "Merge cells"
                | UnmergeCells _ -> id, "Unmerge cells"
            let problem =
                match command with
                | NumberFormat (Some code) -> SheetOps.numberFormatProblem code
                | _ -> None
            match problem with
            | Some error -> Result.Error error
            | None ->
                let next =
                    match command with
                    | ClearCells what -> SheetOps.clearCells what block model
                    | MergeCells when top <= bottom && left <= right ->
                        if model.MergedCells |> List.exists (fun merge -> merge.StartRow = top && merge.StartCol = left && merge.EndRow = bottom && merge.EndCol = right) then model
                        else addMergedCell top left bottom right model
                    | MergeCells -> model
                    | UnmergeCells addr -> removeMergedCell addr model
                    | _ ->
                        let mutable cells = model.Cells
                        let change (addr: CellAddress) (existing: CellData option) =
                            let data =
                                match existing with
                                | Some cell -> cell
                                | None -> { Value = CellValue.Empty; Format = defaultFormat; Comment = None }
                            let format =
                                match command with
                                | Border (side, style, color) -> borderFormatInBlock block side style color addr.Row addr.Col data.Format
                                | ChangeDecimals n ->
                                    match existing |> Option.bind (fun cell -> SheetOps.changeDecimals n data.Format.NumberFormat cell.Value) with
                                    | Some code -> { data.Format with NumberFormat = Some code }
                                    | None -> data.Format
                                | _ -> update data.Format
                            if not (obj.ReferenceEquals(format, data.Format)) && format <> data.Format then
                                cells <- Map.add addr { data with Format = format } cells
                        let area = float (bottom - top + 1) * float (right - left + 1)
                        if top <= bottom && left <= right then
                            if area <= 20000.0 then
                                for row in top .. bottom do
                                    if not (isFilterHiddenRow row model) then
                                        for col in left .. right do
                                            let addr = { Row = row; Col = col }
                                            change addr (Map.tryFind addr model.Cells)
                            else
                                for KeyValue(addr, data) in model.Cells do
                                    if addr.Row >= top && addr.Row <= bottom && addr.Col >= left && addr.Col <= right && not (isFilterHiddenRow addr.Row model) then
                                        change addr (Some data)
                        if obj.ReferenceEquals(cells, model.Cells) then model else { model with Cells = cells } |> invalidateFilterProjection
                if sameSheetContent model next then Ok false
                else
                    let previous = workbook
                    workbook <- { workbook with Sheets = Map.add sheetId { sheet with Model = next } workbook.Sheets }
                    match command with
                    | ClearCells what when what <> SheetOps.ClearFormats -> calculate None beforeFull
                    | _ -> ()
                    recordFor sheetId (defaultArg description label) previous
                    Ok true

    /// Defined-name changes affect formulas throughout the workbook.
    member _.ExecuteName(command: NameCommand, ?description: string, ?beforeFull: unit -> Workbook) : Result<bool, string> =
        let scope, rawName =
            match command with
            | SetDefinedName (scope, name, _, _) | DeleteDefinedName (scope, name) -> scope, name
        let name = if isNull rawName then "" else rawName.Trim()
        let problem =
            if scope |> Option.exists (fun id -> not (Map.containsKey id workbook.Sheets)) then Some "no such sheet"
            elif name.ToUpperInvariant().StartsWith "_XLNM." then Some (sprintf "\"%s\" is a built-in name" name)
            else
                match command with
                | SetDefinedName (_, _, refersTo, _) ->
                    match definedNameProblem name with
                    | Some error -> Some error
                    | None when workbook.Tables |> Map.exists (fun known _ -> System.String.Equals(known, name, System.StringComparison.OrdinalIgnoreCase)) ->
                        Some (sprintf "\"%s\" is the name of a table" name)
                    | None -> SheetOps.refersToProblem refersTo
                | DeleteDefinedName _ -> None
        match problem with
        | Some error -> Result.Error error
        | None ->
            let before = tryFindDefinedName scope name workbook
            let next =
                match command with
                | SetDefinedName (_, _, refersTo, comment) ->
                    let text = if isNull refersTo then "" else refersTo.Trim()
                    Some (setDefinedName scope name text (comment |> Option.filter (fun value -> not (System.String.IsNullOrEmpty value))) workbook)
                | DeleteDefinedName _ -> deleteDefinedName scope name workbook
            match next with
            | None ->
                match scope with
                | Some id -> Result.Error (sprintf "the sheet \"%s\" has no name \"%s\"" workbook.Sheets.[id].Name name)
                | None -> Result.Error (sprintf "the workbook has no name \"%s\" (a sheet-level name needs its sheet index)" name)
            | Some next when before = tryFindDefinedName scope name next -> Ok false
            | Some next ->
                let previous = workbook
                workbook <- next
                calculate None beforeFull
                let label = match command with SetDefinedName _ -> "Define name" | DeleteDefinedName _ -> "Delete name"
                record (defaultArg description label) previous
                Ok true

    /// Value-neutral drawing edits are committed together with authored parts.
    member _.ExecuteDrawing(sheetId: SheetId, command: DrawingCommand, ?description: string) : Result<DrawingCommandResult, string> =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> Result.Error "no such sheet"
        | Some sheet ->
            let loaded = resolveSheet sheet.Model
            let model =
                if loaded.TabColor = sheet.Model.TabColor && loaded.VeryHidden = sheet.Model.VeryHidden then loaded
                else { loaded with TabColor = sheet.Model.TabColor; VeryHidden = sheet.Model.VeryHidden }
            try
                let finite n = not (System.Double.IsNaN n || System.Double.IsInfinity n)
                let geometry (cell: CellAddress) x y width height =
                    if cell.Row < 0 || cell.Col < 0 || cell.Row >= model.RowCount || cell.Col >= model.ColCount then
                        failwith "the anchor cell must fit within the current worksheet"
                    if not (finite x && finite y && finite width && finite height) || x < 0.0 || y < 0.0 || width <= 0.0 || height <= 0.0 then
                        failwith "the picture offsets and size must be finite, with nonnegative offsets and positive size"
                let mutable id = None
                let next, label =
                    match command with
                    | InsertPicture (data, cell, width, height, name) ->
                        geometry cell 0.0 0.0 width height
                        let info = SheetOps.imageInfo data |> Option.defaultWith (fun () -> failwith "the picture is not a PNG, JPEG or GIF image")
                        let number = SheetOps.nextDrawingId model
                        id <- Some number
                        let name = name |> Option.filter (fun text -> not (System.String.IsNullOrWhiteSpace text))
                        SheetOps.insertPicture number info data cell width height name model, "Insert picture"
                    | PlacePicture (number, cell, x, y, width, height) ->
                        geometry cell x y width height
                        (SheetOps.placePicture number cell x y width height model |> Result.defaultWith failwith), "Move picture"
                    | DeletePicture number ->
                        (SheetOps.deletePicture number model |> Result.defaultWith failwith), "Delete picture"
                    | DescribePicture (number, name, description) ->
                        (SheetOps.describePicture number name description model |> Result.defaultWith failwith), "Picture text"
                    | InsertChart (area, kind, title, headers, categories, reserved) ->
                        if model.EditingCell.IsSome then failwith "commit or cancel the cell entry before inserting a chart"
                        if area.From.Row < 0 || area.From.Col < 0 || area.From.Row > area.To.Row || area.From.Col > area.To.Col
                           || area.To.Row >= model.RowCount || area.To.Col >= model.ColCount then
                            failwith "the source range must fit within the current worksheet"
                        if SheetOps.isWholeRows area || SheetOps.isWholeColumns area then failwith "choose a bounded rectangular cell range"
                        if isNull title || title.Length > 255 then failwith "the title exceeds 255 characters"
                        let firstRow = area.From.Row + (if headers then 1 else 0)
                        let firstCol = area.From.Col + (if categories then 1 else 0)
                        let count, series = area.To.Row - firstRow + 1, area.To.Col - firstCol + 1
                        if count < 1 || series < 1 then failwith "the range needs at least one data row and one value column after headers/categories"
                        if kind = ChartAuthoring.Pie && series <> 1 then failwith "a pie chart needs exactly one value column; choose a narrower range"
                        if float count * float series > 20000.0 then failwith "the selected data exceeds 20,000 chart values"
                        let read row col =
                            let address = { Row = row; Col = col }
                            let value = unwrapFormulaValue (getCellValue address model)
                            let number = match value with CellValue.Number n when finite n -> Some n | _ -> None
                            let text =
                                match getCell address model with
                                | Some data -> Rbt.Spreadsheet.CellText.cellDisplayTextIn model.Date1904 None 1.0e6 data.Format value
                                | None -> ""
                            text, number
                        let mutable hasNumbers = false
                        let mutable row = firstRow
                        while not hasNumbers && row <= area.To.Row do
                            let mutable col = firstCol
                            while not hasNumbers && col <= area.To.Col do
                                match unwrapFormulaValue (getCellValue { Row = row; Col = col } model) with
                                | CellValue.Number n when finite n -> hasNumbers <- true
                                | _ -> ()
                                col <- col + 1
                            row <- row + 1
                        if not hasNumbers then failwith "the source has no numeric values"
                        let number = SheetOps.nextDrawingId model
                        id <- Some number
                        let used = System.Collections.Generic.HashSet<string>(reserved)
                        for KeyValue(rel, _) in model.AuthoredCharts do used.Add rel |> ignore
                        for drawing in model.Drawings do
                            match drawing.Content with
                            | KeptObject (_, xml) ->
                                for found in System.Text.RegularExpressions.Regex.Matches(xml, "rIdChart[0-9]+") do
                                    used.Add(found.Value) |> ignore
                            | _ -> ()
                        let mutable n = 1
                        while used.Contains (sprintf "rIdChart%d" n) do n <- n + 1
                        let rel = sprintf "rIdChart%d" n
                        let name = sprintf "Chart %d" number
                        let col = if area.To.Col + 1 < model.ColCount then area.To.Col + 1 else area.From.Col
                        let cell = { Row = area.From.Row; Col = col }
                        let width, height = 480.0, 300.0
                        let anchorXml = ChartAuthoring.standardChartAnchorXml number rel name cell.Row cell.Col width height
                        let chartXml = ChartAuthoring.standardChartXml kind title sheet.Name area.From.Row area.From.Col area.To.Row area.To.Col headers categories read
                        let drawing = { Id = number; Anchor = SheetOps.pictureAnchor cell 0.0 0.0 width height
                                        Name = Some name; Description = None; Content = KeptObject ("chart", anchorXml) }
                        { model with Drawings = model.Drawings @ [ drawing ]; AuthoredCharts = Map.add rel chartXml model.AuthoredCharts }, "Insert chart"
                let changed = not (obj.ReferenceEquals(model.Drawings, next.Drawings)) || not (obj.ReferenceEquals(model.AuthoredCharts, next.AuthoredCharts))
                if changed then
                    let previous =
                        if sheet.Model.Deferred.IsSome then
                            let replace wb = changeSheet sheetId (fun _ -> model) wb
                            committed <- replace committed
                            calculated <- calculated |> Option.map replace
                            engine <- None
                            replace workbook
                        else workbook
                    workbook <- { previous with Sheets = Map.add sheetId { sheet with Model = next } previous.Sheets }
                    recordFor sheetId (defaultArg description label) previous
                Ok { Changed = changed; Id = id }
            with error -> Result.Error error.Message

    /// Value-neutral rule changes share the workbook journal and never recalculate.
    member _.ExecuteRule(sheetId: SheetId, command: RuleCommand, ?description: string) : Result<RuleCommandResult, string> =
        match Map.tryFind sheetId workbook.Sheets with
        | None -> Result.Error "no such sheet"
        | Some sheet ->
            let loaded = resolveSheet sheet.Model
            let model =
                if loaded.TabColor = sheet.Model.TabColor && loaded.VeryHidden = sheet.Model.VeryHidden then loaded
                else { loaded with TabColor = sheet.Model.TabColor; VeryHidden = sheet.Model.VeryHidden }
            let validRanges ranges =
                not (List.isEmpty ranges) &&
                (ranges |> List.forall (fun r ->
                    r.From.Row >= 0 && r.From.Col >= 0 &&
                    r.From.Row <= r.To.Row && r.From.Col <= r.To.Col &&
                    r.To.Row < 1048576 && r.To.Col < 16384))
            let rangesProblem =
                match command with
                | AddConditionalFormat (ranges, _, _) | SetValidation (ranges, _) | ClearValidation ranges -> not (validRanges ranges)
                | UpdateConditionalFormat (_, Some ranges, _, _, _) | ClearConditionalFormats (Some ranges) -> not (validRanges ranges)
                | _ -> false
            if rangesProblem then Result.Error "invalid or empty ranges"
            else
                try
                    let nextId issued highest = max issued highest + 1
                    let mutable issuedRule = None
                    let mutable issuedValidation = None
                    let next, label =
                        match command with
                        | AddConditionalFormat (ranges, stop, format) ->
                            SheetOps.ruleOfDescription "AddConditionalFormat" format.Style (SheetOps.describeRule format) |> ignore
                            let highest = workbook.Sheets |> Map.fold (fun value _ s -> max value (SheetOps.highestConditionalFormatId s.Model)) (SheetOps.highestConditionalFormatId model)
                            let id = nextId lastRuleId highest
                            issuedRule <- Some id
                            SheetOps.addConditionalFormatWithId id ranges stop format model, "Conditional format"
                        | UpdateConditionalFormat (id, ranges, stop, format, priority) ->
                            format |> Option.iter (fun value -> SheetOps.ruleOfDescription "UpdateConditionalFormat" value.Style (SheetOps.describeRule value) |> ignore)
                            let updated = SheetOps.updateConditionalFormat "UpdateConditionalFormat" id (fun entry ->
                                { entry with Ranges = defaultArg ranges entry.Ranges; StopIfTrue = defaultArg stop entry.StopIfTrue; Format = defaultArg format entry.Format }) model
                            (match priority with Some place -> SheetOps.moveConditionalFormat "UpdateConditionalFormat" id place updated | None -> updated), "Edit conditional format"
                        | RemoveConditionalFormat id -> SheetOps.removeConditionalFormat "RemoveConditionalFormat" id model, "Remove conditional format"
                        | ReorderConditionalFormats ids -> SheetOps.reorderConditionalFormats "ReorderConditionalFormats" ids model, "Reorder conditional formats"
                        | ClearConditionalFormats ranges -> SheetOps.clearConditionalFormats ranges model, "Clear conditional formats"
                        | SetValidation (ranges, validation) ->
                            SheetOps.validationOfDescriptionIn model.Date1904 "SetValidation" (SheetOps.describeValidation validation) |> ignore
                            if model.Validations |> List.exists (fun entry -> entry.Ranges = ranges && entry.Validation = validation) then
                                model, "Data validation"
                            else
                                let highest = workbook.Sheets |> Map.fold (fun value _ s -> max value (nextValidationId s.Model - 1)) (nextValidationId model - 1)
                                let id = nextId lastValidationId highest
                                issuedValidation <- Some id
                                setValidationRangesWithId id ranges validation model, "Data validation"
                        | ClearValidation ranges -> clearValidationRanges ranges model, "Clear validation"
                    let changed = model.ConditionalFormats <> next.ConditionalFormats || model.Validations <> next.Validations
                    let affected =
                        match command with
                        | ClearConditionalFormats _ ->
                            model.ConditionalFormats |> List.filter (fun entry -> next.ConditionalFormats |> List.exists (fun kept -> kept.Id = entry.Id && kept.Ranges = entry.Ranges) |> not) |> List.length
                        | ClearValidation _ ->
                            model.Validations |> List.filter (fun entry -> next.Validations |> List.exists (fun kept -> kept.Id = entry.Id && kept.Ranges = entry.Ranges) |> not) |> List.length
                        | _ -> if changed then 1 else 0
                    if changed then
                        let previous =
                            if sheet.Model.Deferred.IsSome then
                                let replace wb = changeSheet sheetId (fun _ -> model) wb
                                committed <- replace committed
                                calculated <- calculated |> Option.map replace
                                engine <- None
                                replace workbook
                            else workbook
                        reserveModelRuleIds model
                        workbook <- { previous with Sheets = Map.add sheetId { sheet with Model = next } previous.Sheets }
                        issuedRule |> Option.iter (fun id -> lastRuleId <- id)
                        issuedValidation |> Option.iter (fun id -> lastValidationId <- id)
                        recordFor sheetId (defaultArg description label) previous
                    let resultId =
                        match issuedRule, issuedValidation, command with
                        | Some id, _, _ | _, Some id, _ -> Some id
                        | _, _, SetValidation (ranges, validation) ->
                            model.Validations |> List.tryFind (fun entry -> entry.Ranges = ranges && entry.Validation = validation) |> Option.map (fun entry -> entry.Id)
                        | _ -> None
                    Ok { Changed = changed; Id = resultId; Affected = affected }
                with error -> Result.Error error.Message

    /// Execute one structural history transaction. By default, commands use the
    /// browser's calculation policy (rename and copy rebase cached calculations);
    /// hosts may request eager recalculation without adding another history step.
    member _.ExecuteStructural(command: StructuralCommand, ?description: string,
                               ?afterCalculate: Workbook -> Workbook, ?beforeFull: unit -> Workbook,
                               ?recalculate: bool) : Result<bool, string> =
        let missing id = if Map.containsKey id workbook.Sheets then None else Some "no such sheet"
        let validateName except name =
            match sheetNameProblem name with
            | Some problem -> Some problem
            | None when sheetNameTaken except name workbook -> Some (sprintf "there is already a sheet named \"%s\"" name)
            | None -> None
        let prepared =
            match command with
            | EditStructure (id, edit) ->
                let at, count =
                    match edit with
                    | InsertRows (at, count) | DeleteRows (at, count)
                    | InsertCols (at, count) | DeleteCols (at, count) -> at, count
                match missing id with
                | Some problem -> Result.Error problem
                | None when at < 0 || count < 0 -> Result.Error "row or column index and count must not be negative"
                | None when count = 0 -> Ok (workbook, "Edit structure", false, false)
                | None -> Ok (Rbt.Spreadsheet.Workbook.applyStructureEdit id edit workbook |> fst, "Edit structure", true, false)
            | AddSheet (name, position, rows, columns) ->
                let name = match name with Some value -> value | None -> nextSheetName workbook
                let at = defaultArg position workbook.SheetOrder.Length
                match validateName None name with
                | Some problem -> Result.Error problem
                | None when at < 0 || at > workbook.SheetOrder.Length -> Result.Error (sprintf "atIndex %d is outside 0 to %d" at workbook.SheetOrder.Length)
                | None when rows <= 0 || columns <= 0 -> Result.Error "sheet dimensions must be positive"
                | None ->
                    let id = freshSheetId issuedSheetIds.Contains workbook
                    let date1904 = workbook.Sheets.[workbook.ActiveSheet].Model.Date1904
                    let sheet = { Id = id; Name = name; Model = { createEmpty rows columns with Date1904 = date1904 }; Visible = true }
                    Ok (insertSheet sheet at workbook, "Add sheet", true, false)
            | RenameSheet (id, name) ->
                match missing id with
                | Some problem -> Result.Error problem
                | None ->
                    match validateName (Some id) name with
                    | Some problem -> Result.Error problem
                    | None -> Ok (renameSheetEverywhere id name workbook, "Rename sheet", false, true)
            | DeleteSheet id ->
                match lastVisibleSheetProblem id workbook with
                | Some problem -> Result.Error problem
                | None ->
                    match removeSheet id workbook with
                    | Some (next, _) -> Ok (next, "Delete sheet", true, false)
                    | None -> Result.Error "cannot delete the last sheet"
            | MoveSheet (id, at) ->
                match missing id with
                | Some problem -> Result.Error problem
                | None when at < 0 || at >= workbook.SheetOrder.Length -> Result.Error (sprintf "to %d is outside 0 to %d" at (workbook.SheetOrder.Length - 1))
                | None -> Ok (moveSheet id at workbook, "Move sheet", true, false)
            | DuplicateSheet id ->
                match Map.tryFind id workbook.Sheets with
                | None -> Result.Error "no such sheet"
                | Some sheet ->
                    let freshId = freshSheetId issuedSheetIds.Contains workbook
                    let name = copySheetName sheet.Name workbook
                    Ok (duplicateSheet id freshId name (sheetIndex id workbook + 1) workbook, "Copy sheet", false, true)
            | SetSheetVisibility (id, visible, veryHidden) ->
                match Map.tryFind id workbook.Sheets with
                | None -> Result.Error "no such sheet"
                | Some sheet ->
                    match if not visible && sheet.Visible then lastVisibleSheetProblem id workbook else None with
                    | Some problem -> Result.Error problem
                    | None -> Ok (setSheetVisibility id visible veryHidden workbook, (if visible then "Unhide sheet" else "Hide sheet"), false, false)
        match prepared with
        | Result.Error problem -> Result.Error problem
        | Ok (next, defaultDescription, recalc, rebase) ->
            let previous = workbook
            if not (changed previous next) then Ok false
            else
                workbook <- next
                reserveChangedRuleIds previous workbook
                reserveNewOrder previous workbook
                if defaultArg recalculate recalc then calculate None beforeFull
                elif rebase then engine <- None; calculated <- Some workbook
                match afterCalculate with
                | Some callback ->
                    let updated = callback workbook
                    reserveChangedRuleIds workbook updated
                    reserveNewOrder workbook updated
                    if changed workbook updated then
                        workbook <- updated
                        calculate None None
                | _ -> ()
                record (defaultArg description defaultDescription) previous
                Ok true
    member _.EnterCell(sheetId, address, text, year, ?accept: CellData -> bool,
                       ?afterCalculate: Workbook -> Workbook, ?beforeFull: unit -> Workbook) =
        input sheetId address text year accept afterCalculate beforeFull
    member _.ClearCell(sheetId, address, ?accept: CellData -> bool,
                       ?afterCalculate: Workbook -> Workbook, ?beforeFull: unit -> Workbook) =
        input sheetId address "" 0 accept afterCalculate beforeFull
    member _.Undo(?restoreActiveSheet: bool) =
        let index = indexFor past
        if index < 0 then false
        else
            let step = past.[index]
            let before = workbook
            if cutRevision = Some step.After then clearClipboard ()
            let atTip = index = past.Count - 1 && revision = step.After
            past.RemoveAt index
            if historyBySheet then applied <- Set.remove step.After applied
            if atTip then
                workbook <- restoreSnapshot step step.BeforeWorkbook (defaultArg restoreActiveSheet false)
                if not (sameMap before.Names workbook.Names) then calculate None None
                else
                    calculated <- Some workbook
                    engine <- None
                revision <- step.Before
            else
                let previous = workbook
                let expected, target = alignHistoryContext step.AfterWorkbook step.BeforeWorkbook previous
                let renamed = rebaseRenamedSheets step.AfterWorkbook step.BeforeWorkbook previous
                let rebased = rebaseStructure step.AfterWorkbook step.BeforeWorkbook renamed
                workbook <- patchWorkbook expected target rebased
                calculated <- Some previous
                engine <- None
                calculate None None
                revision <- fresh ()
            updateClipboard before workbook
            if historyBySheet && savedCheckpoint = Some applied then revision <- savedRevision
            committed <- workbook
            future.Add { step with Owner = workbook.ActiveSheet }
            serial <- serial + 1
            true
    member _.Redo(?restoreActiveSheet: bool) =
        let index = indexFor future
        if index < 0 then false
        else
            let step = future.[index]
            let before = workbook
            let atTip = index = future.Count - 1 && revision = step.Before
            future.RemoveAt index
            if historyBySheet then applied <- Set.add step.After applied
            if atTip then
                workbook <- restoreSnapshot step step.AfterWorkbook (defaultArg restoreActiveSheet false)
                if not (sameMap before.Names workbook.Names) then calculate None None
                else
                    calculated <- Some workbook
                    engine <- None
                revision <- step.After
            else
                let previous = workbook
                let expected, target = alignHistoryContext step.BeforeWorkbook step.AfterWorkbook previous
                let renamed = rebaseRenamedSheets step.BeforeWorkbook step.AfterWorkbook previous
                let rebased = rebaseStructure step.BeforeWorkbook step.AfterWorkbook renamed
                workbook <- patchWorkbook expected target rebased
                calculated <- Some previous
                engine <- None
                calculate None None
                revision <- fresh ()
            updateClipboard before workbook
            if historyBySheet && savedCheckpoint = Some applied then revision <- savedRevision
            committed <- workbook
            past.Add { step with Owner = workbook.ActiveSheet }
            serial <- serial + 1
            true
