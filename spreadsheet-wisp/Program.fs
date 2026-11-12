module SpreadsheetWispApp

open System
open System.IO
open Rbt.Renderer
open Wisp.Reactive
open Wisp.Layout
open Rbt.Office.Ui

module Ui = Wisp.Atrium
module Shell = Wisp.OfficeShell
module XlUi = Rbt.Spreadsheet.Ui
module M = Rbt.Spreadsheet.Model
module Session = Rbt.Spreadsheet.Session
module L = Rbt.Spreadsheet.Layout
module EH = Rbt.SpreadsheetWisp.EditorHost
module Pivot = Rbt.Spreadsheet.Pivot
module CF = Rbt.Spreadsheet.ConditionalFormatting

let private appW = 1200
let private appH = 820

let private existingFile path = if File.Exists path then Some path else None

let private recentFilesPath =
    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".config/redbat-spreadsheet-wisp/recent.txt")

let private loadRecentFiles () =
    if File.Exists recentFilesPath then
        try File.ReadAllLines recentFilesPath |> Array.toList |> List.filter (fun s -> s.Trim() <> "")
        with _ -> []
    else []

let private saveRecentFiles (list: string list) =
    try
        Directory.CreateDirectory(Path.GetDirectoryName recentFilesPath) |> ignore
        File.WriteAllLines(recentFilesPath, list |> List.toArray)
    with _ -> ()

let private recentFiles = refSignal (loadRecentFiles ())

let private rememberRecent path =
    let abs = Path.GetFullPath path
    let next = abs :: (recentFiles.Value |> List.filter ((<>) abs)) |> List.truncate 12
    recentFiles.Value <- next
    saveRecentFiles next

module Bridge = Rbt.Spreadsheet.XlsxBridge

/// Loading and saving XLSX. The conversion between the reader / writer and
/// the editor's models is `Rbt.Spreadsheet.XlsxBridge`, shared with the web
/// component; this host adds file access and the VBA project.
module XlsxRead =
    /// The file the open workbook was imported from: a save carries over from
    /// it what the editor does not model (drawings, named styles, theme, ...).
    let mutable source : Bridge.XlsxSource option = None

    let load path : Rbt.Spreadsheet.Workbook.Workbook option =
        try
            let bytes = File.ReadAllBytes path
            match XlsxReader.loadBytesNative bytes with
            | None -> None
            | Some wb when wb.Sheets.IsEmpty -> None
            | Some wb ->
                let macros = VbaProject.extractFromXlsx bytes |> Option.defaultValue ""
                let imported = Bridge.Import.importWorkbookWith Bridge.editorImportOptions wb
                let context =
                    if String.IsNullOrWhiteSpace macros then imported.Context
                    else
                        { imported.Context with Sheets = imported.Context.Sheets |> Map.map (fun _ sheet -> { sheet with Model = { sheet.Model with Macros = macros } }) }
                let sheets = context.SheetOrder |> List.map (fun id -> context.Sheets.[id].Model) |> List.toArray
                source <- Some { imported.Source with Baseline = sheets }
                Some context
        with ex ->
            eprintfn "[spreadsheet-wisp] load failed: %s" ex.Message
            None

module XlsxWrite =
    let save st path =
        File.WriteAllBytes(path, XlsxWriter.writeWorkbookPackageBytes (EH.exportWorkbook st XlsxRead.source))

let private sampleWorkbook () =
    let put r c v m = M.setCell ({ Row = r; Col = c } : M.CellAddress) ({ Value = v; Format = M.defaultFormat; Comment = None } : M.CellData) m
    let model =
        M.createEmpty 200 40
        |> put 0 0 (M.String "Product")
        |> put 0 1 (M.String "Units")
        |> put 0 2 (M.String "Price")
        |> put 0 3 (M.String "Total")
        |> put 1 0 (M.String "Alpha")
        |> put 1 1 (M.Number 12.0)
        |> put 1 2 (M.Number 4.25)
        |> put 1 3 (M.Formula("B2*C2", M.Number 51.0))
        |> put 2 0 (M.String "Beta")
        |> put 2 1 (M.Number 8.0)
        |> put 2 2 (M.Number 9.5)
        |> put 2 3 (M.Formula("B3*C3", M.Number 76.0))
        |> M.createTable "Table1" 0 0 3 3 true
    [| model |], [| "Sheet1" |]

let private cellFormatAt st =
    let model = EH.activeModel st
    let r, c = EH.focusCell model
    M.getCell ({ Row = r; Col = c } : M.CellAddress) model |> Option.map (fun d -> d.Format) |> Option.defaultValue M.defaultFormat

let private format st command =
    EH.executeFormatting st command |> ignore

let private updateFormat st update =
    format st (Session.TransformFormat update)

let private parseRange (s: string) =
    match s.Split(':') with
    | [| a; b |] ->
        match M.parseCellRef (a.Trim()), M.parseCellRef (b.Trim()) with
        | Some (r1, c1), Some (r2, c2) -> Some(min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    | [| a |] -> M.parseCellRef (a.Trim()) |> Option.map (fun (r, c) -> r, c, r, c)
    | _ -> None

let private selfTest () =
    let st = EH.create [| M.createEmpty 20 10 |] [| "Sheet1" |] 0
    EH.applyCellEdit st 0 0 "2" |> ignore
    EH.applyCellEdit st 0 1 "=A1*5" |> ignore
    let b1 = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st)
    let formulaOk = match b1 with M.Formula(_, M.Number 10.0) -> true | _ -> false
    EH.handleKey st { Key = "ArrowRight"; Shift = false; Ctrl = false; Alt = false } |> ignore
    let navOk = EH.focusCell (EH.activeModel st) = (0, 1)
    EH.handleKey st { Key = "Delete"; Shift = false; Ctrl = false; Alt = false } |> ignore
    let deleted = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) = M.Empty
    EH.undo st |> ignore
    let restored = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) <> M.Empty
    EH.redo st |> ignore
    let redone = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) = M.Empty
    if formulaOk && navOk && deleted && restored && redone then
        printfn "spreadsheet-wisp selftest ok"
        0
    else
        eprintfn "selftest failed formula=%b nav=%b deleted=%b restored=%b redone=%b" formulaOk navOk deleted restored redone
        1

[<EntryPoint>]
let main argv =
    if Array.contains "--selftest" argv then selfTest () else
    let xlsxPath = argv |> Array.tryFind (fun a -> a.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
    let st =
        match xlsxPath |> Option.bind XlsxRead.load with
        | Some loaded ->
            xlsxPath |> Option.iter rememberRecent
            EH.createWithWorkbook loaded
        | None ->
            let sheets, names = sampleWorkbook ()
            EH.create sheets names 0
    let shell = Shell.create ()
    let status = refSignal "Ready"
    let currentPath = refSignal xlsxPath
    let zoom = signal "100"

    let title () =
        let name = currentPath.Value |> Option.map Path.GetFileName |> Option.defaultValue "Book1"
        sprintf "%s - wisp-spreadsheet" name

    let writeTo path =
        let captured = EH.captureSave st
        try
            XlsxWrite.save st path
            EH.markSaved st captured
            currentPath.Value <- Some path
            rememberRecent path
            status.Value <- sprintf "Saved %s" (Path.GetFileName path)
        with ex -> status.Value <- "Save failed: " + ex.Message

    let loadFromPath path =
        match XlsxRead.load path with
        | Some workbook ->
            EH.replaceWorkbook st workbook
            currentPath.Value <- Some path
            rememberRecent path
            status.Value <- sprintf "Opened %s" (Path.GetFileName path)
        | None -> status.Value <- sprintf "Open failed: %s" path

    let openFindReplaceDialog () =
        XlUi.findReplaceDialog
            (fun q matchCase ->
                if q = "" then status.Value <- "Enter a search term"; ""
                else
                    let hits = M.searchCells q matchCase (EH.activeModel st)
                    if hits.Length = 0 then status.Value <- sprintf "No matches for \"%s\"" q
                    else
                        let target = hits.[0]
                        EH.applyNoHistory st (fun m -> { m with Selection = M.SingleCell(target.Row, target.Col); SearchQuery = Some q; SearchCaseSensitive = Some matchCase; SearchResults = hits; SearchIndex = 0 }) |> ignore
                        status.Value <- sprintf "Match 1 of %d for \"%s\"" hits.Length q
                    "")
            (fun q r matchCase ->
                let next, count = M.replaceAllWithCount q r matchCase (EH.activeModel st)
                EH.applyChange st (fun _ -> next) |> ignore
                status.Value <- sprintf "Replaced %d occurrence(s)" count
                "")
        |> Shell.openDialog shell

    let openFilterDialog () =
        let model = EH.activeModel st
        let r, c = EH.focusCell model
        let uniques, label =
            match M.tableAt r c model with
            | Some t -> M.tableColumnUniqueValues t c model, sprintf "%s (table %s)" (M.columnToLetter c) t.Name
            | None -> M.columnUniqueValues c model, M.columnToLetter c
        if List.isEmpty uniques then status.Value <- "No values to filter in this column."
        else
            XlUi.filterDialog label uniques (fun visible ->
                EH.applyChange st (fun m -> if Set.count visible = List.length uniques then M.removeFilter c m else M.applyFilter c visible m) |> ignore
                status.Value <- sprintf "Filter on %s: %d of %d shown" label (Set.count visible) (List.length uniques))
            |> Shell.openDialog shell

    let openInsertTableDialog () =
        XlUi.insertTableDialog (fun rows cols ->
            let model = EH.activeModel st
            let r, c = EH.focusCell model
            let name = sprintf "Table%d" (List.length model.Tables + 1)
            EH.applyChange st (M.createTable name r c (r + rows) (c + cols - 1) true) |> ignore
            status.Value <- sprintf "Inserted table %s" name)
        |> Shell.openDialog shell

    let openMacroDialog () =
        let model = EH.activeModel st
        XlUi.macroDialog model.Macros (fun source ->
            let sheetName = EH.activeSheetName st
            let mutable msg = "Macro ran"
            EH.applyChange st (fun m ->
                match Rbt.Spreadsheet.Vba.runOnModelWithSheetName sheetName { m with Macros = source } source with
                | Result.Ok m2 -> m2
                | Result.Error e -> msg <- "Macro error: " + e; { m with Macros = source }) |> ignore
            status.Value <- msg)
        |> Shell.openDialog shell

    let openPivotDialog () =
        let cellText v =
            match M.unwrapFormulaValue v with
            | M.Number n -> n.ToString("G", Globalization.CultureInfo.InvariantCulture)
            | M.String s -> s
            | M.Boolean b -> if b then "TRUE" else "FALSE"
            | M.Error e -> "#" + e
            | M.Empty -> ""
            | M.Formula _ -> ""
        let preview (req: XlUi.PivotUiRequest) =
            match parseRange req.Source with
            | None -> "Enter a source range like A1:C100."
            | Some (sr, sc, er, ec) ->
                let model = EH.activeModel st
                let grid = [| for r in sr .. er -> [| for c in sc .. ec -> M.getCellValue { Row = r; Col = c } model |] |]
                let headers = if grid.Length > 0 then grid.[0] else [||]
                let rowFields = req.RowFields.Split(',') |> Array.map _.Trim() |> Array.filter ((<>) "") |> Array.toList
                let colOpt = if req.ColField = "" then None else Some req.ColField
                let agg = Pivot.parseAgg req.Agg |> Option.defaultValue Pivot.Sum
                let formula = match Pivot.buildFormula sr sc er ec headers rowFields colOpt req.ValueField req.Agg with Result.Ok f -> f | Result.Error _ -> ""
                match Pivot.pivotFromNames grid rowFields colOpt req.ValueField agg with
                | Result.Error m -> m
                | Result.Ok p ->
                    let rows = p |> Array.truncate 5 |> Array.map (fun row -> row |> Array.truncate 4 |> Array.map cellText |> String.concat "    ") |> String.concat "\n"
                    (if formula <> "" then formula + "\n" else "") + rows
        XlUi.pivotTableDialog preview (fun req ->
            match parseRange req.Source, M.parseCellRef req.Dest with
            | Some (sr, sc, er, ec), Some (dr, dc) ->
                let model = EH.activeModel st
                let headers = [| for c in sc .. ec -> M.getCellValue { Row = sr; Col = c } model |]
                let rowFields = req.RowFields.Split(',') |> Array.map _.Trim() |> Array.filter ((<>) "") |> Array.toList
                let colOpt = if req.ColField = "" then None else Some req.ColField
                match Pivot.buildFormula sr sc er ec headers rowFields colOpt req.ValueField req.Agg with
                | Result.Ok formula ->
                    EH.applyChange st (M.setCell { Row = dr; Col = dc } { Value = M.Formula(formula, M.Empty); Format = M.defaultFormat; Comment = None }) |> ignore
                    status.Value <- sprintf "Inserted pivot formula at %s" (M.formatCellRef dr dc)
                    true
                | Result.Error msg -> status.Value <- "Pivot: " + msg; false
            | _ -> status.Value <- "Pivot: invalid range or destination"; false)
        |> Shell.openDialog shell

    let openInsertSheetDialog () =
        XlUi.insertSheetDialog (fun name ->
            match EH.addSheet st name with
            | Ok () -> status.Value <- sprintf "Added sheet \"%s\"" name
            | Error problem -> status.Value <- problem)
        |> Shell.openDialog shell

    let openRenameSheetDialog () =
        let initial = EH.activeSheetName st
        XlUi.renameSheetDialog initial (fun name ->
            match EH.renameSheet st name with
            | Ok () -> status.Value <- sprintf "Renamed sheet to \"%s\"" name
            | Error problem -> status.Value <- problem)
        |> Shell.openDialog shell

    let actions : XlUi.ShellActions =
        let fmt () = cellFormatAt st
        { PrintPdf = (fun () -> status.Value <- "PDF export not wired in Wisp app")
          OpenFindReplace = openFindReplaceDialog
          Cut = (fun () -> EH.cutSelection st |> ignore; status.Value <- "Cut")
          Copy = (fun () -> EH.copySelection st |> ignore; status.Value <- "Copied")
          Paste = (fun () -> if EH.pasteFromClipboard st then status.Value <- "Pasted" else status.Value <- "Clipboard empty")
          Undo = (fun () -> EH.undo st |> ignore)
          Redo = (fun () -> EH.redo st |> ignore)
          ToggleBold = (fun () -> format st Session.ToggleBold)
          ToggleItalic = (fun () -> format st Session.ToggleItalic)
          ToggleUnderline = (fun () -> format st Session.ToggleUnderline)
          ToggleStrike = (fun () -> format st Session.ToggleStrikethrough)
          ToggleWrapText = (fun () -> format st Session.ToggleWrapText)
          ClearFormatting = (fun () -> format st (Session.ClearCells Rbt.Spreadsheet.SheetOps.ClearFormats))
          ApplyHorizontalAlign = (fun v -> format st (Session.HorizontalAlign (Some(if v = "right" || v = "center" then v else "left"))))
          ApplyVerticalAlign = (fun v -> format st (Session.VerticalAlign (Some(if v = "top" || v = "middle" then v else "bottom"))))
          ApplyNumberFormat = (fun v -> format st (Session.NumberFormat (Some v)))
          IncreaseDecimals = (fun () -> status.Value <- "Increase decimals not wired")
          DecreaseDecimals = (fun () -> status.Value <- "Decrease decimals not wired")
          ApplyFontFamily = (fun v -> updateFormat st (fun f -> { f with FontName = Some v }))
          ApplyFontSize = (fun v -> match Int32.TryParse v with true, n -> updateFormat st (fun f -> { f with FontSize = Some n }) | _ -> ())
          ApplyTextColor = (fun v -> updateFormat st (fun f -> { f with TextColor = Some v }))
          ApplyBgColor = (fun v -> updateFormat st (fun f -> { f with BgColor = Some v }))
          ApplyBorderStyle = (fun side ->
              let edge, style =
                  if side = "thick-outside" then "outside", "thick"
                  else side, "thin"
              format st (Session.Border (edge, style, "#000000")))
          ApplyZoom = (fun v -> zoom.Value <- v)
          MergeCells = (fun () ->
              match (EH.activeModel st).Selection with
              | M.Range _ -> format st Session.MergeCells
              | _ -> status.Value <- "Select a range to merge")
          UnmergeCells = (fun () ->
              let r, c = EH.focusCell (EH.activeModel st)
              format st (Session.UnmergeCells { Row = r; Col = c }))
          InsertRow = (fun () -> let r, _ = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.InsertRows (r, 1)) |> ignore)
          InsertColumn = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.InsertCols (c, 1)) |> ignore)
          DeleteRow = (fun () -> let r, _ = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.DeleteRows (r, 1)) |> ignore)
          DeleteColumn = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.DeleteCols (c, 1)) |> ignore)
          SortAscending = (fun () -> let r, c = EH.focusCell (EH.activeModel st) in EH.applyChange st (M.sortRegionByColumn { Row = r; Col = c } true) |> ignore)
          SortDescending = (fun () -> let r, c = EH.focusCell (EH.activeModel st) in EH.applyChange st (M.sortRegionByColumn { Row = r; Col = c } false) |> ignore)
          ToggleFilter = openFilterDialog
          ClearFilter = (fun () -> EH.applyChange st (fun m -> M.clearColumnFilter None m) |> ignore; status.Value <- "Filters cleared")
          ToggleFreezePanes = (fun () -> EH.applyChange st (fun m -> if m.FrozenRows > 0 || m.FrozenCols > 0 then { m with FrozenRows = 0; FrozenCols = 0 } else let r, c = EH.focusCell m in { m with FrozenRows = r; FrozenCols = c }) |> ignore)
          SplitPanes = (fun () -> status.Value <- "Split panes not wired")
          InsertChart = (fun () -> status.Value <- "Insert chart not wired")
          InsertPivot = openPivotDialog
          InsertTable = openInsertTableDialog
          RunMacro = openMacroDialog
          InsertSheet = openInsertSheetDialog
          DeleteSheet = (fun () ->
              match EH.deleteSheet st with
              | Ok () -> status.Value <- "Deleted sheet"
              | Error problem -> status.Value <- problem)
          RenameSheet = openRenameSheetDialog
          SwitchSheet = (fun name -> st.SheetNames.Value |> Array.tryFindIndex ((=) name) |> Option.iter (EH.switchSheet st))
          EditFormula = EH.editFormula st
          SubmitFormula = (fun () -> EH.submitFormula st) }

    let shellState : XlUi.ShellState =
        let fmt () = cellFormatAt st
        { ActiveCellRef = (fun () -> EH.activeCellRef st)
          FormulaText = (fun () -> EH.formulaText st)
          IsBoldActive = (fun () -> (fmt()).Bold)
          IsItalicActive = (fun () -> (fmt()).Italic)
          IsUnderlineActive = (fun () -> (fmt()).Underline)
          IsStrikeActive = (fun () -> (fmt()).Strikethrough)
          IsWrapTextActive = (fun () -> (fmt()).WrapText)
          IsMergedActive = (fun () -> let r, c = EH.focusCell (EH.activeModel st) in M.findMergedCell { Row = r; Col = c } (EH.activeModel st) |> Option.isSome)
          IsFilterActive = (fun () -> not (Map.isEmpty (EH.activeModel st).Filters))
          IsFreezePanesActive = (fun () -> let m = EH.activeModel st in m.FrozenRows > 0 || m.FrozenCols > 0)
          HorizontalAlign = (fun () -> (fmt()).HorizontalAlign |> Option.defaultValue "left")
          VerticalAlign = (fun () -> (fmt()).VerticalAlign |> Option.defaultValue "bottom")
          NumberFormat = (fun () -> (fmt()).NumberFormat |> Option.defaultValue "general")
          TextColor = (fun () -> (fmt()).TextColor |> Option.defaultValue "#000000")
          BgColor = (fun () -> (fmt()).BgColor |> Option.defaultValue "#FFFFFF")
          FontFamily = (fun () -> (fmt()).FontName |> Option.defaultValue "Calibri")
          FontSize = (fun () -> sprintf "%d" ((fmt()).FontSize |> Option.defaultValue 11))
          BorderStyle = (fun () -> "none")
          ZoomLevel = (fun () -> zoom.Value)
          ActiveSheet = (fun () -> st.SheetNames.Value |> Array.tryItem st.ActiveSheet.Value |> Option.defaultValue "")
          Sheets = (fun () -> st.SheetNames.Value |> Array.map (fun n -> XlUi.choice n n) |> Array.toList)
          SelectionSummary = (fun () -> EH.selectionSummary st)
          CanUndo = (fun () -> EH.canUndo st)
          CanRedo = (fun () -> EH.canRedo st) }

    let menus =
        let fileMenu =
            { TestId = "xl-file-open-save"
              PanelTestId = "xl-file-open-save-panel"
              Label = "File"
              Items =
                [ menuItemShortcut "xl-menu-open" "Open..." "Ctrl+O" None (fun () -> XlUi.openFileDialog loadFromPath |> Shell.openDialog shell)
                  menuItemShortcut "xl-menu-save" "Save" "Ctrl+S" None (fun () -> match currentPath.Value with Some p -> writeTo p | None -> XlUi.saveFileDialog "" writeTo |> Shell.openDialog shell)
                  menuItem "xl-menu-save-as" "Save as..." None (fun () -> XlUi.saveFileDialog (currentPath.Value |> Option.defaultValue "") writeTo |> Shell.openDialog shell) ] }
        fileMenu :: (XlUi.menus shellState actions |> List.tail)
    let toolbarRows = XlUi.toolbarRows shellState actions

    let formulaBar () =
        heightBox 30.0
            (Ui.surfaceBox (fun x y w h -> [ FillRect(rect x y w h, Color.rgb 248uy 249uy 250uy); Cmd.Line(vec2 x (y + h - 1.0), vec2 (x + w) (y + h - 1.0), Color.rgb 218uy 220uy 224uy, 1.0<px>) ])
                (padXY 8.0 3.0
                    (rowA MStart CCenter 8.0
                        [ sizedBox 72.0 22.0 (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Well x y w h Ui.radius) (padXY 6.0 3.0 (textf 12.0 Ui.P.text false (fun () -> EH.activeCellRef st))))
                          expanded 1 (Ui.surfaceBox (fun x y w h -> Ui.face Ui.theme Ui.Well x y w h Ui.radius) (padXY 8.0 3.0 (textf 12.0 Ui.P.text false (fun () -> EH.formulaText st)))) ])))

    let sheetTabs () =
        heightBox 28.0
            (Ui.surfaceBox (fun x y w h -> [ FillRect(rect x y w h, Color.rgb 248uy 249uy 250uy); Cmd.Line(vec2 x y, vec2 (x + w) y, Color.rgb 218uy 220uy 224uy, 1.0<px>) ])
                (padXY 8.0 3.0
                    (rowA MStart CCenter 6.0
                        [ for i in 0 .. st.SheetNames.Value.Length - 1 do
                            let idx = i
                            yield
                                clickable (fun () -> EH.switchSheet st idx)
                                    (Ui.surfaceBox (fun x y w h -> if st.ActiveSheet.Value = idx then [ FillRoundedRect(rect x y w h, 4.0<px>, Color.rgb 232uy 245uy 233uy) ] else [])
                                        (padXY 10.0 3.0 (text 12.0 Ui.P.text st.SheetNames.Value.[i])))
                          yield clickable actions.InsertSheet (padXY 8.0 3.0 (text 13.0 Ui.P.text "+")) ])))

    let workspace () =
        columnA MStart CStretch 0.0
            [ formulaBar ()
              expanded 1
                (canvasFill (fun x y w h ->
                    st.VisibleWidth.Value <- w
                    st.VisibleHeight.Value <- h
                    [ Cmd.Clip([ MoveTo(vec2 x y); LineTo(vec2 (x + w) y); LineTo(vec2 (x + w) (y + h)); LineTo(vec2 x (y + h)); Close ], NonZero,
                        [ Cmd.Group(Matrix3x2.translate x y, 1.0, EH.renderGrid st w h) ]) ]))
                |> pointerRegion (fun kind x y -> EH.handlePointer st st.VisibleWidth.Value st.VisibleHeight.Value kind x y)
                |> fun area ->
                    area.OnScroll <- Some(fun dy -> EH.scrollBy st 0.0 (-dy))
                    area
              sheetTabs ()
              heightBox 22.0 (padXY 8.0 3.0 (textf 11.5 Ui.P.dim false (fun () -> status.Value + "   " + shellState.SelectionSummary()))) ]

    let view () = Shell.chrome shell (title ()) menus toolbarRows (workspace ())

    match Array.tryFindIndex ((=) "--screenshot") argv with
    | Some i when i + 1 < argv.Length ->
        Wisp.Host.shot appW appH view argv.[i + 1]
        printfn "wrote %s" argv.[i + 1]
        0
    | _ ->
        Wisp.Host.keyHandler.Value <- fun kc shift ctrl alt ->
            if ctrl && (kc = 31 || kc = 39) then
                match KeyMap.translate kc shift with
                | Some k when k = "s" || k = "S" -> match currentPath.Value with Some p -> writeTo p | None -> XlUi.saveFileDialog "" writeTo |> Shell.openDialog shell
                | Some k when k = "o" || k = "O" -> XlUi.openFileDialog loadFromPath |> Shell.openDialog shell
                | _ -> ()
            else
                match KeyMap.translate kc shift with
                | Some key -> EH.handleKey st { Key = key; Shift = shift; Ctrl = ctrl; Alt = alt } |> ignore; Scheduler.requestFrame ()
                | None -> ()
        Wisp.Host.run "wisp-spreadsheet" appW appH view
        0
