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
module L = Rbt.Spreadsheet.Layout
module EH = Rbt.SpreadsheetWisp.EditorHost
module GR = Rbt.SpreadsheetWisp.GridRender
module Pivot = Rbt.Spreadsheet.Pivot
module Formula = Rbt.Spreadsheet.Formula

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

module XlsxRead =
    let private colIndex (letters: string) =
        let mutable acc = 0
        for ch in letters do
            let c = Char.ToUpperInvariant ch
            if c >= 'A' && c <= 'Z' then acc <- acc * 26 + (int c - int 'A' + 1)
        acc - 1

    let private parseCellRef (s: string) =
        let cleaned = s.Replace("$", "").Trim()
        let mutable letters = ""
        let mutable digits = ""
        for ch in cleaned do
            if Char.IsLetter ch && digits = "" then letters <- letters + string ch
            elif Char.IsDigit ch then digits <- digits + string ch
        match Int32.TryParse digits with
        | true, row when letters <> "" && row > 0 -> Some(row - 1, colIndex letters)
        | _ -> None

    let private parseRange (s: string) =
        let unqualified =
            match s.LastIndexOf('!') with
            | i when i >= 0 -> s.Substring(i + 1)
            | _ -> s
        match unqualified.Split(':') with
        | [| a; b |] ->
            match parseCellRef a, parseCellRef b with
            | Some (r1, c1), Some (r2, c2) -> Some(min r1 r2, min c1 c2, max r1 r2, max c1 c2)
            | _ -> None
        | [| a |] -> parseCellRef a |> Option.map (fun (r, c) -> r, c, r, c)
        | _ -> None

    let private hexColor (value: string) =
        let s = value.Trim().TrimStart('#')
        let s = if s.Length = 8 then s.Substring 2 else s
        if s.Length = 6 then Some("#" + s) else None

    let private border (b: XlsxReader.CellBorderSide option) =
        b |> Option.map (fun side -> ({ Style = side.Style; Color = side.Color |> Option.bind hexColor } : M.BorderStyle))

    let private format (s: XlsxReader.CellStyle) : M.CellFormat =
        { M.defaultFormat with
            Bold = s.Bold
            Italic = s.Italic
            Underline = s.Underline
            UnderlineDouble = s.UnderlineDouble
            Strikethrough = s.Strikethrough
            FontName = s.FontName
            TextColor = s.FontColor |> Option.bind hexColor
            BgColor = s.FillColor |> Option.bind hexColor
            FontSize = s.FontSize |> Option.map int
            NumberFormat = s.NumberFormat
            HorizontalAlign = s.HorizontalAlign
            VerticalAlign = s.VerticalAlign
            WrapText = s.WrapText
            Indent = s.Indent
            ShrinkToFit = s.ShrinkToFit
            TextRotation = s.TextRotation
            BorderTop = border s.BorderTop
            BorderBottom = border s.BorderBottom
            BorderLeft = border s.BorderLeft
            BorderRight = border s.BorderRight
            BorderDiagonal = border s.BorderDiagonal
            DiagonalUp = s.DiagonalUp
            DiagonalDown = s.DiagonalDown }

    let private runOverride (props: XlsxReader.RunProps) : M.RunOverride =
        { M.emptyRunOverride with
            Bold = Some props.Bold
            Italic = Some props.Italic
            Underline = Some props.Underline
            UnderlineDouble = Some props.UnderlineDouble
            Strikethrough = Some props.Strikethrough
            Color = props.Color |> Option.bind hexColor
            FontName = props.FontName
            FontSize = props.FontSize
            VertAlign =
                match props.VertAlign with
                | XlsxReader.VertAlignSuperscript -> Some M.VertAlignSuperscript
                | XlsxReader.VertAlignSubscript -> Some M.VertAlignSubscript
                | XlsxReader.VertAlignBaseline -> Some M.VertAlignBaseline }

    let private valueOf (wb: XlsxReader.Workbook) (v: XlsxReader.CellValue) =
        match XlsxReader.resolveValue wb v with
        | XlsxReader.Number n -> M.Number n
        | XlsxReader.String s -> M.String s
        | XlsxReader.Boolean b -> M.Boolean b
        | XlsxReader.Date s -> M.String s
        | XlsxReader.Formula f -> M.Formula(f, M.Empty)
        | XlsxReader.Error e -> M.Error(e.TrimStart '#')
        | XlsxReader.Empty -> M.Empty
        | XlsxReader.SharedString _ -> M.Empty

    let private parseDouble (s: string) =
        match Double.TryParse(s, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture) with
        | true, n -> Some n
        | _ -> None

    let private validationRule (dv: XlsxReader.DataValidation) : M.DataValidation option =
        match dv.ValidationType, dv.Formula1, dv.Formula2 with
        | "list", Some f, _ ->
            let values = f.Trim('"').Split(',') |> Array.toList
            Some ({ Rule = M.ListValues values; ErrorMessage = dv.ErrorMessage; InputMessage = None } : M.DataValidation)
        | "whole", lo, hi ->
            Some ({ Rule = M.WholeNumber(lo |> Option.bind parseDouble, hi |> Option.bind parseDouble); ErrorMessage = dv.ErrorMessage; InputMessage = None } : M.DataValidation)
        | "decimal", lo, hi ->
            Some ({ Rule = M.DecimalRange(lo |> Option.bind parseDouble, hi |> Option.bind parseDouble); ErrorMessage = dv.ErrorMessage; InputMessage = None } : M.DataValidation)
        | _ -> None

    let private pageSetup (sheet: XlsxReader.Worksheet) =
        let margins =
            sheet.PageMargins
            |> Option.map (fun m -> ({ Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer } : M.PageMargins))
            |> Option.defaultValue M.defaultPageMargins
        let baseSetup = M.defaultPageSetup
        let setup =
            match sheet.PageSetup with
            | Some p ->
                let w, h = p.PaperSize |> Option.map M.paperSizePoints |> Option.defaultValue (baseSetup.PageWidth, baseSetup.PageHeight)
                { baseSetup with
                    PageWidth = w
                    PageHeight = h
                    Orientation = (match p.Orientation with XlsxReader.OrientationLandscape -> M.LandscapeOrientation | _ -> M.PortraitOrientation)
                    FitToPage = p.FitToPage
                    FitToWidth = p.FitToWidth
                    FitToHeight = p.FitToHeight
                    Scale = p.Scale
                    FirstPageNumber = p.FirstPageNumber
                    UseFirstPageNumber = p.UseFirstPageNumber
                    PageOrder = (match p.PageOrder with XlsxReader.OverThenDown -> M.OverThenDown | _ -> M.DownThenOver)
                    BlackAndWhite = p.BlackAndWhite
                    CellComments =
                        match p.CellComments with
                        | XlsxReader.CellCommentsAsDisplayed -> M.CellCommentsAsDisplayed
                        | XlsxReader.CellCommentsAtEnd -> M.CellCommentsAtEnd
                        | _ -> M.CellCommentsNone }
            | None -> baseSetup
        { setup with
            Margins = margins
            Options =
                { GridLines = sheet.PrintOptions.GridLines
                  Headings = sheet.PrintOptions.Headings
                  HorizontalCentered = sheet.PrintOptions.HorizontalCentered
                  VerticalCentered = sheet.PrintOptions.VerticalCentered }
            HeaderFooter =
                { OddHeader = sheet.HeaderFooter.OddHeader
                  OddFooter = sheet.HeaderFooter.OddFooter
                  EvenHeader = sheet.HeaderFooter.EvenHeader
                  EvenFooter = sheet.HeaderFooter.EvenFooter
                  FirstHeader = sheet.HeaderFooter.FirstHeader
                  FirstFooter = sheet.HeaderFooter.FirstFooter
                  DifferentOddEven = sheet.HeaderFooter.DifferentOddEven
                  DifferentFirst = sheet.HeaderFooter.DifferentFirst
                  ScaleWithDoc = sheet.HeaderFooter.ScaleWithDoc }
            ManualRowBreaks = sheet.RowBreaks
            ManualColBreaks = sheet.ColBreaks
            PrintArea =
                sheet.PrintArea
                |> Option.bind parseRange
                |> Option.map (fun (sr, sc, er, ec) -> ({ From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } : M.CellRange))
            RepeatRows =
                sheet.PrintTitleRows
                |> Option.bind (fun s ->
                    match s.Replace("$", "").Split(':') with
                    | [| a; b |] ->
                        match Int32.TryParse a, Int32.TryParse b with
                        | (true, r1), (true, r2) -> Some(r1 - 1, r2 - 1)
                        | _ -> None
                    | _ -> None)
            RepeatCols =
                sheet.PrintTitleCols
                |> Option.bind (fun s ->
                    match s.Replace("$", "").Split(':') with
                    | [| a; b |] -> Some(colIndex a, colIndex b)
                    | _ -> None) }

    let private tableStyle (style: XlsxReader.TableStyleInfo option) =
        style
        |> Option.map (fun s ->
            let result : M.TableStyle =
                { Name = s.Name
                  ShowFirstColumn = s.ShowFirstColumn
                  ShowLastColumn = s.ShowLastColumn
                  ShowRowStripes = s.ShowRowStripes
                  ShowColumnStripes = s.ShowColumnStripes
                  HeaderFill = Some "#5B9BD5"
                  StripeFill = Some "#DDEBF7" }
            result)

    let toModel (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) =
        let mutable maxRow = 0
        let mutable maxCol = 0
        let comments = System.Collections.Generic.Dictionary<int * int, M.CellComment>()
        for c in sheet.Comments do
            let r = c.CellRef.Row - 1
            let col = colIndex c.CellRef.Column
            if r >= 0 && col >= 0 then comments.[(r, col)] <- { Text = c.Text; Author = c.Author }
        let listDropdowns = System.Collections.Generic.HashSet<int * int>()
        let mutable validations : Map<M.CellAddress, M.DataValidation> = Map.empty
        for dv in sheet.DataValidations do
            match parseRange dv.Range, validationRule dv with
            | Some (sr, sc, er, ec), Some rule ->
                for r in sr .. er do
                    for c in sc .. ec do
                        validations <- Map.add { Row = r; Col = c } rule validations
                        if dv.ValidationType = "list" then listDropdowns.Add(r, c) |> ignore
            | _ -> ()
        let hyperlinks = System.Collections.Generic.HashSet<int * int>()
        for h in sheet.Hyperlinks do
            match parseRange h.Range with
            | Some (sr, sc, er, ec) ->
                for r in sr .. er do for c in sc .. ec do hyperlinks.Add(r, c) |> ignore
            | None -> ()
        let mutable richRuns : Map<M.CellAddress, (string * M.RunOverride) list> = Map.empty
        let cells =
            sheet.Rows
            |> List.collect (fun row ->
                row.Cells
                |> List.choose (fun cell ->
                    let r = cell.Reference.Row - 1
                    let c = colIndex cell.Reference.Column
                    if r < 0 || c < 0 then None
                    else
                        let fmt = cell.Style |> Option.map format |> Option.defaultValue M.defaultFormat
                        let fmt =
                            { fmt with
                                TextColor = if hyperlinks.Contains(r, c) then Some "#0563C1" else fmt.TextColor
                                Underline = fmt.Underline || hyperlinks.Contains(r, c)
                                HasComment = comments.ContainsKey(r, c)
                                HasListDropdown = listDropdowns.Contains(r, c) }
                        let comment = if comments.ContainsKey(r, c) then Some comments.[(r, c)] else None
                        match cell.Value with
                        | XlsxReader.SharedString idx when idx >= 0 && idx < wb.SharedStringRuns.Length && not (List.isEmpty wb.SharedStringRuns.[idx]) ->
                            richRuns <- Map.add { Row = r; Col = c } (wb.SharedStringRuns.[idx] |> List.map (fun (t, p) -> t, runOverride p)) richRuns
                        | _ when not (List.isEmpty cell.InlineRuns) ->
                            richRuns <- Map.add { Row = r; Col = c } (cell.InlineRuns |> List.map (fun (t, p) -> t, runOverride p)) richRuns
                        | _ -> ()
                        maxRow <- max maxRow r
                        maxCol <- max maxCol c
                        // Keep the formula with its cached result, not just the result.
                        let cellValue =
                            match cell.Formula, valueOf wb cell.Value with
                            | Some _, (M.Formula _ as uncached) -> uncached
                            | Some formula, cached -> M.Formula(formula, cached)
                            | None, value -> value
                        Some(({ Row = r; Col = c } : M.CellAddress), ({ Value = cellValue; Format = fmt; Comment = comment } : M.CellData))))
            |> Map.ofList
        let rowHeights =
            sheet.Rows
            |> List.choose (fun r -> if r.Hidden then Some(r.Index - 1, 0.0) else r.Height |> Option.map (fun h -> r.Index - 1, h * 1.333))
            |> Map.ofList
        let colWidths =
            sheet.Columns
            |> List.collect (fun c ->
                [ for idx in c.Index .. c.MaxIndex ->
                    idx - 1, if c.Hidden then 0.0 else c.Width |> Option.map ((*) 7.0) |> Option.defaultValue 100.0 ])
            |> Map.ofList
        let merged =
            sheet.MergedCells
            |> List.choose (fun m ->
                let sr = m.StartRef.Row - 1
                let sc = colIndex m.StartRef.Column
                let er = m.EndRef.Row - 1
                let ec = colIndex m.EndRef.Column
                if sr >= 0 && sc >= 0 && er >= sr && ec >= sc then Some ({ StartRow = sr; StartCol = sc; EndRow = er; EndCol = ec } : M.MergedCell) else None)
        let tables =
            sheet.Tables
            |> List.map (fun t ->
                let result : M.Table =
                    { Name = t.Name
                      StartRow = t.StartRow
                      StartCol = t.StartCol
                      EndRow = t.EndRow
                      EndCol = t.EndCol
                      HeaderRowCount = t.HeaderRowCount
                      Style = tableStyle t.Style
                      AutoFilter = t.AutoFilter
                      HasTotals = t.TotalsRowCount > 0 }
                result)
        let frozenRows, frozenCols =
            match sheet.FrozenPane with
            | Some p -> p.YSplit, p.XSplit
            | None -> 0, 0
        { M.createEmpty (max 200 (maxRow + 200)) (max 40 (maxCol + 50)) with
            Cells = cells
            RichRuns = richRuns
            MergedCells = merged
            RowHeights = rowHeights
            ColumnWidths = colWidths
            Tables = tables
            FrozenRows = frozenRows
            FrozenCols = frozenCols
            ShowGridLines = sheet.ShowGridLines
            Validations = validations
            PageSetup = pageSetup sheet }

    let load path : (M.SpreadsheetModel array * string array) option =
        try
            let bytes = File.ReadAllBytes path
            match XlsxReader.loadBytesNative bytes with
            | None -> None
            | Some wb ->
                let macros = VbaProject.extractFromXlsx bytes |> Option.defaultValue ""
                let sheets =
                    wb.Sheets
                    |> List.map (fun s ->
                        let m = toModel wb s
                        if String.IsNullOrWhiteSpace macros then m else { m with Macros = macros })
                    |> List.toArray
                let names = wb.Sheets |> List.map (fun s -> s.Name) |> List.toArray
                if sheets.Length = 0 then None else Some(sheets, names)
        with ex ->
            eprintfn "[spreadsheet-wisp] load failed: %s" ex.Message
            None

module XlsxWrite =
    let private border (b: M.BorderStyle option) =
        b |> Option.map (fun side -> ({ Style = side.Style; Color = side.Color } : XlsxWriter.XlsxBorderSide))

    let private format (f: M.CellFormat) : XlsxWriter.XlsxCellFormat =
        { XlsxWriter.defaultXlsxCellFormat with
            Bold = f.Bold
            Italic = f.Italic
            Underline = f.Underline
            UnderlineDouble = f.UnderlineDouble
            Strikethrough = f.Strikethrough
            FontName = f.FontName
            TextColor = f.TextColor
            FillColor = f.BgColor
            FontSize = f.FontSize |> Option.map float
            NumberFormat = f.NumberFormat
            HorizontalAlign = f.HorizontalAlign
            VerticalAlign = f.VerticalAlign
            WrapText = f.WrapText
            Indent = f.Indent
            ShrinkToFit = f.ShrinkToFit
            TextRotation = f.TextRotation
            BorderTop = border f.BorderTop
            BorderBottom = border f.BorderBottom
            BorderLeft = border f.BorderLeft
            BorderRight = border f.BorderRight
            BorderDiagonal = border f.BorderDiagonal
            DiagonalUp = f.DiagonalUp
            DiagonalDown = f.DiagonalDown }

    let private richRun (text: string, r: M.RunOverride) : XlsxWriter.XlsxRichRun =
        { Text = text
          Bold = r.Bold
          Italic = r.Italic
          Underline = r.Underline
          UnderlineDouble = r.UnderlineDouble
          Strikethrough = r.Strikethrough
          Color = r.Color
          FontName = r.FontName
          FontSize = r.FontSize
          VertAlign =
            match r.VertAlign with
            | Some M.VertAlignSuperscript -> Some XlsxWriter.XlsxSuperscript
            | Some M.VertAlignSubscript -> Some XlsxWriter.XlsxSubscript
            | Some M.VertAlignBaseline -> Some XlsxWriter.XlsxBaseline
            | None -> None }

    let rec private value (richRuns: XlsxWriter.XlsxRichRun list) (v: M.CellValue) : XlsxWriter.XlsxCellValue =
        match v with
        | M.Empty -> XlsxWriter.XlsxBlank
        | M.Number n -> XlsxWriter.XlsxNumber n
        | M.String s -> XlsxWriter.XlsxText(s, richRuns)
        | M.Boolean b -> XlsxWriter.XlsxBoolean b
        | M.Error e -> XlsxWriter.XlsxError(if e.StartsWith "#" then e else "#" + e)
        | M.Formula (expr, cached) -> XlsxWriter.XlsxFormula(expr, Some(value [] cached))

    let private validationRef row col = sprintf "%s%d" (M.columnToLetter col) (row + 1)

    let private validation (addr: M.CellAddress) (dv: M.DataValidation) =
        let typ, op, f1, f2 =
            match dv.Rule with
            | M.ListValues values -> "list", None, Some("\"" + String.concat "," values + "\""), None
            | M.WholeNumber (lo, hi) -> "whole", Some "between", lo |> Option.map string, hi |> Option.map string
            | M.DecimalRange (lo, hi) -> "decimal", Some "between", lo |> Option.map string, hi |> Option.map string
            | M.TextLength (lo, hi) -> "textLength", Some "between", lo |> Option.map string, hi |> Option.map string
            | M.DateRange (lo, hi) -> "date", Some "between", lo |> Option.map string, hi |> Option.map string
        ({ Sqref = validationRef addr.Row addr.Col
           ValidationType = typ
           Operator = op
           Formula1 = f1
           Formula2 = f2
           AllowBlank = true
           ShowDropDown = true
           ErrorMessage = dv.ErrorMessage
           InputMessage = dv.InputMessage
           ErrorTitle = None
           PromptTitle = None
           ErrorStyle = None
           ShowErrorMessage = true
           ShowInputMessage = dv.InputMessage.IsSome } : XlsxWriter.XlsxDataValidation)

    let private textAt row col (model: M.SpreadsheetModel) =
        let addr : M.CellAddress = { Row = row; Col = col }
        match Map.tryFind addr model.Cells with
        | Some data -> M.displayValue data.Value
        | None -> sprintf "Column%d" (col + 1)

    let private table (model: M.SpreadsheetModel) (t: M.Table) : XlsxWriter.XlsxTableOut =
        { Name = t.Name
          StartRow = t.StartRow
          StartCol = t.StartCol
          EndRow = t.EndRow
          EndCol = t.EndCol
          Columns = [ for c in t.StartCol - 1 .. t.EndCol - 1 -> textAt (t.StartRow - 1) c model ]
          HeaderRowCount = t.HeaderRowCount
          TotalsRowCount = if t.HasTotals then 1 else 0
          AutoFilter = t.AutoFilter
          Style =
            t.Style |> Option.map (fun s ->
                ({ Name = s.Name
                   ShowFirstColumn = s.ShowFirstColumn
                   ShowLastColumn = s.ShowLastColumn
                   ShowRowStripes = s.ShowRowStripes
                   ShowColumnStripes = s.ShowColumnStripes } : XlsxWriter.XlsxTableStyle)) }

    let private margins (m: M.PageMargins) : XlsxWriter.XlsxPageMargins =
        { Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer }

    let private pageSetup (p: M.PageSetup) : XlsxWriter.XlsxPageSetup =
        { Orientation = Some(match p.Orientation with M.LandscapeOrientation -> "landscape" | M.PortraitOrientation -> "portrait")
          FitToPage = p.FitToPage
          FitToWidth = p.FitToWidth
          FitToHeight = p.FitToHeight
          Scale = p.Scale
          PaperSize = None
          FirstPageNumber = p.FirstPageNumber
          UseFirstPageNumber = p.UseFirstPageNumber
          PageOrder = Some(match p.PageOrder with M.OverThenDown -> "overThenDown" | M.DownThenOver -> "downThenOver")
          BlackAndWhite = p.BlackAndWhite
          CellComments =
            match p.CellComments with
            | M.CellCommentsAsDisplayed -> Some "asDisplayed"
            | M.CellCommentsAtEnd -> Some "atEnd"
            | M.CellCommentsNone -> None }

    let private printOptions (p: M.PrintOptions) : XlsxWriter.XlsxPrintOptions =
        { GridLines = p.GridLines; Headings = p.Headings; HorizontalCentered = p.HorizontalCentered; VerticalCentered = p.VerticalCentered }

    let private headerFooter (h: M.HeaderFooter) : XlsxWriter.XlsxHeaderFooter =
        { OddHeader = h.OddHeader; OddFooter = h.OddFooter; EvenHeader = h.EvenHeader; EvenFooter = h.EvenFooter
          FirstHeader = h.FirstHeader; FirstFooter = h.FirstFooter; DifferentOddEven = h.DifferentOddEven
          DifferentFirst = h.DifferentFirst; ScaleWithDoc = h.ScaleWithDoc }

    let private rangeRef (r: M.CellRange) =
        sprintf "$%s$%d:$%s$%d" (M.columnToLetter r.From.Col) (r.From.Row + 1) (M.columnToLetter r.To.Col) (r.To.Row + 1)

    let private modelSheet name (model: M.SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        { Name = name
          Visible = true
          Cells =
            model.Cells
            |> Map.toList
            |> List.map (fun (addr, data) ->
                let runs = model.RichRuns |> Map.tryFind addr |> Option.map (List.map richRun) |> Option.defaultValue []
                { XlsxWriter.Row = addr.Row
                  Col = addr.Col
                  Value = value runs data.Value
                  Format = format data.Format
                  Comment = data.Comment |> Option.map (fun c -> c.Text, c.Author) })
          MergedCells = model.MergedCells |> List.map (fun m -> ({ StartRow = m.StartRow; StartCol = m.StartCol; EndRow = m.EndRow; EndCol = m.EndCol } : XlsxWriter.MergedCellRange))
          ColumnWidths = model.ColumnWidths |> Map.toList |> List.map (fun (c, px) -> c, px / 7.0)
          RowHeights = model.RowHeights |> Map.toList |> List.map (fun (r, px) -> r, px / 1.333)
          FrozenRows = model.FrozenRows
          FrozenCols = model.FrozenCols
          ShowGridLines = model.ShowGridLines
          Tables = model.Tables |> List.map (table model)
          DataValidations = model.Validations |> Map.toList |> List.map (fun (addr, dv) -> validation addr dv)
          ConditionalFormats = []
          Hyperlinks = []
          PageMargins = Some(margins model.PageSetup.Margins)
          PageSetup = Some(pageSetup model.PageSetup)
          PrintOptions = Some(printOptions model.PageSetup.Options)
          HeaderFooter = Some(headerFooter model.PageSetup.HeaderFooter)
          RowBreaks = model.PageSetup.ManualRowBreaks
          ColBreaks = model.PageSetup.ManualColBreaks
          PrintArea = model.PageSetup.PrintArea |> Option.map rangeRef
          PrintTitleRows = model.PageSetup.RepeatRows |> Option.map (fun (a, b) -> sprintf "$%d:$%d" (a + 1) (b + 1))
          PrintTitleCols = model.PageSetup.RepeatCols |> Option.map (fun (a, b) -> sprintf "$%s:$%s" (M.columnToLetter a) (M.columnToLetter b)) }

    let save (sheets: M.SpreadsheetModel array) (names: string array) path =
        let sheetDtos =
            sheets
            |> Array.mapi (fun i model -> modelSheet (if i < names.Length then names.[i] else sprintf "Sheet%d" (i + 1)) model)
            |> Array.toList
        let macroSource = sheets |> Array.tryPick (fun m -> if String.IsNullOrWhiteSpace m.Macros then None else Some m.Macros)
        let workbook = ({ Sheets = sheetDtos; ActiveSheetIndex = 0; DefinedNames = []; VbaProjectSource = macroSource; Dxfs = [] } : XlsxWriter.XlsxWorkbookOut)
        File.WriteAllBytes(path, XlsxWriter.writeWorkbookPackageBytes workbook)

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

let private updateFormat st update =
    let model = EH.activeModel st
    let r, c = EH.focusCell model
    let addr : M.CellAddress = { Row = r; Col = c }
    EH.applyChange st (fun m ->
        let existing = M.getCell addr m |> Option.defaultValue { Value = M.Empty; Format = M.defaultFormat; Comment = None }
        M.setCell addr { existing with Format = update existing.Format } m) |> ignore

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
    let incremental = st.LastRecalcWasIncremental || st.CalcEngine.Value.IsSome
    EH.handleKey st { Key = "ArrowRight"; Shift = false; Ctrl = false; Alt = false } |> ignore
    let navOk = EH.focusCell (EH.activeModel st) = (0, 1)
    EH.handleKey st { Key = "Delete"; Shift = false; Ctrl = false; Alt = false } |> ignore
    let deleted = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) = M.Empty
    EH.undo st |> ignore
    let restored = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) <> M.Empty
    EH.redo st |> ignore
    let redone = M.getCellValue { Row = 0; Col = 1 } (EH.activeModel st) = M.Empty
    if formulaOk && navOk && deleted && restored && redone && incremental then
        printfn "spreadsheet-wisp selftest ok"
        0
    else
        eprintfn "selftest failed formula=%b nav=%b deleted=%b restored=%b redone=%b incremental=%b" formulaOk navOk deleted restored redone incremental
        1

[<EntryPoint>]
let main argv =
    if Array.contains "--selftest" argv then selfTest () else
    let xlsxPath = argv |> Array.tryFind (fun a -> a.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
    let sheets, names =
        match xlsxPath |> Option.bind XlsxRead.load with
        | Some loaded ->
            xlsxPath |> Option.iter rememberRecent
            loaded
        | None -> sampleWorkbook ()
    let st = EH.create sheets names 0
    let shell = Shell.create ()
    let status = refSignal "Ready"
    let currentPath = refSignal xlsxPath
    let zoom = signal "100"

    let title () =
        let name = currentPath.Value |> Option.map Path.GetFileName |> Option.defaultValue "Book1"
        sprintf "%s - wisp-spreadsheet" name

    let writeTo path =
        try
            XlsxWrite.save st.Sheets.Value st.SheetNames.Value path
            currentPath.Value <- Some path
            rememberRecent path
            status.Value <- sprintf "Saved %s" (Path.GetFileName path)
        with ex -> status.Value <- "Save failed: " + ex.Message

    let loadFromPath path =
        match XlsxRead.load path with
        | Some (sheets, names) ->
            st.Sheets.Value <- sheets
            st.SheetNames.Value <- names
            st.ActiveSheet.Value <- 0
            st.UndoStack.Value <- []
            st.RedoStack.Value <- []
            st.CalcEngine.Value <- None
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
            let sheetName = names |> Array.tryItem st.ActiveSheet.Value |> Option.defaultValue "Sheet1"
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
                    EH.applyChange st (fun m ->
                        let m1 = M.setCell { Row = dr; Col = dc } { Value = M.Formula(formula, M.Empty); Format = M.defaultFormat; Comment = None } m
                        Formula.recalcAndSpill None None m1) |> ignore
                    status.Value <- sprintf "Inserted pivot formula at %s" (M.formatCellRef dr dc)
                    true
                | Result.Error msg -> status.Value <- "Pivot: " + msg; false
            | _ -> status.Value <- "Pivot: invalid range or destination"; false)
        |> Shell.openDialog shell

    let openInsertSheetDialog () =
        XlUi.insertSheetDialog (fun name ->
            st.Sheets.Value <- Array.append st.Sheets.Value [| M.createEmpty 100 26 |]
            st.SheetNames.Value <- Array.append st.SheetNames.Value [| name |]
            EH.switchSheet st (st.Sheets.Value.Length - 1)
            status.Value <- sprintf "Added sheet \"%s\"" name)
        |> Shell.openDialog shell

    let openRenameSheetDialog () =
        let idx = st.ActiveSheet.Value
        let initial = st.SheetNames.Value |> Array.tryItem idx |> Option.defaultValue ""
        XlUi.renameSheetDialog initial (fun name ->
            let next = Array.copy st.SheetNames.Value
            if idx >= 0 && idx < next.Length then
                next.[idx] <- name
                st.SheetNames.Value <- next
                st.CalcEngine.Value <- None
                status.Value <- sprintf "Renamed sheet to \"%s\"" name)
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
          ToggleBold = (fun () -> updateFormat st (fun f -> { f with Bold = not f.Bold }))
          ToggleItalic = (fun () -> updateFormat st (fun f -> { f with Italic = not f.Italic }))
          ToggleUnderline = (fun () -> updateFormat st (fun f -> { f with Underline = not f.Underline }))
          ToggleStrike = (fun () -> updateFormat st (fun f -> { f with Strikethrough = not f.Strikethrough }))
          ToggleWrapText = (fun () -> updateFormat st (fun f -> { f with WrapText = not f.WrapText }))
          ClearFormatting = (fun () -> updateFormat st (fun _ -> M.defaultFormat))
          ApplyHorizontalAlign = (fun v -> updateFormat st (fun f -> { f with HorizontalAlign = Some(if v = "right" || v = "center" then v else "left") }))
          ApplyVerticalAlign = (fun v -> updateFormat st (fun f -> { f with VerticalAlign = Some(if v = "top" || v = "middle" then v else "bottom") }))
          ApplyNumberFormat = (fun v -> updateFormat st (fun f -> { f with NumberFormat = Some v }))
          IncreaseDecimals = (fun () -> status.Value <- "Increase decimals not wired")
          DecreaseDecimals = (fun () -> status.Value <- "Decrease decimals not wired")
          ApplyFontFamily = (fun v -> updateFormat st (fun f -> { f with FontName = Some v }))
          ApplyFontSize = (fun v -> match Int32.TryParse v with true, n -> updateFormat st (fun f -> { f with FontSize = Some n }) | _ -> ())
          ApplyTextColor = (fun v -> updateFormat st (fun f -> { f with TextColor = Some v }))
          ApplyBgColor = (fun v -> updateFormat st (fun f -> { f with BgColor = Some v }))
          ApplyBorderStyle = (fun side -> if side <> "none" then updateFormat st (fun f -> { f with BorderTop = Some { Style = "thin"; Color = Some "#000000" }; BorderBottom = Some { Style = "thin"; Color = Some "#000000" }; BorderLeft = Some { Style = "thin"; Color = Some "#000000" }; BorderRight = Some { Style = "thin"; Color = Some "#000000" } }))
          ApplyZoom = (fun v -> zoom.Value <- v)
          MergeCells = (fun () -> match (EH.activeModel st).Selection with M.Range(sr, sc, er, ec) -> EH.applyChange st (M.addMergedCell sr sc er ec) |> ignore | _ -> status.Value <- "Select a range to merge")
          UnmergeCells = (fun () -> let r, c = EH.focusCell (EH.activeModel st) in EH.applyChange st (M.removeMergedCell { Row = r; Col = c }) |> ignore)
          InsertRow = (fun () -> let r, _ = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.InsertRows (r, 1)) |> ignore)
          InsertColumn = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.InsertCols (c, 1)) |> ignore)
          DeleteRow = (fun () -> let r, _ = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.DeleteRows (r, 1)) |> ignore)
          DeleteColumn = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyStructureEdit st (M.DeleteCols (c, 1)) |> ignore)
          SortAscending = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyChange st (M.sortByColumn c true) |> ignore)
          SortDescending = (fun () -> let _, c = EH.focusCell (EH.activeModel st) in EH.applyChange st (M.sortByColumn c false) |> ignore)
          ToggleFilter = openFilterDialog
          ClearFilter = (fun () -> EH.applyChange st (fun m -> { m with Filters = Map.empty; FilterHiddenRows = [||] }) |> ignore; status.Value <- "Filters cleared")
          ToggleFreezePanes = (fun () -> EH.applyChange st (fun m -> if m.FrozenRows > 0 || m.FrozenCols > 0 then { m with FrozenRows = 0; FrozenCols = 0 } else let r, c = EH.focusCell m in { m with FrozenRows = r; FrozenCols = c }) |> ignore)
          SplitPanes = (fun () -> status.Value <- "Split panes not wired")
          InsertChart = (fun () -> status.Value <- "Insert chart not wired")
          InsertPivot = openPivotDialog
          InsertTable = openInsertTableDialog
          RunMacro = openMacroDialog
          InsertSheet = openInsertSheetDialog
          DeleteSheet = (fun () ->
              if st.Sheets.Value.Length <= 1 then status.Value <- "Can't delete the last sheet"
              else
                  let idx = st.ActiveSheet.Value
                  st.Sheets.Value <- st.Sheets.Value |> Array.mapi (fun i s -> i, s) |> Array.filter (fun (i, _) -> i <> idx) |> Array.map snd
                  st.SheetNames.Value <- st.SheetNames.Value |> Array.mapi (fun i s -> i, s) |> Array.filter (fun (i, _) -> i <> idx) |> Array.map snd
                  st.ActiveSheet.Value <- max 0 (min (st.Sheets.Value.Length - 1) idx)
                  st.CalcEngine.Value <- None
                  status.Value <- "Deleted sheet")
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
          CanUndo = (fun () -> not (List.isEmpty st.UndoStack.Value))
          CanRedo = (fun () -> not (List.isEmpty st.RedoStack.Value)) }

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
