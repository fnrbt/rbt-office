/// XLSX <-> editor model, for every host: the browser component (Fable) and
/// the native editors (.NET) share this one conversion in both directions.
///
///   Import.importWorkbook   reader workbook (XlsxReader.Workbook) -> sheet models + workbook context
///   Export.toWorkbookOut    sheet models (+ the workbook they were imported from) -> XlsxWriter.XlsxWorkbookOut
///
/// The export writes the editor's current state. What the editor does not
/// model is carried over from the workbook the sheets were imported from:
/// pictures, charts and other drawing objects, threaded comments, named cell
/// styles, the theme and the other kept parts; and a cell, table, validation
/// or conditional format the user did not touch is written as the file had it.
module Rbt.Spreadsheet.XlsxBridge

open System
open Rbt.Spreadsheet.Model

module CF = Rbt.Spreadsheet.ConditionalFormatting
module Wb = Rbt.Spreadsheet.Workbook

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

/// Zero-based column index of column letters ("A" -> 0, "AA" -> 26).
let columnIndexOfLetters (letters: string) =
    let mutable acc = 0
    for ch in letters do
        let c = Char.ToUpperInvariant ch
        if c >= 'A' && c <= 'Z' then acc <- acc * 26 + (int c - int 'A' + 1)
    acc - 1

/// "#RRGGBB" for an rgb / argb colour as the file has it; None when malformed.
let private hexColor (value: string) =
    let trimmed = value.TrimStart('#')
    let cleaned = if trimmed.Length = 8 then trimmed.Substring(2) else trimmed
    if cleaned.Length = 6 then Some ("#" + cleaned) else None

/// One corner of a range ("$B$7", "B", "7") as (row0, col0); a part that is
/// absent (whole column / whole row) is None.
let private parseCorner (text: string) : (int option * int option) option =
    let mutable col = 0
    let mutable row = 0
    let mutable hasCol = false
    let mutable hasRow = false
    let mutable valid = text <> ""
    for c in text do
        if c = '$' then ()
        elif Char.IsLetter c && not hasRow && col < 20000 then
            hasCol <- true
            col <- col * 26 + (int (Char.ToUpperInvariant c) - int 'A' + 1)
        elif Char.IsDigit c && row < 10000000 then
            hasRow <- true
            row <- row * 10 + (int c - int '0')
        else
            valid <- false
    if not valid || (not hasCol && not hasRow) || (hasRow && row = 0) then None
    else Some ((if hasRow then Some (row - 1) else None), (if hasCol then Some (col - 1) else None))

/// "A1", "$A$1:$C$9", "A:C" or "2:5" as 0-based inclusive bounds, clipped
/// to `maxRow` / `maxCol`. None when malformed or wholly outside the clip.
let parseRangeClipped (maxRow: int) (maxCol: int) (range: string) : (int * int * int * int) option =
    let parts = range.Trim().Split(':')
    let corners =
        match parts.Length with
        | 1 -> parseCorner parts.[0] |> Option.map (fun c -> c, c)
        | 2 ->
            match parseCorner parts.[0], parseCorner parts.[1] with
            | Some a, Some b -> Some (a, b)
            | _ -> None
        | _ -> None
    corners
    |> Option.bind (fun ((r1, c1), (r2, c2)) ->
        let sr = r1 |> Option.defaultValue 0
        let sc = c1 |> Option.defaultValue 0
        let er = r2 |> Option.defaultValue maxRow
        let ec = c2 |> Option.defaultValue maxCol
        let sr, er = min sr er, max sr er
        let sc, ec = min sc ec, max sc ec
        if sr > maxRow || sc > maxCol then None
        else Some (sr, sc, min er maxRow, min ec maxCol))

/// The ranges of an sqref ("A1:A9 C3"), clipped.
let private parseSqref (maxRow: int) (maxCol: int) (sqref: string) =
    sqref.Split([| ' ' |], StringSplitOptions.RemoveEmptyEntries)
    |> Array.choose (parseRangeClipped maxRow maxCol)

/// A plain cell reference without `$` ("B7") as (row0, col0).
let private parsePlainRef (text: string) : (int * int) option =
    let mutable col = 0
    let mutable row = 0
    let mutable hasCol = false
    let mutable hasRow = false
    let mutable valid = text <> ""
    for c in text do
        if Char.IsDigit c then
            hasRow <- true
            if row < 100000000 then row <- row * 10 + (int c - int '0')
        elif Char.IsLetter c && not hasRow then
            hasCol <- true
            if col < 100000 then col <- col * 26 + (int (Char.ToUpperInvariant c) - int 'A' + 1)
        else
            valid <- false
    if valid && hasCol && hasRow && row > 0 && col > 0 then Some (row - 1, col - 1) else None

let private parsePlainRange (range: string) : (int * int * int * int) option =
    let parts = range.Split(':')
    match parts.Length with
    | 1 -> parsePlainRef parts.[0] |> Option.map (fun (r, c) -> r, c, r, c)
    | 2 ->
        match parsePlainRef parts.[0], parsePlainRef parts.[1] with
        | Some (r1, c1), Some (r2, c2) -> Some (min r1 r2, min c1 c2, max r1 r2, max c1 c2)
        | _ -> None
    | _ -> None

/// Sheet id of the sheet at `index`. Must agree with the ids the Spreadsheet
/// component assigns (`sheetN`, 1-based).
let sheetIdOfIndex (index: int) : string =
    sprintf "sheet%d" (index + 1)

let sheetStateText (state: XlsxReader.SheetState) =
    match state with
    | XlsxReader.SheetVisible -> "visible"
    | XlsxReader.SheetHidden -> "hidden"
    | XlsxReader.SheetVeryHidden -> "veryHidden"

// ---------------------------------------------------------------------------
// Import: reader workbook -> models
// ---------------------------------------------------------------------------

/// One sheet of an imported workbook. `State` is "visible", "hidden" or
/// "veryHidden"; hidden sheets are imported (formulas refer to them) but a
/// host should not offer a tab for them.
type ImportedSheet = {
    Name: string
    Model: SpreadsheetModel
    State: string
    Visible: bool
}

/// What an export needs of the workbook the sheets were imported from: the
/// reader's workbook, and each sheet's model exactly as it was installed (so
/// that "the user did not touch this" can be told from "this changed").
type XlsxSource = {
    /// Identifies this import: the models of its sheets carry it in
    /// `SourceSheet.Book`, so a model is only matched with the file it came from.
    Id: int
    Workbook: XlsxReader.Workbook
    /// Parallel to `Workbook.Sheets`.
    Baseline: SpreadsheetModel array
    /// The defined names as imported (the context's `Names`).
    BaselineNames: Map<Wb.SheetId option * string, string>
    /// ... and their spelling, comments and hidden flags (the context's `NameInfo`).
    BaselineNameInfo: Map<Wb.SheetId option * string, Wb.DefinedNameInfo>
}

/// A whole imported workbook: every sheet in file order, and the workbook
/// context (sheets, structured-reference tables, defined names) the formula
/// engine needs to evaluate anything that looks outside its own sheet.
type ImportedWorkbook = {
    Sheets: ImportedSheet array
    Context: Wb.Workbook
    /// The file used the 1904 date system. Date serials have already been
    /// converted to the 1900 system the engine uses.
    Date1904: bool
    /// For `Export.toWorkbookOut`.
    Source: XlsxSource
}

/// How large the models are made around the used range.
type ImportOptions = {
    MinRows: int
    MinCols: int
    /// Empty rows / columns added beyond the used range.
    ExtraRows: int
    ExtraCols: int
}

let mutable private lastSourceId = 0

/// A fresh `XlsxSource.Id`.
let private nextSourceId () =
    lastSourceId <- lastSourceId + 1
    lastSourceId

/// A grid as large as its content (at least 100 x 26): the web component.
let defaultImportOptions = { MinRows = 100; MinCols = 26; ExtraRows = 0; ExtraCols = 0 }

/// Room to keep typing below and to the right of the content: the native editors.
let editorImportOptions = { MinRows = 200; MinCols = 40; ExtraRows = 200; ExtraCols = 50 }

/// Workbook context for a set of sheet models: every table on every sheet is
/// registered for structured references. `tableColumns` supplies declared
/// column names by table name; a table without them reads its header row.
let buildWorkbookContext (models: SpreadsheetModel array) (names: string array) (visible: bool array)
                         (tableColumns: Map<string, string list>)
                         (definedNames: Map<Wb.SheetId option * string, string>)
                         (activeIndex: int) : Wb.Workbook =
    let sheetEntries =
        [ for i in 0 .. models.Length - 1 ->
            let id = sheetIdOfIndex i
            id, ({ Id = id
                   Name = (if i < names.Length then names.[i] else sprintf "Sheet%d" (i + 1))
                   Model = models.[i]
                   Visible = (if i < visible.Length then visible.[i] else true) } : Wb.Sheet) ]
    let tables =
        [ for i in 0 .. models.Length - 1 do
            for t in models.[i].Tables do
                if t.Name <> "" then
                    let columns = Map.tryFind t.Name tableColumns |> Option.defaultValue []
                    yield t.Name,
                          Wb.makeTableFromRegion (sheetIdOfIndex i) t.Name t.StartRow t.StartCol t.EndRow t.EndCol
                              t.HeaderRowCount (if t.HasTotals then 1 else 0) columns models.[i] ]
    { Sheets = Map.ofList sheetEntries
      SheetOrder = sheetEntries |> List.map fst
      ActiveSheet = sheetIdOfIndex (max 0 (min (models.Length - 1) activeIndex))
      Tables = Map.ofList tables
      Names = definedNames
      NameInfo = Map.empty }

/// Pixels at 100% per EMU (the unit drawing anchors use in the file).
let private emuPerPixel = 9525.0

/// A reader anchor (EMU) as a model anchor (pixels).
let private anchorOfReader (anchor: XlsxReader.SheetAnchor) : DrawingAnchor =
    { Kind = anchor.Kind
      FromRow = anchor.FromRow
      FromCol = anchor.FromCol
      FromOffsetX = anchor.FromColOffset / emuPerPixel
      FromOffsetY = anchor.FromRowOffset / emuPerPixel
      ToRow = anchor.ToRow
      ToCol = anchor.ToCol
      ToOffsetX = anchor.ToColOffset / emuPerPixel
      ToOffsetY = anchor.ToRowOffset / emuPerPixel
      X = anchor.X / emuPerPixel
      Y = anchor.Y / emuPerPixel
      Width = anchor.Width |> Option.map (fun emu -> emu / emuPerPixel)
      Height = anchor.Height |> Option.map (fun emu -> emu / emuPerPixel)
      EditAs = anchor.EditAs }

/// A model anchor as a reader anchor; `order` is its z-order on the sheet.
let private anchorToReader (order: int) (anchor: DrawingAnchor) : XlsxReader.SheetAnchor =
    let emu (pixels: float) = Math.Round (pixels * emuPerPixel)
    { Kind = anchor.Kind
      FromCol = anchor.FromCol
      FromColOffset = emu anchor.FromOffsetX
      FromRow = anchor.FromRow
      FromRowOffset = emu anchor.FromOffsetY
      ToCol = anchor.ToCol
      ToColOffset = emu anchor.ToOffsetX
      ToRow = anchor.ToRow
      ToRowOffset = emu anchor.ToOffsetY
      X = emu anchor.X
      Y = emu anchor.Y
      Width = anchor.Width |> Option.map emu
      Height = anchor.Height |> Option.map emu
      EditAs = anchor.EditAs
      Order = order }

/// The sheet's pictures and kept objects as model drawings, back to front.
let private drawingsOfExtras (extras: XlsxReader.SheetExtras) : SheetDrawing list =
    (extras.Images
     |> List.map (fun image ->
        image.Anchor.Order, 0,
        ({ Anchor = anchorOfReader image.Anchor
           Name = image.Name
           Description = image.Description
           Content = Picture (image.ContentType, image.Extension, image.Data) } : SheetDrawing)))
    @ (extras.Objects
       |> List.map (fun item ->
        item.Anchor.Order, 1,
        ({ Anchor = anchorOfReader item.Anchor
           Name = item.Name
           Description = None
           Content = KeptObject (item.Kind, item.Xml) } : SheetDrawing)))
    |> List.sortBy (fun (order, kind, _) -> order, kind)
    |> List.map (fun (_, _, drawing) -> drawing)

let rec private threadOfReader (comment: XlsxReader.ThreadedComment) : CommentThread =
    { Id = comment.Id
      Author = comment.Author
      PersonId = comment.PersonId
      Date = comment.Date
      Text = comment.Text
      Replies = comment.Replies |> List.map threadOfReader }

let rec private threadToReader (addr: CellAddress) (thread: CommentThread) : XlsxReader.ThreadedComment =
    { CellRef = { Column = columnToLetter addr.Col; Row = addr.Row + 1 }
      Id = thread.Id
      Author = thread.Author
      PersonId = thread.PersonId
      Date = thread.Date
      Text = thread.Text
      Replies = thread.Replies |> List.map (threadToReader addr) }

/// The `<from>` / `<to>` marker of a kept object's anchor element, rewritten
/// for a new cell position. The element is otherwise written back unchanged.
let private rewriteMarker (xml: string) (name: string) (col: int) (colOffset: float) (row: int) (rowOffset: float) : string =
    // The marker element with whatever prefix the file uses: <xdr:from> ... </xdr:from>.
    let closeAt = xml.IndexOf (":" + name + ">")
    let openAt = if closeAt < 0 then -1 else xml.LastIndexOf ('<', closeAt)
    if openAt < 0 || xml.Substring(openAt + 1, closeAt - openAt - 1).Contains "/" then xml
    else
        let prefix = xml.Substring(openAt + 1, closeAt - openAt - 1)
        let closing = sprintf "</%s:%s>" prefix name
        let endAt = xml.IndexOf (closing, closeAt)
        if endAt < 0 then xml
        else
            let emu (pixels: float) = string (int64 (Math.Round (pixels * emuPerPixel)))
            let marker =
                sprintf "<%s:%s><%s:col>%d</%s:col><%s:colOff>%s</%s:colOff><%s:row>%d</%s:row><%s:rowOff>%s</%s:rowOff>%s"
                    prefix name prefix col prefix prefix (emu colOffset) prefix prefix row prefix prefix (emu rowOffset) prefix closing
            xml.Substring(0, openAt) + marker + xml.Substring(endAt + closing.Length)

/// The drawing objects and threaded comments of a model in the reader's
/// shape, for the writer. `original` supplies what only the file knows (the
/// relationships kept objects refer to, sheetFormatPr, column definitions).
let private extrasOfModel (original: XlsxReader.SheetExtras) (before: SpreadsheetModel option) (model: SpreadsheetModel) : XlsxReader.SheetExtras =
    let drawingsUnchanged =
        match before with
        | Some baseline -> obj.ReferenceEquals(baseline.Drawings, model.Drawings) || baseline.Drawings = model.Drawings
        | None -> false
    let threadsUnchanged =
        match before with
        | Some baseline -> obj.ReferenceEquals(baseline.Threads, model.Threads) || baseline.Threads = model.Threads
        | None -> false
    let images, objects =
        if drawingsUnchanged then original.Images, original.Objects
        else
            let images = ResizeArray<XlsxReader.SheetImage>()
            let objects = ResizeArray<XlsxReader.SheetObject>()
            model.Drawings
            |> List.iteri (fun order drawing ->
                let anchor = anchorToReader order drawing.Anchor
                match drawing.Content with
                | Picture (contentType, extension, data) ->
                    images.Add { Anchor = anchor; Name = drawing.Name; Description = drawing.Description; ContentType = contentType; Extension = extension; Data = data }
                | KeptObject (kind, xml) ->
                    // The element goes back as it was, at the object's current cells.
                    let moved =
                        if drawing.Anchor.Kind = "absolute" then xml
                        else
                            let withFrom = rewriteMarker xml "from" drawing.Anchor.FromCol drawing.Anchor.FromOffsetX drawing.Anchor.FromRow drawing.Anchor.FromOffsetY
                            match drawing.Anchor.ToCol, drawing.Anchor.ToRow with
                            | Some toCol, Some toRow -> rewriteMarker withFrom "to" toCol drawing.Anchor.ToOffsetX toRow drawing.Anchor.ToOffsetY
                            | _ -> withFrom
                    objects.Add { Anchor = anchor; Kind = kind; Name = drawing.Name; Xml = moved })
            List.ofSeq images, List.ofSeq objects
    let threads =
        if threadsUnchanged then original.ThreadedComments
        else model.Threads |> Map.toList |> List.map (fun (addr, thread) -> threadToReader addr thread)
    { original with Images = images; Objects = objects; ThreadedComments = threads }

module Import =

    type private SheetImport = {
        Model: SpreadsheetModel
        /// `<tableColumn>` names per table, for structured references.
        TableColumns: Map<string, string list>
        State: string
    }

    let private border (side: XlsxReader.CellBorderSide option) : BorderStyle option =
        side
        |> Option.bind (fun b ->
            if String.IsNullOrWhiteSpace b.Style || b.Style = "none" then None
            else Some { Style = b.Style; Color = b.Color |> Option.bind hexColor })

    let private format (s: XlsxReader.CellStyle) : CellFormat =
        { defaultFormat with
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

    /// A run of rich text. Only what the run turns on is recorded; the rest
    /// follows the cell's own format.
    let private runOverride (props: XlsxReader.RunProps) : RunOverride =
        let onlyTrue value = if value then Some true else None
        { Bold = onlyTrue props.Bold
          Italic = onlyTrue props.Italic
          Underline = onlyTrue props.Underline
          UnderlineDouble = onlyTrue props.UnderlineDouble
          Strikethrough = onlyTrue props.Strikethrough
          Color = props.Color |> Option.bind hexColor
          FontName = props.FontName
          FontSize = props.FontSize
          VertAlign =
            match props.VertAlign with
            | XlsxReader.VertAlignSuperscript -> Some VertAlignSuperscript
            | XlsxReader.VertAlignSubscript -> Some VertAlignSubscript
            | XlsxReader.VertAlignBaseline -> None }

    /// Value of a non-formula cell, or the result a formula cell cached in the file.
    let private storedValue (wb: XlsxReader.Workbook) (value: XlsxReader.CellValue) : CellValue =
        match value with
        | XlsxReader.Error e -> CellValue.Error (e.TrimStart '#')
        | _ ->
            match XlsxReader.resolveValue wb value with
            | XlsxReader.Number n -> CellValue.Number n
            | XlsxReader.String s -> CellValue.String s
            | XlsxReader.Boolean b -> CellValue.Boolean b
            | XlsxReader.Date d -> CellValue.String d
            | _ -> CellValue.Empty

    /// A formula cell keeps the result the file cached for it: that is what
    /// Excel shows on open, and it stays correct where our engine lacks a
    /// function. Only a formula the file has no result for is left with an
    /// Empty value, for the workbook-level pass to compute.
    let private cellValue (wb: XlsxReader.Workbook) (cell: XlsxReader.Cell) : CellValue =
        match cell.Formula, cell.Value with
        | Some formula, XlsxReader.Formula _ when not (String.IsNullOrWhiteSpace formula) -> CellValue.Formula (formula, CellValue.Empty)
        | Some formula, value when not (String.IsNullOrWhiteSpace formula) -> CellValue.Formula (formula, storedValue wb value)
        | None, XlsxReader.Formula formula when not (String.IsNullOrWhiteSpace formula) -> CellValue.Formula (formula, CellValue.Empty)
        | _, value ->
            match storedValue wb value with
            | CellValue.String "" -> CellValue.Empty
            | other -> other

    let private tableFills (workbook: XlsxReader.Workbook) (styleName: string) : string * string =
        // Excel's TableStyle{Light,Medium,Dark}N cycle through black and the six accents (N mod 7).
        let mutable n = 0
        let mutable any = false
        for ch in styleName do
            if Char.IsDigit ch && n < 100000 then
                any <- true
                n <- n * 10 + (int ch - int '0')
        let n = if any then n else 2
        let slot = (max 1 n - 1) % 7
        let accent =
            if slot = 0 || slot - 1 >= workbook.ThemeAccents.Length then "000000"
            else workbook.ThemeAccents.[slot - 1]
        "#" + accent, "#" + XlsxReader.Xlsx.applyColorTint accent 0.8

    let private tables (workbook: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : Table list * Map<string, string list> =
        let mutable columns = Map.empty<string, string list>
        let converted =
            sheet.Tables
            |> List.map (fun t ->
                columns <- Map.add t.Name t.Columns columns
                let style =
                    t.Style
                    |> Option.map (fun s ->
                        // HeaderFill / StripeFill make `applyTableStyles` paint the
                        // table at render time; nothing is baked into the cells.
                        let headerFill, stripeFill = tableFills workbook s.Name
                        ({ Name = s.Name
                           ShowFirstColumn = s.ShowFirstColumn
                           ShowLastColumn = s.ShowLastColumn
                           ShowRowStripes = s.ShowRowStripes
                           ShowColumnStripes = s.ShowColumnStripes
                           HeaderFill = Some headerFill
                           StripeFill = Some stripeFill } : TableStyle))
                ({ Name = t.Name
                   StartRow = t.StartRow
                   StartCol = t.StartCol
                   EndRow = t.EndRow
                   EndCol = t.EndCol
                   HeaderRowCount = t.HeaderRowCount
                   Style = style
                   AutoFilter = t.AutoFilter
                   HasTotals = t.TotalsRowCount > 0 } : Table))
        converted, columns

    /// The sheet's conditional formatting as model rules, kept by range (a
    /// whole-column rule stays one).
    let private conditionalFormats (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : ConditionalFormatRange list =
        sheet.ConditionalFormats
        |> List.map (fun cf ->
            ({ Sqref = cf.Range
               Rules =
                cf.Rules
                |> List.map (fun rule ->
                    let dxf =
                        rule.DxfId
                        |> Option.bind (fun id -> if id >= 0 && id < wb.Dxfs.Length then Some wb.Dxfs.[id] else None)
                    ({ Type = rule.Type
                       Operator = rule.Operator
                       Formulas = rule.Formulas
                       Text = rule.Text
                       Rank = rule.Rank
                       Percent = rule.Percent
                       Bottom = rule.Bottom
                       AboveAverage = rule.AboveAverage
                       EqualAverage = rule.EqualAverage
                       StdDev = rule.StdDev
                       ColorScale = rule.ColorScale |> List.map (fun stop -> stop.CfvoType, stop.CfvoValue, stop.Color)
                       DataBarColor = rule.DataBarColor
                       DataBarCfvo = rule.DataBarCfvo
                       IconSet = rule.IconSet
                       IconSetCfvo = rule.IconSetCfvo
                       IconSetReverse = rule.IconSetReverse
                       TimePeriod = rule.TimePeriod
                       Priority = rule.Priority
                       StopIfTrue = rule.StopIfTrue
                       Style =
                        dxf |> Option.map (fun d ->
                            ({ FontColor = d.FontColor |> Option.bind hexColor
                               FillColor = d.FillColor |> Option.bind hexColor
                               Bold = (d.Bold = Some true)
                               Italic = (d.Italic = Some true)
                               Underline = (d.Underline = Some true)
                               Strikethrough = (d.Strike = Some true) } : CF.OoxmlStyle)) } : CF.OoxmlRule)) } : CF.OoxmlConditionalFormat))
        |> CF.ofOoxml

    let private validationOoxml (dv: XlsxReader.DataValidation) : CF.OoxmlValidation =
        { Type = dv.ValidationType
          Operator = dv.Operator
          Formula1 = dv.Formula1
          Formula2 = dv.Formula2
          AllowBlank = dv.AllowBlank
          ShowErrorMessage = dv.ShowErrorMessage
          ShowInputMessage = dv.ShowInputMessage
          ShowDropDown = dv.ShowDropDown
          ErrorStyle = dv.ErrorStyle
          ErrorTitle = dv.ErrorTitle
          Error = dv.ErrorMessage
          PromptTitle = dv.PromptTitle
          Prompt = dv.Prompt }

    /// The sheet's validations, each for the ranges the file gives it (a
    /// whole column stays a whole column), numbered in file order. A list
    /// whose values come from a range or a name keeps its source; it is read
    /// when the values are needed.
    let private validations (sheet: XlsxReader.Worksheet) : ValidationRange list =
        [ for dv in sheet.DataValidations do
            match CF.parseSqref dv.Range with
            | [] -> ()
            | ranges ->
                let ooxml = validationOoxml dv
                match CF.validationRuleOfOoxml ooxml with
                | Some rule -> yield ranges, CF.validationWithOoxml ranges.Head.From ooxml rule
                | None -> () ]
        |> List.mapi (fun index (ranges, validation) -> { Id = index + 1; Ranges = ranges; Validation = { validation with Anchor = None } })

    let private pageSetup (sheet: XlsxReader.Worksheet) : PageSetup =
        let setup =
            match sheet.PageSetup with
            | None -> defaultPageSetup
            | Some p ->
                let w, h =
                    p.PaperSize
                    |> Option.map paperSizePoints
                    |> Option.defaultValue (defaultPageSetup.PageWidth, defaultPageSetup.PageHeight)
                { defaultPageSetup with
                    PageWidth = w
                    PageHeight = h
                    Orientation = (match p.Orientation with XlsxReader.OrientationLandscape -> LandscapeOrientation | _ -> PortraitOrientation)
                    FitToPage = p.FitToPage
                    FitToWidth = p.FitToWidth
                    FitToHeight = p.FitToHeight
                    Scale = p.Scale
                    FirstPageNumber = p.FirstPageNumber
                    UseFirstPageNumber = p.UseFirstPageNumber
                    PageOrder = (match p.PageOrder with XlsxReader.OverThenDown -> OverThenDown | _ -> DownThenOver)
                    BlackAndWhite = p.BlackAndWhite
                    CellComments =
                        match p.CellComments with
                        | XlsxReader.CellCommentsAsDisplayed -> CellCommentsAsDisplayed
                        | XlsxReader.CellCommentsAtEnd -> CellCommentsAtEnd
                        | _ -> CellCommentsNone }
        let margins =
            sheet.PageMargins
            |> Option.map (fun m -> ({ Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer } : PageMargins))
            |> Option.defaultValue defaultPageMargins
        let wholeSheet = parseRangeClipped 1048575 16383
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
                |> Option.bind wholeSheet
                |> Option.map (fun (sr, sc, er, ec) -> ({ From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } : CellRange))
            RepeatRows = sheet.PrintTitleRows |> Option.bind wholeSheet |> Option.map (fun (sr, _, er, _) -> sr, er)
            RepeatCols = sheet.PrintTitleCols |> Option.bind wholeSheet |> Option.map (fun (_, sc, _, ec) -> sc, ec) }

    let private importSheet (options: ImportOptions) (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : SheetImport =
        // Used range, one-based counts: every <row> counts, with or without cells.
        let mutable rowCount = 0
        let mutable colCount = 0
        for row in sheet.Rows do
            if row.Index > rowCount then rowCount <- row.Index
            for cell in row.Cells do
                let col = columnIndexOfLetters cell.Reference.Column + 1
                if col > colCount then colCount <- col

        let comments = Collections.Generic.Dictionary<int * int, CellComment>()
        let commentCells = Collections.Generic.HashSet<int * int>()
        for c in sheet.Comments do
            let r = c.CellRef.Row - 1
            let col = columnIndexOfLetters c.CellRef.Column
            if r >= 0 && col >= 0 then
                commentCells.Add((r, col)) |> ignore
                // The first comment of a cell is the one shown.
                if not (comments.ContainsKey((r, col))) && not (String.IsNullOrWhiteSpace c.Text) then
                    comments.[(r, col)] <- { Text = c.Text; Author = c.Author }
        // One CellFormat per style of the workbook, shared by every cell that uses it.
        let styleCache = Collections.Generic.Dictionary<int, CellFormat>()
        let styleFormat (cell: XlsxReader.Cell) =
            let index =
                match cell.StyleIndex with
                | Some idx when idx >= 0 && idx < wb.Styles.Length -> idx
                | None when wb.Styles.Length > 0 && cell.Style.IsSome -> 0
                | _ -> -1
            if index < 0 then defaultFormat
            else
                match styleCache.TryGetValue index with
                | true, cached -> cached
                | false, _ ->
                    let converted = format wb.Styles.[index]
                    styleCache.[index] <- converted
                    converted

        let mutable cellMap = Map.empty<CellAddress, CellData>
        let mutable richRuns = Map.empty<CellAddress, (string * RunOverride) list>
        // The other cells of a legacy array formula (<f t="array" ref="F94:H94">)
        // hold results the anchor's formula spilled there: they belong to the
        // anchor, so a recalculation spills over them again instead of finding
        // them in its way (#SPILL!).
        let mutable spill = Map.empty<CellAddress, CellAddress>
        for row in sheet.Rows do
            for cell in row.Cells do
                let r = cell.Reference.Row - 1
                let c = columnIndexOfLetters cell.Reference.Column
                if r >= 0 && c >= 0 then
                    match cell.Formula, cell.ArrayRef |> Option.bind parsePlainRange with
                    | Some _, Some (sr, sc, er, ec) when (er > sr || ec > sc) && (er - sr + 1) * (ec - sc + 1) <= 100000 ->
                        let anchor : CellAddress = { Row = r; Col = c }
                        for ghostRow in sr .. er do
                            for ghostCol in sc .. ec do
                                if ghostRow <> r || ghostCol <> c then
                                    spill <- Map.add { Row = ghostRow; Col = ghostCol } anchor spill
                    | _ -> ()
                    let value = cellValue wb cell
                    let baseFormat = styleFormat cell
                    let hasComment = commentCells.Contains((r, c))
                    let comment =
                        match comments.TryGetValue((r, c)) with
                        | true, found -> Some found
                        | false, _ -> None
                    let cellFormat =
                        if (hasComment || comment.IsSome) <> baseFormat.HasComment then
                            { baseFormat with HasComment = baseFormat.HasComment || hasComment || comment.IsSome }
                        else baseFormat
                    // Keep style-only cells: Excel renders fills, borders, comment
                    // markers and validation dropdowns even when the cell has no value.
                    if value <> CellValue.Empty || cellFormat <> defaultFormat || comment.IsSome then
                        let addr : CellAddress = { Row = r; Col = c }
                        cellMap <- Map.add addr { Value = value; Format = cellFormat; Comment = comment } cellMap
                        let runs =
                            if not (List.isEmpty cell.InlineRuns) then cell.InlineRuns
                            else
                                match cell.Value with
                                | XlsxReader.SharedString idx when idx >= 0 && idx < wb.SharedStringRuns.Length -> wb.SharedStringRuns.[idx]
                                | _ -> []
                        match runs with
                        | _ :: _ :: _ -> richRuns <- Map.add addr (runs |> List.map (fun (text, props) -> text, runOverride props)) richRuns
                        | _ -> ()

        // Column widths: file units are characters, the model's are pixels (~7 px per character).
        let defaultColWidth =
            match sheet.Format.DefaultColWidth with
            | Some width -> width
            | None ->
                match sheet.Format.BaseColWidth with
                | Some width -> width + 1.0
                | None -> 8.43
        let mutable columnWidths = Map.empty<int, float>
        if colCount > 0 then
            let widths = Array.create colCount defaultColWidth
            for column in sheet.Columns do
                let first = max 0 (column.Index - 1)
                let last = min (colCount - 1) (column.MaxIndex - 1)
                if first <= last then
                    let width = if column.Hidden then 0.0 else column.Width |> Option.defaultValue 0.0
                    for col = first to last do
                        widths.[col] <- width
            for col in 0 .. colCount - 1 do
                if widths.[col] > 0.0 then columnWidths <- Map.add col (widths.[col] * 7.0) columnWidths
        // Row heights: points in the file, pixels in the model.
        let defaultRowHeight = sheet.Format.DefaultRowHeight |> Option.defaultValue 15.0
        let mutable rowHeights = Map.empty<int, float>
        if rowCount > 0 then
            let heights = Array.create rowCount defaultRowHeight
            for row in sheet.Rows do
                let index = row.Index - 1
                if index >= 0 && index < rowCount then
                    heights.[index] <- row.Height |> Option.defaultValue defaultRowHeight
            for row in 0 .. rowCount - 1 do
                if heights.[row] > 0.0 then rowHeights <- Map.add row (heights.[row] * 1.333) rowHeights
        // A size of 0 is how the model marks a hidden row / column.
        for row in sheet.Rows do
            if row.Hidden then rowHeights <- Map.add (row.Index - 1) 0.0 rowHeights
        for column in sheet.Columns do
            // A hidden range can run to the last column (16384); only the ones near the data matter.
            if column.Hidden then
                for idx in column.Index .. min column.MaxIndex (column.Index + 1023) do
                    columnWidths <- Map.add (idx - 1) 0.0 columnWidths

        let merged =
            sheet.MergedCells
            |> List.choose (fun m ->
                let sr = m.StartRef.Row - 1
                let sc = columnIndexOfLetters m.StartRef.Column
                let er = m.EndRef.Row - 1
                let ec = columnIndexOfLetters m.EndRef.Column
                if sr >= 0 && sc >= 0 && er >= 0 && ec >= 0 then Some ({ StartRow = sr; StartCol = sc; EndRow = er; EndCol = ec } : MergedCell) else None)

        let modelRows = max (rowCount + options.ExtraRows) options.MinRows
        let modelCols = max (colCount + options.ExtraCols) options.MinCols

        let rules = validations sheet
        let sheetTables, tableColumns = tables wb sheet

        // Link cells are shown the way Excel styles them.
        for link in sheet.Hyperlinks do
            match parseRangeClipped (modelRows - 1) (modelCols - 1) link.Range with
            | Some (sr, sc, er, ec) when (er - sr + 1) * (ec - sc + 1) <= 4096 ->
                for r in sr .. er do
                    for c in sc .. ec do
                        let addr : CellAddress = { Row = r; Col = c }
                        match Map.tryFind addr cellMap with
                        | Some data ->
                            let linkFormat =
                                { data.Format with
                                    TextColor = data.Format.TextColor |> Option.orElse (Some "#0563C1")
                                    Underline = true }
                            cellMap <- Map.add addr { data with Format = linkFormat } cellMap
                        | None -> ()
            | _ -> ()

        // Threaded comments, by cell. The cell shows a comment marker, whether
        // or not the file also has the legacy placeholder comment for it.
        let mutable threads = Map.empty<CellAddress, CommentThread>
        for thread in sheet.Extras.ThreadedComments do
            let r = thread.CellRef.Row - 1
            let c = columnIndexOfLetters thread.CellRef.Column
            if r >= 0 && c >= 0 && not (Map.containsKey { Row = r; Col = c } threads) then
                let addr : CellAddress = { Row = r; Col = c }
                threads <- Map.add addr (threadOfReader thread) threads
                match Map.tryFind addr cellMap with
                | Some data when not data.Format.HasComment -> cellMap <- Map.add addr { data with Format = { data.Format with HasComment = true } } cellMap
                | Some _ -> ()
                | None -> cellMap <- Map.add addr { Value = CellValue.Empty; Format = { defaultFormat with HasComment = true }; Comment = None } cellMap

        // Hyperlinks, by range: they follow their cells from here on.
        let hyperlinks =
            sheet.Hyperlinks
            |> List.choose (fun link ->
                parseRangeClipped 1048575 16383 link.Range
                |> Option.map (fun (sr, sc, er, ec) ->
                    ({ Range = { From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } }
                       Target = link.Target
                       IsInternal = link.IsInternal } : SheetHyperlink)))
        // Default cell styles of columns, in the order `XlsxWriter.
        // packageFromReaderWorkbook` lists them (`Source` is that position).
        let columnStyles =
            sheet.Columns
            |> List.choose (fun column ->
                match column.Style with
                | Some index when index > 0 && index < wb.Styles.Length -> Some (column, index)
                | _ -> None)
            |> List.mapi (fun position (column, index) ->
                ({ First = column.Index - 1
                   Last = column.MaxIndex - 1
                   Format =
                    (match styleCache.TryGetValue index with
                     | true, cached -> cached
                     | false, _ -> format wb.Styles.[index])
                   Source = position } : ColumnStyle))
        // What-if data tables: the cells of the range hold the file's results.
        let dataTables =
            sheet.Extras.DataTables
            |> List.choose (fun table ->
                parsePlainRange table.Ref
                |> Option.map (fun (sr, sc, er, ec) ->
                    ({ Range = { From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } }
                       TwoDimensional = table.TwoDimensional
                       RowInput = table.RowInput
                       Input1 = table.Input1
                       Input2 = table.Input2
                       Input1Deleted = table.Input1Deleted
                       Input2Deleted = table.Input2Deleted
                       CalculateAlways = table.CalculateAlways } : DataTable)))
        // Where the file's pivot tables sit.
        let keptRanges =
            sheet.Extras.PivotLocations
            |> List.choose (fun (path, range) ->
                parsePlainRange range
                |> Option.map (fun (sr, sc, er, ec) ->
                    ({ Key = path; Range = { From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } } : KeptRange)))
        // The ranges pivot caches read from this sheet follow its rows and columns too.
        let keptRanges =
            keptRanges
            @ (wb.Preserved.PivotCacheSources
               |> List.choose (fun (path, sheetName, range) ->
                   if not (sheetName.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase)) then None
                   else
                       parsePlainRange (range.Replace("$", ""))
                       |> Option.map (fun (sr, sc, er, ec) ->
                           ({ Key = path; Range = { From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } } : KeptRange))))

        let model =
            { createEmpty modelRows modelCols with
                Hyperlinks = hyperlinks
                ColumnStyles = columnStyles
                DataTables = dataTables
                KeptRanges = keptRanges
                Cells = cellMap
                RichRuns = richRuns
                MergedCells = merged
                RowHeights = rowHeights
                ColumnWidths = columnWidths
                Tables = sheetTables
                FrozenRows = sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0
                FrozenCols = sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0
                ShowGridLines = sheet.ShowGridLines
                Validations = rules
                ConditionalFormats = conditionalFormats wb sheet
                PageSetup = pageSetup sheet
                Drawings = drawingsOfExtras sheet.Extras
                Threads = threads
                Spill = spill }
        { Model = model
          TableColumns = tableColumns
          State = sheetStateText sheet.State }

    let private isUncached (cell: CellData) =
        match cell.Value with
        | CellValue.Formula (_, CellValue.Empty) -> true
        | _ -> false

    /// Compute the formulas the file carries no result for (workbooks written by
    /// libraries rather than by Excel). The calculation runs over the WHOLE
    /// workbook - a sheet recalculated on its own turns every cross-sheet
    /// reference into #REF! and every defined name into #NAME? - and afterwards
    /// every formula that did have a cached result gets it back, so values Excel
    /// computed are never replaced by ours.
    let private computeUncachedFormulas (context: Wb.Workbook) : Wb.Workbook =
        let needsWork =
            context.Sheets |> Map.exists (fun _ sheet -> sheet.Model.Cells |> Map.exists (fun _ cell -> isUncached cell))
        if not needsWork then context
        else
            let recalculated = Rbt.Spreadsheet.Formula.recalculateWorkbook context
            let sheets =
                recalculated.Sheets
                |> Map.map (fun id sheet ->
                    match Map.tryFind id context.Sheets with
                    | None -> sheet
                    | Some original ->
                        let mutable cells = sheet.Model.Cells
                        for KeyValue(addr, cell) in original.Model.Cells do
                            match cell.Value with
                            | CellValue.Formula (_, cached) when cached <> CellValue.Empty ->
                                cells <- Map.add addr cell cells
                            | _ -> ()
                        { sheet with Model = { sheet.Model with Cells = cells } })
            { recalculated with Sheets = sheets }

    /// Defined names of a reader workbook, keyed as the workbook context keys them.
    let private definedNames (wb: XlsxReader.Workbook) =
        wb.NamedRanges
        |> List.map (fun entry ->
            let scope =
                entry.Scope
                |> Option.bind (fun text ->
                    match Int32.TryParse text with
                    | true, index -> Some (sheetIdOfIndex index)
                    | _ -> None)
            scope, entry.Name, entry.RefersTo)
        |> Wb.namesOfTriples

    /// Spelling, comment and hidden flag of a reader workbook's defined names.
    let private definedNameInfo (wb: XlsxReader.Workbook) : Map<Wb.SheetId option * string, Wb.DefinedNameInfo> =
        wb.NamedRanges
        |> List.map (fun entry ->
            let scope =
                entry.Scope
                |> Option.bind (fun text ->
                    match Int32.TryParse text with
                    | true, index -> Some (sheetIdOfIndex index)
                    | _ -> None)
            (scope, entry.Name.ToUpperInvariant()), ({ Name = entry.Name; Comment = entry.Comment; Hidden = entry.Hidden } : Wb.DefinedNameInfo))
        |> Map.ofList

    /// Convert a loaded workbook: every sheet, plus the workbook context.
    /// Throws when the workbook has no sheets.
    let importWorkbookWith (options: ImportOptions) (wb: XlsxReader.Workbook) : ImportedWorkbook =
        let readerSheets = List.toArray wb.Sheets
        if readerSheets.Length = 0 then failwith "XLSX workbook has no worksheets"
        let names = readerSheets |> Array.map (fun sheet -> sheet.Name)
        // Each model knows which sheet of the file it is, whatever happens to
        // its name or position later.
        let sourceId = nextSourceId ()
        let imports =
            readerSheets
            |> Array.mapi (fun index sheet ->
                let import = importSheet options wb sheet
                { import with Model = { import.Model with SourceSheet = Some { Book = sourceId; Sheet = index } } })
        let names' = definedNames wb

        let models = imports |> Array.map (fun import -> import.Model)

        let visible = imports |> Array.map (fun import -> import.State = "visible")
        let tableColumns =
            imports |> Array.fold (fun acc import -> Map.fold (fun m k v -> Map.add k v m) acc import.TableColumns) Map.empty
        let firstVisible = visible |> Array.tryFindIndex id |> Option.defaultValue 0
        let context =
            { buildWorkbookContext models names visible tableColumns names' firstVisible with NameInfo = definedNameInfo wb }
            |> computeUncachedFormulas
        let finalModels =
            Array.init names.Length (fun index ->
                match Map.tryFind (sheetIdOfIndex index) context.Sheets with
                | Some sheet -> sheet.Model
                | None -> models.[index])
        { Sheets =
            names
            |> Array.mapi (fun index name ->
                ({ Name = name
                   Model = finalModels.[index]
                   State = imports.[index].State
                   Visible = visible.[index] } : ImportedSheet))
          Context = context
          Date1904 = wb.Date1904
          Source = { Id = sourceId; Workbook = wb; Baseline = finalModels; BaselineNames = context.Names; BaselineNameInfo = context.NameInfo } }

    let importWorkbook (wb: XlsxReader.Workbook) : ImportedWorkbook =
        importWorkbookWith defaultImportOptions wb

    let private readerCellIsFormula (cell: XlsxReader.Cell) =
        cell.Formula.IsSome || (match cell.Value with XlsxReader.Formula _ -> true | _ -> false)
    /// The formula cells of a reader sheet, as `importSheet` makes them (see `cellValue`).
    let private readerSheetFormulas (sheet: XlsxReader.Worksheet) : (CellAddress * string) array =
        let formulas = ResizeArray<CellAddress * string>()
        for row in sheet.Rows do
            for cell in row.Cells do
                let r = cell.Reference.Row - 1
                let c = columnIndexOfLetters cell.Reference.Column
                if r >= 0 && c >= 0 then
                    match cell.Formula, cell.Value with
                    | Some formula, _ when not (String.IsNullOrWhiteSpace formula) -> formulas.Add ({ Row = r; Col = c }, formula)
                    | None, XlsxReader.Formula formula when not (String.IsNullOrWhiteSpace formula) -> formulas.Add ({ Row = r; Col = c }, formula)
                    | _ -> ()
        formulas.ToArray()
    let private readerCellIsUncached (cell: XlsxReader.Cell) =
        match cell.Value with XlsxReader.Formula _ -> true | _ -> false

    /// Convert a loaded workbook sheet by sheet, when each is first needed:
    /// `Sheets` holds a stand-in (`SpreadsheetModel.Deferred`) for every sheet
    /// `convertNow` does not ask for; `Model.resolveSheet` converts it (when
    /// it is shown, when a formula reads it). The reader's workbook is kept
    /// as `Source`, and an export writes a sheet that was never converted
    /// from it. The stand-ins carry the sheet's tables, so structured
    /// references resolve before their sheet is converted.
    ///
    /// A workbook with formulas the file has no values for (one written by a
    /// library rather than by Excel) is converted whole, as `importWorkbook`
    /// does: its values come from a calculation over every sheet.
    let importWorkbookLazy (convertNow: int -> bool) (wb: XlsxReader.Workbook) : ImportedWorkbook =
        let options = defaultImportOptions
        let readerSheets = List.toArray wb.Sheets
        if readerSheets.Length = 0 then failwith "XLSX workbook has no worksheets"
        let anyUncached =
            readerSheets |> Array.exists (fun sheet -> sheet.Rows |> List.exists (fun row -> row.Cells |> List.exists readerCellIsUncached))
        if anyUncached || readerSheets.Length = 1 then importWorkbookWith options wb
        else
            let names = readerSheets |> Array.map (fun sheet -> sheet.Name)
            let sourceId = nextSourceId ()
            let names' = definedNames wb
            let baseline : SpreadsheetModel array = Array.zeroCreate readerSheets.Length
            let standIns : SpreadsheetModel array = Array.zeroCreate readerSheets.Length
            let rec convert (index: int) : SpreadsheetModel =
                let deferred = standIns.[index].Deferred.Value
                match deferred.Loaded with
                | Some loaded -> loaded
                | None ->
                    let import = importSheet options wb readerSheets.[index]
                    let model = { import.Model with SourceSheet = Some { Book = sourceId; Sheet = index } }
                    deferred.Loaded <- Some model
                    // What an export compares the sheet with: as it was opened.
                    baseline.[index] <- model
                    model
            let tableColumns =
                readerSheets
                |> Array.mapi (fun index sheet ->
                    let sheetTables, columns = tables wb sheet
                    let hasFormulas =
                        not sheet.ConditionalFormats.IsEmpty || not sheet.DataValidations.IsEmpty
                        || sheet.Rows |> List.exists (fun row -> row.Cells |> List.exists readerCellIsFormula)
                    standIns.[index] <-
                        { createEmpty options.MinRows options.MinCols with
                            Tables = sheetTables
                            SourceSheet = Some { Book = sourceId; Sheet = index }
                            Deferred =
                                Some { Loaded = None
                                       Load = (fun () -> convert index)
                                       HasFormulas = hasFormulas
                                       Formulas = (fun () -> readerSheetFormulas sheet) } }
                    baseline.[index] <- standIns.[index]
                    columns)
                |> Array.fold (fun acc columns -> Map.fold (fun m k v -> Map.add k v m) acc columns) Map.empty
            let visible = readerSheets |> Array.map (fun sheet -> sheetStateText sheet.State = "visible")
            let firstVisible = visible |> Array.tryFindIndex id |> Option.defaultValue 0
            let models = standIns |> Array.mapi (fun index standIn -> if convertNow index then convert index else standIn)
            let context = { buildWorkbookContext models names visible tableColumns names' firstVisible with NameInfo = definedNameInfo wb }
            { Sheets =
                names
                |> Array.mapi (fun index name ->
                    ({ Name = name
                       Model = models.[index]
                       State = sheetStateText readerSheets.[index].State
                       Visible = visible.[index] } : ImportedSheet))
              Context = context
              Date1904 = wb.Date1904
              Source = { Id = sourceId; Workbook = wb; Baseline = baseline; BaselineNames = context.Names; BaselineNameInfo = context.NameInfo } }

    /// One sheet, with only that sheet as the context of its formulas - use
    /// `importWorkbook` when other sheets or defined names matter.
    let importSheetWith (options: ImportOptions) (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) : SpreadsheetModel =
        let import = importSheet options wb sheet
        let model = import.Model
        let context =
            buildWorkbookContext [| model |] [| sheet.Name |] [| true |] import.TableColumns Map.empty 0
            |> computeUncachedFormulas
        match Map.tryFind (sheetIdOfIndex 0) context.Sheets with
        | Some entry -> entry.Model
        | None -> model

// ---------------------------------------------------------------------------
// Export: models -> writer workbook
// ---------------------------------------------------------------------------

/// What `Export.toWorkbookOut` writes.
type ExportInput = {
    /// Every sheet's current model, in sheet order.
    Sheets: SpreadsheetModel array
    Names: string array
    /// Parallel to `Sheets`; a missing entry counts as visible.
    Visible: bool array
    ActiveSheetIndex: int
    /// Defined names as the workbook context keys them: (scope sheet id, UPPER name) -> refers-to.
    DefinedNames: Map<Wb.SheetId option * string, string>
    /// Spelling, comment and hidden flag of the defined names (Workbook.NameInfo);
    /// empty = as the file had them, upper-cased for names it did not have.
    NameInfo: Map<Wb.SheetId option * string, Wb.DefinedNameInfo>
    /// The workbook the sheets came from, when they came from a file.
    Source: XlsxSource option
    /// VBA module source to write as the workbook's VBA project (native hosts).
    Macros: string option
}

module Export =

    let private border (b: BorderStyle option) =
        b |> Option.map (fun side -> ({ Style = side.Style; Color = side.Color } : XlsxWriter.XlsxBorderSide))

    let private format (f: CellFormat) : XlsxWriter.XlsxCellFormat =
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

    let private richRun (text: string, r: RunOverride) : XlsxWriter.XlsxRichRun =
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
            | Some VertAlignSuperscript -> Some XlsxWriter.XlsxSuperscript
            | Some VertAlignSubscript -> Some XlsxWriter.XlsxSubscript
            | Some VertAlignBaseline -> Some XlsxWriter.XlsxBaseline
            | None -> None }

    /// A formula is written without its leading "=", with the value the editor holds for it.
    let rec private value (richRuns: XlsxWriter.XlsxRichRun list) (v: CellValue) : XlsxWriter.XlsxCellValue =
        match v with
        | CellValue.Empty -> XlsxWriter.XlsxBlank
        | CellValue.Number n -> XlsxWriter.XlsxNumber n
        | CellValue.String s -> XlsxWriter.XlsxText (s, richRuns)
        | CellValue.Boolean b -> XlsxWriter.XlsxBoolean b
        | CellValue.Error e -> XlsxWriter.XlsxError (if e.StartsWith "#" then e else "#" + e)
        | CellValue.Formula (expr, cached) ->
            let text = if expr.StartsWith "=" then expr.Substring 1 else expr
            match value [] cached with
            | XlsxWriter.XlsxBlank -> XlsxWriter.XlsxFormula (text, None)
            | result -> XlsxWriter.XlsxFormula (text, Some result)

    /// The sheet's validations, grouped back into ranges.
    let private validations (model: SpreadsheetModel) : XlsxWriter.XlsxDataValidation list =
        CF.validationsToOoxml model.Validations
        |> List.map (fun (sqref, dv) ->
            ({ Sqref = sqref
               ValidationType = dv.Type
               Operator = dv.Operator
               Formula1 = dv.Formula1
               Formula2 = dv.Formula2
               AllowBlank = dv.AllowBlank
               ShowDropDown = dv.ShowDropDown
               ErrorMessage = dv.Error
               InputMessage = dv.Prompt
               ErrorTitle = dv.ErrorTitle
               PromptTitle = dv.PromptTitle
               ErrorStyle = dv.ErrorStyle
               ShowErrorMessage = dv.ShowErrorMessage
               ShowInputMessage = dv.ShowInputMessage } : XlsxWriter.XlsxDataValidation))

    /// The sheet's conditional formats in the writer's shape. Each rule's
    /// style becomes a `<dxf>`; `dxfIndex` hands out its index in the
    /// workbook-wide table.
    let private conditionalFormats (dxfIndex: XlsxReader.Dxf -> int) (model: SpreadsheetModel) : XlsxReader.ConditionalFormat list =
        CF.toOoxml model.ConditionalFormats
        |> List.map (fun cf ->
            ({ Range = cf.Sqref
               Rules =
                cf.Rules
                |> List.map (fun rule ->
                    ({ Type = rule.Type
                       Operator = rule.Operator
                       Formulas = rule.Formulas
                       DxfId =
                        rule.Style |> Option.map (fun style ->
                            dxfIndex { FontColor = style.FontColor |> Option.map (fun c -> c.TrimStart '#')
                                       FillColor = style.FillColor |> Option.map (fun c -> c.TrimStart '#')
                                       Bold = (if style.Bold then Some true else None)
                                       Italic = (if style.Italic then Some true else None)
                                       Underline = (if style.Underline then Some true else None)
                                       Strike = (if style.Strikethrough then Some true else None) })
                       ColorScale = rule.ColorScale |> List.map (fun (kind, stopValue, color) -> ({ CfvoType = kind; CfvoValue = stopValue; Color = color } : XlsxReader.CfColorStop))
                       DataBarColor = rule.DataBarColor
                       DataBarCfvo = rule.DataBarCfvo
                       Text = rule.Text
                       IconSet = rule.IconSet
                       IconSetCfvo = rule.IconSetCfvo
                       IconSetReverse = rule.IconSetReverse
                       TimePeriod = rule.TimePeriod
                       Rank = rule.Rank
                       Percent = rule.Percent
                       Bottom = rule.Bottom
                       AboveAverage = rule.AboveAverage
                       EqualAverage = rule.EqualAverage
                       StdDev = rule.StdDev
                       Priority = rule.Priority
                       StopIfTrue = rule.StopIfTrue } : XlsxReader.ConditionalFormatRule)) } : XlsxReader.ConditionalFormat))

    let private textAt row col (model: SpreadsheetModel) =
        let addr : CellAddress = { Row = row; Col = col }
        match Map.tryFind addr model.Cells with
        | Some data when displayValue data.Value <> "" -> displayValue data.Value
        | _ -> sprintf "Column%d" (col + 1)

    let private table (model: SpreadsheetModel) (t: Table) : XlsxWriter.XlsxTableOut =
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

    let private margins (m: PageMargins) : XlsxWriter.XlsxPageMargins =
        { Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer }

    let private pageSetup (paperSize: int option) (p: PageSetup) : XlsxWriter.XlsxPageSetup =
        { Orientation = Some (match p.Orientation with LandscapeOrientation -> "landscape" | PortraitOrientation -> "portrait")
          FitToPage = p.FitToPage
          FitToWidth = p.FitToWidth
          FitToHeight = p.FitToHeight
          Scale = p.Scale
          PaperSize = paperSize
          FirstPageNumber = p.FirstPageNumber
          UseFirstPageNumber = p.UseFirstPageNumber
          PageOrder = Some (match p.PageOrder with OverThenDown -> "overThenDown" | DownThenOver -> "downThenOver")
          BlackAndWhite = p.BlackAndWhite
          CellComments =
            match p.CellComments with
            | CellCommentsAsDisplayed -> Some "asDisplayed"
            | CellCommentsAtEnd -> Some "atEnd"
            | CellCommentsNone -> None }

    let private printOptions (p: PrintOptions) : XlsxWriter.XlsxPrintOptions =
        { GridLines = p.GridLines; Headings = p.Headings; HorizontalCentered = p.HorizontalCentered; VerticalCentered = p.VerticalCentered }

    let private headerFooter (h: HeaderFooter) : XlsxWriter.XlsxHeaderFooter =
        { OddHeader = h.OddHeader; OddFooter = h.OddFooter; EvenHeader = h.EvenHeader; EvenFooter = h.EvenFooter
          FirstHeader = h.FirstHeader; FirstFooter = h.FirstFooter; DifferentOddEven = h.DifferentOddEven
          DifferentFirst = h.DifferentFirst; ScaleWithDoc = h.ScaleWithDoc }

    let private rangeRef (r: CellRange) =
        sprintf "$%s$%d:$%s$%d" (columnToLetter r.From.Col) (r.From.Row + 1) (columnToLetter r.To.Col) (r.To.Row + 1)

    /// Writer formats by the model format they were made from (by identity:
    /// the cells of a sheet share a few format records), so that the writer
    /// in turn finds the style of a cell by identity.
    type private FormatCache = Collections.Generic.Dictionary<CellFormat, XlsxWriter.XlsxCellFormat>

    let private newFormatCache () : FormatCache = Collections.Generic.Dictionary<CellFormat, XlsxWriter.XlsxCellFormat>(HashIdentity.Reference)

    let private cachedFormat (cache: FormatCache) (f: CellFormat) : XlsxWriter.XlsxCellFormat =
        match cache.TryGetValue f with
        | true, converted -> converted
        | false, _ ->
            let converted = format f
            cache.[f] <- converted
            converted

    let private cellOut (cache: FormatCache) (model: SpreadsheetModel) (addr: CellAddress) (data: CellData) : XlsxWriter.XlsxCellOut =
        let runs =
            if model.RichRuns.IsEmpty then []
            else model.RichRuns |> Map.tryFind addr |> Option.map (List.map richRun) |> Option.defaultValue []
        { Row = addr.Row
          Col = addr.Col
          Value = value runs data.Value
          Format = cachedFormat cache data.Format
          Comment = data.Comment |> Option.map (fun c -> c.Text, c.Author) }

    let private cellRangeText (r: CellRange) =
        if r.From = r.To then formatCellRef r.From.Row r.From.Col
        else formatCellRef r.From.Row r.From.Col + ":" + formatCellRef r.To.Row r.To.Col

    let private hyperlinks (model: SpreadsheetModel) : XlsxWriter.XlsxHyperlink list =
        model.Hyperlinks
        |> List.map (fun link -> ({ Range = cellRangeText link.Range; Target = link.Target; IsInternal = link.IsInternal } : XlsxWriter.XlsxHyperlink))

    let private dataTables (model: SpreadsheetModel) : XlsxReader.DataTableFormula list =
        model.DataTables
        |> List.map (fun table ->
            ({ Ref = cellRangeText table.Range
               TwoDimensional = table.TwoDimensional
               RowInput = table.RowInput
               Input1 = table.Input1
               Input2 = table.Input2
               Input1Deleted = table.Input1Deleted
               Input2Deleted = table.Input2Deleted
               CalculateAlways = table.CalculateAlways } : XlsxReader.DataTableFormula))

    /// A sheet written from its model alone. `cells` are its cells in the
    /// writer's shape (a sheet from a file supplies its own, see `sourcedSheet`).
    let private modelSheetWith (cells: XlsxWriter.XlsxCellOut list) (dxfIndex: XlsxReader.Dxf -> int) (name: string) (state: XlsxReader.SheetState) (model: SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        { Name = name
          State = state
          Cells = cells
          MergedCells = model.MergedCells |> List.map (fun m -> ({ StartRow = m.StartRow; StartCol = m.StartCol; EndRow = m.EndRow; EndCol = m.EndCol } : XlsxWriter.MergedCellRange))
          ColumnWidths = model.ColumnWidths |> Map.toList |> List.map (fun (c, px) -> c, px / 7.0)
          RowHeights = model.RowHeights |> Map.toList |> List.map (fun (r, px) -> r, px / 1.333)
          FrozenRows = model.FrozenRows
          FrozenCols = model.FrozenCols
          ShowGridLines = model.ShowGridLines
          Tables = model.Tables |> List.map (table model)
          DataValidations = validations model
          ConditionalFormats = conditionalFormats dxfIndex model
          Hyperlinks = hyperlinks model
          PageMargins = Some (margins model.PageSetup.Margins)
          PageSetup = Some (pageSetup None model.PageSetup)
          PrintOptions = Some (printOptions model.PageSetup.Options)
          HeaderFooter = Some (headerFooter model.PageSetup.HeaderFooter)
          RowBreaks = model.PageSetup.ManualRowBreaks
          ColBreaks = model.PageSetup.ManualColBreaks
          PrintArea = model.PageSetup.PrintArea |> Option.map rangeRef
          PrintTitleRows = model.PageSetup.RepeatRows |> Option.map (fun (a, b) -> sprintf "$%d:$%d" (a + 1) (b + 1))
          PrintTitleCols = model.PageSetup.RepeatCols |> Option.map (fun (a, b) -> sprintf "$%s:$%s" (columnToLetter a) (columnToLetter b))
          // Pictures and threads of the model. Kept objects need the parts of
          // the file they came from, so a sheet without a source has none.
          Extras =
            { extrasOfModel XlsxReader.emptySheetExtras None
                { model with Drawings = model.Drawings |> List.filter (fun drawing -> match drawing.Content with Picture _ -> true | KeptObject _ -> false) }
                with DataTables = dataTables model } }

    let private modelSheet (cache: FormatCache) (dxfIndex: XlsxReader.Dxf -> int) (name: string) (state: XlsxReader.SheetState) (model: SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        let cells = model.Cells |> Map.toList |> List.map (fun (addr, data) -> cellOut cache model addr data)
        modelSheetWith cells dxfIndex name state model

    let private sameRef (a: 'T) (b: 'T) = obj.ReferenceEquals(a, b)

    /// The same formula, its value now an error where the file had a result:
    /// a function this engine lacks. The file's value stands.
    let private onlyBecameError (before: CellData) (after: CellData) =
        match before.Value, after.Value with
        | CellValue.Formula (formulaBefore, cachedBefore), CellValue.Formula (formulaAfter, CellValue.Error _) ->
            formulaBefore = formulaAfter
            && (match cachedBefore with CellValue.Error _ -> false | _ -> true)
            && before.Format = after.Format
            && before.Comment = after.Comment
        | _ -> false

    /// Sizes for the file: what the file had, with every index whose size the
    /// user changed since the import replaced by the model's size.
    let private patchSizes (original: (int * float) list) (before: Map<int, float>) (after: Map<int, float>) (pixelsPerUnit: float) (defaultPixels: float) =
        if sameRef before after then original
        else
            let mutable sizes = Map.ofList original
            for KeyValue(index, pixels) in after do
                match Map.tryFind index before with
                | Some previous when previous = pixels -> ()
                | _ -> sizes <- Map.add index (pixels / pixelsPerUnit) sizes
            for KeyValue(index, _) in before do
                if not (Map.containsKey index after) then
                    // Back to the model's default size.
                    sizes <- Map.add index (defaultPixels / pixelsPerUnit) sizes
            Map.toList sizes

    /// A sheet that came from a file: the model's current state, with
    /// everything the user left alone written as the file had it.
    let private sourcedSheet (cache: FormatCache) (dxfIndex: XlsxReader.Dxf -> int) (name: string) (state: XlsxReader.SheetState)
                             (original: XlsxWriter.XlsxSheetOut) (readerSheet: XlsxReader.Worksheet)
                             (before: SpreadsheetModel) (model: SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        let cells =
                // Cells by position, as one number (16384 columns to a row).
                let key (row: int) (col: int) = float row * 16384.0 + float col
                let originals = Collections.Generic.Dictionary<float, XlsxWriter.XlsxCellOut>()
                for cell in original.Cells do originals.[key cell.Row cell.Col] <- cell
                // The cells as imported. With the cell map itself unchanged every cell is.
                let sameCells = sameRef before.Cells model.Cells
                let previousCells = Collections.Generic.Dictionary<float, CellData>()
                if not sameCells then
                    for KeyValue(addr, data) in before.Cells do previousCells.[key addr.Row addr.Col] <- data
                let sameRuns = sameRef before.RichRuns model.RichRuns
                model.Cells
                |> Map.toList
                |> List.map (fun (addr, data) ->
                    let position = key addr.Row addr.Col
                    let previous =
                        if sameCells then Some data
                        else
                            match previousCells.TryGetValue position with
                            | true, found -> Some found
                            | false, _ -> None
                    let untouched =
                        match previous with
                        | Some previous ->
                            (sameRef previous data || previous = data || onlyBecameError previous data)
                            && (sameRuns || Map.tryFind addr before.RichRuns = Map.tryFind addr model.RichRuns)
                        | None -> false
                    match untouched, originals.TryGetValue position with
                    | true, (true, kept) -> kept
                    | false, (true, kept) when (match previous with Some previous -> sameRef previous.Format data.Format || previous.Format = data.Format | None -> false) ->
                        // Only the value (or the comment) changed: the cell keeps
                        // the format the file gave it, its named cell style included.
                        { cellOut cache model addr data with Format = kept.Format }
                    | _ -> cellOut cache model addr data)
        let fromModel = modelSheetWith cells dxfIndex name state model
        let unchanged (pick: SpreadsheetModel -> 'T) = sameRef (pick before) (pick model) || pick before = pick model
        { fromModel with
            ColumnWidths = patchSizes original.ColumnWidths before.ColumnWidths model.ColumnWidths 7.0 100.0
            RowHeights = patchSizes original.RowHeights before.RowHeights model.RowHeights 1.333 25.0
            Tables = if unchanged (fun m -> m.Tables) then original.Tables else fromModel.Tables
            // A list validation whose values come from a range is expanded on
            // import; unchanged validations go back as the file states them.
            DataValidations = if sameRef before.Validations model.Validations then original.DataValidations else fromModel.DataValidations
            ConditionalFormats = if sameRef before.ConditionalFormats model.ConditionalFormats then original.ConditionalFormats else fromModel.ConditionalFormats
            // Hyperlinks are written where their cells are now.
            Hyperlinks = if unchanged (fun m -> m.Hyperlinks) then original.Hyperlinks else fromModel.Hyperlinks
            PageMargins = if before.PageSetup.Margins = model.PageSetup.Margins then original.PageMargins else fromModel.PageMargins
            PageSetup =
                if unchanged (fun m -> m.PageSetup) then original.PageSetup
                else Some (pageSetup (readerSheet.PageSetup |> Option.bind (fun p -> p.PaperSize)) model.PageSetup)
            PrintOptions = if before.PageSetup.Options = model.PageSetup.Options then original.PrintOptions else fromModel.PrintOptions
            HeaderFooter = if before.PageSetup.HeaderFooter = model.PageSetup.HeaderFooter then original.HeaderFooter else fromModel.HeaderFooter
            PrintArea = if before.PageSetup.PrintArea = model.PageSetup.PrintArea then original.PrintArea else fromModel.PrintArea
            PrintTitleRows = if before.PageSetup.RepeatRows = model.PageSetup.RepeatRows then original.PrintTitleRows else fromModel.PrintTitleRows
            PrintTitleCols = if before.PageSetup.RepeatCols = model.PageSetup.RepeatCols then original.PrintTitleCols else fromModel.PrintTitleCols
            Extras =
                // Drawing objects and threaded comments are written from the
                // model: where the user's row and column edits moved them.
                let extras = extrasOfModel original.Extras (Some before) model
                let extras =
                    if unchanged (fun m -> m.DataTables) then extras
                    else { extras with DataTables = dataTables model }
                // Sparklines and x14 conditional formats name cells; the
                // editor does not model them, so they only go back while the
                // sheet's rows and columns are where the file had them.
                let extras =
                    if before.StructureVersion = model.StructureVersion then extras
                    else { extras with KeptCellExtensions = [] }
                // Column default styles follow their columns. The file's own
                // `<col>` list is by position: it is only kept while no
                // column was inserted or deleted.
                if before.ColCount = model.ColCount && unchanged (fun m -> m.ColumnStyles) then extras
                else
                    let fileStyles = List.toArray original.Extras.ColumnStyles
                    { extras with
                        Columns = []
                        ColumnStyles =
                            model.ColumnStyles
                            |> List.choose (fun style ->
                                if style.Source >= 0 && style.Source < fileStyles.Length then
                                    let _, _, cellStyle = fileStyles.[style.Source]
                                    Some (style.First, style.Last, cellStyle)
                                else None) } }

    /// The defined names to write. Names that still read as imported are
    /// written as the file had them (their spelling and order); otherwise
    /// from the workbook context, with the file's spelling where it has one.
    let private definedNamesOut (input: ExportInput) : XlsxWriter.XlsxDefinedName list =
        let sheetIndexOfId (id: Wb.SheetId) =
            let mutable found = None
            for i in 0 .. input.Sheets.Length - 1 do
                if found.IsNone && sheetIdOfIndex i = id then found <- Some i
            found
        let fromSource (source: XlsxSource) =
            source.Workbook.NamedRanges
            |> List.map (fun n ->
                ({ Name = n.Name
                   RefersTo = n.RefersTo
                   LocalSheetIndex = n.Scope |> Option.bind (fun s -> match Int32.TryParse s with | true, v -> Some v | _ -> None)
                   Comment = n.Comment
                   Hidden = n.Hidden } : XlsxWriter.XlsxDefinedName))
        match input.Source with
        | Some source when (sameRef source.BaselineNames input.DefinedNames || source.BaselineNames = input.DefinedNames)
                           && (input.NameInfo.IsEmpty || sameRef source.BaselineNameInfo input.NameInfo || source.BaselineNameInfo = input.NameInfo) -> fromSource source
        | _ ->
            let spelling =
                match input.Source with
                | Some source -> source.Workbook.NamedRanges |> List.map (fun n -> n.Name.ToUpperInvariant(), n.Name) |> Map.ofList
                | None -> Map.empty
            // Print areas and titles are written from each sheet's page setup.
            let isPrintName (name: string) = name.StartsWith "_XLNM.PRINT_"
            input.DefinedNames
            |> Map.toList
            |> List.choose (fun ((scope, upperName), refersTo) ->
                if isPrintName upperName then None
                else
                    let local = scope |> Option.bind sheetIndexOfId
                    if scope.IsSome && local.IsNone then None
                    else
                        let info = Map.tryFind (scope, upperName) input.NameInfo
                        Some ({ Name =
                                    (match info with
                                     | Some info -> info.Name
                                     | None -> Map.tryFind upperName spelling |> Option.defaultValue upperName)
                                RefersTo = (if refersTo.StartsWith "=" then refersTo.Substring 1 else refersTo)
                                LocalSheetIndex = local
                                Comment = info |> Option.bind (fun info -> info.Comment)
                                Hidden = (match info with Some info -> info.Hidden | None -> false) } : XlsxWriter.XlsxDefinedName))

    /// The editor's workbook in the writer's shape (see the module comment).
    let toWorkbookOut (input: ExportInput) : XlsxWriter.XlsxWorkbookOut =
        let original = input.Source |> Option.map (fun source -> XlsxWriter.packageFromReaderWorkbook source.Workbook)
        // Rules kept from the file refer to the file's dxfs by index, so those
        // keep their places; styles of new rules are added behind them.
        let dxfs = ResizeArray<XlsxReader.Dxf>()
        match original with
        | Some package -> dxfs.AddRange package.Dxfs
        | None -> ()
        let dxfIndex (dxf: XlsxReader.Dxf) =
            match dxfs.IndexOf dxf with
            | -1 -> dxfs.Add dxf; dxfs.Count - 1
            | index -> index
        let cache = newFormatCache ()
        let nameOf index = if index < input.Names.Length then input.Names.[index] else sprintf "Sheet%d" (index + 1)
        let visibleAt index = index >= input.Visible.Length || input.Visible.[index]
        // A sheet of the source is found by the identity its model carries
        // (`SourceSheet`: which sheet of the file it was imported as), so the
        // carry-over survives renames, reordering and added or removed
        // sheets. Two models claiming the same sheet (a copy): the first is it.
        let sourceOf : int -> (XlsxReader.Worksheet * XlsxWriter.XlsxSheetOut * SpreadsheetModel) option =
            match input.Source, original with
            | Some source, Some package ->
                let readerSheets = List.toArray source.Workbook.Sheets
                let outSheets = List.toArray package.Sheets
                let claimed = Collections.Generic.HashSet<int>()
                let owners =
                    input.Sheets
                    |> Array.map (fun model ->
                        match model.SourceSheet with
                        | Some origin when origin.Book = source.Id && origin.Sheet >= 0 && origin.Sheet < readerSheets.Length
                                           && origin.Sheet < outSheets.Length && origin.Sheet < source.Baseline.Length
                                           && claimed.Add origin.Sheet ->
                            Some (readerSheets.[origin.Sheet], outSheets.[origin.Sheet], source.Baseline.[origin.Sheet])
                        | _ -> None)
                fun index -> if index < owners.Length then owners.[index] else None
            | _ -> fun _ -> None
        let sheets =
            input.Sheets
            |> Array.mapi (fun index model ->
                let name = nameOf index
                match sourceOf index with
                | Some (readerSheet, outSheet, before) ->
                    // Hidden and very hidden are both "no tab" to the editor.
                    let state =
                        match visibleAt index, readerSheet.State with
                        | true, _ -> XlsxReader.SheetVisible
                        | false, XlsxReader.SheetVisible -> XlsxReader.SheetHidden
                        | false, hidden -> hidden
                    // A sheet that was never converted goes back as the file had it.
                    if isDeferredSheet model then { outSheet with Name = name; State = state }
                    else sourcedSheet cache dxfIndex name state outSheet readerSheet before model
                | None ->
                    modelSheet cache dxfIndex name (if visibleAt index then XlsxReader.SheetVisible else XlsxReader.SheetHidden) model)
            |> Array.toList
        { Sheets = sheets
          ActiveSheetIndex = max 0 (min (input.Sheets.Length - 1) input.ActiveSheetIndex)
          DefinedNames = definedNamesOut input
          VbaProjectSource = input.Macros |> Option.bind (fun source -> if String.IsNullOrWhiteSpace source then None else Some source)
          Dxfs = List.ofSeq dxfs
          Preserved =
            match input.Source with
            | Some source ->
                // A pivot table whose cells moved (rows / columns inserted or
                // deleted before it) is written at its new place: only the
                // range of its <location> changes, the part is otherwise kept.
                let moved =
                    [ for index in 0 .. input.Sheets.Length - 1 do
                        match sourceOf index with
                        | Some (_, _, before) when not (sameRef before.KeptRanges input.Sheets.[index].KeptRanges) ->
                            for kept in input.Sheets.[index].KeptRanges do
                                match before.KeptRanges |> List.tryFind (fun was -> was.Key = kept.Key) with
                                | Some was when was.Range <> kept.Range -> yield kept.Key, cellRangeText kept.Range
                                | _ -> ()
                        | _ -> () ]
                    |> Map.ofList
                // The calculation chain lists the formula cells of the file: it
                // is kept only while every sheet of the file is there with its
                // formulas in the same cells (Excel builds a new chain for a
                // file without one; a stale one makes it report a damaged file).
                let formulaCells (cells: Map<CellAddress, CellData>) =
                    let found = ResizeArray<CellAddress>()
                    for KeyValue(addr, data) in cells do
                        match data.Value with
                        | CellValue.Formula _ -> found.Add addr
                        | _ -> ()
                    found
                let chainStands =
                    source.Workbook.Preserved.CalcChain.IsSome
                    && input.Sheets.Length = source.Baseline.Length
                    && (input.Sheets
                        |> Array.mapi (fun index model -> index, model)
                        |> Array.forall (fun (index, model) ->
                            match sourceOf index with
                            | Some (_, _, before) ->
                                sameRef before.Cells model.Cells
                                || (before.StructureVersion = model.StructureVersion
                                    && (let was = formulaCells before.Cells
                                        let now = formulaCells model.Cells
                                        was.Count = now.Count && Seq.forall2 (=) was now))
                            | None -> false))
                let preserved =
                    if chainStands then source.Workbook.Preserved else { source.Workbook.Preserved with CalcChain = None }
                if moved.IsEmpty then preserved
                else
                    let cachePaths = preserved.PivotCacheSources |> List.map (fun (path, _, _) -> path) |> Set.ofList
                    { preserved with
                        Parts =
                            preserved.Parts
                            |> List.map (fun (path, data) ->
                                match Map.tryFind path moved with
                                | Some range ->
                                    let text = Text.Encoding.UTF8.GetString data
                                    // A pivot cache's source range, or a pivot table's location.
                                    let rewritten =
                                        if Set.contains path cachePaths then XlsxReader.withPivotCacheSource range text
                                        else XlsxReader.withPivotLocation range text
                                    path, Text.Encoding.UTF8.GetBytes rewritten
                                | None -> path, data) }
            | None -> XlsxReader.emptyWorkbookPreservation }
