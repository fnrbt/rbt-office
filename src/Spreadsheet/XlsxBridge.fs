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
    /// The file is in the 1904 date system: its serials are kept as they are,
    /// and `Model.date1904` (set by the import) makes dates of them.
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
      NameInfo = Map.empty
      Iteration = None }

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
        ({ Id = 0
           Anchor = anchorOfReader image.Anchor
           Name = image.Name
           Description = image.Description
           Content = Picture (image.ContentType, image.Extension, image.Data) } : SheetDrawing)))
    @ (extras.Objects
       |> List.map (fun item ->
        item.Anchor.Order, 1,
        ({ Id = 0
           Anchor = anchorOfReader item.Anchor
           Name = item.Name
           Description = None
           Content = KeptObject (item.Kind, item.Xml) } : SheetDrawing)))
    |> List.sortBy (fun (order, kind, _) -> order, kind)
    |> List.mapi (fun index (_, _, drawing) -> { drawing with Id = index + 1 })

let rec private threadOfReader (comment: XlsxReader.ThreadedComment) : CommentThread =
    { Id = comment.Id
      Author = comment.Author
      PersonId = comment.PersonId
      Date = comment.Date
      Text = comment.Text
      Replies = comment.Replies |> List.map threadOfReader
      Done = comment.Done }

let rec private threadToReader (addr: CellAddress) (thread: CommentThread) : XlsxReader.ThreadedComment =
    { CellRef = { Column = columnToLetter addr.Col; Row = addr.Row + 1 }
      Id = thread.Id
      Author = thread.Author
      PersonId = thread.PersonId
      Date = thread.Date
      Text = thread.Text
      Replies = thread.Replies |> List.map (threadToReader addr)
      Done = thread.Done }

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
    // What rows / columns have besides a size: the size a hidden one keeps
    // (points / characters in the file), outline levels, collapsed marks.
    let axesUnchanged =
        match before with
        | Some baseline ->
            (obj.ReferenceEquals(baseline.Axes, model.Axes) || baseline.Axes = model.Axes)
            && obj.ReferenceEquals(baseline.RowHeights, model.RowHeights) && obj.ReferenceEquals(baseline.ColumnWidths, model.ColumnWidths)
        | None -> false
    let axisOf (sizes: Map<int, float>) (kept: Map<int, float>) (levels: Map<int, int>) (collapsed: Set<int>) (pixelsPerUnit: float) =
        let indices = Collections.Generic.HashSet<int>()
        for KeyValue(index, size) in sizes do if size <= 0.0 then indices.Add index |> ignore
        for KeyValue(index, _) in levels do indices.Add index |> ignore
        // (A group at the first row with its summary above has its mark at -1: not a row.)
        for index in collapsed do if index >= 0 then indices.Add index |> ignore
        [ for index in Seq.sort indices ->
            let hidden = match Map.tryFind index sizes with Some size -> size <= 0.0 | None -> false
            index,
            (if hidden then Map.tryFind index kept |> Option.map (fun pixels -> pixels / pixelsPerUnit) else None),
            (Map.tryFind index levels |> Option.defaultValue 0),
            Set.contains index collapsed ]
    let rowAxis, columnAxis =
        if axesUnchanged then original.RowAxis, original.ColumnAxis
        else
            axisOf model.RowHeights model.Axes.HiddenRowSizes model.Axes.RowLevels model.Axes.CollapsedRows 1.333,
            axisOf model.ColumnWidths model.Axes.HiddenColSizes model.Axes.ColLevels model.Axes.CollapsedCols 7.0
    // A formula that spilled is written as a dynamic array formula over the
    // cells it fills: opened again (here or elsewhere), those cells are its
    // results, not cells in its way. A formula the file had as a legacy
    // (Ctrl+Shift+Enter) array formula stays one, over the range it had; a
    // dynamic array that fills one cell stays a dynamic array.
    let arrayFormulas =
        if Map.isEmpty model.Spill && Map.isEmpty model.ArrayMarks then []
        else
            let bounds = System.Collections.Generic.Dictionary<CellAddress, CellAddress>()
            for KeyValue (cell, anchor) in model.Spill do
                if cell.Row >= anchor.Row && cell.Col >= anchor.Col then
                    match bounds.TryGetValue anchor with
                    | true, corner -> bounds.[anchor] <- { Row = max corner.Row cell.Row; Col = max corner.Col cell.Col }
                    | false, _ -> bounds.[anchor] <- cell
            let rangeText (anchor: CellAddress) (corner: CellAddress) =
                let first = formatCellRef anchor.Row anchor.Col
                if corner = anchor then first else first + ":" + formatCellRef corner.Row corner.Col
            let marked = System.Collections.Generic.HashSet<CellAddress>()
            [ for KeyValue (anchor, _) in model.ArrayMarks do
                match arrayMarkOf anchor model with
                | Some mark ->
                    marked.Add anchor |> ignore
                    let spilled = match bounds.TryGetValue anchor with | true, corner -> Some corner | _ -> None
                    let corner =
                        match spilled with
                        | Some corner when not mark.Legacy -> corner
                        | _ -> { Row = anchor.Row + max 1 mark.Rows - 1; Col = anchor.Col + max 1 mark.Cols - 1 }
                    // A dynamic array that no longer spills fills its own cell.
                    let corner = if mark.Legacy || spilled.IsSome then corner else anchor
                    yield formatCellRef anchor.Row anchor.Col, rangeText anchor corner, not mark.Legacy
                | None -> ()
              for KeyValue (anchor, corner) in bounds do
                if not (marked.Contains anchor) then
                    match Map.tryFind anchor model.Cells with
                    | Some { Value = CellValue.Formula _ } ->
                        yield formatCellRef anchor.Row anchor.Col, rangeText anchor corner, true
                    | _ -> () ]
    // The notes whose text is in runs (a thread's note is made from the thread).
    let commentRuns =
        match before with
        | Some baseline when obj.ReferenceEquals(baseline.Cells, model.Cells) -> original.CommentRuns
        | _ ->
            let props (run: RunOverride) : XlsxReader.RunProps =
                { Bold = run.Bold = Some true
                  Italic = run.Italic = Some true
                  Underline = run.Underline = Some true
                  UnderlineDouble = run.UnderlineDouble = Some true
                  Strikethrough = run.Strikethrough = Some true
                  Color = run.Color |> Option.map (fun color -> color.TrimStart('#'))
                  FontName = run.FontName
                  FontSize = run.FontSize
                  Scheme = ""
                  VertAlign =
                    match run.VertAlign with
                    | Some VertAlignSuperscript -> XlsxReader.VertAlignSuperscript
                    | Some VertAlignSubscript -> XlsxReader.VertAlignSubscript
                    | _ -> XlsxReader.VertAlignBaseline }
            [ for KeyValue (addr, data) in model.Cells do
                match data.Comment with
                | Some note when not note.Runs.IsEmpty && not (Map.containsKey addr model.Threads) ->
                    yield (addr.Row, addr.Col), note.Runs |> List.map (fun (text, run) -> text, props run)
                | _ -> () ]
    { original with
        Images = images; Objects = objects; ThreadedComments = threads; RowAxis = rowAxis; ColumnAxis = columnAxis; ArrayFormulas = arrayFormulas
        CommentRuns = commentRuns
        // (The file's drawing part goes back as it is only with the file's pictures and objects.)
        DrawingPart = if drawingsUnchanged then original.DrawingPart else None }

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
                        // An x14 rule has its format in place.
                        |> Option.orElse (rule.Ext |> Option.bind (fun ext -> ext.Dxf))
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
                       Ext =
                        rule.Ext |> Option.map (fun ext ->
                            ({ DataBar =
                                ext.DataBar |> Option.map (fun bar ->
                                    ({ MinLength = bar.MinLength; MaxLength = bar.MaxLength; ShowValue = bar.ShowValue; Border = bar.Border
                                       Gradient = bar.Gradient; Direction = bar.Direction
                                       NegativeBarColorSameAsPositive = bar.NegativeBarColorSameAsPositive
                                       NegativeBarBorderColorSameAsPositive = bar.NegativeBarBorderColorSameAsPositive
                                       AxisPosition = bar.AxisPosition; Cfvo = bar.Cfvo; BorderColor = bar.BorderColor
                                       NegativeFillColor = bar.NegativeFillColor; NegativeBorderColor = bar.NegativeBorderColor
                                       AxisColor = bar.AxisColor } : CF.OoxmlDataBarExt))
                               IconCustom = ext.IconCustom
                               ShowValue = ext.ShowValue
                               InExtList = ext.InExtList } : CF.OoxmlRuleExt))
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
                    PaperSize = p.PaperSize
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
        let printAreas : CellRange list =
            match sheet.PrintArea with
            | Some text ->
                text.Split(',') |> Array.toList
                |> List.choose (fun part -> wholeSheet (part.Trim()))
                |> List.map (fun (sr, sc, er, ec) -> ({ From = { Row = sr; Col = sc }; To = { Row = er; Col = ec } } : CellRange))
            | None -> []
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
            // (A print area of several areas names them with commas between.)
            PrintArea = printAreas |> List.tryHead
            ExtraPrintAreas = (match printAreas with _ :: rest -> rest | [] -> [])
            RepeatRows = sheet.PrintTitleRows |> Option.bind wholeSheet |> Option.map (fun (sr, _, er, _) -> sr, er)
            RepeatCols = sheet.PrintTitleCols |> Option.bind wholeSheet |> Option.map (fun (_, sc, _, ec) -> sc, ec) }

    /// The sheet's AutoFilter in the model: its range and column filters.
    /// The rows the file hides below the filter's header row are the rows
    /// the filter hid: they become the model's filtered rows (and are rows
    /// of their own height again, not rows the user hid).
    let private withAutoFilter (wb: XlsxReader.Workbook) (sheet: XlsxReader.Worksheet) (model: SpreadsheetModel) : SpreadsheetModel =
        match sheet.Extras.AutoFilter |> Option.bind (fun filter -> CF.parseArea filter.Ref |> Option.map (fun range -> filter, range)) with
        | None -> model
        | Some (filter, range) ->
            let criteria (column: XlsxReader.AutoFilterColumn) : FilterCriteria =
                let kept = FilterKept column.Xml
                match column.Kind with
                | "filters" -> FilterValues (Set.ofList column.Values, column.Blank)
                | "customFilters" ->
                    let conditions =
                        column.Custom
                        |> List.map (fun (operator, value) ->
                            match operator with
                            | "equal" | "" -> Some (OpEqual, value)
                            | "notEqual" -> Some (OpNotEqual, value)
                            | "greaterThan" -> Some (OpGreaterThan, value)
                            | "greaterThanOrEqual" -> Some (OpGreaterOrEqual, value)
                            | "lessThan" -> Some (OpLessThan, value)
                            | "lessThanOrEqual" -> Some (OpLessOrEqual, value)
                            | _ -> None)
                    if conditions.IsEmpty || conditions |> List.exists Option.isNone then kept
                    else FilterCustom (conditions |> List.choose id, column.And || conditions.Length = 1)
                | "top10" -> FilterTop (max 1 (int column.Value), column.Percent, not column.Top)
                | "colorFilter" ->
                    let dxf = column.DxfId |> Option.bind (fun id -> if id >= 0 && id < wb.Dxfs.Length then Some wb.Dxfs.[id] else None)
                    (match dxf with
                     | Some dxf ->
                        let color = (if column.CellColor then dxf.FillColor else dxf.FontColor) |> Option.bind hexColor |> Option.defaultValue ""
                        FilterColor (color.ToUpperInvariant(), not column.CellColor)
                     | None -> kept)
                | "dynamicFilter" when column.DynamicType = "aboveAverage" || column.DynamicType = "belowAverage" -> FilterDynamic column.DynamicType
                | _ -> kept
            let filters =
                filter.Columns
                |> List.fold (fun acc column ->
                    let col = range.From.Col + column.ColId
                    Map.add col ({ Column = col; Criteria = criteria column } : ColumnFilter) acc) Map.empty
            let filtered =
                if Map.isEmpty filters then [||]
                else
                    [| for KeyValue(row, _) in model.Axes.HiddenRowSizes do
                        if row > range.From.Row && row <= range.To.Row then yield row |]
            let shown =
                if filtered.Length = 0 then model
                else
                    let mutable heights = model.RowHeights
                    let mutable kept = model.Axes.HiddenRowSizes
                    for row in filtered do
                        (match Map.tryFind row kept with
                         | Some size -> heights <- Map.add row size heights
                         | None -> heights <- Map.remove row heights)
                        kept <- Map.remove row kept
                    { model with RowHeights = heights; Axes = { model.Axes with HiddenRowSizes = kept } }
            { shown with
                AutoFilter = Some { Range = range; Implicit = false }
                Filters = filters
                FilterHiddenRows = Array.sort filtered }

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
                    comments.[(r, col)] <-
                        { Text = c.Text
                          Author = c.Author
                          // (One run without a format of its own is a plain text.)
                          Runs =
                            (let runs = c.Runs |> List.map (fun (text, props) -> text, runOverride props)
                             if runs |> List.exists (fun (_, run) -> run <> emptyRunOverride) && (runs |> List.map fst |> String.concat "") = c.Text then runs else []) }
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
        // The cells that hold formulas: the sheet's formula index (`Model.formulaIndexOf`).
        let formulaCells = ResizeArray<CellAddress>()
        let mutable richRuns = Map.empty<CellAddress, (string * RunOverride) list>
        // The other cells of a legacy array formula (<f t="array" ref="F94:H94">)
        // hold results the anchor's formula spilled there: they belong to the
        // anchor, so a recalculation spills over them again instead of finding
        // them in its way (#SPILL!).
        let mutable spill = Map.empty<CellAddress, CellAddress>
        // What the file says of each array formula (legacy or dynamic, its range).
        let mutable arrayMarks = Map.empty<CellAddress, ArrayFormulaMark>
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
                    (match value, cell.ArrayRef with
                     | CellValue.Formula (text, _), Some range ->
                        let rows, cols =
                            match parsePlainRange range with
                            | Some (sr, sc, er, ec) -> er - sr + 1, ec - sc + 1
                            | None -> 1, 1
                        arrayMarks <- Map.add { Row = r; Col = c } { Text = text; Legacy = not cell.ArrayDynamic; Rows = rows; Cols = cols } arrayMarks
                     | _ -> ())
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
                        (match value with CellValue.Formula _ -> formulaCells.Add addr | _ -> ())
                        let runs =
                            if not (List.isEmpty cell.InlineRuns) then cell.InlineRuns
                            else
                                match cell.Value with
                                | XlsxReader.SharedString idx when idx >= 0 && idx < wb.SharedStringRuns.Length -> wb.SharedStringRuns.[idx]
                                | _ -> []
                        match runs with
                        | _ :: _ :: _ -> richRuns <- Map.add addr (runs |> List.map (fun (text, props) -> text, runOverride props)) richRuns
                        | _ -> ()

        // A note on a cell the sheet data has no element for (a cell without
        // a value or a style) is on the cell all the same.
        for KeyValue((r, c), comment) in comments do
            let addr : CellAddress = { Row = r; Col = c }
            if r < 1048576 && c < 16384 && not (Map.containsKey addr cellMap) then
                cellMap <- Map.add addr { Value = CellValue.Empty; Format = { defaultFormat with HasComment = true }; Comment = Some comment } cellMap
                if r + 1 > rowCount then rowCount <- r + 1
                if c + 1 > colCount then colCount <- c + 1

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
        // A size of 0 is how the model marks a hidden row / column; the size
        // it is shown at again, and the outline, are kept beside the sizes.
        let mutable hiddenRowSizes = Map.empty<int, float>
        let mutable hiddenColSizes = Map.empty<int, float>
        let mutable rowLevels = Map.empty<int, int>
        let mutable colLevels = Map.empty<int, int>
        let mutable collapsedRows = Set.empty<int>
        let mutable collapsedCols = Set.empty<int>
        for row in sheet.Rows do
            let index = row.Index - 1
            if row.Hidden then
                hiddenRowSizes <- Map.add index ((row.Height |> Option.defaultValue defaultRowHeight) * 1.333) hiddenRowSizes
                rowHeights <- Map.add index 0.0 rowHeights
            if row.OutlineLevel > 0 then rowLevels <- Map.add index (min 7 row.OutlineLevel) rowLevels
            if row.Collapsed then collapsedRows <- Set.add index collapsedRows
        for column in sheet.Columns do
            // A range can run to the last column (16384); only the ones near the data matter.
            if column.Hidden || column.OutlineLevel > 0 || column.Collapsed then
                for idx in column.Index .. min column.MaxIndex (column.Index + 1023) do
                    if column.Hidden then
                        hiddenColSizes <- Map.add (idx - 1) ((column.Width |> Option.defaultValue defaultColWidth) * 7.0) hiddenColSizes
                        columnWidths <- Map.add (idx - 1) 0.0 columnWidths
                    if column.OutlineLevel > 0 then colLevels <- Map.add (idx - 1) (min 7 column.OutlineLevel) colLevels
                    if column.Collapsed then collapsedCols <- Set.add (idx - 1) collapsedCols
        let axes : SheetAxes =
            if hiddenRowSizes.IsEmpty && hiddenColSizes.IsEmpty && rowLevels.IsEmpty && colLevels.IsEmpty && collapsedRows.IsEmpty && collapsedCols.IsEmpty
               && sheet.View.SummaryBelow && sheet.View.SummaryRight && sheet.View.ShowOutlineSymbols then emptySheetAxes
            else
                { HiddenRowSizes = hiddenRowSizes; HiddenColSizes = hiddenColSizes; RowLevels = rowLevels; ColLevels = colLevels
                  CollapsedRows = collapsedRows; CollapsedCols = collapsedCols
                  SummaryBelow = sheet.View.SummaryBelow; SummaryRight = sheet.View.SummaryRight; ShowSymbols = sheet.View.ShowOutlineSymbols }

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
                       IsInternal = link.IsInternal
                       Tooltip = link.Tooltip
                       Display = link.Display } : SheetHyperlink)))
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

        // (A cell the file lists twice is the later one: only what the map holds counts.)
        registerFormulaIndex cellMap
            (formulaCells
             |> Seq.filter (fun addr -> match Map.tryFind addr cellMap with Some { Value = CellValue.Formula _ } -> true | _ -> false)
             |> Set.ofSeq)
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
                Axes = axes
                Tables = sheetTables
                FrozenRows = sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0
                FrozenCols = sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0
                ShowGridLines = sheet.ShowGridLines
                ShowHeaders = sheet.View.ShowRowColHeaders
                ShowFormulas = sheet.View.ShowFormulas
                ShowZeros = sheet.ShowZeros
                Zoom = float sheet.View.ZoomScale / 100.0
                TabColor = sheet.View.TabColor
                VeryHidden = (sheet.State = XlsxReader.SheetVeryHidden)
                Validations = rules
                ConditionalFormats = conditionalFormats wb sheet
                PageSetup = pageSetup sheet
                Drawings = drawingsOfExtras sheet.Extras
                Threads = threads
                Spill = spill
                ArrayMarks = arrayMarks }
            |> withAutoFilter wb sheet
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
        // (Only a formula can lack its result: the sheets' formula indexes say where to look.)
        let needsWork =
            context.Sheets
            |> Map.exists (fun _ sheet ->
                let cells = sheet.Model.Cells
                formulaIndexOf cells |> Set.exists (fun addr -> match Map.tryFind addr cells with Some cell -> isUncached cell | None -> false))
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
    /// The file's iterative calculation of circular references (`<calcPr iterate>`).
    let private iterationOf (wb: XlsxReader.Workbook) : Wb.IterativeCalculation option =
        wb.Preserved.CalcIteration |> Option.map (fun (count, delta) -> { Wb.IterativeCalculation.MaxIterations = count; MaxChange = delta })

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
        // (Dates are made of the workbook's serials in its own date system from here on.)
        Rbt.Spreadsheet.Model.date1904 <- wb.Date1904
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
            { buildWorkbookContext models names visible tableColumns names' firstVisible with NameInfo = definedNameInfo wb; Iteration = iterationOf wb }
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

    /// The formula cells of a reader sheet, as `importSheet` makes them (see `cellValue`).
    let private readerSheetFormulas (sheet: XlsxReader.Worksheet) : (CellAddress * string) array =
        let formulas = ResizeArray<CellAddress * string>()
        let add (cell: XlsxReader.Cell) (formula: string) =
            let r = cell.Reference.Row - 1
            let c = columnIndexOfLetters cell.Reference.Column
            if r >= 0 && c >= 0 then formulas.Add ({ Row = r; Col = c }, formula)
        // (A sheet of several hundred thousand values and a handful of formulas:
        // the reader finds the cells with a formula without reading the others.)
        for cell in sheet.FormulaCells do
            match cell.Formula with
            | Some formula -> if not (String.IsNullOrWhiteSpace formula) then add cell formula
            | None ->
                match cell.Value with
                | XlsxReader.Formula formula when not (String.IsNullOrWhiteSpace formula) -> add cell formula
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
        // (Dates are made of the workbook's serials in its own date system from here on.)
        Rbt.Spreadsheet.Model.date1904 <- wb.Date1904
        let options = defaultImportOptions
        let readerSheets = List.toArray wb.Sheets
        if readerSheets.Length = 0 then failwith "XLSX workbook has no worksheets"
        let anyUncached =
            readerSheets |> Array.exists (fun sheet -> sheet.FormulaCells |> List.exists readerCellIsUncached)
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
                    // (What names other sheets and so has to follow them: formulas, and links into the workbook.)
                    let hasFormulas =
                        not sheet.ConditionalFormats.IsEmpty || not sheet.DataValidations.IsEmpty
                        || sheet.Hyperlinks |> List.exists (fun link -> link.IsInternal)
                        || not sheet.FormulaCells.IsEmpty
                    // (Read off the file once: the graph of a workbook is built more than once.)
                    let mutable formulasRead : (CellAddress * string) array option = None
                    let formulasOf () =
                        match formulasRead with
                        | Some known -> known
                        | None ->
                            let read = readerSheetFormulas sheet
                            formulasRead <- Some read
                            read
                    standIns.[index] <-
                        { createEmpty options.MinRows options.MinCols with
                            Tables = sheetTables
                            // (What can change without the sheet being converted.)
                            TabColor = sheet.View.TabColor
                            VeryHidden = (sheet.State = XlsxReader.SheetVeryHidden)
                            SourceSheet = Some { Book = sourceId; Sheet = index }
                            Deferred =
                                Some { Loaded = None
                                       Load = (fun () -> convert index)
                                       HasFormulas = hasFormulas
                                       Formulas = formulasOf } }
                    baseline.[index] <- standIns.[index]
                    columns)
                |> Array.fold (fun acc columns -> Map.fold (fun m k v -> Map.add k v m) acc columns) Map.empty
            let visible = readerSheets |> Array.map (fun sheet -> sheetStateText sheet.State = "visible")
            let firstVisible = visible |> Array.tryFindIndex id |> Option.defaultValue 0
            let models = standIns |> Array.mapi (fun index standIn -> if convertNow index then convert index else standIn)
            let context = { buildWorkbookContext models names visible tableColumns names' firstVisible with NameInfo = definedNameInfo wb; Iteration = iterationOf wb }
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
                    let dxf =
                        rule.Style |> Option.map (fun style ->
                            ({ FontColor = style.FontColor |> Option.map (fun c -> c.TrimStart '#')
                               FillColor = style.FillColor |> Option.map (fun c -> c.TrimStart '#')
                               Bold = (if style.Bold then Some true else None)
                               Italic = (if style.Italic then Some true else None)
                               Underline = (if style.Underline then Some true else None)
                               Strike = (if style.Strikethrough then Some true else None) } : XlsxReader.Dxf))
                    // A rule of the x14 list has its format in place, not in the styles part.
                    let inExtList = rule.Ext |> Option.exists (fun ext -> ext.InExtList)
                    ({ Type = rule.Type
                       Operator = rule.Operator
                       Formulas = rule.Formulas
                       DxfId = (if inExtList then None else dxf |> Option.map dxfIndex)
                       Ext =
                        rule.Ext |> Option.map (fun ext ->
                            ({ DataBar =
                                ext.DataBar |> Option.map (fun bar ->
                                    ({ MinLength = bar.MinLength; MaxLength = bar.MaxLength; ShowValue = bar.ShowValue; Border = bar.Border
                                       Gradient = bar.Gradient; Direction = bar.Direction
                                       NegativeBarColorSameAsPositive = bar.NegativeBarColorSameAsPositive
                                       NegativeBarBorderColorSameAsPositive = bar.NegativeBarBorderColorSameAsPositive
                                       AxisPosition = bar.AxisPosition; Cfvo = bar.Cfvo; BorderColor = bar.BorderColor
                                       NegativeFillColor = bar.NegativeFillColor; NegativeBorderColor = bar.NegativeBorderColor
                                       AxisColor = bar.AxisColor } : XlsxReader.CfDataBarExt))
                               IconCustom = ext.IconCustom
                               ShowValue = ext.ShowValue
                               InExtList = ext.InExtList
                               Dxf = (if ext.InExtList then dxf else None) } : XlsxReader.CfRuleExt))
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
          Columns =
            (if t.HeaderRowCount > 0 then [ for c in t.StartCol - 1 .. t.EndCol - 1 -> textAt (t.StartRow - 1) c model ]
             else [ for index in 1 .. t.EndCol - t.StartCol + 1 -> sprintf "Column%d" index ])
          HeaderRowCount = t.HeaderRowCount
          TotalsRowCount = if t.HasTotals then 1 else 0
          Totals =
            (if not t.HasTotals then []
             else
                [ for c in t.StartCol - 1 .. t.EndCol - 1 ->
                    match Map.tryFind { Row = t.EndRow - 1; Col = c } model.Cells |> Option.map (fun data -> data.Value) with
                    | Some (CellValue.String label) when label <> "" -> "label", label
                    | Some (CellValue.Formula (formula, _)) ->
                        // SUBTOTAL(10N, the column) is one of Excel's own totals; anything else is a custom one.
                        let body = (if formula.StartsWith "=" then formula.Substring 1 else formula).Trim()
                        let known =
                            [ "101", "average"; "102", "countNums"; "103", "count"; "104", "max"; "105", "min"; "107", "stdDev"; "109", "sum"; "110", "var" ]
                            |> List.tryPick (fun (code, name) -> if body.StartsWith("SUBTOTAL(" + code + ",", StringComparison.OrdinalIgnoreCase) then Some name else None)
                        (match known with
                         | Some name -> "function", name
                         | None -> "custom", body)
                    | _ -> "", "" ])
          AutoFilter = t.AutoFilter
          Style =
            t.Style |> Option.map (fun s ->
                ({ Name = s.Name
                   ShowFirstColumn = s.ShowFirstColumn
                   ShowLastColumn = s.ShowLastColumn
                   ShowRowStripes = s.ShowRowStripes
                   ShowColumnStripes = s.ShowColumnStripes } : XlsxWriter.XlsxTableStyle))
          SourcePart = None }

    let private margins (m: PageMargins) : XlsxWriter.XlsxPageMargins =
        { Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer }

    let private pageSetup (p: PageSetup) : XlsxWriter.XlsxPageSetup =
        { Orientation = Some (match p.Orientation with LandscapeOrientation -> "landscape" | PortraitOrientation -> "portrait")
          FitToPage = p.FitToPage
          FitToWidth = p.FitToWidth
          FitToHeight = p.FitToHeight
          Scale = p.Scale
          PaperSize = p.PaperSize
          FirstPageNumber = p.FirstPageNumber
          UseFirstPageNumber = p.UseFirstPageNumber
          PageOrder = Some (match p.PageOrder with OverThenDown -> "overThenDown" | DownThenOver -> "downThenOver")
          BlackAndWhite = p.BlackAndWhite
          CellComments =
            match p.CellComments with
            | CellCommentsAsDisplayed -> Some "asDisplayed"
            | CellCommentsAtEnd -> Some "atEnd"
            | CellCommentsNone -> None
          Kept = [] }

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
          Comment =
            // A thread has a note beside it in the file (what applications
            // without threads show): written from the thread as it is now.
            match (if model.Threads.IsEmpty then None else Map.tryFind addr model.Threads) with
            | Some thread -> let note = Rbt.Spreadsheet.SheetOps.threadPlaceholderNote thread in Some (note.Text, note.Author)
            | None -> data.Comment |> Option.map (fun c -> c.Text, c.Author) }

    let private cellRangeText (r: CellRange) =
        if r.From = r.To then formatCellRef r.From.Row r.From.Col
        else formatCellRef r.From.Row r.From.Col + ":" + formatCellRef r.To.Row r.To.Col

    let private hyperlinks (model: SpreadsheetModel) : XlsxWriter.XlsxHyperlink list =
        model.Hyperlinks
        |> List.map (fun link -> ({ Range = cellRangeText link.Range; Target = link.Target; IsInternal = link.IsInternal; Tooltip = link.Tooltip; Display = link.Display } : XlsxWriter.XlsxHyperlink))

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
    /// The model as a file states it: the rows its AutoFilter hides are
    /// hidden rows that keep their height.
    let private filterRowsAsHidden (model: SpreadsheetModel) : SpreadsheetModel =
        if model.FilterHiddenRows.Length = 0 || Map.isEmpty model.Filters then model
        else
            let mutable heights = model.RowHeights
            let mutable kept = model.Axes.HiddenRowSizes
            for row in model.FilterHiddenRows do
                match Map.tryFind row model.RowHeights with
                | Some size when size <= 0.0 -> ()
                | size ->
                    kept <- Map.add row (defaultArg size 25.0) kept
                    heights <- Map.add row 0.0 heights
            { model with RowHeights = heights; Axes = { model.Axes with HiddenRowSizes = kept } }

    /// The sheet's AutoFilter in the reader's shape. A colour filter's
    /// colour becomes a `<dxf>` (`dxfIndex` hands out its index).
    let private autoFilterOut (dxfIndex: XlsxReader.Dxf -> int) (model: SpreadsheetModel) : XlsxReader.SheetAutoFilter option =
        model.AutoFilter
        |> Option.map (fun filter ->
            let column (col: int, columnFilter: ColumnFilter) : XlsxReader.AutoFilterColumn =
                let empty = { XlsxReader.emptyAutoFilterColumn with ColId = col - filter.Range.From.Col }
                match columnFilter.Criteria with
                | FilterValues (values, blanks) -> { empty with Kind = "filters"; Values = Set.toList values; Blank = blanks }
                | FilterCustom (conditions, matchAll) ->
                    { empty with
                        Kind = "customFilters"
                        And = matchAll
                        Custom =
                            conditions
                            |> List.map (fun (op, value) ->
                                (match op with
                                 | OpNotEqual -> "notEqual"
                                 | OpGreaterThan -> "greaterThan"
                                 | OpGreaterOrEqual -> "greaterThanOrEqual"
                                 | OpLessThan -> "lessThan"
                                 | OpLessOrEqual -> "lessThanOrEqual"
                                 | _ -> "equal"), value) }
                | FilterTop (count, percent, bottom) -> { empty with Kind = "top10"; Top = not bottom; Percent = percent; Value = float count }
                | FilterColor (color, font) ->
                    let hex = if color = "" then None else Some (color.TrimStart '#')
                    let dxf : XlsxReader.Dxf =
                        { FontColor = (if font then hex else None); FillColor = (if font then None else hex)
                          Bold = None; Italic = None; Underline = None; Strike = None }
                    { empty with Kind = "colorFilter"; DxfId = Some (dxfIndex dxf); CellColor = not font }
                | FilterDynamic kind -> { empty with Kind = "dynamicFilter"; DynamicType = kind }
                | FilterKept xml -> { empty with Kind = "other"; Xml = xml }
            { Ref = CF.sqrefText [ filter.Range ]
              Columns = model.Filters |> Map.toList |> List.filter (fun (col, _) -> col >= filter.Range.From.Col) |> List.map column })

    let private modelSheetWith (cells: XlsxWriter.XlsxCellOut list) (dxfIndex: XlsxReader.Dxf -> int) (name: string) (state: XlsxReader.SheetState) (model: SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        { Name = name
          State = state
          CellSource = XlsxWriter.sheetCells cells
          MergedCells = model.MergedCells |> List.map (fun m -> ({ StartRow = m.StartRow; StartCol = m.StartCol; EndRow = m.EndRow; EndCol = m.EndCol } : XlsxWriter.MergedCellRange))
          ColumnWidths = model.ColumnWidths |> Map.toList |> List.map (fun (c, px) -> c, px / 7.0)
          RowHeights = model.RowHeights |> Map.toList |> List.map (fun (r, px) -> r, px / 1.333)
          FrozenRows = model.FrozenRows
          FrozenCols = model.FrozenCols
          ShowGridLines = model.ShowGridLines
          View =
            { ShowZeros = model.ShowZeros
              ShowHeaders = model.ShowHeaders
              ShowFormulas = model.ShowFormulas
              ZoomScale = max 10 (min 400 (int (System.Math.Round (model.Zoom * 100.0))))
              TabColor = model.TabColor
              SummaryBelow = model.Axes.SummaryBelow
              SummaryRight = model.Axes.SummaryRight
              ShowOutlineSymbols = model.Axes.ShowSymbols }
          Tables = model.Tables |> List.map (table model)
          DataValidations = validations model
          ConditionalFormats = conditionalFormats dxfIndex model
          Hyperlinks = hyperlinks model
          PageMargins = Some (margins model.PageSetup.Margins)
          PageSetup = Some (pageSetup model.PageSetup)
          PrintOptions = Some (printOptions model.PageSetup.Options)
          HeaderFooter = Some (headerFooter model.PageSetup.HeaderFooter)
          RowBreaks = model.PageSetup.ManualRowBreaks
          ColBreaks = model.PageSetup.ManualColBreaks
          PrintArea = (match printAreasOf model.PageSetup with [] -> None | areas -> Some (areas |> List.map rangeRef |> String.concat ","))
          PrintTitleRows = model.PageSetup.RepeatRows |> Option.map (fun (a, b) -> sprintf "$%d:$%d" (a + 1) (b + 1))
          PrintTitleCols = model.PageSetup.RepeatCols |> Option.map (fun (a, b) -> sprintf "$%s:$%s" (columnToLetter a) (columnToLetter b))
          // Pictures and threads of the model. Kept objects need the parts of
          // the file they came from, so a sheet without a source has none.
          Extras =
            { extrasOfModel XlsxReader.emptySheetExtras None
                { model with Drawings = model.Drawings |> List.filter (fun drawing -> match drawing.Content with Picture _ -> true | KeptObject _ -> false) }
                with DataTables = dataTables model; AutoFilter = autoFilterOut dxfIndex model } }

    let private modelSheet (cache: FormatCache) (dxfIndex: XlsxReader.Dxf -> int) (name: string) (state: XlsxReader.SheetState) (model: SpreadsheetModel) : XlsxWriter.XlsxSheetOut =
        let cells = model.Cells |> Map.toList |> List.map (fun (addr, data) -> cellOut cache model addr data)
        modelSheetWith cells dxfIndex name state model

    let private sameRef (a: 'T) (b: 'T) = obj.ReferenceEquals(a, b)

    let private xmlTextDecoded (text: string) =
        if text.IndexOf '&' < 0 then text
        else text.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&")
    let private xmlTextEncoded (text: string) =
        text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")

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
                let basedOnFormats = Collections.Generic.Dictionary<int, Collections.Generic.Dictionary<XlsxWriter.XlsxCellFormat, XlsxWriter.XlsxCellFormat>>()
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
                    | _, (true, kept) when kept.Format.SourceXf.IsSome ->
                        // The format was changed: the new one is made from the
                        // file's (what the change did not touch of it stays).
                        let written = cellOut cache model addr data
                        let basedOn = kept.Format.SourceXf.Value
                        let known =
                            match basedOnFormats.TryGetValue basedOn with
                            | true, known -> known
                            | false, _ ->
                                let known = Collections.Generic.Dictionary<XlsxWriter.XlsxCellFormat, XlsxWriter.XlsxCellFormat>(HashIdentity.Reference)
                                basedOnFormats.[basedOn] <- known
                                known
                        let format =
                            match known.TryGetValue written.Format with
                            | true, format -> format
                            | false, _ ->
                                let format = { written.Format with BasedOnXf = Some basedOn }
                                known.[written.Format] <- format
                                format
                        { written with Format = format }
                    | _ -> cellOut cache model addr data)
        let fromModel = modelSheetWith cells dxfIndex name state model
        let unchanged (pick: SpreadsheetModel -> 'T) = sameRef (pick before) (pick model) || pick before = pick model
        let movedSizes (onRows: bool) (fileSizes: (int * float) list) (sizesBefore: Map<int, float>) (sizesNow: Map<int, float>) (pixelsPerUnit: float) (defaultPixels: float) (modelSizes: (int * float) list) =
            if before.StructureVersion = model.StructureVersion then patchSizes fileSizes sizesBefore sizesNow pixelsPerUnit defaultPixels
            elif model.StructureLog.Length <> model.StructureVersion - before.StructureVersion || before.StructureVersion <> 0 then modelSizes
            else
                let moves = List.rev model.StructureLog |> List.filter (fun move -> move.OnRows = onRows)
                if moves.IsEmpty then patchSizes fileSizes sizesBefore sizesNow pixelsPerUnit defaultPixels
                else
                    let moved (index: int) : int option =
                        moves
                        |> List.fold (fun (at: int option) move ->
                            match at with
                            | None -> None
                            | Some index ->
                                if move.Insert then Some (if index >= move.At then index + move.Count else index)
                                elif index >= move.At + move.Count then Some (index - move.Count)
                                elif index >= move.At then None
                                else Some index) (Some index)
                    let fileMoved = fileSizes |> List.choose (fun (index, size) -> moved index |> Option.map (fun index -> index, size))
                    let beforeMoved = sizesBefore |> Map.toList |> List.choose (fun (index, size) -> moved index |> Option.map (fun index -> index, size)) |> Map.ofList
                    patchSizes fileMoved beforeMoved sizesNow pixelsPerUnit defaultPixels
        { fromModel with
            // (The file's sizes are by index: with rows or columns inserted or
            // deleted, each index has the size the model gives it.)
            // (The file's sizes are by index: with rows or columns inserted or
            // deleted, by where each row and column of the file is now.)
            ColumnWidths = movedSizes false original.ColumnWidths before.ColumnWidths model.ColumnWidths 7.0 100.0 fromModel.ColumnWidths
            RowHeights = movedSizes true original.RowHeights before.RowHeights model.RowHeights 1.333 25.0 fromModel.RowHeights
            Tables =
                if unchanged (fun m -> m.Tables) then original.Tables
                else
                    // Each table that is one of the file's is written as the
                    // file's part of it: found by its name, or - renamed - as
                    // the table of the file it overlaps that no table is named for.
                    let named = fromModel.Tables |> List.map (fun t -> t.Name) |> Set.ofList
                    let orphans = original.Tables |> List.filter (fun t -> t.SourcePart.IsSome && not (Set.contains t.Name named))
                    let overlaps (a: XlsxWriter.XlsxTableOut) (b: XlsxWriter.XlsxTableOut) =
                        a.StartRow <= b.EndRow && b.StartRow <= a.EndRow && a.StartCol <= b.EndCol && b.StartCol <= a.EndCol
                    fromModel.Tables
                    |> List.map (fun t ->
                        match original.Tables |> List.tryFind (fun known -> known.Name = t.Name) with
                        | Some known -> { t with SourcePart = known.SourcePart }
                        | None ->
                            match orphans |> List.filter (overlaps t) with
                            | [ known ] -> { t with SourcePart = known.SourcePart }
                            | _ -> t)
            // A list validation whose values come from a range is expanded on
            // import; unchanged validations go back as the file states them.
            DataValidations = if sameRef before.Validations model.Validations then original.DataValidations else fromModel.DataValidations
            ConditionalFormats = if sameRef before.ConditionalFormats model.ConditionalFormats then original.ConditionalFormats else fromModel.ConditionalFormats
            // Hyperlinks are written where their cells are now.
            Hyperlinks = if unchanged (fun m -> m.Hyperlinks) then original.Hyperlinks else fromModel.Hyperlinks
            PageMargins = if before.PageSetup.Margins = model.PageSetup.Margins then original.PageMargins else fromModel.PageMargins
            PageSetup =
                if unchanged (fun m -> m.PageSetup) then original.PageSetup
                else
                    // What the file says about the printer (resolution, draft,
                    // copies ...) stays through an edit of the setup; a custom
                    // paper's dimensions go when the paper was changed.
                    let kept =
                        match original.PageSetup with
                        | Some known ->
                            if before.PageSetup.PaperSize = model.PageSetup.PaperSize then known.Kept
                            else known.Kept |> List.filter (fun (name, _) -> name <> "paperHeight" && name <> "paperWidth")
                        | None -> []
                    Some { pageSetup model.PageSetup with Kept = kept }
            PrintOptions = if before.PageSetup.Options = model.PageSetup.Options then original.PrintOptions else fromModel.PrintOptions
            HeaderFooter = if before.PageSetup.HeaderFooter = model.PageSetup.HeaderFooter then original.HeaderFooter else fromModel.HeaderFooter
            PrintArea = if printAreasOf before.PageSetup = printAreasOf model.PageSetup then original.PrintArea else fromModel.PrintArea
            PrintTitleRows = if before.PageSetup.RepeatRows = model.PageSetup.RepeatRows then original.PrintTitleRows else fromModel.PrintTitleRows
            PrintTitleCols = if before.PageSetup.RepeatCols = model.PageSetup.RepeatCols then original.PrintTitleCols else fromModel.PrintTitleCols
            Extras =
                // Drawing objects and threaded comments are written from the
                // model: where the user's row and column edits moved them.
                let extras = extrasOfModel original.Extras (Some before) model
                let extras =
                    if unchanged (fun m -> m.DataTables) then extras
                    else { extras with DataTables = dataTables model }
                // A table part of the file goes back as it is while the table
                // is what the part says: its range, its columns, its rows.
                let extras =
                    if extras.TableParts.IsEmpty then extras
                    else
                        let same (name: string) =
                            match original.Tables |> List.tryFind (fun t -> t.Name = name), fromModel.Tables |> List.tryFind (fun t -> t.Name = name) with
                            | Some a, Some b ->
                                a.StartRow = b.StartRow && a.StartCol = b.StartCol && a.EndRow = b.EndRow && a.EndCol = b.EndCol && a.Columns = b.Columns
                                && a.HeaderRowCount = b.HeaderRowCount && a.TotalsRowCount = b.TotalsRowCount && a.Style = b.Style && a.AutoFilter.IsSome = b.AutoFilter.IsSome
                            | _ -> false
                        let kept = extras.TableParts |> List.filter (fun (name, _) -> same name)
                        if kept.Length = extras.TableParts.Length then extras else { extras with TableParts = kept }
                // The file's VML drawing of the comment boxes stands only while
                // the comments are the file's: with other comments it is made anew.
                let commentsOf (m: SpreadsheetModel) = m.Cells |> Map.toList |> List.choose (fun (addr, data) -> data.Comment |> Option.map (fun comment -> addr, comment))
                let commentsSame = sameRef before.Cells model.Cells || commentsOf before = commentsOf model
                // ... and so does its comments part (rich text, shape ids).
                let extras = if commentsSame then extras else { extras with CommentsPart = None }
                let extras =
                    if not (extras.KeptElements |> List.exists (fun (elementName, _) -> elementName = "legacyDrawingNotes")) then extras
                    else
                        if commentsSame then extras
                        else
                            let dropped = extras.KeptElements |> List.filter (fun (elementName, _) -> elementName = "legacyDrawingNotes") |> List.collect (fun (_, xml) -> DocxReader.relationshipIdsInXml xml) |> Set.ofList
                            { extras with
                                KeptElements = extras.KeptElements |> List.filter (fun (elementName, _) -> elementName <> "legacyDrawingNotes")
                                KeptRelationships = extras.KeptRelationships |> List.filter (fun (id, _, _, _) -> not (Set.contains id dropped)) }
                // An AutoFilter that was not touched goes back as the file states it.
                let extras =
                    if before.AutoFilter = model.AutoFilter && unchanged (fun m -> m.Filters) then extras
                    else { extras with AutoFilter = fromModel.Extras.AutoFilter }
                // Sparklines and x14 conditional formats name cells; the
                // editor does not model them, so they only go back while the
                // sheet's rows and columns are where the file had them.
                // With rows or columns inserted or deleted, the cells and ranges they
                // name (`xm:sqref`, `xm:f`) follow, a sparkline whose cell or data is
                // gone is taken out, and any other extension that lost a cell is dropped.
                let extras =
                    if before.StructureVersion = model.StructureVersion || extras.KeptCellExtensions.IsEmpty then extras
                    elif model.StructureLog.Length <> model.StructureVersion || before.StructureVersion <> 0 
                         // (No formula engine at hand to move a reference with.)
                         || adjustFormulaForEdit (InsertRows (0, 1)) (fun _ -> true) "=A1" = "=A1" then { extras with KeptCellExtensions = [] }
                    else
                        let moves = List.rev model.StructureLog
                        let movedSpan (onRows: bool) (low: int) (high: int) : (int * int) option =
                            moves
                            |> List.fold (fun (span: (int * int) option) move ->
                                match span with
                                | Some (low, high) when move.OnRows = onRows ->
                                    if move.Insert then
                                        if low >= move.At then Some (low + move.Count, high + move.Count)
                                        elif high >= move.At then Some (low, high + move.Count)
                                        else span
                                    else
                                        let stop = move.At + move.Count
                                        let newLow = if low >= stop then low - move.Count elif low >= move.At then move.At else low
                                        let newHigh = if high >= stop then high - move.Count elif high >= move.At then move.At - 1 else high
                                        if newHigh < newLow then None else Some (newLow, newHigh)
                                | _ -> span) (Some (low, high))
                        let movedRanges (sqref: string) : string option =
                            let parts =
                                sqref.Split(' ')
                                |> Array.filter (fun part -> part <> "")
                                |> Array.map (fun part ->
                                    parsePlainRange (part.Replace("$", ""))
                                    |> Option.bind (fun (sr, sc, er, ec) ->
                                        match movedSpan true sr er, movedSpan false sc ec with
                                        | Some (sr, er), Some (sc, ec) when er <= 1048575 && ec <= 16383 ->
                                            Some (if sr = er && sc = ec then formatCellRef sr sc else formatCellRef sr sc + ":" + formatCellRef er ec)
                                        | _ -> None))
                            if parts.Length = 0 || parts |> Array.exists Option.isNone then None else Some (parts |> Array.choose id |> String.concat " ")
                        let ownName = readerSheet.Name
                        let onThisSheet (sheet: string option) =
                            match sheet with
                            | None -> true
                            | Some other -> other.Trim('\'').Equals(ownName, StringComparison.OrdinalIgnoreCase)
                        let movedFormula (formula: string) : string option =
                            let adjusted =
                                moves
                                |> List.fold (fun (text: string) move ->
                                    let edit =
                                        match move.OnRows, move.Insert with
                                        | true, true -> InsertRows (move.At, move.Count)
                                        | true, false -> DeleteRows (move.At, move.Count)
                                        | false, true -> InsertCols (move.At, move.Count)
                                        | false, false -> DeleteCols (move.At, move.Count)
                                    adjustFormulaForEdit edit onThisSheet text) ("=" + xmlTextDecoded formula)
                            if adjusted.Contains "#REF!" then None else Some (xmlTextEncoded (if adjusted.StartsWith "=" then adjusted.Substring 1 else adjusted))
                        // The text with every `<name>...</name>` moved; None when one of them is gone.
                        let withMoved (name: string) (move: string -> string option) (xml: string) : string option =
                            let mutable lost = false
                            let moved =
                                Text.RegularExpressions.Regex.Replace(xml, "<" + name + ">([^<]*)</" + name + ">", (fun (found: Text.RegularExpressions.Match) ->
                                    match move found.Groups.[1].Value with
                                    | Some text -> "<" + name + ">" + text + "</" + name + ">"
                                    | None ->
                                        lost <- true
                                        found.Value))
                            if lost then None else Some moved
                        let movedText (xml: string) = withMoved "xm:sqref" movedRanges xml |> Option.bind (withMoved "xm:f" movedFormula)
                        let movedExtension (xml: string) : string option =
                            if not (xml.Contains "<x14:sparklineGroups") then movedText xml
                            else
                                // Sparkline by sparkline; a group, and the list, with none left goes.
                                let sparklines =
                                    Text.RegularExpressions.Regex.Replace(xml, "<x14:sparkline>[\\s\\S]*?</x14:sparkline>", (fun (found: Text.RegularExpressions.Match) ->
                                        movedText found.Value |> Option.defaultValue ""))
                                let groups =
                                    Text.RegularExpressions.Regex.Replace(sparklines, "<x14:sparklineGroup(?=[\\s>])[^>]*>[\\s\\S]*?</x14:sparklineGroup>", (fun (found: Text.RegularExpressions.Match) ->
                                        if not (found.Value.Contains "<x14:sparkline>") then ""
                                        else
                                            // (What the group itself names: the range of its dates.)
                                            let listAt = found.Value.IndexOf "<x14:sparklines>"
                                            if listAt < 0 then found.Value
                                            else
                                                match movedText (found.Value.Substring(0, listAt)) with
                                                | Some head -> head + found.Value.Substring listAt
                                                | None -> ""))
                                if groups.Contains "<x14:sparklineGroup" && (groups.Contains "<x14:sparklineGroup>" || groups.Contains "<x14:sparklineGroup ") then Some groups else None
                        { extras with KeptCellExtensions = extras.KeptCellExtensions |> List.choose (fun xml -> try movedExtension xml with _ -> None) }
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

    // ---- the sheet names inside kept parts ---------------------------------------------
    // Parts the editor keeps without modelling them name sheets too: the series of a
    // chart (`<c:f>Sales!$B$2:$B$9</c:f>`), the source of a pivot cache
    // (`<worksheetSource sheet="Sales">`), the formulas of kept extensions
    // (`<xm:f>`). They follow a sheet that was renamed, and say #REF! of one that
    // was deleted, as the formulas of the cells do.

    /// A sheet name as a reference writes it: quoted when it has to be.
    let private sheetNameInReference (name: string) : string =
        let plain =
            name.Length > 0 && (Char.IsLetter name.[0] || name.[0] = '_')
            && name |> Seq.forall (fun ch -> Char.IsLetterOrDigit ch || ch = '_' || ch = '.')
        let letters = name |> Seq.takeWhile Char.IsLetter |> Seq.length
        let likeCell =
            (letters >= 1 && letters <= 3 && letters < name.Length && name.Substring letters |> Seq.forall Char.IsDigit)
            || (name.Length >= 2 && "RrCc".IndexOf name.[0] >= 0 && name.Substring 1 |> Seq.forall (fun ch -> Char.IsDigit ch || ch = 'C' || ch = 'c'))
        let word = name.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || name.Equals("FALSE", StringComparison.OrdinalIgnoreCase)
        if plain && not likeCell && not word then name else "'" + name.Replace("'", "''") + "'"

    /// `formula` with the sheets its references and names are qualified by
    /// renamed: `renames` gives the new name by the old one in upper case.
    /// Text in double quotes is left alone, and so is a reference into
    /// another workbook (`[1]Sheet1!A1`).
    let private renameSheetQualifiers (renames: Collections.Generic.Dictionary<string, string>) (formula: string) : string =
        if renames.Count = 0 || formula.IndexOf '!' < 0 then formula
        else
            let sb = Text.StringBuilder()
            let renamed (name: string) =
                match renames.TryGetValue (name.ToUpperInvariant()) with
                | true, next -> next
                | false, _ -> name
            /// The qualifier as it is written for the names it holds (one, or the two ends of a 3-D reference).
            let written (qualifier: string) (wasQuoted: bool) : string =
                if qualifier.StartsWith "[" then (if wasQuoted then "'" + qualifier.Replace("'", "''") + "'" else qualifier)
                else
                    let cut = qualifier.IndexOf ':'
                    if cut < 0 then
                        let next = renamed qualifier
                        if next = qualifier then (if wasQuoted then "'" + qualifier.Replace("'", "''") + "'" else qualifier)
                        else sheetNameInReference next
                    else
                        let first, last = renamed (qualifier.Substring(0, cut)), renamed (qualifier.Substring(cut + 1))
                        let both = first + ":" + last
                        if first = qualifier.Substring(0, cut) && last = qualifier.Substring(cut + 1) then (if wasQuoted then "'" + qualifier.Replace("'", "''") + "'" else qualifier)
                        elif sheetNameInReference first = first && sheetNameInReference last = last then both
                        else "'" + both.Replace("'", "''") + "'"
            let isNameChar (ch: char) = Char.IsLetterOrDigit ch || ch = '_' || ch = '.'
            let mutable i = 0
            while i < formula.Length do
                let ch = formula.[i]
                if ch = '"' then
                    // A text: up to its closing quote ("" is a quote inside it).
                    let mutable j = i + 1
                    let mutable closed = false
                    while not closed && j < formula.Length do
                        if formula.[j] = '"' then
                            if j + 1 < formula.Length && formula.[j + 1] = '"' then j <- j + 2
                            else closed <- true
                        else j <- j + 1
                    let stop = min formula.Length (j + 1)
                    sb.Append(formula.Substring(i, stop - i)) |> ignore
                    i <- stop
                elif ch = '\'' then
                    let name = Text.StringBuilder()
                    let mutable j = i + 1
                    let mutable closed = false
                    while not closed && j < formula.Length do
                        if formula.[j] = '\'' then
                            if j + 1 < formula.Length && formula.[j + 1] = '\'' then
                                name.Append('\'') |> ignore
                                j <- j + 2
                            else closed <- true
                        else
                            name.Append(formula.[j]) |> ignore
                            j <- j + 1
                    let stop = min formula.Length (j + 1)
                    if closed && stop < formula.Length && formula.[stop] = '!' then sb.Append(written (name.ToString()) true) |> ignore
                    else sb.Append(formula.Substring(i, stop - i)) |> ignore
                    i <- stop
                elif Char.IsLetter ch || ch = '_' then
                    let mutable j = i
                    while j < formula.Length && isNameChar formula.[j] do j <- j + 1
                    // (`Jan:Mar!B2`: the two ends of a 3-D reference.)
                    let mutable stop = j
                    if j < formula.Length && formula.[j] = ':' then
                        let mutable k = j + 1
                        while k < formula.Length && isNameChar formula.[k] do k <- k + 1
                        if k > j + 1 && k < formula.Length && formula.[k] = '!' then stop <- k
                    if stop < formula.Length && formula.[stop] = '!' then sb.Append(written (formula.Substring(i, stop - i)) false) |> ignore
                    else
                        stop <- j
                        sb.Append(formula.Substring(i, stop - i)) |> ignore
                    i <- stop
                else
                    sb.Append(ch) |> ignore
                    i <- i + 1
            sb.ToString()


    /// `xml` with the text of every element named one of `names` (as written,
    /// with its prefix) passed through `change`.
    let private mapElementTexts (names: string list) (change: string -> string) (xml: string) : string =
        let mutable result = xml
        for name in names do
            let opening = "<" + name
            let closing = "</" + name + ">"
            if result.IndexOf opening >= 0 then
                let sb = Text.StringBuilder()
                let mutable at = 0
                let mutable going = true
                while going do
                    let found = result.IndexOf (opening, at)
                    if found < 0 then going <- false
                    else
                        let after = found + opening.Length
                        let tagEnd = if after < result.Length then result.IndexOf ('>', after) else -1
                        // (`<c:f>` and `<cx:f dir="row">`, not `<c:fmtId>`.)
                        let isIt = after < result.Length && (result.[after] = '>' || result.[after] = ' ') && tagEnd >= 0 && result.[tagEnd - 1] <> '/'
                        let close = if isIt then result.IndexOf (closing, tagEnd) else -1
                        if close < 0 then
                            sb.Append(result.Substring(at, after - at)) |> ignore
                            at <- after
                        else
                            let inner = result.Substring(tagEnd + 1, close - tagEnd - 1)
                            sb.Append(result.Substring(at, tagEnd + 1 - at)) |> ignore
                            if inner.IndexOf '<' >= 0 then sb.Append(inner) |> ignore
                            else
                                let text = xmlTextDecoded inner
                                let changed = change text
                                sb.Append(if changed = text then inner else xmlTextEncoded changed) |> ignore
                            at <- close
                sb.Append(result.Substring at) |> ignore
                result <- sb.ToString()
        result

    /// The formulas of a kept part or extension (`names`: the elements that
    /// hold one) for the sheets as they are named now: `renames` by the old
    /// name in upper case, `deleted` the sheets that are gone (`tabs`: the
    /// file's sheets in tab order, for the 3-D references among them).
    let private keptFormulasFollow (names: string list) (renames: Collections.Generic.Dictionary<string, string>) (deleted: string list) (tabs: string list) (xml: string) : string =
        if renames.Count = 0 && deleted.IsEmpty then xml
        else
            mapElementTexts names (fun formula ->
                let afterDeletes =
                    deleted
                    |> List.fold (fun (text: string) name ->
                        if text.IndexOf(name, StringComparison.OrdinalIgnoreCase) < 0 then text
                        else
                            let rewritten = rewriteFormulaForSheets (DeleteSheetRefs (name, tabs)) ("=" + text)
                            if rewritten.StartsWith "=" then rewritten.Substring 1 else rewritten) formula
                renameSheetQualifiers renames afterDeletes) xml

    /// A pivot cache definition with the sheet of its `<worksheetSource sheet>` renamed.
    let private pivotCacheSheetFollows (renames: Collections.Generic.Dictionary<string, string>) (xml: string) : string =
        let tagAt = xml.IndexOf "<worksheetSource"
        let tagEnd = if tagAt < 0 then -1 else xml.IndexOf ('>', tagAt)
        if tagEnd < 0 then xml
        else
            let tag = xml.Substring(tagAt, tagEnd - tagAt)
            let key = " sheet=\""
            let at = tag.IndexOf key
            if at < 0 then xml
            else
                let start = at + key.Length
                let close = tag.IndexOf ('"', start)
                if close < 0 then xml
                else
                    let name = xmlTextDecoded (tag.Substring(start, close - start))
                    match renames.TryGetValue (name.ToUpperInvariant()) with
                    | true, next -> xml.Substring(0, tagAt + start) + (xmlTextEncoded next).Replace("\"", "&quot;") + xml.Substring(tagAt + close)
                    | false, _ -> xml

    let private isChartPart (path: string) =
        path.StartsWith "xl/charts/" && path.EndsWith ".xml" && path.IndexOf "/_rels/" < 0
    /// The elements of a chart part (chart, chartEx) that hold a formula.
    let private chartFormulaElements = [ "c:f"; "cx:f"; "c15:f"; "c16r3:f" ]

    /// The editor's workbook in the writer's shape (see the module comment).
    /// Whether a sheet is what it was when it was opened, in everything a
    /// save writes into its part: its cells and their texts, sizes, merges,
    /// formats, rules, validations, tables, drawings, comments, links, page
    /// setup, panes, the view's switches and zoom, the tab colour.
    let private untouchedSheet (before: SpreadsheetModel) (model: SpreadsheetModel) : bool =
        let same (pick: SpreadsheetModel -> 'T) = sameRef (pick before) (pick model) || pick before = pick model
        sameRef before model
        || (sameRef before.Cells model.Cells
            && sameRef before.RichRuns model.RichRuns
            && same (fun m -> m.RowHeights) && same (fun m -> m.ColumnWidths) && same (fun m -> m.Axes)
            && before.RowCount = model.RowCount && before.ColCount = model.ColCount && before.StructureVersion = model.StructureVersion
            && same (fun m -> m.MergedCells) && same (fun m -> m.ConditionalFormats) && same (fun m -> m.Validations)
            && same (fun m -> m.Tables) && same (fun m -> m.Drawings) && same (fun m -> m.Threads) && same (fun m -> m.Hyperlinks)
            && same (fun m -> m.ColumnStyles) && same (fun m -> m.DataTables) && same (fun m -> m.KeptRanges) && same (fun m -> m.ArrayMarks)
            && same (fun m -> m.AutoFilter) && same (fun m -> m.Filters) && same (fun m -> m.SortState)
            && same (fun m -> m.FilterVisibleRows) && same (fun m -> m.FilterHiddenRows)
            && same (fun m -> m.PageSetup)
            && before.FrozenRows = model.FrozenRows && before.FrozenCols = model.FrozenCols
            && before.ShowGridLines = model.ShowGridLines && before.ShowHeaders = model.ShowHeaders && before.ShowFormulas = model.ShowFormulas
            && before.ShowZeros = model.ShowZeros && before.Zoom = model.Zoom && before.TabColor = model.TabColor)

    /// What of a sheet is not the file's, when no row or column was inserted
    /// or deleted: the cells (values, formulas, formats) and rows (height,
    /// hidden, outline) that changed, and the elements of the worksheet that
    /// are made anew - column widths, merged cells, conditional formats,
    /// validations, the view (zoom, switches, frozen panes), the tab colour,
    /// the AutoFilter, the page setup. The sheet's part is then written
    /// again with just that replaced. When comments, tables, drawings or
    /// links changed too (`Whole`), the rows are still the file's with the
    /// changed cells, and the rest of the part is made from the sheet. None
    /// when rows or columns were inserted or deleted, or array formulas and
    /// what-if tables changed: such a sheet is written from its model.
    let private patchOfSheet (before: SpreadsheetModel) (model: SpreadsheetModel) : XlsxReader.SheetPatch option =
        let same (pick: SpreadsheetModel -> 'T) = sameRef (pick before) (pick model) || pick before = pick model
        let rest =
            model.Deferred.IsNone && before.StructureVersion = model.StructureVersion
            && same (fun m -> m.DataTables) && same (fun m -> m.KeptRanges) && same (fun m -> m.ArrayMarks) && same (fun m -> m.Spill)
        if not rest then None
        else
            let cells = ResizeArray<int * int * bool>()
            // (A note is in the sheet's comments part, not in its cells.)
            let mutable notes = false
            if not (sameRef before.Cells model.Cells && sameRef before.RichRuns model.RichRuns) then
                let sameRuns = sameRef before.RichRuns model.RichRuns
                let runsSame (addr: CellAddress) = sameRuns || Map.tryFind addr before.RichRuns = Map.tryFind addr model.RichRuns
                for KeyValue(addr, data) in model.Cells do
                    match Map.tryFind addr before.Cells with
                    | Some previous when (sameRef previous data || previous = data || onlyBecameError previous data) && runsSame addr -> ()
                    | Some previous ->
                        if previous.Comment <> data.Comment then notes <- true
                        let valueOnly =
                            match previous.Value, data.Value with
                            | CellValue.Formula (formulaBefore, _), CellValue.Formula (formulaNow, _) ->
                                formulaBefore = formulaNow && (sameRef previous.Format data.Format || previous.Format = data.Format) && runsSame addr
                            | _ -> false
                        cells.Add ((addr.Row, addr.Col, valueOnly))
                    | None ->
                        if data.Comment.IsSome then notes <- true
                        cells.Add ((addr.Row, addr.Col, false))
                for KeyValue(addr, previous) in before.Cells do
                    if not (Map.containsKey addr model.Cells) then
                        if previous.Comment.IsSome then notes <- true
                        cells.Add ((addr.Row, addr.Col, false))
            // The rows as a file states them: those an AutoFilter hides are hidden rows.
            let was = filterRowsAsHidden before
            let now = filterRowsAsHidden model
            let rows =
                if sameRef was.RowHeights now.RowHeights && (sameRef was.Axes now.Axes || was.Axes = now.Axes) then []
                else
                    // Hidden or not, the height it has when it is shown, outline level, collapsed.
                    let stateOf (m: SpreadsheetModel) (row: int) =
                        let hidden, height =
                            match Map.tryFind row m.RowHeights with
                            | Some size when size <= 0.0 -> true, Map.tryFind row m.Axes.HiddenRowSizes
                            | size -> false, size
                        hidden, height, (Map.tryFind row m.Axes.RowLevels |> Option.defaultValue 0), Set.contains row m.Axes.CollapsedRows
                    let touched = Collections.Generic.HashSet<int>()
                    let diff (a: Map<int, 'V>) (b: Map<int, 'V>) =
                        if not (sameRef a b) then
                            for KeyValue(row, value) in a do (if Map.tryFind row b <> Some value then touched.Add row |> ignore)
                            for KeyValue(row, _) in b do (if not (Map.containsKey row a) then touched.Add row |> ignore)
                    diff was.RowHeights now.RowHeights
                    diff was.Axes.HiddenRowSizes now.Axes.HiddenRowSizes
                    diff was.Axes.RowLevels now.Axes.RowLevels
                    if not (sameRef was.Axes.CollapsedRows now.Axes.CollapsedRows) then
                        for row in was.Axes.CollapsedRows do (if not (Set.contains row now.Axes.CollapsedRows) then touched.Add row |> ignore)
                        for row in now.Axes.CollapsedRows do (if not (Set.contains row was.Axes.CollapsedRows) then touched.Add row |> ignore)
                    [ for row in Seq.sort touched do
                        if row >= 0 then
                            let hiddenWas, heightWas, levelWas, collapsedWas = stateOf was row
                            let hidden, height, level, collapsed = stateOf now row
                            // (A row hidden at the height every row has keeps what the file says of its height.)
                            let heightChanged =
                                heightWas <> height
                                && not (heightWas.IsNone && hidden && (match height with Some size -> abs (size - 25.0) < 0.01 | None -> true))
                            if hiddenWas <> hidden || heightChanged || levelWas <> level || collapsedWas <> collapsed then
                                yield ({ Row = row; HeightChanged = heightChanged; Height = height |> Option.map (fun size -> size / 1.333)
                                         Hidden = hidden; Level = level; Collapsed = collapsed } : XlsxReader.SheetRowPatch) ]
            let levelsOf (levels: Map<int, int>) = levels |> Map.fold (fun deepest _ level -> max deepest level) 0
            let setup (pick: PageSetup -> 'T) = pick before.PageSetup = pick model.PageSetup
            let setupSame = same (fun m -> m.PageSetup)
            Some { Cells = List.ofSeq cells
                   Rows = rows
                   Whole =
                    notes || not (same (fun m -> m.Tables)) || not (same (fun m -> m.Drawings)) || not (same (fun m -> m.Threads)) || not (same (fun m -> m.Hyperlinks))
                   Sections =
                    [ if before.TabColor <> model.TabColor || before.Axes.SummaryBelow <> model.Axes.SummaryBelow || before.Axes.SummaryRight <> model.Axes.SummaryRight
                         || before.PageSetup.FitToPage <> model.PageSetup.FitToPage then yield "sheetPr"
                      if before.FrozenRows <> model.FrozenRows || before.FrozenCols <> model.FrozenCols
                         || before.ShowGridLines <> model.ShowGridLines || before.ShowHeaders <> model.ShowHeaders || before.ShowFormulas <> model.ShowFormulas
                         || before.ShowZeros <> model.ShowZeros || before.Zoom <> model.Zoom || before.Axes.ShowSymbols <> model.Axes.ShowSymbols then yield "sheetViews"
                      if levelsOf before.Axes.RowLevels <> levelsOf model.Axes.RowLevels || levelsOf before.Axes.ColLevels <> levelsOf model.Axes.ColLevels then yield "sheetFormatPr"
                      if not (same (fun m -> m.ColumnWidths)) || not (same (fun m -> m.ColumnStyles))
                         || before.Axes.HiddenColSizes <> model.Axes.HiddenColSizes || before.Axes.ColLevels <> model.Axes.ColLevels
                         || before.Axes.CollapsedCols <> model.Axes.CollapsedCols then yield "cols"
                      if not (same (fun m -> m.AutoFilter)) || not (same (fun m -> m.Filters)) then yield "autoFilter"
                      if not (same (fun m -> m.SortState)) then yield "sortState"
                      if not (same (fun m -> m.MergedCells)) then yield "mergeCells"
                      if not (same (fun m -> m.ConditionalFormats)) then yield "conditionalFormatting"
                      if not (same (fun m -> m.Validations)) then yield "dataValidations"
                      if not setupSame then
                          if not (setup (fun p -> p.Options)) then yield "printOptions"
                          if not (setup (fun p -> p.Margins)) then yield "pageMargins"
                          // (Paper, orientation, scale ...: whatever else the setup says.)
                          yield "pageSetup"
                          if not (setup (fun p -> p.HeaderFooter)) then yield "headerFooter"
                          if not (setup (fun p -> p.ManualRowBreaks)) then yield "rowBreaks"
                          if not (setup (fun p -> p.ManualColBreaks)) then yield "colBreaks" ] }

    /// The workbook to write. `sourceParts` are parts of the opened file by
    /// path - its worksheet parts, styles and shared strings - beyond those the
    /// reader kept (`Preserved.SourceParts`; see `verbatimPartsWanted`). With
    /// them, a worksheet nothing was changed in is written as the part it was
    /// read from, byte for byte, beside the sheets that are written from
    /// their models; the styles and shared strings are then the file's with
    /// what is new added behind (no index of the file changes its meaning).
    /// When sheets of the file were renamed or deleted, that is done for the
    /// sheets whose parts do not name any of those. Any other sheet, and any
    /// whose part is not supplied, is written from the model as before.
    let toWorkbookOutWith (sourceParts: (string * byte[]) list) (input: ExportInput) : XlsxWriter.XlsxWorkbookOut =
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
        // A sheet without a tab is hidden, or very hidden when its model says so.
        let stateOf index (model: SpreadsheetModel) =
            if visibleAt index then XlsxReader.SheetVisible
            elif model.VeryHidden then XlsxReader.SheetVeryHidden
            else XlsxReader.SheetHidden
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
        // The sheets of the file as they are named now (by the old name, in upper
        // case), and those that are gone: what kept parts say of them follows.
        let fileTabs = match input.Source with Some source -> source.Workbook.Sheets |> List.map (fun sheet -> sheet.Name) | None -> []
        let renames = Collections.Generic.Dictionary<string, string>()
        let present = Collections.Generic.HashSet<string>()
        for index in 0 .. input.Sheets.Length - 1 do
            match sourceOf index with
            | Some (readerSheet, _, _) ->
                present.Add (readerSheet.Name.ToUpperInvariant()) |> ignore
                if readerSheet.Name <> nameOf index then renames.[readerSheet.Name.ToUpperInvariant()] <- nameOf index
            | None -> ()
        let deletedTabs = fileTabs |> List.filter (fun name -> not (present.Contains (name.ToUpperInvariant())))
        let withKeptFormulas (sheet: XlsxWriter.XlsxSheetOut) : XlsxWriter.XlsxSheetOut =
            if (renames.Count = 0 && deletedTabs.IsEmpty) || sheet.Extras.KeptCellExtensions.IsEmpty then sheet
            else
                let extensions = sheet.Extras.KeptCellExtensions |> List.map (keptFormulasFollow [ "xm:f" ] renames deletedTabs fileTabs)
                if extensions = sheet.Extras.KeptCellExtensions then sheet else { sheet with Extras = { sheet.Extras with KeptCellExtensions = extensions } }
        // A copy of a sheet of the file (Workbook.duplicateSheet) has the kept
        // drawing objects of its source - charts, shapes. The parts behind them
        // are the file's: a chart part is copied under a name of its own (with
        // its relationships part), its series reading the copy where they read
        // the source; what the objects otherwise refer to (the pictures of a
        // group, a chart's style and colours through its relationships) is shared.
        let clonedParts = ResizeArray<string * byte[]>()
        let clonedOverrides = ResizeArray<string * string>()
        let takenPaths = Collections.Generic.HashSet<string>()
        match input.Source with
        | Some source -> for path, _ in source.Workbook.Preserved.Parts do takenPaths.Add path |> ignore
        | None -> ()
        let copyExtras (index: int) (model: SpreadsheetModel) : XlsxReader.SheetExtras option =
            match input.Source, model.KeptFrom with
            | Some source, Some origin when origin.Book = source.Id && origin.Sheet >= 0 && origin.Sheet < source.Workbook.Sheets.Length
                                           && model.Drawings |> List.exists (fun drawing -> match drawing.Content with KeptObject _ -> true | Picture _ -> false) ->
                let readerSheet = List.item origin.Sheet source.Workbook.Sheets
                let xmlOfObjects = model.Drawings |> List.choose (fun drawing -> match drawing.Content with KeptObject (_, xml) -> Some xml | Picture _ -> None)
                let refersTo (relId: string) = xmlOfObjects |> List.exists (fun xml -> xml.IndexOf ("\"" + relId + "\"") >= 0)
                let partOf (path: string) = source.Workbook.Preserved.Parts |> List.tryFind (fun (kept, _) -> kept = path) |> Option.map snd
                let ownRenames = Collections.Generic.Dictionary<string, string>(renames)
                ownRenames.[readerSheet.Name.ToUpperInvariant()] <- nameOf index
                let relationships =
                    readerSheet.Extras.DrawingRelationships
                    |> List.filter (fun (relId, _, _, _) -> refersTo relId)
                    |> List.map (fun (relId, relType, target, isExternal) ->
                        let path = target.TrimStart '/'
                        match (if isExternal || not (isChartPart path) then None else partOf path) with
                        | None -> relId, relType, target, isExternal
                        | Some data ->
                            // "xl/charts/chart3.xml" -> the first "xl/charts/chartN.xml" no part has.
                            let stem = path.Substring(0, path.Length - 4)
                            let mutable digits = 0
                            while digits < stem.Length && Char.IsDigit stem.[stem.Length - 1 - digits] do digits <- digits + 1
                            let prefix = stem.Substring(0, stem.Length - digits)
                            let mutable number = 1
                            while takenPaths.Contains (sprintf "%s%d.xml" prefix number) do number <- number + 1
                            let copyPath = sprintf "%s%d.xml" prefix number
                            takenPaths.Add copyPath |> ignore
                            let text = Text.Encoding.UTF8.GetString data
                            clonedParts.Add ((copyPath, Text.Encoding.UTF8.GetBytes (keptFormulasFollow chartFormulaElements ownRenames deletedTabs fileTabs text)))
                            let relsOf (part: string) = let cut = part.LastIndexOf '/' in part.Substring(0, cut) + "/_rels/" + part.Substring(cut + 1) + ".rels"
                            (match partOf (relsOf path) with
                             | Some rels ->
                                clonedParts.Add ((relsOf copyPath, rels))
                                takenPaths.Add (relsOf copyPath) |> ignore
                             | None -> ())
                            (match source.Workbook.Preserved.Overrides |> List.tryFind (fun (name, _) -> name.TrimStart('/').Equals(path, StringComparison.OrdinalIgnoreCase)) with
                             | Some (_, contentType) -> clonedOverrides.Add (("/" + copyPath, contentType))
                             | None -> ())
                            relId, relType, "/" + copyPath, false)
                Some { XlsxReader.emptySheetExtras with DrawingRelationships = relationships }
            | _ -> None
        // Sheets that go back as the parts they were read from.
        let suppliedParts = Collections.Generic.Dictionary<string, byte[]>()
        match input.Source with
        | Some source ->
            for path, data in source.Workbook.Preserved.SourceParts do suppliedParts.[path] <- data
            for path, data in sourceParts do suppliedParts.[path] <- data
        | None -> ()
        // (Sheets go back as they were read - in either date system: the serials
        // are the file's - while every sheet of the file is there under its
        // name; the styles and shared strings are added to whenever the file's
        // are at hand.)
        let verbatimAllowed = input.Source.IsSome
        // Sheets of the file that were renamed or deleted: a sheet whose part
        // names one of them (a formula, a rule, a link) is written from the
        // model, where the references have followed; a part that does not
        // name any still goes back as it was read. A name is looked for in
        // every way a part can write it: as it is, with an apostrophe doubled
        // (inside a quoted reference), and with `& < >`, the apostrophe and
        // the quotation mark as XML writes them.
        let changedTabs =
            fileTabs |> List.filter (fun name -> renames.ContainsKey (name.ToUpperInvariant()) || not (present.Contains (name.ToUpperInvariant())))
        let changedTabForms =
            changedTabs
            |> List.filter (fun name -> name <> "")
            |> List.collect (fun name ->
                let lower = name.ToLowerInvariant()
                [ for quoted in [ lower; lower.Replace("'", "''") ] do
                    let inXml = quoted.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                    for apostrophe in [ "'"; "&apos;"; "&#39;"; "&#x27;" ] do
                        for quotation in [ "\""; "&quot;"; "&#34;"; "&#x22;" ] do
                            yield inXml.Replace("'", apostrophe).Replace("\"", quotation) ])
            |> List.distinct
        let namesChangedTab (data: byte[]) : bool =
            if changedTabForms.IsEmpty then false
            else
                let text = (Text.Encoding.UTF8.GetString data).ToLowerInvariant()
                changedTabForms |> List.exists (fun form -> text.IndexOf (form, StringComparison.Ordinal) >= 0)
        let inPlaceParts : (string * byte[]) list option =
            match input.Source with
            | Some source ->
                let kept = source.Workbook.Preserved
                let part (path: string option) = path |> Option.bind (fun path -> match suppliedParts.TryGetValue path with | true, data -> Some (path, data) | _ -> None)
                match part kept.StylesPart, kept.SharedStringsPart, part kept.SharedStringsPart with
                | Some styles, None, _ -> Some [ styles ]
                | Some styles, Some _, Some strings -> Some [ styles; strings ]
                | _ -> None
            | _ -> None
        let verbatimParts = ResizeArray<string * byte[]>()
        // A sheet with little changed (cells, widths, merges, rules) goes back
        // as its part with that changed in it.
        let patchOf (outSheet: XlsxWriter.XlsxSheetOut) (before: SpreadsheetModel) (model: SpreadsheetModel) : XlsxReader.SheetPatch option =
            match inPlaceParts, outSheet.Extras.SourcePart with
            | Some _, Some (path, relType) when verbatimAllowed && relType.EndsWith "/worksheet" && suppliedParts.ContainsKey path && not (namesChangedTab suppliedParts.[path]) ->
                match patchOfSheet before model with
                | Some patch ->
                    if not patch.Whole then verbatimParts.Add ((path, suppliedParts.[path]))
                    Some patch
                | None -> None
            | _ -> None
        // A sheet of the file that is written from its model keeps what its part
        // says and a model does not hold (see `SheetExtras.SourceXml`), by where
        // the rows and columns of the file are now.
        let withSource (outSheet: XlsxWriter.XlsxSheetOut) (model: SpreadsheetModel) (written: XlsxWriter.XlsxSheetOut) : XlsxWriter.XlsxSheetOut =
            match outSheet.Extras.SourcePart with
            | Some (path, relType) when relType.EndsWith "/worksheet" && suppliedParts.ContainsKey path ->
                let moves =
                    if model.StructureLog.Length <> model.StructureVersion then None
                    else Some (model.StructureLog |> List.rev |> List.map (fun move -> move.OnRows, move.Insert, move.At, move.Count))
                { written with Extras = { written.Extras with SourceXml = Some (Text.Encoding.UTF8.GetString suppliedParts.[path]); Moves = moves } }
            | _ -> written
        let verbatimOf (outSheet: XlsxWriter.XlsxSheetOut) (before: SpreadsheetModel) (model: SpreadsheetModel) : bool =
            match inPlaceParts, outSheet.Extras.SourcePart with
            | Some _, Some (path, relType) when verbatimAllowed && relType.EndsWith "/worksheet" && suppliedParts.ContainsKey path ->
                let current = match model.Deferred with Some deferred -> (match deferred.Loaded with Some loaded -> loaded | None -> model) | None -> model
                let untouched =
                    if isDeferredSheet model then model.TabColor = outSheet.View.TabColor
                    else untouchedSheet before current
                let untouched = untouched && not (namesChangedTab suppliedParts.[path])
                if untouched then verbatimParts.Add ((path, suppliedParts.[path]))
                untouched
            | _ -> false
        let sheets =
            input.Sheets
            |> Array.mapi (fun index model ->
                let name = nameOf index
                match sourceOf index with
                | Some (_, outSheet, before) when verbatimOf outSheet before model ->
                    // (As the file has it: should the writer not be able to add to
                    // the file's styles, it writes the sheet from the file's cells.)
                    { outSheet with Name = name; State = stateOf index model; Extras = { outSheet.Extras with Verbatim = true } }
                | Some (readerSheet, outSheet, before) ->
                    // A sheet that was never converted goes back as the file had
                    // it (but for its name, visibility and tab colour, which
                    // change without converting it).
                    if isDeferredSheet model then withKeptFormulas { outSheet with Name = name; State = stateOf index model; View = { outSheet.View with TabColor = model.TabColor } }
                    else
                        let written = withKeptFormulas (sourcedSheet cache dxfIndex name (stateOf index model) outSheet readerSheet (filterRowsAsHidden before) (filterRowsAsHidden model))
                        match patchOf outSheet before model with
                        | Some patch when not patch.Whole -> { written with Extras = { written.Extras with Verbatim = true; Patch = Some patch } }
                        | Some patch -> withSource outSheet model { written with Extras = { written.Extras with Patch = Some patch } }
                        | None -> withSource outSheet model written
                | None ->
                    let written = modelSheet cache dxfIndex name (stateOf index model) (filterRowsAsHidden model)
                    match copyExtras index model with
                    | Some kept ->
                        let extras = extrasOfModel kept None model
                        { written with Extras = { written.Extras with Objects = extras.Objects; Images = extras.Images; DrawingRelationships = extras.DrawingRelationships } }
                    | None -> written)
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
                // What the kept parts say of sheets that were renamed or deleted; the parts of copied charts.
                let preserved =
                    if renames.Count = 0 && deletedTabs.IsEmpty && clonedParts.Count = 0 then preserved
                    else
                        let cachePaths = preserved.PivotCacheSources |> List.map (fun (path, _, _) -> path) |> Set.ofList
                        let follows (path: string) (data: byte[]) : byte[] =
                            if renames.Count = 0 && deletedTabs.IsEmpty then data
                            elif isChartPart path then
                                let text = Text.Encoding.UTF8.GetString data
                                let rewritten = keptFormulasFollow chartFormulaElements renames deletedTabs fileTabs text
                                if rewritten = text then data else Text.Encoding.UTF8.GetBytes rewritten
                            elif Set.contains path cachePaths then
                                let text = Text.Encoding.UTF8.GetString data
                                let rewritten = pivotCacheSheetFollows renames text
                                if rewritten = text then data else Text.Encoding.UTF8.GetBytes rewritten
                            else data
                        { preserved with
                            Parts = (preserved.Parts |> List.map (fun (path, data) -> path, follows path data)) @ List.ofSeq clonedParts
                            Overrides = preserved.Overrides @ List.ofSeq clonedOverrides
                            PivotCacheSources =
                                preserved.PivotCacheSources
                                |> List.map (fun (path, sheet, range) ->
                                    match renames.TryGetValue (sheet.ToUpperInvariant()) with
                                    | true, next -> path, next, range
                                    | false, _ -> path, sheet, range) }
                // The parts of the sheets that go back as they were read, and the
                // styles and shared strings they index.
                let preserved =
                    match inPlaceParts with
                    | Some parts ->
                        // (With what their relationships lead to: the writer leaves out what nothing written reaches.)
                        let sheetPaths = source.Workbook.Sheets |> List.choose (fun sheet -> sheet.Extras.SourcePart |> Option.map fst) |> Set.ofList
                        let own = parts @ List.ofSeq verbatimParts
                        let ownNames = own |> List.map fst |> Set.ofList
                        let others =
                            source.Workbook.Preserved.SourceParts
                            |> List.filter (fun (path, _) -> not (Set.contains path sheetPaths) && not (Set.contains path ownNames))
                        let added = own @ others
                        let names = added |> List.map fst |> Set.ofList
                        { preserved with InPlace = true; Parts = (preserved.Parts |> List.filter (fun (path, _) -> not (Set.contains path names))) @ added }
                    | None ->
                        // (Without them: what the reader kept beside the sheets - the workbook part, comments parts.)
                        let known = preserved.Parts |> List.map fst |> Set.ofList
                        let sheetPaths = source.Workbook.Sheets |> List.choose (fun sheet -> sheet.Extras.SourcePart |> Option.map fst) |> Set.ofList
                        let others =
                            source.Workbook.Preserved.SourceParts
                            |> List.filter (fun (path, _) ->
                                not (Set.contains path sheetPaths) && not (Set.contains path known)
                                && Some path <> preserved.StylesPart && Some path <> preserved.SharedStringsPart)
                        if others.IsEmpty then preserved else { preserved with Parts = preserved.Parts @ others }
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

    /// The workbook to write (see `toWorkbookOutWith`): with the parts of the
    /// file the reader kept the bytes of - all of them for a file read by a
    /// .NET host, none of the large ones in a browser.
    let toWorkbookOut (input: ExportInput) : XlsxWriter.XlsxWorkbookOut = toWorkbookOutWith [] input

    /// The parts of the opened file `toWorkbookOutWith` has to be given so
    /// that every sheet nothing was changed in goes back as it was read: the
    /// parts of those sheets, the styles and the shared strings - less those
    /// the reader kept. Empty when no sheet would go back that way.
    let verbatimPartsWanted (input: ExportInput) : string list =
        match input.Source with
        | Some source ->
            let kept = source.Workbook.Preserved
            let readerSheets = List.toArray source.Workbook.Sheets
            let claimed = Collections.Generic.HashSet<int>()
            let sheetParts =
                input.Sheets
                |> Array.toList
                |> List.choose (fun model ->
                    match model.SourceSheet with
                    | Some origin when origin.Book = source.Id && origin.Sheet >= 0 && origin.Sheet < readerSheets.Length && origin.Sheet < source.Baseline.Length && claimed.Add origin.Sheet ->
                        let current = match model.Deferred with Some deferred -> (match deferred.Loaded with Some loaded -> loaded | None -> model) | None -> model
                        let reader = readerSheets.[origin.Sheet]
                        let untouched =
                            if isDeferredSheet model then model.TabColor = reader.View.TabColor
                            else untouchedSheet source.Baseline.[origin.Sheet] current
                        // (... or changed: its part is written again with that changed in it, or the
                        // sheet is written from its model with what the part says beyond a model.)
                        match reader.Extras.SourcePart with
                        | Some (path, relType) when (untouched || not (isDeferredSheet model)) && relType.EndsWith "/worksheet" -> Some path
                        | _ -> None
                    | _ -> None)
            (
                let have = kept.SourceParts |> List.map fst |> Set.ofList
                sheetParts @ Option.toList kept.StylesPart @ Option.toList kept.SharedStringsPart |> List.distinct |> List.filter (fun path -> not (Set.contains path have)))
        | None -> []

    /// The workbook to write with the workbook's setting for the iterative
    /// calculation of circular references (`Workbook.Iteration`).
    let withIteration (iteration: Wb.IterativeCalculation option) (workbook: XlsxWriter.XlsxWorkbookOut) : XlsxWriter.XlsxWorkbookOut =
        { workbook with Preserved = { workbook.Preserved with CalcIteration = iteration |> Option.map (fun setting -> setting.MaxIterations, setting.MaxChange) } }
