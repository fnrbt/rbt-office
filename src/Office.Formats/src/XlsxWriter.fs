module XlsxWriter

open Fable.Core.JsInterop
open XmlGenerator
open ZipWriter
open ZipInterop
open System.Text

/// Convert column index to Excel letter (0 -> A, 25 -> Z, 26 -> AA)
let columnToLetter (col: int) : string =
    let rec loop n acc =
        if n < 0 then acc
        else loop ((n / 26) - 1) (char (65 + (n % 26)) :: acc)
    loop col [] |> Array.ofList |> System.String

/// Convert row/col to cell reference (0, 0 -> A1)
let cellRef (row: int) (col: int) : string =
    sprintf "%s%d" (columnToLetter col) (row + 1)

/// Try to parse a value as a number
let tryParseNumber (value: string) : float option =
    // Invariant: the text is written back into the file as a number, so it
    // must not depend on the machine's culture ("1,5" is not 15).
    match System.Double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
    | (true, num) when not (System.Double.IsNaN num || System.Double.IsInfinity num) -> Some num
    | _ -> None

/// Number as file text: invariant, round-trippable.
let internal numberText (value: float) : string =
#if FABLE_COMPILER
    string value
#else
    value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
#endif

/// Characters XML 1.0 cannot carry at all (C0 controls other than tab, LF,
/// CR, and the two non-characters). Emitting them makes the part unreadable.
let internal isXmlIllegal (ch: char) =
    (ch < ' ' && ch <> '\t' && ch <> '\n' && ch <> '\r') || ch = '\uFFFE' || ch = '\uFFFF'

/// Escape for element text. Illegal characters are dropped; a carriage
/// return is written as a character reference so line-end normalisation
/// does not turn it into a line feed.
let internal appendEscapedChar (sb: StringBuilder) (ch: char) =
    match ch with
    | '&' -> sb.Append("&amp;") |> ignore
    | '<' -> sb.Append("&lt;") |> ignore
    | '>' -> sb.Append("&gt;") |> ignore
    | '"' -> sb.Append("&quot;") |> ignore
    | '\'' -> sb.Append("&apos;") |> ignore
    | '\r' -> sb.Append("&#13;") |> ignore
    | _ -> if not (isXmlIllegal ch) then sb.Append(ch) |> ignore

let internal appendEscaped (sb: StringBuilder) (text: string) =
    for ch in text do
        appendEscapedChar sb ch

/// Escape for an attribute value: as `appendEscaped`, and tab / line feed
/// become character references too, since attribute-value normalisation
/// would otherwise turn them into spaces (multi-line validation messages).
let internal appendEscapedAttr (sb: StringBuilder) (text: string) =
    for ch in text do
        match ch with
        | '\t' -> sb.Append("&#9;") |> ignore
        | '\n' -> sb.Append("&#10;") |> ignore
        | _ -> appendEscapedChar sb ch

let internal hexDigits = "0123456789ABCDEF"

let internal appendXEscape (sb: StringBuilder) (ch: char) =
    let code = int ch
    sb.Append("_x")
        .Append(hexDigits.[(code >>> 12) &&& 15])
        .Append(hexDigits.[(code >>> 8) &&& 15])
        .Append(hexDigits.[(code >>> 4) &&& 15])
        .Append(hexDigits.[code &&& 15])
        .Append('_') |> ignore

let internal isHexDigit (c: char) =
    (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')

/// Escape cell text (ST_Xstring: shared strings, inline strings, cached
/// string results, comments). Characters XML cannot carry — and a carriage
/// return — are written as `_xHHHH_`, the way Excel stores them, and an
/// underscore that would itself read as such an escape is written `_x005F_`.
/// The reader decodes both.
let internal appendCellText (sb: StringBuilder) (text: string) =
    for i in 0 .. text.Length - 1 do
        let ch = text.[i]
        if (ch < ' ' && ch <> '\t' && ch <> '\n') || ch = '\uFFFE' || ch = '\uFFFF' then
            appendXEscape sb ch
        elif ch = '_'
             && i + 6 < text.Length
             && text.[i + 1] = 'x'
             && text.[i + 6] = '_'
             && isHexDigit text.[i + 2] && isHexDigit text.[i + 3] && isHexDigit text.[i + 4] && isHexDigit text.[i + 5] then
            appendXEscape sb '_'
        else
            appendEscapedChar sb ch

let internal appendCellTextElement (sb: StringBuilder) (name: string) (text: string) =
    sb.Append('<').Append(name) |> ignore
    if text = "" then
        sb.Append("/>") |> ignore
    else
        // Leading / trailing whitespace is only kept with xml:space="preserve".
        sb.Append(" xml:space=\"preserve\">") |> ignore
        appendCellText sb text
        sb.Append("</").Append(name).Append('>') |> ignore

let internal appendTextElement (sb: StringBuilder) (name: string) (text: string) =
    sb.Append('<').Append(name) |> ignore
    if text = "" then
        sb.Append("/>") |> ignore
    else
        sb.Append('>') |> ignore
        appendEscaped sb text
        sb.Append("</").Append(name).Append('>') |> ignore

let internal appendCellRef (sb: StringBuilder) (colLetters: string) (rowNumber: string) =
    sb.Append(colLetters).Append(rowNumber) |> ignore

let internal appendSimpleCell (sb: StringBuilder) (rowNumber: string) (colLetters: string) (value: string) =
    sb.Append("<c r=\"") |> ignore
    appendCellRef sb colLetters rowNumber
    match tryParseNumber value with
    | Some num ->
        sb.Append("\"><v>").Append(numberText num).Append("</v></c>") |> ignore
    | None ->
        sb.Append("\" t=\"inlineStr\"><is>") |> ignore
        appendCellTextElement sb "t" value
        sb.Append("</is></c>") |> ignore

/// Generate cell element
let generateCell (rowIdx: int) (colIdx: int) (value: string) : string =
    let ref = cellRef rowIdx colIdx

    match tryParseNumber value with
    | Some num ->
        // Numeric cell
        element "c" [("r", ref)] [
            elementWithText "v" [] (numberText num)
        ]
    | None ->
        // Inline string cell
        element "c" [("r", ref); ("t", "inlineStr")] [
            element "is" [] [
                elementWithText "t" [] value
            ]
        ]

/// Generate sheet XML from 2D array
let generateSheetXml (data: string list list) : string =
    let columnLetters =
        let maxCols =
            data
            |> List.fold (fun maxCols row -> max maxCols row.Length) 0
        [| for col in 0 .. maxCols - 1 -> columnToLetter col |]

    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<worksheet xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\">") |> ignore
    sb.Append("<sheetData>") |> ignore
    let mutable rowIdx = 0
    for row in data do
        let rowNumber = string (rowIdx + 1)
        sb.Append("<row r=\"").Append(rowNumber).Append("\">") |> ignore
        let mutable colIdx = 0
        for value in row do
            appendSimpleCell sb rowNumber columnLetters.[colIdx] value
            colIdx <- colIdx + 1
        sb.Append("</row>") |> ignore
        rowIdx <- rowIdx + 1
    sb.Append("</sheetData></worksheet>") |> ignore
    sb.ToString()

/// Generate workbook.xml
let generateWorkbookXml (sheetNames: string list) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<workbook xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\"><sheets>") |> ignore
    let mutable index = 0
    for name in sheetNames do
        let sheetId = index + 1
        sb.Append("<sheet name=\"") |> ignore
        appendEscaped sb name
        sb.Append("\" sheetId=\"").Append(sheetId)
            .Append("\" r:id=\"rId").Append(sheetId).Append("\"/>") |> ignore
        index <- index + 1
    sb.Append("</sheets></workbook>") |> ignore
    sb.ToString()

/// Generate workbook relationships
let generateWorkbookRels (sheetCount: int) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<Relationships xmlns=\"").Append(Namespaces.relationshipsPackage).Append("\">") |> ignore
    for i in 1 .. sheetCount do
        sb.Append("<Relationship Id=\"rId").Append(i)
            .Append("\" Type=\"").Append(RelationshipTypes.worksheet)
            .Append("\" Target=\"worksheets/sheet").Append(i).Append(".xml\"/>") |> ignore
    sb.Append("</Relationships>") |> ignore
    sb.ToString()

/// Generate [Content_Types].xml for XLSX
let generateContentTypesXlsx (sheetCount: int) : string =
    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<Types xmlns=\"").Append(Namespaces.contentTypes).Append("\">") |> ignore
    sb.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>") |> ignore
    sb.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>") |> ignore
    sb.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>") |> ignore
    for i in 1 .. sheetCount do
        sb.Append("<Override PartName=\"/xl/worksheets/sheet").Append(i)
            .Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>") |> ignore
    sb.Append("</Types>") |> ignore
    sb.ToString()

/// Generate root relationships
let generateRootRelsXlsx () : string =
    let rel = selfClosing "Relationship" [("Id", "rId1"); ("Type", RelationshipTypes.officeDocument); ("Target", "xl/workbook.xml")]
    let root = element "Relationships" [("xmlns", Namespaces.relationshipsPackage)] [rel]
    document root

/// Type for sheet data
type SheetData = {
    Name: string
    Data: string list list
}

/// Type for formatted cell
type FormattedCell = {
    Value: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    FontSize: float option
    FontFamily: string option
    FontColor: string option
    FillColor: string option
    NumberFormat: string option
    Formula: string option  // Excel formula (without =)
}

/// Type for formatted row
type FormattedRow = {
    Cells: FormattedCell list
}

/// Type for merged cell range
type MergedCellRange = {
    StartRow: int  // 0-indexed
    StartCol: int  // 0-indexed
    EndRow: int    // 0-indexed
    EndCol: int    // 0-indexed
}

/// Type for formatted sheet
type FormattedSheet = {
    Name: string
    Rows: FormattedRow list
    MergedCells: MergedCellRange list
}

type XlsxBorderSide = {
    Style: string
    Color: string option
}

type XlsxCellFormat = {
    Bold: bool
    Italic: bool
    Underline: bool
    UnderlineDouble: bool
    Strikethrough: bool
    FontName: string option
    TextColor: string option
    FillColor: string option
    FontSize: float option
    NumberFormat: string option
    HorizontalAlign: string option
    VerticalAlign: string option
    WrapText: bool
    Indent: int
    ShrinkToFit: bool
    TextRotation: int
    BorderTop: XlsxBorderSide option
    BorderBottom: XlsxBorderSide option
    BorderLeft: XlsxBorderSide option
    BorderRight: XlsxBorderSide option
    BorderDiagonal: XlsxBorderSide option
    DiagonalUp: bool
    DiagonalDown: bool
    /// The named cell style the format is based on: an index into the
    /// workbook's `Preserved.NamedStyles`. None is the first one (Normal).
    NamedStyle: int option
    /// The cell format of the opened file this is, exactly (its cellXfs
    /// index). A writer that adds to the file's styles writes the cell with
    /// that entry, so what the model does not hold of it - protection, quote
    /// prefix, pivot button, theme colours - stays. None: a format of the model.
    SourceXf: int option
    /// The cell format of the opened file this was made from by a change of
    /// the user's (its cellXfs index): a writer that adds to the file's
    /// styles makes the new entry from that one, so that what the change did
    /// not touch - protection, quote prefix, pivot button, the named style,
    /// a font, fill or border the model has no full account of - stays.
    BasedOnXf: int option
}

type XlsxVertAlign =
    | XlsxBaseline
    | XlsxSuperscript
    | XlsxSubscript

type XlsxRichRun = {
    Text: string
    Bold: bool option
    Italic: bool option
    Underline: bool option
    UnderlineDouble: bool option
    Strikethrough: bool option
    Color: string option
    FontName: string option
    FontSize: float option
    VertAlign: XlsxVertAlign option
}

type XlsxCellValue =
    | XlsxBlank
    | XlsxNumber of float
    | XlsxText of string * XlsxRichRun list
    | XlsxBoolean of bool
    | XlsxError of string
    | XlsxFormula of string * XlsxCellValue option

type XlsxCellOut = {
    Row: int
    Col: int
    Value: XlsxCellValue
    Format: XlsxCellFormat
    Comment: (string * string option) option
}

type XlsxDataValidation = {
    Sqref: string
    ValidationType: string
    Operator: string option
    Formula1: string option
    Formula2: string option
    AllowBlank: bool
    ShowDropDown: bool
    ErrorMessage: string option
    InputMessage: string option
    /// Caption of the error alert / the input message.
    ErrorTitle: string option
    PromptTitle: string option
    /// "stop" (default), "warning" or "information".
    ErrorStyle: string option
    /// Excel only enforces a validation when showErrorMessage is set, and
    /// only shows the input message when showInputMessage is.
    ShowErrorMessage: bool
    ShowInputMessage: bool
}

type XlsxTableStyle = {
    Name: string
    ShowFirstColumn: bool
    ShowLastColumn: bool
    ShowRowStripes: bool
    ShowColumnStripes: bool
}

type XlsxTableOut = {
    Name: string
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    Columns: string list
    HeaderRowCount: int
    TotalsRowCount: int
    /// What each column shows in the totals row, left to right: ("label",
    /// text), ("function", "sum" | "count" | ...), ("custom", formula) or
    /// ("", "") for nothing ([] when not known).
    Totals: (string * string) list
    AutoFilter: (int * int) option
    Style: XlsxTableStyle option
    /// The part of the opened file the table was read from (one of the
    /// workbook's `Preserved.Parts`). A table that is no longer what that
    /// part says is written as that part with what changed replaced, so it
    /// keeps what the model does not hold: its id and revision ids, the
    /// formats of its columns (dxf ids), calculated-column formulas, its
    /// filter criteria and the binding to its query table.
    SourcePart: string option
}

type XlsxPageMargins = {
    Top: float
    Bottom: float
    Left: float
    Right: float
    Header: float
    Footer: float
}

type XlsxPageSetup = {
    Orientation: string option
    FitToPage: bool
    FitToWidth: int
    FitToHeight: int
    Scale: int
    PaperSize: int option
    FirstPageNumber: int
    UseFirstPageNumber: bool
    PageOrder: string option
    BlackAndWhite: bool
    CellComments: string option
    /// Printer attributes kept from the file (XlsxReader.PageSetup.Kept).
    Kept: (string * string) list
}

type XlsxPrintOptions = {
    GridLines: bool
    Headings: bool
    HorizontalCentered: bool
    VerticalCentered: bool
}

type XlsxHeaderFooter = {
    OddHeader: string option
    OddFooter: string option
    EvenHeader: string option
    EvenFooter: string option
    FirstHeader: string option
    FirstFooter: string option
    DifferentOddEven: bool
    DifferentFirst: bool
    ScaleWithDoc: bool
}

type XlsxHyperlink = {
    Range: string
    Target: string
    IsInternal: bool
    Tooltip: string option
    Display: string option
}

/// What a sheet's view shows besides panes and grid lines.
type XlsxSheetViewOut = {
    /// False: cells with the number 0 show nothing.
    ShowZeros: bool
    /// False: no row numbers and column letters.
    ShowHeaders: bool
    /// True: cells show their formulas.
    ShowFormulas: bool
    /// Zoom in percent (10..400).
    ZoomScale: int
    /// The tab's colour, as XlsxReader.SheetViewOptions.TabColor has it
    /// ("#RRGGBB", "theme:N", "theme:N:tint" or "indexed:N").
    TabColor: string option
    /// Where the outline's summary rows / columns are: below their group
    /// (else above), right of it (else left). Written as `<outlinePr>` when
    /// one of them is false.
    SummaryBelow: bool
    SummaryRight: bool
    /// False: the outline's symbols are hidden (`<sheetView showOutlineSymbols="0">`).
    ShowOutlineSymbols: bool
}

let defaultSheetViewOut : XlsxSheetViewOut = { ShowZeros = true; ShowHeaders = true; ShowFormulas = false; ZoomScale = 100; TabColor = None; SummaryBelow = true; SummaryRight = true; ShowOutlineSymbols = true }

type XlsxSheetOut = {
    Name: string
    /// Visible, hidden, or very hidden (only un-hidable from VBA).
    State: XlsxReader.SheetState
    /// The sheet's cells (see `Cells`). Those of a sheet of a file whose rows
    /// were not read yet are made when something asks for them: a sheet that
    /// goes back as the part it was read from never needs them.
    CellSource: XlsxReader.OnDemand<XlsxCellOut list>
    MergedCells: MergedCellRange list
    ColumnWidths: (int * float) list
    BestFitColumns: Set<int>
    RowHeights: (int * float) list
    FrozenRows: int
    FrozenCols: int
    ShowGridLines: bool
    /// The rest of the sheet view: zeros, headers, formula view, zoom.
    View: XlsxSheetViewOut
    Tables: XlsxTableOut list
    DataValidations: XlsxDataValidation list
    /// Conditional formatting, in the reader's shape. A rule's `DxfId`
    /// indexes `XlsxWorkbookOut.Dxfs`.
    ConditionalFormats: XlsxReader.ConditionalFormat list
    Hyperlinks: XlsxHyperlink list
    PageMargins: XlsxPageMargins option
    PageSetup: XlsxPageSetup option
    PrintOptions: XlsxPrintOptions option
    HeaderFooter: XlsxHeaderFooter option
    RowBreaks: int list
    ColBreaks: int list
    PrintArea: string option
    PrintTitleRows: string option
    PrintTitleCols: string option
    /// Pictures and other drawing objects anchored to cells, and threaded
    /// comments (in the reader's shape; XlsxReader.emptySheetExtras for none).
    Extras: XlsxReader.SheetExtras
}
with
    /// The sheet's cells. (Makes them now when that was left for later.)
    member this.Cells : XlsxCellOut list = this.CellSource.Value

/// Cells for `XlsxSheetOut.CellSource`.
let sheetCells (cells: XlsxCellOut list) : XlsxReader.OnDemand<XlsxCellOut list> = XlsxReader.OnDemand<XlsxCellOut list>.Of cells

type XlsxDefinedName = {
    Name: string
    RefersTo: string
    LocalSheetIndex: int option
    Comment: string option
    Hidden: bool
}

type XlsxWorkbookOut = {
    Sheets: XlsxSheetOut list
    ActiveSheetIndex: int
    DefinedNames: XlsxDefinedName list
    VbaProjectSource: string option
    /// A whole VBA project to write (a project of several modules edited in the app,
    /// `VbaProject.writeProjectModules`); it replaces the kept one and its signatures.
    VbaProjectBytes: byte[] option
    /// Differential formats conditional-formatting rules refer to by index.
    Dxfs: XlsxReader.Dxf list
    /// What was kept of the opened package (theme, chart parts, named cell
    /// styles); XlsxReader.emptyWorkbookPreservation for a new workbook.
    Preserved: XlsxReader.WorkbookPreservation
}

let defaultXlsxCellFormat = {
    Bold = false
    Italic = false
    Underline = false
    UnderlineDouble = false
    Strikethrough = false
    FontName = None
    TextColor = None
    FillColor = None
    FontSize = None
    NumberFormat = None
    HorizontalAlign = None
    VerticalAlign = None
    WrapText = false
    Indent = 0
    ShrinkToFit = false
    TextRotation = 0
    BorderTop = None
    BorderBottom = None
    BorderLeft = None
    BorderRight = None
    BorderDiagonal = None
    DiagonalUp = false
    DiagonalDown = false
    NamedStyle = None
    SourceXf = None
    BasedOnXf = None
}

/// A formula as the file format stores it: prefixes added where they are missing (see `FormulaNames`).
let fileFormula (formula: string) : string = FormulaNames.fileFormula formula

/// A formula as the model holds it and Excel shows it (see `FormulaNames`).
let modelFormula (formula: string) : string = FormulaNames.modelFormula formula

/// A run of a rich text of the file as the writer has it.
let richRunOfReader (props: XlsxReader.RunProps) (text: string) : XlsxRichRun =
    { Text = text
      Bold = Some props.Bold
      Italic = Some props.Italic
      Underline = Some props.Underline
      UnderlineDouble = Some props.UnderlineDouble
      Strikethrough = Some props.Strikethrough
      // (A colour of the theme is not written back as such: the run takes the cell's.)
      Color = props.Color |> Option.filter (fun c -> not (c.StartsWith "theme:")) |> Option.map (fun c -> "#" + c.TrimStart('#'))
      FontName = props.FontName
      FontSize = props.FontSize
      VertAlign =
        match props.VertAlign with
        | XlsxReader.VertAlignSuperscript -> Some XlsxSuperscript
        | XlsxReader.VertAlignSubscript -> Some XlsxSubscript
        | _ -> None }

let private cellRefFromReader (r: XlsxReader.CellReference) =
    let col =
        let mutable acc = 0
        for ch in r.Column.ToUpperInvariant() do
            acc <- acc * 26 + int ch - int 'A' + 1
        acc - 1
    r.Row - 1, col

/// A reader cell style as a writer cell format.
let formatOfReaderStyle (style: XlsxReader.CellStyle option) : XlsxCellFormat =
        match style with
        | None -> defaultXlsxCellFormat
        | Some s ->
            let side (b: XlsxReader.CellBorderSide option) =
                b |> Option.map (fun x -> { Style = x.Style; Color = x.Color })
            { defaultXlsxCellFormat with
                Bold = s.Bold
                Italic = s.Italic
                Underline = s.Underline
                UnderlineDouble = s.UnderlineDouble
                Strikethrough = s.Strikethrough
                FontName = s.FontName
                TextColor = s.FontColor |> Option.map (fun c -> "#" + c.TrimStart('#'))
                FillColor = s.FillColor |> Option.map (fun c -> "#" + c.TrimStart('#'))
                FontSize = s.FontSize
                NumberFormat = s.NumberFormat
                HorizontalAlign = s.HorizontalAlign
                VerticalAlign = s.VerticalAlign
                WrapText = s.WrapText
                Indent = s.Indent
                ShrinkToFit = s.ShrinkToFit
                TextRotation = s.TextRotation
                BorderTop = side s.BorderTop
                BorderBottom = side s.BorderBottom
                BorderLeft = side s.BorderLeft
                BorderRight = side s.BorderRight
                BorderDiagonal = side s.BorderDiagonal
                DiagonalUp = s.DiagonalUp
                DiagonalDown = s.DiagonalDown }

let packageFromReaderWorkbook (workbook: XlsxReader.Workbook) : XlsxWorkbookOut =
    let styleOf = formatOfReaderStyle
    // The named cell style a cell's format is based on (xfId of its cellXfs entry).
    let namedStyleOf (cell: XlsxReader.Cell) =
        match cell.StyleIndex with
        | Some index when index >= 0 && index < workbook.Preserved.CellXfNamedStyles.Length ->
            let named = workbook.Preserved.CellXfNamedStyles.[index]
            if named > 0 then Some named else None
        | _ -> None
    let runOf = richRunOfReader
    let plainValueOf (cell: XlsxReader.Cell) =
        match cell.Value with
        | XlsxReader.Number n -> XlsxNumber n
        | XlsxReader.String s -> XlsxText (s, cell.InlineRuns |> List.map (fun (t, p) -> runOf p t))
        | XlsxReader.SharedString idx when idx >= 0 && idx < workbook.SharedStrings.Length ->
            let runs =
                if idx < workbook.SharedStringRuns.Length then
                    workbook.SharedStringRuns.[idx] |> List.map (fun (t, p) -> runOf p t)
                else []
            XlsxText (workbook.SharedStrings.[idx], runs)
        | XlsxReader.Boolean b -> XlsxBoolean b
        | XlsxReader.Error e -> XlsxError (FormulaNames.errorValue e)
        | _ -> XlsxBlank
    // A formula is written with the result the source file cached for it.
    let valueOf (cell: XlsxReader.Cell) =
        match cell.Formula, cell.Value with
        | Some f, XlsxReader.Formula _ -> XlsxFormula (f, None)
        | Some f, _ -> XlsxFormula (f, Some (plainValueOf cell))
        | None, XlsxReader.Formula f -> XlsxFormula (f, None)
        | None, _ -> plainValueOf cell
    // One format record per cell format of the file, shared by every cell
    // that uses it (the writer then finds a cell's style by identity).
    let formatCache = System.Collections.Generic.Dictionary<int, XlsxCellFormat>()
    let formatOf (cell: XlsxReader.Cell) =
        let key = cell.StyleIndex |> Option.defaultValue -1
        match Portable.Dict.tryFind key formatCache with
        | ValueSome format -> format
        | ValueNone ->
            let format =
                { styleOf cell.Style with NamedStyle = namedStyleOf cell; SourceXf = Some (defaultArg cell.StyleIndex 0) }
            formatCache.[key] <- format
            format
    { Sheets =
        workbook.Sheets
        |> List.map (fun sheet ->
            let comments =
                sheet.Comments
                |> List.map (fun c ->
                    let r, col = cellRefFromReader c.CellRef
                    (r, col), (c.Text, c.Author))
                |> Map.ofList
            { Name = sheet.Name
              State = sheet.State
              CellSource =
                (let cells () =
                    sheet.Rows
                    |> List.collect (fun row ->
                        row.Cells
                        |> List.map (fun cell ->
                            let r, c = cellRefFromReader cell.Reference
                            { Row = r
                              Col = c
                              Value = valueOf cell
                              Format = formatOf cell
                              Comment = (if comments.IsEmpty then None else Map.tryFind (r, c) comments) }))
                 // (Made when asked for: a sheet that goes back as the part it was read from needs none,
                 // and rows that are still in the text of their part stay there.)
                 XlsxReader.OnDemand<XlsxCellOut list>.Deferred cells)
              MergedCells =
                sheet.MergedCells
                |> List.map (fun m ->
                    let sr, sc = cellRefFromReader m.StartRef
                    let er, ec = cellRefFromReader m.EndRef
                    { StartRow = sr; StartCol = sc; EndRow = er; EndCol = ec })
              ColumnWidths =
                sheet.Columns
                |> List.collect (fun c ->
                    if c.Hidden then [ for i in c.Index - 1 .. c.MaxIndex - 1 -> i, 0.0 ]
                    else
                        match c.Width with
                        | Some width -> [ for i in c.Index - 1 .. c.MaxIndex - 1 -> i, width ]
                        | None -> [])
              BestFitColumns =
                sheet.Columns
                |> List.filter (fun c -> c.BestFit)
                |> List.collect (fun c -> [ for i in c.Index - 1 .. c.MaxIndex - 1 -> i ])
                |> Set.ofList
              RowHeights =
                sheet.RowSource.Headers
                |> List.choose (fun r ->
                    if r.Hidden then Some (r.Index - 1, 0.0)
                    else r.Height |> Option.map (fun h -> r.Index - 1, h))
              FrozenRows = sheet.FrozenPane |> Option.map (fun p -> p.YSplit) |> Option.defaultValue 0
              FrozenCols = sheet.FrozenPane |> Option.map (fun p -> p.XSplit) |> Option.defaultValue 0
              ShowGridLines = sheet.ShowGridLines
              View = { ShowZeros = sheet.ShowZeros; ShowHeaders = sheet.View.ShowRowColHeaders; ShowFormulas = sheet.View.ShowFormulas; ZoomScale = sheet.View.ZoomScale; TabColor = sheet.View.TabColor
                       SummaryBelow = sheet.View.SummaryBelow; SummaryRight = sheet.View.SummaryRight; ShowOutlineSymbols = sheet.View.ShowOutlineSymbols }
              Tables =
                sheet.Tables
                |> List.map (fun t ->
                    { Name = t.Name
                      StartRow = t.StartRow
                      StartCol = t.StartCol
                      EndRow = t.EndRow
                      EndCol = t.EndCol
                      Columns = t.Columns
                      HeaderRowCount = t.HeaderRowCount
                      TotalsRowCount = t.TotalsRowCount
                      Totals = t.Totals
                      AutoFilter = t.AutoFilter
                      Style =
                        t.Style |> Option.map (fun s ->
                            { Name = s.Name
                              ShowFirstColumn = s.ShowFirstColumn
                              ShowLastColumn = s.ShowLastColumn
                              ShowRowStripes = s.ShowRowStripes
                              ShowColumnStripes = s.ShowColumnStripes })
                      SourcePart = sheet.Extras.TableParts |> List.tryPick (fun (name, path) -> if name = t.Name then Some path else None) })
              DataValidations =
                sheet.DataValidations
                |> List.map (fun dv ->
                    { Sqref = dv.Range
                      ValidationType = dv.ValidationType
                      Operator = dv.Operator
                      Formula1 = dv.Formula1
                      Formula2 = dv.Formula2
                      AllowBlank = dv.AllowBlank
                      ShowDropDown = dv.ShowDropDown
                      ErrorMessage = dv.ErrorMessage
                      InputMessage = dv.Prompt
                      ErrorTitle = dv.ErrorTitle
                      PromptTitle = dv.PromptTitle
                      ErrorStyle = dv.ErrorStyle
                      ShowErrorMessage = dv.ShowErrorMessage
                      ShowInputMessage = dv.ShowInputMessage })
              ConditionalFormats = sheet.ConditionalFormats
              Hyperlinks = sheet.Hyperlinks |> List.map (fun h -> { Range = h.Range; Target = h.Target; IsInternal = h.IsInternal; Tooltip = h.Tooltip; Display = h.Display })
              PageMargins =
                sheet.PageMargins |> Option.map (fun m -> { Top = m.Top; Bottom = m.Bottom; Left = m.Left; Right = m.Right; Header = m.Header; Footer = m.Footer })
              PageSetup =
                sheet.PageSetup |> Option.map (fun p ->
                    { Orientation =
                        match p.Orientation with
                        | XlsxReader.OrientationLandscape -> Some "landscape"
                        | XlsxReader.OrientationPortrait -> Some "portrait"
                        | _ -> None
                      FitToPage = p.FitToPage
                      FitToWidth = p.FitToWidth
                      FitToHeight = p.FitToHeight
                      Scale = p.Scale
                      PaperSize = p.PaperSize
                      FirstPageNumber = p.FirstPageNumber
                      UseFirstPageNumber = p.UseFirstPageNumber
                      PageOrder = match p.PageOrder with XlsxReader.OverThenDown -> Some "overThenDown" | _ -> None
                      BlackAndWhite = p.BlackAndWhite
                      CellComments =
                        match p.CellComments with
                        | XlsxReader.CellCommentsAsDisplayed -> Some "asDisplayed"
                        | XlsxReader.CellCommentsAtEnd -> Some "atEnd"
                        | _ -> None
                      Kept = p.Kept })
              PrintOptions = Some { GridLines = sheet.PrintOptions.GridLines; Headings = sheet.PrintOptions.Headings; HorizontalCentered = sheet.PrintOptions.HorizontalCentered; VerticalCentered = sheet.PrintOptions.VerticalCentered }
              HeaderFooter =
                Some { OddHeader = sheet.HeaderFooter.OddHeader; OddFooter = sheet.HeaderFooter.OddFooter; EvenHeader = sheet.HeaderFooter.EvenHeader; EvenFooter = sheet.HeaderFooter.EvenFooter; FirstHeader = sheet.HeaderFooter.FirstHeader; FirstFooter = sheet.HeaderFooter.FirstFooter; DifferentOddEven = sheet.HeaderFooter.DifferentOddEven; DifferentFirst = sheet.HeaderFooter.DifferentFirst; ScaleWithDoc = sheet.HeaderFooter.ScaleWithDoc }
              RowBreaks = sheet.RowBreaks
              ColBreaks = sheet.ColBreaks
              PrintArea = sheet.PrintArea
              PrintTitleRows = sheet.PrintTitleRows
              PrintTitleCols = sheet.PrintTitleCols
              Extras =
                { sheet.Extras with
                    Format = Some sheet.Format
                    Columns = sheet.Columns
                    ArrayFormulas =
                        [ for cell in sheet.FormulaCells do
                            match cell.Formula, cell.ArrayRef with
                            | Some _, Some range when range <> "" -> yield cell.Reference.Column + string cell.Reference.Row, range, cell.ArrayDynamic
                            | _ -> () ]
                    UncachedFormulas =
                        Some (sheet.FormulaCells |> List.exists (fun cell -> match cell.Value with XlsxReader.Formula _ -> true | _ -> false))
                    RowAxis =
                        sheet.RowSource.Headers
                        |> List.choose (fun row ->
                            if row.Hidden || row.OutlineLevel > 0 || row.Collapsed then
                                Some (row.Index - 1, (if row.Hidden then row.Height else None), row.OutlineLevel, row.Collapsed)
                            else None)
                    ColumnAxis =
                        sheet.Columns
                        |> List.collect (fun column ->
                            if column.Hidden || column.OutlineLevel > 0 || column.Collapsed then
                                [ for index in column.Index - 1 .. min 16383 (column.MaxIndex - 1) ->
                                    index, (if column.Hidden then column.Width else None), column.OutlineLevel, column.Collapsed ]
                            else [])
                    ColumnStyles =
                        sheet.Columns
                        |> List.choose (fun column ->
                            match column.Style with
                            | Some index when index > 0 && index < workbook.Styles.Length -> Some (column.Index - 1, column.MaxIndex - 1, workbook.Styles.[index])
                            | _ -> None) } })
      ActiveSheetIndex = max 0 (min (workbook.Sheets.Length - 1) workbook.Preserved.ActiveTab)
      DefinedNames = workbook.NamedRanges |> List.map (fun n -> { Name = n.Name; RefersTo = n.RefersTo; LocalSheetIndex = n.Scope |> Option.bind (fun s -> match System.Int32.TryParse s with true, v -> Some v | _ -> None); Comment = n.Comment; Hidden = n.Hidden })
      VbaProjectSource = None
      VbaProjectBytes = None
      Dxfs = workbook.Dxfs |> Array.toList
      Preserved = workbook.Preserved }
