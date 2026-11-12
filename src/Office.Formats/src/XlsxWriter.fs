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
let private numberText (value: float) : string =
#if FABLE_COMPILER
    string value
#else
    value.ToString("R", System.Globalization.CultureInfo.InvariantCulture)
#endif

/// Characters XML 1.0 cannot carry at all (C0 controls other than tab, LF,
/// CR, and the two non-characters). Emitting them makes the part unreadable.
let private isXmlIllegal (ch: char) =
    (ch < ' ' && ch <> '\t' && ch <> '\n' && ch <> '\r') || ch = '\uFFFE' || ch = '\uFFFF'

/// Escape for element text. Illegal characters are dropped; a carriage
/// return is written as a character reference so line-end normalisation
/// does not turn it into a line feed.
let private appendEscapedChar (sb: StringBuilder) (ch: char) =
    match ch with
    | '&' -> sb.Append("&amp;") |> ignore
    | '<' -> sb.Append("&lt;") |> ignore
    | '>' -> sb.Append("&gt;") |> ignore
    | '"' -> sb.Append("&quot;") |> ignore
    | '\'' -> sb.Append("&apos;") |> ignore
    | '\r' -> sb.Append("&#13;") |> ignore
    | _ -> if not (isXmlIllegal ch) then sb.Append(ch) |> ignore

let private appendEscaped (sb: StringBuilder) (text: string) =
    for ch in text do
        appendEscapedChar sb ch

/// Escape for an attribute value: as `appendEscaped`, and tab / line feed
/// become character references too, since attribute-value normalisation
/// would otherwise turn them into spaces (multi-line validation messages).
let private appendEscapedAttr (sb: StringBuilder) (text: string) =
    for ch in text do
        match ch with
        | '\t' -> sb.Append("&#9;") |> ignore
        | '\n' -> sb.Append("&#10;") |> ignore
        | _ -> appendEscapedChar sb ch

let private hexDigits = "0123456789ABCDEF"

let private appendXEscape (sb: StringBuilder) (ch: char) =
    let code = int ch
    sb.Append("_x")
        .Append(hexDigits.[(code >>> 12) &&& 15])
        .Append(hexDigits.[(code >>> 8) &&& 15])
        .Append(hexDigits.[(code >>> 4) &&& 15])
        .Append(hexDigits.[code &&& 15])
        .Append('_') |> ignore

let private isHexDigit (c: char) =
    (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')

/// Escape cell text (ST_Xstring: shared strings, inline strings, cached
/// string results, comments). Characters XML cannot carry — and a carriage
/// return — are written as `_xHHHH_`, the way Excel stores them, and an
/// underscore that would itself read as such an escape is written `_x005F_`.
/// The reader decodes both.
let private appendCellText (sb: StringBuilder) (text: string) =
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

let private appendCellTextElement (sb: StringBuilder) (name: string) (text: string) =
    sb.Append('<').Append(name) |> ignore
    if text = "" then
        sb.Append("/>") |> ignore
    else
        // Leading / trailing whitespace is only kept with xml:space="preserve".
        sb.Append(" xml:space=\"preserve\">") |> ignore
        appendCellText sb text
        sb.Append("</").Append(name).Append('>') |> ignore

let private appendTextElement (sb: StringBuilder) (name: string) (text: string) =
    sb.Append('<').Append(name) |> ignore
    if text = "" then
        sb.Append("/>") |> ignore
    else
        sb.Append('>') |> ignore
        appendEscaped sb text
        sb.Append("</").Append(name).Append('>') |> ignore

let private appendCellRef (sb: StringBuilder) (colLetters: string) (rowNumber: string) =
    sb.Append(colLetters).Append(rowNumber) |> ignore

let private appendSimpleCell (sb: StringBuilder) (rowNumber: string) (colLetters: string) (value: string) =
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

type private FontKey = {
    FontName: string option
    FontSize: float option
    Bold: bool
    Italic: bool
    Underline: bool
    UnderlineDouble: bool
    Strikethrough: bool
    TextColor: string option
}

type private FillKey = string option

type private BorderKey = {
    Top: XlsxBorderSide option
    Bottom: XlsxBorderSide option
    Left: XlsxBorderSide option
    Right: XlsxBorderSide option
    Diagonal: XlsxBorderSide option
    DiagonalUp: bool
    DiagonalDown: bool
}

type private XfKey = {
    NumFmt: string option
    Font: FontKey
    Fill: FillKey
    Border: BorderKey
    HorizontalAlign: string option
    VerticalAlign: string option
    WrapText: bool
    Indent: int
    ShrinkToFit: bool
    TextRotation: int
    /// Index of the named cell style (cellStyleXfs entry) the xf is based on.
    NamedStyle: int
}

type private SharedStringKey = {
    Text: string
    Runs: XlsxRichRun list
}

type private StyleBook = {
    /// Named cell styles kept from the opened workbook.
    NamedStyles: XlsxReader.NamedCellStyle list
    Dxfs: XlsxReader.Dxf list
    Fonts: FontKey list
    Fills: FillKey list
    Borders: BorderKey list
    Xfs: XfKey list
    NumFmtIds: Map<string, int>
    StyleIndex: Map<XfKey, int>
    /// The cellXfs index of a format given as its key (a column's default style).
    KeyIndex: XfKey -> int option
    /// The cellXfs index of a cell format. Formats are looked up by identity
    /// first: the cells of a workbook share a few format records.
    StyleIdOf: XlsxCellFormat -> int
}

let private invariant (value: float) = numberText value

let private invariantFixed (value: float) =
    invariant value

let private attr (name: string) (value: string) =
    let sb = StringBuilder()
    sb.Append(name).Append("=\"") |> ignore
    appendEscapedAttr sb value
    sb.Append('"') |> ignore
    sb.ToString()

let private escaped (value: string) =
    let sb = StringBuilder()
    appendEscaped sb value
    sb.ToString()

let private escapedCellText (value: string) =
    let sb = StringBuilder()
    appendCellText sb value
    sb.ToString()

let private elemText name text =
    let sb = StringBuilder()
    appendTextElement sb name text
    sb.ToString()

let private hasText value =
    not (System.String.IsNullOrEmpty value)

let private cleanRgb (color: string) =
    let c = color.TrimStart('#').ToUpperInvariant()
    if c.Length = 8 then c
    elif c.Length = 6 then "FF" + c
    else c

let private xmlColor color =
    color |> Option.map cleanRgb

let private fontKey (fmt: XlsxCellFormat) =
    { FontName = fmt.FontName
      FontSize = fmt.FontSize
      Bold = fmt.Bold
      Italic = fmt.Italic
      Underline = fmt.Underline
      UnderlineDouble = fmt.UnderlineDouble
      Strikethrough = fmt.Strikethrough
      TextColor = fmt.TextColor |> xmlColor }

let private borderKey (fmt: XlsxCellFormat) =
    let cleanSide (side: XlsxBorderSide option) =
        side |> Option.map (fun (s: XlsxBorderSide) -> { s with Color = s.Color |> xmlColor })
    { Top = cleanSide fmt.BorderTop
      Bottom = cleanSide fmt.BorderBottom
      Left = cleanSide fmt.BorderLeft
      Right = cleanSide fmt.BorderRight
      Diagonal = cleanSide fmt.BorderDiagonal
      DiagonalUp = fmt.DiagonalUp
      DiagonalDown = fmt.DiagonalDown }

let private xfKey (fmt: XlsxCellFormat) =
    { NumFmt = fmt.NumberFormat
      Font = fontKey fmt
      Fill = fmt.FillColor |> xmlColor
      Border = borderKey fmt
      HorizontalAlign = fmt.HorizontalAlign
      VerticalAlign = fmt.VerticalAlign
      WrapText = fmt.WrapText
      Indent = fmt.Indent
      ShrinkToFit = fmt.ShrinkToFit
      TextRotation = fmt.TextRotation
      NamedStyle = fmt.NamedStyle |> Option.defaultValue 0 |> max 0 }

let private defaultFontKey = fontKey defaultXlsxCellFormat
let private defaultBorderKey = borderKey defaultXlsxCellFormat
let private defaultXfKey = xfKey defaultXlsxCellFormat

let private builtInNumFmtId (code: string) =
    match code with
    | "General" -> Some 0
    | "0" -> Some 1
    | "0.00" -> Some 2
    | "#,##0" -> Some 3
    | "#,##0.00" -> Some 4
    | "0%" -> Some 9
    | "0.00%" -> Some 10
    | "0.00E+00" -> Some 11
    | "# ?/?" -> Some 12
    | "# ??/??" -> Some 13
    | "m/d/yyyy" -> Some 14
    | "d-mmm-yy" -> Some 15
    | "d-mmm" -> Some 16
    | "mmm-yy" -> Some 17
    | "h:mm AM/PM" -> Some 18
    | "h:mm:ss AM/PM" -> Some 19
    | "h:mm" -> Some 20
    | "h:mm:ss" -> Some 21
    | "m/d/yyyy h:mm" -> Some 22
    | "#,##0;(#,##0)" -> Some 37
    | "#,##0;[Red](#,##0)" -> Some 38
    | "#,##0.00;(#,##0.00)" -> Some 39
    | "#,##0.00;[Red](#,##0.00)" -> Some 40
    | "mm:ss" -> Some 45
    | "[h]:mm:ss" -> Some 46
    | "mmss.0" -> Some 47
    | "##0.0E+0" -> Some 48
    | "@" -> Some 49
    | _ -> None

let private appendUnique item items =
    if List.contains item items then items else items @ [ item ]

let private columnStyleKey (style: XlsxReader.CellStyle) =
    let side (b: XlsxReader.CellBorderSide option) = b |> Option.map (fun x -> { Style = x.Style; Color = x.Color })
    xfKey
        { defaultXlsxCellFormat with
            Bold = style.Bold
            Italic = style.Italic
            Underline = style.Underline
            UnderlineDouble = style.UnderlineDouble
            Strikethrough = style.Strikethrough
            FontName = style.FontName
            TextColor = style.FontColor |> Option.map (fun c -> "#" + c.TrimStart('#'))
            FillColor = style.FillColor |> Option.map (fun c -> "#" + c.TrimStart('#'))
            FontSize = style.FontSize
            NumberFormat = style.NumberFormat
            HorizontalAlign = style.HorizontalAlign
            VerticalAlign = style.VerticalAlign
            WrapText = style.WrapText
            Indent = style.Indent
            ShrinkToFit = style.ShrinkToFit
            TextRotation = style.TextRotation
            BorderTop = side style.BorderTop
            BorderBottom = side style.BorderBottom
            BorderLeft = side style.BorderLeft
            BorderRight = side style.BorderRight
            BorderDiagonal = side style.BorderDiagonal
            DiagonalUp = style.DiagonalUp
            DiagonalDown = style.DiagonalDown }

let private collectStyles (workbook: XlsxWorkbookOut) =
    // One pass over the cells. A format record seen before (the same object)
    // costs a lookup by identity; only a new one is turned into its key and
    // compared by content. The order of first appearance gives the indexes.
    let xfList = ResizeArray<XfKey>()
    let xfIndex = System.Collections.Generic.Dictionary<XfKey, int>(HashIdentity.Structural)
    let byIdentity = System.Collections.Generic.Dictionary<XlsxCellFormat, int>(HashIdentity.Reference)
    let indexOfKey (key: XfKey) =
        match xfIndex.TryGetValue key with
        | true, index -> index
        | false, _ ->
            let index = xfList.Count
            xfList.Add key
            xfIndex.[key] <- index
            index
    let styleIdOf (format: XlsxCellFormat) =
        match byIdentity.TryGetValue format with
        | true, index -> index
        | false, _ ->
            let index = indexOfKey (xfKey format)
            byIdentity.[format] <- index
            index
    indexOfKey defaultXfKey |> ignore
    for sheet in workbook.Sheets do
        for cell in sheet.Cells do
            styleIdOf cell.Format |> ignore
        // Default cell styles of columns (<col style>).
        for _, _, style in sheet.Extras.ColumnStyles do
            indexOfKey (columnStyleKey style) |> ignore
    let xfs = List.ofSeq xfList
    let fonts =
        xfs |> List.map (fun x -> x.Font) |> List.fold (fun acc key -> appendUnique key acc) [ defaultFontKey ]
    // Custom fills only. Fill 0 (none) and fill 1 (gray125) are reserved by
    // Excel whatever the file says, so custom fills are numbered from 2.
    let fills =
        xfs |> List.choose (fun x -> x.Fill) |> List.fold (fun acc key -> appendUnique (Some key) acc) []
    let borders =
        xfs |> List.map (fun x -> x.Border) |> List.fold (fun acc key -> appendUnique key acc) [ defaultBorderKey ]
    let customNumFmts =
        (xfs |> List.choose (fun x -> x.NumFmt))
        @ (workbook.Preserved.NamedStyles |> List.choose (fun style -> style.NumFmtCode))
        |> List.distinct
        |> List.filter (fun code -> builtInNumFmtId code |> Option.isNone)
        |> List.mapi (fun i code -> code, 164 + i)
        |> Map.ofList
    let styleIndex = xfs |> List.mapi (fun i x -> x, i) |> Map.ofList
    { NamedStyles = workbook.Preserved.NamedStyles
      Dxfs = workbook.Dxfs
      Fonts = fonts
      Fills = fills
      Borders = borders
      Xfs = xfs
      NumFmtIds = customNumFmts
      StyleIndex = styleIndex
      KeyIndex = (fun key -> Map.tryFind key styleIndex)
      StyleIdOf = styleIdOf }

/// The shared strings of the workbook in order of first use, and the index
/// of a text (with its runs). Plain texts - nearly all of them - are found
/// by the string alone.
let private collectSharedStrings (workbook: XlsxWorkbookOut) : SharedStringKey list * (string -> XlsxRichRun list -> int) =
    let items = ResizeArray<SharedStringKey>()
    let plain = System.Collections.Generic.Dictionary<string, int>()
    let rich = System.Collections.Generic.Dictionary<SharedStringKey, int>(HashIdentity.Structural)
    let indexOf (text: string) (runs: XlsxRichRun list) : int =
        if runs.IsEmpty then
            match plain.TryGetValue text with
            | true, index -> index
            | false, _ ->
                let index = items.Count
                items.Add { Text = text; Runs = [] }
                plain.[text] <- index
                index
        else
            let key = { Text = text; Runs = runs }
            match rich.TryGetValue key with
            | true, index -> index
            | false, _ ->
                let index = items.Count
                items.Add key
                rich.[key] <- index
                index
    for sheet in workbook.Sheets do
        for cell in sheet.Cells do
            match cell.Value with
            | XlsxText (text, runs) -> indexOf text runs |> ignore
            | XlsxFormula _ -> ()   // (a formula's text result is written in its cell: t="str")
            | _ -> ()
    List.ofSeq items, indexOf

let private fontXml (font: FontKey) =
    let children = [
        if font.Bold then yield "<b/>"
        if font.Italic then yield "<i/>"
        if font.Strikethrough then yield "<strike/>"
        if font.Underline then
            if font.UnderlineDouble then yield "<u val=\"double\"/>" else yield "<u/>"
        match font.FontSize with
        | Some size -> yield sprintf "<sz val=\"%s\"/>" (invariant size)
        | None -> yield "<sz val=\"11\"/>"
        match font.TextColor with
        | Some c -> yield sprintf "<color rgb=\"%s\"/>" c
        | None -> ()
        yield sprintf "<name %s/>" (attr "val" (font.FontName |> Option.defaultValue "Calibri"))
    ]
    "<font>" + (String.concat "" children) + "</font>"

let private fillXml (fill: FillKey) =
    match fill with
    | None -> "<fill><patternFill patternType=\"none\"/></fill>"
    | Some color -> sprintf "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"%s\"/><bgColor indexed=\"64\"/></patternFill></fill>" color

let private borderSideXml name (side: XlsxBorderSide option) =
    match side with
    | None -> sprintf "<%s/>" name
    | Some s ->
        let color = s.Color |> Option.map (fun c -> sprintf "<color rgb=\"%s\"/>" c) |> Option.defaultValue ""
        sprintf "<%s %s>%s</%s>" name (attr "style" s.Style) color name

let private borderXml (border: BorderKey) =
    let attrs = [
        if border.DiagonalUp then yield "diagonalUp=\"1\""
        if border.DiagonalDown then yield "diagonalDown=\"1\""
    ]
    let openTag =
        if List.isEmpty attrs then "<border>" else "<border " + String.concat " " attrs + ">"
    openTag
    + borderSideXml "left" border.Left
    + borderSideXml "right" border.Right
    + borderSideXml "top" border.Top
    + borderSideXml "bottom" border.Bottom
    + borderSideXml "diagonal" border.Diagonal
    + "</border>"

let private indexOf item items =
    items |> List.findIndex ((=) item)

/// A differential format: `<dxf>` of the styles part, `<x14:dxf>` of an x14 rule.
let private dxfXml (element: string) (d: XlsxReader.Dxf) =
    let flag name value =
        match value with
        | Some true -> sprintf "<%s/>" name
        | Some false -> sprintf "<%s val=\"0\"/>" name
        | None -> ""
    let font =
        let body =
            flag "b" d.Bold + flag "i" d.Italic + flag "strike" d.Strike
            + (match d.Underline with Some true -> "<u/>" | Some false -> "<u val=\"none\"/>" | None -> "")
            + (d.FontColor |> xmlColor |> Option.map (fun c -> sprintf "<color rgb=\"%s\"/>" c) |> Option.defaultValue "")
        if body = "" then "" else "<font>" + body + "</font>"
    let fill =
        d.FillColor
        |> xmlColor
        |> Option.map (fun c -> sprintf "<fill><patternFill><bgColor rgb=\"%s\"/></patternFill></fill>" c)
        |> Option.defaultValue ""
    "<" + element + ">" + font + fill + "</" + element + ">"

/// A cellXfs entry: the format with the ids of its number format, font,
/// fill, border and the named cell style it is based on.
let private xfXmlOf (numFmtId: int) (fontId: int) (fillId: int) (borderId: int) (xfId: int) (xf: XfKey) =
    let alignmentAttrs = [
        match xf.HorizontalAlign with Some v -> yield attr "horizontal" v | None -> ()
        match xf.VerticalAlign with Some v -> yield attr "vertical" v | None -> ()
        if xf.WrapText then yield "wrapText=\"1\""
        if xf.Indent > 0 then yield sprintf "indent=\"%d\"" xf.Indent
        if xf.ShrinkToFit then yield "shrinkToFit=\"1\""
        if xf.TextRotation <> 0 then yield sprintf "textRotation=\"%d\"" xf.TextRotation
    ]
    let applyAlignment = not (List.isEmpty alignmentAttrs)
    let attrs = [
        yield sprintf "numFmtId=\"%d\"" numFmtId
        yield sprintf "fontId=\"%d\"" fontId
        yield sprintf "fillId=\"%d\"" fillId
        yield sprintf "borderId=\"%d\"" borderId
        // The named cell style the format is based on (Normal when the workbook has no such style).
        yield sprintf "xfId=\"%d\"" xfId
        if xf.NumFmt.IsSome then yield "applyNumberFormat=\"1\""
        if fontId <> 0 then yield "applyFont=\"1\""
        if fillId <> 0 then yield "applyFill=\"1\""
        if borderId <> 0 then yield "applyBorder=\"1\""
        if applyAlignment then yield "applyAlignment=\"1\""
    ]
    if applyAlignment then
        sprintf "<xf %s><alignment %s/></xf>" (String.concat " " attrs) (String.concat " " alignmentAttrs)
    else
        sprintf "<xf %s/>" (String.concat " " attrs)

let private stylesXml (book: StyleBook) =
    let numFmts =
        if Map.isEmpty book.NumFmtIds then ""
        else
            let items =
                book.NumFmtIds
                |> Map.toList
                |> List.sortBy snd
                |> List.map (fun (code, id) -> sprintf "<numFmt numFmtId=\"%d\" formatCode=\"%s\"/>" id (escaped code))
                |> String.concat ""
            sprintf "<numFmts count=\"%d\">%s</numFmts>" (Map.count book.NumFmtIds) items
    // The fonts, fills and borders of the kept named styles follow the
    // workbook's own in each table; a named style's xf refers to them there.
    let named = book.NamedStyles |> List.toArray
    let namedSlot (pick: XlsxReader.NamedCellStyle -> string option) (ownCount: int) =
        let mutable next = ownCount
        named
        |> Array.map (fun style ->
            match pick style with
            | Some _ ->
                let id = next
                next <- next + 1
                id
            | None -> 0)
    let namedFontIds = namedSlot (fun style -> style.FontXml) book.Fonts.Length
    let namedFillIds = namedSlot (fun style -> style.FillXml) (2 + book.Fills.Length)
    let namedBorderIds = namedSlot (fun style -> style.BorderXml) book.Borders.Length
    let namedXml (pick: XlsxReader.NamedCellStyle -> string option) = named |> Array.choose pick |> String.concat ""
    let namedCount (pick: XlsxReader.NamedCellStyle -> string option) = named |> Array.filter (fun style -> (pick style).IsSome) |> Array.length
    let fonts = (book.Fonts |> List.map fontXml |> String.concat "") + namedXml (fun style -> style.FontXml)
    let fills =
        let body =
            [ yield fillXml None
              yield "<fill><patternFill patternType=\"gray125\"/></fill>"
              for fill in book.Fills -> fillXml fill
              for style in named do
                  match style.FillXml with
                  | Some xml -> yield xml
                  | None -> () ]
        sprintf "<fills count=\"%d\">%s</fills>" body.Length (String.concat "" body)
    let fillIdOf (fill: FillKey) =
        match fill with
        | None -> 0
        | Some _ -> 2 + indexOf fill book.Fills
    let dxfs =
        sprintf "<dxfs count=\"%d\">%s</dxfs>" book.Dxfs.Length (book.Dxfs |> List.map (dxfXml "dxf") |> String.concat "")
    let borders = (book.Borders |> List.map borderXml |> String.concat "") + namedXml (fun style -> style.BorderXml)
    // cellStyleXfs / cellStyles: the kept named styles, or just Normal.
    let cellStyleXfs, cellStyles =
        if named.Length = 0 then
            "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>",
            "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        else
            let attributes (pairs: (string * string) list) = pairs |> List.map (fun (key, value) -> " " + attr key value) |> String.concat ""
            let xfs =
                named
                |> Array.mapi (fun i style ->
                    let numFmtId =
                        match style.NumFmtCode with
                        | Some code -> builtInNumFmtId code |> Option.orElseWith (fun () -> Map.tryFind code book.NumFmtIds) |> Option.defaultValue 0
                        | None -> style.NumFmtId
                    let ids = sprintf "numFmtId=\"%d\" fontId=\"%d\" fillId=\"%d\" borderId=\"%d\"" numFmtId namedFontIds.[i] namedFillIds.[i] namedBorderIds.[i]
                    if style.XfChildrenXml = "" then sprintf "<xf %s%s/>" ids (attributes style.XfAttributes)
                    else sprintf "<xf %s%s>%s</xf>" ids (attributes style.XfAttributes) style.XfChildrenXml)
                |> String.concat ""
            let styles =
                named
                |> Array.mapi (fun i style -> sprintf "<cellStyle %s xfId=\"%d\"%s/>" (attr "name" style.Name) i (attributes style.StyleAttributes))
                |> String.concat ""
            sprintf "<cellStyleXfs count=\"%d\">%s</cellStyleXfs>" named.Length xfs,
            sprintf "<cellStyles count=\"%d\">%s</cellStyles>" named.Length styles
    let xfXml (xf: XfKey) =
        let numFmtId =
            match xf.NumFmt with
            | Some code ->
                builtInNumFmtId code
                |> Option.orElseWith (fun () -> Map.tryFind code book.NumFmtIds)
                |> Option.defaultValue 0
            | None -> 0
        xfXmlOf numFmtId (indexOf xf.Font book.Fonts) (fillIdOf xf.Fill) (indexOf xf.Border book.Borders)
            (if xf.NamedStyle > 0 && xf.NamedStyle < named.Length then xf.NamedStyle else 0) xf
    let xfs = book.Xfs |> List.map xfXml |> String.concat ""
    declaration + "\n"
    + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    + numFmts
    + sprintf "<fonts count=\"%d\">%s</fonts>" (book.Fonts.Length + namedCount (fun style -> style.FontXml)) fonts
    + fills
    + sprintf "<borders count=\"%d\">%s</borders>" (book.Borders.Length + namedCount (fun style -> style.BorderXml)) borders
    + cellStyleXfs
    + sprintf "<cellXfs count=\"%d\">%s</cellXfs>" book.Xfs.Length xfs
    + cellStyles
    + dxfs + "<tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>"
    + "</styleSheet>"

let private richRunXml (run: XlsxRichRun) =
    let props = [
        match run.Bold with Some true -> yield "<b/>" | Some false -> yield "<b val=\"0\"/>" | None -> ()
        match run.Italic with Some true -> yield "<i/>" | Some false -> yield "<i val=\"0\"/>" | None -> ()
        match run.Strikethrough with Some true -> yield "<strike/>" | Some false -> yield "<strike val=\"0\"/>" | None -> ()
        match run.Underline with
        | Some true ->
            match run.UnderlineDouble with
            | Some true -> yield "<u val=\"double\"/>"
            | _ -> yield "<u/>"
        | Some false -> yield "<u val=\"none\"/>"
        | None -> ()
        match run.Color |> xmlColor with Some c -> yield sprintf "<color rgb=\"%s\"/>" c | None -> ()
        match run.FontName with Some n -> yield sprintf "<rFont %s/>" (attr "val" n) | None -> ()
        match run.FontSize with Some s -> yield sprintf "<sz val=\"%s\"/>" (invariant s) | None -> ()
        match run.VertAlign with
        | Some XlsxSuperscript -> yield "<vertAlign val=\"superscript\"/>"
        | Some XlsxSubscript -> yield "<vertAlign val=\"subscript\"/>"
        | _ -> ()
    ]
    let pr = if List.isEmpty props then "" else "<rPr>" + String.concat "" props + "</rPr>"
    sprintf "<r>%s<t xml:space=\"preserve\">%s</t></r>" pr (escapedCellText run.Text)

let private appendSharedStringItem (sb: StringBuilder) (item: SharedStringKey) =
    if List.isEmpty item.Runs then
        sb.Append("<si><t xml:space=\"preserve\">") |> ignore
        appendCellText sb item.Text
        sb.Append("</t></si>") |> ignore
    else
        sb.Append("<si>") |> ignore
        for run in item.Runs do
            sb.Append(richRunXml run) |> ignore
        sb.Append("</si>") |> ignore

let private sharedStringsXmlPreserved (items: SharedStringKey list) =
    let sb = StringBuilder()
    let count = string items.Length
    sb.Append(declaration).Append("\n<sst xmlns=\"").Append(Namespaces.spreadsheet).Append("\" count=\"").Append(count).Append("\" uniqueCount=\"").Append(count).Append("\">") |> ignore
    for item in items do appendSharedStringItem sb item
    sb.Append("</sst>").ToString()

/// The values of `t="e"` that Excel reads.
let private isFileError (text: string) =
    match text with
    | "#NULL!" | "#DIV/0!" | "#VALUE!" | "#REF!" | "#NAME?" | "#NUM!" | "#N/A" | "#GETTING_DATA"
    | "#SPILL!" | "#CALC!" | "#FIELD!" | "#BLOCKED!" | "#CONNECT!" | "#UNKNOWN!" | "#BUSY!" -> true
    | _ -> false

/// ` t="..."` of a cell of that type ("" for a number).
let private typeAttr (cellType: string) = if cellType = "" then "" else " t=\"" + cellType + "\""

/// The type of a cell that holds the value (`t`: "" for a number) and its `<v>`.
let private cachedValueXml (sharedIndex: string -> XlsxRichRun list -> int) value =
    let rec loop v =
        match v with
        | XlsxBlank -> "", ""
        | XlsxNumber n -> "", "<v>" + invariant n + "</v>"
        | XlsxBoolean b -> "b", (if b then "<v>1</v>" else "<v>0</v>")
        // (What only the engine says - a cycle - is no error of the file format: Excel refuses the file.)
        | XlsxError e when not (isFileError e) -> "", ""
        | XlsxError e -> "e", "<v>" + escaped e + "</v>"
        | XlsxText (text, runs) -> "s", "<v>" + string (sharedIndex text runs) + "</v>"
        | XlsxFormula (_, cached) ->
            cached |> Option.map loop |> Option.defaultValue ("", "")
    loop value

/// The result a formula cell holds: a text is written in the cell itself
/// (`t="str"`), as Excel does - the shared strings are for texts typed into cells.
let private formulaResultXml (sharedIndex: string -> XlsxRichRun list -> int) value =
    match value with
    | XlsxText (text, _) -> "str", "<v>" + escapedCellText text + "</v>"
    | other -> cachedValueXml sharedIndex other

/// `<f t="dataTable" .../>` of a what-if data table.
let private dataTableXml (table: XlsxReader.DataTableFormula) =
    let attrs = [
        yield "t=\"dataTable\""
        yield attr "ref" table.Ref
        if table.TwoDimensional then yield "dt2D=\"1\""
        if table.RowInput then yield "dtr=\"1\""
        yield attr "r1" table.Input1
        match table.Input2 with Some input -> yield attr "r2" input | None -> ()
        if table.Input1Deleted then yield "del1=\"1\""
        if table.Input2Deleted then yield "del2=\"1\""
        if table.CalculateAlways then yield "ca=\"1\""
    ]
    "<f " + String.concat " " attrs + "/>"

/// Column letters by index, made once per column.
let private columnLetterCache = System.Collections.Generic.Dictionary<int, string>()
let private columnLetters (col: int) =
    match columnLetterCache.TryGetValue col with
    | true, letters -> letters
    | false, _ ->
        let letters = columnToLetter col
        columnLetterCache.[col] <- letters
        letters

/// A formula as the file format stores it: prefixes added where they are missing (see `FormulaNames`).
let fileFormula (formula: string) : string = FormulaNames.fileFormula formula

/// A formula as the model holds it and Excel shows it (see `FormulaNames`).
let modelFormula (formula: string) : string = FormulaNames.modelFormula formula

/// Text for an element or attribute, escaped only when it has something to escape.
let private escapedFast (value: string) =
    let mutable plain = true
    let mutable i = 0
    while plain && i < value.Length do
        let ch = value.[i]
        if ch = '&' || ch = '<' || ch = '>' || ch = '"' || ch = '\'' || ch < ' ' || ch = '\uFFFE' || ch = '\uFFFF' then plain <- false
        i <- i + 1
    if plain then value else escaped value

/// One `<c>` appended to `sb`. `rowText` is the row number as text.
let private appendCellXml (sb: StringBuilder) (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (dataTables: Map<string, XlsxReader.DataTableFormula>) (arrayFormulas: Map<string, string * string>) (rowText: string) (cell: XlsxCellOut) =
    let styleId = styles.StyleIdOf cell.Format
    let letters = columnLetters cell.Col
    let dataTable = if dataTables.IsEmpty then None else Map.tryFind (letters + rowText) dataTables
    let openCell (typeAttr: string) =
        sb.Append("<c r=\"").Append(letters).Append(rowText).Append('"') |> ignore
        if styleId <> 0 then sb.Append(" s=\"").Append(string styleId).Append('"') |> ignore
        sb.Append(typeAttr) |> ignore
    match dataTable, cell.Value with
    // The first cell of a what-if data table: the table, and the result the cell holds.
    | Some table, value ->
        let cellType, cachedXml = cachedValueXml sharedIndex value
        openCell (typeAttr cellType)
        sb.Append('>').Append(dataTableXml table).Append(cachedXml).Append("</c>") |> ignore
    | None, XlsxBlank ->
        if styleId <> 0 then
            openCell ""
            sb.Append("/>") |> ignore
    | None, XlsxNumber n ->
        openCell ""
        sb.Append("><v>").Append(invariant n).Append("</v></c>") |> ignore
    | None, XlsxText (text, runs) ->
        openCell " t=\"s\""
        sb.Append("><v>").Append(string (sharedIndex text runs)).Append("</v></c>") |> ignore
    | None, XlsxFormula (formula, cached) ->
        let cellType, cachedXml = cached |> Option.map (formulaResultXml sharedIndex) |> Option.defaultValue ("", "")
        let tAttr = typeAttr cellType
        // An array formula says which cells its results fill; a dynamic one
        // names its cell metadata (`cm`).
        match (if arrayFormulas.IsEmpty then None else Map.tryFind (letters + rowText) arrayFormulas) with
        | Some (range, cmAttr) ->
            openCell (tAttr + cmAttr)
            sb.Append("><f t=\"array\" ref=\"").Append(range).Append("\">") |> ignore
        | None ->
            openCell tAttr
            sb.Append("><f>") |> ignore
        sb.Append(escapedFast (fileFormula formula)).Append("</f>").Append(cachedXml).Append("</c>") |> ignore
    | None, other ->
        let cellType, valueXml = cachedValueXml sharedIndex other
        openCell (typeAttr cellType)
        sb.Append('>').Append(valueXml).Append("</c>") |> ignore

let private sheetDimension (sheet: XlsxSheetOut) =
    let mutable any = false
    let mutable minR = System.Int32.MaxValue
    let mutable minC = System.Int32.MaxValue
    let mutable maxR = 0
    let mutable maxC = 0
    let cover row col =
        any <- true
        if row < minR then minR <- row
        if row > maxR then maxR <- row
        if col < minC then minC <- col
        if col > maxC then maxC <- col
    for c in sheet.Cells do cover c.Row c.Col
    for m in sheet.MergedCells do
        cover m.StartRow m.StartCol
        cover m.EndRow m.EndCol
    for t in sheet.Tables do
        cover (t.StartRow - 1) (t.StartCol - 1)
        cover (t.EndRow - 1) (t.EndCol - 1)
    if not any then "A1"
    else cellRef minR minC + ":" + cellRef maxR maxC

/// `<cols>`: widths (0 = hidden), and the default cell style a column range
/// had in the file. Neighbouring columns with the same width and style are
/// written as one `<col min max>`, as Excel does.
let private colsXml (styles: StyleBook) (sheet: XlsxSheetOut) =
    let widths = Map.ofList sheet.ColumnWidths
    let styleIds =
        sheet.Extras.ColumnStyles
        |> List.collect (fun (first, last, style) ->
            match styles.KeyIndex (columnStyleKey style) with
            | Some id when id > 0 -> [ for col in max 0 first .. min 16383 last -> col, id ]
            | _ -> [])
        |> Map.ofList
    // A hidden column keeps the width the file gave it.
    let fileWidth col =
        sheet.Extras.Columns
        |> List.tryFind (fun column -> col + 1 >= column.Index && col + 1 <= column.MaxIndex)
        |> Option.bind (fun column -> column.Width)
    // The width a hidden column keeps, its outline level and collapsed mark.
    let axis = sheet.Extras.ColumnAxis |> List.map (fun (col, kept, level, collapsed) -> col, (kept, level, collapsed)) |> Map.ofList
    let columns =
        // (Column 16384 is the last: one moved past it by an insertion is gone, as in Excel.)
        Set.unionMany [ widths |> Map.toSeq |> Seq.map fst |> Set.ofSeq; styleIds |> Map.toSeq |> Seq.map fst |> Set.ofSeq
                        axis |> Map.toSeq |> Seq.filter (fun (_, (_, level, collapsed)) -> level > 0 || collapsed) |> Seq.map fst |> Set.ofSeq ]
        |> Set.filter (fun col -> col < 16384)
        |> Set.toList
    if List.isEmpty columns then ""
    else
        let attributes col =
            let width, hidden =
                match Map.tryFind col widths with
                | Some width when width <= 0.0 ->
                    (match Map.tryFind col axis with
                     | Some (Some kept, _, _) -> Some kept
                     | _ -> fileWidth col), true
                | Some width -> Some width, false
                | None -> None, false
            let parts = [
                match width with
                | Some value -> yield sprintf "width=\"%s\" customWidth=\"1\"" (invariantFixed value)
                | None -> ()
                match Map.tryFind col styleIds with
                | Some id -> yield sprintf "style=\"%d\"" id
                | None -> ()
                if hidden then yield "hidden=\"1\""
                match Map.tryFind col axis with
                | Some (_, level, collapsed) ->
                    if level > 0 then yield sprintf "outlineLevel=\"%d\"" level
                    if collapsed then yield "collapsed=\"1\""
                | None -> ()
            ]
            String.concat " " parts
        let body = StringBuilder()
        let flush (first: int) (last: int) (attrs: string) =
            body.Append(sprintf "<col min=\"%d\" max=\"%d\"%s/>" (first + 1) (last + 1) (if attrs = "" then "" else " " + attrs)) |> ignore
        let mutable start = List.head columns
        let mutable previous = start
        let mutable current = attributes start
        for col in List.tail columns do
            let attrs = attributes col
            if col = previous + 1 && attrs = current then
                previous <- col
            else
                flush start previous current
                start <- col
                previous <- col
                current <- attrs
        flush start previous current
        "<cols>" + body.ToString() + "</cols>"

let private sheetViewsXml (sheet: XlsxSheetOut) active =
    let attrs = [
        if sheet.View.ShowFormulas then yield "showFormulas=\"1\""
        if not sheet.ShowGridLines then yield "showGridLines=\"0\""
        if not sheet.View.ShowHeaders then yield "showRowColHeaders=\"0\""
        if not sheet.View.ShowZeros then yield "showZeros=\"0\""
        if not sheet.View.ShowOutlineSymbols then yield "showOutlineSymbols=\"0\""
        if active then yield "tabSelected=\"1\""
        if sheet.View.ZoomScale <> 100 then
            yield sprintf "zoomScale=\"%d\"" sheet.View.ZoomScale
            yield sprintf "zoomScaleNormal=\"%d\"" sheet.View.ZoomScale
        yield "workbookViewId=\"0\""
    ]
    let pane =
        if sheet.FrozenRows > 0 || sheet.FrozenCols > 0 then
            let topLeft = cellRef sheet.FrozenRows sheet.FrozenCols
            let paneName =
                match sheet.FrozenRows > 0, sheet.FrozenCols > 0 with
                | true, true -> "bottomRight"
                | true, false -> "bottomLeft"
                | false, true -> "topRight"
                | _ -> "topLeft"
            sprintf "<pane xSplit=\"%d\" ySplit=\"%d\" topLeftCell=\"%s\" activePane=\"%s\" state=\"frozen\"/>"
                sheet.FrozenCols sheet.FrozenRows topLeft paneName
        else ""
    sprintf "<sheetViews><sheetView %s>%s</sheetView></sheetViews>" (String.concat " " attrs) pane

/// What the rows of a sheet written from its cells keep of the part the
/// sheet was read from (`SheetExtras.SourceXml`).
type private RowCarry = {
    /// By row (from 0, where the row is now): the attributes of the file's
    /// `<row>` that a sheet does not hold - its style (`s`, `customFormat`),
    /// `thickBot`, `x14ac:dyDescent` ... - as text, each with a space before it.
    Attributes: System.Collections.Generic.Dictionary<int, string>
    /// The height the file gives a row without saying that it was set (`ht`
    /// without `customHeight`: what the row's text needed), in points.
    AutoHeights: System.Collections.Generic.Dictionary<int, float>
    /// The rows the file states no height for.
    Heightless: System.Collections.Generic.HashSet<int>
    /// The file writes a formula that reads the same from cell to cell once
    /// (`<f t="shared">`): so do the rows.
    Share: bool
    /// How many of the file's cells name a shared string (None: not known).
    SharedStrings: int option
}

/// A cell with a formula whose `<f>` element is given (a shared formula).
let private appendFormulaCellXml (sb: StringBuilder) (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (rowText: string) (cell: XlsxCellOut) (formulaXml: string) =
    let styleId = styles.StyleIdOf cell.Format
    let cellType, cachedXml =
        match cell.Value with
        | XlsxFormula (_, Some cached) -> formulaResultXml sharedIndex cached
        | _ -> "", ""
    let tAttr = typeAttr cellType
    sb.Append("<c r=\"").Append(columnLetters cell.Col).Append(rowText).Append('"') |> ignore
    if styleId <> 0 then sb.Append(" s=\"").Append(string styleId).Append('"') |> ignore
    sb.Append(tAttr).Append('>').Append(formulaXml).Append(cachedXml).Append("</c>") |> ignore

let private rowsXml (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (dynamicCm: int) (carry: RowCarry option) (sheet: XlsxSheetOut) : string =
    // A data table is written on the first cell of its range.
    let dataTables = sheet.Extras.DataTables |> List.map (fun table -> CellRef.firstCell table.Ref, table) |> Map.ofList
    let arrayFormulas =
        sheet.Extras.ArrayFormulas
        |> List.map (fun (cell, range, dynamic) -> cell, (range, (if dynamic && dynamicCm > 0 then " cm=\"" + string dynamicCm + "\"" else "")))
        |> Map.ofList
    let cells =
        let listed = List.toArray sheet.Cells
        if dataTables.IsEmpty then listed
        else
            // That cell is written even when it holds nothing.
            let present = System.Collections.Generic.HashSet<string>()
            for c in listed do
                present.Add (cellRef c.Row c.Col) |> ignore
            let missing =
                dataTables
                |> Map.toArray
                |> Array.choose (fun (reference, _) ->
                    if present.Contains reference then None
                    else
                        let mutable col = 0
                        let mutable row = 0
                        for ch in reference do
                            if ch >= 'A' && ch <= 'Z' then col <- col * 26 + (int ch - int 'A' + 1)
                            elif ch >= '0' && ch <= '9' then row <- row * 10 + (int ch - int '0')
                        if row < 1 || col < 1 then None
                        else Some { Row = row - 1; Col = col - 1; Value = XlsxBlank; Format = defaultXlsxCellFormat; Comment = None })
            if missing.Length = 0 then listed else Array.append listed missing
    // Cells in row order, then column order. They usually arrive that way.
    let mutable ordered = true
    let mutable i = 1
    while ordered && i < cells.Length do
        let a = cells.[i - 1]
        let b = cells.[i]
        if a.Row > b.Row || (a.Row = b.Row && a.Col > b.Col) then ordered <- false
        i <- i + 1
    let cells =
        if ordered then cells
        // (Stable: cells at the same place keep their order.)
        else cells |> Array.mapi (fun index cell -> index, cell) |> Array.sortWith (fun (ia, a) (ib, b) -> if a.Row <> b.Row then compare a.Row b.Row elif a.Col <> b.Col then compare a.Col b.Col else compare ia ib) |> Array.map snd
    // What a row has besides a height: the height a hidden one keeps, outline level, collapsed mark.
    let axis = System.Collections.Generic.Dictionary<int, float option * int * bool>()
    for (row, kept, level, collapsed) in sheet.Extras.RowAxis do axis.[row] <- (kept, level, collapsed)
    // (A row that only has an outline level is written too: NaN stands for "no height of its own".)
    let given = System.Collections.Generic.HashSet<int>(sheet.RowHeights |> List.map fst)
    let heights =
        [| yield! sheet.RowHeights
           for KeyValue(row, (_, level, collapsed)) in axis do
               if (level > 0 || collapsed) && given.Add row then yield row, nan
           // (So is a row the file gives a style or a height of its own to.)
           match carry with
           | Some carry ->
               for row in carry.Attributes.Keys do
                   if given.Add row then yield row, nan
               for row in carry.AutoHeights.Keys do
                   if given.Add row then yield row, nan
           | None -> () |]
        |> Array.sortBy fst
    // Formulas that read the same from one cell to the next - down a column,
    // or along a row - are written once, by the first of the cells: what a
    // reader makes of the others (`shiftFormulaRefs`) is the formula they hold.
    let sharedFormulas = System.Collections.Generic.Dictionary<int, string>()
    if (match carry with Some carry -> carry.Share | None -> false) then
        let key (row: int) (col: int) = float row * 16384.0 + float col
        let formulaAt = System.Collections.Generic.Dictionary<float, int>()
        let formulaOf (index: int) = match cells.[index].Value with XlsxFormula (formula, _) -> formula | _ -> ""
        for index in 0 .. cells.Length - 1 do
            let cell = cells.[index]
            match cell.Value with
            // (A shared formula cannot read another sheet: Excel refuses the file.)
            | XlsxFormula (formula, _) when formula <> ""
                                            && not (formula.Contains "!")
                                            && (arrayFormulas.IsEmpty || not (Map.containsKey (cellRef cell.Row cell.Col) arrayFormulas))
                                            && (dataTables.IsEmpty || not (Map.containsKey (cellRef cell.Row cell.Col) dataTables)) ->
                formulaAt.[key cell.Row cell.Col] <- index
            | _ -> ()
        let mutable group = 0
        for index in 0 .. cells.Length - 1 do
            let first = cells.[index]
            match formulaAt.TryGetValue (key first.Row first.Col) with
            | true, known when known = index && not (sharedFormulas.ContainsKey index) ->
                let text = formulaOf index
                let run (dRow: int) (dCol: int) =
                    let members = ResizeArray<int>()
                    let mutable going = true
                    while going do
                        let step = members.Count + 1
                        match formulaAt.TryGetValue (key (first.Row + step * dRow) (first.Col + step * dCol)) with
                        | true, next when not (sharedFormulas.ContainsKey next)
                                          && formulaOf next = XlsxReader.Xlsx.shiftFormulaRefs text (step * dRow) (step * dCol) -> members.Add next
                        | _ -> going <- false
                    members
                let down = run 1 0
                let members = if down.Count > 0 then down else run 0 1
                if members.Count > 0 then
                    let last = cells.[members.[members.Count - 1]]
                    sharedFormulas.[index] <-
                        "<f t=\"shared\" ref=\"" + cellRef first.Row first.Col + ":" + cellRef last.Row last.Col + "\" si=\"" + string group + "\">" + escapedFast (fileFormula text) + "</f>"
                    for next in members do sharedFormulas.[next] <- "<f t=\"shared\" si=\"" + string group + "\"/>"
                    group <- group + 1
            | _ -> ()
    // The last height given for a row wins, as with a map built from the list.
    let defaultHeight = sheet.Extras.Format |> Option.bind (fun format -> format.DefaultRowHeight) |> Option.defaultValue 15.0
    let sb = StringBuilder()
    let mutable cellAt = 0
    let mutable heightAt = 0
    while cellAt < cells.Length || heightAt < heights.Length do
        let row =
            if cellAt >= cells.Length then fst heights.[heightAt]
            elif heightAt >= heights.Length then cells.[cellAt].Row
            else min cells.[cellAt].Row (fst heights.[heightAt])
        let mutable height = None
        while heightAt < heights.Length && fst heights.[heightAt] = row do
            height <- Some (snd heights.[heightAt])
            heightAt <- heightAt + 1
        let rowText = string (row + 1)
        sb.Append("<row r=\"").Append(rowText).Append('"') |> ignore
        let extra = match axis.TryGetValue row with | true, found -> Some found | false, _ -> None
        match height with
        // Height 0 is how the sheet model says "hidden" (as for columns); the
        // row keeps the height it is shown at again.
        | Some h when h <= 0.0 ->
            (match extra with
             // (A hidden row the file states no height for has the height every row has.)
             | Some (Some kept, _, _) when kept > 0.0 && not (match carry with Some carry -> carry.Heightless.Contains row && abs (kept - defaultHeight) < 0.01 | None -> false) ->
                sb.Append(" ht=\"").Append(invariantFixed kept).Append("\" customHeight=\"1\"") |> ignore
             | _ -> ())
            sb.Append(" hidden=\"1\"") |> ignore
        | Some h ->
            // (A height the file states without saying it was set stays such a one.)
            let auto = match carry with Some carry -> (match carry.AutoHeights.TryGetValue row with | true, known -> Some known | false, _ -> None) | None -> None
            match auto with
            | Some known when System.Double.IsNaN h || abs (known - h) < 0.01 -> sb.Append(" ht=\"").Append(invariant known).Append('"') |> ignore
            | _ -> if not (System.Double.IsNaN h) then sb.Append(" ht=\"").Append(invariantFixed h).Append("\" customHeight=\"1\"") |> ignore
        | None -> ()
        match extra with
        | Some (_, level, collapsed) ->
            if level > 0 then sb.Append(" outlineLevel=\"").Append(string level).Append('"') |> ignore
            if collapsed then sb.Append(" collapsed=\"1\"") |> ignore
        | None -> ()
        match carry with
        | Some carry ->
            match carry.Attributes.TryGetValue row with
            | true, attributes -> sb.Append(attributes) |> ignore
            | false, _ -> ()
        | None -> ()
        sb.Append('>') |> ignore
        while cellAt < cells.Length && cells.[cellAt].Row = row do
            match sharedFormulas.TryGetValue cellAt with
            | true, formulaXml -> appendFormulaCellXml sb styles sharedIndex rowText cells.[cellAt] formulaXml
            | false, _ -> appendCellXml sb styles sharedIndex dataTables arrayFormulas rowText cells.[cellAt]
            cellAt <- cellAt + 1
        sb.Append("</row>") |> ignore
    sb.ToString()

let private mergeCellsXmlFull (mergedCells: MergedCellRange list) =
    if List.isEmpty mergedCells then ""
    else
        let body =
            mergedCells
            |> List.map (fun r -> sprintf "<mergeCell ref=\"%s:%s\"/>" (cellRef r.StartRow r.StartCol) (cellRef r.EndRow r.EndCol))
            |> String.concat ""
        sprintf "<mergeCells count=\"%d\">%s</mergeCells>" mergedCells.Length body

/// A validation whose formulas read another sheet: Excel keeps those in the
/// sheet's extension list (`<x14:dataValidations>`).
let private isExtValidation (dv: XlsxDataValidation) =
    let reads (formula: string option) = formula |> Option.exists (fun text -> text.Contains "!")
    reads dv.Formula1 || reads dv.Formula2

let private dataValidationAttrs (withSqref: bool) (dv: XlsxDataValidation) = [
    yield attr "type" dv.ValidationType
    if withSqref then yield attr "sqref" dv.Sqref
    match dv.ErrorStyle with Some style when style <> "stop" -> yield attr "errorStyle" style | _ -> ()
    match dv.Operator with Some op -> yield attr "operator" op | None -> ()
    if dv.AllowBlank then yield "allowBlank=\"1\""
    if not dv.ShowDropDown then yield "showDropDown=\"1\""
    if dv.ShowInputMessage then yield "showInputMessage=\"1\""
    if dv.ShowErrorMessage then yield "showErrorMessage=\"1\""
    match dv.ErrorTitle with Some t -> yield attr "errorTitle" t | None -> ()
    match dv.ErrorMessage with Some e -> yield attr "error" e | None -> ()
    match dv.PromptTitle with Some t -> yield attr "promptTitle" t | None -> ()
    match dv.InputMessage with Some p -> yield attr "prompt" p | None -> ()
]

let private dataValidationsXml (all: XlsxDataValidation list) =
    let validations = all |> List.filter (fun dv -> not (isExtValidation dv))
    if List.isEmpty validations then ""
    else
        let body =
            validations
            |> List.map (fun dv ->
                let f1 = dv.Formula1 |> Option.map (elemText "formula1") |> Option.defaultValue ""
                let f2 = dv.Formula2 |> Option.map (elemText "formula2") |> Option.defaultValue ""
                sprintf "<dataValidation %s>%s%s</dataValidation>" (String.concat " " (dataValidationAttrs true dv)) f1 f2)
            |> String.concat ""
        sprintf "<dataValidations count=\"%d\">%s</dataValidations>" validations.Length body

let private x14Namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
let private xmNamespace = "http://schemas.microsoft.com/office/excel/2006/main"

/// The extension of a sheet for the validations that read another sheet, or "".
let private dataValidationsExtXml (all: XlsxDataValidation list) =
    let validations = all |> List.filter isExtValidation
    if List.isEmpty validations then ""
    else
        let formula (name: string) (text: string option) =
            match text with
            | Some text -> "<x14:" + name + "><xm:f>" + escaped text + "</xm:f></x14:" + name + ">"
            | None -> ""
        let body =
            validations
            |> List.map (fun dv ->
                sprintf "<x14:dataValidation %s>%s%s<xm:sqref>%s</xm:sqref></x14:dataValidation>"
                    (String.concat " " (dataValidationAttrs false dv)) (formula "formula1" dv.Formula1) (formula "formula2" dv.Formula2) (escaped dv.Sqref))
            |> String.concat ""
        sprintf "<ext uri=\"{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}\" xmlns:x14=\"%s\"><x14:dataValidations count=\"%d\" xmlns:xm=\"%s\">%s</x14:dataValidations></ext>"
            x14Namespace validations.Length xmNamespace body

/// The id an x14 rule has (and its 2007 twin names): made of the sheet and the rule's place in it.
let private x14RuleId (sheetIndex: int) (ruleNumber: int) =
    let hex (value: int) (width: int) =
        let digits = "0123456789ABCDEF"
        let mutable text = ""
        let mutable rest = value
        for _ in 1 .. width do
            text <- string digits.[rest % 16] + text
            rest <- rest / 16
        text
    "{14000000-0000-4000-" + hex sheetIndex 4 + "-" + hex ruleNumber 12 + "}"

/// The main-part `<conditionalFormatting>` blocks of a sheet and the x14
/// blocks of its extension list. Rules keep their type-specific payload
/// (formulas, colour scale, data bar, icon set, top-N / above-average
/// attributes); a `dxfId` outside the workbook's dxf table is dropped rather
/// than written as a dangling index. A data bar with x14 options is written
/// twice, as Excel does: the 2007 bar, which names its x14 twin by id.
let private conditionalFormattingParts (dxfCount: int) (sheetIndex: int) (formats: XlsxReader.ConditionalFormat list) : string * string =
    let mutable nextPriority = 0
    let mutable ruleNumber = 0
    let cfvo (kind: string, value: string option) =
        match value with
        | Some v -> sprintf "<cfvo %s %s/>" (attr "type" kind) (attr "val" v)
        | None -> sprintf "<cfvo %s/>" (attr "type" kind)
    let x14Cfvo (kind: string, value: string option) =
        match value with
        | Some v -> sprintf "<x14:cfvo %s><xm:f>%s</xm:f></x14:cfvo>" (attr "type" kind) (escaped v)
        | None -> sprintf "<x14:cfvo %s/>" (attr "type" kind)
    let color (c: string) = sprintf "<color rgb=\"%s\"/>" (cleanRgb c)
    let x14Color (name: string) (c: string option) =
        match c with
        | Some c -> sprintf "<x14:%s rgb=\"%s\"/>" name (cleanRgb c)
        | None -> ""
    let ruleAttrs (rule: XlsxReader.ConditionalFormatRule) (priority: int) (withDxfId: bool) = [
        yield attr "type" rule.Type
        match rule.DxfId with
        | Some id when withDxfId && id >= 0 && id < dxfCount -> yield sprintf "dxfId=\"%d\"" id
        | _ -> ()
        yield sprintf "priority=\"%d\"" priority
        if rule.StopIfTrue then yield "stopIfTrue=\"1\""
        if rule.Type = "aboveAverage" && not rule.AboveAverage then yield "aboveAverage=\"0\""
        if rule.Percent then yield "percent=\"1\""
        if rule.Bottom then yield "bottom=\"1\""
        match rule.Operator with Some op -> yield attr "operator" op | None -> ()
        match rule.Text with Some t -> yield attr "text" t | None -> ()
        match rule.TimePeriod with Some period -> yield attr "timePeriod" period | None -> ()
        match rule.Rank with Some r -> yield sprintf "rank=\"%d\"" r | None -> ()
        match rule.StdDev with Some d -> yield sprintf "stdDev=\"%d\"" d | None -> ()
        if rule.EqualAverage then yield "equalAverage=\"1\""
    ]
    let sourceRule (xml: string) (rule: XlsxReader.ConditionalFormatRule) (priority: int) =
        match XmlPatch.readPart xml with
        | None -> failwith "The imported conditional-format rule XML cannot be preserved."
        | Some part ->
            XmlPatch.apply (fun _ -> []) part
                [ XmlPatch.SetAttribute ("priority", string priority)
                  if rule.StopIfTrue then XmlPatch.SetAttribute ("stopIfTrue", "1")
                  else XmlPatch.RemoveAttribute "stopIfTrue" ]
    let extBlocks = ResizeArray<string>()
    let extBlock (cf: XlsxReader.ConditionalFormat) (ruleXml: string) =
        let attributes =
            [ yield attr "xmlns:xm" xmNamespace
              for name, value in cf.GroupAttributes do
                  if name <> "sqref" && name <> "xmlns:xm" then yield attr name value ]
        extBlocks.Add (sprintf "<x14:conditionalFormatting %s>%s<xm:sqref>%s</xm:sqref></x14:conditionalFormatting>" (String.concat " " attributes) ruleXml (escaped cf.Range))
    let main =
        formats
        |> List.filter (fun cf -> not (List.isEmpty cf.Rules))
        |> List.map (fun cf ->
            let rules =
                cf.Rules
                |> List.map (fun (rule: XlsxReader.ConditionalFormatRule) ->
                    nextPriority <- nextPriority + 1
                    ruleNumber <- ruleNumber + 1
                    let priority = rule.Priority |> Option.defaultValue nextPriority
                    let id = x14RuleId sheetIndex ruleNumber
                    match rule.SourceXml, rule.Ext with
                    | Some xml, extension ->
                        let written = sourceRule xml rule priority
                        if extension |> Option.exists (fun ext -> ext.InExtList) then
                            extBlock cf written
                            ""
                        else written
                    // In the x14 list only.
                    | None, Some ext when ext.InExtList ->
                        let body =
                            match rule.Type with
                            | "iconSet" ->
                                sprintf "<x14:iconSet %s%s%s%s>%s%s</x14:iconSet>"
                                    (attr "iconSet" (rule.IconSet |> Option.defaultValue "3TrafficLights1"))
                                    (if ext.ShowValue then "" else " showValue=\"0\"")
                                    (if rule.IconSetReverse then " reverse=\"1\"" else "")
                                    (if ext.IconCustom.IsEmpty then "" else " custom=\"1\"")
                                    (rule.IconSetCfvo |> List.map x14Cfvo |> String.concat "")
                                    (ext.IconCustom |> List.map (fun (family, icon) -> sprintf "<x14:cfIcon %s iconId=\"%d\"/>" (attr "iconSet" family) icon) |> String.concat "")
                            | "colorScale" ->
                                "<x14:colorScale>"
                                + (rule.ColorScale |> List.map (fun stop -> x14Cfvo (stop.CfvoType, stop.CfvoValue)) |> String.concat "")
                                + (rule.ColorScale |> List.map (fun stop -> x14Color "color" (Some stop.Color)) |> String.concat "")
                                + "</x14:colorScale>"
                            | _ ->
                                (rule.Formulas |> List.map (fun f -> "<xm:f>" + escaped f + "</xm:f>") |> String.concat "")
                                + (match ext.Dxf with Some dxf -> dxfXml "x14:dxf" dxf | None -> "")
                        extBlock cf (sprintf "<x14:cfRule %s id=\"%s\">%s</x14:cfRule>" (String.concat " " (ruleAttrs rule priority false)) id body)
                        ""
                    | _ ->
                        let barExt = rule.Ext |> Option.bind (fun ext -> ext.DataBar)
                        let children = [
                            for f in rule.Formulas -> elemText "formula" (fileFormula f)
                            if not (List.isEmpty rule.ColorScale) then
                                yield "<colorScale>"
                                      + (rule.ColorScale |> List.map (fun stop -> cfvo (stop.CfvoType, stop.CfvoValue)) |> String.concat "")
                                      + (rule.ColorScale |> List.map (fun stop -> color stop.Color) |> String.concat "")
                                      + "</colorScale>"
                            if rule.Type = "dataBar" then
                                let barColor = rule.DataBarColor |> Option.defaultValue "638EC6"
                                yield "<dataBar"
                                      + (match barExt with Some bar when not bar.ShowValue -> " showValue=\"0\"" | _ -> "")
                                      + ">"
                                      + (rule.DataBarCfvo |> List.map cfvo |> String.concat "")
                                      + color barColor
                                      + "</dataBar>"
                                match barExt with
                                | Some bar ->
                                    yield sprintf "<extLst><ext xmlns:x14=\"%s\" uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\"><x14:id>%s</x14:id></ext></extLst>" x14Namespace id
                                    let barAttrs = [
                                        yield sprintf "minLength=\"%d\"" bar.MinLength
                                        yield sprintf "maxLength=\"%d\"" bar.MaxLength
                                        if not bar.ShowValue then yield "showValue=\"0\""
                                        if bar.Border then yield "border=\"1\""
                                        if not bar.Gradient then yield "gradient=\"0\""
                                        if bar.Direction <> "context" then yield attr "direction" bar.Direction
                                        if bar.NegativeBarColorSameAsPositive then yield "negativeBarColorSameAsPositive=\"1\""
                                        if bar.Border && not bar.NegativeBarBorderColorSameAsPositive then yield "negativeBarBorderColorSameAsPositive=\"0\""
                                        if bar.AxisPosition <> "automatic" then yield attr "axisPosition" bar.AxisPosition
                                    ]
                                    extBlock cf
                                        (sprintf "<x14:cfRule type=\"dataBar\" id=\"%s\"><x14:dataBar %s>%s%s%s%s%s%s</x14:dataBar></x14:cfRule>"
                                            id (String.concat " " barAttrs)
                                            (bar.Cfvo |> List.map x14Cfvo |> String.concat "")
                                            (x14Color "fillColor" (Some barColor))
                                            (if bar.Border then x14Color "borderColor" (match bar.BorderColor with Some c -> Some c | None -> Some barColor) else "")
                                            (if bar.NegativeBarColorSameAsPositive then "" else x14Color "negativeFillColor" (match bar.NegativeFillColor with Some c -> Some c | None -> Some "FF0000"))
                                            (if bar.Border && not bar.NegativeBarBorderColorSameAsPositive then x14Color "negativeBorderColor" (match bar.NegativeBorderColor with Some c -> Some c | None -> Some "FF0000") else "")
                                            (if bar.AxisPosition = "none" then "" else x14Color "axisColor" (match bar.AxisColor with Some c -> Some c | None -> Some "000000")))
                                | None -> ()
                            match rule.IconSet with
                            | Some family ->
                                yield sprintf "<iconSet %s%s%s>%s</iconSet>" (attr "iconSet" family)
                                        (match rule.Ext with Some ext when not ext.ShowValue -> " showValue=\"0\"" | _ -> "")
                                        (if rule.IconSetReverse then " reverse=\"1\"" else "") (rule.IconSetCfvo |> List.map cfvo |> String.concat "")
                            | None -> ()
                        ]
                        sprintf "<cfRule %s>%s</cfRule>" (String.concat " " (ruleAttrs rule priority true)) (String.concat "" children))
                |> String.concat ""
            if rules = "" then "" else
                let attributes =
                    attr "sqref" cf.Range
                    :: [ for key, value in cf.GroupAttributes do
                             if key <> "sqref" then attr key value ]
                sprintf "<conditionalFormatting %s>%s</conditionalFormatting>" (String.concat " " attributes) rules)
        |> String.concat ""
    main, String.concat "" extBlocks

// ---- XML that was read, changed through its span tree (XmlPatch) ----

let private order = XmlPatch.Spreadsheet.order

/// A worksheet part read in outline - its rows without their cells (`XmlPatch.readPartExcept`). None for a
/// part that is not a worksheet.
let private readWorksheet (text: string) : XmlPatch.SourceElement option =
    XmlPatch.readPartExcept (fun name -> name = "row") text
    |> Option.filter (fun part -> XmlPatch.localName part.Node = "worksheet" && (let span = XmlPatch.span part.Node in span.StartTagEnd < span.End))

/// A worksheet part read with its cells, but for their values (`<v>`, `<is>`).
let private readWorksheetCells (text: string) : XmlPatch.SourceElement option =
    XmlPatch.readPartExcept (fun name -> name = "v" || name = "is") text
    |> Option.filter (fun part -> XmlPatch.localName part.Node = "worksheet")

/// The first child of an element of that name (as written).
let private childNamed (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode option =
    XmlPatch.elements node |> List.tryFind (fun kid -> XmlPatch.name kid = name)

/// The children of an element of that name (as written).
let private childrenNamed (name: string) (node: XmlDsl.XmlNode) : XmlDsl.XmlNode list =
    XmlPatch.elements node |> List.filter (fun kid -> XmlPatch.name kid = name)

/// The child of an element read of that name, as an element of its own.
let private sourceChild (name: string) (source: XmlPatch.SourceElement) : XmlPatch.SourceElement option =
    childNamed name source.Node |> Option.map (XmlPatch.at source)

/// An attribute's edit: the value set, or (None) the attribute taken out.
let private attributeEdit (name: string) (value: string option) : XmlPatch.ElementEdit =
    match value with
    | Some value -> XmlPatch.SetAttribute (name, value)
    | None -> XmlPatch.RemoveAttribute name

let private isSelfClosed (node: XmlDsl.XmlNode) = let span = XmlPatch.span node in span.StartTagEnd = span.End

/// Whether an attribute says yes ("1", "true").
let private isOn (value: string option) = match value with Some "1" | Some "true" -> true | _ -> false

/// The sheet's extension list: the x14 conditional formats and validations
/// written from the rules, and the extensions kept from the file. x14 rules
/// join a kept x14 list (one whose rules were not all understood).
let private sheetExtLstXml (cfBlocks: string) (validationsExt: string) (kept: string list) =
    // (The list of conditional formats of an extension.)
    let cfList (xml: string) =
        XmlPatch.read xml |> Option.bind (fun ext -> XmlPatch.elements ext.Node |> List.tryFind (fun node -> XmlPatch.localName node = "conditionalFormattings") |> Option.map (fun list -> ext, list))
    let keptCf, others = kept |> List.partition (fun xml -> (cfList xml).IsSome)
    let cfExt =
        match keptCf, cfBlocks with
        | [], "" -> []
        | [], blocks ->
            [ sprintf "<ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"%s\"><x14:conditionalFormattings>%s</x14:conditionalFormattings></ext>" x14Namespace blocks ]
        | first :: rest, blocks ->
            match cfList first with
            | Some (ext, list) when blocks <> "" && (XmlPatch.attribute ext.Node "xmlns:x14").IsSome ->
                XmlPatch.apply order ext [ XmlPatch.EditChild (XmlPatch.name list, [ XmlPatch.Append blocks ], None) ] :: rest
            | _ -> keptCf
    match cfExt @ (if validationsExt = "" then [] else [ validationsExt ]) @ others with
    | [] -> ""
    | extensions -> "<extLst>" + String.concat "" extensions + "</extLst>"

let private hyperlinksXml hyperlinks =
    if List.isEmpty hyperlinks then ""
    else
        let body =
            hyperlinks
            |> List.mapi (fun i h ->
                let extra =
                    (match h.Tooltip with Some tip -> " " + attr "tooltip" tip | None -> "")
                    + (match h.Display with Some text -> " " + attr "display" text | None -> "")
                if h.IsInternal then
                    sprintf "<hyperlink %s %s%s/>" (attr "ref" h.Range) (attr "location" h.Target) extra
                else
                    sprintf "<hyperlink %s r:id=\"rIdHyperlink%d\"%s/>" (attr "ref" h.Range) (i + 1) extra)
            |> String.concat ""
        "<hyperlinks>" + body + "</hyperlinks>"

/// `<printOptions>`, `<pageMargins>`, `<pageSetup>` and `<headerFooter>` of a sheet ("" for one that is not written).
let private pageParts (sheet: XlsxSheetOut) : string * string * string * string =
    let printOptions =
        sheet.PrintOptions
        |> Option.map (fun p ->
            let attrs = [
                if p.GridLines then yield "gridLines=\"1\""
                if p.Headings then yield "headings=\"1\""
                if p.HorizontalCentered then yield "horizontalCentered=\"1\""
                if p.VerticalCentered then yield "verticalCentered=\"1\""
            ]
            if List.isEmpty attrs then "" else sprintf "<printOptions %s/>" (String.concat " " attrs))
        |> Option.defaultValue ""
    let margins =
        sheet.PageMargins
        |> Option.map (fun m ->
            sprintf "<pageMargins left=\"%s\" right=\"%s\" top=\"%s\" bottom=\"%s\" header=\"%s\" footer=\"%s\"/>"
                (invariant m.Left) (invariant m.Right) (invariant m.Top) (invariant m.Bottom) (invariant m.Header) (invariant m.Footer))
        |> Option.defaultValue ""
    let pageSetup =
        sheet.PageSetup
        |> Option.map (fun p ->
            let attrs = [
                match p.PaperSize with Some n -> yield sprintf "paperSize=\"%d\"" n | None -> ()
                match p.Orientation with Some o -> yield attr "orientation" o | None -> ()
                yield sprintf "scale=\"%d\"" p.Scale
                // (One page each way is what a file without the attributes says.)
                if p.FitToWidth <> 1 then yield sprintf "fitToWidth=\"%d\"" p.FitToWidth
                if p.FitToHeight <> 1 then yield sprintf "fitToHeight=\"%d\"" p.FitToHeight
                // The number is stated whenever it is not the default, used or not.
                if p.UseFirstPageNumber || p.FirstPageNumber <> 1 then
                    yield sprintf "firstPageNumber=\"%d\"" p.FirstPageNumber
                if p.UseFirstPageNumber then
                    yield "useFirstPageNumber=\"1\""
                match p.PageOrder with Some po -> yield attr "pageOrder" po | None -> ()
                if p.BlackAndWhite then yield "blackAndWhite=\"1\""
                match p.CellComments with Some cc -> yield attr "cellComments" cc | None -> ()
                for name, value in p.Kept do yield attr name value
                // The printer settings part of the file the sheet came from.
                match sheet.Extras.PrinterSettingsRelId with Some id -> yield attr "r:id" id | None -> ()
            ]
            sprintf "<pageSetup %s/>" (String.concat " " attrs))
        |> Option.defaultWith (fun () ->
            match sheet.Extras.PrinterSettingsRelId with
            | Some id -> sprintf "<pageSetup %s/>" (attr "r:id" id)
            | None -> "")
    let headerFooter =
        sheet.HeaderFooter
        |> Option.map (fun h ->
            let attrs = [
                if h.DifferentOddEven then yield "differentOddEven=\"1\""
                if h.DifferentFirst then yield "differentFirst=\"1\""
                if not h.ScaleWithDoc then yield "scaleWithDoc=\"0\""
            ]
            let children = [
                match h.OddHeader with Some v -> yield elemText "oddHeader" v | None -> ()
                match h.OddFooter with Some v -> yield elemText "oddFooter" v | None -> ()
                match h.EvenHeader with Some v -> yield elemText "evenHeader" v | None -> ()
                match h.EvenFooter with Some v -> yield elemText "evenFooter" v | None -> ()
                match h.FirstHeader with Some v -> yield elemText "firstHeader" v | None -> ()
                match h.FirstFooter with Some v -> yield elemText "firstFooter" v | None -> ()
            ]
            // (A header and footer that say nothing are not written.)
            if List.isEmpty attrs && List.isEmpty children then ""
            else sprintf "<headerFooter%s>%s</headerFooter>" (if List.isEmpty attrs then "" else " " + String.concat " " attrs) (String.concat "" children))
        |> Option.defaultValue ""
    printOptions, margins, pageSetup, headerFooter

let private pagePartsXml (sheet: XlsxSheetOut) =
    let printOptions, margins, pageSetup, headerFooter = pageParts sheet
    printOptions + margins + pageSetup + headerFooter

/// The manual page breaks of a sheet. `was`: the file's element ("" for
/// none) - a break the file has is written as the file wrote it (its `min`
/// and `max`: how far the break reaches, which is the print area's extent),
/// and a new one reaches as far as the file's breaks all do. Without that a
/// row break reaches over every column (`max="16383"`) and a column break
/// over every row (`max="1048575"`).
let private breaksXml (name: string) (was: XmlPatch.SourceElement option) (breaks: int list) =
    if List.isEmpty breaks then ""
    else
        let known =
            match was with
            | Some was ->
                XmlPatch.elements was.Node
                |> List.filter (fun brk -> XmlPatch.name brk = "brk")
                |> List.choose (fun brk ->
                    match XmlPatch.attribute brk "id" |> Option.map System.Int32.TryParse, XmlPatch.attribute brk "man" with
                    | Some (true, id), (Some "1" | Some "true") -> Some (id, brk, XmlPatch.bytesOf was.Text brk)
                    | _ -> None)
            | None -> []
        let common (attribute: string) =
            match known |> List.map (fun (_, brk, _) -> XmlPatch.attribute brk attribute) |> List.distinct with
            | [ Some value ] -> Some value
            | _ -> None
        let reach =
            (match common "min" with Some value -> " " + attr "min" value | None -> "")
            + " " + attr "max" (common "max" |> Option.defaultValue (if name = "colBreaks" then "1048575" else "16383"))
        let body =
            breaks
            |> List.map (fun id ->
                match known |> List.tryFind (fun (known, _, _) -> known = id) with
                | Some (_, _, xml) -> xml
                | None -> sprintf "<brk id=\"%d\"%s man=\"1\"/>" id reach)
            |> String.concat ""
        sprintf "<%s count=\"%d\" manualBreakCount=\"%d\">%s</%s>" name breaks.Length breaks.Length body name

let private tablePartsXml count =
    if count = 0 then ""
    else
        let body = [ 1 .. count ] |> List.map (fun i -> sprintf "<tablePart r:id=\"rIdTable%d\"/>" i) |> String.concat ""
        sprintf "<tableParts count=\"%d\">%s</tableParts>" count body

// ---- pictures, kept drawing objects and threaded comments ---------------------

// The sheet's drawing part is written as xl/drawings/sheetDrawing<n>.xml: a
// kept part may be called drawing<n>.xml (the user shapes of a chart).
let private hasDrawing (sheet: XlsxSheetOut) =
    not sheet.Extras.Images.IsEmpty || not sheet.Extras.Objects.IsEmpty

/// The file's own drawing part of a sheet, when it goes back as it is (the
/// sheet's pictures and objects are the file's: see SheetExtras.DrawingPart).
let private keptDrawingPart (workbook: XlsxWorkbookOut) (sheet: XlsxSheetOut) : string option =
    sheet.Extras.DrawingPart |> Option.filter (fun path -> workbook.Preserved.Parts |> List.exists (fun (kept, _) -> kept = path))

/// The file's own part of a sheet that is written as it was read (see
/// SheetExtras.Verbatim): its path and the type of the workbook's
/// relationship to it.
let private verbatimPart (workbook: XlsxWorkbookOut) (sheet: XlsxSheetOut) : (string * string) option =
    if not sheet.Extras.Verbatim then None
    else sheet.Extras.SourcePart |> Option.filter (fun (path, _) -> workbook.Preserved.Parts |> List.exists (fun (kept, _) -> kept = path))

/// The file's comments part of a sheet whose comments are the file's (see SheetExtras.CommentsPart).
let private keptCommentsPart (workbook: XlsxWorkbookOut) (sheet: XlsxSheetOut) : string option =
    sheet.Extras.CommentsPart |> Option.filter (fun path -> workbook.Preserved.Parts |> List.exists (fun (kept, _) -> kept = path))

/// Where the parts of a sheet are written.
type private SheetPlacement = {
    /// The file's part of a sheet written as it was read, and its relationship type.
    Verbatim: (string * string) option
    SheetPath: string
    CommentsPath: string
    /// The VML drawing made for the boxes of the comments.
    VmlPath: string
    ThreadsPath: string
    /// The drawing part made from the sheet's pictures and objects.
    DrawingPath: string
}

/// `xl/worksheets/sheet1.xml` -> `xl/worksheets/_rels/sheet1.xml.rels`.
let private relsPathOf (path: string) = PartPath.relationshipsOf path

/// The parts of every sheet, in sheet order: the usual names by the sheet's
/// number, unless a kept part (or the part of a sheet written as it was
/// read) has that name - then the next number no part has.
let private sheetPlacements (workbook: XlsxWorkbookOut) (keptPaths: Set<string>) : SheetPlacement list =
    let taken = System.Collections.Generic.HashSet<string>(keptPaths)
    let verbatim = workbook.Sheets |> List.map (verbatimPart workbook)
    for part in verbatim do
        match part with
        | Some (path, _) -> taken.Add path |> ignore
        | None -> ()
    // A worksheet of the file is written under the name its part had.
    let own =
        List.zip workbook.Sheets verbatim
        |> List.map (fun (sheet, verbatim) ->
            match verbatim, sheet.Extras.SourcePart with
            | None, Some (path, relType) when OpcXml.relationshipKind relType = "worksheet" && PartPath.extension path = "xml" && taken.Add path -> Some path
            | _ -> None)
    let free (name: int -> string) (number: int) =
        let mutable at = number
        while taken.Contains (name at) do at <- at + 1
        taken.Add (name at) |> ignore
        name at
    List.zip3 own verbatim workbook.Sheets
    |> List.mapi (fun i (own, verbatim, sheet) ->
        let number = i + 1
        match verbatim with
        | Some (path, _) -> { Verbatim = verbatim; SheetPath = path; CommentsPath = ""; VmlPath = ""; ThreadsPath = ""; DrawingPath = "" }
        | None ->
            { Verbatim = None
              SheetPath = (match own with Some path -> path | None -> free (sprintf "xl/worksheets/sheet%d.xml") number)
              CommentsPath = (match keptCommentsPart workbook sheet with Some path -> path | None -> free (sprintf "xl/comments%d.xml") number)
              VmlPath =
                // (The boxes of new comments: not under the name of a kept VML drawing.)
                (let usual = sprintf "xl/drawings/vmlDrawing%d.vml" number
                 if taken.Contains usual then free (sprintf "xl/drawings/commentsDrawing%d.vml") number
                 else (taken.Add usual |> ignore; usual))
              ThreadsPath = free (sprintf "xl/threadedComments/threadedComment%d.xml") number
              DrawingPath = free (sprintf "xl/drawings/sheetDrawing%d.xml") number })

/// The id of a table part (`<table id>`).
let private tableIdOf (xml: string) : int option =
    XmlPatch.readPart xml
    |> Option.bind (fun part -> XmlPatch.attribute part.Node "id")
    |> Option.bind (fun id -> match System.Int32.TryParse id with | true, id -> Some id | _ -> None)

/// Where each table of each sheet is written, sheet by sheet in order: the
/// part, and the id of a table part the writer makes (None: the file's part
/// goes back as it is, see SheetExtras.TableParts). New parts take numbers
/// and ids no kept table part has.
let private tablePlacement (workbook: XlsxWorkbookOut) : (string * int option) list list =
    let parts = System.Collections.Generic.Dictionary<string, byte[]>()
    for path, data in workbook.Preserved.Parts do parts.[path] <- data
    let taken = System.Collections.Generic.HashSet<int>()
    let claimed = System.Collections.Generic.Dictionary<string, XlsxTableOut>()
    let keptOf (sheet: XlsxSheetOut) (table: XlsxTableOut) =
        match sheet.Extras.TableParts |> List.tryPick (fun (name, path) -> if name = table.Name && parts.ContainsKey path then Some path else None) with
        | Some path -> Some path
        | None ->
            // (A changed table of the file: written as its part, changed. One table per part.)
            match table.SourcePart with
            | Some path when parts.ContainsKey path ->
                match claimed.TryGetValue path with
                | true, owner -> if obj.ReferenceEquals (owner, table) then Some path else None
                | false, _ -> claimed.[path] <- table; Some path
            | _ -> None
    for sheet in workbook.Sheets do
        for table in sheet.Tables do
            match keptOf sheet table with
            | Some path ->
                match PartPath.number "table" path with Some number -> taken.Add number |> ignore | None -> ()
                match tableIdOf (System.Text.Encoding.UTF8.GetString parts.[path]) with Some id -> taken.Add id |> ignore | None -> ()
            | None -> ()
    let mutable next = 1
    workbook.Sheets
    |> List.map (fun sheet ->
        // (The tables of a sheet written as it was read are parts of the file.)
        (if (verbatimPart workbook sheet).IsSome then [] else sheet.Tables)
        |> List.map (fun table ->
            match keptOf sheet table with
            | Some path -> path, None
            | None ->
                while taken.Contains next do next <- next + 1
                let id = next
                next <- next + 1
                sprintf "xl/tables/table%d.xml" id, Some id))

let private drawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
let private drawingMainNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
let private threadedCommentsNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

let private emu (value: float) = string (int64 (System.Math.Round value))

/// The drawing part of a sheet: one anchor per picture, written from the model,
/// then the kept objects (charts, shapes) as they were read.
let private drawingXml (sheet: XlsxSheetOut) (imageRelIds: string list) =
    let marker name (col: int) (colOffset: float) (row: int) (rowOffset: float) =
        sprintf "<xdr:%s><xdr:col>%d</xdr:col><xdr:colOff>%s</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%s</xdr:rowOff></xdr:%s>" name col (emu colOffset) row (emu rowOffset) name
    let picture index (relId: string) (image: XlsxReader.SheetImage) =
        let anchor = image.Anchor
        let width = anchor.Width |> Option.defaultValue 0.0
        let height = anchor.Height |> Option.defaultValue 0.0
        let name = image.Name |> Option.defaultValue (sprintf "Picture %d" (index + 1))
        let description = image.Description |> Option.map (fun text -> " " + attr "descr" text) |> Option.defaultValue ""
        let pic =
            sprintf "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"%d\" %s%s/><xdr:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:embed=\"%s\"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"%s\" cy=\"%s\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/>"
                (index + 2) (attr "name" name) description relId (emu width) (emu height)
        let from = marker "from" anchor.FromCol anchor.FromColOffset anchor.FromRow anchor.FromRowOffset
        match anchor.Kind, anchor.ToCol, anchor.ToRow with
        | "twoCell", Some toCol, Some toRow ->
            let editAs = anchor.EditAs |> Option.map (fun value -> " " + attr "editAs" value) |> Option.defaultValue ""
            sprintf "<xdr:twoCellAnchor%s>%s%s%s</xdr:twoCellAnchor>" editAs from (marker "to" toCol anchor.ToColOffset toRow anchor.ToRowOffset) pic
        | "absolute", _, _ ->
            sprintf "<xdr:absoluteAnchor><xdr:pos x=\"%s\" y=\"%s\"/><xdr:ext cx=\"%s\" cy=\"%s\"/>%s</xdr:absoluteAnchor>" (emu anchor.X) (emu anchor.Y) (emu width) (emu height) pic
        | _ ->
            sprintf "<xdr:oneCellAnchor>%s<xdr:ext cx=\"%s\" cy=\"%s\"/>%s</xdr:oneCellAnchor>" from (emu width) (emu height) pic
    // Pictures and kept objects in their z-order (the order of the anchors in
    // the drawing part), whichever kind they are.
    let pictures = List.map3 (fun index relId (image: XlsxReader.SheetImage) -> image.Anchor.Order, 0, picture index relId image) [ 0 .. sheet.Extras.Images.Length - 1 ] imageRelIds sheet.Extras.Images
    let objects = sheet.Extras.Objects |> List.map (fun item -> item.Anchor.Order, 1, item.Xml)
    let ordered = pictures @ objects |> List.sortBy (fun (order, kind, _) -> order, kind) |> List.map (fun (_, _, xml) -> xml)
    // Only the two namespaces every anchor uses are declared here; a kept object
    // carries the others on its own element, and a picture declares r: on its blip.
    sprintf "%s\n<xdr:wsDr xmlns:xdr=\"%s\" xmlns:a=\"%s\">%s</xdr:wsDr>"
        declaration drawingNamespace drawingMainNamespace (String.concat "" ordered)

/// The person behind an author name: the kept one, or a new entry.
let private personsOf (workbook: XlsxWorkbookOut) : (string * string * string * string) list =
    let authors =
        workbook.Sheets
        |> List.collect (fun sheet -> sheet.Extras.ThreadedComments |> List.collect (fun thread -> thread :: thread.Replies))
    let mutable persons = workbook.Preserved.Persons
    let mutable next = 1
    for comment in authors do
        let known =
            match comment.PersonId with
            | Some id -> persons |> List.exists (fun (personId, _, _, _) -> personId = id)
            | None -> persons |> List.exists (fun (_, name, _, _) -> Some name = comment.Author)
        if not known then
            let id =
                match comment.PersonId with
                | Some id -> id
                | None ->
                    let id = sprintf "{00000000-0000-4000-8000-%012d}" next
                    next <- next + 1
                    id
            let name = comment.Author |> Option.defaultValue ""
            persons <- persons @ [ id, name, name, "None" ]
    persons

let private threadedCommentsXml (persons: (string * string * string * string) list) (sheetIndex: int) (sheet: XlsxSheetOut) =
    if sheet.Extras.ThreadedComments.IsEmpty then None
    else
        let personId (comment: XlsxReader.ThreadedComment) =
            match comment.PersonId with
            | Some id -> id
            | None ->
                persons
                |> List.tryPick (fun (id, name, _, _) -> if Some name = comment.Author then Some id else None)
                |> Option.defaultValue "{00000000-0000-4000-8000-000000000000}"
        let mutable next = 0
        let idOf (comment: XlsxReader.ThreadedComment) =
            if comment.Id <> "" then comment.Id
            else
                next <- next + 1
                sprintf "{%08d-0000-4000-8000-%012d}" sheetIndex next
        let one (parent: string option) (comment: XlsxReader.ThreadedComment) (id: string) =
            sprintf "<threadedComment ref=\"%s%d\"%s personId=\"%s\" id=\"%s\"%s%s><text>%s</text></threadedComment>"
                comment.CellRef.Column comment.CellRef.Row
                (comment.Date |> Option.map (fun date -> " " + attr "dT" date) |> Option.defaultValue "")
                (escaped (personId comment)) (escaped id)
                (parent |> Option.map (fun parentId -> " " + attr "parentId" parentId) |> Option.defaultValue "")
                (if parent.IsNone && comment.Done then " done=\"1\"" else "")
                (escapedCellText comment.Text)
        let body =
            sheet.Extras.ThreadedComments
            |> List.collect (fun thread ->
                let rootId = idOf thread
                one None thread rootId :: (thread.Replies |> List.map (fun reply -> one (Some rootId) reply (idOf reply))))
            |> String.concat ""
        Some (sprintf "%s\n<ThreadedComments xmlns=\"%s\" xmlns:x=\"%s\">%s</ThreadedComments>" declaration threadedCommentsNamespace Namespaces.spreadsheet body)

let private personsXml (persons: (string * string * string * string) list) =
    let body =
        persons
        |> List.map (fun (id, name, userId, providerId) -> sprintf "<person %s %s %s %s/>" (attr "displayName" name) (attr "id" id) (attr "userId" userId) (attr "providerId" providerId))
        |> String.concat ""
    sprintf "%s\n<personList xmlns=\"%s\" xmlns:x=\"%s\">%s</personList>" declaration threadedCommentsNamespace Namespaces.spreadsheet body

/// `<sheetFormatPr>`: the file's default column width and row height when the
/// sheet came from a file, Excel's defaults otherwise.
let private sheetFormatXml (sheet: XlsxSheetOut) =
    // The deepest outline level of the rows / columns (Excel sizes its outline gutter by them).
    let deepest (axis: (int * float option * int * bool) list) = axis |> List.fold (fun level (_, _, own, _) -> max level own) 0
    let outline = [
        if deepest sheet.Extras.RowAxis > 0 then yield sprintf "outlineLevelRow=\"%d\"" (deepest sheet.Extras.RowAxis)
        if deepest sheet.Extras.ColumnAxis > 0 then yield sprintf "outlineLevelCol=\"%d\"" (deepest sheet.Extras.ColumnAxis)
    ]
    match sheet.Extras.Format with
    | None -> sprintf "<sheetFormatPr %s/>" (String.concat " " ("defaultRowHeight=\"15\"" :: outline))
    | Some format ->
        let attrs = [
            match format.BaseColWidth with Some width -> yield sprintf "baseColWidth=\"%s\"" (invariant width) | None -> ()
            match format.DefaultColWidth with Some width -> yield sprintf "defaultColWidth=\"%s\"" (invariant width) | None -> ()
            yield sprintf "defaultRowHeight=\"%s\"" (invariant (format.DefaultRowHeight |> Option.defaultValue 15.0))
            if format.CustomDefaultHeight then yield "customHeight=\"1\""
            yield! outline
        ]
        sprintf "<sheetFormatPr %s/>" (String.concat " " attrs)

/// `<autoFilter>` of a sheet (CT_AutoFilter): the range and one
/// `<filterColumn>` per filtered column.
let private autoFilterXml (dxfCount: int) (sheet: XlsxSheetOut) =
    match sheet.Extras.AutoFilter with
    | None -> ""
    | Some filter ->
        let column (c: XlsxReader.AutoFilterColumn) =
            let body =
                match c.Kind with
                | "filters" ->
                    let values = c.Values |> List.map (fun value -> sprintf "<filter %s/>" (attr "val" value)) |> String.concat ""
                    sprintf "<filters%s>%s</filters>" (if c.Blank then " blank=\"1\"" else "") values
                | "customFilters" ->
                    let conditions =
                        c.Custom
                        |> List.map (fun (operator, value) ->
                            sprintf "<customFilter%s %s/>" (if operator = "equal" || operator = "" then "" else " " + attr "operator" operator) (attr "val" value))
                        |> String.concat ""
                    sprintf "<customFilters%s>%s</customFilters>" (if c.And && c.Custom.Length > 1 then " and=\"1\"" else "") conditions
                | "top10" ->
                    sprintf "<top10%s%s val=\"%s\"/>" (if c.Top then "" else " top=\"0\"") (if c.Percent then " percent=\"1\"" else "") (invariant c.Value)
                | "colorFilter" ->
                    (match c.DxfId with
                     | Some id when id >= 0 && id < dxfCount -> sprintf "<colorFilter dxfId=\"%d\"%s/>" id (if c.CellColor then "" else " cellColor=\"0\"")
                     | _ -> "")
                | "dynamicFilter" -> sprintf "<dynamicFilter %s/>" (attr "type" c.DynamicType)
                | _ -> c.Xml
            if body = "" then "" else sprintf "<filterColumn colId=\"%d\">%s</filterColumn>" c.ColId body
        let columns = filter.Columns |> List.map column |> String.concat ""
        if columns = "" then sprintf "<autoFilter %s/>" (attr "ref" filter.Ref)
        else sprintf "<autoFilter %s>%s</autoFilter>" (attr "ref" filter.Ref) columns

/// The `<legacyDrawing>` element kept from the file ("" when none): the
/// sheet's own VML drawing with its form controls ("legacyDrawing"), or the
/// one that only holds the boxes of its comments ("legacyDrawingNotes", kept
/// while the comments are as the file had them).
let private keptLegacyDrawing (sheet: XlsxSheetOut) : string =
    sheet.Extras.KeptElements
    |> List.tryPick (fun (name, xml) -> if name = "legacyDrawing" || name = "legacyDrawingNotes" then Some xml else None)
    |> Option.defaultValue ""

/// `<tabColor>` of a sheet ("" for none).
let private tabColorXml (sheet: XlsxSheetOut) : string =
    match sheet.View.TabColor with
    | Some color ->
        match ColorText.parse color with
        | Some (ColorText.Rgb rgb) -> sprintf "<tabColor rgb=\"FF%s\"/>" (rgb.ToUpperInvariant())
        | Some (ColorText.Theme (index, None)) -> sprintf "<tabColor theme=\"%s\"/>" (escaped index)
        | Some (ColorText.Theme (index, Some tint)) -> sprintf "<tabColor theme=\"%s\" tint=\"%s\"/>" (escaped index) (escaped tint)
        | Some (ColorText.Indexed index) -> sprintf "<tabColor indexed=\"%s\"/>" (escaped index)
        | None -> ""
    | None -> ""

/// A run of a rich text of the file as the writer has it.
let private richRunOfReader (props: XlsxReader.RunProps) (text: string) : XlsxRichRun =
    { Text = text
      Bold = Some props.Bold
      Italic = Some props.Italic
      Underline = Some props.Underline
      UnderlineDouble = Some props.UnderlineDouble
      Strikethrough = Some props.Strikethrough
      Color = props.Color |> Option.map (fun c -> "#" + c.TrimStart('#'))
      FontName = props.FontName
      FontSize = props.FontSize
      VertAlign =
        match props.VertAlign with
        | XlsxReader.VertAlignSuperscript -> Some XlsxSuperscript
        | XlsxReader.VertAlignSubscript -> Some XlsxSubscript
        | _ -> None }

/// The comments part of a sheet. `runs` has the notes whose text is in runs
/// (by row and column); the others are one plain text.
let private commentBodyXml (runsOf: Map<int * int, (string * XlsxReader.RunProps) list>) (cell: XlsxCellOut) (text: string) : string =
    // (Runs that are not the note's text any more are not the note's.)
    match (if runsOf.IsEmpty then None else Map.tryFind (cell.Row, cell.Col) runsOf) with
    | Some textRuns when not textRuns.IsEmpty && (textRuns |> List.map fst |> String.concat "") = text ->
        // (Only what a run turns on is written.)
        let on (flag: bool option) = if flag = Some true then flag else None
        textRuns
        |> List.map (fun (runText, props) ->
            let run = richRunOfReader props runText
            richRunXml { run with Bold = on run.Bold; Italic = on run.Italic; Underline = on run.Underline; UnderlineDouble = on run.UnderlineDouble; Strikethrough = on run.Strikethrough })
        |> String.concat ""
    | _ -> sprintf "<t xml:space=\"preserve\">%s</t>" (escapedCellText text)

/// The file's comments part with the notes of the sheet as they are now: a
/// note that is what the file has - the same text by the same author, in
/// its cell or (rows or columns having moved) in another - is the file's
/// element as it was written (its runs' colours and font families, its
/// `shapeId`, `xr:uid`), with its cell where it is now; any other note is
/// written from the sheet; a note that is gone is left out. Authors are
/// added behind the file's. None when the part has no list of comments.
let private patchedCommentsXml (fileXml: string) (runs: ((int * int) * (string * XlsxReader.RunProps) list) list) (cells: XlsxCellOut list) : string option =
    let comments = cells |> List.choose (fun c -> c.Comment |> Option.map (fun comment -> c, comment))
    match XmlPatch.readPart fileXml with
    | Some part when not comments.IsEmpty ->
      match childNamed "authors" part.Node, childNamed "commentList" part.Node with
      | Some authorsNode, Some listNode when (XmlPatch.span authorsNode).End <= (XmlPatch.span listNode).Start ->
        let authors = ResizeArray<string>(childrenNamed "author" authorsNode |> List.map (fun author -> (XmlPatch.textOf author).Replace("\r", "")))
        let fileAuthors = authors.Count
        // The file's notes: cell, author, text, element.
        let known =
            childrenNamed "comment" listNode
            |> List.map (fun comment ->
                let author =
                    match XmlPatch.attribute comment "authorId" |> Option.map System.Int32.TryParse with
                    | Some (true, id) when id >= 0 && id < fileAuthors -> authors.[id]
                    | _ -> ""
                let text = XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "t") comment |> List.map XmlPatch.textOf |> String.concat ""
                XmlPatch.attribute comment "ref" |> Option.defaultValue "", author, text.Replace("\r", ""), comment)
            |> List.toArray
        let used = Array.create known.Length false
        let runsOf = if runs.IsEmpty then Map.empty else Map.ofList runs
        // (The file's notes as the reader has them, in the order of the part - read when a note changed.)
        let fileNotes = lazy (let read = (try XlsxReader.CellComments.parseComments fileXml with _ -> []) in if read.Length = known.Length then List.toArray read else [||])
        let authorId (author: string) =
            match authors.IndexOf author with
            | -1 -> authors.Add author; authors.Count - 1
            | index -> index
        // Each note of the sheet: the file's element for it, or none.
        let wanted = comments |> List.map (fun (cell, (text, author)) -> cell, text, (author |> Option.defaultValue ""), cellRef cell.Row cell.Col)
        let same (index: int) (text: string) (author: string) =
            let _, knownAuthor, knownText, _ = known.[index]
            not used.[index] && knownAuthor = author && knownText = text.Replace("\r", "")
        let found = System.Collections.Generic.Dictionary<string, int>()
        // (In its own cell first; then - the cells having moved - wherever the file has the same note.)
        for _, text, author, reference in wanted do
            match known |> Array.tryFindIndex (fun (knownRef, _, _, _) -> knownRef = reference) with
            | Some index when same index text author -> used.[index] <- true; found.[reference] <- index
            | _ -> ()
        for _, text, author, reference in wanted do
            if not (found.ContainsKey reference) then
                match [ 0 .. known.Length - 1 ] |> List.tryFind (fun index -> same index text author) with
                | Some index -> used.[index] <- true; found.[reference] <- index
                | None -> ()
        let elements =
            wanted
            |> List.map (fun (cell, text, author, reference) ->
                match found.TryGetValue reference with
                | true, index ->
                    let knownRef, _, _, comment = known.[index]
                    if knownRef = reference then XmlPatch.bytesOf part.Text comment
                    else XmlPatch.apply order (XmlPatch.at part comment) [ XmlPatch.SetAttribute ("ref", reference) ]
                | false, _ ->
                    // A changed note in a cell the file has a note in keeps what the file's element
                    // says of it (`shapeId`, `xr:uid`), and its runs what the file's runs say beyond
                    // what a sheet holds (`<color indexed>`, `<family>`, `<charset>`): a run with the
                    // format the file's run at its place has is written with that run's properties,
                    // and a text without runs with those of the file's last run (the note's own
                    // text, after the author's name in bold).
                    match known |> Array.tryFindIndex (fun (knownRef, _, _, _) -> knownRef = reference) with
                    | Some index when not used.[index] ->
                        used.[index] <- true
                        let _, _, _, comment = known.[index]
                        let fileRuns =
                            XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "r") comment
                            |> List.map (fun run -> XmlPatch.child "rPr" run |> Option.map (XmlPatch.bytesOf part.Text) |> Option.defaultValue "")
                        let readRuns = if index < fileNotes.Value.Length then fileNotes.Value.[index].Runs |> List.map snd else []
                        let runXml (props: string) (runText: string) = sprintf "<r>%s<t xml:space=\"preserve\">%s</t></r>" props (escapedCellText runText)
                        let body =
                            match Map.tryFind (cell.Row, cell.Col) runsOf with
                            | Some textRuns when not textRuns.IsEmpty && (textRuns |> List.map fst |> String.concat "") = text ->
                                if fileRuns.Length <> readRuns.Length then commentBodyXml runsOf cell text
                                else
                                    textRuns
                                    |> List.mapi (fun at (runText, props) ->
                                        if at < readRuns.Length && readRuns.[at] = props then runXml fileRuns.[at] runText
                                        else commentBodyXml (Map.ofList [ (cell.Row, cell.Col), [ runText, props ] ]) cell runText)
                                    |> String.concat ""
                            | _ ->
                                match List.tryLast fileRuns with
                                | Some props when props <> "" -> runXml props text
                                | _ -> commentBodyXml runsOf cell text
                        XmlPatch.apply order (XmlPatch.at part comment)
                            [ XmlPatch.SetAttribute ("authorId", string (authorId author)); XmlPatch.SetChild ("text", "<text>" + body + "</text>") ]
                    | _ -> sprintf "<comment ref=\"%s\" authorId=\"%d\"><text>%s</text></comment>" reference (authorId author) (commentBodyXml runsOf cell text))
        let added = [ for index in fileAuthors .. authors.Count - 1 -> elemText "author" authors.[index] ] |> String.concat ""
        Some (XmlPatch.applyPart order part
                [ if added <> "" then yield XmlPatch.EditChild ("authors", [ XmlPatch.Append added ], None)
                  yield XmlPatch.EditChild ("commentList", [ XmlPatch.SetChildren ("comment", elements) ], None) ])
      | _ -> None
    | _ -> None

let private commentsXml (runs: ((int * int) * (string * XlsxReader.RunProps) list) list) (cells: XlsxCellOut list) =
    let runsOf = if runs.IsEmpty then Map.empty else Map.ofList runs
    let comments = cells |> List.choose (fun c -> c.Comment |> Option.map (fun comment -> c, comment))
    if List.isEmpty comments then None
    else
        // (A note without an author is no one's: it names an empty entry of the list, not the first author.)
        let authors =
            comments
            |> List.map (fun (_, (_, author)) -> author |> Option.defaultValue "")
            |> List.distinct
        let authorIndex (author: string option) =
            authors |> List.tryFindIndex ((=) (author |> Option.defaultValue "")) |> Option.defaultValue 0
        let authorXml = authors |> List.map (elemText "author") |> String.concat ""
        let commentXml =
            comments
            |> List.map (fun (cell, (text, author)) ->
                let body = commentBodyXml runsOf cell text
                sprintf "<comment ref=\"%s\" authorId=\"%d\"><text>%s</text></comment>"
                    (cellRef cell.Row cell.Col)
                    (authorIndex author)
                    body)
            |> String.concat ""
        Some (sprintf "%s\n<comments xmlns=\"%s\"><authors>%s</authors><commentList>%s</commentList></comments>"
                  declaration Namespaces.spreadsheet authorXml commentXml)

let private vmlCommentsXml (sheetIndex: int) (cells: XlsxCellOut list) =
    let comments = cells |> List.filter (fun c -> c.Comment.IsSome)
    if List.isEmpty comments then None
    else
        let shapes =
            comments
            |> List.mapi (fun i c ->
                sprintf "<v:shape id=\"_x0000_s%d\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59.25pt;z-index:%d;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>1, 15, 0, 2, 3, 15, 4, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%d</x:Row><x:Column>%d</x:Column></x:ClientData></v:shape>"
                    (1024 + i) (i + 1) c.Row c.Col)
            |> String.concat ""
        Some (sprintf "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" data=\"%d\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>%s</xml>" sheetIndex shapes)

/// The shapes of a VML drawing (`<v:shape>`, at any depth).
let private vmlShapes (vml: XmlPatch.SourceElement) : XmlDsl.XmlNode list =
    XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = "v:shape") vml.Node

/// Whether a VML shape is the box of a note (its client data is of a "Note").
let private isNoteBox (shape: XmlDsl.XmlNode) : bool =
    not (XmlPatch.descendantsWhere (fun node -> XmlPatch.attribute node "ObjectType" = Some "Note") shape).IsEmpty

/// The cell of the note a VML shape is the box of: its row and column ((-1, -1) when it does not say);
/// None for a shape that is not a note's box.
let private noteBoxCell (shape: XmlDsl.XmlNode) : (int * int) option =
    if not (isNoteBox shape) then None
    else
        let number (name: string) =
            XmlPatch.descendantsWhere (fun node -> XmlPatch.name node = name) shape
            |> List.tryHead
            |> Option.bind (fun node -> match System.Int32.TryParse ((XmlPatch.textOf node).Trim()) with | true, value -> Some value | _ -> None)
        match number "x:Row", number "x:Column" with
        | Some row, Some col -> Some (row, col)
        | _ -> Some (-1, -1)

/// A VML drawing of the file (form controls, camera pictures, the boxes of
/// its comments) with the boxes of the sheet's comments as they are now:
/// the box of a comment that is gone is taken out, a comment without a box
/// gets one, every other shape stays as it is. None when the drawing already
/// has exactly the boxes of the comments (or is not one this can read).
let vmlWithCommentBoxes (vml: string) (commentCells: (int * int) list) : string option =
    match XmlPatch.readPart vml with
    | None -> None
    | Some drawing ->
    let wanted = Set.ofList commentCells
    let shapes = vmlShapes drawing
    let notes = shapes |> List.choose (fun shape -> noteBoxCell shape |> Option.map (fun cell -> shape, cell))
    let boxed = notes |> List.map snd
    if Set.ofList boxed = wanted && boxed.Length = wanted.Count then None
    else
        // The boxes that stay: the first of each comment's cell.
        let kept = System.Collections.Generic.HashSet<int * int>()
        let gone = notes |> List.filter (fun (_, cell) -> not (Set.contains cell wanted && kept.Add cell)) |> List.map fst
        let goneSpans = gone |> List.map XmlPatch.span |> Set.ofList
        let removals = XmlPatch.editsAt drawing.Node [ for shape in gone -> shape, [ XmlPatch.Removed ] ]
        let missing = commentCells |> List.distinct |> List.filter (fun cell -> not (kept.Contains cell))
        if missing.IsEmpty then Some (XmlPatch.applyPart (fun _ -> []) drawing removals)
        else
            // (What stays of the drawing: not in a shape that goes.)
            let staying (node: XmlDsl.XmlNode) =
                let span = XmlPatch.span node
                not (goneSpans |> Set.exists (fun goneSpan -> span.Start >= goneSpan.Start && span.End <= goneSpan.End))
            let rest = XmlPatch.descendantsWhere staying drawing.Node
            // Shape ids of a drawing are those of its block of 1024 (`<o:idmap data>`).
            let highest =
                rest |> List.collect (fun node -> XmlPatch.attributes node |> Map.toList |> List.choose (snd >> VmlIds.shapeNumber)) |> List.fold max 0
            let first =
                if highest > 0 then highest + 1
                else
                    let block =
                        rest |> List.tryPick (fun node -> if XmlPatch.name node = "o:idmap" then XmlPatch.attribute node "data" |> Option.bind VmlIds.firstBlock else None)
                    (defaultArg block 1) * 1024 + 1
            let shapeType =
                if rest |> List.exists (fun node -> XmlPatch.attribute node "id" = Some "_x0000_t202") then ""
                else "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>"
            let added =
                missing
                |> List.mapi (fun i (row, col) ->
                    sprintf "<v:shape id=\"_x0000_s%d\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59.25pt;z-index:%d;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%d, 15, %d, 2, %d, 15, %d, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%d</x:Row><x:Column>%d</x:Column></x:ClientData></v:shape>"
                        (first + i) (shapes.Length + i + 1) (col + 1) (max 0 (row - 1)) (col + 3) (row + 3) row col)
                |> String.concat ""
            Some (XmlPatch.applyPart (fun _ -> []) drawing (removals @ [ XmlPatch.Append (shapeType + added) ]))

let private tableXml id (table: XlsxTableOut) =
    let refText = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1) (table.EndCol - 1))
    let cols =
        let count = table.EndCol - table.StartCol + 1
        [ 0 .. count - 1 ]
        |> List.map (fun i ->
            let name =
                table.Columns
                |> List.tryItem i
                |> Option.defaultValue (sprintf "Column%d" (i + 1))
            // What the column shows in the totals row: Excel checks it against the cell.
            match (if table.TotalsRowCount > 0 then List.tryItem i table.Totals else None) with
            | Some ("label", label) -> sprintf "<tableColumn id=\"%d\" %s %s/>" (i + 1) (attr "name" name) (attr "totalsRowLabel" label)
            | Some ("function", fn) -> sprintf "<tableColumn id=\"%d\" %s %s/>" (i + 1) (attr "name" name) (attr "totalsRowFunction" fn)
            | Some ("custom", formula) ->
                sprintf "<tableColumn id=\"%d\" %s totalsRowFunction=\"custom\"><totalsRowFormula>%s</totalsRowFormula></tableColumn>" (i + 1) (attr "name" name) (escapeXml formula)
            | _ -> sprintf "<tableColumn id=\"%d\" %s/>" (i + 1) (attr "name" name))
        |> String.concat ""
    let autoFilter =
        // (The filter covers the header and the data, not the totals row.)
        let filterRef = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1 - max 0 table.TotalsRowCount) (table.EndCol - 1))
        table.AutoFilter
        |> Option.map (fun _ -> sprintf "<autoFilter ref=\"%s\"/>" filterRef)
        |> Option.defaultValue ""
    let style =
        table.Style
        |> Option.map (fun s ->
            sprintf "<tableStyleInfo %s showFirstColumn=\"%d\" showLastColumn=\"%d\" showRowStripes=\"%d\" showColumnStripes=\"%d\"/>"
                (attr "name" s.Name) (if s.ShowFirstColumn then 1 else 0) (if s.ShowLastColumn then 1 else 0) (if s.ShowRowStripes then 1 else 0) (if s.ShowColumnStripes then 1 else 0))
        |> Option.defaultValue ""
    sprintf "%s\n<table xmlns=\"%s\" id=\"%d\" %s %s ref=\"%s\" headerRowCount=\"%d\" totalsRowCount=\"%d\">%s<tableColumns count=\"%d\">%s</tableColumns>%s</table>"
        declaration Namespaces.spreadsheet id (attr "name" table.Name) (attr "displayName" table.Name) refText table.HeaderRowCount table.TotalsRowCount autoFilter (table.EndCol - table.StartCol + 1) cols style

// ---- a table of the file, changed: the file's part with what changed replaced ----

// ---------------------------------------------------------------------------
// A sheet of the file with some of its cells changed (SheetExtras.Patch)
// ---------------------------------------------------------------------------
// The part of the file is changed by edits of its span tree (XmlPatch): the `<c>` elements that changed
// replaced (added, taken out), the attributes of the rows that changed set, and the top-level elements
// whose content changed made from the sheet as it is now. Everything else is the file's text: sheetPr,
// sheetViews, the rows with their attributes, shared formulas, ignoredErrors, extLst, and every
// relationship id.

/// The cells of a sheet part that name a shared string.
let private sharedStringCells (part: XmlPatch.SourceElement) : int =
    childrenNamed "sheetData" part.Node
    |> List.sumBy (fun data -> childrenNamed "row" data |> List.sumBy (fun row -> childrenNamed "c" row |> List.filter (fun cell -> XmlPatch.attribute cell "t" = Some "s") |> List.length))

type private SheetPatchPlan = {
    /// The part, read in outline: its rows without their cells, but for the rows the patch changes cells of.
    Source: XmlPatch.SourceElement
    /// Its `<sheetData>`.
    Data: XmlDsl.XmlNode
    /// The rows of the file (numbered from 0), ascending.
    Rows: (int * XmlDsl.XmlNode) array
    /// The cells of the rows that have a changed cell: column (from 0) and element, ascending.
    RowCells: System.Collections.Generic.Dictionary<int, (int * XmlDsl.XmlNode) array>
    /// The sections to be made anew can be put into the part as it is.
    SectionsSound: bool
    /// Cells that take over the formula their group shares (its cell is
    /// written anew), by row * 16384 + column: the group's new range and its `si`.
    Rehome: System.Collections.Generic.Dictionary<float, string * string>
}

/// The sections of a sheet a patch can write anew.
let private patchableSections =
    [ "sheetPr"; "sheetViews"; "sheetFormatPr"; "cols"; "autoFilter"; "sortState"; "mergeCells"; "conditionalFormatting"; "dataValidations"
      "hyperlinks"; "printOptions"; "pageMargins"; "pageSetup"; "headerFooter"; "rowBreaks"; "colBreaks" ]

/// How the part of the file is to be patched; None when it cannot be (the
/// sheet is then written from its cells): a part that is not laid out as
/// expected, a changed cell whose formula other cells share over a block.
/// `SectionsSound` is false when a section to be made anew cannot be put
/// into the part as it is (rules and validations of the x14 list): the rows
/// can still be the file's.
let private sheetPatchPlan (dxfCount: int) (sheet: XlsxSheetOut) (patch: XlsxReader.SheetPatch) (source: string) : SheetPatchPlan option =
    match readWorksheet source with
    | None -> None
    | Some part ->
        let top = XmlPatch.elements part.Node
        // (Whether the part has an element of these names, outside its rows.)
        let holds (names: string list) =
            top |> List.exists (fun kid -> XmlPatch.name kid <> "sheetData" && not (XmlPatch.descendantsWhere (fun node -> List.contains (XmlPatch.name node) names) kid).IsEmpty)
        let sectionsSound =
            patch.Sections |> List.forall (fun name -> List.contains name patchableSections)
            // (Rules and validations of the x14 list are in the sheet's extensions.)
            && (not (List.contains "conditionalFormatting" patch.Sections)
                || (not (holds [ "x14:conditionalFormatting"; "x14:conditionalFormattings" ]) && snd (conditionalFormattingParts dxfCount 1 sheet.ConditionalFormats) = ""))
            && (not (List.contains "dataValidations" patch.Sections)
                || (not (holds [ "x14:dataValidation"; "x14:dataValidations" ]) && dataValidationsExtXml sheet.DataValidations = ""))
            // (The links of the file name their relationships as `r:id`, or the list is made with the rest of the part.)
            && (not (List.contains "hyperlinks" patch.Sections)
                || (childrenNamed "hyperlinks" part.Node
                    |> List.forall (fun links -> XmlPatch.allAttributes links |> List.forall (fun (name, _) -> XmlDsl.Ns.stripNs name <> "id" || name = "r:id"))))
        match childrenNamed "sheetData" part.Node with
        | [ data ] ->
            let rows =
                childrenNamed "row" data
                |> List.map (fun row -> (match XmlPatch.attribute row "r" |> Option.map System.Int32.TryParse with Some (true, number) -> number - 1 | _ -> -1), row)
                |> List.toArray
            let mutable ascending = true
            for index in 0 .. rows.Length - 1 do
                if fst rows.[index] < 0 || (index > 0 && fst rows.[index] <= fst rows.[index - 1]) then ascending <- false
            if not ascending then None
            else
                let byNumber = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
                for number, row in rows do byNumber.[number] <- row
                let rowCells = System.Collections.Generic.Dictionary<int, (int * XmlDsl.XmlNode) array>()
                let mutable cellsSound = true
                let scanRow (row: int) =
                    if cellsSound && not (rowCells.ContainsKey row) then
                        match byNumber.TryGetValue row with
                        | true, node when not (isSelfClosed node) ->
                            match XmlPatch.whole part node with
                            | Some whole ->
                                let placed =
                                    childrenNamed "c" whole.Node
                                    |> List.map (fun cell -> (match XmlPatch.attribute cell "r" |> Option.bind CellRef.parse with Some (cellRow, col) when cellRow = row -> col | _ -> -1), cell)
                                    |> List.toArray
                                for index in 0 .. placed.Length - 1 do
                                    if fst placed.[index] < 0 || (index > 0 && fst placed.[index] <= fst placed.[index - 1]) then cellsSound <- false
                                rowCells.[row] <- placed
                            | None -> cellsSound <- false
                        | _ -> rowCells.[row] <- [||]
                for row, _, _ in patch.Cells do scanRow row
                // A changed cell whose formula other cells share (`<f t="shared" ref>`): the
                // first of those cells that stays as it is holds the formula from now on.
                let key (row: int) (col: int) = float row * 16384.0 + float col
                let rehome = System.Collections.Generic.Dictionary<float, string * string>()
                if cellsSound then
                    let replaced = System.Collections.Generic.HashSet<float>()
                    for row, col, valueOnly in patch.Cells do
                        if not valueOnly then replaced.Add (key row col) |> ignore
                    let sharedFormula (cell: XmlDsl.XmlNode) =
                        XmlPatch.child "f" cell |> Option.filter (fun f -> XmlPatch.attribute f "t" = Some "shared")
                    for row, col, valueOnly in patch.Cells do
                        if cellsSound && not valueOnly then
                            match rowCells.[row] |> Array.tryFind (fun (known, _) -> known = col) with
                            | Some (_, cell) ->
                                match sharedFormula cell |> Option.bind (fun f -> XmlPatch.attribute f "ref" |> Option.map (fun range -> f, range)) with
                                | Some (f, range) ->
                                    match XmlPatch.attribute f "si", CellRef.parseRange range with
                                    | Some si, Some ((top, left), (bottom, right)) when left = right || top = bottom ->
                                        let mutable found = false
                                        let mutable at = top
                                        while cellsSound && not found && at <= bottom do
                                            if byNumber.ContainsKey at then
                                                scanRow at
                                                if cellsSound then
                                                    for known, other in rowCells.[at] do
                                                        if not found && known >= left && known <= right && not (replaced.Contains (key at known)) then
                                                            match sharedFormula other with
                                                            | Some formula when XmlPatch.attribute formula "si" = Some si && (XmlPatch.attribute formula "ref").IsNone ->
                                                                found <- true
                                                                rehome.[key at known] <- (cellRef at known + ":" + cellRef bottom right, si)
                                                            | _ -> ()
                                            at <- at + 1
                                    | _ -> cellsSound <- false
                                | None -> ()
                            | None -> ()
                    if cellsSound && rehome.Count > 0 then
                        // (The cell that takes the formula over has to hold one.)
                        let holding = System.Collections.Generic.HashSet<float>()
                        for cell in sheet.Cells do
                            match cell.Value with
                            | XlsxFormula (formula, _) when formula <> "" -> (let position = key cell.Row cell.Col in if rehome.ContainsKey position then holding.Add position |> ignore)
                            | _ -> ()
                        if holding.Count <> rehome.Count then cellsSound <- false
                if not cellsSound then None
                else
                    // (The rows whose cells change, read whole.)
                    let source = XmlPatch.withWhole part [ for KeyValue (row, cells) in rowCells do if cells.Length > 0 then yield (XmlPatch.span byNumber.[row]).Start ]
                    match childNamed "sheetData" source.Node with
                    | Some data ->
                        Some { Source = source; Data = data; Rows = rows; RowCells = rowCells; SectionsSound = sectionsSound; Rehome = rehome }
                    | None -> None
        | _ -> None

/// The attributes of a row as the sheet has it now (`known`: the file's): each with its value, or None for
/// none.
let private rowAttributes (known: Map<string, string>) (row: XlsxReader.SheetRowPatch) : (string * string option) list =
    [ yield "hidden", (if row.Hidden then Some "1" else None)
      if row.HeightChanged then
          match row.Height with
          | Some height ->
              let same = match Map.tryFind "ht" known |> Option.bind tryParseNumber with Some was -> abs (was - height) < 0.01 | None -> false
              if not same then yield! [ "ht", Some (invariantFixed height); "customHeight", Some "1" ]
          | None -> yield! [ "ht", None; "customHeight", None ]
      yield "outlineLevel", (if row.Level > 0 then Some (string row.Level) else None)
      yield "collapsed", (if row.Collapsed then Some "1" else None) ]

/// The edits of the file's `<sheetData>` that give it the changed cells and rows of `sheet`; the first and
/// last row and column written anew (the last row -1 when no cell was); and how many cells of the file that
/// name a shared string are gone.
let private sheetDataEdits (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (dynamicCm: int)
                           (sheet: XlsxSheetOut) (patch: XlsxReader.SheetPatch) (plan: SheetPatchPlan) : XmlPatch.ElementEdit list * (int * int * int * int) * int =
    let dataTables = sheet.Extras.DataTables |> List.map (fun table -> CellRef.firstCell table.Ref, table) |> Map.ofList
    let arrayFormulas =
        sheet.Extras.ArrayFormulas
        |> List.map (fun (cell, range, dynamic) -> cell, (range, (if dynamic && dynamicCm > 0 then " cm=\"" + string dynamicCm + "\"" else "")))
        |> Map.ofList
    let key (row: int) (col: int) = float row * 16384.0 + float col
    // What is done to a cell: 0 - written anew, 1 - the file's with another result, 2 - the file's (it takes a shared formula over).
    let changedRows = System.Collections.Generic.Dictionary<int, ResizeArray<int * int>>()
    let wanted = System.Collections.Generic.HashSet<float>()
    let change (row: int) (col: int) (mode: int) =
        if wanted.Add (key row col) then
            match changedRows.TryGetValue row with
            | true, known -> known.Add ((col, mode))
            | false, _ ->
                let known = ResizeArray<int * int>()
                known.Add ((col, mode))
                changedRows.[row] <- known
    for row, col, valueOnly in patch.Cells do change row col (if valueOnly then 1 else 0)
    for KeyValue(position, _) in plan.Rehome do
        let row = int (floor (position / 16384.0))
        change row (int (position - float row * 16384.0)) 2
    let rowPatches = System.Collections.Generic.Dictionary<int, XlsxReader.SheetRowPatch>()
    for row in patch.Rows do rowPatches.[row.Row] <- row
    let outs = System.Collections.Generic.Dictionary<float, XlsxCellOut>()
    if wanted.Count > 0 then
        for cell in sheet.Cells do
            let position = key cell.Row cell.Col
            if wanted.Contains position then outs.[position] <- cell
    // The cells written, for the sheet's dimension.
    let mutable minRow = System.Int32.MaxValue
    let mutable minCol = System.Int32.MaxValue
    let mutable maxRow = -1
    let mutable maxCol = -1
    let written (row: int) (col: int) =
        if row < minRow then minRow <- row
        if row > maxRow then maxRow <- row
        if col < minCol then minCol <- col
        if col > maxCol then maxCol <- col
    // (The file's cells that named a shared string and do not any more.)
    let mutable sharedGone = 0
    let goes (cell: XmlDsl.XmlNode) = if XmlPatch.attribute cell "t" = Some "s" then sharedGone <- sharedGone + 1
    // A cell of the sheet as it is now ("" for none).
    let cellXml (row: int) (col: int) =
        match outs.TryGetValue (key row col) with
        | true, cell ->
            let sb = StringBuilder()
            appendCellXml sb styles sharedIndex dataTables arrayFormulas (string (row + 1)) cell
            if sb.Length > 0 then written row col
            sb.ToString()
        | false, _ -> ""
    // The file's cell holding the formula its group shares from now on.
    let rehomed (row: int) (col: int) (cell: XmlDsl.XmlNode) : XmlPatch.ElementEdit list =
        match plan.Rehome.TryGetValue (key row col), outs.TryGetValue (key row col) with
        | (true, (reference, si)), (true, { Value = XlsxFormula (formula, _) }) ->
            match XmlPatch.child "f" cell with
            | Some f when XmlPatch.name f = "f" && isSelfClosed f ->
                [ XmlPatch.SetChild ("f", "<f t=\"shared\" ref=\"" + reference + "\" si=\"" + si + "\">" + escapedFast (fileFormula formula) + "</f>") ]
            | _ -> []
        | _ -> []
    // The edits of a changed cell of the file.
    let cellEdits (row: int) (col: int) (mode: int) (cell: XmlDsl.XmlNode) : XmlPatch.ElementEdit list =
        if mode = 2 then rehomed row col cell
        else
            match outs.TryGetValue (key row col) with
            | false, _ ->
                goes cell
                [ XmlPatch.Replaced "" ]
            | true, out ->
                match mode = 1, out.Value with
                | true, XlsxFormula (_, cached) when (XmlPatch.child "f" cell).IsSome ->
                    // The formula is the file's (shared, array ...): only its result is new.
                    let cellType, valueXml =
                        match cached with
                        | Some value -> formulaResultXml sharedIndex value
                        | None -> "", ""
                    if XmlPatch.attribute cell "t" <> Some cellType then goes cell
                    written row col
                    rehomed row col cell
                    @ [ attributeEdit "t" (if cellType = "" then None else Some cellType)
                        XmlPatch.SetChildren ("v", [])
                        XmlPatch.Append valueXml ]
                | _ ->
                    goes cell
                    [ XmlPatch.Replaced (cellXml row col) ]
    let byNumber = System.Collections.Generic.Dictionary<int, XmlDsl.XmlNode>()
    for number, row in plan.Rows do byNumber.[number] <- row
    // The edits of the rows (by where each starts), and the rows put in after the file's.
    let rowEdits = System.Collections.Generic.Dictionary<int, ResizeArray<XmlPatch.ElementEdit>>()
    let editRow (node: XmlDsl.XmlNode) (edits: XmlPatch.ElementEdit list) =
        let start = (XmlPatch.span node).Start
        match rowEdits.TryGetValue start with
        | true, known -> known.AddRange edits
        | false, _ -> rowEdits.[start] <- ResizeArray<XmlPatch.ElementEdit>(edits)
    let appended = ResizeArray<XmlPatch.ElementEdit>()
    let noChanges = ResizeArray<int * int>()
    let touched = System.Collections.Generic.HashSet<int>(changedRows.Keys)
    for row in rowPatches.Keys do touched.Add row |> ignore
    let mutable next = 0
    for row in touched |> Seq.sort do
        let changes = match changedRows.TryGetValue row with | true, known -> known | false, _ -> noChanges
        let rowPatch = match rowPatches.TryGetValue row with | true, known -> Some known | false, _ -> None
        let sorted = changes.ToArray() |> Array.sortBy fst
        match byNumber.TryGetValue row with
        | true, node ->
            let attributes =
                match rowPatch with
                | Some known -> rowAttributes (XmlPatch.attributes node) known |> List.map (fun (name, value) -> attributeEdit name value)
                | None -> []
            if sorted.Length = 0 then editRow node attributes
            else
                let cells = match plan.RowCells.TryGetValue row with | true, cells -> cells | false, _ -> [||]
                // (A cell where the row had none may lie outside the row's `spans`.)
                let added = sorted |> Array.exists (fun (col, _) -> not (cells |> Array.exists (fun (known, _) -> known = col)))
                let edits = ResizeArray<XmlPatch.ElementEdit>(attributes)
                if added then edits.Add (XmlPatch.RemoveAttribute "spans")
                // (New cells go before the first cell of the file after them.)
                let pending = StringBuilder()
                let before () =
                    let xml = pending.ToString()
                    pending.Clear() |> ignore
                    if xml = "" then [] else [ XmlPatch.InsertBefore xml ]
                let mutable i = 0
                let mutable j = 0
                while i < cells.Length || j < sorted.Length do
                    if j >= sorted.Length || (i < cells.Length && fst cells.[i] < fst sorted.[j]) then
                        if pending.Length > 0 then edits.Add (XmlPatch.EditAt ((XmlPatch.span (snd cells.[i])).Start, before ()))
                        i <- i + 1
                    elif i < cells.Length && fst cells.[i] = fst sorted.[j] then
                        let cell = snd cells.[i]
                        let inner = before () @ cellEdits row (fst sorted.[j]) (snd sorted.[j]) cell
                        if not inner.IsEmpty then edits.Add (XmlPatch.EditAt ((XmlPatch.span cell).Start, inner))
                        i <- i + 1
                        j <- j + 1
                    else
                        pending.Append(cellXml row (fst sorted.[j])) |> ignore
                        j <- j + 1
                if pending.Length > 0 then edits.Add (XmlPatch.Append (pending.ToString()))
                editRow node (List.ofSeq edits)
        | false, _ ->
            // A new row (as the sheet has it), before the first row of the file after it.
            let inner = sorted |> Array.map (fun (col, _) -> cellXml row col) |> String.concat ""
            let attributes =
                match rowPatch with
                | Some known -> rowAttributes Map.empty known |> List.choose (fun (name, value) -> value |> Option.map (fun value -> " " + attr name value)) |> String.concat ""
                | None -> ""
            // (A row that says nothing is not written.)
            if inner <> "" || attributes <> "" then
                let xml = "<row r=\"" + string (row + 1) + "\"" + attributes + ">" + inner + "</row>"
                while next < plan.Rows.Length && fst plan.Rows.[next] < row do next <- next + 1
                if next < plan.Rows.Length then editRow (snd plan.Rows.[next]) [ XmlPatch.InsertBefore xml ]
                else appended.Add (XmlPatch.Append xml)
    let edits = [ for KeyValue (start, edits) in rowEdits do yield XmlPatch.EditAt (start, List.ofSeq edits) ] @ List.ofSeq appended
    edits, (minRow, minCol, maxRow, maxCol), sharedGone

/// The file's `<sheetData>` with these edits.
let private patchedSheetData (plan: SheetPatchPlan) (edits: XmlPatch.ElementEdit list) : string =
    XmlPatch.apply order (XmlPatch.at plan.Source plan.Data) edits

// ---- what a part of the file says beyond what a sheet holds --------------------

/// `<sheetPr>` as the file has it (its code name, what else it says), with
/// the tab colour, the outline's summary sides and fit-to-page of the sheet
/// as it is now ("" when it then says nothing).
let private sheetPrFromSource (was: XmlPatch.SourceElement option) (sheet: XlsxSheetOut) : string =
    let source = was |> Option.orElse (XmlPatch.read "<sheetPr/>")
    match source with
    | None -> ""
    | Some source ->
        // A child that holds flags (attributes): the edits that make them say what the sheet does - taken
        // out when it says nothing then, made when the file has none.
        let flags (name: string) (changes: Map<string, string> -> XmlPatch.ElementEdit list) : XmlPatch.ElementEdit list =
            let known = childNamed name source.Node
            let attributes = known |> Option.map XmlPatch.attributes |> Option.defaultValue Map.empty
            let edits = changes attributes
            let after =
                edits |> List.fold (fun (now: Map<string, string>) edit ->
                    match edit with
                    | XmlPatch.SetAttribute (attribute, value) -> Map.add attribute value now
                    | XmlPatch.RemoveAttribute attribute -> Map.remove attribute now
                    | _ -> now) attributes
            let bare = after.IsEmpty && (known |> Option.forall (fun node -> List.isEmpty (XmlPatch.elements node)))
            match known with
            | Some _ when bare -> [ XmlPatch.RemoveChild name ]
            | Some _ -> if edits.IsEmpty then [] else [ XmlPatch.EditChild (name, edits, None) ]
            | None when bare -> []
            | None -> [ XmlPatch.SetChild (name, "<" + name + (after |> Map.toList |> List.map (fun (attribute, value) -> " " + attr attribute value) |> String.concat "") + "/>") ]
        let side (name: string) (below: bool) (attributes: Map<string, string>) =
            let known = match Map.tryFind name attributes with Some "0" | Some "false" -> false | _ -> true
            if known = below then [] else [ attributeEdit name (if below then None else Some "0") ]
        let tabColor =
            match tabColorXml sheet, childNamed "tabColor" source.Node with
            | "", Some _ -> [ XmlPatch.RemoveChild "tabColor" ]
            | "", None -> []
            // (The file's element while it says the same.)
            | made, Some known when (XmlPatch.read made |> Option.map (fun made -> XmlPatch.attributes made.Node)) = Some (XmlPatch.attributes known) -> []
            | made, _ -> [ XmlPatch.SetChild ("tabColor", made) ]
        let fits = sheet.PageSetup |> Option.exists (fun setup -> setup.FitToPage)
        let edits =
            tabColor
            @ flags "outlinePr" (fun attributes -> side "summaryBelow" sheet.View.SummaryBelow attributes @ side "summaryRight" sheet.View.SummaryRight attributes)
            @ flags "pageSetUpPr" (fun attributes -> if isOn (Map.tryFind "fitToPage" attributes) = fits then [] else [ attributeEdit "fitToPage" (if fits then Some "1" else None) ])
        let made = XmlPatch.apply order source edits
        // (A sheetPr that says nothing is not written.)
        if (XmlPatch.attributes source.Node).IsEmpty && XmlPatch.isEmpty made then "" else made

type private SheetMoves = (bool * bool * int * int) list

/// Where rows (or columns) `low` .. `high` are after the edits; None when all of them were deleted.
let private movedSpan (moves: SheetMoves) (onRows: bool) (low: int) (high: int) : (int * int) option =
    let mutable low = low
    let mutable high = high
    let mutable alive = true
    for rows, insert, at, count in moves do
        if alive && rows = onRows then
            if insert then
                if low >= at then
                    low <- low + count
                    high <- high + count
                elif high >= at then high <- high + count
            else
                let stop = at + count
                let newLow = if low >= stop then low - count elif low >= at then at else low
                let newHigh = if high >= stop then high - count elif high >= at then at - 1 else high
                if newHigh < newLow then alive <- false
                else
                    low <- newLow
                    high <- newHigh
    if alive && high <= (if onRows then 1048575 else 16383) then Some (low, high) else None

/// A cell or range reference ("B2", "B2:D9") after the edits; None when its cells are gone.
let private movedRefText (moves: SheetMoves) (reference: string) : string option =
    if moves.IsEmpty then Some reference
    else
        let ends = CellRef.corners (reference.Replace("$", ""))
        if ends.Length > 2 then None
        else
            match CellRef.parseRange (CellRef.ofCorners ends) with
            | Some ((top, left), (bottom, right)) ->
                match movedSpan moves true top bottom, movedSpan moves false left right with
                | Some (top, bottom), Some (left, right) ->
                    Some (if ends.Length = 1 then cellRef top left else cellRef top left + ":" + cellRef bottom right)
                | _ -> None
            | _ -> None

/// A `<sortState>` (the range that was sorted, each `<sortCondition>` the column it was
/// sorted by) after the edits: every range where its cells are now; "" when one of them is gone.
let private movedSortState (moves: SheetMoves) (was: XmlPatch.SourceElement option) : string =
    match was with
    | None -> ""
    | Some was when moves.IsEmpty -> XmlPatch.bytes was
    | Some was ->
        let moved =
            XmlPatch.descendantsWhere (fun node -> (XmlPatch.attribute node "ref").IsSome) was.Node
            |> List.map (fun node -> node, movedRefText moves (XmlPatch.attribute node "ref" |> Option.defaultValue ""))
        if moved |> List.exists (fun (_, reference) -> reference.IsNone) then ""
        else XmlPatch.apply order was (XmlPatch.editsAt was.Node [ for node, reference in moved -> node, [ XmlPatch.SetAttribute ("ref", reference.Value) ] ])

/// An element whose children (`child`) name cells by `sqref`, after the
/// edits: a child whose cells are all gone is taken out, and "" is what is
/// left of an element without children.
let private withMovedChildren (moves: SheetMoves) (child: string) (was: XmlPatch.SourceElement option) : string =
    match was with
    | None -> ""
    | Some was when moves.IsEmpty -> XmlPatch.bytes was
    | Some was ->
        let moved =
            XmlPatch.elements was.Node
            |> List.filter (fun node -> XmlPatch.localName node = child)
            |> List.map (fun node ->
                node, XmlPatch.attribute node "sqref" |> Option.map (fun sqref -> CellRef.ranges sqref |> List.choose (movedRefText moves)))
        if moved |> List.forall (fun (_, ranges) -> ranges = Some []) then ""
        else
            let targets =
                moved |> List.choose (fun (node, ranges) ->
                    match ranges with
                    | None -> None
                    | Some [] -> Some (node, [ XmlPatch.Removed ])
                    | Some ranges -> Some (node, [ XmlPatch.SetAttribute ("sqref", CellRef.ofRanges ranges) ]))
            XmlPatch.apply order was (XmlPatch.editsAt was.Node targets)

/// The edits of a `<sheetViews>` that make its first view say whether the sheet is the one on display
/// (`tabSelected`): two sheets that say so are a group to Excel. None when it says so already.
let private tabSelectedEdits (active: bool) (views: XmlDsl.XmlNode) : XmlPatch.ElementEdit list =
    match XmlPatch.elements views |> List.filter (fun node -> XmlPatch.localName node = "sheetView") with
    | view :: _ when isOn (XmlPatch.attribute view "tabSelected") <> active ->
        [ XmlPatch.EditNth (XmlPatch.name view, 0, [ attributeEdit "tabSelected" (if active then Some "1" else None) ]) ]
    | _ -> []

/// A `<sheetViews>` whose first view says whether the sheet is the one on display.
let private viewsWithTabSelected (active: bool) (views: string) : string =
    match XmlPatch.read views with
    | Some element ->
        match tabSelectedEdits active element.Node with
        | [] -> views
        | edits -> XmlPatch.apply order element edits
    | None -> views

/// The text of a sheet part with its first view saying whether the sheet is
/// the one on display. The text itself when it already says the right thing.
let private withTabSelected (active: bool) (xml: string) : string =
    match XmlPatch.readPartExcept (fun name -> name = "sheetData") xml with
    | Some part ->
        match childrenNamed "sheetViews" part.Node with
        | views :: _ ->
            match tabSelectedEdits active views with
            | [] -> xml
            | edits -> XmlPatch.applyPart order part [ XmlPatch.EditChild (XmlPatch.name views, edits, None) ]
        | [] -> xml
    | None -> xml

/// `<sheetViews>` as the file has it - the selection, where the view is
/// scrolled to, what else it says - with the switches and the zoom of the
/// sheet as it is now (and what names a cell where that cell is after
/// `moves`); with other panes than the file's, the views are made anew.
let private sheetViewsFromSource (was: XmlPatch.SourceElement option) (moves: SheetMoves) (sheet: XlsxSheetOut) : string =
    let made = sheetViewsXml sheet false
    let madeRoot = XmlPatch.read made
    let paneKey (views: XmlDsl.XmlNode option) =
        match views |> Option.bind (fun views -> XmlPatch.descendants "pane" views |> List.tryHead) with
        | Some pane -> String.concat "," [ for name in [ "xSplit"; "ySplit"; "state" ] -> XmlPatch.attribute pane name |> Option.defaultValue "" ]
        | None -> ""
    match was |> Option.bind (fun was -> XmlPatch.descendants "sheetView" was.Node |> List.tryHead |> Option.map (fun view -> was, view)) with
    | Some (was, view) when paneKey (Some was.Node) = paneKey (madeRoot |> Option.map (fun root -> root.Node)) ->
        let madeView = madeRoot |> Option.bind (fun root -> XmlPatch.descendants "sheetView" root.Node |> List.tryHead) |> Option.map XmlPatch.attributes |> Option.defaultValue Map.empty
        let viewEdits =
            [ for name in [ "showFormulas"; "showGridLines"; "showRowColHeaders"; "showZeros"; "showOutlineSymbols"; "zoomScale" ] do
                // (Left as the file has it where that says the same.)
                let fileDefault = if name = "zoomScale" then "100" elif name = "showFormulas" then "0" else "1"
                let normal (value: string option) = match value with Some "true" -> "1" | Some "false" -> "0" | Some value -> value | None -> fileDefault
                if normal (XmlPatch.attribute view name) <> normal (Map.tryFind name madeView) then yield attributeEdit name (Map.tryFind name madeView) ]
        let targets = ResizeArray<XmlDsl.XmlNode * XmlPatch.ElementEdit list>()
        targets.Add ((view, viewEdits))
        if not moves.IsEmpty then
            // A selection of cells that are gone is not kept; a view scrolled to one starts where it starts.
            for selection in XmlPatch.descendants "selection" was.Node do
                let active = XmlPatch.attribute selection "activeCell" |> Option.map (movedRefText moves)
                let ranges = XmlPatch.attribute selection "sqref" |> Option.map (fun sqref -> CellRef.ranges sqref |> List.map (movedRefText moves))
                match active, ranges with
                | Some None, _ -> targets.Add ((selection, [ XmlPatch.Removed ]))
                | _, Some parts when parts |> List.exists Option.isNone -> targets.Add ((selection, [ XmlPatch.Removed ]))
                | _ ->
                    targets.Add ((selection,
                                  [ match active with Some (Some cell) -> yield XmlPatch.SetAttribute ("activeCell", cell) | _ -> ()
                                    match ranges with Some parts -> yield XmlPatch.SetAttribute ("sqref", CellRef.ofRanges (List.choose id parts)) | None -> () ]))
            for node in XmlPatch.descendantsWhere (fun node -> XmlPatch.localName node = "sheetView" || XmlPatch.localName node = "pane") was.Node do
                match XmlPatch.attribute node "topLeftCell" with
                | Some cell -> targets.Add ((node, [ attributeEdit "topLeftCell" (movedRefText moves cell) ]))
                | None -> ()
        XmlPatch.apply order was (XmlPatch.editsAt was.Node (List.ofSeq targets))
    | _ -> made

/// `<sheetFormatPr>` as the file has it, with the outline levels of the sheet as it is now.
let private sheetFormatFromSource (was: XmlPatch.SourceElement option) (sheet: XlsxSheetOut) : string =
    match was with
    | None -> sheetFormatXml sheet
    | Some was ->
        let deepest (axis: (int * float option * int * bool) list) = axis |> List.fold (fun level (_, _, own, _) -> max level own) 0
        let level (name: string) (value: int) =
            let known = match XmlPatch.attribute was.Node name |> Option.map System.Int32.TryParse with Some (true, level) -> level | _ -> 0
            if known = value then [] else [ attributeEdit name (if value > 0 then Some (string value) else None) ]
        XmlPatch.apply order was (level "outlineLevelRow" (deepest sheet.Extras.RowAxis) @ level "outlineLevelCol" (deepest sheet.Extras.ColumnAxis))

/// The part of the file with the changed cells, rows and sections of `sheet`
/// (and its first view saying whether it is the sheet on display: `active`);
/// and how many of the file's cells that name a shared string are gone.
let private patchedSheetXml (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (dynamicCm: int) (sheetIndex: int) (active: bool)
                            (sheet: XlsxSheetOut) (patch: XlsxReader.SheetPatch) (plan: SheetPatchPlan) (added: (string * string) list) : string * int =
    let dataEdits, (minRow, minCol, maxRow, maxCol), sharedGone = sheetDataEdits styles sharedIndex dynamicCm sheet patch plan
    // (`added`: elements the part did not have, put where the schema has them - the
    // `<legacyDrawing>` of a sheet that got its first note.)
    let patch = if added.IsEmpty then patch else { patch with Sections = patch.Sections @ (added |> List.map fst) }
    // The sections made anew.
    let sourceElement (name: string) = sourceChild name plan.Source
    let sectionXml (name: string) =
        match added |> List.tryPick (fun (known, xml) -> if known = name then Some xml else None) with
        | Some xml -> xml
        | None ->
        match name with
        | "sheetPr" -> sheetPrFromSource (sourceElement "sheetPr") sheet
        | "sheetViews" -> viewsWithTabSelected active (sheetViewsFromSource (sourceElement "sheetViews") [] sheet)
        | "sheetFormatPr" -> sheetFormatFromSource (sourceElement "sheetFormatPr") sheet
        | "cols" -> colsXml styles sheet
        | "autoFilter" -> autoFilterXml styles.Dxfs.Length sheet
        | "mergeCells" -> mergeCellsXmlFull sheet.MergedCells
        | "conditionalFormatting" -> fst (conditionalFormattingParts styles.Dxfs.Length sheetIndex sheet.ConditionalFormats)
        | "dataValidations" -> dataValidationsXml sheet.DataValidations
        | "printOptions" -> (let printOptions, _, _, _ = pageParts sheet in printOptions)
        | "pageMargins" -> (let _, margins, _, _ = pageParts sheet in margins)
        | "pageSetup" -> (let _, _, pageSetup, _ = pageParts sheet in pageSetup)
        | "headerFooter" -> (let _, _, _, headerFooter = pageParts sheet in headerFooter)
        | "rowBreaks" -> breaksXml "rowBreaks" (sourceElement "rowBreaks") sheet.RowBreaks
        | "colBreaks" -> breaksXml "colBreaks" (sourceElement "colBreaks") sheet.ColBreaks
        // ("sortState": what the file says of a sort is not what the rows are sorted by now.)
        | _ -> ""
    // (Every element of that name is replaced by the new ones; one the file did not have goes where the schema has it.)
    let sections = patch.Sections |> List.distinct |> List.map (fun name -> XmlPatch.SetChildren (name, [ sectionXml name ]))
    // A dimension that does not hold a written cell is made to.
    let dimension =
        match childNamed "dimension" plan.Source.Node with
        | Some node when maxRow >= 0 ->
            match XmlPatch.attribute node "ref" |> Option.bind CellRef.parseRange with
            | Some ((top, left), (bottom, right)) when minRow < top || minCol < left || maxRow > bottom || maxCol > right ->
                [ XmlPatch.EditChild ("dimension", [ XmlPatch.SetAttribute ("ref", cellRef (min top minRow) (min left minCol) + ":" + cellRef (max bottom maxRow) (max right maxCol)) ], None) ]
            | _ -> []
        | _ -> []
    let views =
        if List.contains "sheetViews" patch.Sections then []
        else
            match childrenNamed "sheetViews" plan.Source.Node with
            | views :: _ -> (match tabSelectedEdits active views with [] -> [] | edits -> [ XmlPatch.EditChild ("sheetViews", edits, None) ])
            | [] -> []
    let data = if dataEdits.IsEmpty then [] else [ XmlPatch.EditChild ("sheetData", dataEdits, None) ]
    XmlPatch.applyPart order plan.Source (dimension @ views @ data @ sections), sharedGone

/// A sheet's part as it was read, for a sheet that is written from its cells.
type private SheetSource = {
    /// The part, read in outline (its rows without their cells).
    Part: XmlPatch.SourceElement
    /// The row and column edits since (see `SheetExtras.Moves`).
    Edits: SheetMoves option
}

let private sheetSourceOf (sheet: XlsxSheetOut) : SheetSource option =
    sheet.Extras.SourceXml |> Option.bind readWorksheet |> Option.map (fun part -> { Part = part; Edits = sheet.Extras.Moves })

/// The attributes of a row that the sheet's rows are written with whatever the file says.
let private rowAttributesOfSheet = [ "r"; "spans"; "ht"; "customHeight"; "hidden"; "outlineLevel"; "collapsed" ]

/// What the rows written from the cells keep of the file's rows (see
/// `RowCarry`), from the text of the part the sheet was read from.
/// `fileStyles`: the styles are the file's, so a row's style index means
/// what it meant.
let private rowCarryOf (fileStyles: bool) (moves: SheetMoves option) (sourceXml: string) : RowCarry option =
    let carry =
        { Attributes = System.Collections.Generic.Dictionary<int, string>()
          AutoHeights = System.Collections.Generic.Dictionary<int, float>()
          Heightless = System.Collections.Generic.HashSet<int>()
          Share = false
          SharedStrings = None }
    match readWorksheetCells sourceXml with
    | None -> None
    | Some part ->
        let rows = childrenNamed "sheetData" part.Node |> List.collect (childrenNamed "row")
        let share =
            rows |> List.exists (fun row ->
                childrenNamed "c" row |> List.exists (fun cell -> XmlPatch.child "f" cell |> Option.exists (fun f -> XmlPatch.attribute f "t" = Some "shared")))
        let carry = { carry with Share = share; SharedStrings = Some (sharedStringCells part) }
        match moves with
        | Some moves ->
            // (The same attributes on row after row are worked out once.)
            let known = System.Collections.Generic.Dictionary<string, string * float option * bool>()
            for row in rows do
                match XmlPatch.attribute row "r" |> Option.map System.Int32.TryParse with
                | Some (true, number) ->
                    let own = XmlPatch.attributeBytes part.Text row |> List.filter (fun (name, _) -> name <> "r" && name <> "spans")
                    let rest = own |> List.map snd |> String.concat ""
                    let attributes, auto, heightless =
                        match known.TryGetValue rest with
                        | true, found -> found
                        | false, _ ->
                            let height = XmlPatch.attribute row "ht" |> Option.bind tryParseNumber
                            let custom = isOn (XmlPatch.attribute row "customHeight")
                            let kept =
                                own
                                |> List.filter (fun (name, _) ->
                                    not (List.contains name rowAttributesOfSheet) && (fileStyles || (name <> "s" && name <> "customFormat")))
                                |> List.map snd
                                |> String.concat ""
                            let made = kept, (if custom then None else height), height.IsNone
                            known.[rest] <- made
                            made
                    if attributes <> "" || auto.IsSome || heightless then
                        match movedSpan moves true (number - 1) (number - 1) with
                        | Some (row, _) ->
                            if heightless then carry.Heightless.Add row |> ignore
                            if attributes <> "" then carry.Attributes.[row] <- attributes
                            match auto with
                            | Some height -> carry.AutoHeights.[row] <- height
                            | None -> ()
                        | None -> ()
                | _ -> ()
            Some carry
        | None -> Some carry

/// The rows of a worksheet part written from the sheet.
type private SheetRows =
    /// The file's `<sheetData>` with the changed cells and rows.
    | FileRows of string
    /// Rows written from the cells, with what they keep of the file's rows (None: a sheet of no file).
    | CellRows of RowCarry option

/// The worksheet part of a sheet written from its cells. With the part the
/// sheet was read from (`SheetExtras.SourceXml`), what that says beyond what
/// a sheet holds stays: the worksheet's namespaces, sheetPr, the views'
/// selection, sheetFormatPr, the rows' own attributes, the sheet's
/// protection, ignoredErrors - by where their rows and cells are now.
/// `keepSort`: the file's `<sortState>` still holds.
let private worksheetXml (styles: StyleBook) sharedIndex (dynamicCm: int) sheetIndex active (hasDrawingPart: bool) (rows: SheetRows) (keepSort: bool) (sheet: XlsxSheetOut) =
    let conditionalFormats, conditionalFormatsExt = conditionalFormattingParts styles.Dxfs.Length sheetIndex sheet.ConditionalFormats
    let keptElement (name: string) =
        sheet.Extras.KeptElements |> List.tryPick (fun (known, xml) -> if known = name then Some xml else None) |> Option.defaultValue ""
    let source = sheetSourceOf sheet
    let keptIds = sheet.Extras.KeptRelationships |> List.map (fun (id, _, _, _) -> id) |> Set.ofList
    // An element of the file's part (None for none, and for one that names a relationship the sheet does not keep).
    let fromFile (name: string) : XmlPatch.SourceElement option =
        source
        |> Option.bind (fun source -> sourceChild name source.Part)
        |> Option.filter (fun element -> OpcXml.relationshipIdsIn (XmlPatch.bytes element) |> List.forall (fun id -> Set.contains id keptIds))
    let fileXml (name: string) = fromFile name |> Option.map XmlPatch.bytes |> Option.defaultValue ""
    let moves = source |> Option.bind (fun source -> source.Edits)
    // ... of one that names cells: with them where they are now, or not at all.
    let unmoved (name: string) = match moves with Some [] -> fileXml name | _ -> ""
    let moved (name: string) (child: string) = match moves with Some moves -> withMovedChildren moves child (fromFile name) | None -> ""
    let sheetPr = sheetPrFromSource (fromFile "sheetPr") sheet
    let root =
        match source with
        | Some source ->
            // (The parts of the sheet are named by relationship.)
            let named = (XmlPatch.attribute source.Part.Node "xmlns:r").IsSome
            XmlPatch.startTag source.Part.Text source.Part.Node (if named then [] else [ XmlPatch.SetAttribute ("xmlns:r", Namespaces.relationships) ])
        | None -> sprintf "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">" Namespaces.spreadsheet Namespaces.relationships
    let views =
        match source, moves with
        | Some _, Some moves -> viewsWithTabSelected active (sheetViewsFromSource (fromFile "sheetViews") moves sheet)
        | _ -> sheetViewsXml sheet active
    let data =
        match rows with
        | FileRows data -> data
        | CellRows carry -> "<sheetData>" + rowsXml styles sharedIndex dynamicCm carry sheet + "</sheetData>"
    declaration + "\n"
    + root
    + sheetPr
    + sprintf "<dimension ref=\"%s\"/>" (sheetDimension sheet)
    + views
    + (match source with Some _ -> sheetFormatFromSource (fromFile "sheetFormatPr") sheet | None -> sheetFormatXml sheet)
    + colsXml styles sheet
    + data
    // Child order is fixed by CT_Worksheet; Excel refuses a part that
    // deviates (e.g. tableParts before legacyDrawing).
    + fileXml "sheetCalcPr"
    + fileXml "sheetProtection"
    + moved "protectedRanges" "protectedRange"
    + unmoved "scenarios"
    + autoFilterXml styles.Dxfs.Length sheet
    // (The file's sort: as it is while it holds, or - rows or columns moved, a sheet written from
    // its cells - with its ranges where the cells are now. A sheet sorted anew comes without it.)
    + (if keepSort then unmoved "sortState"
       else (match moves with Some (_ :: _ as edits) -> movedSortState edits (fromFile "sortState") | _ -> ""))
    + unmoved "dataConsolidate"
    + unmoved "customSheetViews"
    + mergeCellsXmlFull sheet.MergedCells
    + fileXml "phoneticPr"
    + conditionalFormats
    + dataValidationsXml sheet.DataValidations
    + hyperlinksXml sheet.Hyperlinks
    + pagePartsXml sheet
    // (With rows or columns moved the file's breaks are not found by their place: only how far they reach is taken.)
    + breaksXml "rowBreaks" (fromFile "rowBreaks") sheet.RowBreaks
    + breaksXml "colBreaks" (fromFile "colBreaks") sheet.ColBreaks
    + keptElement "customProperties"
    + unmoved "cellWatches"
    + moved "ignoredErrors" "ignoredError"
    + unmoved "smartTags"
    + (if hasDrawingPart then sprintf "<drawing r:id=\"rIdDrawing%d\"/>" sheetIndex else "")
    // The file's own VML drawing (it may hold form controls beside the
    // comment boxes) stands for the one made from the comments.
    + (match keptLegacyDrawing sheet with
       | "" -> if sheet.Cells |> List.exists (fun c -> c.Comment.IsSome) then sprintf "<legacyDrawing r:id=\"rIdCommentsVml%d\"/>" sheetIndex else ""
       | kept -> kept)
    + keptElement "legacyDrawingHF"
    + keptElement "drawingHF"
    + keptElement "picture"
    + keptElement "oleObjects"
    + keptElement "controls"
    + keptElement "webPublishItems"
    + tablePartsXml sheet.Tables.Length
    // The slicer list (and other kept extensions) of the file the sheet came from.
    + sheetExtLstXml conditionalFormatsExt (dataValidationsExtXml sheet.DataValidations) (sheet.Extras.KeptCellExtensions @ sheet.Extras.KeptExtensions)
    + "</worksheet>"

/// The XML of a query table part for its table as written: a field per column of the table, in
/// their order, as Excel asks (it refuses the file otherwise); a column the query does not fill
/// is not data bound. None when the part already says that, or is not a part the writer can follow.
let queryTableXmlFromSource (source: string) (tableXmlText: string) : string option =
    match XmlPatch.readPart source, XmlPatch.readPart tableXmlText with
    | Some part, Some tablePart ->
        match XmlPatch.descendants "queryTableRefresh" part.Node |> List.tryHead with
        | Some refresh ->
            match XmlPatch.child "queryTableFields" refresh, XmlPatch.child "tableColumns" tablePart.Node with
            | Some fields, Some columns ->
                let known = XmlPatch.elements fields |> List.filter (fun node -> XmlPatch.localName node = "queryTableField")
                let knownXml = known |> List.map (XmlPatch.bytesOf part.Text)
                let matched =
                    XmlPatch.elements columns
                    |> List.filter (fun node -> XmlPatch.localName node = "tableColumn")
                    |> List.map (fun column ->
                        let id = XmlPatch.attribute column "id" |> Option.defaultValue "0"
                        match known |> List.tryFind (fun field -> XmlPatch.attribute field "tableColumnId" = Some id) with
                        | Some field -> XmlPatch.bytesOf part.Text field, XmlPatch.attribute field "id"
                        | None ->
                            let fieldId = XmlPatch.attribute column "queryTableFieldId" |> Option.defaultValue id
                            sprintf "<queryTableField %s dataBound=\"0\" %s/>" (attr "id" fieldId) (attr "tableColumnId" id), Some fieldId)
                if List.map fst matched = knownXml then None
                else
                    let top =
                        matched |> List.fold (fun most (_, id) -> match id |> Option.map System.Int32.TryParse with Some (true, id) -> max most id | _ -> most) 0
                    let nextId =
                        match XmlPatch.attribute refresh "nextId" |> Option.map System.Int32.TryParse with
                        | Some (true, id) -> max id (top + 1)
                        | _ -> top + 1
                    let edits =
                        [ XmlPatch.SetAttribute ("nextId", string nextId)
                          XmlPatch.EditChild (XmlPatch.name fields,
                                              [ XmlPatch.SetAttribute ("count", string matched.Length)
                                                XmlPatch.SetChildren ((match known with field :: _ -> XmlPatch.name field | [] -> "queryTableField"), List.map fst matched) ], None) ]
                    Some (XmlPatch.applyPart order part (XmlPatch.editsAt part.Node [ refresh, edits ]))
            | _ -> None
        | None -> None
    | _ -> None

/// The XML of a table part of the file for the table as it is now: its
/// name, range, header and totals rows, filter, columns and style are the
/// model's; everything else (the id, revision ids, dxf ids, each kept
/// column's own id, format and calculated-column formula, the filter's
/// criteria, the query-table attributes) stays as the file has it. Columns
/// are matched by position while their number is the file's, by name
/// otherwise. None when the part is not a table the writer can follow.
let tableXmlFromSource (source: string) (table: XlsxTableOut) : string option =
    match XmlPatch.readPart source with
    | Some part when XmlPatch.localName part.Node = "table" ->
      match childNamed "tableColumns" part.Node with
      | None -> None
      | Some columnsNode ->
        let refText = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1) (table.EndCol - 1))
        let root = part.Node
        let isQueryTable = XmlPatch.attribute root "tableType" = Some "queryTable"
        let refChanged = XmlPatch.attribute root "ref" <> Some refText
        let hadTotals = XmlPatch.attribute root "totalsRowCount" |> Option.exists (fun count -> count <> "0")
        let rootEdits =
            [ XmlPatch.SetAttribute ("name", table.Name)
              XmlPatch.SetAttribute ("displayName", table.Name)
              XmlPatch.SetAttribute ("ref", refText)
              attributeEdit "headerRowCount" (if table.HeaderRowCount = 1 then None else Some (string table.HeaderRowCount))
              attributeEdit "totalsRowCount" (if table.TotalsRowCount > 0 then Some (string table.TotalsRowCount) else None)
              if table.TotalsRowCount > 0 then XmlPatch.RemoveAttribute "totalsRowShown"
              elif hadTotals then XmlPatch.SetAttribute ("totalsRowShown", "0") ]
        // The columns.
        let existing = childrenNamed "tableColumn" columnsNode |> List.toArray
        let count = table.EndCol - table.StartCol + 1
        let names = [| for i in 0 .. count - 1 -> table.Columns |> List.tryItem i |> Option.defaultValue (sprintf "Column%d" (i + 1)) |]
        let nameOf (column: XmlDsl.XmlNode) = XmlPatch.attribute column "name" |> Option.defaultValue ""
        let used = Array.create existing.Length false
        let sourceOf (i: int) : XmlDsl.XmlNode option =
            if existing.Length = count then (used.[i] <- true; Some existing.[i])
            else
                let mutable found = -1
                for k in 0 .. existing.Length - 1 do
                    if found < 0 && not used.[k] && System.String.Equals(nameOf existing.[k], names.[i], System.StringComparison.OrdinalIgnoreCase) then found <- k
                if found < 0 then None else (used.[found] <- true; Some existing.[found])
        let matched = [| for i in 0 .. count - 1 -> sourceOf i |]
        let mutable nextId =
            existing |> Array.fold (fun top column -> match XmlPatch.attribute column "id" |> Option.map System.Int32.TryParse with Some (true, id) -> max top id | _ -> top) 0
        let columnXml (i: int) =
            let total = if table.TotalsRowCount > 0 then List.tryItem i table.Totals else None
            let column =
                match matched.[i] with
                | Some column -> Some (XmlPatch.at part column)
                | None ->
                    nextId <- nextId + 1
                    // (A column of a query table names itself: Excel asks for the attribute there.)
                    XmlPatch.read (if isQueryTable then sprintf "<tableColumn id=\"%d\" uniqueName=\"%d\" name=\"\" queryTableFieldId=\"%d\"/>" nextId nextId nextId
                                   else sprintf "<tableColumn id=\"%d\" name=\"\"/>" nextId)
            let edits =
                XmlPatch.SetAttribute ("name", names.[i])
                :: (match total with
                    | None -> []          // (No totals row: what the file says the column would show stays.)
                    | Some (kind, value) ->
                        [ attributeEdit "totalsRowLabel" (if kind = "label" then Some value else None)
                          attributeEdit "totalsRowFunction" (match kind with "function" -> Some value | "custom" -> Some "custom" | _ -> None)
                          // (After the calculated-column formula, before anything else: CT_TableColumn's order.)
                          XmlPatch.SetChildren ("totalsRowFormula", if kind = "custom" then [ "<totalsRowFormula>" + escapeXml value + "</totalsRowFormula>" ] else []) ])
            column |> Option.map (fun column -> XmlPatch.apply order column edits) |> Option.defaultValue ""
        let columns =
            XmlPatch.EditChild ("tableColumns", [ XmlPatch.SetAttribute ("count", string count); XmlPatch.SetChildren ("tableColumn", [ for i in 0 .. count - 1 -> columnXml i ]) ], None)
        // The filter (it covers the header and the data, not the totals row).
        let filterRef = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1 - max 0 table.TotalsRowCount) (table.EndCol - 1))
        let filter =
            match table.AutoFilter, childNamed "autoFilter" root with
            | None, Some _ -> [ XmlPatch.RemoveChild "autoFilter" ]
            | None, None -> []
            | Some _, Some _ -> [ XmlPatch.EditChild ("autoFilter", [ XmlPatch.SetAttribute ("ref", filterRef) ], None) ]
            | Some _, None -> [ XmlPatch.SetChild ("autoFilter", sprintf "<autoFilter ref=\"%s\"/>" filterRef) ]
        // A sort order names cells of the range it was made for.
        let sort = if refChanged then [ XmlPatch.SetChildren ("sortState", []) ] else []
        let style =
            let made =
                table.Style
                |> Option.map (fun s ->
                    sprintf "<tableStyleInfo %s showFirstColumn=\"%d\" showLastColumn=\"%d\" showRowStripes=\"%d\" showColumnStripes=\"%d\"/>"
                        (attr "name" s.Name) (if s.ShowFirstColumn then 1 else 0) (if s.ShowLastColumn then 1 else 0) (if s.ShowRowStripes then 1 else 0) (if s.ShowColumnStripes then 1 else 0))
            match made, childNamed "tableStyleInfo" root with
            | None, Some _ -> [ XmlPatch.SetChildren ("tableStyleInfo", []) ]
            | None, None -> []
            // (The file's element while it says the same.)
            | Some made, Some known when (XmlPatch.read made |> Option.map (fun made -> XmlPatch.attributes made.Node)) = Some (XmlPatch.attributes known) -> []
            | Some made, _ -> [ XmlPatch.SetChildren ("tableStyleInfo", [ made ]) ]
        Some (XmlPatch.applyPart order part (rootEdits @ filter @ sort @ [ columns ] @ style))
    | _ -> None

let private sheetRelsXml sheetIndex (sheet: XlsxSheetOut) (tableParts: (string * int option) list) (placement: SheetPlacement) (drawingPath: string option) =
    let commentsVmlPath = placement.VmlPath
    let rels = [
        for i, h in sheet.Hyperlinks |> List.mapi (fun i h -> i + 1, h) do
            if not h.IsInternal then
                yield sprintf "<Relationship Id=\"rIdHyperlink%d\" Type=\"%s\" Target=\"%s\" TargetMode=\"External\"/>" i RelationshipTypes.hyperlink (escaped h.Target)
        for i, (path, _) in tableParts |> List.mapi (fun i part -> i + 1, part) do
            yield sprintf "<Relationship Id=\"rIdTable%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"/%s\"/>" i path
        if sheet.Cells |> List.exists (fun c -> c.Comment.IsSome) then
            yield sprintf "<Relationship Id=\"rIdComments%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"/%s\"/>" sheetIndex placement.CommentsPath
            if keptLegacyDrawing sheet = "" then
                yield sprintf "<Relationship Id=\"rIdCommentsVml%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"/%s\"/>" sheetIndex commentsVmlPath
        match drawingPath with
        | Some path -> yield sprintf "<Relationship Id=\"rIdDrawing%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"/%s\"/>" sheetIndex path
        | None -> ()
        if not sheet.Extras.ThreadedComments.IsEmpty then
            yield sprintf "<Relationship Id=\"rIdThreaded%d\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"/%s\"/>" sheetIndex placement.ThreadsPath
        // Pivot tables, slicers and timelines of the file, under the ids the
        // sheet's kept extensions name them by.
        for relId, relType, target, isExternal in sheet.Extras.KeptRelationships do
            yield sprintf "<Relationship %s %s %s%s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" target) (if isExternal then " TargetMode=\"External\"" else "")
    ]
    if List.isEmpty rels then None
    else Some (sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" rels))

/// The relationship id of each sheet in workbook.xml.rels: rId<n>, unless a
/// kept relationship of the workbook (a pivot cache, an external link) has it.
// ---- the parts that hold the package together, kept as the file had them ----

/// A part of the package's plumbing as the file had it (`WorkbookPreservation.PackageTexts`).
let private packageText (workbook: XlsxWorkbookOut) (path: string) : string option =
    workbook.Preserved.PackageTexts |> List.tryPick (fun (known, text) -> if known = path then Some text else None)

/// The relationships of a relationships part (none for one that is not well formed): id, type, target (a
/// part path for an internal one; `ownerDir`: the directory of the part they are of), external.
let private relationshipEntries (ownerDir: string) (xml: string) : (string * string * string * bool) list =
    OpcXml.relationshipsOrEmpty xml |> List.map (fun rel -> rel.Id, rel.Type, OpcXml.targetPath ownerDir rel, rel.External)

/// The ids the file's workbook part names its parts by: part path -> id.
let private fileWorkbookRelationshipIds (workbook: XlsxWorkbookOut) : Map<string, string> =
    match packageText workbook "xl/_rels/workbook.xml.rels" with
    | Some text ->
        relationshipEntries "xl" text
        |> List.choose (fun (id, _, target, isExternal) -> if isExternal || id = "" then None else Some (target, id))
        |> List.distinctBy fst
        |> Map.ofList
    | None -> Map.empty

/// A relationships part of the file for what is written now (`made`): it is
/// the file's text while that names the same parts; else the entries that
/// are gone are taken out and the new ones are added behind the others.
/// `sameIds`: an entry is the file's only under the file's id (other parts
/// use the ids); without it a new entry gets an id none of the file's has.
let private relationshipsFromFile (ownerDir: string) (sameIds: bool) (source: string option) (made: string) : string =
    let key (rel: OpcXml.Relationship) =
        (if sameIds then rel.Id else "") + "\n" + rel.Type + "\n" + OpcXml.targetPath ownerDir rel + (if rel.External then "\nexternal" else "")
    OpcXml.mergeRelationships key true source made

/// `[Content_Types].xml` of the file for what is written now, likewise.
/// `unused`: extensions no written part has (a default for one of them is
/// not added to a file that did not have it).
let private contentTypesFromFile (unused: string list) (source: string option) (made: string) : string =
    OpcXml.mergeContentTypes unused source made

let private sheetRelationshipIds (workbook: XlsxWorkbookOut) : string list =
    let fromFile = fileWorkbookRelationshipIds workbook
    let taken =
        (workbook.Preserved.WorkbookRelationships @ Option.toList workbook.Preserved.CalcChain |> List.map (fun (id, _, _, _) -> id))
        @ (fromFile |> Map.toList |> List.map snd)
        |> Set.ofList
    let used = System.Collections.Generic.HashSet<string>()
    workbook.Sheets
    |> List.mapi (fun i sheet ->
        // A sheet of the file keeps the id the file's workbook names it by.
        match sheet.Extras.SourcePart |> Option.bind (fun (path, _) -> Map.tryFind path fromFile) with
        | Some id when used.Add id -> id
        | _ ->
            let plain = sprintf "rId%d" (i + 1)
            if Set.contains plain taken then sprintf "rIdSheet%d" (i + 1) else plain)

/// The `sheetId` of each sheet: the one its file gave it (other parts name the
/// sheet by it: a slicer cache lists pivot tables by tabId), a free one otherwise.
let private sheetIds (workbook: XlsxWorkbookOut) : int list =
    let mutable used = Set.empty<int>
    let fromFile =
        workbook.Sheets
        |> List.map (fun sheet ->
            match sheet.Extras.SheetId with
            | Some id when id > 0 && not (Set.contains id used) ->
                used <- Set.add id used
                Some id
            | _ -> None)
    let mutable next = 1
    fromFile
    |> List.map (fun id ->
        match id with
        | Some id -> id
        | None ->
            while Set.contains next used do next <- next + 1
            used <- Set.add next used
            next)

/// workbook.xml (or an element of it) without the entries that name relationships that are
/// not written any more (`Preserved.DroppedWorkbookRelationships`): a `<pivotCache>`, a slicer
/// cache or timeline cache of the extension list - and without a list, an extension or an
/// extension list that is left empty by that.
let private droppedCacheTargets (ids: string list) (root: XmlDsl.XmlNode) : (XmlDsl.XmlNode * XmlPatch.ElementEdit list) list =
    let lists = [ "slicerCaches"; "timelineCacheRefs"; "ext"; "pivotCaches"; "extLst" ]
    // (Whether the element goes, and the elements in it that go - all of them, it alone where it goes.)
    let rec visit (node: XmlDsl.XmlNode) : bool * XmlDsl.XmlNode list =
        let kids = XmlPatch.elements node |> List.map visit
        let local = XmlPatch.localName node
        let entry =
            List.contains local [ "pivotCache"; "slicerCache"; "timelineCacheRef" ]
            && (XmlPatch.attribute node "r:id" |> Option.exists (fun id -> List.contains id ids))
        let emptied = List.contains local lists && not kids.IsEmpty && kids |> List.forall fst
        if entry || emptied then true, [ node ] else false, kids |> List.collect snd
    if ids.IsEmpty then [] else [ for node in snd (visit root) -> node, [ XmlPatch.Removed ] ]

/// workbook.xml's entries that name relationships that are not written any more
/// (`Preserved.DroppedWorkbookRelationships`) - a `<pivotCache>`, a slicer cache or timeline cache of
/// the extension list - taken out of XML of it (elements one after another): with a list, an extension
/// or an extension list that is left empty by that.
let private withoutDroppedCaches (ids: string list) (xml: string) : string =
    if ids.IsEmpty || xml = "" then xml
    else
        match XmlPatch.readFragmentRoot xml with
        | Some fragment ->
            match droppedCacheTargets ids fragment.Node with
            | [] -> xml
            | targets -> XmlPatch.applyFragment order fragment (XmlPatch.editsAt fragment.Node targets)
        | None -> xml

let private workbookXmlFull (workbook: XlsxWorkbookOut) =
    let relIds = sheetRelationshipIds workbook |> List.toArray
    let ids = sheetIds workbook |> List.toArray
    let sheets =
        workbook.Sheets
        |> List.mapi (fun i s ->
            let visible =
                match s.State with
                | XlsxReader.SheetVisible -> ""
                | XlsxReader.SheetHidden -> " state=\"hidden\""
                | XlsxReader.SheetVeryHidden -> " state=\"veryHidden\""
            sprintf "<sheet %s sheetId=\"%d\" r:id=\"%s\"%s/>"
                (attr "name" s.Name) ids.[i] relIds.[i] visible)
        |> String.concat ""
    // In a reference a sheet name is quoted with apostrophes, and an
    // apostrophe inside the name is doubled ('Bob''s'!$A$1).
    let quoted (name: string) = "'" + name.Replace("'", "''") + "'"
    let printNames =
        workbook.Sheets
        |> List.mapi (fun i s ->
            [ match s.PrintArea with
              | Some v ->
                  // (Each area of a print area of several is named with its sheet.)
                  let areas = CellRef.areas v |> List.map (fun area -> sprintf "%s!%s" (quoted s.Name) area) |> String.concat ","
                  yield { Name = "_xlnm.Print_Area"; RefersTo = areas; LocalSheetIndex = Some i; Comment = None; Hidden = false }
              | None -> ()
              match s.PrintTitleRows, s.PrintTitleCols with
              | None, None -> ()
              | rows, cols ->
                  let parts = [ rows; cols ] |> List.choose id |> List.map (fun part -> sprintf "%s!%s" (quoted s.Name) part)
                  yield { Name = "_xlnm.Print_Titles"; RefersTo = String.concat "," parts; LocalSheetIndex = Some i; Comment = None; Hidden = false }
              // Excel keeps the range of a sheet's AutoFilter in a hidden name of the sheet.
              match s.Extras.AutoFilter with
              | Some filter ->
                  let reference = CellRef.corners filter.Ref |> List.map CellRef.absolute |> CellRef.ofCorners
                  let fromFile =
                      workbook.DefinedNames
                      |> List.tryFind (fun name -> name.Name = "_xlnm._FilterDatabase" && name.LocalSheetIndex = Some i)
                  let generated = { Name = "_xlnm._FilterDatabase"; RefersTo = ""; LocalSheetIndex = Some i; Comment = None; Hidden = true }
                  yield { defaultArg fromFile generated with RefersTo = sprintf "%s!%s" (quoted s.Name) reference }
              | None -> () ])
        |> List.concat
    let names =
        (workbook.DefinedNames
         |> List.filter (fun name ->
             not (name.Name = "_xlnm._FilterDatabase"
                  && (name.LocalSheetIndex
                      |> Option.exists (fun i -> i >= 0 && i < workbook.Sheets.Length && workbook.Sheets.[i].Extras.AutoFilter.IsSome)))))
        @ printNames
    let definedNames =
        if List.isEmpty names then ""
        else
            let body =
                names
                |> List.map (fun n ->
                    let local = n.LocalSheetIndex |> Option.map (fun i -> sprintf " localSheetId=\"%d\"" i) |> Option.defaultValue ""
                    let comment = n.Comment |> Option.map (fun text -> " " + attr "comment" text) |> Option.defaultValue ""
                    let hidden = if n.Hidden then " hidden=\"1\"" else ""
                    sprintf "<definedName %s%s%s%s>%s</definedName>"
                        (attr "name" n.Name) comment local hidden
                        // (The text of an element: quotes and apostrophes stand as they are.)
                        ((escaped (fileFormula n.RefersTo)).Replace("&apos;", "'").Replace("&quot;", "\"")))
                |> String.concat ""
            "<definedNames>" + body + "</definedNames>"
    // A formula written without a result has to be calculated when the file is opened.
    let hasUncachedFormula =
        workbook.Sheets
        |> List.exists (fun s ->
            // (The cells of a sheet of the file that were not made yet: the reader has looked.)
            match (if s.CellSource.IsLoaded then None else s.Extras.UncachedFormulas) with
            | Some known -> known
            | None -> s.Cells |> List.exists (fun c -> match c.Value with XlsxFormula (_, None) -> true | _ -> false))
    // Iterative calculation of circular references is a setting of the workbook: without it Excel
    // opens a workbook that has them with a warning and zeros.
    let calcPr =
        let iteration =
            match workbook.Preserved.CalcIteration with
            | Some (count, delta) ->
                " iterate=\"1\"" + (if count <> 100 then sprintf " iterateCount=\"%d\"" count else "")
                + (if delta <> 0.001 then " iterateDelta=\"" + numberText delta + "\"" else "")
            | None -> ""
        if hasUncachedFormula || iteration <> "" then "<calcPr" + iteration + (if hasUncachedFormula then " fullCalcOnLoad=\"1\"" else "") + "/>" else ""
    // Child order is fixed by CT_Workbook: sheets, externalReferences,
    // definedNames, calcPr, pivotCaches, extLst.
    let kept = workbook.Preserved
    // The file's own workbook part with the sheet list, the names, the active
    // tab and the calculation settings made anew: its file version, workbook
    // properties and protection, window, views and extensions stay.
    let fromFile =
        kept.WorkbookPart
        |> Option.bind (fun path -> kept.Parts |> List.tryPick (fun (known, data) -> if known = path then Some (System.Text.Encoding.UTF8.GetString data) else None))
        |> Option.bind XmlPatch.readPart
        |> Option.bind (fun part ->
            // (The prefix the part names relationships by.)
            let prefix =
                XmlPatch.attributes part.Node |> Map.toList |> List.tryPick (fun (name, value) ->
                    let local = XmlDsl.Ns.stripNs name
                    if name = "xmlns:" + local && value = Namespaces.relationships then Some local else None)
            match prefix, childNamed "sheets" part.Node with
            | Some prefix, Some sheetsNode when XmlPatch.localName part.Node = "workbook" ->
                let sheetsXml =
                    workbook.Sheets
                    |> List.mapi (fun i s ->
                        let visible =
                            match s.State with
                            | XlsxReader.SheetVisible -> ""
                            | XlsxReader.SheetHidden -> " state=\"hidden\""
                            | XlsxReader.SheetVeryHidden -> " state=\"veryHidden\""
                        sprintf "<sheet %s sheetId=\"%d\" %s:id=\"%s\"%s/>" (attr "name" s.Name) ids.[i] prefix relIds.[i] visible)
                // (An element that says the same in the file stays as the file has it: the same entries, by
                // their attributes and what they hold - in whatever order, the names of a file being sorted by
                // Excel and the print names coming last here.)
                let entries (child: string) (node: XmlDsl.XmlNode option) =
                    match node with
                    | Some node ->
                        XmlPatch.elements node
                        |> List.filter (fun entry -> XmlPatch.name entry = child)
                        |> List.map (fun entry ->
                            (XmlPatch.attributes entry |> Map.toList |> List.map (fun (name, value) -> name + "=" + value) |> List.sort),
                            // (A sheet name in a reference may be quoted or not.)
                            (XmlPatch.textOf entry).Replace("'", ""))
                    | None -> []
                let saysTheSame (child: string) (was: XmlDsl.XmlNode option) (made: string) =
                    List.sort (entries child was) = List.sort (entries child (XmlPatch.read made |> Option.map (fun made -> made.Node)))
                let madeSheets = "<sheets>" + String.concat "" sheetsXml + "</sheets>"
                let sheets =
                    let names (node: XmlDsl.XmlNode) = XmlPatch.elements node |> List.choose (fun entry -> XmlPatch.attribute entry "name")
                    if saysTheSame "sheet" (Some sheetsNode) madeSheets
                       && Some (names sheetsNode) = (XmlPatch.read madeSheets |> Option.map (fun made -> names made.Node)) then []
                    else [ XmlPatch.EditChild ("sheets", [ XmlPatch.SetChildren ("sheet", sheetsXml) ], None) ]
                // The sheet on display.
                let active = string (max 0 workbook.ActiveSheetIndex)
                let views =
                    match XmlPatch.descendants "workbookView" part.Node |> List.tryHead, childNamed "bookViews" part.Node with
                    | Some view, Some _ ->
                        // (The view is the file's while it shows the same sheet.)
                        if (XmlPatch.attribute view "activeTab" |> Option.defaultValue "0") = active then []
                        else XmlPatch.editsAt part.Node [ view, [ XmlPatch.SetAttribute ("activeTab", active); XmlPatch.RemoveAttribute "firstSheet" ] ]
                    | _ when active = "0" -> []
                    | _ -> [ XmlPatch.SetChild ("bookViews", "<bookViews><workbookView activeTab=\"" + active + "\"/></bookViews>") ]
                // The defined names.
                let names =
                    match childNamed "definedNames" part.Node with
                    | Some known ->
                        if saysTheSame "definedName" (Some known) (if definedNames = "" then "<definedNames/>" else definedNames) then []
                        elif definedNames = "" then [ XmlPatch.RemoveChild "definedNames" ]
                        else [ XmlPatch.SetChild ("definedNames", definedNames) ]
                    | None -> if definedNames = "" then [] else [ XmlPatch.SetChild ("definedNames", definedNames) ]
                // Calculation: the file's settings, with what the workbook says of iteration and of formulas without a result.
                let calculation =
                    match childNamed "calcPr" part.Node with
                    | Some _ ->
                        let iterate =
                            match kept.CalcIteration with
                            | Some (count, delta) ->
                                [ XmlPatch.SetAttribute ("iterate", "1")
                                  attributeEdit "iterateCount" (if count <> 100 then Some (string count) else None)
                                  attributeEdit "iterateDelta" (if delta <> 0.001 then Some (numberText delta) else None) ]
                            | None -> [ XmlPatch.RemoveAttribute "iterate"; XmlPatch.RemoveAttribute "iterateCount"; XmlPatch.RemoveAttribute "iterateDelta" ]
                        [ XmlPatch.EditChild ("calcPr", iterate @ (if hasUncachedFormula then [ XmlPatch.SetAttribute ("fullCalcOnLoad", "1") ] else []), None) ]
                    | None -> if calcPr = "" then [] else [ XmlPatch.SetChild ("calcPr", calcPr) ]
                let caches = XmlPatch.editsAt part.Node (droppedCacheTargets kept.DroppedWorkbookRelationships part.Node)
                Some (XmlPatch.applyPart order part (sheets @ views @ names @ calculation @ caches))
            | _ -> None)
    match fromFile with
    | Some text -> text
    | None ->
    sprintf "%s\n<workbook xmlns=\"%s\" xmlns:r=\"%s\">%s<bookViews><workbookView activeTab=\"%d\"/></bookViews><sheets>%s</sheets>%s%s%s%s%s</workbook>"
        declaration Namespaces.spreadsheet Namespaces.relationships (if kept.Date1904 then "<workbookPr date1904=\"1\"/>" else "") (max 0 workbook.ActiveSheetIndex) sheets
        kept.ExternalReferencesXml definedNames calcPr
        (withoutDroppedCaches kept.DroppedWorkbookRelationships kept.PivotCachesXml) (withoutDroppedCaches kept.DroppedWorkbookRelationships kept.WorkbookExtLstXml)

let private workbookRelsXmlFull (workbook: XlsxWorkbookOut) (placements: SheetPlacement list) hasVba (themePath: string option) (hasPersons: bool) (ownMetadata: bool) =
    let sheetRels =
        List.zip (sheetRelationshipIds workbook) placements
        |> List.map (fun (id, placement) ->
            let relType = match placement.Verbatim with Some (_, relType) when relType <> "" -> relType | _ -> RelationshipTypes.worksheet
            sprintf "<Relationship Id=\"%s\" %s Target=\"/%s\"/>" id (attr "Type" relType) (escaped placement.SheetPath))
    // (A part the file's workbook already names keeps the id it has there.)
    let fromFile = fileWorkbookRelationshipIds workbook
    let idOf (path: string) (otherwise: string) = Map.tryFind path fromFile |> Option.defaultValue otherwise
    let baseRels = [
        yield! sheetRels
        yield sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"styles.xml\"/>" (idOf "xl/styles.xml" "rIdStyles") RelationshipTypes.styles
        yield sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"sharedStrings.xml\"/>" (idOf "xl/sharedStrings.xml" "rIdSharedStrings") RelationshipTypes.sharedStrings
        if hasVba then yield sprintf "<Relationship Id=\"%s\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>" (idOf "xl/vbaProject.bin" "rIdVBA")
        // The theme of the opened workbook, kept as it was.
        match themePath with
        | Some path -> yield sprintf "<Relationship Id=\"%s\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"/%s\"/>" (idOf path "rIdTheme") (escaped path)
        | None -> ()
        // The cell metadata dynamic array formulas name, when the opened workbook had none.
        if ownMetadata then yield "<Relationship Id=\"rIdMetadata\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>"
        if hasPersons then yield sprintf "<Relationship Id=\"%s\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"/xl/persons/person.xml\"/>" (idOf "xl/persons/person.xml" "rIdPersons")
        // Pivot caches, slicer caches and external links, under the ids workbook.xml names them by.
        for relId, relType, target, isExternal in workbook.Preserved.WorkbookRelationships @ Option.toList workbook.Preserved.CalcChain do
            yield sprintf "<Relationship %s %s %s%s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" target) (if isExternal then " TargetMode=\"External\"" else "")
    ]
    sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" baseRels)
    |> relationshipsFromFile "xl" true (packageText workbook "xl/_rels/workbook.xml.rels")

let private contentTypesXmlFull (workbook: XlsxWorkbookOut) (placements: SheetPlacement list) (keptParts: (string * byte[]) list) (ownMetadata: bool) (noteParts: string list) (noteBoxes: bool) =
    let hasVba = workbook.VbaProjectSource |> Option.exists hasText
    let workbookType =
        if hasVba then "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    let worksheetType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
    let sheetOverrides =
        placements
        |> List.map (fun placement ->
            // A sheet written as it was read is what the file said it is.
            let contentType =
                match placement.Verbatim with
                | Some (path, relType) ->
                    workbook.Preserved.Overrides
                    |> List.tryPick (fun (name, contentType) -> if name.TrimStart('/').Equals(path, System.StringComparison.OrdinalIgnoreCase) then Some contentType else None)
                    |> Option.defaultValue (
                        if OpcXml.relationshipKind relType = "chartsheet" then "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
                        elif OpcXml.relationshipKind relType = "dialogsheet" then "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"
                        else worksheetType)
                | None -> worksheetType
            sprintf "<Override %s %s/>" (attr "PartName" ("/" + placement.SheetPath)) (attr "ContentType" contentType))
        |> String.concat ""
    let mutable tableId = 1
    let tableOverrides =
        tablePlacement workbook
        |> List.collect (List.choose (fun (path, id) ->
            id |> Option.map (fun _ -> sprintf "<Override PartName=\"/%s\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>" path)))
        |> String.concat ""
    let written = List.zip workbook.Sheets placements |> List.filter (fun (_, placement) -> placement.Verbatim.IsNone)
    let comments =
        written
        |> List.map (fun (s, placement) ->
            // (A comments part of the file is typed with the kept parts.)
            if (keptCommentsPart workbook s).IsNone && s.Cells |> List.exists (fun c -> c.Comment.IsSome) then
                sprintf "<Override PartName=\"/%s\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>" placement.CommentsPath
            else "")
        |> String.concat ""
    // Pictures, the kept parts (charts, theme, embedded data) and threaded comments.
    let verbatimPaths = placements |> List.choose (fun placement -> placement.Verbatim |> Option.map (fun (path, _) -> ("/" + path).ToLowerInvariant())) |> Set.ofList
    let keptPaths = keptParts |> List.map (fun (path, _) -> ("/" + path).ToLowerInvariant()) |> Set.ofList
    let ownDefaults =
        [ yield "rels", "application/vnd.openxmlformats-package.relationships+xml"
          yield "xml", "application/xml"
          yield "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing"
          if hasVba then yield "bin", "application/vnd.ms-office.vbaProject" ]
    let imageDefaults =
        workbook.Sheets
        |> List.collect (fun sheet -> sheet.Extras.Images |> List.map (fun image -> image.Extension, image.ContentType))
    let defaults =
        ownDefaults @ imageDefaults @ (if keptPaths.IsEmpty then [] else workbook.Preserved.Defaults)
        |> List.filter (fun (extension, _) -> extension <> "")
        |> List.distinctBy (fun (extension, _) -> extension.ToLowerInvariant())
        |> List.map (fun (extension, contentType) -> sprintf "<Default %s %s/>" (attr "Extension" extension) (attr "ContentType" contentType))
        |> String.concat ""
    let extraOverrides =
        [ for sheet, placement in written do
              if hasDrawing sheet && (keptDrawingPart workbook sheet).IsNone then
                  yield sprintf "<Override PartName=\"/%s\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>" placement.DrawingPath
              if not sheet.Extras.ThreadedComments.IsEmpty then
                  yield sprintf "<Override PartName=\"/%s\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"/>" placement.ThreadsPath
          if ownMetadata then
              yield "<Override PartName=\"/xl/metadata.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml\"/>"
          if workbook.Sheets |> List.exists (fun sheet -> not sheet.Extras.ThreadedComments.IsEmpty) then
              yield "<Override PartName=\"/xl/persons/person.xml\" ContentType=\"application/vnd.ms-excel.person+xml\"/>"
          for name, contentType in workbook.Preserved.Overrides do
              if Set.contains (name.ToLowerInvariant()) keptPaths && not (Set.contains (name.ToLowerInvariant()) verbatimPaths) then
                  yield sprintf "<Override %s %s/>" (attr "PartName" name) (attr "ContentType" contentType) ]
        |> String.concat ""
    declaration + "\n"
    + sprintf "<Types xmlns=\"%s\">" Namespaces.contentTypes
    + defaults
    + sprintf "<Override PartName=\"/xl/workbook.xml\" ContentType=\"%s\"/>" workbookType
    + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
    + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
    + sheetOverrides + tableOverrides + comments
    // (The comments parts added to sheets that are otherwise the file's parts.)
    + (noteParts |> List.map (fun path -> sprintf "<Override PartName=\"/%s\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>" path) |> String.concat "")
    + extraOverrides
    + "</Types>"
    |> contentTypesFromFile
        // (The default for VML drawings is for the boxes of comments and the kept drawings of that kind.)
        (if noteBoxes || (keptParts |> List.exists (fun (path, _) -> PartPath.extension path = "vml")) || (written |> List.exists (fun (sheet, _) -> sheet.Cells |> List.exists (fun c -> c.Comment.IsSome))) then [] else [ "vml" ])
        (packageText workbook "[Content_Types].xml")

/// The package relationships: the workbook, and the document properties kept from the opened file.
let private rootRelsXmlFull (workbook: XlsxWorkbookOut) =
    let kept = workbook.Preserved.RootRelationships
    let mainId = if kept |> List.exists (fun (id, _, _, _) -> id = "rId1") then "rIdWorkbook" else "rId1"
    let keptXml =
        kept
        |> List.map (fun (relId, relType, target, _) -> sprintf "<Relationship %s %s %s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" (target.TrimStart '/')))
        |> String.concat ""
    sprintf "%s\n<Relationships xmlns=\"%s\"><Relationship Id=\"%s\" Type=\"%s\" Target=\"xl/workbook.xml\"/>%s</Relationships>"
        declaration Namespaces.relationshipsPackage mainId RelationshipTypes.officeDocument keptXml
    |> relationshipsFromFile "" false (packageText workbook "_rels/.rels")

/// The kept parts that something written still refers to: those the written
/// relationships name, and what those parts' own relationships lead to. A
/// part of a sheet that was removed (its pivot table, its printer settings,
/// the chart of its drawing) is left out rather than written as an orphan.
let private reachableKeptParts (workbook: XlsxWorkbookOut) (themePath: string option) : (string * byte[]) list =
    let parts = System.Collections.Generic.Dictionary<string, byte[]>()
    for path, data in workbook.Preserved.Parts do parts.[path] <- data
    let seen = System.Collections.Generic.HashSet<string>()
    let pending = System.Collections.Generic.Queue<string>()
    let want (target: string) (isExternal: bool) =
        if not isExternal && not (OpcXml.isAnchorTarget target) then pending.Enqueue (ZipInterop.normalizePath target)
    for sheet in workbook.Sheets do
        match verbatimPart workbook sheet with
        | Some (path, _) -> want path false
        | None ->
            match keptCommentsPart workbook sheet with
            | Some path when sheet.Cells |> List.exists (fun c -> c.Comment.IsSome) -> want path false
            | _ -> ()
        for _, _, target, isExternal in sheet.Extras.KeptRelationships do want target isExternal
        match keptDrawingPart workbook sheet with
        | Some path -> want path false
        | None ->
            if hasDrawing sheet then
                for _, _, target, isExternal in sheet.Extras.DrawingRelationships do want target isExternal
        for _, path in sheet.Extras.TableParts do
            if sheet.Tables |> List.exists (fun table -> sheet.Extras.TableParts |> List.exists (fun (name, known) -> name = table.Name && known = path)) then want path false
        // (A changed table of the file is written as its part: what the part leads to - a query table - stays.)
        for table in sheet.Tables do
            match table.SourcePart with
            | Some path -> want path false
            | None -> ()
    for _, _, target, isExternal in workbook.Preserved.WorkbookRelationships @ Option.toList workbook.Preserved.CalcChain @ workbook.Preserved.RootRelationships do
        want target isExternal
    match themePath with Some path -> pending.Enqueue path | None -> ()
    while pending.Count > 0 do
        let path = pending.Dequeue()
        if parts.ContainsKey path && seen.Add path then
            let relsPath = PartPath.relationshipsOf path
            match parts.TryGetValue relsPath with
            | true, relsData ->
                seen.Add relsPath |> ignore
                for rel in OpcXml.relationshipsOrEmpty (System.Text.Encoding.UTF8.GetString relsData) do
                    if not rel.External && rel.Target <> "" then pending.Enqueue (OpcXml.targetPath (PartPath.directory path) rel)
            | _ -> ()
    workbook.Preserved.Parts |> List.filter (fun (path, _) -> seen.Contains path)

// ---------------------------------------------------------------------------
// Cell metadata of dynamic array formulas (xl/metadata.xml)
// ---------------------------------------------------------------------------

let private xldaprTypeXml =
    "<metadataType name=\"XLDAPR\" minSupportedVersion=\"120000\" copy=\"1\" pasteAll=\"1\" pasteValues=\"1\" merge=\"1\" splitFirst=\"1\" rowColShift=\"1\" clearFormats=\"1\" clearComments=\"1\" assign=\"1\" coerce=\"1\" cellMeta=\"1\"/>"
let private xdaNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
let private dynamicArrayBlockXml (declareNamespace: bool) =
    "<bk><extLst><ext uri=\"{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}\"><xda:dynamicArrayProperties"
    + (if declareNamespace then " xmlns:xda=\"" + xdaNamespace + "\"" else "")
    + " fDynamic=\"1\" fCollapsed=\"0\"/></ext></extLst></bk>"

/// The metadata part of a workbook whose only metadata is "this formula is a
/// dynamic array", as Excel writes it: cells name its one record with `cm="1"`.
let private dynamicArrayMetadataXml =
    declaration + "\n<metadata xmlns=\"" + Namespaces.spreadsheet + "\" xmlns:xda=\"" + xdaNamespace + "\">"
    + "<metadataTypes count=\"1\">" + xldaprTypeXml + "</metadataTypes>"
    + "<futureMetadata name=\"XLDAPR\" count=\"1\">" + dynamicArrayBlockXml false + "</futureMetadata>"
    + "<cellMetadata count=\"1\"><bk><rc t=\"1\" v=\"0\"/></bk></cellMetadata></metadata>"

/// A metadata part with a cell-metadata record for dynamic array formulas, and
/// the `cm` that names it. A part that has one is returned as it is; one that
/// has other metadata (rich values, MDX) gets the XLDAPR type, its future
/// metadata block and the record added, everything else staying as it was.
let metadataWithDynamicArrays (existing: string) : string * int =
    let known = try XlsxReader.dynamicArrayCellMetadata existing with _ -> Set.empty
    if not known.IsEmpty then existing, Set.minElement known
    else
        match XmlPatch.readPart existing with
        | Some part when XmlPatch.localName part.Node = "metadata" ->
            let root = part.Node
            // (An element that holds nothing is written again in full.)
            let holding (name: string) = childNamed name root |> Option.filter (fun node -> not (isSelfClosed node))
            let count (name: string) (node: XmlDsl.XmlNode) = XmlPatch.elements node |> List.filter (fun kid -> XmlPatch.localName kid = name) |> List.length
            // 1. The metadata type.
            let typeEdits, typeIndex =
                match holding "metadataTypes" with
                | Some types ->
                    let all = XmlPatch.elements types |> List.filter (fun kid -> XmlPatch.localName kid = "metadataType")
                    match all |> List.tryFindIndex (fun kind -> XmlPatch.attribute kind "name" = Some "XLDAPR") with
                    | Some index -> [], index + 1
                    | None -> [ XmlPatch.EditChild ("metadataTypes", [ XmlPatch.SetAttribute ("count", string (all.Length + 1)); XmlPatch.Append xldaprTypeXml ], None) ], all.Length + 1
                | None -> [ XmlPatch.SetChild ("metadataTypes", "<metadataTypes count=\"1\">" + xldaprTypeXml + "</metadataTypes>") ], 1
            // 2. The future metadata block that says "dynamic".
            let futureEdits, blockIndex =
                match childrenNamed "futureMetadata" root |> List.tryFind (fun node -> XmlPatch.attribute node "name" = Some "XLDAPR" && not (isSelfClosed node)) with
                | Some future ->
                    let had = count "bk" future
                    [ XmlPatch.EditChildWhere ("futureMetadata", "name", "XLDAPR", [ XmlPatch.SetAttribute ("count", string (had + 1)); XmlPatch.Append (dynamicArrayBlockXml true) ]) ], had
                | None -> [ XmlPatch.AddChild ("futureMetadata", "<futureMetadata name=\"XLDAPR\" count=\"1\">" + dynamicArrayBlockXml true + "</futureMetadata>") ], 0
            // 3. The cell-metadata record that points at it.
            let record = sprintf "<bk><rc t=\"%d\" v=\"%d\"/></bk>" typeIndex blockIndex
            let cellEdits, cm =
                match holding "cellMetadata" with
                | Some cells ->
                    let had = count "bk" cells
                    [ XmlPatch.EditChild ("cellMetadata", [ XmlPatch.SetAttribute ("count", string (had + 1)); XmlPatch.Append record ], None) ], had + 1
                | None -> [ XmlPatch.SetChild ("cellMetadata", "<cellMetadata count=\"1\">" + record + "</cellMetadata>") ], 1
            XmlPatch.applyPart order part (typeEdits @ futureEdits @ cellEdits), cm
        | _ -> dynamicArrayMetadataXml, 1

// ---------------------------------------------------------------------------
// Styles and shared strings extended in place (WorkbookPreservation.InPlace)
// ---------------------------------------------------------------------------

/// The file's styles part as a style book that only adds: a format the file
/// has (the same font, fill, border, number format, alignment and named
/// style) is its cellXfs entry of the file, any other becomes a new entry
/// behind the file's - with its font, fill, border and number format added
/// behind the file's own. So every style index the file's sheets use means
/// what it meant. The second value gives the part when all cells have been
/// written. None when the part is not one this can be done to (no cellXfs,
/// fonts, fills or borders; entries the reader did not count the same way).
let private inPlaceStyles (workbook: XlsxWorkbookOut) (stylesXml: string) : (StyleBook * (unit -> string)) option =
    let kept = workbook.Preserved
    match XmlPatch.readPart stylesXml |> Option.filter (fun part -> XmlPatch.localName part.Node = "styleSheet") with
    | None -> None
    | Some part ->
    // A list of the style sheet (it may close itself), and its items.
    let list (name: string) = XmlPatch.elements part.Node |> List.tryFind (fun node -> XmlPatch.localName node = name)
    let itemsOf (name: string) (item: string) =
        list name |> Option.map (fun node -> XmlPatch.elements node |> List.filter (fun kid -> XmlPatch.localName kid = item))
    match itemsOf "fonts" "font", itemsOf "fills" "fill", itemsOf "borders" "border", itemsOf "cellXfs" "xf" with
    | Some fonts, Some fills, Some borders, Some sourceXfs ->
        let fontCount, fillCount, borderCount, xfCount = fonts.Length, fills.Length, borders.Length, sourceXfs.Length
        let dxfCount = itemsOf "dxfs" "dxf" |> Option.map List.length |> Option.defaultValue 0
        if xfCount = 0 || xfCount <> kept.CellFormats.Length || dxfCount <> kept.DxfCount || workbook.Dxfs.Length < dxfCount || fontCount = 0 || fillCount < 2 || borderCount = 0 then None
        else
            // The xfId of each named style, as the reader numbers them.
            let styleXfCount = itemsOf "cellStyleXfs" "xf" |> Option.map List.length |> Option.defaultValue 0
            let namedXfIds =
                itemsOf "cellStyles" "cellStyle" |> Option.defaultValue []
                |> List.choose (fun style ->
                    match XmlPatch.attribute style "name", XmlPatch.attribute style "xfId" |> Option.map System.Int32.TryParse with
                    | Some _, Some (true, xfId) when xfId >= 0 && xfId < styleXfCount -> Some xfId
                    | _ -> None)
                |> List.distinct |> List.sort |> List.toArray
            // The custom number formats of the file.
            let numFmtIds = System.Collections.Generic.Dictionary<string, int>()
            let mutable nextNumFmt = 164
            for numFmt in itemsOf "numFmts" "numFmt" |> Option.defaultValue [] do
                match XmlPatch.attribute numFmt "numFmtId" |> Option.map System.Int32.TryParse, XmlPatch.attribute numFmt "formatCode" with
                | Some (true, id), Some code ->
                    nextNumFmt <- max nextNumFmt (id + 1)
                    if not (numFmtIds.ContainsKey code) then numFmtIds.[code] <- id
                | _ -> ()
            // The file's formats by what the writer knows of a format.
            let xfIndex = System.Collections.Generic.Dictionary<XfKey, int>(HashIdentity.Structural)
            kept.CellFormats
            |> Array.iteri (fun index style ->
                let named = if index < kept.CellXfNamedStyles.Length then max 0 kept.CellXfNamedStyles.[index] else 0
                let key = { columnStyleKey style with NamedStyle = named }
                // (Not entry 0: a cell is written without a style for it, and a blank
                // one not at all - only the writer's own default format is that.)
                if index > 0 && not (xfIndex.ContainsKey key) then xfIndex.[key] <- index)
            xfIndex.[defaultXfKey] <- 0
            let newNumFmts, newFonts, newFills, newBorders, newXfs = ResizeArray<string>(), ResizeArray<string>(), ResizeArray<string>(), ResizeArray<string>(), ResizeArray<string>()
            let fontIds = System.Collections.Generic.Dictionary<FontKey, int>(HashIdentity.Structural)
            let fillIds = System.Collections.Generic.Dictionary<string, int>()
            let borderIds = System.Collections.Generic.Dictionary<BorderKey, int>(HashIdentity.Structural)
            let numFmtIdOf (key: XfKey) =
                match key.NumFmt with
                | None -> 0
                | Some code ->
                    match builtInNumFmtId code with
                    | Some id -> id
                    | None ->
                        match numFmtIds.TryGetValue code with
                        | true, id -> id
                        | false, _ ->
                            let id = nextNumFmt
                            nextNumFmt <- nextNumFmt + 1
                            numFmtIds.[code] <- id
                            newNumFmts.Add (sprintf "<numFmt numFmtId=\"%d\" formatCode=\"%s\"/>" id (escaped code))
                            id
            let fontIdOf (key: XfKey) =
                if key.Font = defaultFontKey then 0
                else
                    match fontIds.TryGetValue key.Font with
                    | true, id -> id
                    | false, _ ->
                        let id = fontCount + newFonts.Count
                        fontIds.[key.Font] <- id
                        newFonts.Add (fontXml key.Font)
                        id
            let fillIdOf (key: XfKey) =
                match key.Fill with
                | None -> 0
                | Some color ->
                    match fillIds.TryGetValue color with
                    | true, id -> id
                    | false, _ ->
                        let id = fillCount + newFills.Count
                        fillIds.[color] <- id
                        newFills.Add (fillXml key.Fill)
                        id
            let borderIdOf (key: XfKey) =
                if key.Border = defaultBorderKey then 0
                else
                    match borderIds.TryGetValue key.Border with
                    | true, id -> id
                    | false, _ ->
                        let id = borderCount + newBorders.Count
                        borderIds.[key.Border] <- id
                        newBorders.Add (borderXml key.Border)
                        id
            let indexOfKey (key: XfKey) =
                match xfIndex.TryGetValue key with
                | true, index -> index
                | false, _ ->
                    let numFmtId = numFmtIdOf key
                    let fontId = fontIdOf key
                    let fillId = fillIdOf key
                    let borderId = borderIdOf key
                    let xfId = if key.NamedStyle > 0 && key.NamedStyle < namedXfIds.Length then namedXfIds.[key.NamedStyle] else 0
                    let index = xfCount + newXfs.Count
                    newXfs.Add (xfXmlOf numFmtId fontId fillId borderId xfId key)
                    xfIndex.[key] <- index
                    index
            // A format the user made from one of the file's (BasedOnXf): the
            // file's entry with what the change touched replaced - its number
            // format, font, fill, border or alignment, each only when it is
            // not what the reader saw in that entry. Everything else of the
            // entry stays: its named style, protection, quote prefix, pivot
            // button, attributes and extensions this writer knows nothing of.
            let sourceXfs = List.toArray sourceXfs
            let derived = System.Collections.Generic.Dictionary<struct (int * XfKey), int>(HashIdentity.Structural)
            let derivedIndex (basedOn: int) (key: XfKey) : int option =
                if basedOn < 0 || basedOn >= xfCount then None
                else
                    match derived.TryGetValue (struct (basedOn, key)) with
                    | true, index -> Some index
                    | false, _ ->
                        let was = columnStyleKey kept.CellFormats.[basedOn]
                        let source = sourceXfs.[basedOn]
                        let ids =
                            [ if key.NumFmt <> was.NumFmt then yield! [ XmlPatch.SetAttribute ("numFmtId", string (numFmtIdOf key)); XmlPatch.SetAttribute ("applyNumberFormat", "1") ]
                              if key.Font <> was.Font then yield! [ XmlPatch.SetAttribute ("fontId", string (fontIdOf key)); XmlPatch.SetAttribute ("applyFont", "1") ]
                              if key.Fill <> was.Fill then yield! [ XmlPatch.SetAttribute ("fillId", string (fillIdOf key)); XmlPatch.SetAttribute ("applyFill", "1") ]
                              if key.Border <> was.Border then yield! [ XmlPatch.SetAttribute ("borderId", string (borderIdOf key)); XmlPatch.SetAttribute ("applyBorder", "1") ] ]
                        let sameAlignment =
                            key.HorizontalAlign = was.HorizontalAlign && key.VerticalAlign = was.VerticalAlign && key.WrapText = was.WrapText
                            && key.Indent = was.Indent && key.ShrinkToFit = was.ShrinkToFit && key.TextRotation = was.TextRotation
                        let alignment =
                            if sameAlignment then []
                            else
                                let known = XmlPatch.child "alignment" source
                                let edits =
                                    [ attributeEdit "horizontal" key.HorizontalAlign
                                      attributeEdit "vertical" key.VerticalAlign
                                      attributeEdit "wrapText" (if key.WrapText then Some "1" else None)
                                      attributeEdit "indent" (if key.Indent > 0 then Some (string key.Indent) else None)
                                      attributeEdit "shrinkToFit" (if key.ShrinkToFit then Some "1" else None)
                                      attributeEdit "textRotation" (if key.TextRotation <> 0 then Some (string key.TextRotation) else None) ]
                                let after =
                                    edits |> List.fold (fun (now: Map<string, string>) edit ->
                                        match edit with
                                        | XmlPatch.SetAttribute (name, value) -> Map.add name value now
                                        | XmlPatch.RemoveAttribute name -> Map.remove name now
                                        | _ -> now) (known |> Option.map XmlPatch.attributes |> Option.defaultValue Map.empty)
                                let made =
                                    "<alignment"
                                    + (edits |> List.choose (function XmlPatch.SetAttribute (name, value) -> Some (" " + attr name value) | _ -> None) |> String.concat "")
                                    + "/>"
                                XmlPatch.SetAttribute ("applyAlignment", "1")
                                :: (match known with
                                    | Some _ when after.IsEmpty -> [ XmlPatch.RemoveChild "alignment" ]
                                    | Some node -> [ XmlPatch.EditChild (XmlPatch.name node, edits, None) ]
                                    | None when after.IsEmpty -> []
                                    // (The first child of an xf.)
                                    | None -> [ XmlPatch.SetChild ("alignment", made) ])
                        let index = xfCount + newXfs.Count
                        newXfs.Add (XmlPatch.apply order (XmlPatch.at part source) (ids @ alignment))
                        derived.[struct (basedOn, key)] <- index
                        Some index
            let byIdentity = System.Collections.Generic.Dictionary<XlsxCellFormat, int>(HashIdentity.Reference)
            let styleIdOf (format: XlsxCellFormat) =
                match format.SourceXf with
                | Some index when index >= 0 && index < xfCount -> index
                | _ ->
                    match byIdentity.TryGetValue format with
                    | true, index -> index
                    | false, _ ->
                        let key = xfKey format
                        let index =
                            match format.BasedOnXf |> Option.bind (fun basedOn -> derivedIndex basedOn key) with
                            | Some index -> index
                            | None -> indexOfKey key
                        byIdentity.[format] <- index
                        index
            let sourceDxfs = itemsOf "dxfs" "dxf" |> Option.defaultValue [] |> List.toArray
            let dxfArray = workbook.Dxfs |> List.toArray
            let derivedDxfXml (d: XlsxReader.Dxf) =
                match d.SourceDxfIndex with
                | None -> dxfXml "dxf" d
                | Some index when index >= 0 && index < sourceDxfs.Length ->
                    let original = dxfArray.[index]
                    let putBool name before after =
                        if before = after then None
                        else Some (match after with
                                   | Some true -> XmlPatch.SetChild (name, "<" + name + "/>")
                                   | Some false -> XmlPatch.SetChild (name, if name = "u" then "<u val=\"none\"/>" else "<" + name + " val=\"0\"/>")
                                   | None -> XmlPatch.RemoveChild name)
                    let colorEdit name before after =
                        if before = after then None
                        else Some (match xmlColor after with
                                   | Some rgb -> XmlPatch.SetChild (name, sprintf "<%s rgb=\"%s\"/>" name rgb)
                                   | None -> XmlPatch.RemoveChild name)
                    let fontEdits =
                        [ putBool "b" original.Bold d.Bold
                          putBool "i" original.Italic d.Italic
                          putBool "u" original.Underline d.Underline
                          putBool "strike" original.Strike d.Strike
                          colorEdit "color" original.FontColor d.FontColor ]
                        |> List.choose id
                    let fillEdits =
                        if original.FillColor = d.FillColor then []
                        else
                            match xmlColor d.FillColor with
                            | None -> [ XmlPatch.RemoveChild "fill" ]
                            | Some rgb ->
                                let colorName =
                                    XmlPatch.whole part sourceDxfs.[index]
                                    |> Option.bind (fun source ->
                                        XmlPatch.child "fill" source.Node
                                        |> Option.bind (XmlPatch.child "patternFill")
                                        |> Option.bind (XmlPatch.child "fgColor"))
                                    |> Option.map (fun _ -> "fgColor")
                                    |> Option.defaultValue "bgColor"
                                let child = sprintf "<%s rgb=\"%s\"/>" colorName rgb
                                [ XmlPatch.EditChild ("fill",
                                    [ XmlPatch.EditChild ("patternFill", [ XmlPatch.SetChild (colorName, child) ], Some "<patternFill/>") ],
                                    Some "<fill/>") ]
                    match XmlPatch.whole part sourceDxfs.[index] with
                    | Some source ->
                        let edits = (if fontEdits.IsEmpty then [] else [ XmlPatch.EditChild ("font", fontEdits, Some "<font/>") ]) @ fillEdits
                        XmlPatch.apply order source edits
                    | None -> failwith "Cannot preserve the imported conditional-format differential style."
                | _ -> failwith "The original conditional-format differential style is unavailable."
            let finish () =
                let newDxfs = workbook.Dxfs |> List.skip dxfCount |> List.map derivedDxfXml
                // (The part is the file's text; a list written `<dxfs count="0"/>` is opened only to add to it.)
                let appendTo (name: string) (item: string) (items: string list) =
                    if items.IsEmpty then []
                    else
                        match list name with
                        | Some node ->
                            let had = itemsOf name item |> Option.map List.length |> Option.defaultValue 0
                            [ XmlPatch.EditChild (XmlPatch.name node, [ XmlPatch.SetAttribute ("count", string (had + items.Length)); XmlPatch.Append (String.concat "" items) ], None) ]
                        | None -> [ XmlPatch.SetChild (name, sprintf "<%s count=\"%d\">%s</%s>" name items.Length (String.concat "" items) name) ]
                let edits =
                    appendTo "numFmts" "numFmt" (List.ofSeq newNumFmts)
                    @ appendTo "fonts" "font" (List.ofSeq newFonts)
                    @ appendTo "fills" "fill" (List.ofSeq newFills)
                    @ appendTo "borders" "border" (List.ofSeq newBorders)
                    @ appendTo "cellXfs" "xf" (List.ofSeq newXfs)
                    @ appendTo "dxfs" "dxf" newDxfs
                if edits.IsEmpty then stylesXml else XmlPatch.applyPart order part edits
            Some ({ NamedStyles = kept.NamedStyles; Dxfs = workbook.Dxfs; Fonts = []; Fills = []; Borders = []; Xfs = []; NumFmtIds = Map.empty; StyleIndex = Map.empty
                    KeyIndex = (fun key -> Some (indexOfKey key)); StyleIdOf = styleIdOf },
                  finish)
    | _ -> None

/// The file's shared strings as a table that only adds: a text the file has
/// is its entry of the file, a new one is added behind the file's. The second
/// value gives the part when all cells have been written. None when the part
/// does not have the entries the reader counted.
let private inPlaceSharedStrings (workbook: XlsxWorkbookOut) (sharedStringsXml: string) : ((string -> XlsxRichRun list -> int) * (unit -> string)) option =
    let texts, runs = workbook.Preserved.SharedStrings
    let source =
        if sharedStringsXml.Trim() = "" then declaration + "\n<sst xmlns=\"" + Namespaces.spreadsheet + "\" count=\"0\" uniqueCount=\"0\"></sst>"
        else sharedStringsXml
    match XmlPatch.readPartExcept (fun name -> name = "si") source with
    | Some root when XmlPatch.localName root.Node = "sst" && (XmlPatch.elements root.Node |> List.filter (fun node -> XmlPatch.localName node = "si") |> List.length) = texts.Length ->
        let plain = System.Collections.Generic.Dictionary<string, int>()
        let rich = System.Collections.Generic.Dictionary<SharedStringKey, int>(HashIdentity.Structural)
        for index in 0 .. texts.Length - 1 do
            if index >= runs.Length || runs.[index].IsEmpty then
                if not (plain.ContainsKey texts.[index]) then plain.[texts.[index]] <- index
            else
                let key = { Text = texts.[index]; Runs = runs.[index] |> List.map (fun (text, props) -> richRunOfReader props text) }
                if not (rich.ContainsKey key) then rich.[key] <- index
        let added = StringBuilder()
        let mutable count = texts.Length
        let indexOf (text: string) (textRuns: XlsxRichRun list) : int =
            if textRuns.IsEmpty then
                match plain.TryGetValue text with
                | true, index -> index
                | false, _ ->
                    let index = count
                    count <- count + 1
                    appendSharedStringItem added { Text = text; Runs = [] }
                    plain.[text] <- index
                    index
            else
                let key = { Text = text; Runs = textRuns }
                match rich.TryGetValue key with
                | true, index -> index
                | false, _ ->
                    let index = count
                    count <- count + 1
                    appendSharedStringItem added key
                    rich.[key] <- index
                    index
        let finish () =
            if count = texts.Length then source
            else
                let was = match XmlPatch.attribute root.Node "count" |> Option.map System.Int32.TryParse with Some (true, n) -> n | _ -> texts.Length
                XmlPatch.applyPart order root
                    [ XmlPatch.SetAttribute ("uniqueCount", string count)
                      XmlPatch.SetAttribute ("count", string (was + count - texts.Length))
                      XmlPatch.AddChild ("si", added.ToString()) ]
        Some (indexOf, finish)
    | _ -> None

/// A drawing part of the file with its anchors where they are now
/// (`SheetExtras.DrawingAnchors`): the `<from>` / `<to>` of an anchor that
/// moved are changed - a marker that only changed its row or column keeps
/// the offsets the file wrote -, everything else is the part's text. None
/// when the part's anchors are not the ones given (another count, another
/// kind, an anchor inside `mc:AlternateContent`).
let private movedDrawingXml (xml: string) (anchors: XlsxReader.SheetAnchor list) : string option =
    match XmlPatch.readPart xml with
    | None -> None
    | Some part ->
        let isAnchor (node: XmlDsl.XmlNode) = List.contains (XmlPatch.localName node) [ "twoCellAnchor"; "oneCellAnchor"; "absoluteAnchor" ]
        let all = XmlPatch.descendantsWhere isAnchor part.Node
        let kids = XmlPatch.elements part.Node
        let placed = kids |> List.indexed |> List.filter (fun (_, kid) -> isAnchor kid)
        // (Every anchor is a child of the drawing, one after another.)
        let sound =
            all.Length = anchors.Length && placed.Length = all.Length
            && (placed.IsEmpty || fst (List.last placed) - fst placed.Head + 1 = placed.Length)
        if not sound then None
        else
            let mutable sound = true
            let targets =
                [ for node, anchor in List.zip all (anchors |> List.sortBy (fun anchor -> anchor.Order)) do
                    let expected = (match anchor.Kind with "twoCell" -> "twoCellAnchor" | "oneCell" -> "oneCellAnchor" | _ -> "absoluteAnchor")
                    if XmlPatch.localName node <> expected then sound <- false
                    else
                        // A marker: its cell and offsets set where they moved.
                        let marker (name: string) (col: int) (colOffset: float) (row: int) (rowOffset: float) =
                            match XmlPatch.child name node with
                            | None -> sound <- false; []
                            | Some markerNode ->
                                let part (tag: string) =
                                    XmlPatch.child tag markerNode
                                    |> Option.bind (fun value -> match System.Int64.TryParse ((XmlPatch.textOf value).Trim()) with | true, number -> Some (float number) | _ -> None)
                                match part "col", part "colOff", part "row", part "rowOff" with
                                | Some fileCol, Some fileColOff, Some fileRow, Some fileRowOff ->
                                    // (Offsets are held in pixels: within a pixel of the file's, the file's stands.)
                                    let colOff = if abs (colOffset - fileColOff) < 9525.0 then fileColOff else colOffset
                                    let rowOff = if abs (rowOffset - fileRowOff) < 9525.0 then fileRowOff else rowOffset
                                    if int fileCol <> col || int fileRow <> row || colOff <> fileColOff || rowOff <> fileRowOff then
                                        let set (tag: string) (value: string) =
                                            let child = (XmlPatch.child tag markerNode).Value
                                            XmlPatch.EditChild (XmlPatch.name child, [ XmlPatch.SetText value ], None)
                                        [ markerNode, [ set "col" (string col); set "colOff" (emu colOff); set "row" (string row); set "rowOff" (emu rowOff) ] ]
                                    else []
                                | _ -> sound <- false; []
                        if anchor.Kind <> "absolute" then
                            yield! marker "from" anchor.FromCol anchor.FromColOffset anchor.FromRow anchor.FromRowOffset
                            match anchor.ToCol, anchor.ToRow with
                            | Some toCol, Some toRow -> yield! marker "to" toCol anchor.ToColOffset toRow anchor.ToRowOffset
                            | _ -> () ]
            if not sound then None
            elif targets.IsEmpty then Some xml
            else Some (XmlPatch.applyPart (fun _ -> []) part (XmlPatch.editsAt part.Node targets))

/// The `<hyperlinks>` of a sheet whose part stays the file's, for the links the sheet has
/// now: a link that is what the file says of it - its cells, where it leads, its tip and
/// its text - is the file's element as it is written (its relationship stays); another one
/// is a new element, an external one with a relationship under an id from `freeId`.
/// `fileLinks`: the file's `<hyperlinks>` element ("" when it has none); `entries`: the
/// sheet's relationships as the file has them; `namespaceText`: what declares the `r`
/// prefix on a new element when the worksheet does not. Gives the element ("" for no
/// link), the relationships to add (id, target) and the entries of the file to take out
/// (the relationships of links that are gone).
let private hyperlinksInPlace (fileLinks: XmlPatch.SourceElement option) (entries: (string * string * string * bool) list) (links: XlsxHyperlink list)
                              (freeId: unit -> string) (namespaceText: string) : string * (string * string) list * string list =
    let linkRels = entries |> List.filter (fun (_, relType, _, isExternal) -> isExternal && OpcXml.relationshipKind relType = "hyperlink")
    let known =
        match fileLinks with
        | Some fileLinks ->
            XmlPatch.elements fileLinks.Node
            |> List.filter (fun node -> XmlPatch.name node = "hyperlink")
            |> List.map (fun node ->
                let value (name: string) = XmlPatch.attribute node name
                let id = value "r:id"
                let url = id |> Option.bind (fun id -> linkRels |> List.tryPick (fun (known, _, target, _) -> if known = id then Some target else None))
                // (As the reader takes it.)
                let target, isInternal =
                    match url, value "location" with
                    | Some url, Some place when place <> "" -> url + "#" + place, false
                    | Some url, _ -> url, false
                    | None, Some place -> place, true
                    | None, None -> defaultArg id "", false
                XmlPatch.bytesOf fileLinks.Text node, id,
                ({ Range = defaultArg (value "ref") ""; Target = target; IsInternal = isInternal; Tooltip = value "tooltip"; Display = value "display" } : XlsxHyperlink))
        | None -> []
    let taken = System.Collections.Generic.HashSet<int>()
    let keptIds = System.Collections.Generic.HashSet<string>()
    let added = ResizeArray<string * string>()
    let body = ResizeArray<string>()
    for link in links do
        let same = known |> List.indexed |> List.tryFind (fun (at, (_, _, fileLink)) -> not (taken.Contains at) && fileLink = link)
        match same with
        | Some (at, (text, id, _)) ->
            taken.Add at |> ignore
            (match id with Some id -> keptIds.Add id |> ignore | None -> ())
            body.Add text
        | None ->
            let extra =
                (match link.Tooltip with Some tip -> " " + attr "tooltip" tip | None -> "")
                + (match link.Display with Some text -> " " + attr "display" text | None -> "")
            if link.IsInternal then body.Add (sprintf "<hyperlink %s %s%s/>" (attr "ref" link.Range) (attr "location" link.Target) extra)
            else
                let id = freeId ()
                added.Add ((id, link.Target))
                body.Add (sprintf "<hyperlink%s %s r:id=\"%s\"%s/>" namespaceText (attr "ref" link.Range) id extra)
    let gone = linkRels |> List.choose (fun (id, _, _, _) -> if keptIds.Contains id then None else Some id)
    let element =
        if body.Count = 0 then ""
        else
            match fileLinks with
            | Some fileLinks -> XmlPatch.apply order fileLinks [ XmlPatch.SetChildren ("hyperlink", List.ofSeq body) ]
            | None -> "<hyperlinks>" + String.concat "" body + "</hyperlinks>"
    element, List.ofSeq added, gone

/// The notes of a sheet whose part goes back as the file's (see
/// `SheetPatch.Notes`): the parts they are written into.
type private NotesInPlace = {
    /// The file's comments part (None: the sheet had none, a part is added).
    CommentsPath: string option
    CommentsXml: string
    /// The file's VML drawing (None: the sheet had none, a part is added).
    VmlPath: string option
    /// The drawing with the boxes of the notes (None: the file's has them).
    VmlXml: string option
    /// Parts of the file that go, with their relationships - the sheet's last note was
    /// removed: its comments part, and the VML drawing when it held nothing but the boxes
    /// of the notes (the sheet's part then loses its `<legacyDrawing>`).
    Gone: string list
}


let private workbookFiles (workbook: XlsxWorkbookOut) =
    // Styles and shared strings: the file's, added to (so that a sheet written
    // as it was read still means what it meant), or made from the cells.
    let partText (path: string) =
        workbook.Preserved.Parts |> List.tryPick (fun (kept, data) -> if kept = path then Some (System.Text.Encoding.UTF8.GetString data) else None)
    let inPlace =
        if not workbook.Preserved.InPlace then None
        else
            let sharedText =
                match workbook.Preserved.SharedStringsPart with
                | Some path -> partText path
                | None -> Some ""
            match workbook.Preserved.StylesPart |> Option.bind partText |> Option.bind (inPlaceStyles workbook), sharedText |> Option.bind (inPlaceSharedStrings workbook) with
            | Some styles, Some strings -> Some (styles, strings)
            | _ -> None
    if inPlace.IsNone && (workbook.Dxfs |> List.exists (fun dxf -> dxf.SourceDxfIndex.IsSome)) then
        failwith "The original styles part is unavailable; cannot save an edited imported conditional format without losing its differential style."
    // Without them a worksheet cannot go back as it was read: it is written from its cells.
    let workbook =
        if inPlace.IsSome then workbook
        else
            { workbook with
                Sheets =
                    workbook.Sheets
                    |> List.map (fun sheet ->
                        match sheet.Extras.Verbatim, sheet.Extras.SourcePart with
                        | true, Some (_, relType) when OpcXml.relationshipKind relType = "worksheet" -> { sheet with Extras = { sheet.Extras with Verbatim = false } }
                        | _ -> sheet) }
    // A sheet that is the file's but for some cells: how its part is patched,
    // or - when it cannot be - the sheet is written from its cells.
    // (Or its rows are the file's with those cells, in a part otherwise made
    // from the sheet: `rowPlans`. `sortKept`: the file's sort state holds.)
    // A drawing part whose objects only moved: the file's part with its anchors moved - or, when
    // its anchors are not what the sheet has, a part made from the sheet as for any other change.
    let movedDrawings = System.Collections.Generic.Dictionary<string, string>()
    let workbook =
        if workbook.Sheets |> List.forall (fun sheet -> sheet.Extras.DrawingAnchors.IsNone) then workbook
        else
            { workbook with
                Sheets =
                    workbook.Sheets
                    |> List.map (fun sheet ->
                        match sheet.Extras.DrawingAnchors, sheet.Extras.DrawingPart with
                        | Some anchors, Some path ->
                            match partText path |> Option.bind (fun xml -> movedDrawingXml xml anchors) with
                            | Some xml ->
                                movedDrawings.[path] <- xml
                                sheet
                            | None -> { sheet with Extras = { sheet.Extras with DrawingPart = None; DrawingAnchors = None } }
                        | Some _, None -> { sheet with Extras = { sheet.Extras with DrawingAnchors = None } }
                        | None, _ -> sheet) }
    let patchPlans = System.Collections.Generic.Dictionary<int, SheetPatchPlan>()
    // A sheet whose notes changed and whose part can stay the file's: the notes go into the
    // comments part and the VML drawing the file has for them, or into new ones (the part then
    // gets a `<legacyDrawing>` and its relationships two entries). None when that cannot be
    // done here (no note left, a comments part without a list, a part that names a drawing
    // the reader did not keep): the part is then written from the sheet, its rows the file's.
    let notePlans = System.Collections.Generic.Dictionary<int, NotesInPlace>()
    let notesInPlace (number: int) (sheet: XlsxSheetOut) (source: XmlPatch.SourceElement) : NotesInPlace option =
        // (The `<legacyDrawing>` of the sheet's part, and those that name a drawing.)
        let legacyDrawings = childrenNamed "legacyDrawing" source.Node
        let legacyIds = legacyDrawings |> List.choose (fun node -> XmlPatch.attribute node "r:id")
        let commentCells = sheet.Cells |> List.choose (fun c -> if c.Comment.IsSome then Some (c.Row, c.Col) else None)
        if commentCells.IsEmpty then
            // No note is left: the comments part goes; the VML drawing goes with it when it held
            // nothing but their boxes, else it stays without them. (By what the file's part and
            // its relationships say: the sheet itself no longer knows of either.)
            match sheet.Extras.SourcePart with
            | Some (sheetPath, _) ->
                let baseDir = PartPath.directory sheetPath
                let entries = partText (relsPathOf sheetPath) |> Option.map (relationshipEntries baseDir) |> Option.defaultValue []
                let comments = entries |> List.filter (fun (_, relType, _, isExternal) -> not isExternal && OpcXml.relationshipKind relType = "comments")
                match comments with
                | [ (_, _, commentsPath, _) ] when legacyIds.Length = 1 && (partText commentsPath).IsSome ->
                    let id = legacyIds.Head
                    let drawing =
                        entries |> List.tryPick (fun (known, _, target, isExternal) ->
                            if known <> id || isExternal then None else partText target |> Option.map (fun vml -> target, vml))
                    match drawing with
                    | Some (vmlPath, vml) ->
                        let notesOnly = XmlPatch.readPart vml |> Option.exists (fun drawing -> vmlShapes drawing |> List.forall isNoteBox)
                        if notesOnly then Some { CommentsPath = Some commentsPath; CommentsXml = ""; VmlPath = Some vmlPath; VmlXml = None; Gone = [ commentsPath; vmlPath ] }
                        else
                            match vmlWithCommentBoxes vml [] with
                            | Some without -> Some { CommentsPath = Some commentsPath; CommentsXml = ""; VmlPath = Some vmlPath; VmlXml = Some without; Gone = [ commentsPath ] }
                            | None -> None
                    | None -> None
                // (A file whose notes have no boxes - not written by Excel: only the comments part goes.
                // `VmlPath` names no part: none is made, none goes.)
                | [ (_, _, commentsPath, _) ] when legacyDrawings.IsEmpty && (partText commentsPath).IsSome ->
                    Some { CommentsPath = Some commentsPath; CommentsXml = ""; VmlPath = Some ""; VmlXml = None; Gone = [ commentsPath ] }
                | _ -> None
            | None -> None
        else
            let relsText = sheet.Extras.SourcePart |> Option.bind (fun (path, _) -> partText (relsPathOf path))
            let comments =
                match keptCommentsPart workbook sheet with
                | Some path ->
                    partText path
                    |> Option.bind (fun fileXml -> patchedCommentsXml fileXml sheet.Extras.CommentRuns sheet.Cells)
                    |> Option.map (fun xml -> Some path, xml)
                | None ->
                    if relsText |> Option.exists (fun text -> OpcXml.relationshipsOrEmpty text |> List.exists (fun rel -> OpcXml.relationshipKind rel.Type = "comments")) then None
                    else commentsXml sheet.Extras.CommentRuns sheet.Cells |> Option.map (fun xml -> None, xml)
            let boxes =
                match keptLegacyDrawing sheet with
                | "" ->
                    if not legacyDrawings.IsEmpty then None
                    else vmlCommentsXml number sheet.Cells |> Option.map (fun xml -> None, Some xml)
                | keptElement ->
                    let ids = OpcXml.relationshipIdsIn keptElement
                    sheet.Extras.KeptRelationships
                    |> List.tryPick (fun (id, _, target, isExternal) ->
                        if isExternal || not (List.contains id ids) then None
                        else
                            let path = ZipInterop.normalizePath target
                            partText path |> Option.map (fun vml -> Some path, vmlWithCommentBoxes vml commentCells))
            match comments, boxes with
            | Some (commentsPath, commentsXml), Some (vmlPath, vmlXml) ->
                Some { CommentsPath = commentsPath; CommentsXml = commentsXml; VmlPath = vmlPath; VmlXml = vmlXml; Gone = [] }
            | _ -> None
    let rowPlans = System.Collections.Generic.Dictionary<int, SheetPatchPlan * XlsxReader.SheetPatch>()
    let sortKept = System.Collections.Generic.HashSet<int>()
    let workbook =
        if workbook.Sheets |> List.forall (fun sheet -> sheet.Extras.Patch.IsNone) then workbook
        else
            { workbook with
                Sheets =
                    workbook.Sheets
                    |> List.mapi (fun index sheet ->
                        match sheet.Extras.Patch with
                        | None -> sheet
                        | Some patch ->
                            let sourceText =
                                if sheet.Extras.Verbatim then sheet.Extras.SourcePart |> Option.bind (fun (path, _) -> partText path)
                                else sheet.Extras.SourceXml
                            let plan = sourceText |> Option.bind (sheetPatchPlan workbook.Dxfs.Length sheet patch)
                            if not (List.contains "sortState" patch.Sections) then sortKept.Add index |> ignore
                            // (Written from the sheet, it keeps what its part says beyond that.)
                            let fromSheet = { sheet.Extras with Verbatim = false; Patch = None; SourceXml = sourceText; Moves = (match sheet.Extras.Moves with Some moves -> Some moves | None -> Some []) }
                            match plan with
                            | Some plan when sheet.Extras.Verbatim && plan.SectionsSound && not patch.Whole
                                             && (not patch.Notes
                                                 || (match notesInPlace (index + 1) sheet plan.Source with
                                                     | Some notes -> notePlans.[index] <- notes; true
                                                     | None -> false)) ->
                                patchPlans.[index] <- plan
                                sheet
                            | Some plan when inPlace.IsSome ->
                                rowPlans.[index] <- (plan, patch)
                                { sheet with Extras = fromSheet }
                            | _ -> { sheet with Extras = fromSheet }) }
    let styles, stylesText, sharedIndex, sharedText =
        match inPlace with
        | Some ((styles, stylesText), (sharedIndex, sharedText)) -> styles, stylesText, sharedIndex, sharedText
        | None ->
            let styles = collectStyles workbook
            let sharedItems, sharedIndex = collectSharedStrings workbook
            styles, (fun () -> stylesXml styles), sharedIndex, (fun () -> sharedStringsXmlPreserved sharedItems)
    let hasVba = workbook.VbaProjectSource |> Option.exists hasText
    let mutable tableId = 1
    let textFiles = ResizeArray<PackageFile>()
    // The theme stays only if its part was kept.
    let themePath = workbook.Preserved.ThemePath |> Option.filter (fun path -> workbook.Preserved.Parts |> List.exists (fun (kept, _) -> kept = path))
    // (Less the parts that go with a sheet's last note: its part stays the file's, without them.)
    let goneParts = System.Collections.Generic.HashSet<string>()
    for KeyValue(_, notes) in notePlans do
        for path in notes.Gone do
            goneParts.Add path |> ignore
            goneParts.Add (relsPathOf path) |> ignore
    let keptParts =
        let reached = reachableKeptParts workbook themePath
        if goneParts.Count = 0 then reached else reached |> List.filter (fun (path, _) -> not (goneParts.Contains path))
    // Dynamic array formulas name a record of the cell metadata: the opened
    // workbook's metadata part with that record added to it, or a new part.
    let hasDynamicArrays =
        workbook.Sheets |> List.exists (fun sheet -> sheet.Extras.ArrayFormulas |> List.exists (fun (_, _, dynamic) -> dynamic))
    let keptMetadata =
        workbook.Preserved.WorkbookRelationships
        |> List.tryPick (fun (_, relType, target, isExternal) ->
            if isExternal || not (OpcXml.relationshipKind relType = "sheetMetadata") then None
            else
                let path = ZipInterop.normalizePath target
                keptParts |> List.tryFind (fun (kept, _) -> kept = path))
    let ownMetadata = hasDynamicArrays && keptMetadata.IsNone
    let dynamicCm =
        if not hasDynamicArrays then 0
        else
            match keptMetadata with
            | Some (path, data) ->
                let xml, cm = metadataWithDynamicArrays (System.Text.Encoding.UTF8.GetString data)
                textFiles.Add { Path = path; Content = xml }
                cm
            | None ->
                textFiles.Add { Path = "xl/metadata.xml"; Content = dynamicArrayMetadataXml }
                1
    let keptPartPaths = keptParts |> List.map fst |> Set.ofList
    let placements = sheetPlacements workbook keptPartPaths
    // The parts added for the notes of sheets that had none (see `notePlans`), under names no part has.
    let noteParts = System.Collections.Generic.Dictionary<int, string * string>()
    if notePlans.Count > 0 then
        let taken = System.Collections.Generic.HashSet<string>(keptPartPaths)
        for path, _ in workbook.Preserved.Parts do taken.Add path |> ignore
        for placed in placements do
            for path in [ placed.SheetPath; placed.CommentsPath; placed.VmlPath; placed.ThreadsPath; placed.DrawingPath ] do taken.Add path |> ignore
        let free (name: int -> string) =
            let mutable at = 1
            while taken.Contains (name at) do at <- at + 1
            taken.Add (name at) |> ignore
            name at
        for KeyValue(index, notes) in notePlans do
            noteParts.[index] <-
                ((if notes.CommentsPath.IsNone then free (sprintf "xl/comments%d.xml") else ""),
                 (if notes.VmlPath.IsNone then free (sprintf "xl/drawings/vmlDrawing%d.vml") else ""))
    textFiles.Add
        { Path = "[Content_Types].xml"
          Content =
            contentTypesXmlFull workbook placements keptParts ownMetadata
                [ for KeyValue(_, (commentsPath, _)) in noteParts do if commentsPath <> "" then yield commentsPath ] (notePlans.Count > 0) }
    textFiles.Add { Path = "_rels/.rels"; Content = rootRelsXmlFull workbook }
    for KeyValue(path, xml) in movedDrawings do
        if keptPartPaths.Contains path then textFiles.Add { Path = path; Content = xml }
    textFiles.Add { Path = "xl/workbook.xml"; Content = workbookXmlFull workbook }
    let persons = personsOf workbook
    let hasThreads = workbook.Sheets |> List.exists (fun sheet -> not sheet.Extras.ThreadedComments.IsEmpty)
    textFiles.Add { Path = "xl/_rels/workbook.xml.rels"; Content = workbookRelsXmlFull workbook placements hasVba themePath hasThreads ownMetadata }
    if hasThreads then textFiles.Add { Path = "xl/persons/person.xml"; Content = personsXml persons }
    let mediaFiles = ResizeArray<string * byte[]>()
    let mutable nextMedia = 1
    // The same picture several times in the workbook is one media part.
    let sameBytes (a: byte[]) (b: byte[]) =
        obj.ReferenceEquals(a, b)
        || (a.Length = b.Length
            && (let mutable same = true
                let mutable i = 0
                while same && i < a.Length do
                    if a.[i] <> b.[i] then same <- false
                    i <- i + 1
                same))
    let mediaExtensions = System.Collections.Generic.Dictionary<string, string>()
    let placement = tablePlacement workbook
    // How many more cells name a shared string, by sheet written (None: not known).
    let sharedDelta = System.Collections.Generic.Dictionary<int, int option>()
    List.zip workbook.Sheets placements
    |> List.iteri (fun i (sheet, placed) ->
      // (The cells written that name a shared string.)
      let sharedUses = ref 0
      let countedShared (text: string) (runs: XlsxRichRun list) =
          sharedUses.Value <- sharedUses.Value + 1
          sharedIndex text runs
      // (A sheet written as it was read is one of the kept parts; only whether it is the one on display may differ.)
      if placed.Verbatim.IsSome && patchPlans.ContainsKey i then
        // The file's part with the cells and sections that changed.
        // Its relationships are the file's, with what the edit adds behind them - under ids the
        // file does not use, the targets relative to the sheet's folder as Excel writes them -
        // and without what it ends (the relationship of a link that is gone).
        let relsPath = relsPathOf placed.SheetPath
        let relsWas =
            partText relsPath
            |> Option.defaultValue (sprintf "%s\n<Relationships xmlns=\"%s\"></Relationships>" declaration Namespaces.relationshipsPackage)
        let ids = System.Collections.Generic.HashSet<string>(OpcXml.relationshipsOrEmpty relsWas |> List.map (fun rel -> rel.Id))
        let freeId () =
            let mutable number = 1
            while not (ids.Add (sprintf "rId%d" number)) do number <- number + 1
            sprintf "rId%d" number
        let targetOf (path: string) =
            if PartPath.directory placed.SheetPath = "xl/worksheets" && PartPath.isIn "xl" path then PartPath.relative "xl/worksheets" path
            else "/" + path
        let removedRels = ResizeArray<string>()
        let addedRels = ResizeArray<OpcXml.Relationship>()
        let plan = patchPlans.[i]
        let namespaceText = if (XmlPatch.attribute plan.Source.Node "xmlns:r").IsSome then "" else sprintf " xmlns:r=\"%s\"" Namespaces.relationships
        // (Its notes, when they changed: in the comments part and the VML drawing; the part itself
        // says nothing of them but that it has a drawing of that kind.)
        let legacy =
            match notePlans.TryGetValue i with
            | true, notes ->
                let newComments, newVml = noteParts.[i]
                let gone (path: string option) = match path with Some path -> List.contains path notes.Gone | None -> false
                if not (gone notes.CommentsPath) then
                    textFiles.Add { Path = (match notes.CommentsPath with Some path -> path | None -> newComments); Content = notes.CommentsXml }
                match notes.VmlXml with
                | Some xml when not (gone notes.VmlPath) -> textFiles.Add { Path = (match notes.VmlPath with Some path -> path | None -> newVml); Content = xml }
                | _ -> ()
                if not notes.Gone.IsEmpty then
                    for id, _, target, isExternal in relationshipEntries (PartPath.directory placed.SheetPath) relsWas do
                        if not isExternal && List.contains target notes.Gone then removedRels.Add id
                if newComments <> "" then
                    addedRels.Add { Id = freeId (); Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"; Target = targetOf newComments; External = false }
                if gone notes.VmlPath then [ "legacyDrawing", "" ]
                elif newVml = "" then []
                else
                    let id = freeId ()
                    addedRels.Add { Id = id; Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"; Target = targetOf newVml; External = false }
                    [ "legacyDrawing", sprintf "<legacyDrawing%s r:id=\"%s\"/>" namespaceText id ]
            | false, _ -> []
        // Its links, when they changed: the list in the part, an external one's address in the relationships.
        let links =
            if not (List.contains "hyperlinks" sheet.Extras.Patch.Value.Sections) then []
            else
                let element, added, gone =
                    hyperlinksInPlace (sourceChild "hyperlinks" plan.Source) (relationshipEntries (PartPath.directory placed.SheetPath) relsWas) sheet.Hyperlinks freeId namespaceText
                for id, target in added do
                    addedRels.Add { Id = id; Type = RelationshipTypes.hyperlink; Target = target; External = true }
                removedRels.AddRange gone
                [ "hyperlinks", element ]
        if addedRels.Count > 0 || removedRels.Count > 0 then
            textFiles.Add { Path = relsPath; Content = OpcXml.editRelationships (List.ofSeq removedRels) (List.ofSeq addedRels) relsWas }
        let added = legacy @ links
        let xml, sharedGone = patchedSheetXml styles countedShared dynamicCm (i + 1) (i = workbook.ActiveSheetIndex) sheet sheet.Extras.Patch.Value plan added
        sharedDelta.[i] <- Some (sharedUses.Value - sharedGone)
        textFiles.Add { Path = placed.SheetPath; Content = xml }
      elif placed.Verbatim.IsSome then
        match keptParts |> List.tryFind (fun (path, _) -> path = placed.SheetPath) with
        | Some (_, data) ->
            let xml = System.Text.Encoding.UTF8.GetString data
            let shown = withTabSelected (i = workbook.ActiveSheetIndex) xml
            if not (obj.ReferenceEquals (shown, xml)) then textFiles.Add { Path = placed.SheetPath; Content = shown }
        | None -> ()
      else
        let sheetIndex = i + 1
        let tableParts = placement.[i]
        // The sheet's drawing: the file's part as it is, or one made from the pictures and objects.
        let keptDrawing = keptDrawingPart workbook sheet
        let drawingPath =
            match keptDrawing with
            | Some path -> Some path
            | None -> if hasDrawing sheet then Some placed.DrawingPath else None
        // (The cells of the file that named a shared string and are not written: all of a sheet written
        // from its cells, those that changed of rows that are the file's.)
        let rows, sharedGone =
            match rowPlans.TryGetValue i with
            | true, (plan, patch) ->
                let edits, _, gone = sheetDataEdits styles countedShared dynamicCm sheet patch plan
                FileRows (patchedSheetData plan edits), Some gone
            | false, _ ->
                match sheet.Extras.SourceXml with
                | Some xml ->
                    let carry = rowCarryOf inPlace.IsSome sheet.Extras.Moves xml
                    CellRows carry, (carry |> Option.bind (fun carry -> carry.SharedStrings))
                | None ->
                    CellRows None,
                    (match sheet.Extras.SourcePart with
                     | Some (path, _) -> partText path |> Option.bind readWorksheetCells |> Option.map sharedStringCells
                     | None -> Some 0)          // (a sheet the file does not have)
        let sheetXml = worksheetXml styles countedShared dynamicCm sheetIndex (i = workbook.ActiveSheetIndex) drawingPath.IsSome rows (sortKept.Contains i) sheet
        sharedDelta.[i] <- sharedGone |> Option.map (fun gone -> sharedUses.Value - gone)
        let commentsVmlPath = placed.VmlPath
        // The sheet's relationships: the file's part while it says what is written - the same
        // parts, by the ids the file names them by (the sheet's part then names them by those
        // too), with the targets as the file wrote them -, else with what is new behind it.
        let sheetXml, relsXml =
            let made = sheetRelsXml sheetIndex sheet tableParts placed drawingPath
            let fileRels =
                match sheet.Extras.SourcePart with
                | Some (path, _) when path = placed.SheetPath -> partText (relsPathOf path)
                | _ -> None
            match made, fileRels with
            | Some made, Some source ->
                let baseDir = PartPath.directory placed.SheetPath
                let was = relationshipEntries baseDir source
                let now = relationshipEntries baseDir made
                let key (relType: string) (target: string) (isExternal: bool) = relType + "\n" + target + (if isExternal then "\nexternal" else "")
                let taken = System.Collections.Generic.HashSet<string>()
                let renames =
                    now
                    |> List.choose (fun (id, relType, target, isExternal) ->
                        was
                        |> List.tryFind (fun (fileId, fileType, fileTarget, fileExternal) -> fileId <> "" && key fileType fileTarget fileExternal = key relType target isExternal && not (taken.Contains fileId))
                        |> Option.map (fun (fileId, _, _, _) -> taken.Add fileId |> ignore; id, fileId))
                let changed = renames |> List.filter (fun (id, fileId) -> id <> fileId) |> Map.ofList
                // (An id the sheet made for another part must not be one the file uses.)
                let clash = now |> List.exists (fun (id, _, _, _) -> not (renames |> List.exists (fun (known, _) -> known = id)) && taken.Contains id)
                if clash then sheetXml, Some made
                elif changed.IsEmpty then sheetXml, Some (relationshipsFromFile baseDir true (Some source) made)
                else
                    // (Every entry renamed at once: an id may be taken over by another entry.)
                    let renamedMade =
                        match XmlPatch.readPart made with
                        | Some rels ->
                            XmlPatch.applyPart order rels
                                [ for KeyValue (id, fileId) in changed -> XmlPatch.EditChildWhere ("Relationship", "Id", id, [ XmlPatch.SetAttribute ("Id", fileId) ]) ]
                        | None -> made
                    let renamedSheet =
                        match readWorksheet sheetXml with
                        | Some part ->
                            let named =
                                XmlPatch.elements part.Node
                                |> List.filter (fun node -> XmlPatch.name node <> "sheetData")
                                |> List.collect (XmlPatch.descendantsWhere (fun node -> XmlPatch.attribute node "r:id" |> Option.exists changed.ContainsKey))
                            XmlPatch.applyPart order part
                                (XmlPatch.editsAt part.Node [ for node in named -> node, [ XmlPatch.SetAttribute ("r:id", changed.[XmlPatch.attribute node "r:id" |> Option.get]) ] ])
                        | None -> sheetXml
                    renamedSheet, Some (relationshipsFromFile baseDir true (Some source) renamedMade)
            | made, _ -> sheetXml, made
        textFiles.Add { Path = placed.SheetPath; Content = sheetXml }
        match relsXml with
        | Some rels -> textFiles.Add { Path = relsPathOf placed.SheetPath; Content = rels }
        | None -> ()
        match keptCommentsPart workbook sheet with
        | Some path when sheet.Extras.CommentsChanged ->
            // The file's part with the notes that changed (or, a part without a list of comments, one made from the notes).
            match partText path |> Option.bind (fun fileXml -> patchedCommentsXml fileXml sheet.Extras.CommentRuns sheet.Cells) with
            | Some xml -> textFiles.Add { Path = path; Content = xml }
            | None ->
                match commentsXml sheet.Extras.CommentRuns sheet.Cells with
                | Some xml -> textFiles.Add { Path = path; Content = xml }
                | None -> ()
        | Some _ -> ()
        | None ->
            match commentsXml sheet.Extras.CommentRuns sheet.Cells with
            | Some xml -> textFiles.Add { Path = placed.CommentsPath; Content = xml }
            | None -> ()
        match keptLegacyDrawing sheet with
        | "" ->
            match vmlCommentsXml sheetIndex sheet.Cells with
            | Some xml -> textFiles.Add { Path = commentsVmlPath; Content = xml }
            | None -> ()
        | keptElement ->
            // The file's own VML drawing stands (it may hold form controls):
            // the boxes of the comments in it are made those of the comments now.
            let ids = OpcXml.relationshipIdsIn keptElement
            let part =
                sheet.Extras.KeptRelationships
                |> List.tryPick (fun (id, _, target, isExternal) ->
                    if isExternal || not (List.contains id ids) then None
                    else
                        let path = ZipInterop.normalizePath target
                        keptParts |> List.tryFind (fun (kept, _) -> kept = path))
            match part with
            | Some (path, data) ->
                let commentCells = sheet.Cells |> List.choose (fun c -> if c.Comment.IsSome then Some (c.Row, c.Col) else None)
                match vmlWithCommentBoxes (System.Text.Encoding.UTF8.GetString data) commentCells with
                | Some xml -> textFiles.Add { Path = path; Content = xml }
                | None -> ()
            | None -> ()
        match threadedCommentsXml persons sheetIndex sheet with
        | Some xml -> textFiles.Add { Path = placed.ThreadsPath; Content = xml }
        | None -> ()
        if hasDrawing sheet && keptDrawing.IsNone then
            // One media part per picture (not under the name of a kept part), and
            // the relationships the kept objects refer to, with their ids.
            let imageRels =
                sheet.Extras.Images
                |> List.mapi (fun i image ->
                    let extension = if image.Extension = "" then "bin" else image.Extension
                    let known = mediaFiles |> Seq.tryFind (fun (path, data) -> mediaExtensions.[path] = extension && sameBytes data image.Data)
                    match known with
                    | Some (path, _) -> sprintf "rIdImg%d" (i + 1), path
                    | None ->
                        while Set.contains (sprintf "xl/media/image%d.%s" nextMedia extension) keptPartPaths do
                            nextMedia <- nextMedia + 1
                        let path = sprintf "xl/media/image%d.%s" nextMedia extension
                        nextMedia <- nextMedia + 1
                        mediaFiles.Add (path, image.Data)
                        mediaExtensions.[path] <- extension
                        sprintf "rIdImg%d" (i + 1), path)
            let relsBody =
                [ for relId, path in imageRels do
                      yield sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"/%s\"/>" relId RelationshipTypes.image path
                  for relId, relType, target, isExternal in sheet.Extras.DrawingRelationships do
                      yield sprintf "<Relationship %s %s %s%s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" target) (if isExternal then " TargetMode=\"External\"" else "") ]
            textFiles.Add { Path = placed.DrawingPath; Content = drawingXml sheet (imageRels |> List.map fst) }
            if not relsBody.IsEmpty then
                textFiles.Add
                    { Path = relsPathOf placed.DrawingPath
                      Content = sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" relsBody) }
        for table, (path, id) in List.zip sheet.Tables tableParts do
            match id with
            | Some id -> textFiles.Add { Path = path; Content = tableXml id table }
            | None ->
                // The file's part: as it is while the table is what it says, else with what changed replaced.
                let asFile = sheet.Extras.TableParts |> List.exists (fun (name, known) -> name = table.Name && known = path)
                if not asFile then
                    let source = keptParts |> List.tryPick (fun (kept, data) -> if kept = path then Some (System.Text.Encoding.UTF8.GetString data) else None)
                    match source |> Option.bind (fun xml -> tableXmlFromSource xml table) with
                    | Some xml ->
                        textFiles.Add { Path = path; Content = xml }
                        // (The fields of its query go with the table's columns.)
                        let kept (wanted: string) = workbook.Preserved.Parts |> List.tryPick (fun (known, data) -> if known = wanted then Some (System.Text.Encoding.UTF8.GetString data) else None)
                        for query in OpcXml.relationshipsOrEmpty (kept (relsPathOf path) |> Option.defaultValue "") do
                          if OpcXml.relationshipKind query.Type = "queryTable" && not query.External then
                            let queryPath = OpcXml.targetPath (PartPath.directory path) query
                            match kept queryPath |> Option.bind (fun text -> queryTableXmlFromSource text xml) with
                            | Some updated -> textFiles.Add { Path = queryPath; Content = updated }
                            | None -> ()
                    | None ->
                        // (Not a part the writer can follow: a table of its own making under the part's name and id.)
                        let id = source |> Option.bind tableIdOf |> Option.defaultValue 1
                        textFiles.Add { Path = path; Content = tableXml id table })
    // (After the sheets: styles and strings that are added to grow as the cells are written.)
    textFiles.Add { Path = "xl/styles.xml"; Content = stylesText () }
    // The shared strings' `count` is the number of cells that name one: what the file says, less
    // the uses in the parts of the sheets that were written again, plus the uses in what was written.
    let sharedXml =
        let xml = sharedText ()
        let rootOf (text: string) = XmlPatch.readPartExcept (fun name -> name = "si") text
        let fileCount =
            workbook.Preserved.SharedStringsPart |> Option.bind partText |> Option.bind rootOf
            |> Option.bind (fun was -> XmlPatch.attribute was.Node "count")
            |> Option.map System.Int32.TryParse
        match inPlace, fileCount with
        | Some _, Some (true, was) ->
            let deltas = [ for KeyValue (_, delta) in sharedDelta -> delta ]
            match rootOf xml with
            | Some root when deltas |> List.forall Option.isSome ->
                let count = string (max 0 (was + (deltas |> List.sumBy (fun delta -> delta.Value))))
                if XmlPatch.attribute root.Node "count" = Some count then xml
                else XmlPatch.applyPart order root [ XmlPatch.SetAttribute ("count", count) ]
            | _ -> xml
        | _ -> xml
    textFiles.Add { Path = "xl/sharedStrings.xml"; Content = sharedXml }
    let written = textFiles |> Seq.map (fun file -> file.Path) |> Set.ofSeq
    let binaryFiles =
        (match workbook.VbaProjectSource with
         | Some source when source <> "" -> [ "xl/vbaProject.bin", VbaProject.writeProject "Module1" source ]
         | _ -> [])
        @ List.ofSeq mediaFiles
        // Kept parts go back as they were, unless the writer produced a part of that name.
        @ (keptParts |> List.filter (fun (path, _) -> not (Set.contains path written)))
    textFiles |> Seq.toList, binaryFiles

let writeWorkbookPackage (workbook: XlsxWorkbookOut) : Fable.Core.JS.Promise<obj> =
    let textFiles, binaryFiles = workbookFiles workbook
    createPackageWithBinary textFiles binaryFiles

/// `writeWorkbookPackage` for a workbook some of whose parts are entries of
/// the opened package (`source`, a raw archive; `reuse`: the bytes read from
/// it by path): a part that is written as those very bytes is put into the
/// new package as the compressed entry it was.
let writeWorkbookPackageReusing (source: obj) (reuse: (string * byte[]) list) (workbook: XlsxWorkbookOut) : Fable.Core.JS.Promise<obj> =
    let textFiles, binaryFiles = workbookFiles workbook
    createPackageReusing source reuse textFiles binaryFiles

#if !FABLE_COMPILER
let writeWorkbookPackageBytes (workbook: XlsxWorkbookOut) : byte[] =
    let textFiles, binaryFiles = workbookFiles workbook
    createPackageBytesWithBinary textFiles binaryFiles
#endif

let private formattedToWorkbook (sheets: FormattedSheet list) =
    let convertCell (row: int) (col: int) (cell: FormattedCell) =
        let fmt =
            { defaultXlsxCellFormat with
                Bold = cell.Bold |> Option.defaultValue false
                Italic = cell.Italic |> Option.defaultValue false
                Underline = cell.Underline |> Option.defaultValue false
                FontName = cell.FontFamily
                TextColor = cell.FontColor
                FillColor = cell.FillColor
                FontSize = cell.FontSize
                NumberFormat = cell.NumberFormat }
        let value =
            match cell.Formula with
            | Some formula ->
                let cached =
                    match tryParseNumber cell.Value with
                    | Some n -> XlsxNumber n
                    | None -> XlsxText (cell.Value, [])
                XlsxFormula (formula, Some cached)
            | None ->
                match tryParseNumber cell.Value with
                | Some n -> XlsxNumber n
                | None -> XlsxText (cell.Value, [])
        { Row = row; Col = col; Value = value; Format = fmt; Comment = None }
    { Sheets =
        sheets
        |> List.map (fun s ->
            { Name = s.Name
              State = XlsxReader.SheetVisible
              CellSource = XlsxReader.OnDemand<XlsxCellOut list>.Of (s.Rows |> List.mapi (fun r row -> row.Cells |> List.mapi (fun c cell -> convertCell r c cell)) |> List.concat)
              MergedCells = s.MergedCells
              ColumnWidths = []
              RowHeights = []
              FrozenRows = 0
              FrozenCols = 0
              ShowGridLines = true
              View = defaultSheetViewOut
              Tables = []
              DataValidations = []
              ConditionalFormats = []
              Hyperlinks = []
              PageMargins = None
              PageSetup = None
              PrintOptions = None
              HeaderFooter = None
              RowBreaks = []
              ColBreaks = []
              PrintArea = None
              PrintTitleRows = None
              PrintTitleCols = None
              Extras = XlsxReader.emptySheetExtras })
      ActiveSheetIndex = 0
      DefinedNames = []
      VbaProjectSource = None
      Dxfs = []
      Preserved = XlsxReader.emptyWorkbookPreservation }

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
        match formatCache.TryGetValue key with
        | true, format -> format
        | false, _ ->
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
      Dxfs = workbook.Dxfs |> Array.toList
      Preserved = workbook.Preserved }

/// Helper to create a simple cell from string
let simpleCell (value: string) : FormattedCell = {
    Value = value
    Bold = None
    Italic = None
    Underline = None
    FontSize = None
    FontFamily = None
    FontColor = None
    FillColor = None
    NumberFormat = None
    Formula = None
}

/// Convert hex color to RGB (removes # prefix)
let hexToRgb (color: string) : string =
    ColorText.argb color

let private appendFormattedCell (sb: StringBuilder) (rowNumber: string) (colLetters: string) (cell: FormattedCell) =
    sb.Append("<c r=\"") |> ignore
    appendCellRef sb colLetters rowNumber
    match cell.Formula with
    | Some formula ->
        sb.Append("\">") |> ignore
        appendTextElement sb "f" formula
        appendTextElement sb "v" cell.Value
        sb.Append("</c>") |> ignore
    | None ->
        match tryParseNumber cell.Value with
        | Some num ->
            sb.Append("\"><v>").Append(numberText num).Append("</v></c>") |> ignore
        | None ->
            sb.Append("\" t=\"inlineStr\"><is>") |> ignore
            appendCellTextElement sb "t" cell.Value
            sb.Append("</is></c>") |> ignore

let private appendMergeCells (sb: StringBuilder) (mergedCells: MergedCellRange list) =
    if not (List.isEmpty mergedCells) then
        sb.Append("<mergeCells count=\"").Append(List.length mergedCells).Append("\">") |> ignore
        for range in mergedCells do
            sb.Append("<mergeCell ref=\"")
                .Append(cellRef range.StartRow range.StartCol)
                .Append(':')
                .Append(cellRef range.EndRow range.EndCol)
                .Append("\"/>") |> ignore
        sb.Append("</mergeCells>") |> ignore

/// Generate cell style XML (inline for Phase 2, will be deduplicated in Phase 3)
let generateCellStyle (cell: FormattedCell) : int option * string =
    let hasStyle =
        cell.Bold.IsSome || cell.Italic.IsSome || cell.Underline.IsSome ||
        cell.FontSize.IsSome || cell.FontFamily.IsSome ||
        cell.FontColor.IsSome || cell.FillColor.IsSome

    if not hasStyle then
        (None, "")
    else
        // For Phase 2, we'll use inline styling (simplified)
        // In Phase 3, we'll implement proper style deduplication
        (None, "")

/// Generate formatted cell element
let generateFormattedCell (rowIdx: int) (colIdx: int) (cell: FormattedCell) : string =
    let ref = cellRef rowIdx colIdx

    match cell.Formula with
    | Some formula ->
        // Cell with formula
        element "c" [("r", ref)] [
            elementWithText "f" [] formula
            elementWithText "v" [] cell.Value  // Cached value
        ]
    | None ->
        match tryParseNumber cell.Value with
        | Some num ->
            // Numeric cell
            element "c" [("r", ref)] [
                elementWithText "v" [] (numberText num)
            ]
        | None ->
            // Inline string cell
            element "c" [("r", ref); ("t", "inlineStr")] [
                element "is" [] [
                    elementWithText "t" [] cell.Value
                ]
            ]

/// Generate merged cells XML
let generateMergedCells (mergedCells: MergedCellRange list) : string =
    if List.isEmpty mergedCells then
        ""
    else
        let mergeElements =
            mergedCells
            |> List.map (fun range ->
                let startRef = cellRef range.StartRow range.StartCol
                let endRef = cellRef range.EndRow range.EndCol
                let rangeRef = sprintf "%s:%s" startRef endRef
                selfClosing "mergeCell" [("ref", rangeRef)]
            )
            |> String.concat ""

        element "mergeCells" [("count", string (List.length mergedCells))] [mergeElements]

/// Generate formatted sheet XML
let generateFormattedSheetXml (sheet: FormattedSheet) : string =
    let columnLetters =
        let maxCols =
            sheet.Rows
            |> List.fold (fun maxCols row -> max maxCols row.Cells.Length) 0
        [| for col in 0 .. maxCols - 1 -> columnToLetter col |]

    let sb = StringBuilder()
    sb.Append(declaration).Append('\n') |> ignore
    sb.Append("<worksheet xmlns=\"").Append(Namespaces.spreadsheet)
        .Append("\" xmlns:r=\"").Append(Namespaces.relationships).Append("\">") |> ignore
    sb.Append("<sheetData>") |> ignore
    let mutable rowIdx = 0
    for row in sheet.Rows do
        let rowNumber = string (rowIdx + 1)
        sb.Append("<row r=\"").Append(rowNumber).Append("\">") |> ignore
        let mutable colIdx = 0
        for cell in row.Cells do
            appendFormattedCell sb rowNumber columnLetters.[colIdx] cell
            colIdx <- colIdx + 1
        sb.Append("</row>") |> ignore
        rowIdx <- rowIdx + 1
    sb.Append("</sheetData>") |> ignore
    appendMergeCells sb sheet.MergedCells
    sb.Append("</worksheet>") |> ignore
    sb.ToString()

/// Helper to create a cell with formula
let cellWithFormula (formula: string) (cachedValue: string) : FormattedCell = {
    Value = cachedValue
    Bold = None
    Italic = None
    Underline = None
    FontSize = None
    FontFamily = None
    FontColor = None
    FillColor = None
    NumberFormat = None
    Formula = Some formula
}

/// Helper to create a merged cell range
let createMergedRange (startRow: int) (startCol: int) (endRow: int) (endCol: int) : MergedCellRange = {
    StartRow = startRow
    StartCol = startCol
    EndRow = endRow
    EndCol = endCol
}

/// Helper to create a merged range from Excel-style references (e.g., "A1:B2")
let parseMergedRange (range: string) : MergedCellRange option =
    let parts = CellRef.corners range
    if parts.Length <> 2 then
        None
    else
        // This is a simplified parser - in production you'd want full Excel reference parsing
        // For now, just return None - users should use createMergedRange directly
        None

/// Create XLSX from formatted sheets
let writeFormattedXlsx (sheets: FormattedSheet list) : Fable.Core.JS.Promise<obj> =
    writeWorkbookPackage (formattedToWorkbook sheets)

/// Create minimal XLSX from 2D arrays
let writeMinimalXlsx (sheets: SheetData list) : Fable.Core.JS.Promise<obj> =
    let sheetNames = sheets |> List.map (fun s -> s.Name)

    let baseFiles = [
        { Path = "[Content_Types].xml"; Content = generateContentTypesXlsx (List.length sheets) }
        { Path = "_rels/.rels"; Content = generateRootRelsXlsx () }
        { Path = "xl/workbook.xml"; Content = generateWorkbookXml sheetNames }
        { Path = "xl/_rels/workbook.xml.rels"; Content = generateWorkbookRels (List.length sheets) }
    ]

    // Add each sheet
    let sheetFiles =
        sheets
        |> List.mapi (fun i sheet ->
            { Path = sprintf "xl/worksheets/sheet%d.xml" (i + 1); Content = generateSheetXml sheet.Data }
        )

    let files = baseFiles @ sheetFiles

    createPackage files
