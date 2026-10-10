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
    Cells: XlsxCellOut list
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
            | XlsxFormula (_, Some (XlsxText (text, runs))) -> indexOf text runs |> ignore
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

let private cachedValueXml (sharedIndex: string -> XlsxRichRun list -> int) value =
    let rec loop v =
        match v with
        | XlsxBlank -> "", ""
        | XlsxNumber n -> "", "<v>" + invariant n + "</v>"
        | XlsxBoolean b -> " t=\"b\"", (if b then "<v>1</v>" else "<v>0</v>")
        | XlsxError e -> " t=\"e\"", "<v>" + escaped e + "</v>"
        | XlsxText (text, runs) -> " t=\"s\"", "<v>" + string (sharedIndex text runs) + "</v>"
        | XlsxFormula (_, cached) ->
            cached |> Option.map loop |> Option.defaultValue ("", "")
    loop value

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
        let tAttr, cachedXml = cachedValueXml sharedIndex value
        openCell tAttr
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
        let tAttr, cachedXml = cached |> Option.map (cachedValueXml sharedIndex) |> Option.defaultValue ("", "")
        // An array formula says which cells its results fill; a dynamic one
        // names its cell metadata (`cm`).
        match (if arrayFormulas.IsEmpty then None else Map.tryFind (letters + rowText) arrayFormulas) with
        | Some (range, cmAttr) ->
            openCell (tAttr + cmAttr)
            sb.Append("><f t=\"array\" ref=\"").Append(range).Append("\">") |> ignore
        | None ->
            openCell tAttr
            sb.Append("><f>") |> ignore
        sb.Append(escapedFast formula).Append("</f>").Append(cachedXml).Append("</c>") |> ignore
    | None, other ->
        let tAttr, valueXml = cachedValueXml sharedIndex other
        openCell tAttr
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
        Set.unionMany [ widths |> Map.toSeq |> Seq.map fst |> Set.ofSeq; styleIds |> Map.toSeq |> Seq.map fst |> Set.ofSeq
                        axis |> Map.toSeq |> Seq.filter (fun (_, (_, level, collapsed)) -> level > 0 || collapsed) |> Seq.map fst |> Set.ofSeq ]
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

/// The first cell of a range reference ("B5:F10" -> "B5"), without "$".
let private firstCellOfRange (range: string) =
    (match range.IndexOf ':' with -1 -> range | index -> range.Substring(0, index)).Replace("$", "").ToUpperInvariant()

let private rowsXml (styles: StyleBook) (sharedIndex: string -> XlsxRichRun list -> int) (dynamicCm: int) (sheet: XlsxSheetOut) : string =
    // A data table is written on the first cell of its range.
    let dataTables = sheet.Extras.DataTables |> List.map (fun table -> firstCellOfRange table.Ref, table) |> Map.ofList
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
               if (level > 0 || collapsed) && not (given.Contains row) then yield row, nan |]
        |> Array.sortBy fst
    // The last height given for a row wins, as with a map built from the list.
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
             | Some (Some kept, _, _) when kept > 0.0 -> sb.Append(" ht=\"").Append(invariantFixed kept).Append("\" customHeight=\"1\"") |> ignore
             | _ -> ())
            sb.Append(" hidden=\"1\"") |> ignore
        | Some h when System.Double.IsNaN h -> ()
        | Some h -> sb.Append(" ht=\"").Append(invariantFixed h).Append("\" customHeight=\"1\"") |> ignore
        | None -> ()
        match extra with
        | Some (_, level, collapsed) ->
            if level > 0 then sb.Append(" outlineLevel=\"").Append(string level).Append('"') |> ignore
            if collapsed then sb.Append(" collapsed=\"1\"") |> ignore
        | None -> ()
        sb.Append('>') |> ignore
        while cellAt < cells.Length && cells.[cellAt].Row = row do
            appendCellXml sb styles sharedIndex dataTables arrayFormulas rowText cells.[cellAt]
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
    let extBlocks = ResizeArray<string>()
    let extBlock (sqref: string) (ruleXml: string) =
        extBlocks.Add (sprintf "<x14:conditionalFormatting xmlns:xm=\"%s\">%s<xm:sqref>%s</xm:sqref></x14:conditionalFormatting>" xmNamespace ruleXml (escaped sqref))
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
                    match rule.Ext with
                    // In the x14 list only.
                    | Some ext when ext.InExtList ->
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
                        extBlock cf.Range (sprintf "<x14:cfRule %s id=\"%s\">%s</x14:cfRule>" (String.concat " " (ruleAttrs rule priority false)) id body)
                        ""
                    | _ ->
                        let barExt = rule.Ext |> Option.bind (fun ext -> ext.DataBar)
                        let children = [
                            for f in rule.Formulas -> elemText "formula" f
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
                                    extBlock cf.Range
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
            if rules = "" then "" else sprintf "<conditionalFormatting %s>%s</conditionalFormatting>" (attr "sqref" cf.Range) rules)
        |> String.concat ""
    main, String.concat "" extBlocks

/// The sheet's extension list: the x14 conditional formats and validations
/// written from the rules, and the extensions kept from the file. x14 rules
/// join a kept x14 list (one whose rules were not all understood).
let private sheetExtLstXml (cfBlocks: string) (validationsExt: string) (kept: string list) =
    let isCfExt (xml: string) = xml.Contains "conditionalFormattings>"
    let keptCf, others = kept |> List.partition isCfExt
    let cfExt =
        match keptCf, cfBlocks with
        | [], "" -> []
        | [], blocks ->
            [ sprintf "<ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" xmlns:x14=\"%s\"><x14:conditionalFormattings>%s</x14:conditionalFormattings></ext>" x14Namespace blocks ]
        | first :: rest, blocks ->
            let close = first.LastIndexOf "</"
            let listClose = if close < 0 then -1 else first.LastIndexOf ("</", close - 1)
            if blocks = "" || listClose < 0 || not (first.Contains "xmlns:x14=") then keptCf
            else (first.Substring(0, listClose) + blocks + first.Substring listClose) :: rest
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

let private pagePartsXml (sheet: XlsxSheetOut) =
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
                yield sprintf "fitToWidth=\"%d\"" p.FitToWidth
                yield sprintf "fitToHeight=\"%d\"" p.FitToHeight
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
            sprintf "<headerFooter%s>%s</headerFooter>" (if List.isEmpty attrs then "" else " " + String.concat " " attrs) (String.concat "" children))
        |> Option.defaultValue ""
    printOptions + margins + pageSetup + headerFooter

let private breaksXml name breaks =
    if List.isEmpty breaks then ""
    else
        let body = breaks |> List.map (fun id -> sprintf "<brk id=\"%d\" max=\"16383\" man=\"1\"/>" id) |> String.concat ""
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
let private relsPathOf (path: string) =
    match path.LastIndexOf '/' with
    | -1 -> "_rels/" + path + ".rels"
    | cut -> path.Substring(0, cut) + "/_rels/" + path.Substring(cut + 1) + ".rels"

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
            | None, Some (path, relType) when relType.EndsWith "/worksheet" && path.EndsWith ".xml" && taken.Add path -> Some path
            | _ -> None)
    let free (name: int -> string) (number: int) =
        let mutable at = number
        while taken.Contains (name at) do at <- at + 1
        taken.Add (name at) |> ignore
        name at
    List.zip own verbatim
    |> List.mapi (fun i (own, verbatim) ->
        let number = i + 1
        match verbatim with
        | Some (path, _) -> { Verbatim = verbatim; SheetPath = path; CommentsPath = ""; VmlPath = ""; ThreadsPath = ""; DrawingPath = "" }
        | None ->
            { Verbatim = None
              SheetPath = (match own with Some path -> path | None -> free (sprintf "xl/worksheets/sheet%d.xml") number)
              CommentsPath = free (sprintf "xl/comments%d.xml") number
              VmlPath =
                // (The boxes of new comments: not under the name of a kept VML drawing.)
                (let usual = sprintf "xl/drawings/vmlDrawing%d.vml" number
                 if taken.Contains usual then free (sprintf "xl/drawings/commentsDrawing%d.vml") number
                 else (taken.Add usual |> ignore; usual))
              ThreadsPath = free (sprintf "xl/threadedComments/threadedComment%d.xml") number
              DrawingPath = free (sprintf "xl/drawings/sheetDrawing%d.xml") number })

let private tabSelectedPattern = System.Text.RegularExpressions.Regex("\\s+tabSelected=\"([^\"]*)\"")

/// The XML of a sheet part with its first view saying whether the sheet is
/// the one on display (`tabSelected`): two sheets that say so are a group to
/// Excel. The text itself when it already says the right thing.
let private withTabSelected (active: bool) (xml: string) : string =
    let rec viewAt (from: int) =
        let at = xml.IndexOf ("<sheetView", from, System.StringComparison.Ordinal)
        if at < 0 || at + 10 >= xml.Length then -1
        elif xml.[at + 10] = 's' then viewAt (at + 10)
        else at
    let at = viewAt 0
    let tagEnd = if at < 0 then -1 else xml.IndexOf ('>', at)
    if tagEnd < 0 then xml
    else
        let tag = xml.Substring(at, tagEnd - at + 1)
        let found = tabSelectedPattern.Match tag
        let selected = found.Success && (found.Groups.[1].Value = "1" || found.Groups.[1].Value = "true")
        if selected = active then xml
        else
            let bare = if found.Success then tag.Remove(found.Index, found.Length) else tag
            let tag = if active then "<sheetView tabSelected=\"1\"" + bare.Substring 10 else bare
            xml.Substring(0, at) + tag + xml.Substring(tagEnd + 1)

let private tableIdPattern = System.Text.RegularExpressions.Regex("<table\\b[^>]*?\\sid=\"(\\d+)\"")
let private tableNumberPattern = System.Text.RegularExpressions.Regex("table(\\d+)\\.xml$")

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
                let number = tableNumberPattern.Match path
                if number.Success then taken.Add (int number.Groups.[1].Value) |> ignore
                let id = tableIdPattern.Match (System.Text.Encoding.UTF8.GetString parts.[path])
                if id.Success then taken.Add (int id.Groups.[1].Value) |> ignore
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

let private worksheetXml (styles: StyleBook) sharedIndex (dynamicCm: int) sheetIndex active (hasDrawingPart: bool) (sheet: XlsxSheetOut) =
    let conditionalFormats, conditionalFormatsExt = conditionalFormattingParts styles.Dxfs.Length sheetIndex sheet.ConditionalFormats
    let hasFitToPage = sheet.PageSetup |> Option.exists (fun p -> p.FitToPage)
    // (CT_SheetPr: tabColor comes before pageSetUpPr.)
    let tabColor =
        match sheet.View.TabColor with
        | Some color when color.StartsWith "#" && color.Length = 7 -> sprintf "<tabColor rgb=\"FF%s\"/>" (color.Substring(1).ToUpperInvariant())
        | Some color when color.StartsWith "theme:" ->
            (match color.Split(':') with
             | [| _; index |] -> sprintf "<tabColor theme=\"%s\"/>" (escaped index)
             | [| _; index; tint |] -> sprintf "<tabColor theme=\"%s\" tint=\"%s\"/>" (escaped index) (escaped tint)
             | _ -> "")
        | Some color when color.StartsWith "indexed:" -> sprintf "<tabColor indexed=\"%s\"/>" (escaped (color.Substring 8))
        | _ -> ""
    // (... and outlinePr between the two.)
    let outlinePr =
        if sheet.View.SummaryBelow && sheet.View.SummaryRight then ""
        else
            "<outlinePr" + (if sheet.View.SummaryBelow then "" else " summaryBelow=\"0\"") + (if sheet.View.SummaryRight then "" else " summaryRight=\"0\"") + "/>"
    let keptElement (name: string) =
        sheet.Extras.KeptElements |> List.tryPick (fun (known, xml) -> if known = name then Some xml else None) |> Option.defaultValue ""
    let sheetPr =
        if hasFitToPage || tabColor <> "" || outlinePr <> "" then
            "<sheetPr>" + tabColor + outlinePr + (if hasFitToPage then "<pageSetUpPr fitToPage=\"1\"/>" else "") + "</sheetPr>"
        else ""
    declaration + "\n"
    + sprintf "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">" Namespaces.spreadsheet Namespaces.relationships
    + sheetPr
    + sprintf "<dimension ref=\"%s\"/>" (sheetDimension sheet)
    + sheetViewsXml sheet active
    + sheetFormatXml sheet
    + colsXml styles sheet
    + ("<sheetData>" + rowsXml styles sharedIndex dynamicCm sheet + "</sheetData>")
    // Child order is fixed by CT_Worksheet; Excel refuses a part that
    // deviates (e.g. tableParts before legacyDrawing).
    + autoFilterXml styles.Dxfs.Length sheet
    + mergeCellsXmlFull sheet.MergedCells
    + conditionalFormats
    + dataValidationsXml sheet.DataValidations
    + hyperlinksXml sheet.Hyperlinks
    + pagePartsXml sheet
    + breaksXml "rowBreaks" sheet.RowBreaks
    + breaksXml "colBreaks" sheet.ColBreaks
    + keptElement "customProperties"
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

let private commentsXml (cells: XlsxCellOut list) =
    let comments = cells |> List.choose (fun c -> c.Comment |> Option.map (fun comment -> c, comment))
    if List.isEmpty comments then None
    else
        let authors =
            comments
            |> List.choose (fun (_, (_, author)) -> author)
            |> fun xs -> if List.isEmpty xs then [ "" ] else xs
            |> List.distinct
        let authorIndex author =
            match author with
            | Some a -> authors |> List.tryFindIndex ((=) a) |> Option.defaultValue 0
            | None -> 0
        let authorXml = authors |> List.map (elemText "author") |> String.concat ""
        let commentXml =
            comments
            |> List.map (fun (cell, (text, author)) ->
                sprintf "<comment ref=\"%s\" authorId=\"%d\"><text><t xml:space=\"preserve\">%s</t></text></comment>"
                    (cellRef cell.Row cell.Col)
                    (authorIndex author)
                    (escapedCellText text))
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

let private vmlShapePattern = System.Text.RegularExpressions.Regex("<v:shape\\b[\\s\\S]*?</v:shape>")
let private vmlRowPattern = System.Text.RegularExpressions.Regex("<x:Row>\\s*(\\d+)\\s*</x:Row>")
let private vmlColumnPattern = System.Text.RegularExpressions.Regex("<x:Column>\\s*(\\d+)\\s*</x:Column>")
let private vmlShapeIdPattern = System.Text.RegularExpressions.Regex("_x0000_s(\\d+)")
let private vmlIdMapPattern = System.Text.RegularExpressions.Regex("<o:idmap\\b[^>]*\\bdata=\"(\\d+)")

/// A VML drawing of the file (form controls, camera pictures, the boxes of
/// its comments) with the boxes of the sheet's comments as they are now:
/// the box of a comment that is gone is taken out, a comment without a box
/// gets one, every other shape stays as it is. None when the drawing already
/// has exactly the boxes of the comments.
let vmlWithCommentBoxes (vml: string) (commentCells: (int * int) list) : string option =
    let wanted = Set.ofList commentCells
    let shapes = vmlShapePattern.Matches vml |> Seq.cast<System.Text.RegularExpressions.Match> |> Seq.toList
    let noteCell (shape: System.Text.RegularExpressions.Match) : (int * int) option =
        if shape.Value.IndexOf ("ObjectType=\"Note\"", System.StringComparison.Ordinal) < 0 then None
        else
            let row = vmlRowPattern.Match shape.Value
            let column = vmlColumnPattern.Match shape.Value
            if row.Success && column.Success then Some (int row.Groups.[1].Value, int column.Groups.[1].Value) else Some (-1, -1)
    let notes = shapes |> List.choose (fun shape -> noteCell shape |> Option.map (fun cell -> shape, cell))
    let boxed = notes |> List.map snd
    if Set.ofList boxed = wanted && boxed.Length = wanted.Count then None
    else
        // The boxes that stay: the first of each comment's cell.
        let kept = System.Collections.Generic.HashSet<int * int>()
        let sb = StringBuilder()
        let mutable at = 0
        for shape, cell in notes do
            if not (Set.contains cell wanted && kept.Add cell) then
                sb.Append(vml.Substring(at, shape.Index - at)) |> ignore
                at <- shape.Index + shape.Length
        sb.Append(vml.Substring at) |> ignore
        let body = sb.ToString()
        let missing = commentCells |> List.distinct |> List.filter (fun cell -> not (kept.Contains cell))
        if missing.IsEmpty then Some body
        else
            // Shape ids of a drawing are those of its block of 1024 (`<o:idmap data>`).
            let highest =
                vmlShapeIdPattern.Matches body |> Seq.cast<System.Text.RegularExpressions.Match>
                |> Seq.fold (fun top found -> max top (int found.Groups.[1].Value)) 0
            let first =
                if highest > 0 then highest + 1
                else
                    let block = vmlIdMapPattern.Match body
                    (if block.Success then int block.Groups.[1].Value else 1) * 1024 + 1
            let shapeType =
                if body.IndexOf ("id=\"_x0000_t202\"", System.StringComparison.Ordinal) >= 0 then ""
                else "<v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\"><v:stroke joinstyle=\"miter\"/><v:path gradientshapeok=\"t\" o:connecttype=\"rect\"/></v:shapetype>"
            let added =
                missing
                |> List.mapi (fun i (row, col) ->
                    sprintf "<v:shape id=\"_x0000_s%d\" type=\"#_x0000_t202\" style=\"position:absolute;margin-left:80pt;margin-top:5pt;width:108pt;height:59.25pt;z-index:%d;visibility:hidden\" fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/><v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\"></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>%d, 15, %d, 2, %d, 15, %d, 16</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>%d</x:Row><x:Column>%d</x:Column></x:ClientData></v:shape>"
                        (first + i) (shapes.Length + i + 1) (col + 1) (max 0 (row - 1)) (col + 3) (row + 3) row col)
                |> String.concat ""
            let close = body.LastIndexOf ("</xml>", System.StringComparison.Ordinal)
            if close < 0 then None else Some (body.Substring(0, close) + shapeType + added + body.Substring close)

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

let private tableRootPattern = System.Text.RegularExpressions.Regex("<table(?=[\\s>])[^>]*>")
let private tableColumnsPattern = System.Text.RegularExpressions.Regex("<tableColumns(?=[\\s>])[^>]*>([\\s\\S]*?)</tableColumns>")
let private tableColumnPattern = System.Text.RegularExpressions.Regex("<tableColumn(?=[\\s>/])[^>]*?/>|<tableColumn(?=[\\s>])[^>]*>[\\s\\S]*?</tableColumn>")
let private tableAutoFilterPattern = System.Text.RegularExpressions.Regex("<autoFilter(?=[\\s>/])[^>]*?/>|<autoFilter(?=[\\s>])[^>]*>[\\s\\S]*?</autoFilter>")
let private tableSortStatePattern = System.Text.RegularExpressions.Regex("<sortState(?=[\\s>/])[^>]*?/>|<sortState(?=[\\s>])[^>]*>[\\s\\S]*?</sortState>")
let private tableStyleInfoPattern = System.Text.RegularExpressions.Regex("<tableStyleInfo(?=[\\s>/])[^>]*?/>|<tableStyleInfo(?=[\\s>])[^>]*>[\\s\\S]*?</tableStyleInfo>")
let private totalsFormulaPattern = System.Text.RegularExpressions.Regex("<totalsRowFormula(?=[\\s>/])[^>]*?/>|<totalsRowFormula(?=[\\s>])[^>]*>[\\s\\S]*?</totalsRowFormula>")

/// The value of an attribute of a start tag (as the file wrote it, not decoded).
let private tagAttribute (tag: string) (name: string) : string option =
    let found = System.Text.RegularExpressions.Regex.Match(tag, "\\s" + name + "=\"([^\"]*)\"")
    if found.Success then Some found.Groups.[1].Value else None

/// A start tag (`<x ...>` or `<x .../>`) with an attribute set to a value
/// (given escaped), or without it.
let private withTagAttribute (name: string) (value: string option) (tag: string) : string =
    let pattern = System.Text.RegularExpressions.Regex("\\s" + name + "=\"[^\"]*\"")
    let found = pattern.Match tag
    match value with
    | Some value ->
        let text = " " + name + "=\"" + value + "\""
        if found.Success then tag.Substring(0, found.Index) + text + tag.Substring(found.Index + found.Length)
        else
            let cut = if tag.EndsWith "/>" then tag.Length - 2 else tag.Length - 1
            tag.Substring(0, cut).TrimEnd() + text + tag.Substring cut
    | None -> if found.Success then tag.Substring(0, found.Index) + tag.Substring(found.Index + found.Length) else tag

/// The start tag of an element's XML (up to its first `>`), and the rest.
let private splitStartTag (xml: string) : string * string =
    let close = xml.IndexOf '>'
    xml.Substring(0, close + 1), xml.Substring(close + 1)

/// The XML of a table part of the file for the table as it is now: its
/// name, range, header and totals rows, filter, columns and style are the
/// model's; everything else (the id, revision ids, dxf ids, each kept
/// column's own id, format and calculated-column formula, the filter's
/// criteria, the query-table attributes) stays as the file has it. Columns
/// are matched by position while their number is the file's, by name
/// otherwise. None when the part is not a table the writer can follow.
let tableXmlFromSource (source: string) (table: XlsxTableOut) : string option =
    let root = tableRootPattern.Match source
    let columns = tableColumnsPattern.Match source
    if not root.Success || not columns.Success || columns.Index < root.Index then None
    else
        let refText = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1) (table.EndCol - 1))
        let escapedValue (text: string) =
            let sb = StringBuilder()
            appendEscapedAttr sb text
            sb.ToString()
        let rootTag = root.Value
        let refChanged = tagAttribute rootTag "ref" <> Some refText
        let hadTotals = tagAttribute rootTag "totalsRowCount" |> Option.exists (fun count -> count <> "0")
        let newRoot =
            rootTag
            |> withTagAttribute "name" (Some (escapedValue table.Name))
            |> withTagAttribute "displayName" (Some (escapedValue table.Name))
            |> withTagAttribute "ref" (Some refText)
            |> withTagAttribute "headerRowCount" (if table.HeaderRowCount = 1 then None else Some (string table.HeaderRowCount))
            |> withTagAttribute "totalsRowCount" (if table.TotalsRowCount > 0 then Some (string table.TotalsRowCount) else None)
            |> fun tag ->
                if table.TotalsRowCount > 0 then withTagAttribute "totalsRowShown" None tag
                elif hadTotals then withTagAttribute "totalsRowShown" (Some "0") tag
                else tag
        // The columns.
        let existing =
            tableColumnPattern.Matches columns.Groups.[1].Value |> Seq.cast<System.Text.RegularExpressions.Match> |> Seq.map (fun found -> found.Value) |> Seq.toArray
        let count = table.EndCol - table.StartCol + 1
        let names = [| for i in 0 .. count - 1 -> table.Columns |> List.tryItem i |> Option.defaultValue (sprintf "Column%d" (i + 1)) |]
        let nameOf (xml: string) = tagAttribute (fst (splitStartTag xml)) "name" |> Option.defaultValue ""
        let used = Array.create existing.Length false
        let sourceOf (i: int) : string option =
            if existing.Length = count then (used.[i] <- true; Some existing.[i])
            else
                let wanted = escapedValue names.[i]
                let mutable found = -1
                for k in 0 .. existing.Length - 1 do
                    if found < 0 && not used.[k] && System.String.Equals(nameOf existing.[k], wanted, System.StringComparison.OrdinalIgnoreCase) then found <- k
                if found < 0 then None else (used.[found] <- true; Some existing.[found])
        let matched = [| for i in 0 .. count - 1 -> sourceOf i |]
        let mutable nextId =
            existing |> Array.fold (fun top xml -> match tagAttribute (fst (splitStartTag xml)) "id" |> Option.map System.Int32.TryParse with Some (true, id) -> max top id | _ -> top) 0
        let columnXml (i: int) =
            let total = if table.TotalsRowCount > 0 then List.tryItem i table.Totals else None
            let xml =
                match matched.[i] with
                | Some xml -> xml
                | None ->
                    nextId <- nextId + 1
                    sprintf "<tableColumn id=\"%d\" name=\"\"/>" nextId
            let startTag, rest = splitStartTag xml
            let startTag = withTagAttribute "name" (Some (escapedValue names.[i])) startTag
            match total with
            | None -> startTag + rest          // (No totals row: what the file says the column would show stays.)
            | Some (kind, value) ->
                let startTag =
                    startTag
                    |> withTagAttribute "totalsRowLabel" (if kind = "label" then Some (escapedValue value) else None)
                    |> withTagAttribute "totalsRowFunction" (match kind with "function" -> Some (escapedValue value) | "custom" -> Some "custom" | _ -> None)
                let rest = totalsFormulaPattern.Replace(rest, "")
                if kind <> "custom" then startTag + rest
                else
                    let formula = "<totalsRowFormula>" + escapeXml value + "</totalsRowFormula>"
                    if startTag.EndsWith "/>" then startTag.Substring(0, startTag.Length - 2).TrimEnd() + ">" + formula + "</tableColumn>"
                    else
                        // (After the calculated-column formula, before anything else: CT_TableColumn's order.)
                        let close = "</calculatedColumnFormula>"
                        let at = rest.IndexOf (close, System.StringComparison.Ordinal)
                        if at >= 0 then startTag + rest.Substring(0, at + close.Length) + formula + rest.Substring(at + close.Length)
                        else startTag + formula + rest
        let columnsTag, _ = splitStartTag columns.Value
        let newColumns =
            withTagAttribute "count" (Some (string count)) columnsTag
            + (String.concat "" [ for i in 0 .. count - 1 -> columnXml i ])
            + "</tableColumns>"
        let between = source.Substring(root.Index + root.Length, columns.Index - root.Index - root.Length)
        let after = source.Substring(columns.Index + columns.Length)
        // The filter (it covers the header and the data, not the totals row).
        let filterRef = sprintf "%s:%s" (cellRef (table.StartRow - 1) (table.StartCol - 1)) (cellRef (table.EndRow - 1 - max 0 table.TotalsRowCount) (table.EndCol - 1))
        let between =
            let found = tableAutoFilterPattern.Match between
            match table.AutoFilter, found.Success with
            | None, true -> between.Substring(0, found.Index) + between.Substring(found.Index + found.Length)
            | None, false -> between
            | Some _, true ->
                let startTag, rest = splitStartTag found.Value
                between.Substring(0, found.Index) + withTagAttribute "ref" (Some filterRef) startTag + rest + between.Substring(found.Index + found.Length)
            | Some _, false -> sprintf "<autoFilter ref=\"%s\"/>" filterRef + between
        // A sort order names cells of the range it was made for.
        let between = if refChanged then tableSortStatePattern.Replace(between, "") else between
        let style =
            table.Style
            |> Option.map (fun s ->
                sprintf "<tableStyleInfo %s showFirstColumn=\"%d\" showLastColumn=\"%d\" showRowStripes=\"%d\" showColumnStripes=\"%d\"/>"
                    (attr "name" s.Name) (if s.ShowFirstColumn then 1 else 0) (if s.ShowLastColumn then 1 else 0) (if s.ShowRowStripes then 1 else 0) (if s.ShowColumnStripes then 1 else 0))
            |> Option.defaultValue ""
        let after =
            let found = tableStyleInfoPattern.Match after
            if found.Success then after.Substring(0, found.Index) + style + after.Substring(found.Index + found.Length) else style + after
        Some (source.Substring(0, root.Index) + newRoot + between + newColumns + after)

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
let private sheetRelationshipIds (workbook: XlsxWorkbookOut) : string list =
    let taken = workbook.Preserved.WorkbookRelationships @ Option.toList workbook.Preserved.CalcChain |> List.map (fun (id, _, _, _) -> id) |> Set.ofList
    workbook.Sheets
    |> List.mapi (fun i _ ->
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
              | Some v -> yield { Name = "_xlnm.Print_Area"; RefersTo = sprintf "%s!%s" (quoted s.Name) v; LocalSheetIndex = Some i; Comment = None; Hidden = false }
              | None -> ()
              match s.PrintTitleRows, s.PrintTitleCols with
              | None, None -> ()
              | rows, cols ->
                  let parts = [ rows; cols ] |> List.choose id |> List.map (fun part -> sprintf "%s!%s" (quoted s.Name) part)
                  yield { Name = "_xlnm.Print_Titles"; RefersTo = String.concat "," parts; LocalSheetIndex = Some i; Comment = None; Hidden = false }
              // Excel keeps the range of a sheet's AutoFilter in a hidden name of the sheet.
              match s.Extras.AutoFilter with
              | Some filter ->
                  let absolute (corner: string) =
                      let letters = corner |> Seq.takeWhile (fun c -> c < '0' || c > '9') |> Seq.filter (fun c -> c <> '$') |> Seq.toArray |> System.String
                      let digits = corner |> Seq.filter (fun c -> c >= '0' && c <= '9') |> Seq.toArray |> System.String
                      (if letters = "" then "" else "$" + letters) + (if digits = "" then "" else "$" + digits)
                  let reference = filter.Ref.Split(':') |> Array.map absolute |> String.concat ":"
                  yield { Name = "_xlnm._FilterDatabase"; RefersTo = sprintf "%s!%s" (quoted s.Name) reference; LocalSheetIndex = Some i; Comment = None; Hidden = true }
              | None -> () ])
        |> List.concat
    let names = workbook.DefinedNames @ printNames
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
                        (escaped n.RefersTo))
                |> String.concat ""
            "<definedNames>" + body + "</definedNames>"
    // A formula written without a result has to be calculated when the file is opened.
    let hasUncachedFormula =
        workbook.Sheets
        |> List.exists (fun s -> s.Cells |> List.exists (fun c -> match c.Value with XlsxFormula (_, None) -> true | _ -> false))
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
    sprintf "%s\n<workbook xmlns=\"%s\" xmlns:r=\"%s\"><bookViews><workbookView activeTab=\"%d\"/></bookViews><sheets>%s</sheets>%s%s%s%s%s</workbook>"
        declaration Namespaces.spreadsheet Namespaces.relationships (max 0 workbook.ActiveSheetIndex) sheets
        kept.ExternalReferencesXml definedNames calcPr kept.PivotCachesXml kept.WorkbookExtLstXml

let private workbookRelsXmlFull (workbook: XlsxWorkbookOut) (placements: SheetPlacement list) hasVba (themePath: string option) (hasPersons: bool) (ownMetadata: bool) =
    let sheetRels =
        List.zip (sheetRelationshipIds workbook) placements
        |> List.map (fun (id, placement) ->
            let relType = match placement.Verbatim with Some (_, relType) when relType <> "" -> relType | _ -> RelationshipTypes.worksheet
            sprintf "<Relationship Id=\"%s\" %s Target=\"/%s\"/>" id (attr "Type" relType) (escaped placement.SheetPath))
    let baseRels = [
        yield! sheetRels
        yield sprintf "<Relationship Id=\"rIdStyles\" Type=\"%s\" Target=\"styles.xml\"/>" RelationshipTypes.styles
        yield sprintf "<Relationship Id=\"rIdSharedStrings\" Type=\"%s\" Target=\"sharedStrings.xml\"/>" RelationshipTypes.sharedStrings
        if hasVba then yield "<Relationship Id=\"rIdVBA\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>"
        // The theme of the opened workbook, kept as it was.
        match themePath with
        | Some path -> yield sprintf "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"/%s\"/>" (escaped path)
        | None -> ()
        // The cell metadata dynamic array formulas name, when the opened workbook had none.
        if ownMetadata then yield "<Relationship Id=\"rIdMetadata\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata\" Target=\"metadata.xml\"/>"
        if hasPersons then yield "<Relationship Id=\"rIdPersons\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"/xl/persons/person.xml\"/>"
        // Pivot caches, slicer caches and external links, under the ids workbook.xml names them by.
        for relId, relType, target, isExternal in workbook.Preserved.WorkbookRelationships @ Option.toList workbook.Preserved.CalcChain do
            yield sprintf "<Relationship %s %s %s%s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" target) (if isExternal then " TargetMode=\"External\"" else "")
    ]
    sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" baseRels)

let private contentTypesXmlFull (workbook: XlsxWorkbookOut) (placements: SheetPlacement list) (keptParts: (string * byte[]) list) (ownMetadata: bool) =
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
                        if relType.EndsWith "/chartsheet" then "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"
                        elif relType.EndsWith "/dialogsheet" then "application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"
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
            if s.Cells |> List.exists (fun c -> c.Comment.IsSome) then
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
    + sheetOverrides + tableOverrides + comments + extraOverrides
    + "</Types>"

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

let private relationshipTargetPattern = System.Text.RegularExpressions.Regex("<Relationship\\b[^>]*>")

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
        if not isExternal && not (target.StartsWith "#") then pending.Enqueue (ZipInterop.normalizePath target)
    for sheet in workbook.Sheets do
        match verbatimPart workbook sheet with
        | Some (path, _) -> want path false
        | None -> ()
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
    let directoryOf (path: string) = match path.LastIndexOf '/' with -1 -> "" | at -> path.Substring(0, at)
    while pending.Count > 0 do
        let path = pending.Dequeue()
        if parts.ContainsKey path && seen.Add path then
            let directory = directoryOf path
            let relsPath = (if directory = "" then "" else directory + "/") + "_rels/" + path.Substring(if directory = "" then 0 else directory.Length + 1) + ".rels"
            match parts.TryGetValue relsPath with
            | true, relsData ->
                seen.Add relsPath |> ignore
                let relsText = System.Text.Encoding.UTF8.GetString relsData
                for m in relationshipTargetPattern.Matches relsText |> Seq.cast<System.Text.RegularExpressions.Match> do
                    let tag = m.Value
                    let attribute (name: string) =
                        let key = " " + name + "=\""
                        let at = tag.IndexOf key
                        if at < 0 then None
                        else
                            let start = at + key.Length
                            let close = tag.IndexOf ('"', start)
                            if close < 0 then None else Some (tag.Substring(start, close - start))
                    match attribute "Target" with
                    | Some target when attribute "TargetMode" <> Some "External" ->
                        pending.Enqueue (ZipInterop.resolveTarget directory (target.Replace("&amp;", "&")))
                    | _ -> ()
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
        let countOf (text: string) (token: string) =
            let mutable count = 0
            let mutable at = text.IndexOf(token, System.StringComparison.Ordinal)
            while at >= 0 do
                count <- count + 1
                at <- text.IndexOf(token, at + token.Length, System.StringComparison.Ordinal)
            count
        // An element that closes itself holds nothing: it is written again in full.
        let withoutEmpty (name: string) (text: string) =
            System.Text.RegularExpressions.Regex.Replace(text, "<" + name + "\\b[^>]*/>", "")
        let withCount (openTag: string) (count: int) =
            let pattern = System.Text.RegularExpressions.Regex("\\bcount=\"\\d*\"")
            if pattern.IsMatch openTag then pattern.Replace(openTag, "count=\"" + string count + "\"")
            else openTag.Substring(0, openTag.Length - 1) + " count=\"" + string count + "\">"
        /// An element found by the start of its tag: (tag start, end of the start tag, start of the end tag).
        let elementAt (text: string) (start: string) (name: string) =
            let at = text.IndexOf(start, System.StringComparison.Ordinal)
            if at < 0 then None
            else
                let tagEnd = text.IndexOf('>', at)
                let close = if tagEnd < 0 then -1 else text.IndexOf("</" + name + ">", tagEnd, System.StringComparison.Ordinal)
                if close < 0 then None else Some (at, tagEnd, close)
        /// Adds `item` to the element as its last child, with the count of its `child`ren kept right.
        let appended (text: string) (at: int, tagEnd: int, close: int) (child: string) (item: string) =
            let had = countOf (text.Substring(tagEnd + 1, close - tagEnd - 1)) child
            text.Substring(0, at) + withCount (text.Substring(at, tagEnd - at + 1)) (had + 1)
            + text.Substring(tagEnd + 1, close - tagEnd - 1) + item + text.Substring close,
            had
        let text = existing |> withoutEmpty "metadataTypes" |> withoutEmpty "cellMetadata"
        let rootAt = text.IndexOf("<metadata", System.StringComparison.Ordinal)
        let rootEnd = if rootAt < 0 then -1 else text.IndexOf('>', rootAt)
        let rootClose = text.LastIndexOf("</metadata>", System.StringComparison.Ordinal)
        if rootEnd < 0 || rootClose < rootEnd then dynamicArrayMetadataXml, 1
        else
            // 1. The metadata type.
            let text, typeIndex =
                match elementAt text "<metadataTypes" "metadataTypes" with
                | Some (at, tagEnd, close) ->
                    let body = text.Substring(tagEnd + 1, close - tagEnd - 1)
                    let named = body.IndexOf("name=\"XLDAPR\"", System.StringComparison.Ordinal)
                    if named >= 0 then text, countOf (body.Substring(0, named)) "<metadataType "
                    else
                        let extended, had = appended text (at, tagEnd, close) "<metadataType " xldaprTypeXml
                        extended, had + 1
                | None ->
                    text.Substring(0, rootEnd + 1) + "<metadataTypes count=\"1\">" + xldaprTypeXml + "</metadataTypes>" + text.Substring(rootEnd + 1), 1
            // Where the elements that come after the future metadata begin:
            // cellMetadata, valueMetadata, the part's own extLst.
            let tailAt (text: string) (names: string list) =
                let last =
                    [ "</metadataTypes>"; "</metadataStrings>"; "</mdxMetadata>"; "</futureMetadata>"; "</cellMetadata>"; "</valueMetadata>" ]
                    |> List.map (fun close -> text.LastIndexOf(close, System.StringComparison.Ordinal))
                    |> List.max
                let candidates =
                    [ for name in names do
                        let at = text.IndexOf("<" + name, System.StringComparison.Ordinal)
                        if at >= 0 then yield at
                      let ext = text.IndexOf("<extLst", max 0 last, System.StringComparison.Ordinal)
                      if ext >= 0 then yield ext
                      yield text.LastIndexOf("</metadata>", System.StringComparison.Ordinal) ]
                List.min candidates
            // 2. The future metadata block that says "dynamic".
            let text, blockIndex =
                match elementAt text "<futureMetadata name=\"XLDAPR\"" "futureMetadata" with
                | Some found -> appended text found "<bk>" (dynamicArrayBlockXml true)
                | None ->
                    let at = tailAt text [ "cellMetadata"; "valueMetadata" ]
                    text.Substring(0, at) + "<futureMetadata name=\"XLDAPR\" count=\"1\">" + dynamicArrayBlockXml true + "</futureMetadata>" + text.Substring at, 0
            // 3. The cell-metadata record that points at it.
            let record = sprintf "<bk><rc t=\"%d\" v=\"%d\"/></bk>" typeIndex blockIndex
            match elementAt text "<cellMetadata" "cellMetadata" with
            | Some found ->
                let extended, had = appended text found "<bk>" record
                extended, had + 1
            | None ->
                let at = tailAt text [ "valueMetadata" ]
                text.Substring(0, at) + "<cellMetadata count=\"1\">" + record + "</cellMetadata>" + text.Substring at, 1

// ---------------------------------------------------------------------------
// Styles and shared strings extended in place (WorkbookPreservation.InPlace)
// ---------------------------------------------------------------------------

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

let private regexCount (pattern: string) (text: string) =
    System.Text.RegularExpressions.Regex.Matches(text, pattern).Count

/// The first element `name` of the text that has an end tag: where its start
/// tag begins and ends and where its end tag begins.
let private elementSpan (xml: string) (name: string) : (int * int * int) option =
    let found = System.Text.RegularExpressions.Regex.Match(xml, "<" + name + "(?=[\\s>])[^>]*>")
    if not found.Success then None
    else
        let close = xml.IndexOf ("</" + name + ">", found.Index + found.Length, System.StringComparison.Ordinal)
        if close < 0 then None else Some (found.Index, found.Index + found.Length - 1, close)

/// The element `name` with `items` added behind its children and its `count` set.
let private withAppendedChildren (name: string) (childPattern: string) (items: string list) (xml: string) : string =
    match (if items.IsEmpty then None else elementSpan xml name) with
    | None -> xml
    | Some (at, tagEnd, close) ->
        let body = xml.Substring(tagEnd + 1, close - tagEnd - 1)
        let total = regexCount childPattern body + items.Length
        xml.Substring(0, at) + withTagAttribute "count" (Some (string total)) (xml.Substring(at, tagEnd - at + 1))
        + body + String.concat "" items + xml.Substring close

let private xmlAttributeDecoded (text: string) =
    if text.IndexOf '&' < 0 then text
    else text.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&")

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
    let selfClosed (name: string) (text: string) =
        System.Text.RegularExpressions.Regex.Replace(text, "<" + name + "((?:\\s[^>]*)?)/>", "<" + name + "$1></" + name + ">")
    let xml = [ "numFmts"; "fonts"; "fills"; "borders"; "cellXfs"; "dxfs" ] |> List.fold (fun text name -> selfClosed name text) stylesXml
    let bodyOf (name: string) = elementSpan xml name |> Option.map (fun (_, tagEnd, close) -> xml.Substring(tagEnd + 1, close - tagEnd - 1))
    let fontPattern, fillPattern, borderPattern, xfPattern, dxfPattern, numFmtPattern =
        "<font[\\s>/]", "<fill[\\s>/]", "<border[\\s>/]", "<xf[\\s>/]", "<dxf[\\s>/]", "<numFmt[\\s>/]"
    match bodyOf "fonts", bodyOf "fills", bodyOf "borders", bodyOf "cellXfs", elementSpan xml "styleSheet" with
    | Some fontsBody, Some fillsBody, Some bordersBody, Some xfsBody, Some _ ->
        let fontCount, fillCount, borderCount, xfCount = regexCount fontPattern fontsBody, regexCount fillPattern fillsBody, regexCount borderPattern bordersBody, regexCount xfPattern xfsBody
        let dxfCount = bodyOf "dxfs" |> Option.map (regexCount dxfPattern) |> Option.defaultValue 0
        if xfCount = 0 || xfCount <> kept.CellFormats.Length || dxfCount <> kept.DxfCount || workbook.Dxfs.Length < dxfCount || fontCount = 0 || fillCount < 2 || borderCount = 0 then None
        else
            // The xfId of each named style, as the reader numbers them.
            let styleXfCount = bodyOf "cellStyleXfs" |> Option.map (regexCount xfPattern) |> Option.defaultValue 0
            let namedXfIds =
                bodyOf "cellStyles" |> Option.defaultValue ""
                |> fun body -> System.Text.RegularExpressions.Regex.Matches(body, "<cellStyle(?=[\\s>/])[^>]*>") |> Seq.cast<System.Text.RegularExpressions.Match>
                |> Seq.choose (fun found ->
                    match tagAttribute found.Value "name", tagAttribute found.Value "xfId" |> Option.map System.Int32.TryParse with
                    | Some _, Some (true, xfId) when xfId >= 0 && xfId < styleXfCount -> Some xfId
                    | _ -> None)
                |> Seq.distinct |> Seq.sort |> Seq.toArray
            // The custom number formats of the file.
            let numFmtIds = System.Collections.Generic.Dictionary<string, int>()
            let mutable nextNumFmt = 164
            for found in System.Text.RegularExpressions.Regex.Matches(bodyOf "numFmts" |> Option.defaultValue "", "<numFmt(?=[\\s>/])[^>]*>") |> Seq.cast<System.Text.RegularExpressions.Match> do
                match tagAttribute found.Value "numFmtId" |> Option.map System.Int32.TryParse, tagAttribute found.Value "formatCode" with
                | Some (true, id), Some code ->
                    nextNumFmt <- max nextNumFmt (id + 1)
                    let code = xmlAttributeDecoded code
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
            let indexOfKey (key: XfKey) =
                match xfIndex.TryGetValue key with
                | true, index -> index
                | false, _ ->
                    let numFmtId =
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
                    let fontId =
                        if key.Font = defaultFontKey then 0
                        else
                            match fontIds.TryGetValue key.Font with
                            | true, id -> id
                            | false, _ ->
                                let id = fontCount + newFonts.Count
                                fontIds.[key.Font] <- id
                                newFonts.Add (fontXml key.Font)
                                id
                    let fillId =
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
                    let borderId =
                        if key.Border = defaultBorderKey then 0
                        else
                            match borderIds.TryGetValue key.Border with
                            | true, id -> id
                            | false, _ ->
                                let id = borderCount + newBorders.Count
                                borderIds.[key.Border] <- id
                                newBorders.Add (borderXml key.Border)
                                id
                    let xfId = if key.NamedStyle > 0 && key.NamedStyle < namedXfIds.Length then namedXfIds.[key.NamedStyle] else 0
                    let index = xfCount + newXfs.Count
                    newXfs.Add (xfXmlOf numFmtId fontId fillId borderId xfId key)
                    xfIndex.[key] <- index
                    index
            let byIdentity = System.Collections.Generic.Dictionary<XlsxCellFormat, int>(HashIdentity.Reference)
            let styleIdOf (format: XlsxCellFormat) =
                match byIdentity.TryGetValue format with
                | true, index -> index
                | false, _ ->
                    let index = indexOfKey (xfKey format)
                    byIdentity.[format] <- index
                    index
            let finish () =
                let newDxfs = workbook.Dxfs |> List.skip dxfCount |> List.map (dxfXml "dxf")
                let insertedAfter (marks: string list) (element: string) (text: string) =
                    let at = marks |> List.tryPick (fun mark -> match text.IndexOf (mark, System.StringComparison.Ordinal) with -1 -> None | found -> Some (found + mark.Length))
                    match at with
                    | Some at -> text.Substring(0, at) + element + text.Substring at
                    | None -> text
                let text =
                    if newNumFmts.Count = 0 then xml
                    elif (elementSpan xml "numFmts").IsSome then withAppendedChildren "numFmts" numFmtPattern (List.ofSeq newNumFmts) xml
                    else
                        // (The first child of the style sheet.)
                        match elementSpan xml "styleSheet" with
                        | Some (_, tagEnd, _) -> xml.Substring(0, tagEnd + 1) + sprintf "<numFmts count=\"%d\">%s</numFmts>" newNumFmts.Count (String.concat "" newNumFmts) + xml.Substring(tagEnd + 1)
                        | None -> xml
                let text =
                    text
                    |> withAppendedChildren "fonts" fontPattern (List.ofSeq newFonts)
                    |> withAppendedChildren "fills" fillPattern (List.ofSeq newFills)
                    |> withAppendedChildren "borders" borderPattern (List.ofSeq newBorders)
                    |> withAppendedChildren "cellXfs" xfPattern (List.ofSeq newXfs)
                if newDxfs.IsEmpty then text
                elif (elementSpan text "dxfs").IsSome then withAppendedChildren "dxfs" dxfPattern newDxfs text
                else insertedAfter [ "</cellStyles>"; "</cellXfs>" ] (sprintf "<dxfs count=\"%d\">%s</dxfs>" newDxfs.Length (String.concat "" newDxfs)) text
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
    let xml =
        if sharedStringsXml.Trim() = "" then declaration + "\n<sst xmlns=\"" + Namespaces.spreadsheet + "\" count=\"0\" uniqueCount=\"0\"></sst>"
        else System.Text.RegularExpressions.Regex.Replace(sharedStringsXml, "<sst((?:\\s[^>]*)?)/>", "<sst$1></sst>")
    let root = System.Text.RegularExpressions.Regex.Match(xml, "<sst(?=[\\s>])[^>]*>")
    let close = xml.LastIndexOf ("</sst>", System.StringComparison.Ordinal)
    if not root.Success || close < root.Index || regexCount "<si[\\s>/]" xml <> texts.Length then None
    else
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
            if count = texts.Length then xml
            else
                let was = match tagAttribute root.Value "count" |> Option.map System.Int32.TryParse with Some (true, n) -> n | _ -> texts.Length
                let rootTag = root.Value |> withTagAttribute "uniqueCount" (Some (string count)) |> withTagAttribute "count" (Some (string (was + count - texts.Length)))
                xml.Substring(0, root.Index) + rootTag + xml.Substring(root.Index + root.Length, close - root.Index - root.Length) + added.ToString() + xml.Substring close
        Some (indexOf, finish)

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
    // Without them a worksheet cannot go back as it was read: it is written from its cells.
    let workbook =
        if inPlace.IsSome then workbook
        else
            { workbook with
                Sheets =
                    workbook.Sheets
                    |> List.map (fun sheet ->
                        match sheet.Extras.Verbatim, sheet.Extras.SourcePart with
                        | true, Some (_, relType) when relType.EndsWith "/worksheet" -> { sheet with Extras = { sheet.Extras with Verbatim = false } }
                        | _ -> sheet) }
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
    let keptParts = reachableKeptParts workbook themePath
    // Dynamic array formulas name a record of the cell metadata: the opened
    // workbook's metadata part with that record added to it, or a new part.
    let hasDynamicArrays =
        workbook.Sheets |> List.exists (fun sheet -> sheet.Extras.ArrayFormulas |> List.exists (fun (_, _, dynamic) -> dynamic))
    let keptMetadata =
        workbook.Preserved.WorkbookRelationships
        |> List.tryPick (fun (_, relType, target, isExternal) ->
            if isExternal || not (relType.EndsWith "/sheetMetadata") then None
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
    textFiles.Add { Path = "[Content_Types].xml"; Content = contentTypesXmlFull workbook placements keptParts ownMetadata }
    textFiles.Add { Path = "_rels/.rels"; Content = rootRelsXmlFull workbook }
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
    List.zip workbook.Sheets placements
    |> List.iteri (fun i (sheet, placed) ->
      // (A sheet written as it was read is one of the kept parts; only whether it is the one on display may differ.)
      if placed.Verbatim.IsSome then
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
        textFiles.Add { Path = placed.SheetPath; Content = worksheetXml styles sharedIndex dynamicCm sheetIndex (i = workbook.ActiveSheetIndex) drawingPath.IsSome sheet }
        let commentsVmlPath = placed.VmlPath
        match sheetRelsXml sheetIndex sheet tableParts placed drawingPath with
        | Some rels -> textFiles.Add { Path = relsPathOf placed.SheetPath; Content = rels }
        | None -> ()
        match commentsXml sheet.Cells with
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
            let ids = DocxReader.relationshipIdsInXml keptElement
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
                    | Some xml -> textFiles.Add { Path = path; Content = xml }
                    | None ->
                        // (Not a part the writer can follow: a table of its own making under the part's name and id.)
                        let id = source |> Option.map (fun xml -> tableIdPattern.Match xml) |> Option.filter (fun found -> found.Success) |> Option.map (fun found -> int found.Groups.[1].Value) |> Option.defaultValue 1
                        textFiles.Add { Path = path; Content = tableXml id table })
    // (After the sheets: styles and strings that are added to grow as the cells are written.)
    textFiles.Add { Path = "xl/styles.xml"; Content = stylesText () }
    textFiles.Add { Path = "xl/sharedStrings.xml"; Content = sharedText () }
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
              Cells = s.Rows |> List.mapi (fun r row -> row.Cells |> List.mapi (fun c cell -> convertCell r c cell)) |> List.concat
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
        | XlsxReader.Error e -> XlsxError (if e.StartsWith "#" then e else "#" + e)
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
                match namedStyleOf cell with
                | None -> styleOf cell.Style
                | named -> { styleOf cell.Style with NamedStyle = named }
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
              Cells =
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
                sheet.Rows
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
                        [ for row in sheet.Rows do
                            for cell in row.Cells do
                                match cell.Formula, cell.ArrayRef with
                                | Some _, Some range when range <> "" -> yield cell.Reference.Column + string cell.Reference.Row, range, cell.ArrayDynamic
                                | _ -> () ]
                    RowAxis =
                        sheet.Rows
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
    if color.StartsWith("#") then
        "FF" + color.Substring(1).ToUpper()  // Add alpha channel
    else
        "FF" + color.ToUpper()

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
    let parts = range.Split(':')
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
