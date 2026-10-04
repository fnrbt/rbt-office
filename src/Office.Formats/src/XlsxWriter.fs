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
    AutoFilter: (int * int) option
    Style: XlsxTableStyle option
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
}

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
    let xfs =
        workbook.Sheets
        |> List.collect (fun s ->
            (s.Cells |> List.map (fun c -> xfKey c.Format))
            // Default cell styles of columns (<col style>).
            @ (s.Extras.ColumnStyles |> List.map (fun (_, _, style) -> columnStyleKey style)))
        |> List.fold (fun acc key -> appendUnique key acc) [ defaultXfKey ]
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
      StyleIndex = styleIndex }

let private sharedTextOf value =
    match value with
    | XlsxText (text, runs) -> Some { Text = text; Runs = runs }
    | XlsxFormula (_, Some (XlsxText (text, runs))) -> Some { Text = text; Runs = runs }
    | _ -> None

let private collectSharedStrings (workbook: XlsxWorkbookOut) =
    let items =
        workbook.Sheets
        |> List.collect (fun s -> s.Cells |> List.choose (fun c -> sharedTextOf c.Value))
        |> List.distinct
    let index = items |> List.mapi (fun i key -> key, i) |> Map.ofList
    items, index

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
        let dxfXml (d: XlsxReader.Dxf) =
            let flag name value =
                match value with
                | Some true -> sprintf "<%s/>" name
                | Some false -> sprintf "<%s val=\"0\"/>" name
                | None -> ""
            let font =
                let body =
                    flag "b" d.Bold + flag "i" d.Italic
                    + (d.FontColor |> xmlColor |> Option.map (fun c -> sprintf "<color rgb=\"%s\"/>" c) |> Option.defaultValue "")
                if body = "" then "" else "<font>" + body + "</font>"
            let fill =
                d.FillColor
                |> xmlColor
                |> Option.map (fun c -> sprintf "<fill><patternFill><bgColor rgb=\"%s\"/></patternFill></fill>" c)
                |> Option.defaultValue ""
            "<dxf>" + font + fill + "</dxf>"
        sprintf "<dxfs count=\"%d\">%s</dxfs>" book.Dxfs.Length (book.Dxfs |> List.map dxfXml |> String.concat "")
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
        let fontId = indexOf xf.Font book.Fonts
        let fillId = fillIdOf xf.Fill
        let borderId = indexOf xf.Border book.Borders
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
            yield sprintf "xfId=\"%d\"" (if xf.NamedStyle > 0 && xf.NamedStyle < named.Length then xf.NamedStyle else 0)
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

let private sharedStringsXmlPreserved (items: SharedStringKey list) =
    let body =
        items
        |> List.map (fun item ->
            if List.isEmpty item.Runs then
                sprintf "<si><t xml:space=\"preserve\">%s</t></si>" (escapedCellText item.Text)
            else
                "<si>" + (item.Runs |> List.map richRunXml |> String.concat "") + "</si>")
        |> String.concat ""
    sprintf "%s\n<sst xmlns=\"%s\" count=\"%d\" uniqueCount=\"%d\">%s</sst>"
        declaration Namespaces.spreadsheet items.Length items.Length body

let private cachedValueXml sharedIndex value =
    let rec loop v =
        match v with
        | XlsxBlank -> "", ""
        | XlsxNumber n -> "", sprintf "<v>%s</v>" (invariant n)
        | XlsxBoolean b -> " t=\"b\"", sprintf "<v>%d</v>" (if b then 1 else 0)
        | XlsxError e -> " t=\"e\"", sprintf "<v>%s</v>" (escaped e)
        | XlsxText (text, runs) ->
            let key = { Text = text; Runs = runs }
            let idx = Map.find key sharedIndex
            " t=\"s\"", sprintf "<v>%d</v>" idx
        | XlsxFormula (_, cached) ->
            cached |> Option.map loop |> Option.defaultValue ("", "")
    loop value

let private cellXmlFull (styles: StyleBook) sharedIndex (cell: XlsxCellOut) =
    let reference = cellRef cell.Row cell.Col
    let xf = xfKey cell.Format
    let styleId = Map.find xf styles.StyleIndex
    let styleAttr = if styleId = 0 then "" else sprintf " s=\"%d\"" styleId
    match cell.Value with
    | XlsxBlank ->
        if styleId = 0 then ""
        else sprintf "<c r=\"%s\"%s/>" reference styleAttr
    | XlsxFormula (formula, cached) ->
        let tAttr, cachedXml = cached |> Option.map (cachedValueXml sharedIndex) |> Option.defaultValue ("", "")
        sprintf "<c r=\"%s\"%s%s><f>%s</f>%s</c>" reference styleAttr tAttr (escaped formula) cachedXml
    | other ->
        let tAttr, valueXml = cachedValueXml sharedIndex other
        sprintf "<c r=\"%s\"%s%s>%s</c>" reference styleAttr tAttr valueXml

let private sheetDimension (sheet: XlsxSheetOut) =
    let refs =
        [ for c in sheet.Cells -> c.Row, c.Col
          for m in sheet.MergedCells do
              yield m.StartRow, m.StartCol
              yield m.EndRow, m.EndCol
          for t in sheet.Tables do
              yield t.StartRow - 1, t.StartCol - 1
              yield t.EndRow - 1, t.EndCol - 1 ]
    match refs with
    | [] -> "A1"
    | _ ->
        let minR = refs |> List.map fst |> List.min
        let minC = refs |> List.map snd |> List.min
        let maxR = refs |> List.map fst |> List.max
        let maxC = refs |> List.map snd |> List.max
        sprintf "%s:%s" (cellRef minR minC) (cellRef maxR maxC)

/// `<cols>`: widths (0 = hidden), and the default cell style a column range
/// had in the file. Neighbouring columns with the same width and style are
/// written as one `<col min max>`, as Excel does.
let private colsXml (styles: StyleBook) (sheet: XlsxSheetOut) =
    let widths = Map.ofList sheet.ColumnWidths
    let styleIds =
        sheet.Extras.ColumnStyles
        |> List.collect (fun (first, last, style) ->
            match Map.tryFind (columnStyleKey style) styles.StyleIndex with
            | Some id when id > 0 -> [ for col in max 0 first .. min 16383 last -> col, id ]
            | _ -> [])
        |> Map.ofList
    // A hidden column keeps the width the file gave it.
    let fileWidth col =
        sheet.Extras.Columns
        |> List.tryFind (fun column -> col + 1 >= column.Index && col + 1 <= column.MaxIndex)
        |> Option.bind (fun column -> column.Width)
    let columns =
        Set.union (widths |> Map.toSeq |> Seq.map fst |> Set.ofSeq) (styleIds |> Map.toSeq |> Seq.map fst |> Set.ofSeq)
        |> Set.toList
    if List.isEmpty columns then ""
    else
        let attributes col =
            let width, hidden =
                match Map.tryFind col widths with
                | Some width when width <= 0.0 -> fileWidth col, true
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
        if not sheet.ShowGridLines then yield "showGridLines=\"0\""
        if active then yield "tabSelected=\"1\""
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

let private rowsXml styles sharedIndex (sheet: XlsxSheetOut) =
    let cellsByRow = sheet.Cells |> List.groupBy (fun c -> c.Row) |> Map.ofList
    let rowHeights = sheet.RowHeights |> Map.ofList
    let rowSet =
        Set.union
            (cellsByRow |> Map.toSeq |> Seq.map fst |> Set.ofSeq)
            (rowHeights |> Map.toSeq |> Seq.map fst |> Set.ofSeq)
    rowSet
    |> Set.toList
    |> List.sort
    |> List.map (fun row ->
        let rowAttrs = [
            yield sprintf "r=\"%d\"" (row + 1)
            match Map.tryFind row rowHeights with
            // Height 0 is how the sheet model says "hidden" (as for columns).
            | Some h when h <= 0.0 -> yield "hidden=\"1\""
            | Some h ->
                yield sprintf "ht=\"%s\"" (invariantFixed h)
                yield "customHeight=\"1\""
            | None -> ()
        ]
        let cells =
            Map.tryFind row cellsByRow
            |> Option.defaultValue []
            |> List.sortBy (fun c -> c.Col)
            |> List.map (cellXmlFull styles sharedIndex)
            |> String.concat ""
        sprintf "<row %s>%s</row>" (String.concat " " rowAttrs) cells)
    |> String.concat ""

let private mergeCellsXmlFull (mergedCells: MergedCellRange list) =
    if List.isEmpty mergedCells then ""
    else
        let body =
            mergedCells
            |> List.map (fun r -> sprintf "<mergeCell ref=\"%s:%s\"/>" (cellRef r.StartRow r.StartCol) (cellRef r.EndRow r.EndCol))
            |> String.concat ""
        sprintf "<mergeCells count=\"%d\">%s</mergeCells>" mergedCells.Length body

let private dataValidationsXml validations =
    if List.isEmpty validations then ""
    else
        let body =
            validations
            |> List.map (fun dv ->
                let attrs = [
                    yield attr "type" dv.ValidationType
                    yield attr "sqref" dv.Sqref
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
                let f1 = dv.Formula1 |> Option.map (elemText "formula1") |> Option.defaultValue ""
                let f2 = dv.Formula2 |> Option.map (elemText "formula2") |> Option.defaultValue ""
                sprintf "<dataValidation %s>%s%s</dataValidation>" (String.concat " " attrs) f1 f2)
            |> String.concat ""
        sprintf "<dataValidations count=\"%d\">%s</dataValidations>" validations.Length body

/// `<conditionalFormatting>` blocks. Rules keep their type-specific payload
/// (formulas, colour scale, data bar, icon set, top-N / above-average
/// attributes); a `dxfId` outside the workbook's dxf table is dropped rather
/// than written as a dangling index.
let private conditionalFormattingXml (dxfCount: int) (formats: XlsxReader.ConditionalFormat list) =
    let mutable nextPriority = 0
    let cfvo (kind: string, value: string option) =
        match value with
        | Some v -> sprintf "<cfvo %s %s/>" (attr "type" kind) (attr "val" v)
        | None -> sprintf "<cfvo %s/>" (attr "type" kind)
    let color (c: string) = sprintf "<color rgb=\"%s\"/>" (cleanRgb c)
    formats
    |> List.filter (fun cf -> not (List.isEmpty cf.Rules))
    |> List.map (fun cf ->
        let rules =
            cf.Rules
            |> List.map (fun (rule: XlsxReader.ConditionalFormatRule) ->
                nextPriority <- nextPriority + 1
                let attrs = [
                    yield attr "type" rule.Type
                    match rule.DxfId with
                    | Some id when id >= 0 && id < dxfCount -> yield sprintf "dxfId=\"%d\"" id
                    | _ -> ()
                    yield sprintf "priority=\"%d\"" (rule.Priority |> Option.defaultValue nextPriority)
                    if rule.StopIfTrue then yield "stopIfTrue=\"1\""
                    if rule.Type = "aboveAverage" && not rule.AboveAverage then yield "aboveAverage=\"0\""
                    if rule.Percent then yield "percent=\"1\""
                    if rule.Bottom then yield "bottom=\"1\""
                    match rule.Operator with Some op -> yield attr "operator" op | None -> ()
                    match rule.Text with Some t -> yield attr "text" t | None -> ()
                    match rule.Rank with Some r -> yield sprintf "rank=\"%d\"" r | None -> ()
                    match rule.StdDev with Some d -> yield sprintf "stdDev=\"%d\"" d | None -> ()
                    if rule.EqualAverage then yield "equalAverage=\"1\""
                ]
                let children = [
                    for f in rule.Formulas -> elemText "formula" f
                    if not (List.isEmpty rule.ColorScale) then
                        yield "<colorScale>"
                              + (rule.ColorScale |> List.map (fun stop -> cfvo (stop.CfvoType, stop.CfvoValue)) |> String.concat "")
                              + (rule.ColorScale |> List.map (fun stop -> color stop.Color) |> String.concat "")
                              + "</colorScale>"
                    if rule.Type = "dataBar" then
                        yield "<dataBar>"
                              + (rule.DataBarCfvo |> List.map cfvo |> String.concat "")
                              + color (rule.DataBarColor |> Option.defaultValue "638EC6")
                              + "</dataBar>"
                    match rule.IconSet with
                    | Some family ->
                        yield sprintf "<iconSet %s%s>%s</iconSet>" (attr "iconSet" family) (if rule.IconSetReverse then " reverse=\"1\"" else "") (rule.IconSetCfvo |> List.map cfvo |> String.concat "")
                    | None -> ()
                ]
                sprintf "<cfRule %s>%s</cfRule>" (String.concat " " attrs) (String.concat "" children))
            |> String.concat ""
        sprintf "<conditionalFormatting %s>%s</conditionalFormatting>" (attr "sqref" cf.Range) rules)
    |> String.concat ""

let private hyperlinksXml hyperlinks =
    if List.isEmpty hyperlinks then ""
    else
        let body =
            hyperlinks
            |> List.mapi (fun i h ->
                if h.IsInternal then
                    sprintf "<hyperlink %s %s/>" (attr "ref" h.Range) (attr "location" h.Target)
                else
                    sprintf "<hyperlink %s r:id=\"rIdHyperlink%d\"/>" (attr "ref" h.Range) (i + 1))
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
            ]
            sprintf "<pageSetup %s/>" (String.concat " " attrs))
        |> Option.defaultValue ""
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
            sprintf "<threadedComment ref=\"%s%d\"%s personId=\"%s\" id=\"%s\"%s><text>%s</text></threadedComment>"
                comment.CellRef.Column comment.CellRef.Row
                (comment.Date |> Option.map (fun date -> " " + attr "dT" date) |> Option.defaultValue "")
                (escaped (personId comment)) (escaped id)
                (parent |> Option.map (fun parentId -> " " + attr "parentId" parentId) |> Option.defaultValue "")
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
    match sheet.Extras.Format with
    | None -> "<sheetFormatPr defaultRowHeight=\"15\"/>"
    | Some format ->
        let attrs = [
            match format.BaseColWidth with Some width -> yield sprintf "baseColWidth=\"%s\"" (invariant width) | None -> ()
            match format.DefaultColWidth with Some width -> yield sprintf "defaultColWidth=\"%s\"" (invariant width) | None -> ()
            yield sprintf "defaultRowHeight=\"%s\"" (invariant (format.DefaultRowHeight |> Option.defaultValue 15.0))
            if format.CustomDefaultHeight then yield "customHeight=\"1\""
        ]
        sprintf "<sheetFormatPr %s/>" (String.concat " " attrs)

let private worksheetXml (styles: StyleBook) sharedIndex sheetIndex active (sheet: XlsxSheetOut) =
    let hasFitToPage = sheet.PageSetup |> Option.exists (fun p -> p.FitToPage)
    let sheetPr = if hasFitToPage then "<sheetPr><pageSetUpPr fitToPage=\"1\"/></sheetPr>" else ""
    declaration + "\n"
    + sprintf "<worksheet xmlns=\"%s\" xmlns:r=\"%s\">" Namespaces.spreadsheet Namespaces.relationships
    + sheetPr
    + sprintf "<dimension ref=\"%s\"/>" (sheetDimension sheet)
    + sheetViewsXml sheet active
    + sheetFormatXml sheet
    + colsXml styles sheet
    + ("<sheetData>" + rowsXml styles sharedIndex sheet + "</sheetData>")
    // Child order is fixed by CT_Worksheet; Excel refuses a part that
    // deviates (e.g. tableParts before legacyDrawing).
    + mergeCellsXmlFull sheet.MergedCells
    + conditionalFormattingXml styles.Dxfs.Length sheet.ConditionalFormats
    + dataValidationsXml sheet.DataValidations
    + hyperlinksXml sheet.Hyperlinks
    + pagePartsXml sheet
    + breaksXml "rowBreaks" sheet.RowBreaks
    + breaksXml "colBreaks" sheet.ColBreaks
    + (if hasDrawing sheet then sprintf "<drawing r:id=\"rIdDrawing%d\"/>" sheetIndex else "")
    + (if sheet.Cells |> List.exists (fun c -> c.Comment.IsSome) then sprintf "<legacyDrawing r:id=\"rIdCommentsVml%d\"/>" sheetIndex else "")
    + tablePartsXml sheet.Tables.Length
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
            let totals = if table.TotalsRowCount > 0 then " totalsRowFunction=\"none\"" else ""
            sprintf "<tableColumn id=\"%d\" %s%s/>" (i + 1) (attr "name" name) totals)
        |> String.concat ""
    let autoFilter =
        table.AutoFilter
        |> Option.map (fun _ -> sprintf "<autoFilter ref=\"%s\"/>" refText)
        |> Option.defaultValue ""
    let style =
        table.Style
        |> Option.map (fun s ->
            sprintf "<tableStyleInfo %s showFirstColumn=\"%d\" showLastColumn=\"%d\" showRowStripes=\"%d\" showColumnStripes=\"%d\"/>"
                (attr "name" s.Name) (if s.ShowFirstColumn then 1 else 0) (if s.ShowLastColumn then 1 else 0) (if s.ShowRowStripes then 1 else 0) (if s.ShowColumnStripes then 1 else 0))
        |> Option.defaultValue ""
    sprintf "%s\n<table xmlns=\"%s\" id=\"%d\" %s %s ref=\"%s\" headerRowCount=\"%d\" totalsRowCount=\"%d\">%s<tableColumns count=\"%d\">%s</tableColumns>%s</table>"
        declaration Namespaces.spreadsheet id (attr "name" table.Name) (attr "displayName" table.Name) refText table.HeaderRowCount table.TotalsRowCount autoFilter (table.EndCol - table.StartCol + 1) cols style

let private sheetRelsXml sheetIndex (sheet: XlsxSheetOut) firstTableId =
    let rels = [
        for i, h in sheet.Hyperlinks |> List.mapi (fun i h -> i + 1, h) do
            if not h.IsInternal then
                yield sprintf "<Relationship Id=\"rIdHyperlink%d\" Type=\"%s\" Target=\"%s\" TargetMode=\"External\"/>" i RelationshipTypes.hyperlink (escaped h.Target)
        for i in 1 .. sheet.Tables.Length do
            yield sprintf "<Relationship Id=\"rIdTable%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table%d.xml\"/>" i (firstTableId + i - 1)
        if sheet.Cells |> List.exists (fun c -> c.Comment.IsSome) then
            yield sprintf "<Relationship Id=\"rIdComments%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments%d.xml\"/>" sheetIndex sheetIndex
            yield sprintf "<Relationship Id=\"rIdCommentsVml%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing%d.vml\"/>" sheetIndex sheetIndex
        if hasDrawing sheet then
            yield sprintf "<Relationship Id=\"rIdDrawing%d\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"/xl/drawings/sheetDrawing%d.xml\"/>" sheetIndex sheetIndex
        if not sheet.Extras.ThreadedComments.IsEmpty then
            yield sprintf "<Relationship Id=\"rIdThreaded%d\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/threadedComment\" Target=\"/xl/threadedComments/threadedComment%d.xml\"/>" sheetIndex sheetIndex
    ]
    if List.isEmpty rels then None
    else Some (sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" rels))

let private workbookXmlFull (workbook: XlsxWorkbookOut) =
    let sheets =
        workbook.Sheets
        |> List.mapi (fun i s ->
            let visible =
                match s.State with
                | XlsxReader.SheetVisible -> ""
                | XlsxReader.SheetHidden -> " state=\"hidden\""
                | XlsxReader.SheetVeryHidden -> " state=\"veryHidden\""
            sprintf "<sheet %s sheetId=\"%d\" r:id=\"rId%d\"%s/>"
                (attr "name" s.Name) (i + 1) (i + 1) visible)
        |> String.concat ""
    // In a reference a sheet name is quoted with apostrophes, and an
    // apostrophe inside the name is doubled ('Bob''s'!$A$1).
    let quoted (name: string) = "'" + name.Replace("'", "''") + "'"
    let printNames =
        workbook.Sheets
        |> List.mapi (fun i s ->
            [ match s.PrintArea with
              | Some v -> yield { Name = "_xlnm.Print_Area"; RefersTo = sprintf "%s!%s" (quoted s.Name) v; LocalSheetIndex = Some i }
              | None -> ()
              match s.PrintTitleRows, s.PrintTitleCols with
              | None, None -> ()
              | rows, cols ->
                  let parts = [ rows; cols ] |> List.choose id |> List.map (fun part -> sprintf "%s!%s" (quoted s.Name) part)
                  yield { Name = "_xlnm.Print_Titles"; RefersTo = String.concat "," parts; LocalSheetIndex = Some i } ])
        |> List.concat
    let names = workbook.DefinedNames @ printNames
    let definedNames =
        if List.isEmpty names then ""
        else
            let body =
                names
                |> List.map (fun n ->
                    let local = n.LocalSheetIndex |> Option.map (fun i -> sprintf " localSheetId=\"%d\"" i) |> Option.defaultValue ""
                    sprintf "<definedName %s%s>%s</definedName>"
                        (attr "name" n.Name) local
                        (escaped n.RefersTo))
                |> String.concat ""
            "<definedNames>" + body + "</definedNames>"
    // A formula written without a result has to be calculated when the file is opened.
    let hasUncachedFormula =
        workbook.Sheets
        |> List.exists (fun s -> s.Cells |> List.exists (fun c -> match c.Value with XlsxFormula (_, None) -> true | _ -> false))
    let calcPr = if hasUncachedFormula then "<calcPr fullCalcOnLoad=\"1\"/>" else ""
    sprintf "%s\n<workbook xmlns=\"%s\" xmlns:r=\"%s\"><bookViews><workbookView activeTab=\"%d\"/></bookViews><sheets>%s</sheets>%s%s</workbook>"
        declaration Namespaces.spreadsheet Namespaces.relationships (max 0 workbook.ActiveSheetIndex) sheets definedNames calcPr

let private workbookRelsXmlFull (sheetCount: int) hasVba (themePath: string option) (hasPersons: bool) =
    let sheetRels =
        [ 1 .. sheetCount ]
        |> List.map (fun i -> sprintf "<Relationship Id=\"rId%d\" Type=\"%s\" Target=\"worksheets/sheet%d.xml\"/>" i RelationshipTypes.worksheet i)
    let baseRels = [
        yield! sheetRels
        yield sprintf "<Relationship Id=\"rIdStyles\" Type=\"%s\" Target=\"styles.xml\"/>" RelationshipTypes.styles
        yield sprintf "<Relationship Id=\"rIdSharedStrings\" Type=\"%s\" Target=\"sharedStrings.xml\"/>" RelationshipTypes.sharedStrings
        if hasVba then yield "<Relationship Id=\"rIdVBA\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>"
        // The theme of the opened workbook, kept as it was.
        match themePath with
        | Some path -> yield sprintf "<Relationship Id=\"rIdTheme\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"/%s\"/>" (escaped path)
        | None -> ()
        if hasPersons then yield "<Relationship Id=\"rIdPersons\" Type=\"http://schemas.microsoft.com/office/2017/10/relationships/person\" Target=\"/xl/persons/person.xml\"/>"
    ]
    sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" baseRels)

let private contentTypesXmlFull (workbook: XlsxWorkbookOut) =
    let hasVba = workbook.VbaProjectSource |> Option.exists hasText
    let workbookType =
        if hasVba then "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    let sheetOverrides =
        workbook.Sheets
        |> List.mapi (fun i _ -> sprintf "<Override PartName=\"/xl/worksheets/sheet%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" (i + 1))
        |> String.concat ""
    let mutable tableId = 1
    let tableOverrides =
        workbook.Sheets
        |> List.collect (fun s ->
            s.Tables |> List.map (fun _ ->
                let id = tableId
                tableId <- tableId + 1
                sprintf "<Override PartName=\"/xl/tables/table%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>" id))
        |> String.concat ""
    let comments =
        workbook.Sheets
        |> List.mapi (fun i s ->
            if s.Cells |> List.exists (fun c -> c.Comment.IsSome) then
                sprintf "<Override PartName=\"/xl/comments%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>" (i + 1)
            else "")
        |> String.concat ""
    // Pictures, the kept parts (charts, theme, embedded data) and threaded comments.
    let keptPaths = workbook.Preserved.Parts |> List.map (fun (path, _) -> ("/" + path).ToLowerInvariant()) |> Set.ofList
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
        [ for i, sheet in workbook.Sheets |> List.mapi (fun i sheet -> i + 1, sheet) do
              if hasDrawing sheet then
                  yield sprintf "<Override PartName=\"/xl/drawings/sheetDrawing%d.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>" i
              if not sheet.Extras.ThreadedComments.IsEmpty then
                  yield sprintf "<Override PartName=\"/xl/threadedComments/threadedComment%d.xml\" ContentType=\"application/vnd.ms-excel.threadedcomments+xml\"/>" i
          if workbook.Sheets |> List.exists (fun sheet -> not sheet.Extras.ThreadedComments.IsEmpty) then
              yield "<Override PartName=\"/xl/persons/person.xml\" ContentType=\"application/vnd.ms-excel.person+xml\"/>"
          for name, contentType in workbook.Preserved.Overrides do
              if Set.contains (name.ToLowerInvariant()) keptPaths then
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

let private rootRelsXmlFull =
    sprintf "%s\n<Relationships xmlns=\"%s\"><Relationship Id=\"rId1\" Type=\"%s\" Target=\"xl/workbook.xml\"/></Relationships>"
        declaration Namespaces.relationshipsPackage RelationshipTypes.officeDocument

let private workbookFiles (workbook: XlsxWorkbookOut) =
    let styles = collectStyles workbook
    let sharedItems, sharedIndex = collectSharedStrings workbook
    let hasVba = workbook.VbaProjectSource |> Option.exists hasText
    let mutable tableId = 1
    let textFiles = ResizeArray<PackageFile>()
    textFiles.Add { Path = "[Content_Types].xml"; Content = contentTypesXmlFull workbook }
    textFiles.Add { Path = "_rels/.rels"; Content = rootRelsXmlFull }
    textFiles.Add { Path = "xl/workbook.xml"; Content = workbookXmlFull workbook }
    let persons = personsOf workbook
    let hasThreads = workbook.Sheets |> List.exists (fun sheet -> not sheet.Extras.ThreadedComments.IsEmpty)
    // The theme stays only if its part was kept.
    let themePath = workbook.Preserved.ThemePath |> Option.filter (fun path -> workbook.Preserved.Parts |> List.exists (fun (kept, _) -> kept = path))
    textFiles.Add { Path = "xl/_rels/workbook.xml.rels"; Content = workbookRelsXmlFull workbook.Sheets.Length hasVba themePath hasThreads }
    if hasThreads then textFiles.Add { Path = "xl/persons/person.xml"; Content = personsXml persons }
    let keptPartPaths = workbook.Preserved.Parts |> List.map fst |> Set.ofList
    let mediaFiles = ResizeArray<string * byte[]>()
    let mutable nextMedia = 1
    textFiles.Add { Path = "xl/styles.xml"; Content = stylesXml styles }
    textFiles.Add { Path = "xl/sharedStrings.xml"; Content = sharedStringsXmlPreserved sharedItems }
    workbook.Sheets
    |> List.iteri (fun i sheet ->
        let sheetIndex = i + 1
        let firstTableId = tableId
        textFiles.Add { Path = sprintf "xl/worksheets/sheet%d.xml" sheetIndex; Content = worksheetXml styles sharedIndex sheetIndex (i = workbook.ActiveSheetIndex) sheet }
        match sheetRelsXml sheetIndex sheet firstTableId with
        | Some rels -> textFiles.Add { Path = sprintf "xl/worksheets/_rels/sheet%d.xml.rels" sheetIndex; Content = rels }
        | None -> ()
        match commentsXml sheet.Cells with
        | Some xml -> textFiles.Add { Path = sprintf "xl/comments%d.xml" sheetIndex; Content = xml }
        | None -> ()
        match vmlCommentsXml sheetIndex sheet.Cells with
        | Some xml -> textFiles.Add { Path = sprintf "xl/drawings/vmlDrawing%d.vml" sheetIndex; Content = xml }
        | None -> ()
        match threadedCommentsXml persons sheetIndex sheet with
        | Some xml -> textFiles.Add { Path = sprintf "xl/threadedComments/threadedComment%d.xml" sheetIndex; Content = xml }
        | None -> ()
        if hasDrawing sheet then
            // One media part per picture (not under the name of a kept part), and
            // the relationships the kept objects refer to, with their ids.
            let imageRels =
                sheet.Extras.Images
                |> List.mapi (fun i image ->
                    let extension = if image.Extension = "" then "bin" else image.Extension
                    while Set.contains (sprintf "xl/media/image%d.%s" nextMedia extension) keptPartPaths do
                        nextMedia <- nextMedia + 1
                    let path = sprintf "xl/media/image%d.%s" nextMedia extension
                    nextMedia <- nextMedia + 1
                    mediaFiles.Add (path, image.Data)
                    sprintf "rIdImg%d" (i + 1), path)
            let relsBody =
                [ for relId, path in imageRels do
                      yield sprintf "<Relationship Id=\"%s\" Type=\"%s\" Target=\"/%s\"/>" relId RelationshipTypes.image path
                  for relId, relType, target, isExternal in sheet.Extras.DrawingRelationships do
                      yield sprintf "<Relationship %s %s %s%s/>" (attr "Id" relId) (attr "Type" relType) (attr "Target" target) (if isExternal then " TargetMode=\"External\"" else "") ]
            textFiles.Add { Path = sprintf "xl/drawings/sheetDrawing%d.xml" sheetIndex; Content = drawingXml sheet (imageRels |> List.map fst) }
            if not relsBody.IsEmpty then
                textFiles.Add
                    { Path = sprintf "xl/drawings/_rels/sheetDrawing%d.xml.rels" sheetIndex
                      Content = sprintf "%s\n<Relationships xmlns=\"%s\">%s</Relationships>" declaration Namespaces.relationshipsPackage (String.concat "" relsBody) }
        for table in sheet.Tables do
            textFiles.Add { Path = sprintf "xl/tables/table%d.xml" tableId; Content = tableXml tableId table }
            tableId <- tableId + 1)
    let written = textFiles |> Seq.map (fun file -> file.Path) |> Set.ofSeq
    let binaryFiles =
        (match workbook.VbaProjectSource with
         | Some source when source <> "" -> [ "xl/vbaProject.bin", VbaProject.writeProject "Module1" source ]
         | _ -> [])
        @ List.ofSeq mediaFiles
        // Kept parts go back as they were, unless the writer produced a part of that name.
        @ (workbook.Preserved.Parts |> List.filter (fun (path, _) -> not (Set.contains path written)))
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
    let runOf (props: XlsxReader.RunProps) text =
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
                          Format =
                            (match namedStyleOf cell with
                             | None -> styleOf cell.Style
                             | named -> { styleOf cell.Style with NamedStyle = named })
                          Comment = Map.tryFind (r, c) comments }))
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
                      AutoFilter = t.AutoFilter
                      Style =
                        t.Style |> Option.map (fun s ->
                            { Name = s.Name
                              ShowFirstColumn = s.ShowFirstColumn
                              ShowLastColumn = s.ShowLastColumn
                              ShowRowStripes = s.ShowRowStripes
                              ShowColumnStripes = s.ShowColumnStripes }) })
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
              Hyperlinks = sheet.Hyperlinks |> List.map (fun h -> { Range = h.Range; Target = h.Target; IsInternal = h.IsInternal })
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
                        | _ -> None })
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
                    ColumnStyles =
                        sheet.Columns
                        |> List.choose (fun column ->
                            match column.Style with
                            | Some index when index > 0 && index < workbook.Styles.Length -> Some (column.Index - 1, column.MaxIndex - 1, workbook.Styles.[index])
                            | _ -> None) } })
      ActiveSheetIndex = 0
      DefinedNames = workbook.NamedRanges |> List.map (fun n -> { Name = n.Name; RefersTo = n.RefersTo; LocalSheetIndex = n.Scope |> Option.bind (fun s -> match System.Int32.TryParse s with true, v -> Some v | _ -> None) })
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
