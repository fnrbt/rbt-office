module XlsxReader

open XmlDsl
open OfficeDsl
open ZipInterop
open Fable.Core.JsInterop
open System.Text

// XLSX domain model
type CellReference = {
    Column: string
    Row: int
}

type CellValue =
    | Number of float
    | String of string
    | SharedString of int
    | Boolean of bool
    | Date of string
    | Formula of string
    | Error of string
    | Empty

/// `<vertAlign val="..."/>` on an rPr or font element. Excel uses these
/// for chemistry (H₂O), math (x²), and footnote markers. Renderers shift
/// the baseline + scale the glyph size; we map both onto a small DU so
/// downstream code doesn't have to string-match.
type VertAlign =
    | VertAlignBaseline
    | VertAlignSuperscript
    | VertAlignSubscript

/// `<u val="..."/>` style on an rPr or font element. None = no underline,
/// Single = standard one-line, Double = two parallel lines (used for
/// accounting totals etc.). Excel also defines singleAccounting /
/// doubleAccounting which render at the full cell width rather than just
/// the text width; we collapse those onto Single/Double respectively since
/// our underline stroke already spans the run's measured width.
type UnderlineStyle =
    | UnderlineNone
    | UnderlineSingle
    | UnderlineDouble

/// Optional per-run formatting from an `<rPr>` inside an `<si><r>` element.
/// Any field set on the run overrides the cell's base style for that text;
/// fields left None inherit from the surrounding cell format.
type RunProps = {
    Bold: bool
    Italic: bool
    Underline: bool
    /// When Underline is true, this flag distinguishes the `double`
    /// (and `doubleAccounting`) styles from the default single line.
    /// Renderers emit two parallel strokes when true.
    UnderlineDouble: bool
    Strikethrough: bool
    Color: string option       // ARGB hex without '#' prefix
    FontName: string option
    FontSize: float option     // points
    /// `<scheme val="major|minor"/>` — when present (and no explicit
    /// `<name>` overrides), the renderer resolves this against the
    /// workbook theme's `majorFont`/`minorFont` slot. "" = absent.
    Scheme: string
    VertAlign: VertAlign
}

let emptyRunProps = {
    Bold = false
    Italic = false
    Underline = false
    UnderlineDouble = false
    Strikethrough = false
    Color = None
    FontName = None
    FontSize = None
    Scheme = ""
    VertAlign = VertAlignBaseline
}

/// A parsed shared-string entry. Plain SI entries have one Run with empty
/// props and the entire text. Rich entries have one Run per `<r>` element
/// each carrying its own `<rPr>`-derived RunProps.
type SharedStringEntry = {
    Text: string                       // concatenated for indexing / search
    Runs: (string * RunProps) list     // empty list ≡ no rich runs (use Text)
}

/// One side of a cell border. `Style` is the OOXML border style name
/// (thin/medium/thick/dashed/dotted/double/...), `Color` is an ARGB hex
/// string when explicit (or None when Excel's "automatic" black).
type CellBorderSide = {
    Style: string
    Color: string option
}

type CellStyle = {
    FontName: string option
    FontColor: string option
    FillColor: string option
    Bold: bool
    Italic: bool
    Underline: bool
    UnderlineDouble: bool
    Strikethrough: bool
    FontSize: float option
    NumberFormat: string option
    HorizontalAlign: string option
    VerticalAlign: string option
    WrapText: bool
    /// `<alignment indent="N"/>` — N units of left-padding when horizontal
    /// align is left/right/distributed. Each unit is ~3 spaces wide.
    Indent: int
    /// `<alignment shrinkToFit="1"/>` — auto-shrink the cell's text size to
    /// fit the cell width when it would otherwise overflow.
    ShrinkToFit: bool
    /// `<alignment textRotation="N"/>` — degrees CCW (0..90), or 91..180
    /// for CW (Excel encodes CW as 90 + |angle|), or 255 for stacked
    /// (each character on its own line). 0 = horizontal (default).
    TextRotation: int
    BorderTop: CellBorderSide option
    BorderBottom: CellBorderSide option
    BorderLeft: CellBorderSide option
    BorderRight: CellBorderSide option
    BorderDiagonal: CellBorderSide option
    DiagonalUp: bool
    DiagonalDown: bool
}

type MergeCell = {
    StartRef: CellReference
    EndRef: CellReference
}

type Cell = {
    Reference: CellReference
    /// The cell's value. For a formula cell this is the result Excel cached
    /// in `<v>`; only a formula with no cached result (an uncalculated
    /// workbook) surfaces as `Formula text` here.
    Value: CellValue
    /// Formula text without the leading `=`, when the cell has an `<f>`.
    /// Shared formulas (`t="shared"`) are expanded: every cell in the group
    /// carries the master's text with its relative references shifted.
    Formula: string option
    /// For the anchor of a legacy array formula (`<f t="array" ref="B1:B3">`),
    /// the range the formula fills. None for ordinary formulas.
    ArrayRef: string option
    StyleIndex: int option
    Style: CellStyle option
    /// Per-run rich-text formatting for inline-string cells whose
    /// `<is>` element contains `<r>` children (instead of a single
    /// `<t>`). Mirrors the sharedStrings <r> grammar, but the runs
    /// live on the cell rather than in the workbook-level
    /// SharedStringRuns table. Empty for non-inline cells and for
    /// plain inline strings (single <t>).
    InlineRuns: (string * RunProps) list
}

type Row = {
    Index: int
    Cells: Cell list
    /// Height to apply when rendering. Some when the row's `customHeight`
    /// is `true` (or absent — defaults to true), in which case the stored
    /// `ht` is the row's explicit height. None when `customHeight="false"`,
    /// meaning Excel auto-derives the height from the row's font.
    Height: float option
    /// Raw `ht` attribute value as written in the XML, regardless of
    /// `customHeight`. Lets the runner detect the case where
    /// `customHeight="false" ht="X"` and `X == sheetFormatPr.defaultRowHeight`,
    /// which Excel treats as implicit (uses defaultRowHeight literally
    /// rather than font-derived auto-height). openpyxl writes this pattern
    /// for every row whose height matches the sheet default.
    RawHeight: float option
    /// Value of the row's `customHeight` attribute, or None when absent.
    CustomHeight: bool option
    /// `<row hidden="1">` collapses the row in display + print (height = 0).
    Hidden: bool
}

type ColumnInfo = {
    Index: int
    MaxIndex: int
    Width: float option
    Hidden: bool
    /// `<col style="N"/>` — default cellXf index for every cell in this
    /// column that doesn't carry its own `s="..."`. None when absent.
    Style: int option
}

// Advanced features
type CellComment = {
    CellRef: CellReference
    Author: string option
    Text: string
}

/// One `<cfvo>` + `<color>` pair inside a `<colorScale>`. The cfvo
/// describes where on the value axis the stop sits (min / max /
/// percentile / num / percent); Color is the RRGGBB applied there.
type CfColorStop = {
    CfvoType: string
    CfvoValue: string option
    Color: string
}

type ConditionalFormatRule = {
    Type: string  // "cellIs", "expression", "colorScale", "dataBar", etc.
    /// cellIs operator: greaterThan / lessThan / between / equal /
    /// notEqual / greaterThanOrEqual / lessThanOrEqual. None for
    /// non-cellIs rule types.
    Operator: string option
    /// One formula for most cellIs operators; two for `between` /
    /// `notBetween`. Empty for colorScale / dataBar / iconSet.
    Formulas: string list
    /// Index into the workbook's `<dxfs>` differential-format table —
    /// the overlay applied when a cellIs / expression rule matches.
    DxfId: int option
    /// Ordered stops for a `colorScale` rule (2 or 3 entries). Empty
    /// for other rule types.
    ColorScale: CfColorStop list
    /// `dataBar` bar fill color (RRGGBB). None for non-dataBar rules.
    DataBarColor: string option
    /// `dataBar` min/max cfvo bounds — (type, optional value) pairs.
    /// Two entries when present; empty otherwise.
    DataBarCfvo: (string * string option) list
    /// `text` attribute of containsText / notContainsText / beginsWith /
    /// endsWith rules — the literal the rule searches for.
    Text: string option
    /// `iconSet` family name (e.g. "3TrafficLights1", "5Arrows").
    /// None for non-iconSet rules.
    IconSet: string option
    /// `iconSet` threshold cfvos — N entries for an N-icon set, each
    /// (type, optional value). The first is always the implicit 0%
    /// floor; icon i shows when the value is ≥ threshold i.
    IconSetCfvo: (string * string option) list
    /// `top10` rank — the N in "top/bottom N". When Percent is set, N
    /// is a percentage of the range count. None for non-top10 rules.
    Rank: int option
    /// `top10` percent flag — Rank is a percentage, not a count.
    Percent: bool
    /// `top10` bottom flag — match the bottom N instead of the top N.
    Bottom: bool
    /// `aboveAverage` direction — true = above the range mean (the XML
    /// default), false = below. Only meaningful for aboveAverage rules.
    AboveAverage: bool
    /// `aboveAverage` equalAverage flag — include cells exactly equal
    /// to the mean in the match.
    EqualAverage: bool
    /// `aboveAverage` stdDev — when > 0, match cells beyond N standard
    /// deviations from the mean (in the AboveAverage direction) rather
    /// than simply above/below it.
    StdDev: int option
    Priority: int option
    StopIfTrue: bool
}

type ConditionalFormat = {
    Range: string
    Rules: ConditionalFormatRule list
}

/// A `<dxf>` "differential format" from styles.xml — the partial style
/// overlay a conditional-formatting rule applies on top of a cell's
/// existing format. Only the fields a dxf can carry are modeled; all
/// optional, since a dxf typically sets just a fill or just a color.
type Dxf = {
    FontColor: string option
    FillColor: string option
    Bold: bool option
    Italic: bool option
}

type DataValidation = {
    Range: string
    ValidationType: string  // "list", "whole", "decimal", "date", "time", "textLength", "custom"
    Operator: string option
    Formula1: string option
    Formula2: string option
    AllowBlank: bool
    ShowDropDown: bool
    ErrorMessage: string option
    /// `errorTitle` — caption of the error alert.
    ErrorTitle: string option
    /// `errorStyle`: "stop" (default, rejects the entry), "warning" or
    /// "information". None when the attribute is absent.
    ErrorStyle: string option
    /// `showErrorMessage`. Excel only enforces a validation (shows the
    /// alert and rejects the entry) when this is set.
    ShowErrorMessage: bool
    /// `prompt` — the input message shown while the cell is selected.
    Prompt: string option
    /// `promptTitle` — caption of the input message.
    PromptTitle: string option
    /// `showInputMessage`.
    ShowInputMessage: bool
}

type NamedRange = {
    Name: string
    RefersTo: string
    Scope: string option  // worksheet name or "Workbook"
}

type ChartType =
    | LineChart
    | BarChart
    | PieChart
    | AreaChart
    | ScatterChart
    | OtherChart of string

type ChartSeries = {
    Name: string option
    Values: string  // Reference like "Sheet1!$B$2:$B$10"
    Categories: string option  // Reference like "Sheet1!$A$2:$A$10"
}

type Chart = {
    Id: string option
    Type: ChartType
    Title: string option
    Series: ChartSeries list
}

type PivotTableField = {
    Name: string
    Index: int
}

type PivotTable = {
    Name: string
    DataSource: string  // Reference to source range
    Location: string  // Reference to pivot table location
    RowFields: PivotTableField list
    ColumnFields: PivotTableField list
    DataFields: PivotTableField list
    FilterFields: PivotTableField list
}

/// sheetFormatPr defaults captured at parse time so renderers can fall back
/// without re-reading the sheet XML.
type SheetFormat = {
    BaseColWidth: float option
    DefaultColWidth: float option
    DefaultRowHeight: float option
    /// `<sheetFormatPr customHeight="1"/>` — when true, Excel treats
    /// `defaultRowHeight` as an explicit row height for every row that
    /// doesn't carry its own `ht`, rather than as a hint the font-derived
    /// auto-height can override. Without honoring this flag, fixtures
    /// like pandas-ints_spelled_with_decimals (which sets
    /// `customHeight="1" defaultRowHeight="15.75"` to lock every row at
    /// 15.75 pt regardless of the Arial 10 cells inside it) drift below
    /// Excel's grid.
    CustomDefaultHeight: bool
}

/// XLSX table style reference (e.g. `<tableStyleInfo name="TableStyleMedium7"
/// showRowStripes="1"/>`). The named style is one of Excel's built-ins; we
/// pair it with the workbook theme accent colors at render time.
type TableStyleInfo = {
    Name: string
    ShowFirstColumn: bool
    ShowLastColumn: bool
    ShowRowStripes: bool
    ShowColumnStripes: bool
}

/// XLSX table definition (`xl/tables/tableN.xml`). Range coordinates are
/// 1-indexed to match the rest of the reader.
type TableDefinition = {
    /// `displayName` (fallback `name`) of the `<table>` element — the identifier
    /// structured references use (`Table1[Amount]`).
    Name: string
    /// Column names in left→right order, from `<tableColumns><tableColumn name>`.
    Columns: string list
    StartRow: int
    StartCol: int
    EndRow: int
    EndCol: int
    HeaderRowCount: int
    /// `totalsRowCount` — number of totals rows at the bottom (0 or 1 in practice).
    TotalsRowCount: int
    Style: TableStyleInfo option
    /// (startCol, endCol) span of `<autoFilter ref="…">`, 1-indexed. None when
    /// the table has no autoFilter child. Excel renders a dropdown button on
    /// each header cell whose column falls in this range; row coords are
    /// implicit (header is always row StartRow).
    AutoFilter: (int * int) option
}

/// `<pane xSplit ySplit state="frozen" topLeftCell=…/>` inside a sheet's
/// `<sheetView>`. xSplit / ySplit are the column / row counts that stay
/// pinned (NOT pixel offsets — Excel treats them as integer cell counts
/// when state="frozen"). Sheets without a frozen pane omit this entirely.
type FrozenPane = {
    XSplit: int
    YSplit: int
}

/// `<hyperlink ref="A1" r:id="..." location="..." display="..."/>` on a
/// sheet. External links use `r:id` to point at a Relationship target;
/// internal links use `location` directly (e.g. `Sheet2!A1`). `Display`
/// holds the text Excel shows when it differs from the cell value;
/// usually None. Range is the same string Excel writes (single cell
/// "A1" or block "A1:C3").
type Hyperlink = {
    Range: string
    Target: string
    IsInternal: bool
}

/// Print-related sheet metadata, parsed from <pageSetup>, <pageMargins>,
/// <printOptions>, <headerFooter>, <rowBreaks>/<colBreaks>, and the
/// workbook's _xlnm.Print_Area / _xlnm.Print_Titles defined names.

type PrintOrientation =
    | OrientationDefault
    | OrientationPortrait
    | OrientationLandscape

/// `<pageMargins>` — all values in inches (matching OOXML's encoding).
type PageMargins = {
    Top: float
    Bottom: float
    Left: float
    Right: float
    Header: float
    Footer: float
}

/// `<pageSetup pageOrder>` enum. "downThenOver" (Excel default) walks all
/// row pages of the first column slice before advancing to the next column
/// slice. "overThenDown" does the inverse.
type PageOrder =
    | DownThenOver
    | OverThenDown

/// `<pageSetup cellComments>` controls how cell comments appear in print
/// output: "none" (default) suppresses them, "asDisplayed" paints the
/// red corner marker as it appears on screen, "atEnd" lists comments
/// after the last printed page.
type CellCommentsPrint =
    | CellCommentsNone
    | CellCommentsAsDisplayed
    | CellCommentsAtEnd

/// `<pageSetup>` plus `<sheetPr><pageSetUpPr fitToPage>`. PaperSize is
/// OOXML's enum (1 = Letter, 9 = A4, …); None when unspecified.
type PageSetup = {
    PaperSize: int option
    Orientation: PrintOrientation
    /// `fitToWidth` / `fitToHeight` — only meaningful when FitToPage = true.
    /// Default is 1/1 (fit to one page each way).
    FitToWidth: int
    FitToHeight: int
    /// Manual percent scale (10..400); 100 = no scale. Ignored when
    /// FitToPage = true.
    Scale: int
    /// `<pageSetUpPr fitToPage="1"/>`.
    FitToPage: bool
    /// Page number to display for the first page; `&P` substitutes this
    /// value on page 1 and increments from there. Only honored when
    /// `UseFirstPageNumber` is true.
    FirstPageNumber: int
    /// `useFirstPageNumber="1"` opts in to FirstPageNumber; otherwise
    /// Excel ignores it and numbers from 1.
    UseFirstPageNumber: bool
    /// Direction of multi-page walk; default DownThenOver.
    PageOrder: PageOrder
    /// `blackAndWhite="1"` — strip fills, force text + borders to black.
    BlackAndWhite: bool
    /// `cellComments` attribute. Default `CellCommentsNone` (Excel's
    /// default is to suppress comments in printed output).
    CellComments: CellCommentsPrint
}

/// `<printOptions>` — boolean flags that affect what gets printed.
type PrintOptions = {
    GridLines: bool
    Headings: bool
    HorizontalCentered: bool
    VerticalCentered: bool
}

/// `<headerFooter>` — Excel's `&L`/`&C`/`&R` field-coded strings for each
/// section. `differentOddEven` / `differentFirst` flag whether the even /
/// first variants are honored. `scaleWithDoc` (default true) controls
/// whether H/F text scales together with the page's print scale.
type HeaderFooter = {
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

/// Per-sheet print metadata derived from the workbook's `_xlnm.*` defined
/// names. `PrintArea` is a range like `$A$1:$F$20`; `PrintTitleRows` /
/// `PrintTitleCols` come from a `_xlnm.Print_Titles` value such as
/// `Sheet1!$1:$3,Sheet1!$A:$B` (either part may be missing).
type WorksheetPrintNames = {
    PrintArea: string option
    PrintTitleRows: string option
    PrintTitleCols: string option
}

let emptyPrintNames =
    { PrintArea = None; PrintTitleRows = None; PrintTitleCols = None }

let defaultPrintOptions =
    { GridLines = false; Headings = false; HorizontalCentered = false; VerticalCentered = false }

let defaultHeaderFooter =
    { OddHeader = None; OddFooter = None; EvenHeader = None; EvenFooter = None
      FirstHeader = None; FirstFooter = None; DifferentOddEven = false; DifferentFirst = false
      ScaleWithDoc = true }

/// `<sheet state="…">` in workbook.xml. Hidden sheets can be unhidden from
/// the UI; very hidden ones only through VBA. Neither shows a tab.
type SheetState =
    | SheetVisible
    | SheetHidden
    | SheetVeryHidden

type Worksheet = {
    Name: string
    /// Tab visibility, from the workbook's `<sheet state>`.
    State: SheetState
    Rows: Row list
    MergedCells: MergeCell list
    Columns: ColumnInfo list
    Format: SheetFormat
    Comments: CellComment list
    ConditionalFormats: ConditionalFormat list
    DataValidations: DataValidation list
    Charts: Chart list
    PivotTables: PivotTable list
    Tables: TableDefinition list
    FrozenPane: FrozenPane option
    /// `<sheetView showZeros="false"/>`. When false, cells with numeric
    /// value 0 render as blank (Excel default = true).
    ShowZeros: bool
    /// `<sheetView showGridLines="0"/>`. False = hide the screen gridlines.
    ShowGridLines: bool
    Hyperlinks: Hyperlink list
    /// `<pageSetup>` + `<sheetPr><pageSetUpPr>`. None when neither is present.
    PageSetup: PageSetup option
    /// `<pageMargins>`. None when absent.
    PageMargins: PageMargins option
    /// `<printOptions>` (defaulted when absent).
    PrintOptions: PrintOptions
    /// `<headerFooter>` (defaulted when absent).
    HeaderFooter: HeaderFooter
    /// `<rowBreaks>` — 1-based row numbers where a page break sits AFTER
    /// the given row.
    RowBreaks: int list
    /// `<colBreaks>` — 1-based column numbers where a break sits AFTER
    /// the given column.
    ColBreaks: int list
    /// Print area, from `_xlnm.Print_Area` (sheet-qualifier stripped).
    PrintArea: string option
    /// Repeat-row rows, from the row part of `_xlnm.Print_Titles`
    /// (e.g. `$1:$3`).
    PrintTitleRows: string option
    /// Repeat-col columns, from the col part of `_xlnm.Print_Titles`
    /// (e.g. `$A:$B`).
    PrintTitleCols: string option
}

type Workbook = {
    Sheets: Worksheet list
    /// Concatenated text of each shared-string entry, indexed by SI position.
    /// For cells, this is what the value displays as.
    SharedStrings: string array
    /// Parsed runs for shared-string entries that have rich text. Same length
    /// as `SharedStrings`; entries with no rich runs have an empty list.
    SharedStringRuns: (string * RunProps) list array
    Styles: CellStyle array
    /// Differential formats from styles.xml `<dxfs>`, indexed by dxfId.
    /// Conditional-formatting cellIs / expression rules reference these.
    Dxfs: Dxf array
    Relationships: Relationship list
    NamedRanges: NamedRange list
    ThemeAccents: string array
    /// True when the source workbook uses the 1904 date system
    /// (`<workbookPr date1904="1"/>`). Every consumer of this model assumes
    /// 1900-system serials, so the reader has already shifted the value of
    /// each date-formatted numeric cell by +1462 days; the flag records the
    /// origin only. Numbers that are dates but carry no date number format
    /// are left as stored.
    Date1904: bool
}

// File content is always written with invariant formatting, whatever the
// machine's culture: every numeric parse of package text goes through these.
// (Under de-DE a culture-sensitive parse reads "1.5" as 15; under sv-SE it
// rejects "-1" because the culture's minus sign is U+2212.)
let internal tryParseInt (text: string) : int option =
    match System.Int32.TryParse(text, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture) with
    | true, value -> Some value
    | _ -> None

let internal tryParseFloat (text: string) : float option =
    match System.Double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
    | true, value -> Some value
    | _ -> None

let internal isTrueText (text: string) =
    text = "1" || text = "true"

// DSL for XLSX-specific patterns
module Xlsx =
    // Namespace-aware helpers
    let inline sElem name node = Ns.childNs name node
    let inline sElems name node = Ns.descendantsNs name node

    let private childElems localName node =
        Xml.children node
        |> List.choose (fun child ->
            match Xml.name child with
            | Some name when Ns.stripNs name = localName -> Some child
            | _ -> None)

    /// Undo SpreadsheetML's `_xHHHH_` escaping (ST_Xstring): characters that
    /// XML cannot carry — control characters, and a carriage return that
    /// must survive line-end normalisation — are stored as `_x000D_` etc.
    /// A literal underscore that would otherwise read as an escape is itself
    /// written `_x005F_`, which this decodes back to `_`.
    let internal decodeXmlEscapes (text: string) : string =
        if not (text.Contains "_x") then text
        else
            let isHex (c: char) =
                (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F')
            let sb = StringBuilder(text.Length)
            let mutable i = 0
            while i < text.Length do
                if text.[i] = '_'
                   && i + 6 < text.Length
                   && text.[i + 1] = 'x'
                   && text.[i + 6] = '_'
                   && isHex text.[i + 2] && isHex text.[i + 3] && isHex text.[i + 4] && isHex text.[i + 5] then
                    sb.Append(char (System.Convert.ToInt32(text.Substring(i + 2, 4), 16))) |> ignore
                    i <- i + 7
                else
                    sb.Append(text.[i]) |> ignore
                    i <- i + 1
            sb.ToString()

    /// Concatenated text of every `<localName>` element under `node`, in
    /// document order, skipping phonetic runs. Each element's text is
    /// `_xHHHH_`-decoded.
    let internal descendantTextConcat localName node =
        let sb = StringBuilder()
        let rec walk current =
            match current with
            | XmlDsl.Element(name, _, children) ->
                if Ns.stripNs name = localName then
                    let own = StringBuilder()
                    for child in children do
                        match child with
                        | XmlDsl.Text text -> own.Append(text) |> ignore
                        | _ -> walk child
                    sb.Append(decodeXmlEscapes (own.ToString())) |> ignore
                elif Ns.stripNs name = "rPh" then
                    // Phonetic (furigana) run: a reading aid, not cell text.
                    ()
                else
                    for child in children do
                        walk child
            | _ -> ()
        walk node
        sb.ToString()

    /// 1-based column index -> letters ("A", "B", ..., "AA").
    let internal indexToColumnLetters (n: int) : string =
        let mutable k = n
        let mutable letters = ""
        while k > 0 do
            let rem = (k - 1) % 26
            letters <- string (char (int 'A' + rem)) + letters
            k <- (k - 1) / 26
        letters

    /// Column letters (A, B, ..., AA) -> 1-based column index; 0 when empty.
    let internal columnLettersToIndex (letters: string) =
        let mutable acc = 0
        for c in letters do
            if System.Char.IsLetter c then
                acc <- acc * 26 + (int (System.Char.ToUpperInvariant c) - int 'A' + 1)
        acc

    let private richRunText (runs: (string * RunProps) list) =
        let sb = StringBuilder()
        for text, _ in runs do
            sb.Append(text) |> ignore
        sb.ToString()

    // Cell reference parsing DSL
    let private parseCellRefParts (value: string) =
        let mutable index = 0
        while index < value.Length && value.[index] = '$' do
            index <- index + 1

        let letterStart = index
        while index < value.Length && System.Char.IsLetter value.[index] do
            index <- index + 1

        let letters =
            if index > letterStart then value.Substring(letterStart, index - letterStart)
            else ""

        while index < value.Length && value.[index] = '$' do
            index <- index + 1

        let mutable row = 0
        let mutable hasRow = false
        while index < value.Length && System.Char.IsDigit value.[index] do
            hasRow <- true
            row <- row * 10 + int value.[index] - int '0'
            index <- index + 1

        letters, if hasRow then row else 0

    let parseCellRef (ref: string) : CellReference =
        let letters, row = parseCellRefParts ref
        {
            Column = letters
            Row = row
        }

    // Rich-text run extraction. Used by both the sharedStrings parser (one
    // <r>* block per <si>) and the inline-string cellValue path (one <r>*
    // block per <is> on a t="inlineStr" cell). The two grammars are
    // identical; only the parent element differs.
    let internal parseRunProps (rPr: XmlNode) : RunProps =
        let boolChild key =
            // <b/> means bold; <b val="0"/> means explicitly not bold.
            match sElem key rPr with
            | Some node ->
                Xml.attr "val" node
                |> Option.map (fun v -> v <> "0" && v <> "false")
                |> Option.defaultValue true
            | None -> false
        let colorHex =
            match sElem "color" rPr with
            | Some node ->
                Xml.attr "rgb" node
                |> Option.map (fun raw ->
                    let t = raw.TrimStart('#')
                    if t.Length = 8 then t.Substring(2) else t)
            | None -> None
        let fontName =
            sElem "rFont" rPr |> Option.bind (Xml.attr "val")
        let fontSize =
            sElem "sz" rPr
            |> Option.bind (Xml.attr "val")
            |> Option.bind tryParseFloat
        let vertAlign =
            sElem "vertAlign" rPr
            |> Option.bind (Xml.attr "val")
            |> Option.map (fun v ->
                match v with
                | "superscript" -> VertAlignSuperscript
                | "subscript" -> VertAlignSubscript
                | _ -> VertAlignBaseline)
            |> Option.defaultValue VertAlignBaseline
        let underline = boolChild "u"
        let underlineDouble =
            sElem "u" rPr
            |> Option.bind (Xml.attr "val")
            |> Option.map (fun v -> v = "double" || v = "doubleAccounting")
            |> Option.defaultValue false
        let scheme =
            sElem "scheme" rPr
            |> Option.bind (Xml.attr "val")
            |> Option.defaultValue ""
        {
            Bold = boolChild "b"
            Italic = boolChild "i"
            Underline = underline
            UnderlineDouble = underlineDouble
            Strikethrough = boolChild "strike"
            Color = colorHex
            FontName = fontName
            FontSize = fontSize
            Scheme = scheme
            VertAlign = vertAlign
        }

    let internal richRunsOf (parent: XmlNode) : (string * RunProps) list =
        // Each <r> = run; the <t> inside it is the run's text. <rPr> is the
        // run's formatting (may be absent → defaults). Iterate in source
        // order so the spans stay contiguous.
        Xml.children parent
        |> List.choose (fun child ->
            match Xml.name child with
            | Some "r" ->
                let text = descendantTextConcat "t" child
                let props =
                    sElem "rPr" child
                    |> Option.map parseRunProps
                    |> Option.defaultValue emptyRunProps
                Some (text, props)
            | _ -> None)

#if !FABLE_COMPILER
    let private sharedStringBoolAttr (reader: System.Xml.XmlReader) =
        match reader.GetAttribute("val") with
        | null -> true
        | value -> value <> "0" && not (System.String.Equals(value, "false", System.StringComparison.OrdinalIgnoreCase))

    let private sharedStringTextPreserve (reader: System.Xml.XmlReader) =
        let xmlSpace = reader.GetAttribute("space", "http://www.w3.org/XML/1998/namespace")
        let xmlSpace = if isNull xmlSpace then reader.GetAttribute("xml:space") else xmlSpace
        xmlSpace = "preserve"

    let private readSharedStringText (reader: System.Xml.XmlReader) =
        let preserve = sharedStringTextPreserve reader
        if reader.IsEmptyElement then
            ""
        else
            let depth = reader.Depth
            let sb = StringBuilder()
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Text
                | System.Xml.XmlNodeType.CDATA
                | System.Xml.XmlNodeType.SignificantWhitespace
                | System.Xml.XmlNodeType.Whitespace ->
                    sb.Append(reader.Value) |> ignore
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
            let text = sb.ToString()
            decodeXmlEscapes (if preserve then text else text.Trim())

    let private parseSharedStringRunProps (reader: System.Xml.XmlReader) =
        let mutable bold = false
        let mutable italic = false
        let mutable underline = false
        let mutable underlineDouble = false
        let mutable strikethrough = false
        let mutable color = None
        let mutable fontName = None
        let mutable fontSize = None
        let mutable scheme = ""
        let mutable vertAlign = VertAlignBaseline
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "b" -> bold <- sharedStringBoolAttr reader
                    | "i" -> italic <- sharedStringBoolAttr reader
                    | "strike" -> strikethrough <- sharedStringBoolAttr reader
                    | "u" ->
                        underline <- sharedStringBoolAttr reader
                        match reader.GetAttribute("val") with
                        | "double" | "doubleAccounting" -> underlineDouble <- true
                        | _ -> ()
                    | "color" ->
                        match reader.GetAttribute("rgb") with
                        | null -> ()
                        | raw ->
                            let trimmed = raw.TrimStart('#')
                            color <- Some (if trimmed.Length = 8 then trimmed.Substring(2) else trimmed)
                    | "rFont" ->
                        match reader.GetAttribute("val") with
                        | null -> ()
                        | value -> fontName <- Some value
                    | "sz" ->
                        match reader.GetAttribute("val") with
                        | null -> ()
                        | value ->
                            match System.Double.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture) with
                            | true, parsed -> fontSize <- Some parsed
                            | false, _ -> ()
                    | "scheme" ->
                        match reader.GetAttribute("val") with
                        | null -> ()
                        | value -> scheme <- value
                    | "vertAlign" ->
                        match reader.GetAttribute("val") with
                        | "superscript" -> vertAlign <- VertAlignSuperscript
                        | "subscript" -> vertAlign <- VertAlignSubscript
                        | _ -> vertAlign <- VertAlignBaseline
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        {
            Bold = bold
            Italic = italic
            Underline = underline
            UnderlineDouble = underlineDouble
            Strikethrough = strikethrough
            Color = color
            FontName = fontName
            FontSize = fontSize
            Scheme = scheme
            VertAlign = vertAlign
        }

    let private parseSharedStringRun (reader: System.Xml.XmlReader) =
        let depth = reader.Depth
        let sb = StringBuilder()
        let mutable props = emptyRunProps
        if not reader.IsEmptyElement then
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "rPr" -> props <- parseSharedStringRunProps reader
                    | "t" -> sb.Append(readSharedStringText reader) |> ignore
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        sb.ToString(), props

    let private parseSharedStringItem (reader: System.Xml.XmlReader) =
        let depth = reader.Depth
        let plainText = StringBuilder()
        let runs = System.Collections.Generic.List<string * RunProps>()
        if not reader.IsEmptyElement then
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    if reader.Depth = depth + 1 && reader.LocalName = "r" then
                        runs.Add(parseSharedStringRun reader)
                    elif runs.Count = 0 && reader.Depth = depth + 1 && reader.LocalName = "t" then
                        plainText.Append(readSharedStringText reader) |> ignore
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        if runs.Count = 0 then
            { Text = plainText.ToString(); Runs = [] }
        else
            let sb = StringBuilder()
            for text, _ in runs do
                sb.Append(text) |> ignore
            { Text = sb.ToString(); Runs = runs |> Seq.toList }

    let private sharedStringsFast (sstXml: string) =
        let settings = System.Xml.XmlReaderSettings()
        settings.DtdProcessing <- System.Xml.DtdProcessing.Prohibit
        settings.XmlResolver <- null
        use stringReader = new System.IO.StringReader(sstXml)
        use reader = System.Xml.XmlReader.Create(stringReader, settings)
        let entries = System.Collections.Generic.List<SharedStringEntry>()
        while reader.Read() do
            if reader.NodeType = System.Xml.XmlNodeType.Element && reader.LocalName = "si" then
                entries.Add(parseSharedStringItem reader)
        entries.ToArray()
#endif

    let private cellValueFromParts (cellType: string option) (value: string option) (inlineText: string option) : CellValue =
        match cellType, value with
        | Some "inlineStr", _ ->
            inlineText |> Option.defaultValue "" |> String
        | Some "s", Some v ->
            // The index is an integer; tolerate "3.0" from sloppy writers and
            // treat anything unusable as an empty cell rather than throwing.
            match tryParseInt (v.Trim()) with
            | Some idx -> SharedString idx
            | None ->
                match tryParseFloat v with
                | Some n when n >= 0.0 && n < 2147483647.0 && n = System.Math.Floor n -> SharedString (int n)
                | _ -> Empty
        | Some "b", Some v -> Boolean (isTrueText (v.Trim()))
        | Some "str", Some v -> String (decodeXmlEscapes v)
        | Some "e", Some v ->
            // Error cell: <c t="e"><v>#DIV/0!</v></c>. Strip the leading
            // `#` so downstream `Error` rendering doesn't double up the
            // prefix (the Paint pass adds it via `sprintf "#%s" e`).
            let stripped = if v.StartsWith "#" then v.Substring 1 else v
            CellValue.Error stripped
        | _, Some v ->
            match tryParseFloat v with
            | Some n -> Number n
            | None -> if v = "" then Empty else String v
        | _ ->
            // No <v> — but an <is>/<t> might still be present even if the
            // cell omits an explicit t="inlineStr".
            match inlineText with
            | Some t when t <> "" -> String t
            | _ -> Empty

    let private isAsciiLetter (c: char) = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z')
    let private isAsciiDigit (c: char) = c >= '0' && c <= '9'
    /// Characters that continue a name / function / sheet identifier.
    let private isNameChar (c: char) =
        isAsciiLetter c || isAsciiDigit c || c = '_' || c = '.' || c = '\\' || c > '\127'

    /// Rewrite a formula for a cell `dRow` rows and `dCol` columns away from
    /// where it was written, the way Excel fills a shared formula: relative
    /// parts of A1 references, whole-column (`A:C`) and whole-row (`1:3`)
    /// ranges move, `$`-anchored parts stay. String literals, quoted sheet
    /// names, structured references and function names are left alone. A
    /// reference pushed off the sheet becomes `#REF!`.
    let internal shiftFormulaRefs (formula: string) (dRow: int) (dCol: int) : string =
        if dRow = 0 && dCol = 0 then formula
        else
            let n = formula.Length
            let sb = StringBuilder(n + 8)
            let maxRow = 1048576
            let maxCol = 16384
            // `$`? + 1..3 letters at p -> (absolute, column index, next); index 0 = no match.
            let readCol (p: int) : bool * int * int =
                let mutable q = p
                let abs = q < n && formula.[q] = '$'
                if abs then q <- q + 1
                let start = q
                let mutable col = 0
                while q < n && isAsciiLetter formula.[q] && q - start < 4 do
                    col <- col * 26 + (int (System.Char.ToUpperInvariant formula.[q]) - int 'A' + 1)
                    q <- q + 1
                if q = start || q - start > 3 || col > maxCol then abs, 0, p else abs, col, q
            // `$`? + digits at p -> (absolute, row, next); row 0 = no match.
            let readRow (p: int) : bool * int * int =
                let mutable q = p
                let abs = q < n && formula.[q] = '$'
                if abs then q <- q + 1
                let start = q
                let mutable row = 0
                while q < n && isAsciiDigit formula.[q] && q - start < 8 do
                    row <- row * 10 + (int formula.[q] - int '0')
                    q <- q + 1
                if q = start || q - start > 7 || row < 1 || row > maxRow then abs, 0, p else abs, row, q
            // A reference must not run into more identifier text, a call
            // (`LOG10(`) or a sheet qualifier (`AB1!`).
            let endsToken (p: int) =
                p >= n || not (isNameChar formula.[p] || formula.[p] = '(' || formula.[p] = '!')
            let appendCol (abs: bool) (col: int) =
                if abs then sb.Append('$') |> ignore
                sb.Append(indexToColumnLetters col) |> ignore
            let appendRow (abs: bool) (row: int) =
                if abs then sb.Append('$') |> ignore
                sb.Append(row) |> ignore
            let mutable i = 0
            // `1:3` / `$1:$3` at p: emit shifted and advance, or leave untouched.
            let tryRowRange (p: int) : bool =
                let a1, r1, q1 = readRow p
                if r1 > 0 && q1 < n && formula.[q1] = ':' then
                    let a2, r2, q2 = readRow (q1 + 1)
                    if r2 > 0 && endsToken q2 then
                        let n1 = if a1 then r1 else r1 + dRow
                        let n2 = if a2 then r2 else r2 + dRow
                        if n1 < 1 || n1 > maxRow || n2 < 1 || n2 > maxRow then
                            sb.Append("#REF!") |> ignore
                        else
                            appendRow a1 n1
                            sb.Append(':') |> ignore
                            appendRow a2 n2
                        i <- q2
                        true
                    else false
                else false
            while i < n do
                let ch = formula.[i]
                if ch = '"' || ch = '\'' then
                    // String literal or quoted sheet name; a doubled quote is an escape.
                    sb.Append(ch) |> ignore
                    i <- i + 1
                    let mutable closed = false
                    while not closed && i < n do
                        let c = formula.[i]
                        sb.Append(c) |> ignore
                        i <- i + 1
                        if c = ch then
                            if i < n && formula.[i] = ch then
                                sb.Append(ch) |> ignore
                                i <- i + 1
                            else
                                closed <- true
                elif ch = '[' then
                    // Structured reference / external-workbook index: copy to the matching bracket.
                    let mutable depth = 0
                    let mutable finished = false
                    while not finished && i < n do
                        let c = formula.[i]
                        if c = '[' then depth <- depth + 1
                        elif c = ']' then depth <- depth - 1
                        sb.Append(c) |> ignore
                        i <- i + 1
                        if depth = 0 then finished <- true
                elif ch = '$' || isAsciiLetter ch then
                    let cAbs, col, afterCol = readCol i
                    let mutable handled = false
                    if col > 0 then
                        let rAbs, row, afterRow = readRow afterCol
                        if row > 0 && endsToken afterRow then
                            let c2 = if cAbs then col else col + dCol
                            let r2 = if rAbs then row else row + dRow
                            if c2 < 1 || c2 > maxCol || r2 < 1 || r2 > maxRow then
                                sb.Append("#REF!") |> ignore
                            else
                                appendCol cAbs c2
                                appendRow rAbs r2
                            i <- afterRow
                            handled <- true
                        elif row = 0 && afterCol < n && formula.[afterCol] = ':' then
                            let cAbs2, col2, afterCol2 = readCol (afterCol + 1)
                            if col2 > 0 && endsToken afterCol2 then
                                let n1 = if cAbs then col else col + dCol
                                let n2 = if cAbs2 then col2 else col2 + dCol
                                if n1 < 1 || n1 > maxCol || n2 < 1 || n2 > maxCol then
                                    sb.Append("#REF!") |> ignore
                                else
                                    appendCol cAbs n1
                                    sb.Append(':') |> ignore
                                    appendCol cAbs2 n2
                                i <- afterCol2
                                handled <- true
                    elif ch = '$' then
                        handled <- tryRowRange i
                    if not handled then
                        if ch = '$' then
                            sb.Append(ch) |> ignore
                            i <- i + 1
                        else
                            // Function / defined name / sheet name: copy the whole run.
                            while i < n && isNameChar formula.[i] do
                                sb.Append(formula.[i]) |> ignore
                                i <- i + 1
                elif isAsciiDigit ch then
                    if not (tryRowRange i) then
                        while i < n && isNameChar formula.[i] do
                            sb.Append(formula.[i]) |> ignore
                            i <- i + 1
                else
                    sb.Append(ch) |> ignore
                    i <- i + 1
            sb.ToString()

    /// True when a number-format code renders a date or time — i.e. the
    /// number under it is a date serial. Literal text ("..."), escaped
    /// characters and bracketed sections ([Red], [$-409]) are ignored, except
    /// the elapsed-time brackets [h] [m] [s].
    let internal isDateFormatCode (code: string) : bool =
        if code = "" || code = "@" || code.Equals("General", System.StringComparison.OrdinalIgnoreCase) then false
        else
            let mutable found = false
            let mutable i = 0
            let n = code.Length
            while not found && i < n do
                let ch = code.[i]
                if ch = '"' then
                    i <- i + 1
                    while i < n && code.[i] <> '"' do i <- i + 1
                    i <- i + 1
                elif ch = '\\' || ch = '_' || ch = '*' then
                    i <- i + 2
                elif ch = '[' then
                    let close = code.IndexOf(']', i)
                    let inner = if close > i then code.Substring(i + 1, close - i - 1).ToLowerInvariant() else ""
                    if inner <> "" && (inner |> Seq.forall (fun c -> c = 'h') || inner |> Seq.forall (fun c -> c = 'm') || inner |> Seq.forall (fun c -> c = 's')) then
                        found <- true
                    i <- if close > i then close + 1 else n
                else
                    match System.Char.ToLowerInvariant ch with
                    | 'd' | 'm' | 'y' | 'h' | 's' -> found <- true
                    | 'e' ->
                        // `E+00` is scientific notation; a bare `e` is the era year.
                        if i + 1 < n && (code.[i + 1] = '+' || code.[i + 1] = '-') then i <- i + 1
                        else found <- true
                    | _ -> ()
                    i <- i + 1
            found

    /// State one worksheet parse carries from cell to cell. Shared by the
    /// DOM and the streaming parser so both resolve cells identically.
    type internal SheetParseContext = {
        Styles: CellStyle array option
        /// Per cellXf index: the number under this style is a 1904-system
        /// date serial to move into the 1900 system. Empty for 1900 workbooks.
        DateShift: bool array
        /// Shared-formula masters seen so far: `si` -> (row, column, text).
        SharedFormulas: System.Collections.Generic.Dictionary<string, int * int * string>
    }

    /// Days between the 1904 and 1900 date systems' epochs.
    let internal date1904Offset = 1462.0

    let internal newSheetContext (styles: CellStyle array option) (date1904: bool) : SheetParseContext =
        {
            Styles = styles
            DateShift =
                match styles, date1904 with
                | Some styles, true ->
                    styles |> Array.map (fun style -> style.NumberFormat |> Option.exists isDateFormatCode)
                | _ -> [||]
            SharedFormulas = System.Collections.Generic.Dictionary<string, int * int * string>()
        }

    let private resolvedStyle (styles: CellStyle array option) styleIdx =
        match styles, styleIdx with
        | Some styles, Some idx when idx >= 0 && idx < styles.Length -> Some styles.[idx]
        | Some styles, None when styles.Length > 0 -> Some styles.[0]
        | _ -> None

    /// Assemble a cell from the pieces either parser extracted. Excel stores
    /// both the formula and its last computed result; the model keeps both.
    let private buildCell (ctx: SheetParseContext) (column: string) (colIndex: int) (rowIndex: int)
                          (cellType: string option) (styleIdx: int option)
                          (fText: string option) (fType: string option) (fRef: string option) (fSi: string option)
                          (value: string option) (inlineText: string option) (inlineRuns: (string * RunProps) list) : Cell =
        let formula, arrayRef =
            match fText with
            | None -> None, None
            | Some text ->
                match fType with
                | Some "shared" ->
                    if text <> "" then
                        // Master cell of a shared group: later cells reuse its text.
                        match fSi with
                        | Some si -> ctx.SharedFormulas.[si] <- (rowIndex, colIndex, text)
                        | None -> ()
                        Some text, None
                    else
                        match fSi with
                        | Some si ->
                            match ctx.SharedFormulas.TryGetValue si with
                            | true, (masterRow, masterCol, masterText) ->
                                Some (shiftFormulaRefs masterText (rowIndex - masterRow) (colIndex - masterCol)), None
                            | _ -> None, None
                        | None -> None, None
                | Some "array" -> (if text <> "" then Some text else None), fRef
                // A what-if data table: its "formula" is a pair of input-cell
                // references, not an expression. Keep the cached result only.
                | Some "dataTable" -> None, None
                | _ -> (if text <> "" then Some text else None), None
        let cached =
            match cellValueFromParts cellType value inlineText with
            | Number n ->
                let idx = styleIdx |> Option.defaultValue 0
                if idx >= 0 && idx < ctx.DateShift.Length && ctx.DateShift.[idx] then Number (n + date1904Offset)
                else Number n
            | other -> other
        {
            Reference = { Column = column; Row = rowIndex }
            Value =
                match cached, formula with
                | Empty, Some f -> Formula f
                | _ -> cached
            Formula = formula
            ArrayRef = arrayRef
            StyleIndex = styleIdx
            Style = resolvedStyle ctx.Styles styleIdx
            InlineRuns =
                match cellType with
                | Some "inlineStr" -> inlineRuns
                | _ -> []
        }

    // Cell value extraction with pattern matching. Returns the cached
    // result when there is one, the formula text otherwise (uncalculated
    // workbooks); the `Cell` record carries both.
    let cellValue (cNode: XmlNode) : CellValue =
        let cellType = Xml.attr "t" cNode
        let value = sElem "v" cNode |> Option.bind Xml.text
        let inlineText =
            sElem "is" cNode
            |> Option.map (descendantTextConcat "t")
        match cellValueFromParts cellType value inlineText, sElem "f" cNode |> Option.bind Xml.text with
        | Empty, Some f when f <> "" -> Formula f
        | cached, _ -> cached

    /// Parse one `<c>`. `r` is optional (streaming writers omit it): a cell
    /// without it sits one column right of the previous cell in its row.
    /// Returns the cell and its 1-based column index.
    let private cellFromNode (ctx: SheetParseContext) (rowIndex: int) (prevCol: int) (cNode: XmlNode) : Cell * int =
        let column, colIndex, row =
            let inferred () = indexToColumnLetters (prevCol + 1), prevCol + 1, rowIndex
            match Xml.attr "r" cNode with
            | Some ref ->
                let letters, parsedRow = parseCellRefParts ref
                if letters = "" then inferred ()
                else letters, columnLettersToIndex letters, (if parsedRow > 0 then parsedRow else rowIndex)
            | None -> inferred ()
        let cellType = Xml.attr "t" cNode
        let styleIdx = Xml.attr "s" cNode |> Option.bind tryParseInt
        let fNode = sElem "f" cNode
        let fText = fNode |> Option.map (fun f -> Xml.text f |> Option.defaultValue "")
        let value = sElem "v" cNode |> Option.bind Xml.text
        // `t="inlineStr"` cells carry their text directly inside <is><t>...</t></is>
        // instead of pointing at a sharedStrings index. Rare in normal
        // Excel exports but emitted by various code generators (POI, the
        // .NET OpenXmlSdk's row-streaming writer, sql-server-style mass
        // importers) where deduping into sharedStrings isn't worth the cost.
        let isNode = sElem "is" cNode
        let inlineText = isNode |> Option.map (descendantTextConcat "t")
        // Inline-string cells with <r> children carry per-run formatting
        // identical in shape to sharedStrings rich text.
        let inlineRuns =
            match cellType with
            | Some "inlineStr" -> isNode |> Option.map richRunsOf |> Option.defaultValue []
            | _ -> []
        let cell =
            buildCell ctx column colIndex row cellType styleIdx
                fText (fNode |> Option.bind (Xml.attr "t")) (fNode |> Option.bind (Xml.attr "ref")) (fNode |> Option.bind (Xml.attr "si"))
                value inlineText inlineRuns
        cell, colIndex

    // Cell extraction with applicative style
    let cellWithStyles (styles: CellStyle array option) (cNode: XmlNode) : Cell option =
        Xml.attr "r" cNode
        |> Option.map (fun _ -> cellFromNode (newSheetContext styles false) 0 0 cNode |> fst)

    let cell cNode = cellWithStyles None cNode

    /// Cells a `<row s=… customFormat="1" spans="a:b">` implies but does not
    /// list: the row style (e.g. a fill bar) applies to every cell across the
    /// span, so synthesize the missing ones (empty value, row style) — e.g. a
    /// full-width title bar where only A1 carries text.
    let private rowSpanCells (ctx: SheetParseContext) (rowIndex: int) (rowStyle: int option) (customFormat: bool) (spans: string option) (explicitCells: Cell list) : Cell list =
        match (if customFormat then rowStyle else None), spans with
        | Some s, Some spans ->
            match spans.Split(':') with
            | [| a; b |] ->
                match tryParseInt a, tryParseInt b with
                | Some c1, Some c2 when c2 >= c1 && c2 - c1 <= 256 ->
                    let present = explicitCells |> List.map (fun c -> c.Reference.Column) |> Set.ofList
                    let synth =
                        [ for col in c1 .. c2 do
                            let letters = indexToColumnLetters col
                            if not (present.Contains letters) then
                                yield { Reference = { Column = letters; Row = rowIndex }
                                        Value = Empty
                                        Formula = None
                                        ArrayRef = None
                                        StyleIndex = Some s
                                        Style = resolvedStyle ctx.Styles (Some s)
                                        InlineRuns = [] } ]
                    explicitCells @ synth
                | _ -> explicitCells
            | _ -> explicitCells
        | _ -> explicitCells

    /// Parse one `<row>`. `r` is optional: a row without it follows the
    /// previous row.
    let private rowFromNode (ctx: SheetParseContext) (prevRow: int) (rowNode: XmlNode) : Row =
        let rowIndex =
            match Xml.attr "r" rowNode |> Option.bind tryParseInt with
            | Some r when r > 0 -> r
            | _ -> prevRow + 1
        let customHeightAttr = Xml.attr "customHeight" rowNode |> Option.map isTrueText
        // OOXML §18.3.1.73 default for customHeight is FALSE. When
        // absent, the row's stored `ht` is just a hint Excel may
        // override with the font-derived auto height. Only honor `ht`
        // as the row's final height when customHeight is explicitly
        // set to true.
        let customHeight = customHeightAttr |> Option.defaultValue false
        let rawHeight = Xml.attr "ht" rowNode |> Option.bind tryParseFloat
        let height = if customHeight then rawHeight else None
        let hidden = Xml.attr "hidden" rowNode |> Option.map isTrueText |> Option.defaultValue false
        let mutable prevCol = 0
        let explicitCells =
            childElems "c" rowNode
            |> List.map (fun cNode ->
                let cell, colIndex = cellFromNode ctx rowIndex prevCol cNode
                prevCol <- colIndex
                cell)
        let rowStyle = Xml.attr "s" rowNode |> Option.bind tryParseInt
        let customFormat = Xml.attr "customFormat" rowNode |> Option.map isTrueText |> Option.defaultValue false
        {
            Index = rowIndex
            Cells = rowSpanCells ctx rowIndex rowStyle customFormat (Xml.attr "spans" rowNode) explicitCells
            Height = height
            RawHeight = rawHeight
            CustomHeight = customHeightAttr
            Hidden = hidden
        }

    // Row extraction
    let rowWithStyles (styles: CellStyle array option) (rowNode: XmlNode) : Row option =
        Xml.attr "r" rowNode
        |> Option.map (fun _ -> rowFromNode (newSheetContext styles false) 0 rowNode)

    let row rowNode = rowWithStyles None rowNode

    // Merged cells extraction
    let parseMergeCell (ref: string) : MergeCell option =
        match ref.Split(':') with
        | [| start; finish |] ->
            Some {
                StartRef = parseCellRef start
                EndRef = parseCellRef finish
            }
        | _ -> None

    let mergeCells (sheet: XmlNode) : MergeCell list =
        sElem "mergeCells" sheet
        |> Option.map (sElems "mergeCell")
        |> Option.defaultValue []
        |> List.choose (fun mc ->
            Xml.attr "ref" mc
            |> Option.bind parseMergeCell)

    // Column info extraction. <col min="4" max="7"/> describes columns 4..7
    // sharing the same width/hidden setting, so we keep both ends of the
    // range so the runner can fan the width out to every column in it.
    let columnInfo (colNode: XmlNode) : ColumnInfo option =
        Xml.attr "min" colNode
        |> Option.bind tryParseInt
        |> Option.map (fun minIdx ->
            let maxIdx =
                Xml.attr "max" colNode
                |> Option.bind tryParseInt
                |> Option.defaultValue minIdx
            let width = Xml.attr "width" colNode |> Option.bind tryParseFloat
            let hidden = Xml.attr "hidden" colNode |> Option.map isTrueText |> Option.defaultValue false
            let style = Xml.attr "style" colNode |> Option.bind tryParseInt
            {
                Index = minIdx
                MaxIndex = maxIdx
                Width = width
                Hidden = hidden
                Style = style
            })

    let columns (sheet: XmlNode) : ColumnInfo list =
        sElem "cols" sheet
        |> Option.map (sElems "col")
        |> Option.defaultValue []
        |> List.choose columnInfo

    /// Excel uses sheetFormatPr defaults for any row/column without an
    /// explicit <row ht="..."/> or <col width="..."/> entry. Without this,
    /// renderers treat unspecified columns as if they had no defaults, which
    /// over-widens them and drifts every column past A.
    let emptySheetFormat : SheetFormat = {
        BaseColWidth = None
        DefaultColWidth = None
        DefaultRowHeight = None
        CustomDefaultHeight = false
    }

    let sheetFormatPr (sheet: XmlNode) : SheetFormat =
        match sElem "sheetFormatPr" sheet with
        | None -> emptySheetFormat
        | Some node ->
            {
                BaseColWidth = Xml.attr "baseColWidth" node |> Option.bind tryParseFloat
                DefaultColWidth = Xml.attr "defaultColWidth" node |> Option.bind tryParseFloat
                DefaultRowHeight = Xml.attr "defaultRowHeight" node |> Option.bind tryParseFloat
                CustomDefaultHeight =
                    Xml.attr "customHeight" node
                    |> Option.map (fun v -> v = "1" || v = "true")
                    |> Option.defaultValue false
            }

    // Worksheet extraction (advanced features parsed separately)
    /// Parse `<hyperlinks>` block: a list of `<hyperlink ref location
    /// r:id display/>` entries. External links use `r:id` to point at a
    /// Relationship in the sheet's rels file; internal links use the
    /// `location` attribute directly (e.g. `Sheet2!A1`). When the caller
    /// supplies a resolver function for r:id, external targets get
    /// filled in; otherwise we store the r:id verbatim (good enough for
    /// the "render as blue + underlined" use case which only cares that
    /// the cell IS a hyperlink, not where it points).
    let private hyperlinks (resolveRel: string -> string option) (sheet: XmlNode) : Hyperlink list =
        sElem "hyperlinks" sheet
        |> Option.map (sElems "hyperlink")
        |> Option.defaultValue []
        |> List.choose (fun node ->
            match Xml.attr "ref" node with
            | None -> None
            | Some range ->
                let location = Xml.attr "location" node
                let rid = relationshipId node
                let target, isInternal =
                    match location with
                    | Some loc -> loc, true
                    | None ->
                        match rid |> Option.bind resolveRel with
                        | Some url -> url, false
                        | None -> rid |> Option.defaultValue "", false
                Some { Range = range; Target = target; IsInternal = isInternal })

    /// `<sheetView showZeros="false"/>`. Excel's default is true; opt-out
    /// hides numeric-zero cells while keeping other content intact.
    let private showZeros (sheet: XmlNode) : bool =
        sElem "sheetViews" sheet
        |> Option.bind (sElem "sheetView")
        |> Option.bind (Xml.attr "showZeros")
        |> Option.map (fun v -> not (v = "0" || v = "false"))
        |> Option.defaultValue true

    /// `<sheetView showGridLines="0"/>`. Excel's default is true; polished
    /// workbooks turn it off and rely on explicit cell borders.
    let private showGridLines (sheet: XmlNode) : bool =
        sElem "sheetViews" sheet
        |> Option.bind (sElem "sheetView")
        |> Option.bind (Xml.attr "showGridLines")
        |> Option.map (fun v -> not (v = "0" || v = "false"))
        |> Option.defaultValue true

    /// Read `<sheetView><pane xSplit ySplit state="frozen"/></sheetView>`.
    /// Excel only renders the freeze divider when state is "frozen" (or
    /// "frozenSplit"); plain "split" puts a draggable bar in instead and
    /// the screenshots in our corpus never use it. Returns None when the
    /// pane is missing or in a non-frozen state.
    let private frozenPane (sheet: XmlNode) : FrozenPane option =
        sElem "sheetViews" sheet
        |> Option.bind (sElem "sheetView")
        |> Option.bind (sElem "pane")
        |> Option.bind (fun pane ->
            let state = Xml.attr "state" pane |> Option.defaultValue ""
            if state <> "frozen" && state <> "frozenSplit" then None
            else
                let parseSplit a =
                    Xml.attr a pane
                    |> Option.bind tryParseFloat
                    |> Option.map int
                    |> Option.defaultValue 0
                let x = parseSplit "xSplit"
                let y = parseSplit "ySplit"
                if x <= 0 && y <= 0 then None
                else Some { XSplit = x; YSplit = y })

    // ----- print-related parsers -----

    let private parseFloatAttr (node: XmlNode) (name: string) : float option =
        Xml.attr name node |> Option.bind tryParseFloat

    let private parseIntAttr (node: XmlNode) (name: string) : int option =
        Xml.attr name node |> Option.bind tryParseInt

    let private parseBoolAttr (node: XmlNode) (name: string) : bool =
        match Xml.attr name node with
        | Some "1" | Some "true" -> true
        | _ -> false

    /// `<pageMargins left right top bottom header footer/>` — all in inches.
    let private pageMargins (sheet: XmlNode) : PageMargins option =
        sElem "pageMargins" sheet
        |> Option.map (fun m ->
            let get a = parseFloatAttr m a |> Option.defaultValue 0.0
            { Top = get "top"
              Bottom = get "bottom"
              Left = get "left"
              Right = get "right"
              Header = get "header"
              Footer = get "footer" })

    /// `<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`.
    let private fitToPageFlag (sheet: XmlNode) : bool =
        sElem "sheetPr" sheet
        |> Option.bind (sElem "pageSetUpPr")
        |> Option.map (fun n -> parseBoolAttr n "fitToPage")
        |> Option.defaultValue false

    /// `<pageSetup paperSize orientation fitToWidth fitToHeight scale .../>`
    /// plus the fitToPage flag from sheetPr.
    let private pageSetup (sheet: XmlNode) : PageSetup option =
        let elem = sElem "pageSetup" sheet
        let fit = fitToPageFlag sheet
        match elem, fit with
        | None, false -> None
        | _ ->
            let n = elem
            let orientation =
                n |> Option.bind (Xml.attr "orientation") |> function
                | Some "portrait" -> OrientationPortrait
                | Some "landscape" -> OrientationLandscape
                | _ -> OrientationDefault
            let intOr fallback name = n |> Option.bind (fun e -> parseIntAttr e name) |> Option.defaultValue fallback
            let useFirst = n |> Option.map (fun e -> parseBoolAttr e "useFirstPageNumber") |> Option.defaultValue false
            let pageOrder =
                n |> Option.bind (Xml.attr "pageOrder") |> function
                | Some "overThenDown" -> OverThenDown
                | _ -> DownThenOver
            Some {
                PaperSize = n |> Option.bind (fun e -> parseIntAttr e "paperSize")
                Orientation = orientation
                FitToWidth = intOr 1 "fitToWidth"
                FitToHeight = intOr 1 "fitToHeight"
                Scale = intOr 100 "scale"
                FitToPage = fit
                FirstPageNumber = intOr 1 "firstPageNumber"
                UseFirstPageNumber = useFirst
                PageOrder = pageOrder
                BlackAndWhite = n |> Option.map (fun e -> parseBoolAttr e "blackAndWhite") |> Option.defaultValue false
                CellComments =
                    n |> Option.bind (Xml.attr "cellComments") |> function
                    | Some "asDisplayed" -> CellCommentsAsDisplayed
                    | Some "atEnd" -> CellCommentsAtEnd
                    | _ -> CellCommentsNone
            }

    /// `<printOptions gridLines headings horizontalCentered verticalCentered/>`.
    let private printOptions (sheet: XmlNode) : PrintOptions =
        match sElem "printOptions" sheet with
        | None -> defaultPrintOptions
        | Some n ->
            { GridLines = parseBoolAttr n "gridLines"
              Headings = parseBoolAttr n "headings"
              HorizontalCentered = parseBoolAttr n "horizontalCentered"
              VerticalCentered = parseBoolAttr n "verticalCentered" }

    /// `<headerFooter differentOddEven="1" differentFirst="1"><oddHeader>…</oddHeader>…</headerFooter>`.
    let private headerFooter (sheet: XmlNode) : HeaderFooter =
        match sElem "headerFooter" sheet with
        | None -> defaultHeaderFooter
        | Some hf ->
            let part name = sElem name hf |> Option.bind Xml.text
            { OddHeader = part "oddHeader"
              OddFooter = part "oddFooter"
              EvenHeader = part "evenHeader"
              EvenFooter = part "evenFooter"
              FirstHeader = part "firstHeader"
              FirstFooter = part "firstFooter"
              DifferentOddEven = parseBoolAttr hf "differentOddEven"
              DifferentFirst = parseBoolAttr hf "differentFirst"
              // `scaleWithDoc` defaults to true when absent — matches Excel.
              ScaleWithDoc =
                  Xml.attr "scaleWithDoc" hf
                  |> Option.map (fun v -> v = "1" || v = "true")
                  |> Option.defaultValue true }

    /// `<rowBreaks count="N"><brk id="R" max="C" man="1"/>…</rowBreaks>`.
    /// We only collect manual breaks (man="1"); automatic breaks that Excel
    /// emitted are reconstructed by pagination on our side.
    let private breaks (container: string) (sheet: XmlNode) : int list =
        sElem container sheet
        |> Option.map (fun node ->
            Ns.descendantsNs "brk" node
            |> List.choose (fun b ->
                if parseBoolAttr b "man" then parseIntAttr b "id" else None)
            |> List.filter (fun i -> i > 0)
            |> List.sort)
        |> Option.defaultValue []

#if !FABLE_COMPILER
    let private worksheetFastNativeThreshold = 1_000_000
    let private relationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    let private resizeListToList (items: System.Collections.Generic.List<'T>) =
        let mutable result = []
        for i = items.Count - 1 downto 0 do
            result <- items.[i] :: result
        result

    let private nativeAttr (reader: System.Xml.XmlReader) (name: string) =
        match reader.GetAttribute(name) with
        | null -> None
        | value -> Some value

    let private strictRelationshipsNamespace = "http://purl.oclc.org/ooxml/officeDocument/relationships"

    /// The element's relationship id, looked up by namespace so any prefix
    /// (and the Strict-conformance namespace) works.
    let private nativeRelId (reader: System.Xml.XmlReader) =
        match reader.GetAttribute("id", relationshipsNamespace) with
        | null ->
            match reader.GetAttribute("id", strictRelationshipsNamespace) with
            | null -> nativeAttr reader "r:id"
            | value -> Some value
        | value -> Some value

    let private tryParseNativeInt (value: string) = tryParseInt value

    let private tryParseNativeFloat (value: string) = tryParseFloat value

    let private nativeIntAttr reader name =
        nativeAttr reader name |> Option.bind tryParseNativeInt

    let private nativeFloatAttr reader name =
        nativeAttr reader name |> Option.bind tryParseNativeFloat

    let private nativeBoolValue (value: string) =
        value = "1" || System.String.Equals(value, "true", System.StringComparison.OrdinalIgnoreCase)

    let private nativeBoolAttr reader name =
        nativeAttr reader name
        |> Option.map nativeBoolValue
        |> Option.defaultValue false

    let private nativeOptionalBoolAttr reader name =
        nativeAttr reader name |> Option.map nativeBoolValue

    let private nativeBoolAttrDefault fallback reader name =
        nativeAttr reader name
        |> Option.map nativeBoolValue
        |> Option.defaultValue fallback

    let private readElementTextNative (reader: System.Xml.XmlReader) =
        if reader.IsEmptyElement then
            ""
        else
            let depth = reader.Depth
            let sb = StringBuilder()
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Text
                | System.Xml.XmlNodeType.CDATA
                | System.Xml.XmlNodeType.SignificantWhitespace
                | System.Xml.XmlNodeType.Whitespace ->
                    sb.Append(reader.Value) |> ignore
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
            sb.ToString()

    let private nativeColumnName (columnsByIndex: System.Collections.Generic.Dictionary<int, string>) index =
        match columnsByIndex.TryGetValue index with
        | true, value -> value
        | false, _ ->
            let value = indexToColumnLetters index
            columnsByIndex.[index] <- value
            value

    /// Split an A1 reference into (1-based column index, row); 0 = absent.
    let private cellPositionNative (value: string) : int * int =
        let mutable index = 0
        while index < value.Length && value.[index] = '$' do
            index <- index + 1

        let mutable colIndex = 0
        while index < value.Length && System.Char.IsLetter value.[index] do
            let ch = System.Char.ToUpperInvariant value.[index]
            colIndex <- colIndex * 26 + (int ch - int 'A' + 1)
            index <- index + 1

        while index < value.Length && value.[index] = '$' do
            index <- index + 1

        let mutable row = 0
        while index < value.Length && System.Char.IsDigit value.[index] do
            row <- row * 10 + int value.[index] - int '0'
            index <- index + 1

        colIndex, row

    let private inlineStringNative (reader: System.Xml.XmlReader) =
        let depth = reader.Depth
        let plainText = StringBuilder()
        let runs = System.Collections.Generic.List<string * RunProps>()
        if not reader.IsEmptyElement then
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "r" -> runs.Add(parseSharedStringRun reader)
                    | "t" when runs.Count = 0 ->
                        plainText.Append(readSharedStringText reader) |> ignore
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        if runs.Count = 0 then
            plainText.ToString(), []
        else
            let sb = StringBuilder()
            for text, _ in runs do
                sb.Append(text) |> ignore
            sb.ToString(), resizeListToList runs

    /// Streaming twin of `cellFromNode`: same inference for a missing `r`,
    /// same `buildCell`.
    let private cellNative (ctx: SheetParseContext) (columnsByIndex: System.Collections.Generic.Dictionary<int, string>) (rowIndex: int) (prevCol: int) (reader: System.Xml.XmlReader) : Cell * int =
        let colIndex, row =
            match reader.GetAttribute("r") with
            | null -> prevCol + 1, rowIndex
            | ref ->
                let col, parsedRow = cellPositionNative ref
                if col = 0 then prevCol + 1, rowIndex
                else col, (if parsedRow > 0 then parsedRow else rowIndex)
        let cellType = nativeAttr reader "t"
        let styleIdx = nativeIntAttr reader "s"
        let mutable fText = None
        let mutable fType = None
        let mutable fRef = None
        let mutable fSi = None
        let mutable value = None
        let mutable inlineText = None
        let mutable inlineRuns = []
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "f" ->
                        fType <- nativeAttr reader "t"
                        fRef <- nativeAttr reader "ref"
                        fSi <- nativeAttr reader "si"
                        fText <- Some (readElementTextNative reader)
                    | "v" -> value <- Some (readElementTextNative reader)
                    | "is" ->
                        let text, runs = inlineStringNative reader
                        inlineText <- Some text
                        inlineRuns <- runs
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        let cell =
            buildCell ctx (nativeColumnName columnsByIndex colIndex) colIndex row cellType styleIdx
                fText fType fRef fSi value inlineText inlineRuns
        cell, colIndex

    /// Streaming twin of `rowFromNode`.
    let private rowNative (ctx: SheetParseContext) (columnsByIndex: System.Collections.Generic.Dictionary<int, string>) (prevRow: int) (reader: System.Xml.XmlReader) : Row =
        let rowIndex =
            match nativeIntAttr reader "r" with
            | Some r when r > 0 -> r
            | _ -> prevRow + 1
        let customHeightAttr = nativeOptionalBoolAttr reader "customHeight"
        let customHeight = customHeightAttr |> Option.defaultValue false
        let rawHeight = nativeFloatAttr reader "ht"
        let height = if customHeight then rawHeight else None
        let hidden = nativeBoolAttr reader "hidden"
        let rowStyle = nativeIntAttr reader "s"
        let customFormat = nativeBoolAttr reader "customFormat"
        let spans = nativeAttr reader "spans"
        let explicitCells = System.Collections.Generic.List<Cell>()
        let mutable prevCol = 0
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element when reader.LocalName = "c" ->
                    let cell, colIndex = cellNative ctx columnsByIndex rowIndex prevCol reader
                    prevCol <- colIndex
                    explicitCells.Add(cell)
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        {
            Index = rowIndex
            Cells = rowSpanCells ctx rowIndex rowStyle customFormat spans (resizeListToList explicitCells)
            Height = height
            RawHeight = rawHeight
            CustomHeight = customHeightAttr
            Hidden = hidden
        }

    let private columnInfoNative (reader: System.Xml.XmlReader) : ColumnInfo option =
        nativeIntAttr reader "min"
        |> Option.map (fun minIdx ->
            let maxIdx = nativeIntAttr reader "max" |> Option.defaultValue minIdx
            {
                Index = minIdx
                MaxIndex = maxIdx
                Width = nativeFloatAttr reader "width"
                Hidden = nativeBoolAttr reader "hidden"
                Style = nativeIntAttr reader "style"
            })

    let private sheetFormatPrNative (reader: System.Xml.XmlReader) : SheetFormat =
        {
            BaseColWidth = nativeFloatAttr reader "baseColWidth"
            DefaultColWidth = nativeFloatAttr reader "defaultColWidth"
            DefaultRowHeight = nativeFloatAttr reader "defaultRowHeight"
            CustomDefaultHeight = nativeBoolAttr reader "customHeight"
        }

    let private sheetViewsNative (reader: System.Xml.XmlReader) =
        let mutable showZerosValue = true
        let mutable showGridLinesValue = true
        let mutable pane = None
        let mutable sheetViewDepth = -1
        let mutable sawSheetView = false
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element when reader.LocalName = "sheetView" && not sawSheetView ->
                    sawSheetView <- true
                    showZerosValue <- nativeBoolAttrDefault true reader "showZeros"
                    showGridLinesValue <- nativeBoolAttrDefault true reader "showGridLines"
                    if not reader.IsEmptyElement then
                        sheetViewDepth <- reader.Depth
                | System.Xml.XmlNodeType.Element when reader.LocalName = "pane" && sheetViewDepth >= 0 ->
                    let state = nativeAttr reader "state" |> Option.defaultValue ""
                    if state = "frozen" || state = "frozenSplit" then
                        let parseSplit name =
                            nativeFloatAttr reader name
                            |> Option.map int
                            |> Option.defaultValue 0
                        let x = parseSplit "xSplit"
                        let y = parseSplit "ySplit"
                        if x > 0 || y > 0 then
                            pane <- Some { XSplit = x; YSplit = y }
                | System.Xml.XmlNodeType.EndElement when reader.Depth = sheetViewDepth ->
                    sheetViewDepth <- -1
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        showZerosValue, showGridLinesValue, pane

    let private fitToPageNative (reader: System.Xml.XmlReader) =
        let mutable fit = false
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element when reader.LocalName = "pageSetUpPr" ->
                    fit <- nativeBoolAttr reader "fitToPage"
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        fit

    let private defaultPageSetupNative fit =
        {
            PaperSize = None
            Orientation = OrientationDefault
            FitToWidth = 1
            FitToHeight = 1
            Scale = 100
            FitToPage = fit
            FirstPageNumber = 1
            UseFirstPageNumber = false
            PageOrder = DownThenOver
            BlackAndWhite = false
            CellComments = CellCommentsNone
        }

    let private pageSetupNative fit (reader: System.Xml.XmlReader) =
        let orientation =
            nativeAttr reader "orientation"
            |> function
                | Some "portrait" -> OrientationPortrait
                | Some "landscape" -> OrientationLandscape
                | _ -> OrientationDefault
        {
            PaperSize = nativeIntAttr reader "paperSize"
            Orientation = orientation
            FitToWidth = nativeIntAttr reader "fitToWidth" |> Option.defaultValue 1
            FitToHeight = nativeIntAttr reader "fitToHeight" |> Option.defaultValue 1
            Scale = nativeIntAttr reader "scale" |> Option.defaultValue 100
            FitToPage = fit
            FirstPageNumber = nativeIntAttr reader "firstPageNumber" |> Option.defaultValue 1
            UseFirstPageNumber = nativeBoolAttr reader "useFirstPageNumber"
            PageOrder =
                nativeAttr reader "pageOrder"
                |> function
                    | Some "overThenDown" -> OverThenDown
                    | _ -> DownThenOver
            BlackAndWhite = nativeBoolAttr reader "blackAndWhite"
            CellComments =
                nativeAttr reader "cellComments"
                |> function
                    | Some "asDisplayed" -> CellCommentsAsDisplayed
                    | Some "atEnd" -> CellCommentsAtEnd
                    | _ -> CellCommentsNone
        }

    let private pageMarginsNative (reader: System.Xml.XmlReader) =
        let get name = nativeFloatAttr reader name |> Option.defaultValue 0.0
        {
            Top = get "top"
            Bottom = get "bottom"
            Left = get "left"
            Right = get "right"
            Header = get "header"
            Footer = get "footer"
        }

    let private printOptionsNative (reader: System.Xml.XmlReader) =
        {
            GridLines = nativeBoolAttr reader "gridLines"
            Headings = nativeBoolAttr reader "headings"
            HorizontalCentered = nativeBoolAttr reader "horizontalCentered"
            VerticalCentered = nativeBoolAttr reader "verticalCentered"
        }

    let private headerFooterNative (reader: System.Xml.XmlReader) =
        let mutable oddHeader = None
        let mutable oddFooter = None
        let mutable evenHeader = None
        let mutable evenFooter = None
        let mutable firstHeader = None
        let mutable firstFooter = None
        let differentOddEven = nativeBoolAttr reader "differentOddEven"
        let differentFirst = nativeBoolAttr reader "differentFirst"
        let scaleWithDoc = nativeBoolAttrDefault true reader "scaleWithDoc"
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "oddHeader" -> oddHeader <- Some (readElementTextNative reader)
                    | "oddFooter" -> oddFooter <- Some (readElementTextNative reader)
                    | "evenHeader" -> evenHeader <- Some (readElementTextNative reader)
                    | "evenFooter" -> evenFooter <- Some (readElementTextNative reader)
                    | "firstHeader" -> firstHeader <- Some (readElementTextNative reader)
                    | "firstFooter" -> firstFooter <- Some (readElementTextNative reader)
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        {
            OddHeader = oddHeader
            OddFooter = oddFooter
            EvenHeader = evenHeader
            EvenFooter = evenFooter
            FirstHeader = firstHeader
            FirstFooter = firstFooter
            DifferentOddEven = differentOddEven
            DifferentFirst = differentFirst
            ScaleWithDoc = scaleWithDoc
        }

    let private breaksNative (reader: System.Xml.XmlReader) =
        let ids = System.Collections.Generic.List<int>()
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element when reader.LocalName = "brk" && nativeBoolAttr reader "man" ->
                    match nativeIntAttr reader "id" with
                    | Some id when id > 0 -> ids.Add(id)
                    | _ -> ()
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        ids |> resizeListToList |> List.sort

    let private mergeCellNative (reader: System.Xml.XmlReader) =
        nativeAttr reader "ref" |> Option.bind parseMergeCell

    let private hyperlinkNative resolveRel (reader: System.Xml.XmlReader) =
        nativeAttr reader "ref"
        |> Option.map (fun range ->
            let location = nativeAttr reader "location"
            let rid = nativeRelId reader
            let target, isInternal =
                match location with
                | Some loc -> loc, true
                | None ->
                    match rid |> Option.bind resolveRel with
                    | Some url -> url, false
                    | None -> rid |> Option.defaultValue "", false
            { Range = range; Target = target; IsInternal = isInternal })

    let private worksheetFastNativeSupported (sheetXml: string) =
        sheetXml.Length >= worksheetFastNativeThreshold
        && sheetXml.IndexOf("conditionalFormatting", System.StringComparison.Ordinal) < 0
        && sheetXml.IndexOf("dataValidation", System.StringComparison.Ordinal) < 0

    let internal worksheetNativeFastCore (ctx: SheetParseContext) (name: string) (state: SheetState) (sheetXml: string) (comments: CellComment list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet option =
        if not (worksheetFastNativeSupported sheetXml) then
            None
        else
            let settings = System.Xml.XmlReaderSettings()
            settings.DtdProcessing <- System.Xml.DtdProcessing.Prohibit
            settings.XmlResolver <- null
            use stringReader = new System.IO.StringReader(sheetXml)
            use reader = System.Xml.XmlReader.Create(stringReader, settings)
            let rows = System.Collections.Generic.List<Row>()
            let columns = System.Collections.Generic.List<ColumnInfo>()
            let columnNamesByIndex = System.Collections.Generic.Dictionary<int, string>()
            let mergedCells = System.Collections.Generic.List<MergeCell>()
            let hyperlinks = System.Collections.Generic.List<Hyperlink>()
            let mutable format = emptySheetFormat
            let mutable frozenPaneValue = None
            let mutable showZerosValue = true
            let mutable showGridLinesValue = true
            let mutable fitToPage = false
            let mutable pageSetupValue = None
            let mutable pageMarginsValue = None
            let mutable printOptionsValue = defaultPrintOptions
            let mutable headerFooterValue = defaultHeaderFooter
            let mutable rowBreaksValue = []
            let mutable colBreaksValue = []
            let mutable prevRow = 0
            let resolveRel rid =
                sheetRels
                |> List.tryFind (fun r -> r.Id = rid && r.Type.EndsWith("/hyperlink"))
                |> Option.map (fun r -> r.Target)

            while reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element ->
                    match reader.LocalName with
                    | "sheetPr" ->
                        fitToPage <- fitToPageNative reader
                    | "sheetViews" ->
                        let showZeros, showGridLines, frozenPane = sheetViewsNative reader
                        showZerosValue <- showZeros
                        showGridLinesValue <- showGridLines
                        frozenPaneValue <- frozenPane
                    | "sheetFormatPr" ->
                        format <- sheetFormatPrNative reader
                    | "col" ->
                        match columnInfoNative reader with
                        | Some col -> columns.Add(col)
                        | None -> ()
                    | "row" ->
                        let row = rowNative ctx columnNamesByIndex prevRow reader
                        prevRow <- row.Index
                        rows.Add(row)
                    | "mergeCell" ->
                        match mergeCellNative reader with
                        | Some mergeCell -> mergedCells.Add(mergeCell)
                        | None -> ()
                    | "hyperlink" ->
                        match hyperlinkNative resolveRel reader with
                        | Some link -> hyperlinks.Add(link)
                        | None -> ()
                    | "pageMargins" ->
                        pageMarginsValue <- Some (pageMarginsNative reader)
                    | "pageSetup" ->
                        pageSetupValue <- Some (pageSetupNative fitToPage reader)
                    | "printOptions" ->
                        printOptionsValue <- printOptionsNative reader
                    | "headerFooter" ->
                        headerFooterValue <- headerFooterNative reader
                    | "rowBreaks" ->
                        rowBreaksValue <- breaksNative reader
                    | "colBreaks" ->
                        colBreaksValue <- breaksNative reader
                    | _ -> ()
                | _ -> ()

            let pageSetupFinal =
                match pageSetupValue, fitToPage with
                | Some setup, fit -> Some { setup with FitToPage = fit }
                | None, true -> Some (defaultPageSetupNative true)
                | None, false -> None

            Some {
                Name = name
                State = state
                Rows = resizeListToList rows
                MergedCells = resizeListToList mergedCells
                Columns = resizeListToList columns
                Format = format
                Comments = comments
                ConditionalFormats = []
                DataValidations = []
                Charts = []
                PivotTables = pivotTables
                Tables = tables
                FrozenPane = frozenPaneValue
                ShowZeros = showZerosValue
                ShowGridLines = showGridLinesValue
                Hyperlinks = resizeListToList hyperlinks
                PageSetup = pageSetupFinal
                PageMargins = pageMarginsValue
                PrintOptions = printOptionsValue
                HeaderFooter = headerFooterValue
                RowBreaks = rowBreaksValue
                ColBreaks = colBreaksValue
                PrintArea = printNames.PrintArea
                PrintTitleRows = printNames.PrintTitleRows
                PrintTitleCols = printNames.PrintTitleCols
            }

    let worksheetNativeFastWithStyles (styles: CellStyle array option) (name: string) (sheetXml: string) (comments: CellComment list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet option =
        worksheetNativeFastCore (newSheetContext styles false) name SheetVisible sheetXml comments [] tables sheetRels printNames
#endif

    let internal worksheetFromNodeCore (ctx: SheetParseContext) (name: string) (state: SheetState) (sheet: XmlNode) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet =
        let sheetData = sElem "sheetData" sheet

        let mutable prevRow = 0
        let rows =
            sheetData
            |> Option.map (childElems "row")
            |> Option.defaultValue []
            |> List.map (fun rowNode ->
                let row = rowFromNode ctx prevRow rowNode
                prevRow <- row.Index
                row)

        // Resolve hyperlink r:id values against the sheet rels: external
        // hyperlinks (Target on the Relationship is the actual URL) get
        // routed straight in. Tables / drawings live in the same rels
        // file but with a different Type — filter to hyperlink type.
        let resolveRel rid =
            sheetRels
            |> List.tryFind (fun r -> r.Id = rid && r.Type.EndsWith("/hyperlink"))
            |> Option.map (fun r -> r.Target)

        {
            Name = name
            State = state
            Rows = rows
            MergedCells = mergeCells sheet
            Columns = columns sheet
            Format = sheetFormatPr sheet
            Comments = comments
            ConditionalFormats = condFormats
            DataValidations = dataVals
            Charts = charts
            PivotTables = pivotTables
            Tables = tables
            FrozenPane = frozenPane sheet
            ShowZeros = showZeros sheet
            ShowGridLines = showGridLines sheet
            Hyperlinks = hyperlinks resolveRel sheet
            PageSetup = pageSetup sheet
            PageMargins = pageMargins sheet
            PrintOptions = printOptions sheet
            HeaderFooter = headerFooter sheet
            RowBreaks = breaks "rowBreaks" sheet
            ColBreaks = breaks "colBreaks" sheet
            PrintArea = printNames.PrintArea
            PrintTitleRows = printNames.PrintTitleRows
            PrintTitleCols = printNames.PrintTitleCols
        }

    let worksheetFromNodeWithStyles (styles: CellStyle array option) (name: string) (sheet: XmlNode) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet =
        worksheetFromNodeCore (newSheetContext styles false) name SheetVisible sheet comments condFormats dataVals charts pivotTables tables sheetRels printNames

    let worksheetWithStyles (styles: CellStyle array option) (name: string) (sheetXml: string) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet =
        let sheet = XmlDsl.parse sheetXml
        worksheetFromNodeWithStyles styles name sheet comments condFormats dataVals charts pivotTables tables sheetRels printNames

    let worksheet name sheetXml comments condFormats dataVals charts pivotTables tables sheetRels printNames =
        worksheetWithStyles None name sheetXml comments condFormats dataVals charts pivotTables tables sheetRels printNames

    /// Parse an `A1:F5` style range into 1-indexed inclusive bounds.
    let private parseRange (ref: string) =
        let parts = ref.Split(':')
        let parseCorner (s: string) =
            let letters, parsedRow = parseCellRefParts s
            let row = if parsedRow > 0 then parsedRow else 1
            (row, columnLettersToIndex letters)
        if parts.Length = 1 then
            let (r, c) = parseCorner parts.[0]
            Some (r, c, r, c)
        elif parts.Length = 2 then
            let (r1, c1) = parseCorner parts.[0]
            let (r2, c2) = parseCorner parts.[1]
            Some (r1, c1, r2, c2)
        else None

    /// Parse `xl/tables/tableN.xml`. Skips tables with no usable range.
    let parseTable (tableXml: string) : TableDefinition option =
        try
            let doc = XmlDsl.parse tableXml
            match Xml.attr "ref" doc with
            | None -> None
            | Some refStr ->
                match parseRange refStr with
                | None -> None
                | Some (r1, c1, r2, c2) ->
                    let headerCount =
                        Xml.attr "headerRowCount" doc
                        |> Option.bind tryParseInt
                        |> Option.defaultValue 1
                    let style =
                        sElem "tableStyleInfo" doc
                        |> Option.bind (fun node ->
                            Xml.attr "name" node
                            |> Option.map (fun name ->
                                let boolAttr key =
                                    Xml.attr key node
                                    |> Option.map (fun v -> v = "1" || v = "true")
                                    |> Option.defaultValue false
                                {
                                    Name = name
                                    ShowFirstColumn = boolAttr "showFirstColumn"
                                    ShowLastColumn = boolAttr "showLastColumn"
                                    ShowRowStripes = boolAttr "showRowStripes"
                                    ShowColumnStripes = boolAttr "showColumnStripes"
                                }))
                    let autoFilter =
                        sElem "autoFilter" doc
                        |> Option.bind (fun node -> Xml.attr "ref" node)
                        |> Option.bind parseRange
                        |> Option.map (fun (_, fc1, _, fc2) -> (fc1, fc2))
                    let tableName =
                        Xml.attr "displayName" doc
                        |> Option.orElse (Xml.attr "name" doc)
                        |> Option.defaultValue ""
                    let columns =
                        Ns.descendantsNs "tableColumn" doc
                        |> List.choose (Xml.attr "name")
                    let totalsRowCount =
                        Xml.attr "totalsRowCount" doc
                        |> Option.bind tryParseInt
                        |> Option.defaultValue 0
                    Some {
                        Name = tableName
                        Columns = columns
                        StartRow = r1
                        StartCol = c1
                        EndRow = r2
                        EndCol = c2
                        HeaderRowCount = headerCount
                        TotalsRowCount = totalsRowCount
                        Style = style
                        AutoFilter = autoFilter
                    }
        with _ -> None

    /// Parse `xl/theme/theme1.xml` accent1..accent6 sRGB colors. Returns an
    /// array of six "RRGGBB" strings indexed 0..5 (accent1 == [0]). When the
    /// theme is missing or unparseable, falls back to Excel's 2013 defaults.
    let parseThemeAccents (themeXml: string) : string array =
        let defaults = [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]
        try
            let doc = XmlDsl.parse themeXml
            let pickAccent name =
                Ns.descendantsNs name doc
                |> List.tryHead
                |> Option.bind (Ns.descendantsNs "srgbClr" >> List.tryHead)
                |> Option.bind (Xml.attr "val")
            let names = [| "accent1"; "accent2"; "accent3"; "accent4"; "accent5"; "accent6" |]
            names
            |> Array.mapi (fun i n ->
                pickAccent n |> Option.defaultValue defaults.[i])
        with _ -> defaults

    /// Parse the full 12-slot theme color scheme indexed in the order that
    /// XLSX `<color theme="N"/>` integers use:
    ///   0=lt1 1=dk1 2=lt2 3=dk2 4..9=accent1..6 10=hlink 11=folHlink
    /// OOXML swaps dk1/lt1 (and dk2/lt2) between the theme XML element order
    /// (`dk1` first) and the theme-N integer space — that swap is folded into
    /// the names array below so callers can index by raw theme="N".
    /// Defaults are Office 2013's scheme.
    let parseThemeColors (themeXml: string) : string array =
        // Slot defaults in XLSX theme-N order.
        let defaults = [|
            "FFFFFF"  // 0 lt1 (background 1)
            "000000"  // 1 dk1 (text 1)
            "E7E6E6"  // 2 lt2 (background 2)
            "44546A"  // 3 dk2 (text 2)
            "4472C4"  // 4 accent1
            "ED7D31"  // 5 accent2
            "A5A5A5"  // 6 accent3
            "FFC000"  // 7 accent4
            "5B9BD5"  // 8 accent5
            "70AD47"  // 9 accent6
            "0563C1"  // 10 hlink
            "954F72"  // 11 folHlink
        |]
        try
            let doc = XmlDsl.parse themeXml
            let pickScheme name =
                // <a:clrScheme> children carry either <a:srgbClr val="..."/>
                // or <a:sysClr val="windowText" lastClr="..."/>. For sysClr we
                // prefer lastClr (the persisted RGB) over val (which is just a
                // system-color label).
                Ns.descendantsNs name doc
                |> List.tryHead
                |> Option.bind (fun node ->
                    let srgb = Ns.descendantsNs "srgbClr" node |> List.tryHead |> Option.bind (Xml.attr "val")
                    match srgb with
                    | Some s -> Some s
                    | None ->
                        Ns.descendantsNs "sysClr" node
                        |> List.tryHead
                        |> Option.bind (fun sn ->
                            Xml.attr "lastClr" sn
                            |> Option.orElseWith (fun () -> Xml.attr "val" sn)))
            // Names in XLSX theme-N order (note dk1/lt1 swap vs the theme XML
            // element order).
            let names = [|
                "lt1"; "dk1"; "lt2"; "dk2"
                "accent1"; "accent2"; "accent3"; "accent4"; "accent5"; "accent6"
                "hlink"; "folHlink"
            |]
            names
            |> Array.mapi (fun i n -> pickScheme n |> Option.defaultValue defaults.[i])
        with _ -> defaults

    /// Parse the theme's majorFont / minorFont latin typeface names —
    /// used when a font in styles.xml has `<scheme val="major|minor"/>`
    /// without an explicit `<name>`. Returns (major, minor); each is
    /// None when the slot is missing, defaulting to Calibri-Light /
    /// Calibri in the caller's resolution if absent.
    let parseThemeFonts (themeXml: string) : string option * string option =
        try
            let doc = XmlDsl.parse themeXml
            let latinOf slotName =
                Ns.descendantsNs slotName doc
                |> List.tryHead
                |> Option.bind (fun slot ->
                    Ns.descendantsNs "latin" slot
                    |> List.tryHead
                    |> Option.bind (Xml.attr "typeface"))
            latinOf "majorFont", latinOf "minorFont"
        with _ -> None, None

    /// Excel 8 / BIFF8 legacy indexed palette. 64 indexed slots plus indices
    /// 64/65 reserved for the OS default foreground/background. Values match
    /// Microsoft's documented palette (see `IndexedColors` in the OOXML
    /// transitional spec). Returns None when the index is out of range.
    let indexedPaletteColor (i: int) : string option =
        let palette = [|
            "000000"; "FFFFFF"; "FF0000"; "00FF00"; "0000FF"; "FFFF00"; "FF00FF"; "00FFFF"
            "000000"; "FFFFFF"; "FF0000"; "00FF00"; "0000FF"; "FFFF00"; "FF00FF"; "00FFFF"
            "800000"; "008000"; "000080"; "808000"; "800080"; "008080"; "C0C0C0"; "808080"
            "9999FF"; "993366"; "FFFFCC"; "CCFFFF"; "660066"; "FF8080"; "0066CC"; "CCCCFF"
            "000080"; "FF00FF"; "FFFF00"; "00FFFF"; "800080"; "800000"; "008080"; "0000FF"
            "00CCFF"; "CCFFFF"; "CCFFCC"; "FFFF99"; "99CCFF"; "FF99CC"; "CC99FF"; "FFCC99"
            "3366FF"; "33CCCC"; "99CC00"; "FFCC00"; "FF9900"; "FF6600"; "666699"; "969696"
            "003366"; "339966"; "003300"; "333300"; "993300"; "993366"; "333399"; "333333"
        |]
        match i with
        | n when n >= 0 && n < palette.Length -> Some palette.[n]
        | 64 -> Some "000000"  // system foreground
        | 65 -> Some "FFFFFF"  // system background
        | _ -> None

    /// Apply an OOXML `tint` modifier to an RRGGBB color. Per the spec, tint
    /// is in [-1, 1]: negative darkens, positive lightens, against the HLS
    /// luminance channel. Cells in poi-51585 / poi-51710 lean on this for
    /// banded "lighter accent" fills — without it, every banded row painted
    /// the same accent color.
    let applyColorTint (rgb: string) (tint: float) : string =
        if System.Double.IsNaN tint || tint = 0.0 then rgb
        else
            let clean = if rgb.StartsWith "#" then rgb.Substring 1 else rgb
            let hex = if clean.Length = 8 then clean.Substring 2 else clean
            if hex.Length <> 6 then rgb
            else
                let r = System.Convert.ToInt32(hex.Substring(0, 2), 16) |> float
                let g = System.Convert.ToInt32(hex.Substring(2, 2), 16) |> float
                let b = System.Convert.ToInt32(hex.Substring(4, 2), 16) |> float
                let r1, g1, b1 = r / 255.0, g / 255.0, b / 255.0
                let mx = max r1 (max g1 b1)
                let mn = min r1 (min g1 b1)
                let l = (mx + mn) / 2.0
                let d = mx - mn
                let s = if d = 0.0 then 0.0 elif l < 0.5 then d / (mx + mn) else d / (2.0 - mx - mn)
                let h =
                    if d = 0.0 then 0.0
                    elif mx = r1 then ((g1 - b1) / d + (if g1 < b1 then 6.0 else 0.0)) / 6.0
                    elif mx = g1 then ((b1 - r1) / d + 2.0) / 6.0
                    else ((r1 - g1) / d + 4.0) / 6.0
                let t = max -1.0 (min 1.0 tint)
                let l' =
                    if t < 0.0 then l * (1.0 + t)
                    else l * (1.0 - t) + t
                let hueToRgb p q t =
                    let t = if t < 0.0 then t + 1.0 elif t > 1.0 then t - 1.0 else t
                    if t < 1.0 / 6.0 then p + (q - p) * 6.0 * t
                    elif t < 0.5 then q
                    elif t < 2.0 / 3.0 then p + (q - p) * (2.0 / 3.0 - t) * 6.0
                    else p
                let r2, g2, b2 =
                    if s = 0.0 then l', l', l'
                    else
                        let q = if l' < 0.5 then l' * (1.0 + s) else l' + s - l' * s
                        let p = 2.0 * l' - q
                        hueToRgb p q (h + 1.0 / 3.0),
                        hueToRgb p q h,
                        hueToRgb p q (h - 1.0 / 3.0)
                let toByte v = int (max 0.0 (min 255.0 (v * 255.0 + 0.5)))
                sprintf "%02X%02X%02X" (toByte r2) (toByte g2) (toByte b2)

    /// Resolve a color element (`<fgColor .../>`, `<color .../>` etc.) into an
    /// RRGGBB string. Tries `rgb` → `theme` (+ optional `tint`) → `indexed`.
    /// Returns None when the element has no recognized color attribute.
    let resolveColorElement (themeColors: string array) (node: XmlNode) : string option =
        let rgb = Xml.attr "rgb" node
        let themed =
            Xml.attr "theme" node
            |> Option.bind (fun s ->
                match tryParseInt s with
                | Some n when n >= 0 && n < themeColors.Length -> Some themeColors.[n]
                | _ -> None)
        let indexed =
            Xml.attr "indexed" node
            |> Option.bind (fun s -> tryParseInt s |> Option.bind indexedPaletteColor)
        let baseColor =
            rgb
            |> Option.orElseWith (fun () -> themed)
            |> Option.orElseWith (fun () -> indexed)
        match baseColor with
        | None -> None
        | Some raw ->
            // OOXML rgb attributes are AARRGGBB (alpha first). Strip the
            // leading alpha so every consumer sees a plain 6-char RRGGBB —
            // some renderers' parseHexColor treat an 8-char string as
            // RRGGBBAA (alpha last), which would mangle the channels.
            let stripped = raw.TrimStart('#')
            let c = if stripped.Length = 8 then stripped.Substring 2 else stripped
            let tint =
                Xml.attr "tint" node
                |> Option.bind tryParseFloat
                |> Option.defaultValue 0.0
            if tint = 0.0 then Some c
            else Some (applyColorTint c tint)

    // Style parsing. parseFont returns an 9-tuple covering every <font>
    // child plus the scheme hint — callers pick the slots they care
    // about. The scheme is "" when absent, "major"/"minor" otherwise;
    // theme-font resolution happens later, after the workbook theme
    // majorFont/minorFont arrays have been parsed.
    let parseFont (themeColors: string array) (fontNode: XmlNode) : (bool * bool * bool * bool * bool * string option * float option * string option * string) =
        let bold = sElem "b" fontNode |> Option.isSome
        let italic = sElem "i" fontNode |> Option.isSome
        let underline = sElem "u" fontNode |> Option.isSome
        let underlineDouble =
            sElem "u" fontNode
            |> Option.bind (Xml.attr "val")
            |> Option.map (fun v -> v = "double" || v = "doubleAccounting")
            |> Option.defaultValue false
        let strike = sElem "strike" fontNode |> Option.isSome
        let color = sElem "color" fontNode |> Option.bind (resolveColorElement themeColors)
        let size = sElem "sz" fontNode |> Option.bind (Xml.attr "val") |> Option.bind tryParseFloat
        let name = sElem "name" fontNode |> Option.bind (Xml.attr "val")
        let scheme =
            sElem "scheme" fontNode
            |> Option.bind (Xml.attr "val")
            |> Option.defaultValue ""
        (bold, italic, underline, underlineDouble, strike, color, size, name, scheme)

    /// Blend two RRGGBB hex colors by `t ∈ [0,1]` — 0.0 returns `a`, 1.0
    /// returns `b`. Used for gradient-stop approximation and named-gray
    /// pattern types (we don't have a tile primitive, so a tinted blend
    /// of the fg toward the bg conveys the same visual cue).
    let private blendHex (a: string) (b: string) (t: float) : string =
        let stripA = a.TrimStart('#')
        let stripB = b.TrimStart('#')
        let aHex = if stripA.Length = 8 then stripA.Substring 2 else stripA
        let bHex = if stripB.Length = 8 then stripB.Substring 2 else stripB
        if aHex.Length <> 6 || bHex.Length <> 6 then a
        else
            let chan i =
                let av = System.Convert.ToInt32(aHex.Substring(i, 2), 16)
                let bv = System.Convert.ToInt32(bHex.Substring(i, 2), 16)
                let v = float av * (1.0 - t) + float bv * t
                int (System.Math.Round v) |> max 0 |> min 255
            sprintf "%02X%02X%02X" (chan 0) (chan 2) (chan 4)

    let parseFill (themeColors: string array) (fillNode: XmlNode) : string option =
        // Excel-as-written XLSXs sometimes use bgColor instead of fgColor for
        // solid fills (POI does this; phpsp-27template has indexed bgColor on
        // table-row template entries). Prefer fgColor when present, fall back
        // to bgColor for solid/non-empty patterns.
        let pattern = sElem "patternFill" fillNode
        let patternType = pattern |> Option.bind (Xml.attr "patternType")
        // <gradientFill> is the other top-level fill kind; carries
        // ordered <stop position="..."><color/></stop> children. The
        // renderer has no gradient primitive, so collapse the stops into
        // a single representative color by blending equally between
        // adjacent stops weighted by position. For the typical 2-stop
        // case (most authors use start + end), this lands midway between
        // the two endpoints — a recognizable "purple" for yellow→red,
        // "tan" for white→orange, etc. — without a separate render path.
        let gradientStops =
            sElem "gradientFill" fillNode
            |> Option.map (fun gf ->
                Ns.descendantsNs "stop" gf
                |> List.choose (fun stop ->
                    let pos =
                        Xml.attr "position" stop
                        |> Option.bind tryParseFloat
                        |> Option.defaultValue 0.0
                    sElem "color" stop
                    |> Option.bind (resolveColorElement themeColors)
                    |> Option.map (fun hex -> pos, hex))
                |> List.sortBy fst)
            |> Option.defaultValue []
        match patternType, gradientStops with
        | _, (s :: rest) when not (List.isEmpty rest) ->
            // Multi-stop gradient: blend pairwise to the midpoint, then
            // average. Simpler approximation: take the midpoint of the
            // first and last stops.
            let _, firstHex = s
            let _, lastHex = List.last gradientStops
            Some (blendHex firstHex lastHex 0.5)
        | _, [ (_, hex) ] -> Some hex
        | Some "none", _ -> None
        | Some "gray125", _ -> None
        // Named gray patterns have no fgColor in OOXML — the pattern *is*
        // the color. Approximate via a tint between the cell's bg and a
        // "named-gray" tone:
        //   lightGray  ≈ #E0E0E0 (25% tone)
        //   mediumGray ≈ #A0A0A0 (50% tone)
        //   darkGray   ≈ #606060 (75% tone)
        | Some "lightGray", _ -> Some "E0E0E0"
        | Some "mediumGray", _ -> Some "A0A0A0"
        | Some "darkGray", _ -> Some "606060"
        | Some patternName, _ ->
            // Directional patterns (darkHorizontal/Vertical/Down/Up/Grid/
            // Trellis, lightHorizontal/Vertical/Down/Up/Grid/Trellis) carry
            // an explicit fgColor (the stripe color) and optionally bgColor
            // (the field color). Without a tile primitive we blend fg→bg by
            // an empirical density factor that matches the pattern's
            // density: dark-* ≈ 70% fg, light-* ≈ 30% fg.
            let fg = pattern |> Option.bind (sElem "fgColor") |> Option.bind (resolveColorElement themeColors)
            let bg = pattern |> Option.bind (sElem "bgColor") |> Option.bind (resolveColorElement themeColors)
            let density =
                if patternName.StartsWith "dark" then 0.7
                elif patternName.StartsWith "light" then 0.3
                else 1.0
            match fg, bg with
            | Some fgHex, Some bgHex when density < 1.0 -> Some (blendHex bgHex fgHex density)
            | Some fgHex, None when density < 1.0 -> Some (blendHex "FFFFFF" fgHex density)
            | Some fgHex, _ -> Some fgHex
            | None, Some bgHex -> Some bgHex
            | None, None -> None
        | None, _ ->
            let fg = pattern |> Option.bind (sElem "fgColor") |> Option.bind (resolveColorElement themeColors)
            let bg = pattern |> Option.bind (sElem "bgColor") |> Option.bind (resolveColorElement themeColors)
            fg |> Option.orElseWith (fun () -> bg)

    /// Parse a single <left>/<right>/<top>/<bottom> child of a <border>.
    /// Returns Some when the side has a `style` attribute; None for empty
    /// (unstyled) sides.
    let parseBorderSide (themeColors: string array) (sideNode: XmlNode) : CellBorderSide option =
        Xml.attr "style" sideNode
        |> Option.map (fun style ->
            let color =
                sElem "color" sideNode
                |> Option.bind (resolveColorElement themeColors)
            { Style = style; Color = color })

    type BorderRecord = {
        Top: CellBorderSide option
        Bottom: CellBorderSide option
        Left: CellBorderSide option
        Right: CellBorderSide option
        /// `<diagonal>` style + color, paired with the parent border's
        /// `diagonalUp`/`diagonalDown` attributes. Excel renders one
        /// stroke per direction flag; both true gives the X pattern.
        Diagonal: CellBorderSide option
        DiagonalUp: bool
        DiagonalDown: bool
    }

    let emptyBorder : BorderRecord = {
        Top = None
        Bottom = None
        Left = None
        Right = None
        Diagonal = None
        DiagonalUp = false
        DiagonalDown = false
    }

    let parseBorder (themeColors: string array) (borderNode: XmlNode) : BorderRecord =
        let flag attr =
            Xml.attr attr borderNode
            |> Option.map (fun v -> v = "1" || v = "true")
            |> Option.defaultValue false
        {
            Top = sElem "top" borderNode |> Option.bind (parseBorderSide themeColors)
            Bottom = sElem "bottom" borderNode |> Option.bind (parseBorderSide themeColors)
            Left = sElem "left" borderNode |> Option.bind (parseBorderSide themeColors)
            Right = sElem "right" borderNode |> Option.bind (parseBorderSide themeColors)
            Diagonal = sElem "diagonal" borderNode |> Option.bind (parseBorderSide themeColors)
            DiagonalUp = flag "diagonalUp"
            DiagonalDown = flag "diagonalDown"
        }

    /// Strip OOXML `\X` literal-char escapes from a format string. The XML
    /// stores `yyyy\-mm\-dd\ hh:mm:ss` to mean "yyyy-mm-dd hh:mm:ss" with
    /// the dashes/space as literals (so the `-` doesn't trip the negative-
    /// subformat parser). Excel COM's `cell.NumberFormat` strips the
    /// backslashes when reporting the format. Mirrors that — also keeps
    /// `\\` as a single literal backslash, leaves quoted "..." strings
    /// intact (the contents are already literal).
    let private stripBackslashEscapes (code: string) : string =
        let sb = System.Text.StringBuilder(code.Length)
        let mutable i = 0
        let mutable inQuote = false
        while i < code.Length do
            let ch = code.[i]
            if inQuote then
                sb.Append ch |> ignore
                if ch = '"' then inQuote <- false
                i <- i + 1
            elif ch = '"' then
                sb.Append ch |> ignore
                inQuote <- true
                i <- i + 1
            elif ch = '\\' && i + 1 < code.Length then
                sb.Append code.[i + 1] |> ignore
                i <- i + 2
            else
                sb.Append ch |> ignore
                i <- i + 1
        sb.ToString()

    /// Strip quotes from currency-symbol literals — `"$"` → `$`, `"€"` → `€`,
    /// etc. Excel COM unwraps these single-char currency literals when
    /// reporting the format. Other quoted strings (like `"-"` inside an
    /// accounting placeholder) keep their quotes since stripping them would
    /// confuse the format parser. The synthetic-numfmt fixture's accounting
    /// format `_("$"* #,##0.00_);…` is reported as `_($* #,##0.00_);…`.
    let private stripCurrencyQuotes (code: string) : string =
        let isCurrencySymbol (c: char) =
            c = '$' || c = '€' || c = '£' || c = '¥' || c = '₩'
            || c = '¢' || c = '₹' || c = '₽' || c = '₪' || c = '₱'
        System.Text.RegularExpressions.Regex.Replace(
            code,
            "\"(.)\"",
            fun m ->
                let inner = m.Groups.[1].Value.[0]
                if isCurrencySymbol inner then string inner else m.Value)

    let parseNumberFormat (numFmtNode: XmlNode) : string option =
        // Normalize the special "GENERAL" sentinel (case-insensitive) to
        // canonical "General". Excel COM's cell.NumberFormat always reports
        // this form regardless of how the file stored it, and our renderer
        // / harness expect the canonical spelling — pandas-test_converters
        // ships <numFmt formatCode="GENERAL"/> and Excel reports it as
        // "General" through COM, so a literal pass-through diverged on
        // every cell.
        //
        // Also strip OOXML's `\X` literal-char escapes: Excel COM returns
        // the format without them. Without this, pandas-testmultiindex's
        // `yyyy\-mm\-dd\ hh:mm:ss` diverged from Excel's reported
        // `yyyy-mm-dd hh:mm:ss` on every cell.
        Xml.attr "formatCode" numFmtNode
        |> Option.map (fun code ->
            if code.Equals("general", System.StringComparison.OrdinalIgnoreCase) then "General"
            else stripBackslashEscapes code |> stripCurrencyQuotes)

    /// Built-in numFmtId → format code table. Per the OOXML spec, ids 0–49
    /// are pre-defined and don't appear in the sheet's `<numFmts>` block, so
    /// readers need their own lookup. Custom formats start at 164.
    let builtInNumberFormat (id: int) : string option =
        match id with
        | 0 -> None  // General → no explicit code; renderer applies General trimming
        | 1 -> Some "0"
        | 2 -> Some "0.00"
        | 3 -> Some "#,##0"
        | 4 -> Some "#,##0.00"
        | 9 -> Some "0%"
        | 10 -> Some "0.00%"
        | 11 -> Some "0.00E+00"
        | 12 -> Some "# ?/?"
        | 13 -> Some "# ??/??"
        | 14 -> Some "m/d/yyyy"
        | 15 -> Some "d-mmm-yy"
        | 16 -> Some "d-mmm"
        | 17 -> Some "mmm-yy"
        | 18 -> Some "h:mm AM/PM"
        | 19 -> Some "h:mm:ss AM/PM"
        | 20 -> Some "h:mm"
        | 21 -> Some "h:mm:ss"
        | 22 -> Some "m/d/yyyy h:mm"
        | 37 -> Some "#,##0;(#,##0)"
        | 38 -> Some "#,##0;[Red](#,##0)"
        | 39 -> Some "#,##0.00;(#,##0.00)"
        | 40 -> Some "#,##0.00;[Red](#,##0.00)"
        | 45 -> Some "mm:ss"
        | 46 -> Some "[h]:mm:ss"
        | 47 -> Some "mmss.0"
        | 48 -> Some "##0.0E+0"
        | 49 -> Some "@"
        | _ -> None

    let parseCellXf (xfNode: XmlNode) (fonts: (bool * bool * bool * bool * bool * string option * float option * string option * string) array) (fills: string option array) (borders: BorderRecord array) (numFmts: Map<int, string>) (themeMajorFont: string option) (themeMinorFont: string option) (cellStyleXfs: (int * int * int * int option) array) : CellStyle =
        // OOXML cellXf inheritance: when the cellXf references a base
        // style via `xfId`, fields not explicitly overridden via the
        // matching `applyXxx="1"` flag fall through to the base style's
        // id. Excel's actual behaviour: if `applyXxx` is absent OR `0`,
        // the base wins; if `1`, the cellXf wins. We resolve the
        // effective fontId/fillId/borderId/numFmtId by checking each
        // applyXxx flag against the base's id from cellStyleXfs.
        let xfId = Xml.attr "xfId" xfNode |> Option.bind tryParseInt
        let baseFontId, baseFillId, baseBorderId, baseNumFmtId =
            match xfId with
            | Some id when id >= 0 && id < cellStyleXfs.Length ->
                let (f, fi, b, n) = cellStyleXfs.[id]
                f, fi, b, n
            | _ -> 0, 0, 0, None
        let applyFlag name =
            Xml.attr name xfNode
            |> Option.map (fun v -> v = "1" || v = "true")
        let ownInt name = Xml.attr name xfNode |> Option.bind tryParseInt
        let resolveId (flagName: string) (ownAttr: string) (baseId: int) (defaultId: int) =
            let own = ownInt ownAttr
            match applyFlag flagName, own with
            // applyXxx="1" → own value wins (using defaultId if absent)
            | Some true, Some v -> v
            | Some true, None -> defaultId
            // applyXxx absent: per Excel's de-facto behavior the default is
            // "apply" (treat as true), so an own value overrides the base.
            // pandas-test_none_type has <xf numFmtId="165"> with no
            // applyNumberFormat → Excel uses 165 (@), not inherited 0.
            | None, Some v -> v
            // applyXxx="0" (or both flag and own absent) → inherit base when
            // we have an xfId, otherwise fall through to own/default.
            | Some false, _ ->
                match xfId with
                | Some _ -> baseId
                | None -> own |> Option.defaultValue defaultId
            | None, None ->
                match xfId with
                | Some _ -> baseId
                | None -> defaultId
        let fontId = resolveId "applyFont" "fontId" baseFontId 0
        let fillId = resolveId "applyFill" "fillId" baseFillId 0
        let borderId = resolveId "applyBorder" "borderId" baseBorderId 0
        let numFmtId =
            // Number format: nullable. If applyNumberFormat="1" or absent
            // (Excel's default is "apply"), use own; if "0" inherit base.
            // Matches resolveId's handling of the other apply flags — see
            // pandas-test_none_type for the case this catches.
            let own = ownInt "numFmtId"
            match applyFlag "applyNumberFormat", own with
            | Some true, _ -> own
            | None, Some _ -> own
            | _, _ ->
                match xfId with
                | Some _ -> baseNumFmtId
                | None -> own

        let (bold, italic, underline, underlineDouble, strikethrough, fontColor, fontSize, fontNameRaw, scheme) =
            if fontId < fonts.Length then fonts.[fontId]
            else (false, false, false, false, false, None, None, None, "")
        // Resolve scheme="major|minor" to the theme font when no explicit
        // <name> overrides it. Falls through to None when neither side
        // names a font; downstream renderers default to the workbook
        // default font in that case.
        let fontName =
            match fontNameRaw with
            | Some _ -> fontNameRaw
            | None ->
                match scheme with
                | "major" -> themeMajorFont
                | "minor" -> themeMinorFont
                | _ -> None

        let fillColor =
            if fillId < fills.Length then fills.[fillId]
            else None

        let border =
            if borderId < borders.Length then borders.[borderId]
            else emptyBorder

        let numFmt =
            numFmtId
            |> Option.bind (fun id ->
                match Map.tryFind id numFmts with
                | Some custom -> Some custom
                | None -> builtInNumberFormat id)

        // <xf><alignment horizontal="..." vertical="..." wrapText="1"/></xf>
        let alignmentNode = sElem "alignment" xfNode
        let horizontalAlign = alignmentNode |> Option.bind (Xml.attr "horizontal")
        let verticalAlign = alignmentNode |> Option.bind (Xml.attr "vertical")
        let wrapText =
            alignmentNode
            |> Option.bind (Xml.attr "wrapText")
            |> Option.map (fun v -> v = "1" || v = "true")
            |> Option.defaultValue false
        let indent =
            alignmentNode
            |> Option.bind (Xml.attr "indent")
            |> Option.bind tryParseInt
            |> Option.defaultValue 0
        let shrinkToFit =
            alignmentNode
            |> Option.bind (Xml.attr "shrinkToFit")
            |> Option.map (fun v -> v = "1" || v = "true")
            |> Option.defaultValue false
        let textRotation =
            alignmentNode
            |> Option.bind (Xml.attr "textRotation")
            |> Option.bind tryParseInt
            |> Option.defaultValue 0

        {
            FontName = fontName
            FontColor = fontColor
            FillColor = fillColor
            Bold = bold
            Italic = italic
            Underline = underline
            UnderlineDouble = underlineDouble
            Strikethrough = strikethrough
            FontSize = fontSize
            NumberFormat = numFmt
            HorizontalAlign = horizontalAlign
            VerticalAlign = verticalAlign
            WrapText = wrapText
            Indent = indent
            ShrinkToFit = shrinkToFit
            TextRotation = textRotation
            BorderTop = border.Top
            BorderBottom = border.Bottom
            BorderLeft = border.Left
            BorderRight = border.Right
            BorderDiagonal = border.Diagonal
            DiagonalUp = border.DiagonalUp
            DiagonalDown = border.DiagonalDown
        }

    let parseStyles (themeColors: string array) (themeMajorFont: string option) (themeMinorFont: string option) (stylesXml: string) : CellStyle array =
        let doc = XmlDsl.parse stylesXml

        // Parse fonts. When a font has no own <sz val=...> we inherit the
        // workbook default font's size (the first <font> entry). Excel COM
        // reports `Font.Size = 10` for pandas-ints_spelled_with_decimals'
        // font[1] (an Arial without <sz/>) because that's what's painted
        // on screen — the size comes from font[0]'s <sz val="10"/>. Without
        // this fallback our renderer defaulted to 11 pt and underreported
        // the inherited size.
        let fontsRaw =
            sElem "fonts" doc
            |> Option.map (sElems "font")
            |> Option.defaultValue []
            |> List.map (parseFont themeColors)
            |> List.toArray
        let defaultParsedSize =
            if fontsRaw.Length > 0 then
                let (_, _, _, _, _, _, size, _, _) = fontsRaw.[0]
                size
            else None
        let inheritFontSize (font: (bool * bool * bool * bool * bool * string option * float option * string option * string)) =
            let (b, i, u, ud, s, c, sz, n, sch) = font
            match sz with
            | None -> (b, i, u, ud, s, c, defaultParsedSize, n, sch)
            | _ -> font
        let fonts = fontsRaw |> Array.map inheritFontSize

        // Parse fills
        let fills =
            sElem "fills" doc
            |> Option.map (sElems "fill")
            |> Option.defaultValue []
            |> List.map (parseFill themeColors)
            |> List.toArray

        // Parse borders
        let borders =
            sElem "borders" doc
            |> Option.map (sElems "border")
            |> Option.defaultValue []
            |> List.map (parseBorder themeColors)
            |> List.toArray

        // Parse number formats
        let numFmts =
            sElem "numFmts" doc
            |> Option.map (sElems "numFmt")
            |> Option.defaultValue []
            |> List.choose (fun nf ->
                xml {
                    let! id = Xml.attr "numFmtId" nf |> Option.bind tryParseInt
                    let! code = parseNumberFormat nf
                    return (id, code)
                })
            |> Map.ofList

        // Parse cellStyleXfs first — these are the "named" style base
        // records (Normal, Comma, Currency, etc.) that cellXfs can
        // reference via `xfId`. We don't need the full CellStyle here,
        // just the id pointers (fontId/fillId/borderId/numFmtId) used
        // by inheritance resolution in parseCellXf.
        let cellStyleXfs : (int * int * int * int option) array =
            sElem "cellStyleXfs" doc
            |> Option.map (sElems "xf")
            |> Option.defaultValue []
            |> List.map (fun xf ->
                let intOf name = Xml.attr name xf |> Option.bind tryParseInt
                let fontId = intOf "fontId" |> Option.defaultValue 0
                let fillId = intOf "fillId" |> Option.defaultValue 0
                let borderId = intOf "borderId" |> Option.defaultValue 0
                let numFmtId = intOf "numFmtId"
                (fontId, fillId, borderId, numFmtId))
            |> List.toArray

        // Parse cell formats (cellXfs) — these can inherit from
        // cellStyleXfs via `xfId`. parseCellXf resolves the chain.
        sElem "cellXfs" doc
        |> Option.map (sElems "xf")
        |> Option.defaultValue []
        |> List.map (fun xf -> parseCellXf xf fonts fills borders numFmts themeMajorFont themeMinorFont cellStyleXfs)
        |> List.toArray

    let sharedStrings (sstXml: string) : SharedStringEntry array =
#if !FABLE_COMPILER
        sharedStringsFast sstXml
#else
        let sst = XmlDsl.parse sstXml
        childElems "si" sst
        |> List.map (fun si ->
            let runs = richRunsOf si
            if List.isEmpty runs then
                // Plain SI — all text in direct <t> descendants (single or
                // none); preserve original concat-all-<t> behavior.
                let text = descendantTextConcat "t" si
                { Text = text; Runs = [] }
            else
                let text = richRunText runs
                { Text = text; Runs = runs })
        |> List.toArray
#endif

    // Workbook sheets metadata
    type SheetInfo = {
        Name: string
        SheetId: string
        RelationshipId: string
        State: SheetState
    }

    /// `<workbookPr date1904="1"/>`: serial 0 is 1904-01-01, not 1899-12-30.
    let date1904 (wbNode: XmlNode) : bool =
        sElem "workbookPr" wbNode
        |> Option.bind (Xml.attr "date1904")
        |> Option.map isTrueText
        |> Option.defaultValue false

    let sheets (wbNode: XmlNode) : SheetInfo list =
        let sheetsElem = sElem "sheets" wbNode

        sheetsElem
        |> Option.map (sElems "sheet")
        |> Option.defaultValue []
        |> List.choose (fun sheet ->
            xml {
                let! name = Xml.attr "name" sheet
                // The relationships namespace may be bound to any prefix.
                let! rId = relationshipId sheet
                return {
                    Name = name
                    SheetId = Xml.attr "sheetId" sheet |> Option.defaultValue ""
                    RelationshipId = rId
                    State =
                        match Xml.attr "state" sheet with
                        | Some "hidden" -> SheetHidden
                        | Some "veryHidden" -> SheetVeryHidden
                        | _ -> SheetVisible
                }
            })

let private sheetInfosForLoad loadAll sheetName (sheetInfos: Xlsx.SheetInfo list) =
    let indexed = sheetInfos |> List.mapi (fun index info -> index, info)
    if loadAll then
        indexed
    else
        match sheetName with
        | Some name ->
            indexed |> List.filter (fun (_, info) -> info.Name = name)
        | None ->
            match indexed with
            | first :: _ -> [ first ]
            | [] -> []

// Cell Comments parsing
module CellComments =
    /// Parse one `<comment ref authorId>`. The author is an index into the
    /// part's `<authors>` list, and the text is the concatenation of the
    /// `<t>` runs under `<text>` exactly as written — runs are pieces of one
    /// string, not words to be spaced apart.
    let parseCommentWithAuthors (authors: string array) (commentNode: XmlNode) : CellComment option =
        Xml.attr "ref" commentNode
        |> Option.map (fun ref ->
            let cellRef = Xlsx.parseCellRef ref
            let author =
                Xml.attr "authorId" commentNode
                |> Option.bind tryParseInt
                |> Option.bind (fun idx -> if idx >= 0 && idx < authors.Length then Some authors.[idx] else None)
                |> Option.filter (fun name -> name <> "")
            let text =
                match Ns.childNs "text" commentNode with
                | Some textNode -> Xlsx.descendantTextConcat "t" textNode
                | None -> Xlsx.descendantTextConcat "t" commentNode
            {
                CellRef = cellRef
                Author = author
                Text = text
            })

    let parseComment (commentNode: XmlNode) : CellComment option =
        parseCommentWithAuthors [||] commentNode

    let parseComments (xml: string) : CellComment list =
        let doc = XmlDsl.parse xml
        let authors =
            Ns.childNs "authors" doc
            |> Option.map (fun node ->
                Xml.children node
                |> List.choose (fun child ->
                    match Xml.name child with
                    | Some name when Ns.stripNs name = "author" -> Some (Xml.text child |> Option.defaultValue "")
                    | _ -> None))
            |> Option.defaultValue []
            |> List.toArray
        Ns.descendantsNs "comment" doc
        |> List.choose (parseCommentWithAuthors authors)

// Conditional Formatting parsing
module ConditionalFormats =
    let parseRule (themeColors: string array) (ruleNode: XmlNode) : ConditionalFormatRule =
        let ruleType = Xml.attr "type" ruleNode |> Option.defaultValue "cellIs"
        let priority = Xml.attr "priority" ruleNode |> Option.bind tryParseInt
        let stopIfTrue = Xml.attr "stopIfTrue" ruleNode |> Option.map isTrueText |> Option.defaultValue false
        let operator = Xml.attr "operator" ruleNode
        let dxfId = Xml.attr "dxfId" ruleNode |> Option.bind tryParseInt
        // cellIs uses one or two <formula> children (two for between);
        // collect them in document order.
        let formulas =
            Ns.descendantsNs "formula" ruleNode
            |> List.choose Xml.text
        // colorScale carries N <cfvo> stops then N matching <color>
        // entries; zip them positionally.
        let colorScale =
            Ns.childNs "colorScale" ruleNode
            |> Option.map (fun cs ->
                let cfvos = Ns.descendantsNs "cfvo" cs
                let colors = Ns.descendantsNs "color" cs
                List.zip
                    (cfvos |> List.truncate (List.length colors))
                    (colors |> List.truncate (List.length cfvos))
                |> List.map (fun (cfvo, colorNode) ->
                    {
                        CfvoType = Xml.attr "type" cfvo |> Option.defaultValue "min"
                        CfvoValue = Xml.attr "val" cfvo
                        Color =
                            Xlsx.resolveColorElement themeColors colorNode
                            |> Option.defaultValue "FFFFFF"
                    }))
            |> Option.defaultValue []
        let stripAlpha (raw: string) =
            let t = raw.TrimStart('#')
            if t.Length = 8 then t.Substring 2 else t
        let cfvoPairs (parent: XmlNode) =
            Ns.descendantsNs "cfvo" parent
            |> List.map (fun cfvo ->
                Xml.attr "type" cfvo |> Option.defaultValue "min",
                Xml.attr "val" cfvo)
        // <dataBar><cfvo type="min"/><cfvo type="max"/><color rgb=…/></dataBar>
        let dataBarColor, dataBarCfvo =
            match Ns.childNs "dataBar" ruleNode with
            | Some db ->
                let color =
                    Ns.descendantsNs "color" db
                    |> List.tryHead
                    |> Option.bind (Xlsx.resolveColorElement themeColors)
                color, cfvoPairs db
            | None -> None, []
        // <iconSet iconSet="3TrafficLights1"><cfvo .../>...</iconSet>
        let iconSet, iconSetCfvo =
            match Ns.childNs "iconSet" ruleNode with
            | Some ic ->
                // Excel defaults the family to 3TrafficLights1 when the
                // attribute is omitted.
                Some (Xml.attr "iconSet" ic |> Option.defaultValue "3TrafficLights1"),
                cfvoPairs ic
            | None -> None, []
        // top10 / aboveAverage carry their config as attributes on the
        // <cfRule> element itself rather than a child element.
        let attrInt name =
            Xml.attr name ruleNode
            |> Option.bind tryParseInt
        let attrBool name fallback =
            Xml.attr name ruleNode
            |> Option.map (fun v -> v = "1" || v = "true")
            |> Option.defaultValue fallback
        let rank = attrInt "rank"
        let percent = attrBool "percent" false
        let bottom = attrBool "bottom" false
        // aboveAverage attr defaults to true (the rule means "above" when
        // the attribute is absent); below-average rules set it to 0.
        let aboveAverage = attrBool "aboveAverage" true
        let equalAverage = attrBool "equalAverage" false
        let stdDev = attrInt "stdDev"
        {
            Type = ruleType
            Operator = operator
            Formulas = formulas
            DxfId = dxfId
            ColorScale = colorScale
            DataBarColor = dataBarColor
            DataBarCfvo = dataBarCfvo
            Text = Xml.attr "text" ruleNode
            IconSet = iconSet
            IconSetCfvo = iconSetCfvo
            Rank = rank
            Percent = percent
            Bottom = bottom
            AboveAverage = aboveAverage
            EqualAverage = equalAverage
            StdDev = stdDev
            Priority = priority
            StopIfTrue = stopIfTrue
        }

    let parseConditionalFormat (themeColors: string array) (cfNode: XmlNode) : ConditionalFormat option =
        Xml.attr "sqref" cfNode
        |> Option.map (fun range ->
            let rules =
                Ns.descendantsNs "cfRule" cfNode
                |> List.map (parseRule themeColors)
            {
                Range = range
                Rules = rules
            })

    let parseConditionalFormats (themeColors: string array) (sheet: XmlNode) : ConditionalFormat list =
        Ns.descendantsNs "conditionalFormatting" sheet
        |> List.choose (parseConditionalFormat themeColors)

    /// Parse the `<dxfs>` table out of styles.xml. Each `<dxf>` carries a
    /// partial font + fill; conditional-format rules reference them by
    /// index. dxf fills conventionally use `<bgColor>` (not fgColor),
    /// so check both.
    let parseDxfs (themeColors: string array) (stylesDoc: XmlNode) : Dxf array =
        Xlsx.sElem "dxfs" stylesDoc
        |> Option.map (Xlsx.sElems "dxf")
        |> Option.defaultValue []
        |> List.map (fun dxf ->
            let fontColor =
                Xlsx.sElem "font" dxf
                |> Option.bind (Xlsx.sElem "color")
                |> Option.bind (Xlsx.resolveColorElement themeColors)
            let fillColor =
                Xlsx.sElem "fill" dxf
                |> Option.bind (Xlsx.sElem "patternFill")
                |> Option.bind (fun pf ->
                    Xlsx.sElem "bgColor" pf |> Option.bind (Xlsx.resolveColorElement themeColors)
                    |> Option.orElseWith (fun () ->
                        Xlsx.sElem "fgColor" pf |> Option.bind (Xlsx.resolveColorElement themeColors)))
            let boolChild key =
                Xlsx.sElem "font" dxf
                |> Option.bind (Xlsx.sElem key)
                |> Option.map (fun node ->
                    Xml.attr "val" node
                    |> Option.map (fun v -> v <> "0" && v <> "false")
                    |> Option.defaultValue true)
            {
                FontColor = fontColor
                FillColor = fillColor
                Bold = boolChild "b"
                Italic = boolChild "i"
            })
        |> List.toArray

// Data Validation parsing
module DataValidations =
    let parseDataValidation (dvNode: XmlNode) : DataValidation option =
        Xml.attr "sqref" dvNode
        |> Option.map (fun range ->
            let valType = Xml.attr "type" dvNode |> Option.defaultValue "none"
            let operator = Xml.attr "operator" dvNode
            let flag name fallback =
                Xml.attr name dvNode |> Option.map isTrueText |> Option.defaultValue fallback
            // `showDropDown="1"` means the in-cell arrow is SUPPRESSED (the
            // attribute is named backwards in the schema).
            let showDropDown = not (flag "showDropDown" false)
            let formula1 = Ns.childNs "formula1" dvNode |> Option.bind Xml.text
            let formula2 = Ns.childNs "formula2" dvNode |> Option.bind Xml.text
            {
                Range = range
                ValidationType = valType
                Operator = operator
                Formula1 = formula1
                Formula2 = formula2
                AllowBlank = flag "allowBlank" false
                ShowDropDown = showDropDown
                ErrorMessage = Xml.attr "error" dvNode
                ErrorTitle = Xml.attr "errorTitle" dvNode
                ErrorStyle = Xml.attr "errorStyle" dvNode
                ShowErrorMessage = flag "showErrorMessage" false
                Prompt = Xml.attr "prompt" dvNode
                PromptTitle = Xml.attr "promptTitle" dvNode
                ShowInputMessage = flag "showInputMessage" false
            })

    let parseDataValidations (sheet: XmlNode) : DataValidation list =
        Ns.childNs "dataValidations" sheet
        |> Option.map (Ns.descendantsNs "dataValidation")
        |> Option.defaultValue []
        |> List.choose parseDataValidation

// Named Ranges parsing
module NamedRanges =
    let parseDefinedName (nameNode: XmlNode) : NamedRange option =
        Xml.attr "name" nameNode
        |> Option.bind (fun name ->
            Xml.text nameNode
            |> Option.map (fun refersTo ->
                let scope = Xml.attr "localSheetId" nameNode
                {
                    Name = name
                    RefersTo = refersTo
                    Scope = scope
                }))

    let parseNamedRanges (wbNode: XmlNode) : NamedRange list =
        Ns.childNs "definedNames" wbNode
        |> Option.map (fun dn -> Ns.descendantsNs "definedName" dn)
        |> Option.defaultValue []
        |> List.choose parseDefinedName

    let private stripSheetQualifier (s: string) =
        // 'Sheet One'!$A$1 / Sheet1!$A$1 — drop everything up to and
        // including the '!'. (Sheet names can contain a comma; we don't
        // handle that pathological case in print-titles splitting yet.)
        match s.IndexOf '!' with
        | -1 -> s
        | i -> s.Substring(i + 1)

    let private isRowsRef (s: string) =
        // "$1:$3" — only digits, '$', ':'.
        s.Length > 0
        && s |> Seq.forall (fun c -> System.Char.IsDigit c || c = '$' || c = ':')

    let private isColsRef (s: string) =
        // "$A:$X" — only letters, '$', ':'.
        s.Length > 0
        && s |> Seq.forall (fun c -> System.Char.IsLetter c || c = '$' || c = ':')

    /// Pull the print-related `_xlnm.*` defined names out of the workbook's
    /// name list and group their values per sheet (by `localSheetId`, the
    /// 0-based sheet index). Returns the rest of the names (regular user
    /// names) and a per-sheet lookup. Other `_xlnm.*` built-ins
    /// (`_FilterDatabase`, …) are dropped — they're Excel internals.
    let splitPrintNames (names: NamedRange list) : NamedRange list * Map<int, WorksheetPrintNames> =
        let mutable regular = []
        let mutable perSheet : Map<int, WorksheetPrintNames> = Map.empty
        let update idx upd =
            let cur = Map.tryFind idx perSheet |> Option.defaultValue emptyPrintNames
            perSheet <- Map.add idx (upd cur) perSheet
        let parseScope (scope: string option) =
            scope
            |> Option.bind (fun s ->
                tryParseInt s)
        for nr in names do
            match nr.Name, parseScope nr.Scope with
            | "_xlnm.Print_Area", Some idx ->
                update idx (fun w -> { w with PrintArea = Some (stripSheetQualifier nr.RefersTo) })
            | "_xlnm.Print_Titles", Some idx ->
                // Comma-separated rows part + cols part; either may be absent.
                let parts = nr.RefersTo.Split(',') |> Array.map stripSheetQualifier
                for p in parts do
                    if isRowsRef p then update idx (fun w -> { w with PrintTitleRows = Some p })
                    elif isColsRef p then update idx (fun w -> { w with PrintTitleCols = Some p })
            | n, _ when n.StartsWith "_xlnm." -> ()  // other built-ins: drop
            | _ -> regular <- nr :: regular
        (List.rev regular, perSheet)

// Charts parsing
module Charts =
    let parseChartType (chartNode: XmlNode) : ChartType =
        if Ns.descendantsNs "lineChart" chartNode |> List.isEmpty |> not then LineChart
        elif Ns.descendantsNs "barChart" chartNode |> List.isEmpty |> not then BarChart
        elif Ns.descendantsNs "pieChart" chartNode |> List.isEmpty |> not then PieChart
        elif Ns.descendantsNs "areaChart" chartNode |> List.isEmpty |> not then AreaChart
        elif Ns.descendantsNs "scatterChart" chartNode |> List.isEmpty |> not then ScatterChart
        else OtherChart "unknown"

    let parseSeries (serNode: XmlNode) : ChartSeries option =
        let name =
            Ns.descendantsNs "tx" serNode
            |> List.tryHead
            |> Option.bind (Ns.childNs "v")
            |> Option.bind Xml.text

        let values =
            Ns.descendantsNs "val" serNode
            |> List.tryHead
            |> Option.bind (Ns.childNs "f")
            |> Option.bind Xml.text

        let categories =
            Ns.descendantsNs "cat" serNode
            |> List.tryHead
            |> Option.bind (Ns.childNs "f")
            |> Option.bind Xml.text

        values |> Option.map (fun v -> {
            Name = name
            Values = v
            Categories = categories
        })

    let parseChart (chartXml: string) : Chart option =
        let doc = XmlDsl.parse chartXml
        let chartNode = Ns.childNs "chartSpace" doc

        chartNode |> Option.map (fun chart ->
            let title =
                Ns.descendantsNs "title" chart
                |> List.tryHead
                |> Option.bind (fun titleNode ->
                    Ns.descendantsNs "t" titleNode
                    |> List.tryHead)
                |> Option.bind Xml.text

            let chartType = parseChartType chart

            let series =
                Ns.descendantsNs "ser" chart
                |> List.choose parseSeries

            {
                Id = None
                Type = chartType
                Title = title
                Series = series
            })

// Pivot Tables parsing
module PivotTables =
    let parseField (fieldNode: XmlNode) : PivotTableField option =
        xml {
            let! name = Xml.attr "name" fieldNode
            let! idx = Xml.attr "fld" fieldNode |> Option.bind tryParseInt
            return {
                Name = name
                Index = idx
            }
        }

    let parsePivotTable (ptNode: XmlNode) : PivotTable option =
        Xml.attr "name" ptNode
        |> Option.map (fun name ->
            let dataSource = Xml.attr "dataSource" ptNode |> Option.defaultValue ""
            let location = Xml.attr "location" ptNode |> Option.defaultValue ""

            let rowFields =
                Ns.childNs "rowFields" ptNode
                |> Option.map (Ns.descendantsNs "field")
                |> Option.defaultValue []
                |> List.choose parseField

            let colFields =
                Ns.childNs "colFields" ptNode
                |> Option.map (Ns.descendantsNs "field")
                |> Option.defaultValue []
                |> List.choose parseField

            let dataFields =
                Ns.childNs "dataFields" ptNode
                |> Option.map (Ns.descendantsNs "dataField")
                |> Option.defaultValue []
                |> List.choose parseField

            {
                Name = name
                DataSource = dataSource
                Location = location
                RowFields = rowFields
                ColumnFields = colFields
                DataFields = dataFields
                FilterFields = []
            })

    let parsePivotTables (xml: string) : PivotTable list =
        let doc = XmlDsl.parse xml
        Ns.descendantsNs "pivotTableDefinition" doc
        |> List.choose parsePivotTable

// Apply styles to cells
let applyStyling (styles: CellStyle array) (sheet: Worksheet) : Worksheet =
    // Note: `<col style="N"/>` is intentionally NOT applied to cells that
    // exist in sheetData. Per the OOXML spec (§18.3.1.13), the column's
    // default style applies only to cells that are NOT present in
    // sheetData — i.e. completely missing `<c>` elements. Cells with a
    // `<c>` element use either their own `s=` or cellXf 0 (default).
    // Excel matches that exactly; our earlier attempt to apply col-style
    // as cell fallback over-painted col A / col C against Excel.
    // Cells without `s=` on the `<c>` element still use cellXfs[0]
    // (the default cell format) — this picks up the workbook's default
    // font, fill, number format from cellXfs[0]. Without this, plain
    // cells fell through to the renderer's hardcoded Calibri-11 default
    // and missed the workbook's chosen default font/size.
    let defaultStyle =
        if styles.Length > 0 then Some styles.[0] else None
    let applyToCell (cell: Cell) : Cell =
        match cell.StyleIndex with
        | Some idx when idx < styles.Length ->
            { cell with Style = Some styles.[idx] }
        | None when defaultStyle.IsSome ->
            { cell with Style = defaultStyle }
        | _ -> cell

    let applyToRow (row: Row) : Row =
        { row with Cells = row.Cells |> List.map applyToCell }

    { sheet with Rows = sheet.Rows |> List.map applyToRow }

// ---------------------------------------------------------------------------
// Workbook loading. One implementation, written against a synchronous part
// source: the native entry points read the archive directly, the JSZip entry
// points first await the parts each stage needs and then run the same code.
// ---------------------------------------------------------------------------

/// Text of package parts by path, plus an existence test for parts that
/// have not been fetched yet.
type private PartSource = {
    Exists: string -> bool
    Text: string -> string option
}

let private partDirectory (partPath: string) =
    match partPath.LastIndexOf '/' with
    | -1 -> ""
    | idx -> partPath.Substring(0, idx)

/// `xl/worksheets/sheet1.xml` -> `xl/worksheets/_rels/sheet1.xml.rels`.
let private relsPathFor (partPath: string) =
    let dir = partDirectory partPath
    if dir = "" then "_rels/" + partPath + ".rels"
    else dir + "/_rels/" + partPath.Substring(dir.Length + 1) + ".rels"

let private relationshipsOf (source: PartSource) (partPath: string) : Relationship list =
    source.Text (relsPathFor partPath)
    |> Option.map (XmlDsl.parse >> parseRelationships)
    |> Option.defaultValue []

let private isInternal (rel: Relationship) =
    rel.TargetMode <> Some "External"

/// Path of the part a relationship of `ownerPath` points at.
let private relTargetPath (ownerPath: string) (rel: Relationship) =
    ZipInterop.resolveTarget (partDirectory ownerPath) rel.Target

let private defaultWorkbookPath = "xl/workbook.xml"

/// The main part is whatever the package relationships say it is
/// (`_rels/.rels`, type …/officeDocument); `xl/workbook.xml` is only a
/// convention, kept as the fallback for packages without a usable root rel.
let private mainPartPath (source: PartSource) =
    let fromRels =
        try
            source.Text "_rels/.rels"
            |> Option.map (XmlDsl.parse >> parseRelationships)
            |> Option.defaultValue []
            |> List.tryFind (fun r -> r.Type.EndsWith "/officeDocument" && isInternal r)
            |> Option.map (fun r -> ZipInterop.resolveTarget "" r.Target)
        with _ -> None
    match fromRels with
    | Some path when source.Exists path -> path
    | _ -> defaultWorkbookPath

type private WorkbookHeader = {
    WorkbookPath: string
    Rels: Relationship list
    SheetInfos: Xlsx.SheetInfo list
    Date1904: bool
    NamedRanges: NamedRange list
    PrintNames: Map<int, WorksheetPrintNames>
    StylesPath: string
    SharedStringsPath: string
    ThemePath: string
}

let private readWorkbookHeader (source: PartSource) (workbookPath: string) : WorkbookHeader option =
    source.Text workbookPath
    |> Option.map (fun wbXml ->
        let wb = XmlDsl.parse wbXml
        let rels = relationshipsOf source workbookPath
        let partOfType (suffix: string) (fallback: string) =
            rels
            |> List.tryFind (fun r -> r.Type.EndsWith suffix && isInternal r)
            |> Option.map (relTargetPath workbookPath)
            |> Option.filter source.Exists
            |> Option.defaultValue (ZipInterop.resolveTarget (partDirectory workbookPath) fallback)
        let namedRanges, printNames = NamedRanges.parseNamedRanges wb |> NamedRanges.splitPrintNames
        {
            WorkbookPath = workbookPath
            Rels = rels
            SheetInfos = Xlsx.sheets wb
            Date1904 = Xlsx.date1904 wb
            NamedRanges = namedRanges
            PrintNames = printNames
            StylesPath = partOfType "/styles" "styles.xml"
            SharedStringsPath = partOfType "/sharedStrings" "sharedStrings.xml"
            ThemePath = partOfType "/theme" "theme/theme1.xml"
        })

/// The sheets to load, with their index in the workbook and their part path.
let private selectedSheets (header: WorkbookHeader) loadAll sheetName : (int * Xlsx.SheetInfo * string) list =
    sheetInfosForLoad loadAll sheetName header.SheetInfos
    |> List.choose (fun (sheetIdx, info) ->
        Office.relById info.RelationshipId header.Rels
        |> Option.map (fun rel -> sheetIdx, info, relTargetPath header.WorkbookPath rel))

/// Parts a sheet reaches through its own relationships: comments, tables
/// and pivot tables. Their paths come from the rels, never from a guess
/// based on the sheet number.
let private sheetDependencyPaths (source: PartSource) (sheetPath: string) : string list =
    relationshipsOf source sheetPath
    |> List.filter (fun r ->
        isInternal r
        && (r.Type.EndsWith "/comments" || r.Type.EndsWith "/table" || r.Type.EndsWith "/pivotTable"))
    |> List.map (relTargetPath sheetPath)

let private defaultThemeColors = [|
    "FFFFFF"; "000000"; "E7E6E6"; "44546A"
    "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47"
    "0563C1"; "954F72"
|]

let private buildWorkbook (source: PartSource) (header: WorkbookHeader) (sheetsToLoad: (int * Xlsx.SheetInfo * string) list) : Workbook =
    let sharedStringEntries =
        source.Text header.SharedStringsPath
        |> Option.map Xlsx.sharedStrings
        |> Option.defaultValue [||]

    // The theme is needed before styles so font/fill/border colors can
    // resolve theme="N" + tint references.
    let themeXml = source.Text header.ThemePath
    let themeColors =
        themeXml
        |> Option.map Xlsx.parseThemeColors
        |> Option.defaultValue defaultThemeColors
    let themeMajorFont, themeMinorFont =
        themeXml
        |> Option.map Xlsx.parseThemeFonts
        |> Option.defaultValue (None, None)
    let themeAccents =
        themeXml
        |> Option.map Xlsx.parseThemeAccents
        |> Option.defaultValue [| "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47" |]

    let stylesXml = source.Text header.StylesPath
    let styles =
        stylesXml
        |> Option.map (Xlsx.parseStyles themeColors themeMajorFont themeMinorFont)
        |> Option.defaultValue [||]
    let dxfs =
        stylesXml
        |> Option.map (XmlDsl.parse >> ConditionalFormats.parseDxfs themeColors)
        |> Option.defaultValue [||]

    let worksheets =
        sheetsToLoad
        |> List.choose (fun (sheetIdx, info, sheetPath) ->
            source.Text sheetPath
            |> Option.map (fun xml ->
                let printNames =
                    Map.tryFind sheetIdx header.PrintNames
                    |> Option.defaultValue emptyPrintNames
                let sheetRels = relationshipsOf source sheetPath
                let partsOfType (suffix: string) =
                    sheetRels
                    |> List.filter (fun r -> r.Type.EndsWith suffix && isInternal r)
                    |> List.choose (fun r -> source.Text (relTargetPath sheetPath r))
                let tables = partsOfType "/table" |> List.choose Xlsx.parseTable
                let comments = partsOfType "/comments" |> List.collect CellComments.parseComments
                let pivots =
                    partsOfType "/pivotTable"
                    |> List.collect (fun pivotXml -> try PivotTables.parsePivotTables pivotXml with _ -> [])
                // A fresh context per sheet: shared-formula groups are sheet-local.
                let ctx = Xlsx.newSheetContext (Some styles) header.Date1904
                let viaDom () =
                    let sheet = XmlDsl.parse xml
                    let condFormats = ConditionalFormats.parseConditionalFormats themeColors sheet
                    let dataVals = DataValidations.parseDataValidations sheet
                    Xlsx.worksheetFromNodeCore ctx info.Name info.State sheet comments condFormats dataVals [] pivots tables sheetRels printNames
#if FABLE_COMPILER
                viaDom ()
#else
                match Xlsx.worksheetNativeFastCore ctx info.Name info.State xml comments pivots tables sheetRels printNames with
                | Some sheet -> sheet
                | None -> viaDom ()
#endif
            ))

    {
        Sheets = worksheets
        SharedStrings = sharedStringEntries |> Array.map (fun e -> e.Text)
        SharedStringRuns = sharedStringEntries |> Array.map (fun e -> e.Runs)
        Styles = styles
        Dxfs = dxfs
        Relationships = header.Rels
        NamedRanges = header.NamedRanges
        ThemeAccents = themeAccents
        Date1904 = header.Date1904
    }

// JSZip entry: await the parts each stage needs into a cache, then run the
// shared synchronous stages over it.
let private parseWorkbookFiltered loadAll sheetName (archive: ZipArchive) : Fable.Core.JS.Promise<Workbook option> =
    let cache = System.Collections.Generic.Dictionary<string, string>()
    let source = {
        Exists = fun path -> (ZipInterop.tryFindFile archive path).IsSome
        Text = fun path ->
            match cache.TryGetValue path with
            | true, text -> Some text
            | _ -> None
    }
    let fetch (paths: string list) : Fable.Core.JS.Promise<unit> =
        let wanted =
            paths
            |> List.distinct
            |> List.filter (fun path -> not (cache.ContainsKey path))
            |> List.choose (fun path -> ZipInterop.tryFindFile archive path |> Option.map (fun file -> path, file))
            |> List.toArray
        wanted
        |> Array.map (fun (_, file) -> file.GetText())
        |> Promise.all
        |> Promise.map (fun texts ->
            texts |> Array.iteri (fun i text -> cache.[fst wanted.[i]] <- text))
    promise {
        do! fetch [ "_rels/.rels" ]
        let workbookPath = mainPartPath source
        do! fetch [ workbookPath; relsPathFor workbookPath ]
        match readWorkbookHeader source workbookPath with
        | None -> return None
        | Some header ->
            let sheetsToLoad = selectedSheets header loadAll sheetName
            let sheetPaths = sheetsToLoad |> List.map (fun (_, _, path) -> path)
            do! fetch ([ header.StylesPath; header.SharedStringsPath; header.ThemePath ]
                       @ sheetPaths
                       @ (sheetPaths |> List.map relsPathFor))
            do! fetch (sheetPaths |> List.collect (sheetDependencyPaths source))
            return Some (buildWorkbook source header sheetsToLoad)
    }

let parseWorkbook (archive: ZipArchive) : Fable.Core.JS.Promise<Workbook option> =
    parseWorkbookFiltered true None archive

let parseWorkbookSelected sheetName (archive: ZipArchive) : Fable.Core.JS.Promise<Workbook option> =
    parseWorkbookFiltered false sheetName archive

// Formula Evaluation Engine (stubbed - basic implementation for future)
module FormulaEngine =
    type EvalContext = {
        Workbook: Workbook
        CurrentSheet: Worksheet option
    }

    // Stub: Formula evaluation would parse and compute Excel formulas
    // This would include: arithmetic, functions (SUM, AVERAGE, IF, VLOOKUP, etc.),
    // cell references, range operations, and more.
    // Full implementation requires a lexer, parser, and evaluator for Excel formula syntax.
    let evaluateFormula (ctx: EvalContext) (formula: string) : float option =
        // Stub: Just try to parse as a number
        tryParseFloat formula

    // Stub: Evaluate a cell's formula
    let evaluateCell (wb: Workbook) (sheet: Worksheet) (cell: Cell) : CellValue =
        match cell.Value with
        | Formula f ->
            // Stub: Return formula as-is without evaluation
            cell.Value
        | other -> other

// Public API
let load (data: obj) : Fable.Core.JS.Promise<Workbook option> =
    promise {
        let! archive = loadArchive data
        return! parseWorkbook archive
    }

let loadSelected (data: obj) (sheetName: string option) : Fable.Core.JS.Promise<Workbook option> =
    promise {
        let! archive = loadArchive data
        return! parseWorkbookSelected sheetName archive
    }

// Resolve shared string references
let resolveValue (wb: Workbook) (value: CellValue) : CellValue =
    match value with
    | SharedString idx when idx < wb.SharedStrings.Length ->
        String wb.SharedStrings.[idx]
    | other -> other

// Extract sheet as 2D array
let sheetToArray (wb: Workbook) (sheet: Worksheet) : obj array array =
    let columnNumber (letters: string) =
        let mutable acc = 0
        for i = 0 to letters.Length - 1 do
            let ch = System.Char.ToUpperInvariant letters.[i]
            if ch >= 'A' && ch <= 'Z' then
                acc <- acc * 26 + (int ch - int 'A' + 1)
        acc

    let emptyCell = box ""
    let valueObject value =
        match resolveValue wb value with
        | String s -> box s
        | Number n -> box n
        | Boolean b -> box b
        | Date d -> box d
        | Formula f -> box f
        | _ -> emptyCell

    let mutable maxRow = 0
    let mutable maxCol = 0
    for row in sheet.Rows do
        if row.Index > maxRow then
            maxRow <- row.Index
        for cell in row.Cells do
            let col = columnNumber cell.Reference.Column
            if col > maxCol then
                maxCol <- col

    let result = Array.init maxRow (fun _ -> Array.create maxCol emptyCell)
    let seenRows = Array.zeroCreate<bool> maxRow
    for row in sheet.Rows do
        if row.Index > 0 && row.Index <= maxRow then
            let rowOffset = row.Index - 1
            if not seenRows.[rowOffset] then
                seenRows.[rowOffset] <- true
                let values = result.[rowOffset]
                for cell in row.Cells do
                    let col = columnNumber cell.Reference.Column
                    if col > 0 && col <= maxCol then
                        values.[col - 1] <- valueObject cell.Value

    result

#if !FABLE_COMPILER
// Synchronous native loader: the same stages as the JSZip entry, reading
// parts straight from the archive so native runners need no Promise scheduler.
let private parseWorkbookNativeFiltered loadAll sheetName (archive: ZipArchive) : Workbook option =
    let source = {
        Exists = fun path -> (ZipInterop.tryFindFile archive path).IsSome
        Text = ZipInterop.getFileTextNative archive
    }
    readWorkbookHeader source (mainPartPath source)
    |> Option.map (fun header -> buildWorkbook source header (selectedSheets header loadAll sheetName))

let parseWorkbookNative (archive: ZipArchive) : Workbook option =
    parseWorkbookNativeFiltered true None archive

let parseWorkbookNativeSelected sheetName (archive: ZipArchive) : Workbook option =
    parseWorkbookNativeFiltered false sheetName archive

/// Top-level native loader: bytes -> Workbook option.
let loadBytesNative (data: byte[]) : Workbook option =
    let archive = ZipInterop.loadArchiveNative data
    parseWorkbookNative archive

/// Top-level native selected-sheet loader: bytes -> Workbook option.
let loadBytesNativeSelected (data: byte[]) (sheetName: string option) : Workbook option =
    let archive = ZipInterop.loadArchiveNative data
    parseWorkbookNativeSelected sheetName archive
#endif
