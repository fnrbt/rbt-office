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
    /// The array formula is a dynamic array (the cell's `cm` names an XLDAPR
    /// entry of xl/metadata.xml): it spills. False: a legacy (Ctrl+Shift+Enter)
    /// array formula over `ArrayRef`.
    ArrayDynamic: bool
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
    /// `outlineLevel` (0 = in no group) and `collapsed` (the summary row of
    /// a collapsed group).
    OutlineLevel: int
    Collapsed: bool
}

type ColumnInfo = {
    Index: int
    MaxIndex: int
    Width: float option
    Hidden: bool
    /// `outlineLevel` (0 = in no group) and `collapsed`.
    OutlineLevel: int
    Collapsed: bool
    /// `<col style="N"/>` — default cellXf index for every cell in this
    /// column that doesn't carry its own `s="..."`. None when absent.
    Style: int option
}

// Advanced features
type CellComment = {
    CellRef: CellReference
    Author: string option
    Text: string
    /// The runs of a text written in runs (`<r>` with a font, a size, bold,
    /// a colour ...), in order; `Text` is their texts joined. Empty for a
    /// text without runs.
    Runs: (string * RunProps) list
}

/// One `<cfvo>` + `<color>` pair inside a `<colorScale>`. The cfvo
/// describes where on the value axis the stop sits (min / max /
/// percentile / num / percent); Color is the RRGGBB applied there.
type CfColorStop = {
    CfvoType: string
    CfvoValue: string option
    Color: string
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
    /// `<u/>` (any kind but "none") and `<strike/>` of the dxf's font.
    Underline: bool option
    Strike: bool option
}

/// The x14 form of a data bar (`<x14:dataBar>` in the sheet's extension
/// list), attribute by attribute with the schema's defaults.
type CfDataBarExt = {
    MinLength: int
    MaxLength: int
    ShowValue: bool
    Border: bool
    Gradient: bool
    /// "context", "leftToRight", "rightToLeft".
    Direction: string
    NegativeBarColorSameAsPositive: bool
    NegativeBarBorderColorSameAsPositive: bool
    /// "automatic", "middle", "none".
    AxisPosition: string
    /// The two ends: (type, value) with the x14 types (autoMin, autoMax,
    /// min, max, num, percent, percentile, formula).
    Cfvo: (string * string option) list
    BorderColor: string option
    NegativeFillColor: string option
    NegativeBorderColor: string option
    AxisColor: string option
}

/// What the x14 form of a conditional-formatting rule holds beyond the
/// `<cfRule>` (the `{78C0D931-…}` extension of the sheet).
type CfRuleExt = {
    DataBar: CfDataBarExt option
    /// Custom icons of an icon set (`<x14:cfIcon>`): (icon set, icon id) per step.
    IconCustom: (string * int) list
    /// `showValue` of an icon set.
    ShowValue: bool
    /// The rule is in the x14 list only: an icon set of Excel 2010+, custom
    /// icons, a rule whose formulas read another sheet (`<xm:f>`).
    InExtList: bool
    /// The format of such a rule: it has its `<x14:dxf>` in place, not an
    /// index into the styles part.
    Dxf: Dxf option
}

let emptyCfRuleExt : CfRuleExt = { DataBar = None; IconCustom = []; ShowValue = true; InExtList = false; Dxf = None }

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
    /// `iconSet` reverse flag — the icons run from best to worst.
    IconSetReverse: bool
    /// `timePeriod` of a "date occurring" rule ("today", "last7Days", ...).
    TimePeriod: string option
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
    /// The rule's x14 extension, when it has one.
    Ext: CfRuleExt option
}

type ConditionalFormat = {
    Range: string
    Rules: ConditionalFormatRule list
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
    /// The `comment` attribute (the name manager's comment).
    Comment: string option
    /// `hidden="1"`: not listed in the name manager.
    Hidden: bool
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
    /// What each column shows in the totals row, left to right: ("label",
    /// text), ("function", "sum" | "count" | ...), ("custom", formula) or
    /// ("", "") for nothing.
    Totals: (string * string) list
    Style: TableStyleInfo option
    /// (startCol, endCol) span of `<autoFilter ref="…">`, 1-indexed. None when
    /// the table has no autoFilter child. Excel renders a dropdown button on
    /// each header cell whose column falls in this range; row coords are
    /// implicit (header is always row StartRow).
    AutoFilter: (int * int) option
}

/// `<pane xSplit ySplit state="frozen" topLeftCell=…/>` inside a sheet's
/// What a `<sheetView>` says besides its panes, grid lines and zeros.
type SheetViewOptions = {
    /// `showRowColHeaders="0"` hides the row numbers and column letters.
    ShowRowColHeaders: bool
    /// `showFormulas="1"`: cells show their formulas instead of the results.
    ShowFormulas: bool
    /// `zoomScale`, in percent (100 when absent).
    ZoomScale: int
    /// The colour of the sheet's tab (`<sheetPr><tabColor/>`; it is kept here
    /// with the other things a sheet says about how it is shown): "#RRGGBB"
    /// for an rgb colour, else "theme:N" / "theme:N:tint" / "indexed:N" as
    /// the file has it. None = no colour.
    TabColor: string option
    /// `<sheetPr><outlinePr summaryBelow summaryRight/>`: whether the summary
    /// row of a group of rows is the one below it (else the one above) and
    /// the summary column of a group of columns the one to its right (else
    /// to its left). Both true when the file does not say.
    SummaryBelow: bool
    SummaryRight: bool
    /// `<sheetView showOutlineSymbols="0">`: the outline's brackets and
    /// buttons are not shown although rows / columns are grouped.
    ShowOutlineSymbols: bool
}

let defaultSheetView : SheetViewOptions = { ShowRowColHeaders = true; ShowFormulas = false; ZoomScale = 100; TabColor = None; SummaryBelow = true; SummaryRight = true; ShowOutlineSymbols = true }

/// The colour of a `<tabColor>` element from its attributes (see SheetViewOptions.TabColor).
let tabColorOfAttributes (rgb: string option) (theme: string option) (tint: string option) (indexed: string option) : string option =
    match rgb, theme, indexed with
    | Some value, _, _ when value.Length = 8 || value.Length = 6 -> Some ("#" + value.Substring(value.Length - 6).ToUpperInvariant())
    | _, Some index, _ when index <> "" ->
        (match tint with
         | Some tint when tint <> "" && tint <> "0" -> Some ("theme:" + index + ":" + tint)
         | _ -> Some ("theme:" + index))
    | _, _, Some index when index <> "" -> Some ("indexed:" + index)
    | _ -> None

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
    /// `tooltip` (the screen tip) and `display` (the text the link shows).
    Tooltip: string option
    Display: string option
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

/// The `<pageSetup>` attributes that are kept as read (see `PageSetup.Kept`).
let keptPageSetupAttributes =
    [ "paperHeight"; "paperWidth"; "usePrinterDefaults"; "draft"; "errors"; "horizontalDpi"; "verticalDpi"; "copies" ]

/// `<pageSetup>` plus `<sheetPr><pageSetUpPr fitToPage>`. PaperSize is
/// OOXML's enum (1 = Letter, 9 = A4, …); None when unspecified.
type PageSetup = {
    /// What `<pageSetup>` says about the printer and is not modelled
    /// (resolution, draft quality, how errors print, copies, a custom paper):
    /// (attribute, value) pairs, written back as they were read.
    Kept: (string * string) list
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

/// Where a drawing object sits on the sheet (xdr:twoCellAnchor / oneCellAnchor /
/// absoluteAnchor). Columns and rows are 0-based; offsets and sizes are EMU.
type SheetAnchor = {
    /// "twoCell" | "oneCell" | "absolute"
    Kind: string
    FromCol: int
    FromColOffset: float
    FromRow: int
    FromRowOffset: float
    /// The bottom-right corner of a two-cell anchor.
    ToCol: int option
    ToColOffset: float
    ToRow: int option
    ToRowOffset: float
    /// Position of an absolute anchor.
    X: float
    Y: float
    /// Size given by the anchor (one-cell, absolute) or by the object's own a:ext.
    Width: float option
    Height: float option
    /// editAs: twoCell | oneCell | absolute.
    EditAs: string option
    /// Position of the anchor in the drawing part: its z-order among the
    /// sheet's pictures and other objects (later is in front).
    Order: int
}

/// A picture anchored to cells.
type SheetImage = {
    Anchor: SheetAnchor
    Name: string option
    Description: string option
    /// image/png, image/jpeg, ...
    ContentType: string
    /// File extension without the dot.
    Extension: string
    Data: byte[]
}

/// A drawing object that is not a picture (a chart, a shape, a group, ...):
/// its anchor element as it is in the file, kept to be written back unchanged.
type SheetObject = {
    Anchor: SheetAnchor
    /// chart | shape | other
    Kind: string
    Name: string option
    /// The anchor element, with the namespace declarations it needs.
    Xml: string
}

/// A threaded comment (a comment with an author and replies).
type ThreadedComment = {
    CellRef: CellReference
    Id: string
    Author: string option
    /// The person entry (id, user id, provider id) behind the author.
    PersonId: string option
    Date: string option
    Text: string
    Replies: ThreadedComment list
    /// The thread is resolved (`done="1"`, on the first comment of a thread).
    Done: bool
}

/// A what-if data table (`<f t="dataTable" ref="B5:F10" dt2D="1" dtr="1" r1="B1" r2="B2"/>`
/// on the first cell of the range it fills). It is not an expression: Excel
/// recomputes the range by substituting the input cells. The cells keep the
/// results the file cached.
type DataTableFormula = {
    /// The range the table fills.
    Ref: string
    /// Two input cells (`dt2D`).
    TwoDimensional: bool
    /// A one-input table whose input runs along the row (`dtr`).
    RowInput: bool
    /// The input cells (`r1`, `r2`).
    Input1: string
    Input2: string option
    /// The input cell was deleted (`del1`, `del2`).
    Input1Deleted: bool
    Input2Deleted: bool
    /// `ca`: always calculated.
    CalculateAlways: bool
}

/// One `<filterColumn>` of a sheet's `<autoFilter>`.
type AutoFilterColumn = {
    /// The column's place in the filter's range, 0-based (`colId`).
    ColId: int
    /// "filters" (a list of values), "customFilters", "top10", "colorFilter",
    /// "dynamicFilter", or "other" for what is only kept (`Xml`): a list with
    /// date groups, an icon filter.
    Kind: string
    /// filters: the values, and whether blank cells pass.
    Values: string list
    Blank: bool
    /// customFilters: (operator, value) - "equal" | "notEqual" | "greaterThan"
    /// | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual" - joined by
    /// AND (`And`) or OR.
    Custom: (string * string) list
    And: bool
    /// top10: the largest (`Top`) or smallest `Value` items, or percent.
    Top: bool
    Percent: bool
    Value: float
    /// colorFilter: the differential format holding the colour, and whether
    /// it is the cell's fill (else its font colour).
    DxfId: int option
    CellColor: bool
    /// dynamicFilter: its `type` ("aboveAverage", "today", ...).
    DynamicType: string
    /// The children of the `<filterColumn>` as the file has them.
    Xml: string
}

/// A sheet's `<autoFilter>`: its range and the filtered columns.
type SheetAutoFilter = {
    Ref: string
    Columns: AutoFilterColumn list
}

let emptyAutoFilterColumn : AutoFilterColumn =
    { ColId = 0; Kind = "other"; Values = []; Blank = false; Custom = []; And = false; Top = true; Percent = false; Value = 10.0
      DxfId = None; CellColor = true; DynamicType = ""; Xml = "" }

/// For a writer: what of a sheet is not what its part of the file says, when
/// that is little enough for the part to be written again with only this
/// changed (see `SheetExtras.Patch`).
type SheetPatch = {
    /// The cells that are not the file's - changed, new, or gone: row and
    /// column (from 0), and whether the cell still holds the formula it had
    /// with another result (its `<f>` then stays the file's, shared or not).
    Cells: (int * int * bool) list
    /// The elements of the worksheet made from the sheet as it is now:
    /// "sheetPr", "sheetViews", "sheetFormatPr", "cols", "autoFilter",
    /// "sortState" (taken out), "mergeCells", "conditionalFormatting",
    /// "dataValidations", "printOptions", "pageMargins", "pageSetup",
    /// "headerFooter", "rowBreaks", "colBreaks".
    Sections: string list
    /// The rows whose own attributes are not the file's (see `SheetRowPatch`).
    Rows: SheetRowPatch list
    /// More changed than the part's own text says - comments, tables,
    /// drawings, links (they are in other parts, named by relationship):
    /// everything of the worksheet but its rows is then made from the sheet
    /// as it is now (with what the file's part says beyond what a sheet
    /// holds, see `SheetExtras.SourceXml`); the rows are the file's with
    /// `Cells` and `Rows` changed in them.
    Whole: bool
}

/// A row of a sheet whose attributes are to be those of the sheet as it is
/// now: hidden or not, its outline level and collapsed mark; and, with
/// `HeightChanged`, its height in points (None: the sheet's default height).
and SheetRowPatch = {
    Row: int
    HeightChanged: bool
    Height: float option
    Hidden: bool
    Level: int
    Collapsed: bool
}

/// What a sheet holds beyond its cells: pictures, other drawing objects and threaded comments.
/// A value that is worked out when it is first asked for (and then kept;
/// what it is worked out from is let go). Two of these are equal when their
/// values are.
[<AllowNullLiteral(false)>]
type OnDemand<'T when 'T : equality> (load: (unit -> 'T) option, known: 'T) =
    let mutable pending : (unit -> 'T) option = load
    let mutable value : 'T = known
    static member Of (value: 'T) : OnDemand<'T> = OnDemand<'T> (None, value)
    static member Deferred (load: unit -> 'T) : OnDemand<'T> = OnDemand<'T> (Some load, Unchecked.defaultof<'T>)
    member _.IsLoaded : bool = pending.IsNone
    member _.Value : 'T =
        match pending with
        | None -> value
        | Some work ->
            let found = work ()
            value <- found
            pending <- None
            found
    override this.Equals (other: obj) =
        match other with
        | :? OnDemand<'T> as that -> this.Value = that.Value
        | _ -> false
    override this.GetHashCode () = 0

/// The rows of a sheet. A reader that has the text of the sheet's part may
/// leave the rows unread until something asks for them (`Value`): a workbook
/// opens without the cost of the sheets nobody looks at. Two of these are
/// equal when their rows are.
[<AllowNullLiteral(false)>]
type SheetRows (load: (unit -> Row list) option, known: Row list, quickFormulas: (unit -> Cell list option) option, quickHeaders: (unit -> Row list option) option) =
    let mutable pending : (unit -> Row list) option = load
    let mutable rows : Row list = known
    let mutable quick : (unit -> Cell list option) option = quickFormulas
    let mutable formulas : Cell list option = None
    let mutable quickRows : (unit -> Row list option) option = quickHeaders
    let mutable headers : Row list option = None
    /// Rows that are already read.
    static member Of (rows: Row list) : SheetRows = SheetRows (None, rows, None, None)
    /// Rows read by `load` when they are first asked for; `quickFormulas`
    /// (when there is one) finds the cells with a formula without them, and
    /// `quickHeaders` the rows without their cells - or give None when they cannot.
    static member Deferred (load: unit -> Row list, quickFormulas: (unit -> Cell list option) option, quickHeaders: (unit -> Row list option) option) : SheetRows =
        SheetRows (Some load, [], quickFormulas, quickHeaders)
    /// What the rows say of themselves (index, height, hidden, outline) - the
    /// rows themselves when they are read, else the rows without their cells
    /// where that can be had without reading them. Rows with no height of
    /// their own that are neither hidden nor in an outline may be left out.
    member this.Headers : Row list =
        match pending, headers with
        | None, _ -> rows
        | Some _, Some known -> known
        | Some _, None ->
            match (match quickRows with Some scan -> scan () | None -> None) with
            | Some found ->
                headers <- Some found
                quickRows <- None
                found
            | None -> this.Value
    /// Whether the rows have been read.
    member _.IsLoaded : bool = pending.IsNone
    /// The rows (read now when they were not yet).
    member _.Value : Row list =
        match pending with
        | None -> rows
        | Some read ->
            let found = read ()
            rows <- found
            pending <- None
            quick <- None
            quickRows <- None
            headers <- None
            found
    /// The cells that hold a formula, in the order of the part: as `Value`
    /// has them, but found without reading every row when the rows are not
    /// read yet (a sheet of a million values and a dozen formulas).
    member this.FormulaCells : Cell list =
        match formulas with
        | Some known -> known
        | None ->
            let scanned =
                match pending, quick with
                | Some _, Some scan -> scan ()
                | _ -> None
            let found =
                match scanned with
                | Some cells -> cells
                | None ->
                    this.Value
                    |> List.collect (fun row ->
                        row.Cells |> List.filter (fun cell -> cell.Formula.IsSome || (match cell.Value with Formula _ -> true | _ -> false)))
            formulas <- Some found
            quick <- None
            found
    override this.Equals (other: obj) =
        match other with
        | :? SheetRows as rows -> this.Value = rows.Value
        | _ -> false
    override this.GetHashCode () = hash (List.length this.Value)

type SheetExtras = {
    Images: SheetImage list
    Objects: SheetObject list
    ThreadedComments: ThreadedComment list
    /// Relationships of the sheet's drawing part that the objects refer to
    /// (id, type, absolute target or external URL, external).
    DrawingRelationships: (string * string * string * bool) list
    /// The sheet's `<sheetFormatPr>` (default column width and row height),
    /// for the writer: None writes Excel's defaults. `XlsxWriter.
    /// packageFromReaderWorkbook` fills it in from `Worksheet.Format`.
    Format: SheetFormat option
    /// The sheet's `<col>` definitions as the file has them (ranges, default
    /// cell style), likewise filled in by `packageFromReaderWorkbook`.
    Columns: ColumnInfo list
    /// Default cell style of column ranges (`<col style>`): first and last
    /// column (0-based) and the style, likewise filled in by
    /// `packageFromReaderWorkbook`.
    ColumnStyles: (int * int * CellStyle) list
    /// The sheet's `sheetId` in workbook.xml. Other parts name a sheet by it
    /// (a slicer cache lists its pivot tables by `tabId`), so it is written back.
    SheetId: int option
    /// Relationships of the sheet part to parts the model does not hold
    /// (pivot tables, slicers, timelines): id, type, absolute target, external.
    KeptRelationships: (string * string * string * bool) list
    /// `<ext>` elements of the sheet's `<extLst>` that refer to those parts
    /// (the slicer list), with the namespace declarations they need.
    KeptExtensions: string list
    /// Elements of the sheet that name such parts (`keptSheetElementNames`:
    /// legacyDrawing with form controls, oleObjects, controls, the header
    /// pictures, custom properties ...) as (name, XML), written back as they are.
    KeptElements: (string * string) list
    /// The sheet's drawing part in the file. While the sheet's pictures and
    /// objects are the file's, that part goes back as it is, under its name.
    DrawingPart: string option
    /// The sheet's table parts in the file: (table name, part). A table that
    /// is as the file had it goes back as that part (with what the model
    /// does not hold of it: revision ids, per-column formats and formulas,
    /// its query table).
    TableParts: (string * string) list
    /// The sheet's pivot tables: part path and the range of `<location ref>`.
    PivotLocations: (string * string) list
    /// What-if data tables, in file order.
    DataTables: DataTableFormula list
    /// `<ext>` elements of the sheet's `<extLst>` that name cells of the
    /// sheet (sparkline groups, x14 conditional formats). They are written
    /// back while the sheet's rows and columns are where the file had them.
    KeptCellExtensions: string list
    /// The relationship id of the `<pageSetup>` element (the printer
    /// settings part, one of `KeptRelationships`).
    PrinterSettingsRelId: string option
    /// What rows / columns have besides a size, for the writer: 0-based
    /// index, the height (points) / width (characters) a hidden one keeps,
    /// the outline level, and the collapsed mark. `packageFromReaderWorkbook`
    /// fills them in from the sheet's rows and columns.
    RowAxis: (int * float option * int * bool) list
    ColumnAxis: (int * float option * int * bool) list
    /// The sheet's own AutoFilter (a table's is part of the table).
    AutoFilter: SheetAutoFilter option
    /// Array formulas for the writer: the cell that holds the formula and the
    /// range its results fill ("C1", "C1:C3"), written as `<f t="array" ref>`;
    /// true for a dynamic array (the cell gets the `cm` of the XLDAPR entry of
    /// xl/metadata.xml), false for a legacy Ctrl+Shift+Enter array formula.
    ArrayFormulas: (string * string * bool) list
    /// The part the sheet was read from and the type of the workbook's
    /// relationship to it (a worksheet, a chart sheet, a dialog sheet).
    SourcePart: (string * string) option
    /// The sheet's comments part in the file. While the sheet's comments are
    /// the file's, that part goes back as it is (rich text, shape ids,
    /// authors), also for a sheet that is otherwise written anew.
    CommentsPart: string option
    /// For the writer: the runs of the notes whose text is written in runs,
    /// by cell (row and column from 0). A note that is not here is written
    /// as one plain text.
    CommentRuns: ((int * int) * (string * RunProps) list) list
    /// For the writer, of a sheet of a file whose cells were not made yet:
    /// whether it has a formula the file holds no value for (None: not known,
    /// the cells say).
    UncachedFormulas: bool option
    /// The sheet is written as that part, byte for byte, with everything the
    /// part's own relationships lead to (the part's bytes are one of the
    /// workbook's `Preserved.Parts`). The reader says so for the sheets the
    /// editor has no model of (chart sheets, dialog sheets); a writer's caller
    /// says so for a worksheet nothing was changed in.
    Verbatim: bool
    /// With `Verbatim`: the sheet is written as that part with these cells
    /// and sections replaced by the sheet's (the rest of the part - sheetPr,
    /// the views, the rows and their attributes, shared formulas,
    /// ignoredErrors, the extensions - stays as it was read). A writer that
    /// cannot patch the part writes the sheet from its cells.
    Patch: SheetPatch option
    /// For a writer, of a sheet that is written from its cells: the text of
    /// the part it was read from. What that part says and a sheet does not
    /// hold is taken from it - the worksheet's namespaces, sheetPr (the code
    /// name), the views' selection, sheetFormatPr, the rows' own attributes
    /// (a row's style), the sheet's protection, ignoredErrors.
    SourceXml: string option
    /// With `SourceXml`: the rows and columns inserted and deleted since the
    /// part was read, in the order it was done - on rows (else columns), an
    /// insertion (else a deletion), the index (from 0) and the count. What
    /// is taken from the part by row or cell follows them. None: not known
    /// (nothing that names a row or cell is then taken from the part).
    Moves: (bool * bool * int * int) list option
}

let emptySheetExtras =
    { Images = []; Objects = []; ThreadedComments = []; DrawingRelationships = []; Format = None; Columns = []; ColumnStyles = []
      SheetId = None; KeptRelationships = []; KeptExtensions = []; KeptElements = []; DrawingPart = None; TableParts = []; PivotLocations = []; DataTables = []
      KeptCellExtensions = []; PrinterSettingsRelId = None; RowAxis = []; ColumnAxis = []; AutoFilter = None; ArrayFormulas = []
      SourcePart = None; CommentsPart = None; CommentRuns = []; UncachedFormulas = None; Verbatim = false; Patch = None; SourceXml = None; Moves = None }

/// A named cell style (cellStyles / cellStyleXfs of styles.xml), with the
/// font, fill, border and number format its xf refers to given as their XML so
/// that it does not depend on the indexes of the file it came from.
type NamedCellStyle = {
    Name: string
    /// Attributes of the cellStyle element other than name and xfId (builtinId, customBuiltin, ...).
    StyleAttributes: (string * string) list
    /// Attributes of the xf other than the font / fill / border / number format ids.
    XfAttributes: (string * string) list
    /// Children of the xf (alignment, protection).
    XfChildrenXml: string
    FontXml: string option
    FillXml: string option
    BorderXml: string option
    /// Built-in number format id (< 164), or the format code of a custom one.
    NumFmtId: int
    NumFmtCode: string option
}

/// Parts of the workbook package the model does not hold, kept to be written
/// back: the theme, the parts behind drawing objects (charts, their embedded
/// data), and the named cell styles.
type WorkbookPreservation = {
    /// (path, data) of parts written back as they are.
    Parts: (string * byte[]) list
    /// [Content_Types].xml entries of the original package.
    Defaults: (string * string) list
    Overrides: (string * string) list
    /// Path of the theme part, when the workbook has one.
    ThemePath: string option
    NamedStyles: NamedCellStyle list
    /// Persons of the threaded comments: (id, display name, user id, provider id).
    Persons: (string * string * string * string) list
    /// For each cell format of the file (cellXfs entry, the index a cell's
    /// `StyleIndex` holds): the index in `NamedStyles` of the named cell style
    /// it is based on (its xfId), or -1.
    CellXfNamedStyles: int array
    /// Relationships of the workbook part to kept parts (pivot caches, slicer
    /// caches, external links, connections): id, type, absolute target, external.
    WorkbookRelationships: (string * string * string * bool) list
    /// The elements of workbook.xml that refer to them, as they were:
    /// `<externalReferences>`, `<pivotCaches>` and `<extLst>` ("" when absent).
    ExternalReferencesXml: string
    PivotCachesXml: string
    WorkbookExtLstXml: string
    /// Relationships of the package to its document properties
    /// (docProps/core.xml, app.xml, custom.xml): id, type, absolute target, external.
    RootRelationships: (string * string * string * bool) list
    /// The workbook's relationship to its calculation chain. The chain lists
    /// the formula cells: a writer keeps it (and this relationship) only when
    /// no formula cell was added, removed or moved; Excel builds a new one
    /// for a file without.
    CalcChain: (string * string * string * bool) option
    /// The cell ranges pivot caches read (`<worksheetSource ref sheet>`):
    /// cache definition part, sheet name, range.
    PivotCacheSources: (string * string * string) list
    /// Iterative calculation of circular references (`<calcPr iterate="1">`):
    /// most iterations and the largest change that ends them. None: off.
    CalcIteration: (int * float) option
    /// The workbook is in the 1904 date system (`workbookPr date1904`): a
    /// workbook part made anew says so.
    Date1904: bool
    /// The sheet that was on display when the file was saved (`<workbookView
    /// activeTab>`, an index into the workbook's sheets).
    ActiveTab: int
    /// The styles and shared strings parts of the file, and what the file's
    /// indexes into them mean: its cell formats by cellXfs index (the
    /// workbook's `Styles`), its shared strings by index (texts and runs,
    /// the workbook's own arrays), the number of its dxfs.
    StylesPart: string option
    SharedStringsPart: string option
    /// The workbook part (among `SourceParts`): a writer puts it back with its
    /// sheet list, defined names, active tab and calculation settings made
    /// anew and everything else as it is.
    WorkbookPart: string option
    CellFormats: CellStyle array
    SharedStrings: string array * (string * RunProps) list array
    DxfCount: int
    /// Parts of the file a writer needs the bytes of to write a sheet as it
    /// was read beside sheets it makes: the worksheet parts, the styles and
    /// the shared strings, and what the worksheets' relationships lead to
    /// that is not among `Parts` (comments, their boxes, threaded comments),
    /// with those relationship parts. A reader that does not hold the bytes
    /// of the large parts (the browser's: worksheets, shared strings) leaves
    /// those out; a caller then supplies them.
    SourceParts: (string * byte[]) list
    /// The parts that hold the package together, as the file had them, by
    /// path: `[Content_Types].xml`, `_rels/.rels` and the workbook's
    /// relationships. A writer keeps them as they are while they still say
    /// what is written, and otherwise takes entries out and adds entries.
    PackageTexts: (string * string) list
    /// For a writer: the styles and shared strings are the file's parts
    /// (`StylesPart` and `SharedStringsPart`, among `Parts`) with what is new
    /// added behind what they have, so that every index the file uses means
    /// what it meant. Sheets marked `SheetExtras.Verbatim` need it.
    InPlace: bool
}

let emptyWorkbookPreservation =
    { Parts = []; Defaults = []; Overrides = []; ThemePath = None; NamedStyles = []; Persons = []; CellXfNamedStyles = [||]
      WorkbookRelationships = []; ExternalReferencesXml = ""; PivotCachesXml = ""; WorkbookExtLstXml = ""
      RootRelationships = []; CalcChain = None; PivotCacheSources = []; CalcIteration = None; Date1904 = false; ActiveTab = 0
      StylesPart = None; SharedStringsPart = None; WorkbookPart = None; CellFormats = [||]; SharedStrings = [||], [||]; DxfCount = 0; SourceParts = []; PackageTexts = []; InPlace = false }

type Worksheet = {
    Name: string
    /// Tab visibility, from the workbook's `<sheet state>`.
    State: SheetState
    /// The sheet's rows (see `Rows`): read from the part when first asked for.
    RowSource: SheetRows
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
    /// Headers, formula view and zoom of the first `<sheetView>`.
    View: SheetViewOptions
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
    /// Pictures, charts and other drawing objects, and threaded comments.
    Extras: SheetExtras
}
with
    /// The sheet's rows. (Reads them from the part when that was left for later.)
    member this.Rows : Row list = this.RowSource.Value
    /// The cells with a formula, without reading the other rows where the reader can tell.
    member this.FormulaCells : Cell list = this.RowSource.FormulaCells

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
    /// What the package holds beyond the model (see WorkbookPreservation).
    Preserved: WorkbookPreservation
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
/// xl/metadata.xml: the cell-metadata records (the 1-based values of `<c cm>`)
/// that mark a dynamic array formula. Such a record points
/// (`<rc t v>`) into the future metadata of type XLDAPR at a block whose
/// `xda:dynamicArrayProperties` has `fDynamic="1"`.
let dynamicArrayCellMetadata (metadataXml: string) : Set<int> =
    let doc = XmlDsl.parse metadataXml
    let named (local: string) (node: XmlNode) =
        match node with
        | XmlDsl.Element (name, _, _) -> Ns.stripNs name = local
        | _ -> false
    let kids (local: string) (node: XmlNode) = Xml.children node |> List.filter (named local)
    let types =
        kids "metadataTypes" doc
        |> List.collect (kids "metadataType")
        |> List.map (fun node -> Xml.attr "name" node |> Option.defaultValue "")
        |> List.toArray
    let dynamicBlocks =
        kids "futureMetadata" doc
        |> List.filter (fun node -> Xml.attr "name" node = Some "XLDAPR")
        |> List.collect (kids "bk")
        |> List.map (fun block ->
            Ns.descendantsNs "dynamicArrayProperties" block
            |> List.exists (fun properties -> Xml.attr "fDynamic" properties |> Option.exists isTrueText))
        |> List.toArray
    kids "cellMetadata" doc
    |> List.collect (kids "bk")
    |> List.mapi (fun index block ->
        let dynamic =
            kids "rc" block
            |> List.exists (fun record ->
                match Xml.attr "t" record |> Option.bind tryParseInt, Xml.attr "v" record |> Option.bind tryParseInt with
                | Some t, Some v ->
                    t >= 1 && t <= types.Length && types.[t - 1] = "XLDAPR" && v >= 0 && v < dynamicBlocks.Length && dynamicBlocks.[v]
                | _ -> false)
        if dynamic then Some (index + 1) else None)
    |> List.choose id
    |> Set.ofList

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
        // (<u/> is a single underline; <u val="none"/> says there is none.)
        let underline = boolChild "u" && (sElem "u" rPr |> Option.bind (Xml.attr "val")) <> Some "none"
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
                        // (<u val="none"/> says there is no underline.)
                        | "none" -> underline <- false
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
        /// What-if data tables seen so far.
        DataTables: ResizeArray<DataTableFormula>
        /// The cell-metadata indices (`<c cm>`, 1-based) that mark a dynamic
        /// array formula; see `dynamicArrayCellMetadata`.
        mutable DynamicCm: Set<int>
    }

    /// Days between the 1904 and 1900 date systems' epochs.
    let internal date1904Offset = 1462.0

    let internal newSheetContext (styles: CellStyle array option) (date1904: bool) : SheetParseContext =
        {
            Styles = styles
            // (The serials of a workbook in the 1904 date system stay what the
            // file has: `Workbook.Date1904` says which system they are in.)
            DateShift = (ignore date1904; [||])
            SharedFormulas = System.Collections.Generic.Dictionary<string, int * int * string>()
            DataTables = ResizeArray<DataTableFormula>()
            DynamicCm = Set.empty
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
                          (fAttr: string -> string option)
                          (value: string option) (inlineText: string option) (inlineRuns: (string * RunProps) list) : Cell =
        let arrayDynamic =
            fType = Some "array" && not ctx.DynamicCm.IsEmpty
            && (match fAttr "cm" |> Option.bind tryParseInt with
                | Some cm -> ctx.DynamicCm.Contains cm
                | None -> false)
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
                // references, not an expression. The cell keeps the cached
                // result; the table itself is recorded for the sheet.
                | Some "dataTable" ->
                    let flag name = fAttr name |> Option.map isTrueText |> Option.defaultValue false
                    ctx.DataTables.Add
                        { Ref = fRef |> Option.defaultValue (column + string rowIndex)
                          TwoDimensional = flag "dt2D"
                          RowInput = flag "dtr"
                          Input1 = fAttr "r1" |> Option.defaultValue ""
                          Input2 = fAttr "r2"
                          Input1Deleted = flag "del1"
                          Input2Deleted = flag "del2"
                          CalculateAlways = flag "ca" }
                    None, None
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
            ArrayDynamic = arrayDynamic && arrayRef.IsSome
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
                (fun name -> if name = "cm" then Xml.attr "cm" cNode else fNode |> Option.bind (Xml.attr name))
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
                                        ArrayDynamic = false
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
            OutlineLevel = Xml.attr "outlineLevel" rowNode |> Option.bind tryParseInt |> Option.defaultValue 0
            Collapsed = Xml.attr "collapsed" rowNode |> Option.map isTrueText |> Option.defaultValue false
        }

    // Row extraction
    let rowWithStyles (styles: CellStyle array option) (rowNode: XmlNode) : Row option =
        Xml.attr "r" rowNode
        |> Option.map (fun _ -> rowFromNode (newSheetContext styles false) 0 rowNode)

    let row rowNode = rowWithStyles None rowNode

    // ---- <sheetData> read straight from the text of the part -------------------
    // A large sheet is almost all <sheetData>: rows of cells with a formula and
    // a value. Building a DOM and then a node tree for it costs several times
    // the part's size in memory and most of the load time in a browser. The
    // rows are therefore scanned from the text; everything else of the sheet
    // (columns, merges, rules, page setup) still goes through the XML parser,
    // on the part with its rows cut out. Cells are assembled by `buildCell`,
    // as in the other two parsers.

    /// Decode the five predefined entities and character references.
    let private decodeXmlText (text: string) : string =
        if text.IndexOf '&' < 0 then text
        else
            let sb = StringBuilder(text.Length)
            let mutable i = 0
            while i < text.Length do
                let ch = text.[i]
                if ch = '&' then
                    let close = text.IndexOf(';', i)
                    if close < 0 || close - i > 10 then
                        sb.Append ch |> ignore
                        i <- i + 1
                    else
                        let name = text.Substring(i + 1, close - i - 1)
                        let replacement =
                            match name with
                            | "amp" -> "&"
                            | "lt" -> "<"
                            | "gt" -> ">"
                            | "quot" -> "\""
                            | "apos" -> "'"
                            | _ when name.StartsWith "#" ->
                                let code =
                                    if name.Length > 2 && (name.[1] = 'x' || name.[1] = 'X') then
                                        let mutable value = 0
                                        let mutable valid = true
                                        for k in 2 .. name.Length - 1 do
                                            let c = name.[k]
                                            if c >= '0' && c <= '9' then value <- value * 16 + (int c - int '0')
                                            elif c >= 'a' && c <= 'f' then value <- value * 16 + (int c - int 'a' + 10)
                                            elif c >= 'A' && c <= 'F' then value <- value * 16 + (int c - int 'A' + 10)
                                            else valid <- false
                                        if valid then value else -1
                                    else
                                        match tryParseInt (name.Substring 1) with
                                        | Some value -> value
                                        | None -> -1
                                if code < 0 then "&" + name + ";"
                                elif code < 0x10000 then string (char code)
                                else
                                    // A character beyond the BMP is a surrogate pair.
                                    let offset = code - 0x10000
                                    string (char (0xD800 + (offset >>> 10))) + string (char (0xDC00 + (offset &&& 0x3FF)))
                            | _ -> "&" + name + ";"
                        sb.Append replacement |> ignore
                        i <- close + 1
                else
                    sb.Append ch |> ignore
                    i <- i + 1
            sb.ToString()

    // The scanner works on character codes and reads the attributes of a tag
    // in one pass into a small array: it runs over tens of megabytes of text,
    // and it is kept in small functions of one shape each so that a JavaScript
    // engine compiles (and keeps) them optimised.

    let inline private codeAt (text: string) (index: int) : int = int text.[index]

    let inline private isSpaceCode (code: int) = code = 32 || code = 10 || code = 13 || code = 9

    /// Ordinal search from `start` (the default search of .NET is by culture, and slow).
#if FABLE_COMPILER
    let inline private indexOfFrom (text: string) (value: string) (start: int) : int = text.IndexOf(value, start)
#else
    let inline private indexOfFrom (text: string) (value: string) (start: int) : int = text.IndexOf(value, start, System.StringComparison.Ordinal)
#endif

    /// Read the attributes of a start tag (`xml.[from .. upTo - 1]`, the text
    /// after the element name) once: `values.[i]` becomes the value of the
    /// attribute `names.[i]`, or null when the tag does not have it.
    let private scanAttributes (xml: string) (from: int) (upTo: int) (names: string array) (values: string array) : unit =
        for k in 0 .. values.Length - 1 do
            values.[k] <- null
        let mutable i = from
        while i < upTo do
            while i < upTo && isSpaceCode (codeAt xml i) do i <- i + 1
            let nameStart = i
            let mutable inName = true
            while inName && i < upTo do
                let code = codeAt xml i
                if code = 61 || code = 47 || isSpaceCode code then inName <- false else i <- i + 1
            let nameEnd = i
            while i < upTo && isSpaceCode (codeAt xml i) do i <- i + 1
            if i < upTo && codeAt xml i = 61 then
                i <- i + 1
                while i < upTo && isSpaceCode (codeAt xml i) do i <- i + 1
                if i < upTo then
                    let quote = xml.[i]
                    let valueStart = i + 1
                    let valueEnd = xml.IndexOf(quote, valueStart)
                    if valueEnd < 0 || valueEnd >= upTo then i <- upTo
                    else
                        let length = nameEnd - nameStart
                        let mutable k = 0
                        while k < names.Length do
                            let name = names.[k]
                            let mutable same = name.Length = length
                            let mutable j = 0
                            while same && j < length do
                                if codeAt xml (nameStart + j) <> codeAt name j then same <- false
                                j <- j + 1
                            if same then
                                values.[k] <- decodeXmlText (xml.Substring(valueStart, valueEnd - valueStart))
                                k <- names.Length
                            else k <- k + 1
                        i <- valueEnd + 1
            else
                i <- i + 1

    /// A run of decimal digits as a number, or -1.
    let private digitsValue (text: string) : int =
        if isNull text || text.Length = 0 || text.Length > 9 then -1
        else
            let mutable value = 0
            let mutable i = 0
            while value >= 0 && i < text.Length do
                let code = codeAt text i
                if code >= 48 && code <= 57 then value <- value * 10 + (code - 48) else value <- -1
                i <- i + 1
            value

    /// An integer attribute as the XML parsers read it (`tryParseInt`).
    let private intAttribute (text: string) : int option =
        if isNull text then None
        else
            match digitsValue text with
            | -1 -> tryParseInt text
            | value -> Some value

    let private boolAttribute (text: string) : bool option =
        if isNull text then None else Some (isTrueText text)

    /// Text content of an element as the XML parser hands it over: line ends
    /// normalised, entities decoded, and trimmed unless the element says
    /// xml:space="preserve". `tagFrom .. tagEnd` is the text of the start tag
    /// after the element name.
    let private scanText (xml: string) (tagFrom: int) (tagEnd: int) (contentEnd: int) : string =
        let raw = xml.Substring(tagEnd + 1, contentEnd - tagEnd - 1)
        // An XML parser reads every line end (CR LF, CR) as LF; a character
        // reference (&#13;) is the way to say CR.
        let raw = if raw.IndexOf '\r' < 0 then raw else raw.Replace("\r\n", "\n").Replace("\r", "\n")
        let decoded = decodeXmlText raw
        let preserve = tagEnd > tagFrom && indexOfFrom (xml.Substring(tagFrom, tagEnd - tagFrom)) "preserve" 0 >= 0
        if preserve then decoded else decoded.Trim()

    /// What a scan of a sheet's rows carries from cell to cell.
    type private RowScan = {
        Xml: string
        Ctx: SheetParseContext
        ColumnNames: System.Collections.Generic.Dictionary<int, string>
        RowNames: string array
        RowValues: string array
        CellNames: string array
        CellValues: string array
        FormulaNames: string array
        FormulaValues: string array
        /// Where the next inline string (`<is`) is, or -1 when there is none ahead.
        mutable NextInline: int
        mutable Valid: bool
        mutable PrevCol: int
    }

    let private scanColumnName (scan: RowScan) (index: int) : string =
        match scan.ColumnNames.TryGetValue index with
        | true, name -> name
        | false, _ ->
            let name = indexToColumnLetters index
            scan.ColumnNames.[index] <- name
            name

    /// The `<f>` and `<v>` children of a cell: (formula text or null, its tag
    /// bounds, value text or null). Other children are skipped.
    let private scanCellChildren (scan: RowScan) (from: int) (cellEnd: int) (formulaTag: int array) : string * string =
        let xml = scan.Xml
        let mutable formula : string = null
        let mutable value : string = null
        let mutable childPos = from
        while childPos < cellEnd do
            let childAt = xml.IndexOf('<', childPos)
            if childAt < 0 || childAt >= cellEnd - 1 then childPos <- cellEnd
            else
                let childTagEnd = xml.IndexOf('>', childAt)
                if childTagEnd < 0 || childTagEnd > cellEnd then childPos <- cellEnd
                else
                    let selfClosed = codeAt xml (childTagEnd - 1) = 47
                    let nameCode = codeAt xml (childAt + 1)
                    let afterCode = codeAt xml (childAt + 2)
                    let single = isSpaceCode afterCode || afterCode = 62 || afterCode = 47
                    // 102 'f', 118 'v'
                    if single && (nameCode = 102 || nameCode = 118) then
                        let tagAttrEnd = if selfClosed then childTagEnd - 1 else childTagEnd
                        let contentEnd =
                            if selfClosed then childTagEnd
                            else
                                match indexOfFrom xml (if nameCode = 102 then "</f>" else "</v>") childTagEnd with
                                | -1 -> cellEnd
                                | index -> min index cellEnd
                        let text = if selfClosed then "" else scanText xml (childAt + 2) tagAttrEnd contentEnd
                        if nameCode = 102 then
                            formula <- text
                            formulaTag.[0] <- childAt + 2
                            formulaTag.[1] <- tagAttrEnd
                        else value <- text
                        childPos <- if selfClosed then childTagEnd + 1 else contentEnd + 4
                    elif nameCode = 47 || selfClosed then childPos <- childTagEnd + 1
                    else
                        // An element this reader has no use for (extLst): skip over it.
                        let mutable nameEnd = childAt + 1
                        while nameEnd < childTagEnd && not (isSpaceCode (codeAt xml nameEnd)) do nameEnd <- nameEnd + 1
                        let closing = "</" + xml.Substring(childAt + 1, nameEnd - childAt - 1) + ">"
                        match indexOfFrom xml closing childTagEnd with
                        | -1 -> childPos <- cellEnd
                        | index -> childPos <- index + closing.Length
        formula, value

    /// One `<c>` that has no inline string: `cellAt` is its "<", `cellTagEnd`
    /// the ">" of its start tag, `cellEnd` its "</c>" (or `cellTagEnd` when
    /// the tag closes itself).
    let private scanCell (scan: RowScan) (rowIndex: int) (cellAt: int) (cellTagEnd: int) (cellEnd: int) (selfClosed: bool) : Cell =
        let xml = scan.Xml
        scanAttributes xml (cellAt + 2) (if selfClosed then cellTagEnd - 1 else cellTagEnd) scan.CellNames scan.CellValues
        let reference = scan.CellValues.[0]
        let mutable colIndex = 0
        let mutable row = 0
        if not (isNull reference) then
            for i in 0 .. reference.Length - 1 do
                let code = codeAt reference i
                if code >= 65 && code <= 90 then colIndex <- colIndex * 26 + (code - 64)
                elif code >= 97 && code <= 122 then colIndex <- colIndex * 26 + (code - 96)
                elif code >= 48 && code <= 57 then row <- row * 10 + (code - 48)
        if colIndex = 0 then
            // No (usable) reference: one column right of the previous cell.
            colIndex <- scan.PrevCol + 1
            row <- rowIndex
        elif row <= 0 then row <- rowIndex
        scan.PrevCol <- colIndex
        let cellType = match scan.CellValues.[2] with | null -> None | text -> Some text
        let styleIdx = intAttribute scan.CellValues.[1]
        if selfClosed then
            buildCell scan.Ctx (scanColumnName scan colIndex) colIndex row cellType styleIdx None None None None (fun _ -> None) None None []
        else
            let formulaTag = [| 0; 0 |]
            let formula, value = scanCellChildren scan (cellTagEnd + 1) cellEnd formulaTag
            if isNull formula then
                buildCell scan.Ctx (scanColumnName scan colIndex) colIndex row cellType styleIdx None None None None (fun _ -> None)
                    (if isNull value then None else Some value) None []
            else
                scanAttributes xml formulaTag.[0] formulaTag.[1] scan.FormulaNames scan.FormulaValues
                let optional (text: string) = if isNull text then None else Some text
                let fType = optional scan.FormulaValues.[0]
                let fRef = optional scan.FormulaValues.[1]
                let fSi = optional scan.FormulaValues.[2]
                // The other attributes are only asked for by a data table.
                // The cell's own `cm` was read with its other attributes.
                let cellCm = optional scan.CellValues.[3]
                let fAttr (name: string) =
                    if name = "cm" then cellCm
                    else
                        let names = [| name |]
                        let values : string array = [| null |]
                        scanAttributes xml formulaTag.[0] formulaTag.[1] names values
                        optional values.[0]
                buildCell scan.Ctx (scanColumnName scan colIndex) colIndex row cellType styleIdx (Some formula) fType fRef fSi fAttr
                    (if isNull value then None else Some value) None []

    /// The cells of one row: `from` is just after the row's start tag, `rowEnd` its "</row>".
    let private scanRowCells (scan: RowScan) (rowIndex: int) (from: int) (rowEnd: int) : Cell list =
        let xml = scan.Xml
        let cells = System.Collections.Generic.List<Cell>()
        scan.PrevCol <- 0
        let mutable cellPos = from
        while scan.Valid && cellPos < rowEnd do
            let cellAt = indexOfFrom xml "<c" cellPos
            if cellAt < 0 || cellAt >= rowEnd then cellPos <- rowEnd
            else
                let cellTagEnd = xml.IndexOf('>', cellAt)
                let next = codeAt xml (cellAt + 2)
                if cellTagEnd < 0 || cellTagEnd > rowEnd || not (isSpaceCode next || next = 62 || next = 47) then scan.Valid <- false
                else
                    let selfClosed = codeAt xml (cellTagEnd - 1) = 47
                    let cellEnd = if selfClosed then cellTagEnd else indexOfFrom xml "</c>" cellTagEnd
                    if cellEnd < 0 || cellEnd > rowEnd then scan.Valid <- false
                    else
                        // Where the next inline string is: looked for again only once it is passed.
                        if scan.NextInline >= 0 && scan.NextInline < cellAt then scan.NextInline <- indexOfFrom xml "<is" cellAt
                        if not selfClosed && scan.NextInline >= 0 && scan.NextInline < cellEnd then
                            // An inline string (plain or rich) is rare: the XML parser reads this one cell.
                            let cell, colIndex = cellFromNode scan.Ctx rowIndex scan.PrevCol (XmlDsl.parse (xml.Substring(cellAt, cellEnd + 4 - cellAt)))
                            scan.PrevCol <- colIndex
                            cells.Add cell
                        else
                            cells.Add (scanCell scan rowIndex cellAt cellTagEnd cellEnd selfClosed)
                        cellPos <- if selfClosed then cellTagEnd + 1 else cellEnd + 4
        List.ofSeq cells

    /// The rows of a sheet part, from its text. None when the part is not in
    /// the plain shape this scanner reads (a prefixed sheetData element,
    /// comments or CDATA among the rows): the caller then parses the whole part.
    /// Also returns the part with the rows cut out.
    let private scanSheetRowsWith (headersOnly: bool) (ctx: SheetParseContext) (sheetXml: string) : (Row list * string) option =
        let openAt = indexOfFrom sheetXml "<sheetData" 0
        if openAt < 0 then None
        else
            let openEnd = sheetXml.IndexOf('>', openAt)
            if openEnd < 0 then None
            elif codeAt sheetXml (openEnd - 1) = 47 then Some ([], sheetXml)
            else
                let closeAt = indexOfFrom sheetXml "</sheetData>" openEnd
                let commentAt = if closeAt < 0 then -1 else indexOfFrom sheetXml "<!" openEnd
                if closeAt < 0 || (commentAt >= 0 && commentAt < closeAt) then None
                else
                    let scan =
                        { Xml = sheetXml
                          Ctx = ctx
                          ColumnNames = System.Collections.Generic.Dictionary<int, string>()
                          RowNames = [| "r"; "ht"; "customHeight"; "hidden"; "s"; "customFormat"; "spans"; "outlineLevel"; "collapsed" |]
                          RowValues = Array.create 9 null
                          CellNames = [| "r"; "s"; "t"; "cm" |]
                          CellValues = Array.create 4 null
                          FormulaNames = [| "t"; "ref"; "si" |]
                          FormulaValues = Array.create 3 null
                          NextInline = indexOfFrom sheetXml "<is" openEnd
                          Valid = true
                          PrevCol = 0 }
                    let rows = System.Collections.Generic.List<Row>()
                    let mutable prevRow = 0
                    let mutable pos = openEnd + 1
                    while scan.Valid && pos < closeAt do
                        let rowAt = indexOfFrom sheetXml "<row" pos
                        if rowAt < 0 || rowAt >= closeAt then pos <- closeAt
                        else
                            let rowTagEnd = sheetXml.IndexOf('>', rowAt)
                            let afterName = codeAt sheetXml (rowAt + 4)
                            if rowTagEnd < 0 || rowTagEnd > closeAt || not (isSpaceCode afterName || afterName = 62 || afterName = 47) then scan.Valid <- false
                            else
                                let selfClosed = codeAt sheetXml (rowTagEnd - 1) = 47
                                scanAttributes sheetXml (rowAt + 4) (if selfClosed then rowTagEnd - 1 else rowTagEnd) scan.RowNames scan.RowValues
                                let rowIndex =
                                    match intAttribute scan.RowValues.[0] with
                                    | Some r when r > 0 -> r
                                    | _ -> prevRow + 1
                                prevRow <- rowIndex
                                let rawHeight = match scan.RowValues.[1] with | null -> None | text -> tryParseFloat text
                                let customHeightAttr = boolAttribute scan.RowValues.[2]
                                let hidden = boolAttribute scan.RowValues.[3] |> Option.defaultValue false
                                let rowStyle = intAttribute scan.RowValues.[4]
                                let customFormat = boolAttribute scan.RowValues.[5] |> Option.defaultValue false
                                let spans = match scan.RowValues.[6] with | null -> None | text -> Some text
                                let rowEnd = if selfClosed then rowTagEnd else indexOfFrom sheetXml "</row>" rowTagEnd
                                if rowEnd < 0 || rowEnd > closeAt then scan.Valid <- false
                                else
                                    let cells = if headersOnly || selfClosed then [] else scanRowCells scan rowIndex (rowTagEnd + 1) rowEnd
                                    if scan.Valid then
                                        rows.Add
                                            { Index = rowIndex
                                              Cells = (if headersOnly then [] else rowSpanCells ctx rowIndex rowStyle customFormat spans cells)
                                              Height = (if customHeightAttr = Some true then rawHeight else None)
                                              RawHeight = rawHeight
                                              CustomHeight = customHeightAttr
                                              Hidden = hidden
                                              OutlineLevel = intAttribute scan.RowValues.[7] |> Option.defaultValue 0
                                              Collapsed = boolAttribute scan.RowValues.[8] |> Option.defaultValue false }
                                    pos <- if selfClosed then rowTagEnd + 1 else rowEnd + 6
                    if not scan.Valid then None
                    elif headersOnly then Some (List.ofSeq rows, "")
                    else Some (List.ofSeq rows, sheetXml.Substring(0, openAt) + "<sheetData/>" + sheetXml.Substring(closeAt + 12))

    // Merged cells extraction
    let internal scanSheetRows (ctx: SheetParseContext) (sheetXml: string) : (Row list * string) option =
        scanSheetRowsWith false ctx sheetXml

    /// The rows without their cells (what `scanSheetRows` gives, less the cells): see `SheetRows.Headers`.
    /// Rows that say nothing of themselves may be left out - all of them are
    /// when no row of the part has a height, is hidden or is in an outline.
    let internal scanRowHeaders (ctx: SheetParseContext) (sheetXml: string) : Row list option =
        let openAt = indexOfFrom sheetXml "<sheetData" 0
        let closeAt = if openAt < 0 then -1 else indexOfFrom sheetXml "</sheetData>" openAt
        let within (attribute: string) =
            let at = indexOfFrom sheetXml attribute openAt
            at >= 0 && at < closeAt
        // (A cell has none of these attributes; text holds no quote that is not escaped... but for a
        // quote in a formula or a value, which then only costs the reading of the rows' tags.)
        if openAt >= 0 && closeAt >= 0 && not (within " ht=\"" || within " hidden=\"" || within " outlineLevel=\"" || within " collapsed=\"" || within " ht='" || within " hidden='" || within " outlineLevel='" || within " collapsed='") then
            (if (indexOfFrom sheetXml "<!" openAt |> fun at -> at >= 0 && at < closeAt) then None else Some [])
        else scanSheetRowsWith true ctx sheetXml |> Option.map fst

    /// Where the rows of a sheet part are, when they are in the shape the
    /// scanner reads: (the ">" of the sheetData start tag, the "<" of its end
    /// tag). None for an empty or prefixed sheetData, or one with comments /
    /// CDATA among the rows.
    let private sheetDataBounds (sheetXml: string) : (int * int) option =
        let openAt = indexOfFrom sheetXml "<sheetData" 0
        if openAt < 0 then None
        else
            let openEnd = sheetXml.IndexOf('>', openAt)
            if openEnd < 0 || codeAt sheetXml (openEnd - 1) = 47 then None
            else
                let closeAt = indexOfFrom sheetXml "</sheetData>" openEnd
                let commentAt = if closeAt < 0 then -1 else indexOfFrom sheetXml "<!" openEnd
                if closeAt < 0 || (commentAt >= 0 && commentAt < closeAt) then None else Some (openEnd, closeAt)

    /// The part without its rows, for the XML parser (None: see `sheetDataBounds`).
    let internal sheetWithoutRows (sheetXml: string) : string option =
        match sheetDataBounds sheetXml with
        | None -> None
        | Some (_, closeAt) ->
            let openAt = indexOfFrom sheetXml "<sheetData" 0
            Some (sheetXml.Substring(0, openAt) + "<sheetData/>" + sheetXml.Substring(closeAt + 12))

    /// The cells of a sheet part that hold a formula, in the order of the
    /// part, exactly as `scanSheetRows` makes them - found by their `<f>`
    /// without reading the other cells. `ctx` is a context of its own (the
    /// masters of shared formulas are collected in it). None when the part
    /// is not in the shape this reads: the caller then reads every row.
    let internal scanFormulaCells (ctx: SheetParseContext) (sheetXml: string) : Cell list option =
        match sheetDataBounds sheetXml with
        | None -> (if indexOfFrom sheetXml "<sheetData" 0 >= 0 && indexOfFrom sheetXml "<f" 0 < 0 then Some [] else None)
        | Some (openEnd, closeAt) ->
            let scan =
                { Xml = sheetXml
                  Ctx = ctx
                  ColumnNames = System.Collections.Generic.Dictionary<int, string>()
                  RowNames = [||]
                  RowValues = [||]
                  CellNames = [| "r"; "s"; "t"; "cm" |]
                  CellValues = Array.create 4 null
                  FormulaNames = [| "t"; "ref"; "si" |]
                  FormulaValues = Array.create 3 null
                  NextInline = indexOfFrom sheetXml "<is" openEnd
                  Valid = true
                  PrevCol = 0 }
            let cells = System.Collections.Generic.List<Cell>()
            let mutable pos = openEnd + 1
            while scan.Valid && pos < closeAt do
                let at = indexOfFrom sheetXml "<f" pos
                if at < 0 || at >= closeAt then pos <- closeAt
                else
                    let after = codeAt sheetXml (at + 2)
                    if not (isSpaceCode after || after = 62 || after = 47) then pos <- at + 2
                    else
                        // The cell the formula is in: the nearest "<c" before it.
                        let mutable cellAt = at - 1
                        let isCellStart (index: int) =
                            codeAt sheetXml index = 60 && codeAt sheetXml (index + 1) = 99
                            && (let next = codeAt sheetXml (index + 2) in isSpaceCode next || next = 62)
                        while cellAt > openEnd && not (isCellStart cellAt) do cellAt <- cellAt - 1
                        let cellTagEnd = if cellAt > openEnd then sheetXml.IndexOf('>', cellAt) else -1
                        let cellEnd = if cellTagEnd < 0 then -1 else indexOfFrom sheetXml "</c>" cellTagEnd
                        if cellTagEnd < 0 || cellTagEnd > at || cellEnd < at || cellEnd > closeAt then scan.Valid <- false
                        else
                            scanAttributes sheetXml (cellAt + 2) cellTagEnd scan.CellNames scan.CellValues
                            let reference = scan.CellValues.[0]
                            let mutable row = 0
                            let mutable letters = 0
                            if not (isNull reference) then
                                for i in 0 .. reference.Length - 1 do
                                    let code = codeAt reference i
                                    if code >= 48 && code <= 57 then row <- row * 10 + (code - 48)
                                    elif (code >= 65 && code <= 90) || (code >= 97 && code <= 122) then letters <- letters + 1
                            // (A cell that does not say where it is takes its place from the cells before it.)
                            if row <= 0 || letters = 0 then scan.Valid <- false
                            else
                                if scan.NextInline >= 0 && scan.NextInline < cellAt then scan.NextInline <- indexOfFrom sheetXml "<is" cellAt
                                let cell =
                                    if scan.NextInline >= 0 && scan.NextInline < cellEnd then
                                        cellFromNode ctx row 0 (XmlDsl.parse (sheetXml.Substring(cellAt, cellEnd + 4 - cellAt))) |> fst
                                    else scanCell scan row cellAt cellTagEnd cellEnd false
                                if cell.Formula.IsSome || (match cell.Value with Formula _ -> true | _ -> false) then cells.Add cell
                                pos <- cellEnd + 4
            if scan.Valid then Some (List.ofSeq cells) else None

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
                OutlineLevel = Xml.attr "outlineLevel" colNode |> Option.bind tryParseInt |> Option.defaultValue 0
                Collapsed = Xml.attr "collapsed" colNode |> Option.map isTrueText |> Option.defaultValue false
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
                    // A link out of the workbook is a relationship; `location`
                    // is then the place inside what it points at.
                    match rid |> Option.bind resolveRel, location with
                    | Some url, Some loc when loc <> "" -> url + "#" + loc, false
                    | Some url, _ -> url, false
                    | None, Some loc -> loc, true
                    | None, None -> rid |> Option.defaultValue "", false
                Some { Range = range; Target = target; IsInternal = isInternal; Tooltip = Xml.attr "tooltip" node; Display = Xml.attr "display" node })

    /// `<sheetView showZeros="false"/>`. Excel's default is true; opt-out
    /// hides numeric-zero cells while keeping other content intact.
    let private showZeros (sheet: XmlNode) : bool =
        sElem "sheetViews" sheet
        |> Option.bind (sElem "sheetView")
        |> Option.bind (Xml.attr "showZeros")
        |> Option.map (fun v -> not (v = "0" || v = "false"))
        |> Option.defaultValue true

    /// Headers, formula view and zoom of the first `<sheetView>`.
    let private sheetViewOptions (sheet: XmlNode) : SheetViewOptions =
        match sElem "sheetViews" sheet |> Option.bind (sElem "sheetView") with
        | None -> defaultSheetView
        | Some view ->
            let flag name fallback =
                match Xml.attr name view with
                | Some v -> not (v = "0" || v = "false")
                | None -> fallback
            { ShowRowColHeaders = flag "showRowColHeaders" true
              ShowFormulas = flag "showFormulas" false
              ZoomScale =
                (match Xml.attr "zoomScale" view |> Option.bind tryParseInt with
                 | Some percent when percent >= 10 && percent <= 400 -> percent
                 | _ -> 100)
              TabColor = None; SummaryBelow = true; SummaryRight = true; ShowOutlineSymbols = flag "showOutlineSymbols" true }

    /// `<sheetPr><tabColor rgb | theme tint | indexed/></sheetPr>`.
    let private tabColorOf (sheet: XmlNode) : string option =
        sElem "sheetPr" sheet
        |> Option.bind (sElem "tabColor")
        |> Option.bind (fun node -> tabColorOfAttributes (Xml.attr "rgb" node) (Xml.attr "theme" node) (Xml.attr "tint" node) (Xml.attr "indexed" node))

    /// `<sheetPr><outlinePr summaryBelow summaryRight/></sheetPr>` (both true when absent).
    let private outlineSummaryOf (sheet: XmlNode) : bool * bool =
        match sElem "sheetPr" sheet |> Option.bind (sElem "outlinePr") with
        | Some node ->
            let flag (name: string) = match Xml.attr name node with Some v -> not (v = "0" || v = "false") | None -> true
            flag "summaryBelow", flag "summaryRight"
        | None -> true, true

    /// `<sheetPr><outlinePr showOutlineSymbols/>` (true when absent): with the
    /// sheet view's flag of that name, whether the outline's symbols are shown.
    let private outlineSymbolsOf (sheet: XmlNode) : bool =
        match sElem "sheetPr" sheet |> Option.bind (sElem "outlinePr") |> Option.bind (Xml.attr "showOutlineSymbols") with
        | Some v -> not (v = "0" || v = "false")
        | None -> true

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
                Kept =
                    match n with
                    | Some e -> keptPageSetupAttributes |> List.choose (fun name -> Xml.attr name e |> Option.map (fun value -> name, value))
                    | None -> []
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
        let mutable fExtra : Map<string, string> =
            if ctx.DynamicCm.IsEmpty then Map.empty
            else match nativeAttr reader "cm" with Some cm -> Map.add "cm" cm Map.empty | None -> Map.empty
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
                        if fType = Some "dataTable" then
                            for name in [ "dt2D"; "dtr"; "r1"; "r2"; "del1"; "del2"; "ca" ] do
                                match nativeAttr reader name with
                                | Some text -> fExtra <- Map.add name text fExtra
                                | None -> ()
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
                fText fType fRef fSi (fun name -> Map.tryFind name fExtra) value inlineText inlineRuns
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
        let outlineLevel = nativeIntAttr reader "outlineLevel" |> Option.defaultValue 0
        let collapsed = nativeBoolAttr reader "collapsed"
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
            OutlineLevel = outlineLevel
            Collapsed = collapsed
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
                OutlineLevel = nativeIntAttr reader "outlineLevel" |> Option.defaultValue 0
                Collapsed = nativeBoolAttr reader "collapsed"
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
        let mutable view = defaultSheetView
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
                    view <-
                        { ShowRowColHeaders = nativeBoolAttrDefault true reader "showRowColHeaders"
                          ShowFormulas = nativeBoolAttrDefault false reader "showFormulas"
                          ZoomScale =
                            (match nativeFloatAttr reader "zoomScale" |> Option.map int with
                             | Some percent when percent >= 10 && percent <= 400 -> percent
                             | _ -> 100)
                          TabColor = None; SummaryBelow = true; SummaryRight = true
                          ShowOutlineSymbols = nativeBoolAttrDefault true reader "showOutlineSymbols" }
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
        showZerosValue, showGridLinesValue, pane, view

    /// What `<sheetPr>` holds: the fit-to-page flag, the tab colour and where
    /// the summary rows / columns of the outline are (below, right).
    let private fitToPageNative (reader: System.Xml.XmlReader) : bool * string option * (bool * bool) * bool =
        let mutable fit = false
        let mutable tabColor = None
        let mutable summary = true, true
        let mutable symbols = true
        if not reader.IsEmptyElement then
            let depth = reader.Depth
            let mutable doneReading = false
            while not doneReading && reader.Read() do
                match reader.NodeType with
                | System.Xml.XmlNodeType.Element when reader.LocalName = "pageSetUpPr" ->
                    fit <- nativeBoolAttr reader "fitToPage"
                | System.Xml.XmlNodeType.Element when reader.LocalName = "tabColor" ->
                    let attribute (name: string) = match reader.GetAttribute name with null -> None | value -> Some value
                    tabColor <- tabColorOfAttributes (attribute "rgb") (attribute "theme") (attribute "tint") (attribute "indexed")
                | System.Xml.XmlNodeType.Element when reader.LocalName = "outlinePr" ->
                    summary <- nativeBoolAttrDefault true reader "summaryBelow", nativeBoolAttrDefault true reader "summaryRight"
                    symbols <- nativeBoolAttrDefault true reader "showOutlineSymbols"
                | System.Xml.XmlNodeType.EndElement when reader.Depth = depth ->
                    doneReading <- true
                | _ -> ()
        fit, tabColor, summary, symbols

    let private defaultPageSetupNative fit =
        {
            Kept = []
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
            Kept = keptPageSetupAttributes |> List.choose (fun name -> nativeAttr reader name |> Option.map (fun value -> name, value))
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
                match rid |> Option.bind resolveRel, location with
                | Some url, Some loc when loc <> "" -> url + "#" + loc, false
                | Some url, _ -> url, false
                | None, Some loc -> loc, true
                | None, None -> rid |> Option.defaultValue "", false
            { Range = range; Target = target; IsInternal = isInternal; Tooltip = nativeAttr reader "tooltip"; Display = nativeAttr reader "display" })

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
            let mutable summaryValue = true, true
            let mutable symbolsValue = true
            let mutable frozenPaneValue = None
            let mutable showZerosValue = true
            let mutable viewValue = defaultSheetView
            let mutable showGridLinesValue = true
            let mutable fitToPage = false
            let mutable tabColorValue = None
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
                        let fit, tabColor, summary, symbols = fitToPageNative reader
                        symbolsValue <- symbols
                        fitToPage <- fit
                        tabColorValue <- tabColor
                        summaryValue <- summary
                    | "sheetViews" ->
                        let showZeros, showGridLines, frozenPane, view = sheetViewsNative reader
                        viewValue <- view
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
                RowSource = SheetRows.Of (resizeListToList rows)
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
                View = { viewValue with TabColor = tabColorValue; SummaryBelow = fst summaryValue; SummaryRight = snd summaryValue; ShowOutlineSymbols = viewValue.ShowOutlineSymbols && symbolsValue }
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
                Extras = emptySheetExtras
            }

    let worksheetNativeFastWithStyles (styles: CellStyle array option) (name: string) (sheetXml: string) (comments: CellComment list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet option =
        worksheetNativeFastCore (newSheetContext styles false) name SheetVisible sheetXml comments [] tables sheetRels printNames
#endif

    /// The rows of a parsed sheet part.
    let internal rowsFromNode (ctx: SheetParseContext) (sheet: XmlNode) : Row list =
        let mutable prevRow = 0
        sElem "sheetData" sheet
        |> Option.map (childElems "row")
        |> Option.defaultValue []
        |> List.map (fun rowNode ->
            let row = rowFromNode ctx prevRow rowNode
            prevRow <- row.Index
            row)

    let internal worksheetFromNodeCore (ctx: SheetParseContext) (name: string) (state: SheetState) (sheet: XmlNode) (comments: CellComment list) (condFormats: ConditionalFormat list) (dataVals: DataValidation list) (charts: Chart list) (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet =
        let rows = rowsFromNode ctx sheet

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
            RowSource = SheetRows.Of rows
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
            View =
                (let below, right = outlineSummaryOf sheet
                 let view = sheetViewOptions sheet
                 { view with TabColor = tabColorOf sheet; SummaryBelow = below; SummaryRight = right; ShowOutlineSymbols = view.ShowOutlineSymbols && outlineSymbolsOf sheet })
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
            Extras = emptySheetExtras
        }

    /// A sheet from the text of its part: the rows are scanned from the text
    /// (see `scanSheetRows`), the rest goes through the XML parser. `rest`
    /// reads what it needs (rules, validations) from the parsed part without
    /// its rows. None when the rows are not in the shape the scanner reads.
    let internal worksheetScannedCore (ctx: SheetParseContext) (name: string) (state: SheetState) (sheetXml: string) (comments: CellComment list)
                                      (rest: XmlNode -> ConditionalFormat list * DataValidation list)
                                      (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet option =
        // Rows first: a shared formula's master must be seen before its followers.
        match scanSheetRows ctx sheetXml with
        | None -> None
        | Some (rows, withoutRows) ->
            let sheet = XmlDsl.parse withoutRows
            let condFormats, dataVals = rest sheet
            Some { worksheetFromNodeCore ctx name state sheet comments condFormats dataVals [] pivotTables tables sheetRels printNames with RowSource = SheetRows.Of rows }

    /// As `worksheetScannedCore`, with the rows left in the text of the part
    /// until something asks for them (`Worksheet.Rows`); the cells with a
    /// formula can be had without the others (`Worksheet.FormulaCells`).
    /// `newContext` makes a context for one reading of the rows. None when the
    /// rows are not in the shape the scanner reads, or hold what the sheet
    /// has to know of at once (a what-if data table).
    let internal worksheetDeferredCore (newContext: unit -> SheetParseContext) (name: string) (state: SheetState) (sheetXml: string) (comments: CellComment list)
                                       (rest: XmlNode -> ConditionalFormat list * DataValidation list)
                                       (pivotTables: PivotTable list) (tables: TableDefinition list) (sheetRels: Relationship list) (printNames: WorksheetPrintNames) : Worksheet option =
        if indexOfFrom sheetXml "\"dataTable\"" 0 >= 0 || indexOfFrom sheetXml "'dataTable'" 0 >= 0 then None
        else
            match sheetWithoutRows sheetXml with
            | None -> None
            | Some withoutRows ->
                let sheet = XmlDsl.parse withoutRows
                let condFormats, dataVals = rest sheet
                let readRows () =
                    match scanSheetRows (newContext ()) sheetXml with
                    | Some (rows, _) -> rows
                    // (Rows the scanner gives up on: the XML parser reads the part.)
                    | None -> rowsFromNode (newContext ()) (XmlDsl.parse sheetXml)
                let readFormulas () = scanFormulaCells (newContext ()) sheetXml
                let readHeaders () = scanRowHeaders (newContext ()) sheetXml
                Some { worksheetFromNodeCore (newContext ()) name state sheet comments condFormats dataVals [] pivotTables tables sheetRels printNames
                         with RowSource = SheetRows.Deferred (readRows, Some readFormulas, Some readHeaders) }

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
                    let totals =
                        Ns.descendantsNs "tableColumn" doc
                        |> List.filter (fun node -> (Xml.attr "name" node).IsSome)
                        |> List.map (fun node ->
                            match Xml.attr "totalsRowLabel" node, Xml.attr "totalsRowFunction" node with
                            | Some label, _ -> "label", label
                            | None, Some "custom" ->
                                (match sElem "totalsRowFormula" node |> Option.bind Xml.text with
                                 | Some formula -> "custom", formula
                                 | None -> "", "")
                            | None, Some "none" | None, None -> "", ""
                            | None, Some name -> "function", name)
                    Some {
                        Name = tableName
                        Columns = columns
                        StartRow = r1
                        StartCol = c1
                        EndRow = r2
                        EndCol = c2
                        HeaderRowCount = headerCount
                        TotalsRowCount = totalsRowCount
                        Totals = totals
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
        // (<b/> is bold, <b val="0"/> is not; <u val="none"/> is no underline.)
        let on (name: string) = match sElem name fontNode with Some node -> (match Xml.attr "val" node with Some ("0" | "false") -> false | _ -> true) | None -> false
        let bold = on "b"
        let italic = on "i"
        let underline = on "u" && (sElem "u" fontNode |> Option.bind (Xml.attr "val")) <> Some "none"
        let underlineDouble =
            sElem "u" fontNode
            |> Option.bind (Xml.attr "val")
            |> Option.map (fun v -> v = "double" || v = "doubleAccounting")
            |> Option.defaultValue false
        let strike = on "strike"
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

    /// `<calcPr iterate="1" iterateCount iterateDelta>`: iterative calculation
    /// of circular references - (most iterations, largest change that ends
    /// them), with Excel's defaults 100 and 0.001 - or None.
    let calcIteration (wbNode: XmlNode) : (int * float) option =
        match sElem "calcPr" wbNode with
        | Some calcPr when (Xml.attr "iterate" calcPr |> Option.map isTrueText |> Option.defaultValue false) ->
            let count =
                match Xml.attr "iterateCount" calcPr |> Option.map System.Int32.TryParse with
                | Some (true, n) when n > 0 -> n
                | _ -> 100
            let delta =
                match Xml.attr "iterateDelta" calcPr |> Option.map (fun text -> System.Double.TryParse(text, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture)) with
                | Some (true, d) when d >= 0.0 -> d
                | _ -> 0.001
            Some (count, delta)
        | _ -> None

    /// `<workbookPr date1904="1"/>`: serial 0 is 1904-01-01, not 1899-12-30.
    let date1904 (wbNode: XmlNode) : bool =
        sElem "workbookPr" wbNode
        |> Option.bind (Xml.attr "date1904")
        |> Option.map isTrueText
        |> Option.defaultValue false

    /// `<workbookView activeTab>`: the sheet on display.
    let activeTab (wbNode: XmlNode) : int =
        sElem "bookViews" wbNode
        |> Option.bind (sElem "workbookView")
        |> Option.bind (Xml.attr "activeTab")
        |> Option.bind tryParseInt
        |> Option.defaultValue 0

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
            let runs =
                match Ns.childNs "text" commentNode with
                | Some textNode -> Xlsx.richRunsOf textNode
                | None -> []
            {
                CellRef = cellRef
                Author = author
                Text = text
                Runs = runs
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
        let iconSet, iconSetCfvo, iconSetReverse =
            match Ns.childNs "iconSet" ruleNode with
            | Some ic ->
                // Excel defaults the family to 3TrafficLights1 when the
                // attribute is omitted.
                Some (Xml.attr "iconSet" ic |> Option.defaultValue "3TrafficLights1"),
                cfvoPairs ic,
                (Xml.attr "reverse" ic |> Option.map isTrueText |> Option.defaultValue false)
            | None -> None, [], false
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
            IconSetReverse = iconSetReverse
            TimePeriod = Xml.attr "timePeriod" ruleNode
            Rank = rank
            Percent = percent
            Bottom = bottom
            AboveAverage = aboveAverage
            EqualAverage = equalAverage
            StdDev = stdDev
            Priority = priority
            StopIfTrue = stopIfTrue
            Ext =
                // `showValue="0"` of the 2007 icon set.
                match Ns.childNs "iconSet" ruleNode |> Option.bind (Xml.attr "showValue") with
                | Some text when not (isTrueText text) -> Some { emptyCfRuleExt with ShowValue = false }
                | _ -> None
        }

    /// One `<dxf>` (of the styles part, or the `<x14:dxf>` of an x14 rule).
    let parseDxf (themeColors: string array) (dxf: XmlNode) : Dxf =
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
            Underline =
                Xlsx.sElem "font" dxf
                |> Option.bind (Xlsx.sElem "u")
                |> Option.map (fun node -> Xml.attr "val" node <> Some "none")
            Strike = boolChild "strike"
        }

    /// The rule types of the x14 list this reader makes rules of.
    let x14RuleTypes =
        [ "dataBar"; "iconSet"; "colorScale"; "expression"; "cellIs"; "containsText"; "notContainsText"; "beginsWith"; "endsWith"
          "top10"; "aboveAverage"; "duplicateValues"; "uniqueValues"; "containsBlanks"; "notContainsBlanks"; "containsErrors"
          "notContainsErrors"; "timePeriod" ]

    let private directChildren (local: string) (node: XmlNode) =
        Xml.children node
        |> List.filter (fun child -> match child with XmlDsl.Element (name, _, _) -> Ns.stripNs name = local | _ -> false)

    /// The x14 rules of a sheet (`<x14:conditionalFormatting>` blocks of its
    /// extension list): (sqref, rule node) in document order.
    let private x14RuleNodes (sheet: XmlNode) : (string * XmlNode) list =
        match Ns.childNs "extLst" sheet with
        | None -> []
        | Some extLst ->
            Ns.descendantsNs "conditionalFormatting" extLst
            |> List.collect (fun block ->
                let sqref = directChildren "sqref" block |> List.tryHead |> Option.bind Xml.text |> Option.map (fun text -> text.Trim()) |> Option.defaultValue ""
                directChildren "cfRule" block |> List.map (fun rule -> sqref, rule))

    /// Whether every x14 rule of the sheet is one this reader makes a rule
    /// of. When not, the extension is kept as it is and none is read.
    let x14Understood (sheet: XmlNode) : bool =
        x14RuleNodes sheet
        |> List.forall (fun (sqref, rule) ->
            let kind = Xml.attr "type" rule |> Option.defaultValue ""
            Xml.attr "pivot" rule |> Option.forall (fun text -> not (isTrueText text))
            && List.contains kind x14RuleTypes
            && (sqref <> "" || kind = "dataBar"))

    let private x14Cfvo (parent: XmlNode) : (string * string option) list =
        directChildren "cfvo" parent
        |> List.map (fun cfvo ->
            Xml.attr "type" cfvo |> Option.defaultValue "min",
            (Ns.childNs "f" cfvo |> Option.bind Xml.text |> Option.filter (fun text -> text <> "")))

    let private x14DataBar (themeColors: string array) (bar: XmlNode) : CfDataBarExt * string option =
        let flag (name: string) (fallback: bool) = Xml.attr name bar |> Option.map isTrueText |> Option.defaultValue fallback
        let number (name: string) (fallback: int) = Xml.attr name bar |> Option.bind tryParseInt |> Option.defaultValue fallback
        let colorOf (name: string) = Ns.childNs name bar |> Option.bind (Xlsx.resolveColorElement themeColors)
        { MinLength = number "minLength" 10
          MaxLength = number "maxLength" 90
          ShowValue = flag "showValue" true
          Border = flag "border" false
          Gradient = flag "gradient" true
          Direction = Xml.attr "direction" bar |> Option.defaultValue "context"
          NegativeBarColorSameAsPositive = flag "negativeBarColorSameAsPositive" false
          NegativeBarBorderColorSameAsPositive = flag "negativeBarBorderColorSameAsPositive" true
          AxisPosition = Xml.attr "axisPosition" bar |> Option.defaultValue "automatic"
          Cfvo = x14Cfvo bar
          BorderColor = colorOf "borderColor"
          NegativeFillColor = colorOf "negativeFillColor"
          NegativeBorderColor = colorOf "negativeBorderColor"
          AxisColor = colorOf "axisColor" },
        colorOf "fillColor"

    /// A rule that is in the x14 list only.
    let private x14Rule (themeColors: string array) (ruleNode: XmlNode) : ConditionalFormatRule =
        let basic = parseRule themeColors ruleNode
        let formulas = directChildren "f" ruleNode |> List.choose Xml.text
        let ext = { emptyCfRuleExt with InExtList = true; Dxf = Ns.childNs "dxf" ruleNode |> Option.map (parseDxf themeColors) }
        match Ns.childNs "dataBar" ruleNode, Ns.childNs "iconSet" ruleNode, Ns.childNs "colorScale" ruleNode with
        | Some bar, _, _ ->
            let barExt, fill = x14DataBar themeColors bar
            let legacy (kind: string, value: string option) =
                match kind with
                | "autoMin" -> "min", None
                | "autoMax" -> "max", None
                | _ -> kind, value
            { basic with
                Formulas = []
                DataBarColor = fill
                DataBarCfvo = barExt.Cfvo |> List.map legacy
                Ext = Some { emptyCfRuleExt with DataBar = Some barExt } }
        | None, Some icons, _ ->
            let custom =
                directChildren "cfIcon" icons
                |> List.map (fun icon -> Xml.attr "iconSet" icon |> Option.defaultValue "NoIcons", (Xml.attr "iconId" icon |> Option.bind tryParseInt |> Option.defaultValue 0))
            { basic with
                Formulas = []
                IconSetCfvo = x14Cfvo icons
                Ext = Some { ext with
                                IconCustom = custom
                                ShowValue = Xml.attr "showValue" icons |> Option.map isTrueText |> Option.defaultValue true } }
        | None, None, Some scale ->
            let stops = x14Cfvo scale
            let colors = directChildren "color" scale |> List.map (fun node -> Xlsx.resolveColorElement themeColors node |> Option.defaultValue "FFFFFF")
            { basic with
                Formulas = []
                ColorScale =
                    List.zip (stops |> List.truncate colors.Length) (colors |> List.truncate stops.Length)
                    |> List.map (fun ((kind, value), color) -> { CfvoType = kind; CfvoValue = value; Color = color })
                Ext = Some ext }
        | None, None, None -> { basic with Formulas = formulas; Ext = Some ext }

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
        let x14 = if x14Understood sheet then x14RuleNodes sheet else []
        // The x14 data bar a 2007 data bar names by id (`<x14:id>` in the rule's own extension).
        let barsById = System.Collections.Generic.Dictionary<string, XmlNode>()
        for _, rule in x14 do
            match Xml.attr "type" rule, Xml.attr "id" rule, Ns.childNs "dataBar" rule with
            | Some "dataBar", Some id, Some bar -> barsById.[id] <- bar
            | _ -> ()
        let paired = System.Collections.Generic.HashSet<string>()
        let main =
            Ns.descendantsNs "conditionalFormatting" sheet
            |> List.choose (fun cfNode ->
                Xml.attr "sqref" cfNode
                |> Option.map (fun range ->
                    let rules =
                        Ns.descendantsNs "cfRule" cfNode
                        |> List.map (fun ruleNode ->
                            let rule = parseRule themeColors ruleNode
                            let id = Ns.descendantsNs "id" ruleNode |> List.tryHead |> Option.bind Xml.text |> Option.map (fun text -> text.Trim())
                            match id with
                            | Some id when rule.Type = "dataBar" && barsById.ContainsKey id ->
                                paired.Add id |> ignore
                                let barExt, fill = x14DataBar themeColors barsById.[id]
                                { rule with
                                    DataBarColor = (match fill with Some _ -> fill | None -> rule.DataBarColor)
                                    Ext = Some { emptyCfRuleExt with DataBar = Some barExt } }
                            | _ -> rule)
                    { Range = range; Rules = rules }))
        let extOnly =
            x14
            |> List.choose (fun (sqref, ruleNode) ->
                match Xml.attr "id" ruleNode with
                | Some id when paired.Contains id -> None
                | _ when sqref = "" -> None
                | _ -> Some { Range = sqref; Rules = [ x14Rule themeColors ruleNode ] })
        main @ extOnly

    /// Parse the `<dxfs>` table out of styles.xml. Each `<dxf>` carries a
    /// partial font + fill; conditional-format rules reference them by
    /// index. dxf fills conventionally use `<bgColor>` (not fgColor),
    /// so check both.
    let parseDxfs (themeColors: string array) (stylesDoc: XmlNode) : Dxf array =
        Xlsx.sElem "dxfs" stylesDoc
        |> Option.map (Xlsx.sElems "dxf")
        |> Option.defaultValue []
        |> List.map (parseDxf themeColors)
        |> List.toArray

// Data Validation parsing
module DataValidations =
    let parseDataValidation (dvNode: XmlNode) : DataValidation option =
        // The x14 form has its range in `<xm:sqref>` and its formulas in `<xm:f>`.
        let childText (name: string) = Ns.childNs name dvNode |> Option.bind Xml.text |> Option.map (fun text -> text.Trim()) |> Option.filter (fun text -> text <> "")
        let formulaOf (name: string) =
            Ns.childNs name dvNode
            |> Option.bind (fun node ->
                match Ns.childNs "f" node with
                | Some inner -> Xml.text inner
                | None -> Xml.text node)
        Xml.attr "sqref" dvNode
        |> Option.orElseWith (fun () -> childText "sqref")
        |> Option.map (fun range ->
            let valType = Xml.attr "type" dvNode |> Option.defaultValue "none"
            let operator = Xml.attr "operator" dvNode
            let flag name fallback =
                Xml.attr name dvNode |> Option.map isTrueText |> Option.defaultValue fallback
            // `showDropDown="1"` means the in-cell arrow is SUPPRESSED (the
            // attribute is named backwards in the schema).
            let showDropDown = not (flag "showDropDown" false)
            let formula1 = formulaOf "formula1"
            let formula2 = formulaOf "formula2"
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
        let main =
            Ns.childNs "dataValidations" sheet
            |> Option.map (Ns.descendantsNs "dataValidation")
            |> Option.defaultValue []
        // The validations whose formulas read another sheet (`<x14:dataValidations>` of the extension list).
        let extended =
            Ns.childNs "extLst" sheet
            |> Option.map (Ns.descendantsNs "dataValidations")
            |> Option.defaultValue []
            |> List.collect (Ns.descendantsNs "dataValidation")
        main @ extended |> List.choose parseDataValidation

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
                    Comment = Xml.attr "comment" nameNode |> Option.filter (fun text -> text <> "")
                    Hidden = (match Xml.attr "hidden" nameNode with Some ("1" | "true") -> true | _ -> false)
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

    /// The areas of a reference to several (`A,B`): split at the commas that
    /// are not inside a quoted sheet name.
    let private splitAreas (text: string) : string list =
        let parts = ResizeArray<string>()
        let mutable start = 0
        let mutable quoted = false
        for i in 0 .. text.Length - 1 do
            if text.[i] = '\'' then quoted <- not quoted
            elif text.[i] = ',' && not quoted then
                parts.Add (text.Substring(start, i - start))
                start <- i + 1
        parts.Add (text.Substring start)
        parts |> Seq.map (fun part -> part.Trim()) |> Seq.filter (fun part -> part <> "") |> Seq.toList

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
                // (Several areas: `Sheet1!$A$1:$C$10,Sheet1!$E$1:$G$10`, kept as `$A$1:$C$10,$E$1:$G$10`.)
                update idx (fun w -> { w with PrintArea = Some (splitAreas nr.RefersTo |> List.map stripSheetQualifier |> String.concat ",") })
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

    // (Rows that are not read yet are styled when they are.)
    let source = sheet.RowSource
    { sheet with RowSource = (if source.IsLoaded then SheetRows.Of (source.Value |> List.map applyToRow) else SheetRows.Deferred ((fun () -> source.Value |> List.map applyToRow), None, Some (fun () -> if source.IsLoaded then None else Some source.Headers))) }

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
    /// Raw bytes of a part (pictures, and parts kept to be written back).
    Bytes: string -> byte[] option
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
    /// Relationships of the workbook part to kept parts, and the elements of
    /// workbook.xml that refer to them (see WorkbookPreservation).
    KeptRels: (string * string * string * bool) list
    ExternalReferencesXml: string
    PivotCachesXml: string
    ExtLstXml: string
    CalcIteration: (int * float) option
    ActiveTab: int
}

/// Relationship types of the workbook part whose targets are kept as they
/// are: every one but those of the parts a writer makes from the model (the
/// sheets, styles, shared strings, the theme, the calculation chain, the
/// persons of threaded comments, the macros). So pivot and slicer caches,
/// external links, connections, cell metadata and rich values, the Power
/// Pivot data model, custom XML, XML maps ... all go back as they were.
let private isKeptWorkbookRelType (relType: string) =
    not ([ "/worksheet"; "/chartsheet"; "/dialogsheet"; "/macrosheet"; "/xlMacrosheet"; "/xlIntlMacrosheet"; "/styles"; "/sharedStrings"; "/theme"
           "/calcChain"; "/person"; "/vbaProject"; "/officeDocument" ]
         |> List.exists (fun suffix -> relType.EndsWith suffix))

/// Sheets the editor has no model of: a chart sheet, a dialog sheet. They are
/// shown as an empty sheet (a chart sheet with its chart on it) and written
/// back as the parts they were read from.
let private isForeignSheetRelType (relType: string) =
    relType.EndsWith "/chartsheet" || relType.EndsWith "/dialogsheet"

/// Relationship types of a sheet part whose targets are kept as they are:
/// every one but those of what the model holds (the drawing, tables,
/// comments, threaded comments, hyperlinks). A VML drawing is kept when the
/// sheet names it (see `keptSheetElements`): it may hold form controls.
let private isKeptSheetRelType (relType: string) =
    not ([ "/drawing"; "/table"; "/comments"; "/threadedComment"; "/hyperlink"; "/vmlDrawing" ] |> List.exists (fun suffix -> relType.EndsWith suffix))

/// The elements of a worksheet that only name parts the model does not hold,
/// in the order CT_Worksheet has them.
let keptSheetElementNames = [ "customProperties"; "legacyDrawing"; "legacyDrawingHF"; "drawingHF"; "picture"; "oleObjects"; "controls"; "webPublishItems" ]

/// Those elements of a sheet as (name, XML with the namespace declarations
/// it needs). Read from the text of the part, after the rows: both sheet
/// parsers leave them alone.
let private keptSheetElements (sheetXml: string) : (string * string) list =
    let rowsEnd = max (sheetXml.LastIndexOf "</sheetData>") (sheetXml.LastIndexOf "<sheetData/>")
    let from = if rowsEnd < 0 then 0 else rowsEnd
    let found =
        keptSheetElementNames
        |> List.choose (fun name ->
            let openAt = sheetXml.IndexOf ("<" + name, from)
            // (Not a longer name that begins the same: legacyDrawing / legacyDrawingHF.)
            let rec startOf (at: int) =
                if at < 0 then -1
                else
                    let next = at + name.Length + 1
                    if next < sheetXml.Length && (sheetXml.[next] = ' ' || sheetXml.[next] = '>' || sheetXml.[next] = '/' || sheetXml.[next] = '\r' || sheetXml.[next] = '\n' || sheetXml.[next] = '\t') then at
                    else startOf (sheetXml.IndexOf ("<" + name, next))
            let start = startOf openAt
            if start < 0 then None
            else
                let tagEnd = sheetXml.IndexOf ('>', start)
                if tagEnd < 0 then None
                elif sheetXml.[tagEnd - 1] = '/' then Some (name, sheetXml.Substring(start, tagEnd + 1 - start))
                else
                    let close = "</" + name + ">"
                    let closeAt = sheetXml.IndexOf (close, tagEnd)
                    if closeAt < 0 then None else Some (name, sheetXml.Substring(start, closeAt + close.Length - start)))
    if found.IsEmpty then []
    else
        try
            let rootAt = sheetXml.IndexOf "<worksheet"
            let rootEnd = sheetXml.IndexOf ('>', rootAt)
            match XmlDsl.parse (sheetXml.Substring(rootAt, rootEnd - rootAt).TrimEnd('/') + "/>") with
            | XmlDsl.Element (_, rootAttributes, _) -> found |> List.map (fun (name, xml) -> name, DocxReader.selfContainedXml rootAttributes xml)
            | _ -> []
        with _ -> []

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
        let keptRels =
            rels
            |> List.filter (fun r -> isKeptWorkbookRelType r.Type)
            |> List.map (fun r -> r.Id, r.Type, (if isInternal r then "/" + relTargetPath workbookPath r else r.Target), not (isInternal r))
        let keptIds = keptRels |> List.map (fun (id, _, _, _) -> id) |> Set.ofList
        let rootAttributes = match wb with | XmlDsl.Element (_, attrs, _) -> attrs | _ -> Map.empty
        let fragment (node: XmlNode) = DocxReader.selfContainedXml rootAttributes (DocxReader.Docx.xmlToString node)
        // An element is kept when every relationship it names is a kept one.
        let refersOnlyToKept (xml: string) =
            DocxReader.relationshipIdsInXml xml |> List.forall (fun id -> Set.contains id keptIds)
        let keptChild (name: string) =
            match Ns.childNs name wb with
            | Some node ->
                let xml = fragment node
                if refersOnlyToKept xml then xml else ""
            | None -> ""
        let extLst =
            match Ns.childNs "extLst" wb with
            | Some node ->
                let kept = Xml.children node |> List.map fragment |> List.filter refersOnlyToKept
                if kept.IsEmpty then "" else "<extLst>" + String.concat "" kept + "</extLst>"
            | None -> ""
        {
            KeptRels = keptRels
            ExternalReferencesXml = keptChild "externalReferences"
            PivotCachesXml = keptChild "pivotCaches"
            ExtLstXml = extLst
            CalcIteration = Xlsx.calcIteration wb
            ActiveTab = Xlsx.activeTab wb
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

// ---- drawings: pictures and other objects anchored to cells -----------------

let private imageContentType (path: string) =
    let lower = path.ToLowerInvariant()
    let extension = match lower.LastIndexOf '.' with | -1 -> "" | index -> lower.Substring(index + 1)
    let contentType =
        match extension with
        | "png" -> "image/png"
        | "jpg" | "jpeg" -> "image/jpeg"
        | "gif" -> "image/gif"
        | "bmp" -> "image/bmp"
        | "svg" -> "image/svg+xml"
        | "webp" -> "image/webp"
        | "emf" -> "image/x-emf"
        | "wmf" -> "image/x-wmf"
        | "tif" | "tiff" -> "image/tiff"
        | _ -> "application/octet-stream"
    extension, contentType

let private anchorOf (node: XmlNode) : SheetAnchor option =
    let number (text: string option) = text |> Option.bind tryParseFloat |> Option.defaultValue 0.0
    let childText name parent = Ns.childNs name parent |> Option.bind Xml.text
    let marker name =
        Ns.childNs name node
        |> Option.map (fun m ->
            (childText "col" m |> Option.bind tryParseInt |> Option.defaultValue 0), number (childText "colOff" m),
            (childText "row" m |> Option.bind tryParseInt |> Option.defaultValue 0), number (childText "rowOff" m))
    let size =
        Ns.childNs "ext" node
        |> Option.orElseWith (fun () -> Ns.descendantsNs "xfrm" node |> List.tryHead |> Option.bind (Ns.childNs "ext"))
    let width = size |> Option.bind (Xml.attr "cx") |> Option.bind tryParseFloat
    let height = size |> Option.bind (Xml.attr "cy") |> Option.bind tryParseFloat
    let kind =
        match Xml.name node |> Option.map Ns.stripNs with
        | Some "twoCellAnchor" -> Some "twoCell"
        | Some "oneCellAnchor" -> Some "oneCell"
        | Some "absoluteAnchor" -> Some "absolute"
        | _ -> None
    kind
    |> Option.map (fun kind ->
        let fromCol, fromColOffset, fromRow, fromRowOffset = marker "from" |> Option.defaultValue (0, 0.0, 0, 0.0)
        let toMarker = marker "to"
        let pos = Ns.childNs "pos" node
        { Kind = kind
          FromCol = fromCol; FromColOffset = fromColOffset; FromRow = fromRow; FromRowOffset = fromRowOffset
          ToCol = toMarker |> Option.map (fun (col, _, _, _) -> col)
          ToColOffset = toMarker |> Option.map (fun (_, offset, _, _) -> offset) |> Option.defaultValue 0.0
          ToRow = toMarker |> Option.map (fun (_, _, row, _) -> row)
          ToRowOffset = toMarker |> Option.map (fun (_, _, _, offset) -> offset) |> Option.defaultValue 0.0
          X = pos |> Option.bind (Xml.attr "x") |> number
          Y = pos |> Option.bind (Xml.attr "y") |> number
          Width = width
          Height = height
          EditAs = Xml.attr "editAs" node
          Order = 0 })

/// The pictures and other objects of a sheet's drawing part. A picture is read
/// into the model (anchor, size, data); anything else (a chart, a shape, a
/// group) is kept as its anchor element, to be written back as it is.
let private readDrawing (source: PartSource) (drawingPath: string) : SheetImage list * SheetObject list * (string * string * string * bool) list =
    match source.Text drawingPath |> Option.map XmlDsl.parse with
    | Some (XmlDsl.Element (_, rootAttrs, anchors) as root) ->
        let rels = relationshipsOf source drawingPath
        let images = ResizeArray<SheetImage>()
        let objects = ResizeArray<SheetObject>()
        let mutable order = 0
        for node in anchors do
            match anchorOf node with
            | Some anchor ->
                let anchor = { anchor with Order = order }
                order <- order + 1
                let picture =
                    Ns.childNs "pic" node
                    |> Option.bind (fun pic ->
                        let blip = Ns.descendantsNs "blip" pic |> List.tryHead
                        let relId = blip |> Option.bind (fun b -> Xml.attr "r:embed" b)
                        let rel = relId |> Option.bind (fun id -> rels |> List.tryFind (fun r -> r.Id = id && isInternal r))
                        let props = Ns.descendantsNs "cNvPr" pic |> List.tryHead
                        rel
                        |> Option.bind (fun r ->
                            let path = relTargetPath drawingPath r
                            source.Bytes path
                            |> Option.map (fun data ->
                                let extension, contentType = imageContentType path
                                { Anchor = anchor
                                  Name = props |> Option.bind (Xml.attr "name")
                                  Description = props |> Option.bind (Xml.attr "descr")
                                  ContentType = contentType
                                  Extension = extension
                                  Data = data })))
                match picture with
                | Some image -> images.Add image
                | None ->
                    let kind =
                        if not (Ns.descendantsNs "chart" node).IsEmpty then "chart"
                        elif (Ns.childNs "sp" node).IsSome || (Ns.childNs "cxnSp" node).IsSome then "shape"
                        else "other"
                    objects.Add
                        { Anchor = anchor
                          Kind = kind
                          Name = Ns.descendantsNs "cNvPr" node |> List.tryHead |> Option.bind (Xml.attr "name")
                          Xml = DocxReader.selfContainedXml rootAttrs (DocxReader.Docx.xmlToString node) }
            | None -> ()
        ignore root
        // Relationships the kept objects refer to, with absolute targets.
        let referenced = objects |> Seq.collect (fun item -> DocxReader.relationshipIdsInXml item.Xml) |> Set.ofSeq
        let keptRels =
            rels
            |> List.filter (fun r -> Set.contains r.Id referenced)
            // A target like "#Sheet1!A1" is a place in the workbook, not a part: kept as written.
            |> List.map (fun r ->
                let isPart = isInternal r && not (r.Target.StartsWith "#")
                r.Id, r.Type, (if isPart then "/" + relTargetPath drawingPath r else r.Target), not (isInternal r))
        List.ofSeq images, List.ofSeq objects, keptRels
    | _ -> [], [], []

// ---- threaded comments ---------------------------------------------------------

let private readPersons (source: PartSource) (path: string) : (string * string * string * string) list =
    match source.Text path |> Option.map XmlDsl.parse with
    | Some doc ->
        Ns.descendantsNs "person" doc
        |> List.choose (fun person ->
            Xml.attr "id" person
            |> Option.map (fun id ->
                id, (Xml.attr "displayName" person |> Option.defaultValue ""), (Xml.attr "userId" person |> Option.defaultValue ""), (Xml.attr "providerId" person |> Option.defaultValue "None")))
    | None -> []

/// Threaded comments of a sheet: each thread is its first comment with the
/// later ones (parentId) as replies, in file order.
let private readThreadedComments (persons: (string * string * string * string) list) (xml: string) : ThreadedComment list =
    let names = persons |> List.map (fun (id, name, _, _) -> id, name) |> Map.ofList
    let all =
        Ns.descendantsNs "threadedComment" (XmlDsl.parse xml)
        |> List.choose (fun node ->
            match Xml.attr "ref" node, Xml.attr "id" node with
            | Some cellRef, Some id ->
                let personId = Xml.attr "personId" node
                Some (Xml.attr "parentId" node,
                      { CellRef = Xlsx.parseCellRef cellRef
                        Id = id
                        Author = personId |> Option.bind (fun person -> Map.tryFind person names) |> Option.filter (fun name -> name <> "")
                        PersonId = personId
                        Date = Xml.attr "dT" node
                        Text = Ns.childNs "text" node |> Option.bind Xml.text |> Option.defaultValue ""
                        Replies = []
                        Done = (match Xml.attr "done" node with Some ("1" | "true") -> true | _ -> false) })
            | _ -> None)
    all
    |> List.filter (fun (parent, _) -> parent.IsNone)
    |> List.map (fun (_, root) ->
        { root with Replies = all |> List.filter (fun (parent, _) -> parent = Some root.Id) |> List.map snd })

// ---- named cell styles ---------------------------------------------------------

let private readNamedStyles (stylesXml: string) : NamedCellStyle list * int array =
    let doc = XmlDsl.parse stylesXml
    let entries (containerName: string) =
        Ns.childNs containerName doc |> Option.map Xml.children |> Option.defaultValue [] |> List.filter (fun node -> (Xml.name node).IsSome) |> List.toArray
    let fonts, fills, borders = entries "fonts", entries "fills", entries "borders"
    let numFmts =
        entries "numFmts"
        |> Array.choose (fun node -> Option.map2 (fun id code -> id, code) (Xml.attr "numFmtId" node |> Option.bind tryParseInt) (Xml.attr "formatCode" node))
        |> Map.ofArray
    let styleXfs = entries "cellStyleXfs"
    // A fragment may use a prefix the styles part declares on its root (x14ac, xr, ...).
    let rootAttributes = match doc with | XmlDsl.Element (_, attrs, _) -> attrs | _ -> Map.empty
    let fragment (node: XmlNode) = DocxReader.selfContainedXml rootAttributes (DocxReader.Docx.xmlToString node)
    let itemXml (items: XmlNode array) (index: int option) =
        index |> Option.bind (fun i -> if i >= 0 && i < items.Length then Some (fragment items.[i]) else None)
    entries "cellStyles"
    |> Array.toList
    |> List.choose (fun cellStyle ->
        match Xml.attr "name" cellStyle, Xml.attr "xfId" cellStyle |> Option.bind tryParseInt with
        | Some name, Some xfId when xfId >= 0 && xfId < styleXfs.Length ->
            let xf = styleXfs.[xfId]
            let attributes (node: XmlNode) (skip: string list) =
                match node with
                // Prefixed attributes (xr:uid) are editor bookkeeping and need their namespace; left out.
                | XmlDsl.Element (_, attrs, _) -> attrs |> Map.toList |> List.filter (fun (key, _) -> not (List.contains key skip) && not (key.Contains ":"))
                | _ -> []
            let intAttr key = Xml.attr key xf |> Option.bind tryParseInt
            let numFmtId = intAttr "numFmtId" |> Option.defaultValue 0
            Some (xfId,
                  { Name = name
                    StyleAttributes = attributes cellStyle [ "name"; "xfId" ]
                    XfAttributes = attributes xf [ "numFmtId"; "fontId"; "fillId"; "borderId" ]
                    XfChildrenXml = Xml.children xf |> List.map fragment |> String.concat ""
                    FontXml = itemXml fonts (intAttr "fontId")
                    FillXml = itemXml fills (intAttr "fillId")
                    BorderXml = itemXml borders (intAttr "borderId")
                    // A custom format is identified by its code; its id is the file's own numbering.
                    NumFmtId = if numFmtId >= 164 then 0 else numFmtId
                    NumFmtCode = if numFmtId >= 164 then Map.tryFind numFmtId numFmts else None })
        | _ -> None)
    // In the order of their xf, so that xf 0 (Normal, which the cell formats are based on) stays first.
    |> List.sortBy fst
    |> List.distinctBy fst
    |> fun named ->
        // Which named style each cell format is based on (its xfId).
        let position = named |> List.mapi (fun index (xfId, _) -> xfId, index) |> Map.ofList
        let cellXfNamedStyles =
            entries "cellXfs"
            |> Array.map (fun xf ->
                Xml.attr "xfId" xf
                |> Option.bind tryParseInt
                |> Option.bind (fun xfId -> Map.tryFind xfId position)
                |> Option.defaultValue -1)
        List.map snd named, cellXfNamedStyles

// ---- parts kept for the next save ----------------------------------------------

/// The `<ext>` elements of a sheet's own `<extLst>` (the one that closes the
/// worksheet element) that hold a slicer list or timeline references, each
/// with the namespace declarations it needs. Read from the text of the part:
/// both sheet parsers leave the extension list alone.
let private isPartExtension (xml: string) = xml.Contains "slicerList" || xml.Contains "timelineRefs"
/// Sparkline groups and the x14 conditional formats (data bars, icon sets of Excel 2010+).
let private isCellExtension (xml: string) = xml.Contains "sparklineGroups" || xml.Contains "conditionalFormattings"

let private keptSheetExtensions (sheetXml: string) : string list =
    let closeAt = sheetXml.LastIndexOf "</extLst>"
    let rootAt = sheetXml.IndexOf "<worksheet"
    if closeAt < 0 || rootAt < 0 then []
    else
        let after = sheetXml.Substring(closeAt + 9).Trim()
        let rootEnd = sheetXml.IndexOf ('>', rootAt)
        if not (after.StartsWith "</") || after.IndexOf '<' <> after.LastIndexOf '<' || rootEnd < 0 then []
        else
            // The matching start tag: the last "<extLst" at nesting depth 0 before it.
            let mutable depth = 0
            let mutable openAt = -1
            let mutable at = closeAt
            while openAt < 0 && at > rootEnd do
                let previousOpen = sheetXml.LastIndexOf ("<extLst", at - 1)
                let previousClose = sheetXml.LastIndexOf ("</extLst>", at - 1)
                if previousOpen < 0 then at <- -1
                elif previousClose > previousOpen then
                    depth <- depth + 1
                    at <- previousClose
                elif depth > 0 then
                    depth <- depth - 1
                    at <- previousOpen
                else openAt <- previousOpen
            let fragment = if openAt < 0 then "" else sheetXml.Substring(openAt, closeAt + 9 - openAt)
            if not (isPartExtension fragment || isCellExtension fragment) then []
            else
                try
                    let rootTag = sheetXml.Substring(rootAt, rootEnd - rootAt).TrimEnd('/')
                    match XmlDsl.parse (rootTag + ">" + fragment + "</worksheet>") with
                    | XmlDsl.Element (_, rootAttributes, [ extLst ]) as parsed ->
                        // The x14 conditional formats are read as rules (and written from
                        // them); they are only kept as they are when one is not understood.
                        let rulesRead = ConditionalFormats.x14Understood parsed
                        Xml.children extLst
                        |> List.map (fun ext -> DocxReader.selfContainedXml rootAttributes (DocxReader.Docx.xmlToString ext))
                        |> List.filter (fun xml ->
                            isPartExtension xml || xml.Contains "sparklineGroups" || (xml.Contains "conditionalFormattings" && not rulesRead))
                    | _ -> []
                with _ -> []

let private pageSetupRelPattern = System.Text.RegularExpressions.Regex("<(?:[A-Za-z0-9_]+:)?pageSetup\\b[^>]*?\\sr:id=\"([^\"]*)\"")

/// The relationship id on a sheet's `<pageSetup>` (its printer settings).
let private pageSetupRelId (sheetXml: string) : string option =
    // The element is near the end of the part, after the rows.
    let at = sheetXml.LastIndexOf "<pageSetup"
    if at < 0 then None
    else
        let close = sheetXml.IndexOf ('>', at)
        if close < 0 then None
        else
            let m = pageSetupRelPattern.Match (sheetXml.Substring(at, close - at + 1))
            if m.Success then Some m.Groups.[1].Value else None

let private worksheetSourcePattern = System.Text.RegularExpressions.Regex("<(?:[A-Za-z0-9_]+:)?worksheetSource\\b[^>]*>")
let private attributeIn (name: string) (tag: string) : string option =
    let key = " " + name + "=\""
    let at = tag.IndexOf key
    if at < 0 then None
    else
        let start = at + key.Length
        let close = tag.IndexOf ('"', start)
        if close < 0 then None else Some (tag.Substring(start, close - start))

let private unescapeAttribute (text: string) : string =
    if not (text.Contains "&") then text
    else text.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"").Replace("&apos;", "'").Replace("&amp;", "&")

let private autoFilterTagPattern = System.Text.RegularExpressions.Regex("<(filter|customFilter|top10|colorFilter|dynamicFilter|filters|customFilters|dateGroupItem|iconFilter)\\b[^>]*>")

/// The sheet's `<autoFilter>` (the one after the rows; a table's is in the
/// table's own part), read off the text of the sheet part.
let private sheetAutoFilterOf (sheetXml: string) : SheetAutoFilter option =
    let dataEnd = max 0 (sheetXml.LastIndexOf "</sheetData>")
    let at = sheetXml.IndexOf ("<autoFilter", dataEnd)
    let tagEnd = if at < 0 then -1 else sheetXml.IndexOf ('>', at)
    if tagEnd < 0 then None
    else
        let tag = sheetXml.Substring(at, tagEnd - at + 1)
        match attributeIn "ref" tag with
        | None -> None
        | Some reference ->
            let body =
                if tag.EndsWith "/>" then ""
                else
                    let close = sheetXml.IndexOf ("</autoFilter>", tagEnd)
                    if close < 0 then "" else sheetXml.Substring(tagEnd + 1, close - tagEnd - 1)
            let flag (name: string) (fallback: bool) (element: string) =
                match attributeIn name element with
                | Some "1" | Some "true" -> true
                | Some _ -> false
                | None -> fallback
            let columns = ResizeArray<AutoFilterColumn>()
            let mutable from = 0
            let mutable more = true
            while more do
                let start = body.IndexOf ("<filterColumn", from)
                let startEnd = if start < 0 then -1 else body.IndexOf ('>', start)
                if startEnd < 0 then more <- false
                else
                    let startTag = body.Substring(start, startEnd - start + 1)
                    let close = if startTag.EndsWith "/>" then -1 else body.IndexOf ("</filterColumn>", startEnd)
                    let inner = if close < 0 then "" else body.Substring(startEnd + 1, close - startEnd - 1)
                    from <- (if close < 0 then startEnd + 1 else close + 15)
                    let tags = [ for m in autoFilterTagPattern.Matches inner -> m.Groups.[1].Value, m.Value ]
                    let first (name: string) = tags |> List.tryFind (fun (found, _) -> found = name) |> Option.map snd
                    let column = { emptyAutoFilterColumn with ColId = (attributeIn "colId" startTag |> Option.bind tryParseInt |> Option.defaultValue 0); Xml = inner }
                    let parsed =
                        if tags |> List.exists (fun (name, _) -> name = "dateGroupItem" || name = "iconFilter") then column
                        else
                            match first "filters", first "customFilters", first "top10", first "colorFilter", first "dynamicFilter" with
                            | Some filters, _, _, _, _ ->
                                { column with
                                    Kind = "filters"
                                    Blank = flag "blank" false filters
                                    Values = tags |> List.choose (fun (name, element) -> if name = "filter" then attributeIn "val" element |> Option.map unescapeAttribute else None) }
                            | None, Some custom, _, _, _ ->
                                { column with
                                    Kind = "customFilters"
                                    And = flag "and" false custom
                                    Custom =
                                        tags
                                        |> List.choose (fun (name, element) ->
                                            if name = "customFilter" then
                                                Some ((attributeIn "operator" element |> Option.defaultValue "equal"), (attributeIn "val" element |> Option.map unescapeAttribute |> Option.defaultValue ""))
                                            else None) }
                            | None, None, Some top, _, _ ->
                                { column with
                                    Kind = "top10"
                                    Top = flag "top" true top
                                    Percent = flag "percent" false top
                                    Value = (attributeIn "val" top |> Option.bind tryParseFloat |> Option.defaultValue 10.0) }
                            | None, None, None, Some color, _ ->
                                { column with Kind = "colorFilter"; DxfId = (attributeIn "dxfId" color |> Option.bind tryParseInt); CellColor = flag "cellColor" true color }
                            | None, None, None, None, Some dynamic ->
                                { column with Kind = "dynamicFilter"; DynamicType = (attributeIn "type" dynamic |> Option.defaultValue "") }
                            | _ -> column
                    columns.Add parsed
            Some { Ref = reference; Columns = List.ofSeq columns }

/// The sheet and range a pivot cache definition reads (`<worksheetSource
/// ref sheet>`); None for a cache over a named range, a table or external data.
let pivotCacheSourceOf (cacheXml: string) : (string * string) option =
    let m = worksheetSourcePattern.Match cacheXml
    if not m.Success then None
    else
        match attributeIn "sheet" m.Value, attributeIn "ref" m.Value with
        | Some sheet, Some range when (attributeIn "r:id" m.Value).IsNone ->
            Some (sheet.Replace("&amp;", "&").Replace("&apos;", "'").Replace("&quot;", "\"").Replace("&lt;", "<").Replace("&gt;", ">"), range)
        | _ -> None

/// `cacheXml` with the range of its `<worksheetSource ref>` replaced.
let withPivotCacheSource (range: string) (cacheXml: string) : string =
    let m = worksheetSourcePattern.Match cacheXml
    if not m.Success then cacheXml
    else
        let tag = m.Value
        let at = tag.IndexOf " ref=\""
        if at < 0 then cacheXml
        else
            let start = m.Index + at + 6
            let close = cacheXml.IndexOf ('"', start)
            cacheXml.Substring(0, start) + range + cacheXml.Substring close

let private pivotLocationPattern = System.Text.RegularExpressions.Regex("<(?:[A-Za-z0-9_]+:)?location\\b[^>]*?\\sref=\"([^\"]*)\"")

/// The range of a pivot table part's `<location ref>`.
let pivotLocationOf (pivotXml: string) : string option =
    let m = pivotLocationPattern.Match pivotXml
    if m.Success then Some m.Groups.[1].Value else None

/// `pivotXml` with the range of its `<location ref>` replaced.
let withPivotLocation (range: string) (pivotXml: string) : string =
    let m = pivotLocationPattern.Match pivotXml
    if not m.Success then pivotXml
    else
        // (Not through the group's index: a JS regex match has none.)
        let whole = m.Value
        let at = m.Index + whole.LastIndexOf "ref=\"" + 5
        pivotXml.Substring(0, at) + range + pivotXml.Substring(at + m.Groups.[1].Value.Length)

/// Every part reachable from `roots` through the parts' own relationships
/// (charts with their embedded data and style parts), with those .rels parts.
let private reachablePartsFrom (source: PartSource) (roots: string list) : (string * byte[]) list =
    let seen = System.Collections.Generic.HashSet<string>()
    let parts = ResizeArray<string * byte[]>()
    let pending = System.Collections.Generic.Queue<string>(roots)
    while pending.Count > 0 do
        let path = pending.Dequeue()
        if seen.Add path then
            match source.Bytes path with
            | Some data ->
                parts.Add (path, data)
                let relsPath = relsPathFor path
                match source.Bytes relsPath with
                | Some relsData ->
                    if seen.Add relsPath then parts.Add (relsPath, relsData)
                    for rel in relationshipsOf source path do
                        if isInternal rel then pending.Enqueue (relTargetPath path rel)
                | None -> ()
            | None -> ()
    parts |> Seq.sortBy fst |> Seq.toList

let private contentTypesOf (source: PartSource) : (string * string) list * (string * string) list =
    match source.Text "[Content_Types].xml" |> Option.map XmlDsl.parse with
    | Some types ->
        let entries (elementName: string) (keyName: string) =
            Xml.children types
            |> List.choose (fun node ->
                if Xml.name node |> Option.map Ns.stripNs = Some elementName then
                    Option.map2 (fun key contentType -> key, contentType) (Xml.attr keyName node) (Xml.attr "ContentType" node)
                else None)
        entries "Default" "Extension", entries "Override" "PartName"
    | None -> [], []

let private defaultThemeColors = [|
    "FFFFFF"; "000000"; "E7E6E6"; "44546A"
    "4472C4"; "ED7D31"; "A5A5A5"; "FFC000"; "5B9BD5"; "70AD47"
    "0563C1"; "954F72"
|]

/// `scanRows` (native only; a browser always does): read the rows of each
/// sheet from the text of its part instead of with the streaming XML reader.
let private buildWorkbook (scanRows: bool) (source: PartSource) (header: WorkbookHeader) (sheetsToLoad: (int * Xlsx.SheetInfo * string) list) : Workbook =
    ignore scanRows
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

    // Which `<c cm>` values say "dynamic array formula".
    let dynamicCm =
        header.Rels
        |> List.tryFind (fun r -> r.Type.EndsWith "/sheetMetadata" && isInternal r)
        |> Option.bind (fun r -> source.Text (relTargetPath header.WorkbookPath r))
        |> Option.map (fun xml -> try dynamicArrayCellMetadata xml with _ -> Set.empty)
        |> Option.defaultValue Set.empty

    let persons =
        header.Rels
        |> List.filter (fun r -> r.Type.EndsWith "/person" && isInternal r)
        |> List.collect (fun r -> readPersons source (relTargetPath header.WorkbookPath r))

    let worksheets =
        sheetsToLoad
        |> List.choose (fun (sheetIdx, info, sheetPath) ->
            source.Text sheetPath
            |> Option.map (fun xml ->
                let printNames =
                    Map.tryFind sheetIdx header.PrintNames
                    |> Option.defaultValue emptyPrintNames
                let sheetRels = relationshipsOf source sheetPath
                let sheetRelType =
                    Office.relById info.RelationshipId header.Rels |> Option.map (fun rel -> rel.Type) |> Option.defaultValue ""
                let partsOfType (suffix: string) =
                    sheetRels
                    |> List.filter (fun r -> r.Type.EndsWith suffix && isInternal r)
                    |> List.choose (fun r -> source.Text (relTargetPath sheetPath r))
                let tablesWithParts =
                    sheetRels
                    |> List.filter (fun r -> r.Type.EndsWith "/table" && isInternal r)
                    |> List.choose (fun r ->
                        let path = relTargetPath sheetPath r
                        source.Text path |> Option.bind Xlsx.parseTable |> Option.map (fun table -> table, path))
                let tables = tablesWithParts |> List.map fst
                let drawingPart =
                    sheetRels
                    |> List.tryFind (fun r -> r.Type.EndsWith "/drawing" && isInternal r)
                    |> Option.map (relTargetPath sheetPath)
                    |> Option.filter source.Exists
                let comments = partsOfType "/comments" |> List.collect CellComments.parseComments
                let pivots =
                    partsOfType "/pivotTable"
                    |> List.collect (fun pivotXml -> try PivotTables.parsePivotTables pivotXml with _ -> [])
                // Pictures, charts and other objects: the sheet's drawing part.
                let images, objects, drawingRels =
                    sheetRels
                    |> List.tryFind (fun r -> r.Type.EndsWith "/drawing" && isInternal r)
                    |> Option.map (fun r -> try readDrawing source (relTargetPath sheetPath r) with _ -> [], [], [])
                    |> Option.defaultValue ([], [], [])
                let threads =
                    partsOfType "/threadedComment"
                    |> List.collect (fun threadXml -> try readThreadedComments persons threadXml with _ -> [])
                // What the model does not hold (pivot tables, slicers, timelines,
                // printer settings, controls, embedded objects ...): the parts
                // are kept as they are, with the elements that name them.
                let keptElements =
                    let all = try keptSheetElements xml with _ -> []
                    let known = sheetRels |> List.map (fun r -> r.Id) |> Set.ofList
                    all |> List.filter (fun (_, element) -> DocxReader.relationshipIdsInXml element |> List.forall (fun id -> Set.contains id known))
                // A VML drawing that holds nothing but the boxes of the sheet's
                // comments is only good while the comments stay as they are.
                let keptElements =
                    keptElements
                    |> List.map (fun (name, element) ->
                        if name <> "legacyDrawing" then name, element
                        else
                            let vml =
                                DocxReader.relationshipIdsInXml element
                                |> List.tryPick (fun id -> sheetRels |> List.tryFind (fun r -> r.Id = id && isInternal r))
                                |> Option.bind (fun r -> source.Text (relTargetPath sheetPath r))
                                |> Option.defaultValue ""
                            let count (needle: string) =
                                let mutable n = 0
                                let mutable at = vml.IndexOf needle
                                while at >= 0 do
                                    n <- n + 1
                                    at <- vml.IndexOf (needle, at + needle.Length)
                                n
                            let shapes = count "ClientData"
                            let notes = count "ObjectType=\"Note\""
                            // (Each shape has an opening and a closing ClientData tag.)
                            if not comments.IsEmpty && shapes <= notes * 2 then "legacyDrawingNotes", element else name, element)
                let namedIds = keptElements |> List.collect (fun (_, element) -> DocxReader.relationshipIdsInXml element) |> Set.ofList
                let keptRels =
                    sheetRels
                    |> List.filter (fun r -> isKeptSheetRelType r.Type || (r.Type.EndsWith "/vmlDrawing" && Set.contains r.Id namedIds))
                    |> List.map (fun r -> r.Id, r.Type, (if isInternal r then "/" + relTargetPath sheetPath r else r.Target), not (isInternal r))
                let keptIds = keptRels |> List.map (fun (id, _, _, _) -> id) |> Set.ofList
                let allExtensions =
                    keptSheetExtensions xml
                    |> List.filter (fun ext -> DocxReader.relationshipIdsInXml ext |> List.forall (fun id -> Set.contains id keptIds))
                let keptExtensions = allExtensions |> List.filter isPartExtension
                let cellExtensions = allExtensions |> List.filter (fun ext -> not (isPartExtension ext))
                let pivotLocations =
                    sheetRels
                    |> List.filter (fun r -> r.Type.EndsWith "/pivotTable" && isInternal r)
                    |> List.choose (fun r ->
                        let path = relTargetPath sheetPath r
                        source.Text path |> Option.bind pivotLocationOf |> Option.map (fun range -> path, range))
                let extras =
                    { emptySheetExtras with
                        Images = images; Objects = objects; ThreadedComments = threads; DrawingRelationships = drawingRels
                        SheetId = tryParseInt info.SheetId
                        KeptRelationships = keptRels
                        KeptExtensions = keptExtensions
                        KeptElements = keptElements
                        DrawingPart = drawingPart
                        TableParts = tablesWithParts |> List.map (fun (table, path) -> table.Name, path)
                        KeptCellExtensions = cellExtensions
                        AutoFilter = (try sheetAutoFilterOf xml with _ -> None)
                        PrinterSettingsRelId = pageSetupRelId xml |> Option.filter (fun id -> Set.contains id keptIds)
                        PivotLocations = pivotLocations
                        SourcePart = Some (sheetPath, sheetRelType)
                        CommentsPart =
                            sheetRels
                            |> List.tryFind (fun r -> r.Type.EndsWith "/comments" && isInternal r)
                            |> Option.map (relTargetPath sheetPath)
                            |> Option.filter (fun path -> (source.Bytes path).IsSome)
                        Verbatim = isForeignSheetRelType sheetRelType && (source.Bytes sheetPath).IsSome }
                // A fresh context per sheet: shared-formula groups are sheet-local.
                let ctx = Xlsx.newSheetContext (Some styles) header.Date1904
                ctx.DynamicCm <- dynamicCm
                let withExtras (sheet: Worksheet) = { sheet with Extras = { extras with DataTables = List.ofSeq ctx.DataTables } }
                let viaDom () =
                    let sheet = XmlDsl.parse xml
                    let condFormats = ConditionalFormats.parseConditionalFormats themeColors sheet
                    let dataVals = DataValidations.parseDataValidations sheet
                    Xlsx.worksheetFromNodeCore ctx info.Name info.State sheet comments condFormats dataVals [] pivots tables sheetRels printNames
                // In a browser the rows are read from the text of the part: a
                // DOM plus a node tree of a large sheet is most of the load time.
                let viaScan () =
                    Xlsx.worksheetScannedCore ctx info.Name info.State xml comments
                        (fun sheet -> ConditionalFormats.parseConditionalFormats themeColors sheet, DataValidations.parseDataValidations sheet)
                        pivots tables sheetRels printNames
                // The rows of a sheet stay in the text of its part until they
                // are asked for: opening a workbook does not read the sheets
                // nobody looks at.
                let newContext () =
                    let fresh = Xlsx.newSheetContext (Some styles) header.Date1904
                    fresh.DynamicCm <- dynamicCm
                    fresh
                let deferred () =
                    Xlsx.worksheetDeferredCore newContext info.Name info.State xml comments
                        (fun sheet -> ConditionalFormats.parseConditionalFormats themeColors sheet, DataValidations.parseDataValidations sheet)
                        pivots tables sheetRels printNames
                let scanned () =
                    match deferred () with
                    | Some sheet -> { sheet with Extras = extras }
                    | None ->
                    match viaScan () with
                    | Some sheet -> withExtras sheet
                    | None ->
                        // The scanner may have recorded shared formulas and data
                        // tables before it gave up: start again.
                        ctx.SharedFormulas.Clear()
                        ctx.DataTables.Clear()
                        withExtras (viaDom ())
#if FABLE_COMPILER
                scanned ()
#else
                if scanRows then scanned ()
                else
                    match Xlsx.worksheetNativeFastCore ctx info.Name info.State xml comments pivots tables sheetRels printNames with
                    | Some sheet -> withExtras sheet
                    | None -> withExtras (viaDom ())
#endif
            ))

    // Parts behind the kept drawing objects (chart parts, their embedded data),
    // and the theme: written back as they are.
    let keptRoots =
        (worksheets
         |> List.collect (fun sheet -> sheet.Extras.DrawingRelationships)
         |> List.choose (fun (_, _, target, isExternal) -> if isExternal || target.StartsWith "#" then None else Some (ZipInterop.normalizePath target)))
        @ (if source.Exists header.ThemePath then [ header.ThemePath ] else [])
        // A sheet that is written as the part it was read from.
        @ (worksheets |> List.choose (fun sheet -> if sheet.Extras.Verbatim then sheet.Extras.SourcePart |> Option.map fst else None))
        // The drawing and table parts themselves, for the sheets that leave them as they are.
        @ (worksheets |> List.collect (fun sheet -> Option.toList sheet.Extras.DrawingPart @ List.map snd sheet.Extras.TableParts))
        // Pivot tables and slicers of the sheets, and the workbook's pivot
        // caches, slicer caches and external links.
        @ ((worksheets |> List.collect (fun sheet -> sheet.Extras.KeptRelationships)) @ header.KeptRels
           |> List.choose (fun (_, _, target, isExternal) -> if isExternal then None else Some (ZipInterop.normalizePath target)))
    // The document properties: parts of the package itself, not of the workbook.
    let rootRels =
        (try source.Text "_rels/.rels" |> Option.map (XmlDsl.parse >> parseRelationships) |> Option.defaultValue [] with _ -> [])
        |> List.filter (fun r ->
            isInternal r
            // (Every part of the package but the workbook: the document properties, a thumbnail, sensitivity labels ...)
            && not (r.Type.EndsWith "/officeDocument")
            && source.Exists (ZipInterop.resolveTarget "" r.Target))
        |> List.map (fun r -> r.Id, r.Type, "/" + ZipInterop.resolveTarget "" r.Target, false)
    let calcChain =
        header.Rels
        |> List.tryFind (fun r -> r.Type.EndsWith "/calcChain" && isInternal r && source.Exists (relTargetPath header.WorkbookPath r))
        |> Option.map (fun r -> r.Id, r.Type, "/" + relTargetPath header.WorkbookPath r, false)
    let keptRoots =
        keptRoots
        @ (rootRels @ Option.toList calcChain |> List.map (fun (_, _, target, _) -> ZipInterop.normalizePath target))
    let pivotCacheSources =
        header.KeptRels
        |> List.choose (fun (_, relType, target, isExternal) ->
            if isExternal || not (relType.EndsWith "/pivotCacheDefinition") then None
            else
                let path = ZipInterop.normalizePath target
                source.Text path |> Option.bind pivotCacheSourceOf |> Option.map (fun (sheet, range) -> path, sheet, range))
    let defaults, overrides = contentTypesOf source
    let sharedStringTexts = sharedStringEntries |> Array.map (fun e -> e.Text)
    let sharedStringRuns = sharedStringEntries |> Array.map (fun e -> e.Runs)
    let namedStyles = stylesXml |> Option.map (fun xml -> try readNamedStyles xml with _ -> [], [||]) |> Option.defaultValue ([], [||])
    let preserved =
        { Parts = reachablePartsFrom source (List.distinct keptRoots)
          Defaults = defaults
          Overrides = overrides
          ThemePath = if source.Exists header.ThemePath then Some header.ThemePath else None
          NamedStyles = fst namedStyles
          Persons = persons
          CellXfNamedStyles = snd namedStyles
          WorkbookRelationships = header.KeptRels
          ExternalReferencesXml = header.ExternalReferencesXml
          PivotCachesXml = header.PivotCachesXml
          WorkbookExtLstXml = header.ExtLstXml
          RootRelationships = rootRels
          CalcChain = calcChain
          PivotCacheSources = pivotCacheSources
          CalcIteration = header.CalcIteration
          Date1904 = header.Date1904
          ActiveTab = header.ActiveTab
          StylesPart = if source.Exists header.StylesPath then Some header.StylesPath else None
          SharedStringsPart = if source.Exists header.SharedStringsPath then Some header.SharedStringsPath else None
          WorkbookPart = Some header.WorkbookPath
          CellFormats = styles
          SharedStrings = sharedStringTexts, sharedStringRuns
          DxfCount = dxfs.Length
          SourceParts =
            [ header.WorkbookPath; header.StylesPath; header.SharedStringsPath ] @ (worksheets |> List.choose (fun sheet -> sheet.Extras.SourcePart |> Option.map fst))
            |> List.distinct
            |> List.choose (fun path -> source.Bytes path |> Option.map (fun data -> path, data))
          PackageTexts =
            [ "[Content_Types].xml"; "_rels/.rels"; relsPathFor header.WorkbookPath ]
            |> List.choose (fun path -> (try source.Text path with _ -> None) |> Option.map (fun text -> path, text))
          InPlace = false }
    // ... and what the sheets' own relationships lead to (comments, their
    // boxes, threaded comments ...) with those relationship parts.
    let preserved =
        let known = System.Collections.Generic.HashSet<string>(preserved.Parts @ preserved.SourceParts |> List.map fst)
        let sheetPaths = worksheets |> List.choose (fun sheet -> sheet.Extras.SourcePart |> Option.map fst)
        let targets =
            sheetPaths
            |> List.collect (fun path -> relationshipsOf source path |> List.filter isInternal |> List.map (relTargetPath path))
            |> List.filter (fun path -> not (known.Contains path))
        let more =
            (reachablePartsFrom source (List.distinct targets)
             @ (sheetPaths |> List.choose (fun path -> let rels = relsPathFor path in source.Bytes rels |> Option.map (fun data -> rels, data))))
            |> List.filter (fun (path, _) -> known.Add path)
        if more.IsEmpty then preserved else { preserved with SourceParts = preserved.SourceParts @ more }

    {
        Sheets = worksheets
        SharedStrings = sharedStringTexts
        SharedStringRuns = sharedStringRuns
        Styles = styles
        Dxfs = dxfs
        Relationships = header.Rels
        NamedRanges = header.NamedRanges
        ThemeAccents = themeAccents
        Date1904 = header.Date1904
        Preserved = preserved
    }

#if FABLE_COMPILER
[<Fable.Core.Emit("new Uint8Array($0)")>]
let private bytesOfBuffer (buffer: obj) : byte[] = Fable.Core.Util.jsNative

[<Fable.Core.Emit("new TextDecoder('utf-8').decode($0)")>]
let private textOfBytes (bytes: byte[]) : string = Fable.Core.Util.jsNative
#else
let private bytesOfBuffer (buffer: obj) : byte[] = unbox buffer
let private textOfBytes (bytes: byte[]) : string = System.Text.Encoding.UTF8.GetString bytes
#endif

/// Parts that are read as text on demand and never needed as bytes: the
/// sheets and the shared strings, which are the large ones.
let private isLargeTextPart (path: string) =
    let lower = path.ToLowerInvariant()
    (lower.StartsWith "xl/worksheets/" && not (lower.Contains "/_rels/")) || lower.EndsWith "sharedstrings.xml"

// JSZip entry: await the parts each stage needs into a cache, then run the
// shared synchronous stages over it.
let private parseWorkbookFiltered loadAll sheetName (archive: ZipArchive) : Fable.Core.JS.Promise<Workbook option> =
    let cache = System.Collections.Generic.Dictionary<string, string>()
    // Bytes of everything but the large text parts: pictures, and the parts
    // kept for the next save (drawings, charts, theme).
    let byteCache = System.Collections.Generic.Dictionary<string, byte[]>()
    let source = {
        Exists = fun path -> (ZipInterop.tryFindFile archive path).IsSome
        Text = fun path ->
            match cache.TryGetValue path with
            | true, text -> Some text
            | _ ->
                match byteCache.TryGetValue path with
                | true, bytes ->
                    let text = textOfBytes bytes
                    cache.[path] <- text
                    Some text
                | _ -> None
        Bytes = fun path ->
            match byteCache.TryGetValue path with
            | true, bytes -> Some bytes
            | _ -> None
    }
    let fetchBytes () : Fable.Core.JS.Promise<unit> =
        let wanted =
            archive.Files
            |> Map.toArray
            |> Array.filter (fun (name, _) -> not (isLargeTextPart name))
        wanted
        |> Array.map (fun (_, file) -> file.GetArrayBuffer())
        |> Promise.all
        |> Promise.map (fun buffers ->
            buffers |> Array.iteri (fun i buffer -> byteCache.[fst wanted.[i]] <- bytesOfBuffer buffer))
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
            do! fetchBytes ()
            return Some (buildWorkbook true source header sheetsToLoad)
    }

/// Name, sheetId, relationship id and state of every sheet, from the
/// workbook part alone: no sheet, style or string part is read.
let private parseSheetInfos (archive: ZipArchive) : Fable.Core.JS.Promise<Xlsx.SheetInfo list option> =
    let cache = System.Collections.Generic.Dictionary<string, string>()
    let source = {
        Exists = fun path -> (ZipInterop.tryFindFile archive path).IsSome
        Text = fun path -> match cache.TryGetValue path with | true, text -> Some text | _ -> None
        Bytes = fun _ -> None
    }
    let fetch (paths: string list) : Fable.Core.JS.Promise<unit> =
        let wanted =
            paths
            |> List.choose (fun path -> ZipInterop.tryFindFile archive path |> Option.map (fun file -> path, file))
            |> List.toArray
        wanted
        |> Array.map (fun (_, file) -> file.GetText())
        |> Promise.all
        |> Promise.map (fun texts -> texts |> Array.iteri (fun i text -> cache.[fst wanted.[i]] <- text))
    promise {
        do! fetch [ "_rels/.rels" ]
        let workbookPath = mainPartPath source
        do! fetch [ workbookPath; relsPathFor workbookPath ]
        return
            readWorkbookHeader source workbookPath
            |> Option.map (fun header ->
                // The sheets a load would give: those whose part is in the package.
                selectedSheets header true None
                |> List.filter (fun (_, _, path) -> source.Exists path)
                |> List.map (fun (_, info, _) -> info))
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

/// The sheets a workbook lists (name, state), without reading any of them:
/// for a sheet picker, or to check that written bytes open as a workbook.
/// None when the package has no workbook part.
let loadSheetInfos (data: obj) : Fable.Core.JS.Promise<Xlsx.SheetInfo list option> =
    promise {
        let! archive = loadArchive data
        return! parseSheetInfos archive
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
let private parseWorkbookNativeWith (scanRows: bool) loadAll sheetName (archive: ZipArchive) : Workbook option =
    let source = {
        Exists = fun path -> (ZipInterop.tryFindFile archive path).IsSome
        Text = ZipInterop.getFileTextNative archive
        Bytes = ZipInterop.getFileBytesNative archive
    }
    readWorkbookHeader source (mainPartPath source)
    |> Option.map (fun header -> buildWorkbook scanRows source header (selectedSheets header loadAll sheetName))

let private parseWorkbookNativeFiltered loadAll sheetName (archive: ZipArchive) : Workbook option =
    parseWorkbookNativeWith false loadAll sheetName archive

/// The whole workbook, with the rows of every sheet read by the text scanner
/// the browser build uses (`Xlsx.scanSheetRows`) rather than the streaming
/// XML reader. For tests: both must give the same workbook.
let loadBytesNativeScanned (data: byte[]) : Workbook option =
    parseWorkbookNativeWith true true None (ZipInterop.loadArchiveNative data)

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
